#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
COBOL Support Agent — IMAP watcher + SMTP sender + OpenRouter
- Lê INBOX por IMAP em polling
- Classifica/gera ação via OpenRouter
- Responde por SMTP OU move para INBOX.Escalar/INBOX.Respondidos
- Exibe rotas / e /diag/* para health-check

Requisitos: apenas libs padrão + requests (Render já tem).
"""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import json
import time
import ssl
import imaplib
import smtplib
import logging
import threading
import traceback
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from email.header import decode_header, make_header
from email.utils import parseaddr, formataddr, formatdate, make_msgid

import requests

# -----------------------------
# Logging
# -----------------------------
def _parse_log_level(s: str):
    if not s:
        return logging.INFO
    s = s.strip().upper()
    return {
        "CRITICAL": logging.CRITICAL,
        "ERROR": logging.ERROR,
        "WARNING": logging.WARNING,
        "INFO": logging.INFO,
        "DEBUG": logging.DEBUG,
        "NOTSET": logging.NOTSET,
    }.get(s, logging.INFO)

logging.basicConfig(
    level=_parse_log_level(os.getenv("LOG_LEVEL", "INFO")),
    format="%(asctime)s [%(levelname)s] %(message)s",
)

log = logging.getLogger(__name__)

# -----------------------------
# Config (env)
# -----------------------------
APP_PUBLIC_URL     = os.getenv("APP_PUBLIC_URL", "")
APP_TITLE          = os.getenv("APP_TITLE", "COBOL Support Agent")

IMAP_HOST          = os.getenv("IMAP_HOST", "localhost")
IMAP_PORT          = int(os.getenv("IMAP_PORT", "993"))
MAIL_USER          = os.getenv("MAIL_USER") or os.getenv("IMAP_USER", "")
MAIL_PASS          = os.getenv("MAIL_PASS") or os.getenv("IMAP_PASS", "")

FOLDER_ESCALATE    = os.getenv("FOLDER_ESCALATE", "Escalar")
FOLDER_PROCESSED   = os.getenv("FOLDER_PROCESSED", "Respondidos")
SENT_FOLDER        = os.getenv("SENT_FOLDER", "INBOX.Sent")  # você pode mudar para INBOX.Enviados
SAVE_SENT_COPY     = os.getenv("SAVE_SENT_COPY", "true").lower() == "true"

CHECK_INTERVAL     = int(os.getenv("CHECK_INTERVAL_SECONDS", "60"))
EXPUNGE_AFTER_COPY = os.getenv("EXPUNGE_AFTER_COPY", "true").lower() == "true"

# Limiar de resposta
RESPOND_THRESHOLD = float(
    os.getenv("RESPOND_THRESHOLD", os.getenv("CONFIDENCE_THRESHOLD", "0.8"))
)

# SMTP
SMTP_HOST      = os.getenv("SMTP_HOST", IMAP_HOST)
SMTP_PORT      = int(os.getenv("SMTP_PORT", "587"))
SMTP_TLS_MODE  = os.getenv("SMTP_TLS_MODE", "starttls").lower()  # starttls | ssl | off

# Assinatura
SIGNATURE_NAME   = os.getenv("SIGNATURE_NAME", "Equipe Aprenda COBOL — Suporte")
SIGNATURE_LINKS  = os.getenv("SIGNATURE_LINKS", "https://aprendacobol.com.br/assinatura/")
SIGNATURE_FOOTER = os.getenv("SIGNATURE_FOOTER", "").strip().strip('"')

# Backends LLM
LLM_BACKEND = os.getenv("LLM_BACKEND", "openrouter").lower()

# OpenRouter
OR_API_KEY      = os.getenv("OPENROUTER_API_KEY", "")
OR_MODEL        = os.getenv("OPENROUTER_MODEL", "openrouter/auto")
OR_MAX_TOKENS   = int(os.getenv("OPENROUTER_MAX_TOKENS", "512"))
OR_APP_NAME     = os.getenv("OPENROUTER_APP_NAME", APP_TITLE)
OR_SITE_URL     = os.getenv("OPENROUTER_SITE_URL", APP_PUBLIC_URL)

# Ollama (placeholder)
OLLAMA_HOST     = os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434")
OLLAMA_MODEL    = os.getenv("OLLAMA_MODEL", "llama3.1:8b")

PORT            = int(os.getenv("PORT", "10000"))

# cache da pasta de enviados detectada
_DETECTED_SENT_FOLDER = None

# -----------------------------
# Prompt do modelo
# -----------------------------
SYSTEM_PROMPT = (
    "Você é um assistente de suporte de um curso de COBOL. "
    "SEMPRE produza um JSON VÁLIDO e nada além disso. "
    "Formato do JSON (minificado, sem comentários, sem markdown, sem texto extra): "
    '{"assunto": "...", "corpo_markdown": "...", "nivel_confianca": 0.0, "acao": "responder|escalar"} '
    "Regras: "
    "1) NUNCA inclua crases ou ``` no output. "
    "2) NUNCA acrescente explicações fora do JSON. "
    "3) Sempre use chaves exatamente como no esquema. "
    "4) PT-BR no corpo. "
    "5) 'nivel_confianca' entre 0 e 1. "
    "6) Se pedido estiver claro e respondível, 'acao'='responder' com nivel_confianca>=0.8; "
    "   se ambíguo/incompleto, 'acao'='escalar' com nivel_confianca<=0.6."
    "Analise o e-mail (texto e código anexo ou no corpo do e-mail). Responda educadamente. "
    "Cite pontos do COBOL (DIVISION, SECTION, PIC, níveis, I/O, SQLCA etc.). "
    "Sugira passos práticos. Se no e-mail tiver qualquer solicitação para entrar no grupo do Telegram, dê o link 'https://t.me/aprendacobol'. "
    "Sempre ofereça o curso de Formação Completa de Programadore Aprenda COBOl disponível por assinatura em: https://assinatura.aprendacobol.com.br."
)

# -----------------------------
# Utils
# -----------------------------
def ensure_inbox_prefix(folder: str) -> str:
    if not folder:
        return "INBOX.Sent"
    if folder.upper().startswith("INBOX"):
        return folder
    return f"INBOX.{folder}"

def decode_mime_header(value: str) -> str:
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value or ""

def get_msg_text(msg) -> str:
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition", "")).lower()
            if ctype == "text/plain" and "attachment" not in disp:
                try:
                    return part.get_content().strip()
                except Exception:
                    try:
                        return part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="replace").strip()
                    except Exception:
                        continue
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition", "")).lower()
            if ctype == "text/html" and "attachment" not in disp:
                try:
                    html = part.get_content()
                except Exception:
                    html = part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", errors="replace")
                text = re.sub(r"(?s)<(script|style).*?>.*?</\1>", "", html, flags=re.I)
                text = re.sub(r"(?s)<br\s*/?>", "\n", text, flags=re.I)
                text = re.sub(r"(?s)</p\s*>", "\n\n", text, flags=re.I)
                text = re.sub(r"(?s)<[^>]+>", "", text)
                return text.strip()
    else:
        if msg.get_content_type() == "text/plain":
            try:
                return msg.get_content().strip()
            except Exception:
                return msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8", errors="replace").strip()
        if msg.get_content_type() == "text/html":
            html = msg.get_content()
            text = re.sub(r"(?s)<(script|style).*?>.*?</\1>", "", html, flags=re.I)
            text = re.sub(r"(?s)<br\s*/?>", "\n", text, flags=re.I)
            text = re.sub(r"(?s)</p\s*>", "\n\n", text, flags=re.I)
            text = re.sub(r"(?s)<[^>]+>", "", text)
            return text.strip()
    return ""

def make_signature() -> str:
    parts = []
    if SIGNATURE_NAME:
        parts.append(SIGNATURE_NAME)
    if SIGNATURE_FOOTER:
        parts.append(SIGNATURE_FOOTER)
    if SIGNATURE_LINKS:
        parts.append(SIGNATURE_LINKS)
    return "\n\n" + "\n".join(parts).strip() if parts else ""

def sanitize_model_text_to_json(text: str) -> str:
    if not text:
        return ""
    text = text.strip().strip("`").strip()
    start = text.find("{")
    if start == -1:
        return ""
    brace = 0
    in_str = False
    esc = False
    for i, ch in enumerate(text[start:], start=start):
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
        else:
            if ch == '"':
                in_str = True
            elif ch == "{":
                brace += 1
            elif ch == "}":
                brace -= 1
                if brace == 0:
                    return text[start:i+1]
    return ""

def parse_llm_json(text: str):
    if not text:
        return None
    def try_load(s):
        try:
            return json.loads(s)
        except Exception:
            return None
    data = try_load(text)
    if data is not None:
        return data
    cand = sanitize_model_text_to_json(text)
    data = try_load(cand)
    if data is not None:
        return data
    cand2 = cand.replace("\u201c", '"').replace("\u201d", '"').replace("\u2018", "'").replace("\u2019", "'")
    data = try_load(cand2)
    return data

def normalize_decision(d: dict):
    if not isinstance(d, dict):
        return None
    acao = str(d.get("acao", "")).strip().lower()
    if acao not in ("responder", "escalar"):
        acao = "escalar"
    try:
        nivel = float(d.get("nivel_confianca", 0.0))
    except Exception:
        nivel = 0.0
    nivel = max(0.0, min(1.0, nivel))
    assunto = str(d.get("assunto", "") or "").strip()
    corpo = str(d.get("corpo_markdown", "") or "").strip()
    return acao, nivel, assunto, corpo

# -----------------------------
# LLM clients
# -----------------------------
def call_openrouter(messages):
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OR_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": OR_SITE_URL or APP_PUBLIC_URL or "https://example.com",
        "X-Title": OR_APP_NAME or APP_TITLE or "App",
    }
    payload = {
        "model": OR_MODEL,
        "messages": messages,
        "max_tokens": OR_MAX_TOKENS,
        "temperature": 0.2,
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    j = resp.json()
    content = (j.get("choices", [{}])[0].get("message", {}) or {}).get("content", "")
    return content

def classify_email_with_llm(subject: str, sender: str, body: str) -> dict | None:
    if LLM_BACKEND == "openrouter":
        if not OR_API_KEY:
            log.warning("OPENROUTER_API_KEY ausente; escalando por segurança.")
            return None
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"Remetente: {sender}\nAssunto: {subject}\n\nCorpo do e-mail:\n{body}"},
        ]
        try:
            raw = call_openrouter(messages)
            log.debug("OpenRouter bruto (preview): %s", (raw[:1000] if raw else ""))
            data = parse_llm_json(raw)
            return data
        except Exception as e:
            log.warning("Falha ao chamar OpenRouter: %s", e)
            return None
    else:
        log.warning("LLM_BACKEND desconhecido ou não implementado: %s; escalando.", LLM_BACKEND)
        return None

# -----------------------------
# SMTP + salvar cópia em Enviados
# -----------------------------
def imap_ensure_folder(imap, folder_name: str):
    try:
        typ, _ = imap.create(folder_name)
        # se já existe, vem NO [ALREADYEXISTS]; ignoramos
    except Exception:
        pass

def imap_detect_sent_folder(imap) -> str:
    """
    Tenta detectar a pasta marcada com \Sent. Se não achar, usa SENT_FOLDER (com INBOX.*).
    Cacheia o resultado para não listar toda hora.
    """
    global _DETECTED_SENT_FOLDER
    if _DETECTED_SENT_FOLDER:
        return _DETECTED_SENT_FOLDER

    prefer = ensure_inbox_prefix(SENT_FOLDER)
    try:
        typ, data = imap.list("", "*")
        if typ == "OK" and data:
            for raw in data:
                try:
                    line = raw.decode("utf-7") if isinstance(raw, bytes) else str(raw)
                except Exception:
                    line = str(raw)
                # Exemplo: * LIST (\HasNoChildren \UnMarked \Sent) "." INBOX.Sent
                if "\\Sent" in line:
                    # nome da mailbox é o último token depois do separador
                    mbox = line.split(' "', 1)[-1]
                    # acima não é robusto pra todos; melhor pegar após o último espaço
                    mbox = line.split(" ", maxsplit=3)[-1].strip()
                    # remove aspas, se houver
                    mbox = mbox.strip('"')
                    _DETECTED_SENT_FOLDER = mbox
                    log.info("Pasta \\Sent detectada: %s", _DETECTED_SENT_FOLDER)
                    return _DETECTED_SENT_FOLDER
    except Exception as e:
        log.debug("Falha ao listar pastas para detectar \\Sent: %s", e)

    _DETECTED_SENT_FOLDER = prefer
    return _DETECTED_SENT_FOLDER

def imap_append_sent_copy(imap, msg: EmailMessage):
    try:
        dest = imap_detect_sent_folder(imap)
        imap_ensure_folder(imap, dest)
        date_time = imaplib.Time2Internaldate(time.time())
        # marca como lida na pasta de enviados
        imap.append(dest, r"(\Seen)", date_time, msg.as_bytes())
        log.info("Cópia da resposta salva em %s", dest)
    except Exception as e:
        log.warning("Não foi possível salvar cópia em enviados: %s", e)

def smtp_send_reply(to_addr: str, subject: str, body_markdown: str,
                    in_reply_to_msgid: str | None, original_msg,
                    imap_for_copy=None) -> EmailMessage:
    from_addr = MAIL_USER
    if not from_addr:
        raise RuntimeError("MAIL_USER não definido")

    full_body = body_markdown.strip() + make_signature()

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = formataddr(("Suporte Aprenda COBOL", from_addr))
    msg["To"] = to_addr
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid()
    if in_reply_to_msgid:
        msg["In-Reply-To"] = in_reply_to_msgid
        refs = original_msg.get_all("References", [])
        if in_reply_to_msgid not in (refs or []):
            refs = (refs or []) + [in_reply_to_msgid]
        if refs:
            msg["References"] = " ".join(refs)

    msg.set_content(full_body, subtype="plain", charset="utf-8")

    server = None
    try:
        if SMTP_TLS_MODE == "ssl":
            context = ssl.create_default_context()
            server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context, timeout=60)
        else:
            server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60)
            server.ehlo()
            if SMTP_TLS_MODE == "starttls":
                context = ssl.create_default_context()
                server.starttls(context=context)
                server.ehlo()
        server.login(MAIL_USER, MAIL_PASS)
        server.send_message(msg)
        log.info("Resposta enviada via SMTP para %s", to_addr)
    finally:
        try:
            if server:
                server.quit()
        except Exception:
            pass

    # salva cópia em Enviados pelo IMAP (na mesma sessão de leitura)
    if SAVE_SENT_COPY and imap_for_copy is not None:
        imap_append_sent_copy(imap_for_copy, msg)

    return msg

# -----------------------------
# IMAP helpers
# -----------------------------
def imap_connect() -> imaplib.IMAP4_SSL:
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    try:
        typ, caps = imap.capability()
        if typ == "OK":
            # Pre-login caps
            caps_str = ""
            try:
                caps_str = " ".join(caps[0].decode().split()[2:]) if caps else ""
            except Exception:
                caps_str = str(caps)
            log.debug("CAPABILITIES: %s", caps_str)
    except Exception:
        pass
    imap.login(MAIL_USER, MAIL_PASS)
    return imap

def imap_select_inbox(imap):
    imap.select("INBOX")

def move_to_folder(imap, seq: bytes, folder_short: str):
    dest = ensure_inbox_prefix(folder_short)
    imap_ensure_folder(imap, dest)
    imap.copy(seq, dest)
    imap.store(seq, "+FLAGS", r"(\Deleted)")
    if EXPUNGE_AFTER_COPY:
        imap.expunge()

# -----------------------------
# Processamento principal
# -----------------------------
def process_unseen_once():
    imap = None
    try:
        imap = imap_connect()
        imap_select_inbox(imap)

        typ, data = imap.search(None, "UNSEEN")
        if typ != "OK":
            log.warning("SEARCH UNSEEN falhou: %s", typ)
            return

        ids = data[0].split()
        log.debug("UNSEEN: %s", ids)
        if not ids:
            return

        for seq in ids:
            typ, msgdata = imap.fetch(seq, "(BODY.PEEK[])")
            if typ != "OK" or not msgdata or not isinstance(msgdata[0], tuple):
                log.warning("FETCH falhou para seq=%s", seq)
                continue

            raw = msgdata[0][1]
            msg = BytesParser(policy=policy.default).parsebytes(raw)

            subj = decode_mime_header(msg.get("Subject", ""))
            from_name, from_addr = parseaddr(msg.get("From", ""))
            sender_disp = formataddr((from_name, from_addr))
            in_reply_to = msg.get("Message-ID", None)

            body = get_msg_text(msg)

            decision_raw = classify_email_with_llm(subj, sender_disp, body)
            if decision_raw is None:
                log.info("Decisão do modelo ausente → escalar por segurança.")
                move_to_folder(imap, seq, FOLDER_ESCALATE)
                continue

            parsed = normalize_decision(decision_raw)
            if not parsed:
                log.info("JSON inválido do modelo → escalar por segurança.")
                move_to_folder(imap, seq, FOLDER_ESCALATE)
                continue

            acao, nivel, assunto_resp, corpo_md = parsed
            log.info("Decisão do modelo: acao=%s conf=%.2f", acao, nivel)

            if acao == "responder" and nivel >= RESPOND_THRESHOLD and corpo_md.strip():
                try:
                    reply_to = from_addr
                    reply_subject = assunto_resp if assunto_resp else f"Re: {subj or ''}".strip()
                    smtp_send_reply(
                        reply_to,
                        reply_subject,
                        corpo_md,
                        in_reply_to,
                        msg,
                        imap_for_copy=imap  # <<< salva cópia no Enviados
                    )
                    move_to_folder(imap, seq, FOLDER_PROCESSED)
                    log.info("Resposta enviada e e-mail movido para INBOX.%s", FOLDER_PROCESSED)
                except Exception as e:
                    log.warning("Falha ao enviar resposta (%s) → escalar.", e)
                    move_to_folder(imap, seq, FOLDER_ESCALATE)
            else:
                move_to_folder(imap, seq, FOLDER_ESCALATE)
                log.info("E-mail movido para INBOX.%s", FOLDER_ESCALATE)

    except Exception as e:
        log.error("Erro no ciclo IMAP: %s", e)
        log.debug("Traceback:\n%s", traceback.format_exc())
    finally:
        try:
            if imap:
                imap.logout()
        except Exception:
            pass

def watcher_loop():
    log.info("Watcher IMAP — envio via SMTP HostGator")
    if APP_PUBLIC_URL:
        log.info("App público em: %s", APP_PUBLIC_URL)
    while True:
        process_unseen_once()
        time.sleep(CHECK_INTERVAL)

# -----------------------------
# Flask app
# -----------------------------
from flask import Flask, jsonify

app = Flask(__name__)

@app.route("/")
def index():
    return f"{APP_TITLE} está rodando."

@app.route("/health")
def health():
    return jsonify({"ok": True, "service": APP_TITLE})

@app.route("/diag/openrouter-chat")
def diag_or():
    if LLM_BACKEND != "openrouter":
        return jsonify({"ok": False, "error": "LLM_BACKEND != openrouter", "backend": LLM_BACKEND}), 400
    if not OR_API_KEY:
        return jsonify({"ok": False, "error": "OPENROUTER_API_KEY ausente"}), 400
    try:
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": "eco"},
        ]
        content = call_openrouter(messages)
        return jsonify({
            "ok": True,
            "model": OR_MODEL,
            "headers_sent": {
                "HTTP-Referer": OR_SITE_URL or APP_PUBLIC_URL,
                "Referer": OR_SITE_URL or APP_PUBLIC_URL,
                "X-Title": OR_APP_NAME or APP_TITLE,
            },
            "body_json": {"ok": True, "eco": "teste"},
            "raw": content[:2000] if content else "",
        })
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

def start_background_thread():
    t = threading.Thread(target=watcher_loop, name="imap-watcher", daemon=True)
    t.start()

if __name__ == "__main__":
    if SMTP_TLS_MODE == "ssl" and SMTP_PORT == 587:
        log.warning("SMTP_TLS_MODE=ssl com porta 587 — normalmente 587 requer STARTTLS. "
                    "Considere usar SMTP_TLS_MODE=starttls OU porta 465 com ssl.")
    if SMTP_TLS_MODE == "starttls" and SMTP_PORT == 465:
        log.warning("SMTP_TLS_MODE=starttls com porta 465 — normalmente 465 é SSL implícito. "
                    "Considere SMTP_TLS_MODE=ssl OU porta 587 com starttls.")

    start_background_thread()
    app.run(host="0.0.0.0", port=PORT, debug=False)
