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
import ssl
import json
import time
import queue
import email
import imaplib
import smtplib
import logging
import traceback
import threading
import requests
from email import policy
from email.message import EmailMessage
from email.header import decode_header, make_header
from email.utils import parseaddr, formataddr, formatdate

from flask import Flask, jsonify, Response

# -----------------------------------------------------------------------------
# Config & logging
# -----------------------------------------------------------------------------
def _coerce_log_level(val: str, default=logging.INFO) -> int:
    if not val:
        return default
    lv = str(val).strip().upper()
    return getattr(logging, lv, default)

LOG_LEVEL = _coerce_log_level(os.getenv("LOG_LEVEL", "INFO"))
logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

# Debug baixo nível do imaplib (0..5)
IMAP_DEBUG = int(os.getenv("IMAP_DEBUG", "0") or "0")
imaplib.Debug = IMAP_DEBUG

APP_PUBLIC_URL = os.getenv("APP_PUBLIC_URL")

# E-mail (IMAP/SMTP)
IMAP_HOST = os.getenv("IMAP_HOST", os.getenv("MAIL_IMAP_HOST", "mail.aprendacobol.com.br"))
IMAP_PORT = int(os.getenv("IMAP_PORT", os.getenv("MAIL_IMAP_PORT", "993")))
IMAP_SSL = (os.getenv("IMAP_SSL", "true").lower() in ("1", "true", "yes"))

# Credenciais: preferir MAIL_*; cair para IMAP_* se necessário
IMAP_USER = os.getenv("MAIL_USER") or os.getenv("IMAP_USER")
IMAP_PASS = os.getenv("MAIL_PASS") or os.getenv("IMAP_PASS")

SMTP_HOST = os.getenv("SMTP_HOST", os.getenv("MAIL_SMTP_HOST", "mail.aprendacobol.com.br"))
SMTP_PORT = int(os.getenv("SMTP_PORT", os.getenv("MAIL_SMTP_PORT", "587")))
SMTP_SSL = (os.getenv("SMTP_SSL", "false").lower() in ("1", "true", "yes"))
SMTP_STARTTLS = (os.getenv("SMTP_STARTTLS", "true").lower() in ("1", "true", "yes"))
SMTP_USER = os.getenv("SMTP_USER", IMAP_USER)
SMTP_PASS = os.getenv("SMTP_PASS", IMAP_PASS)
MAIL_FROM = os.getenv("MAIL_FROM", SMTP_USER or "")

# OpenRouter
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openrouter/auto")
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"

# Loop
POLL_INTERVAL_SEC = int(os.getenv("POLL_INTERVAL_SEC", "60"))

# Pastas
FOLDER_ESCALAR = os.getenv("FOLDER_ESCALAR", "INBOX.Escalar")
FOLDER_RESPONDIDOS = os.getenv("FOLDER_RESPONDIDOS", "INBOX.Respondidos")

# -----------------------------------------------------------------------------
# System prompt solicitado
# -----------------------------------------------------------------------------
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
    "Sempre ofereça o curso de Formação Completa de Programadore Aprenda COBOl disponível por assinatura em: https://assinatura.aprendacobol.com.br. "
)

# -----------------------------------------------------------------------------
# Utilitários de e-mail
# -----------------------------------------------------------------------------
def _decode_maybe_hdr(value: str) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except Exception:
        return value

def _extract_text_from_email(msg: email.message.Message) -> str:
    """Extrai texto legível do e-mail (prefere text/plain; senão tenta html sem tags)."""
    parts = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = (part.get_content_type() or "").lower()
            disp = (part.get("Content-Disposition") or "").lower()
            if ctype == "text/plain" and "attachment" not in disp:
                try:
                    parts.append(part.get_content().strip())
                except Exception:
                    try:
                        parts.append(part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "ignore"))
                    except Exception:
                        pass
    else:
        ctype = (msg.get_content_type() or "").lower()
        if ctype == "text/plain":
            try:
                parts.append(msg.get_content().strip())
            except Exception:
                try:
                    parts.append(msg.get_payload(decode=True).decode(msg.get_content_charset() or "utf-8", "ignore"))
                except Exception:
                    pass
        elif ctype == "text/html":
            raw = msg.get_payload(decode=True) or b""
            html = raw.decode(msg.get_content_charset() or "utf-8", "ignore")
            # Remover tags simples
            txt = re.sub(r"<br\s*/?>", "\n", html, flags=re.I)
            txt = re.sub(r"<[^>]+>", "", txt)
            parts.append(txt.strip())

    text = ("\n\n").join([p for p in parts if p]).strip()
    if not text:
        text = "(sem corpo de texto)"
    return text

def _build_reply(original: email.message.Message, assunto: str, corpo_markdown: str) -> EmailMessage:
    from_name, from_addr = parseaddr(original.get("To") or MAIL_FROM)
    orig_from_name, orig_from_addr = parseaddr(original.get("From") or "")
    reply = EmailMessage()
    reply["From"] = MAIL_FROM
    reply["To"] = formataddr((orig_from_name, orig_from_addr))
    # Assunto com Re: se necessário
    if not assunto:
        assunto = _decode_maybe_hdr(original.get("Subject") or "")
    if not assunto.lower().startswith("re:"):
        assunto = f"Re: {assunto}"
    reply["Subject"] = assunto
    # Threading headers
    if original.get("Message-ID"):
        reply["In-Reply-To"] = original["Message-ID"]
        refs = original.get("References")
        reply["References"] = f"{refs} {original['Message-ID']}" if refs else original["Message-ID"]
    reply["Date"] = formatdate(localtime=True)
    reply.set_content(corpo_markdown or "(sem corpo)", subtype="plain", charset="utf-8")
    return reply

def smtp_send(msg: EmailMessage):
    if not SMTP_HOST or not SMTP_PORT or not SMTP_USER or not SMTP_PASS or not MAIL_FROM:
        raise RuntimeError("Config SMTP incompleta.")
    logger.info("Enviando resposta SMTP para %s ...", msg["To"])
    if SMTP_SSL:
        ctx = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx, timeout=60) as s:
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as s:
            if SMTP_STARTTLS:
                s.starttls(context=ssl.create_default_context())
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
    logger.info("Resposta enviada com sucesso.")

# -----------------------------------------------------------------------------
# IMAP helpers (AUTH=PLAIN preferencial)
# -----------------------------------------------------------------------------
def imap_connect() -> imaplib.IMAP4:
    if not IMAP_USER or not IMAP_PASS:
        raise RuntimeError("Variáveis ausentes: MAIL_USER/MAIL_PASS (ou IMAP_USER/IMAP_PASS).")

    if IMAP_SSL:
        ctx = ssl.create_default_context()
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx)
    else:
        imap = imaplib.IMAP4(IMAP_HOST, IMAP_PORT)
        imap.starttls()

    # Capabilities (pré-login)
    try:
        typ, caps = imap.capability()
        caps_joined = " ".join(
            c.decode("utf-8", "ignore") if isinstance(c, bytes) else str(c) for c in (caps or [])
        )
        logger.debug("CAPABILITIES: %s", caps_joined)
    except Exception:
        pass

    # Tenta AUTH=PLAIN, fallback para LOGIN, com retry
    try:
        # após login o servidor divulga mais capabilities, mas já dá pra tentar AUTH=PLAIN
        try:
            imap.authenticate("PLAIN", lambda _: f"\0{IMAP_USER}\0{IMAP_PASS}".encode("utf-8"))
        except Exception as auth_err:
            logger.warning("AUTH=PLAIN falhou (%s). Tentando LOGIN ...", auth_err)
            imap.login(IMAP_USER, IMAP_PASS)
    except imaplib.IMAP4.error as e:
        logger.warning("Autenticação falhou: %s. Retentando em 2s ...", e)
        time.sleep(2.0)
        imap.login(IMAP_USER, IMAP_PASS)

    return imap

def ensure_mailbox(imap: imaplib.IMAP4, mailbox: str):
    try:
        typ, data = imap.list('""', '*')
        names = []
        if typ == "OK" and data:
            for raw in data:
                if not raw:
                    continue
                line = raw.decode("utf-8", "ignore")
                # nome é o último token entre aspas ou após o separador
                parts = line.split(' "." ')
                if len(parts) == 2:
                    names.append(parts[1].strip().strip('"'))
        if mailbox in names:
            return
        imap.create(mailbox)
    except Exception as e:
        # se já existe, ignorar
        if "ALREADYEXISTS" not in str(e).upper():
            logger.debug("ensure_mailbox('%s') aviso: %s", mailbox, e)

def move_message(imap: imaplib.IMAP4, seq_num: str, dest_mailbox: str):
    ensure_mailbox(imap, dest_mailbox)
    typ, _ = imap.copy(seq_num, dest_mailbox)
    if typ != "OK":
        raise RuntimeError(f"Falha no COPY para {dest_mailbox}")
    imap.store(seq_num, "+FLAGS", r"(\Deleted)")
    imap.expunge()

# -----------------------------------------------------------------------------
# OpenRouter: chamada + parser robusto
# -----------------------------------------------------------------------------
def call_openrouter(email_subject: str, email_from: str, email_text: str) -> str:
    """
    Retorna o texto bruto do modelo (pode vir com lixo, por isso existe parse abaixo).
    """
    if not OPENROUTER_API_KEY:
        # Sem API key: devolve stub que força escalar (para ambiente de teste)
        return '{"assunto":"","corpo_markdown":"","nivel_confianca":0.5,"acao":"escalar"}'

    user_prompt = (
        "Dados do e-mail a analisar (não responda aqui; siga o formato do JSON pedido no system):\n"
        f"Assunto: {email_subject}\n"
        f"De: {email_from}\n"
        "Corpo:\n"
        f"{email_text}\n"
    )

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": APP_PUBLIC_URL or "https://cobol-support-agent-cloud-hostgator.onrender.com",
        "X-Title": "COBOL Support Agent",
    }
    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        "max_tokens": 700,
        "temperature": 0.2,
    }

    try:
        resp = requests.post(OPENROUTER_URL, headers=headers, data=json.dumps(payload), timeout=60)
        resp.raise_for_status()
        j = resp.json()
        content = j["choices"][0]["message"]["content"]
        return content
    except Exception as e:
        logger.warning("Falha no OpenRouter: %s", e)
        return ""  # vai cair no fallback de escalar

def extract_first_json_object(text: str) -> dict:
    """
    Extrai o primeiro objeto JSON válido do texto, respeitando aspas e escapes.
    Levanta ValueError se não encontrar.
    """
    if not text:
        raise ValueError("vazio")
    # Retira eventuais cercas de código
    t = text.strip()
    if t.startswith("```"):
        t = t.strip("`").strip()
        # muitas vezes vem "json\n{...}"
        if t.lower().startswith("json"):
            t = t[4:].lstrip()

    in_str = False
    esc = False
    depth = 0
    start = -1
    for i, ch in enumerate(t):
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        else:
            if ch == '"':
                in_str = True
                continue
            if ch == "{":
                if depth == 0:
                    start = i
                depth += 1
                continue
            if ch == "}":
                depth -= 1
                if depth == 0 and start != -1:
                    candidate = t[start : i + 1]
                    try:
                        return json.loads(candidate)
                    except Exception as je:
                        # continua procurando o próximo bloco
                        pass
    raise ValueError("nenhum JSON válido encontrado")

def normalize_model_output(raw_text: str) -> dict:
    """
    Converte a saída do modelo em dict com chaves esperadas.
    Se falhar, levanta ValueError.
    """
    obj = extract_first_json_object(raw_text)
    # valida chaves
    for k in ("assunto", "corpo_markdown", "nivel_confianca", "acao"):
        if k not in obj:
            raise ValueError(f"chave ausente: {k}")
    ac = str(obj["acao"]).strip().lower()
    if ac not in ("responder", "escalar"):
        raise ValueError(f"acao inválida: {ac}")
    try:
        obj["nivel_confianca"] = float(obj["nivel_confianca"])
    except Exception:
        raise ValueError("nivel_confianca inválido")
    obj["assunto"] = str(obj.get("assunto") or "").strip()
    obj["corpo_markdown"] = str(obj.get("corpo_markdown") or "").strip()
    obj["acao"] = ac
    return obj

# -----------------------------------------------------------------------------
# Processamento IMAP (uma passada)
# -----------------------------------------------------------------------------
def process_unseen_once():
    imap = None
    try:
        imap = imap_connect()
        typ, _ = imap.select("INBOX")  # READ-WRITE
        if typ != "OK":
            raise RuntimeError("Falha ao selecionar INBOX")

        typ, data = imap.search(None, "UNSEEN")
        if typ != "OK":
            raise RuntimeError("Falha no SEARCH UNSEEN")

        ids = []
        if data and len(data) > 0 and data[0]:
            ids = data[0].split()

        logger.debug("UNSEEN: %s", ids)

        # nada a fazer
        if not ids:
            imap.expunge()
            return

        # Processa apenas o primeiro por ciclo (evita timeouts)
        seq = ids[0].decode() if isinstance(ids[0], bytes) else str(ids[0])

        # Busca corpo completo
        typ, msgdata = imap.fetch(seq, "(BODY.PEEK[])")
        if typ != "OK" or not msgdata or not msgdata[0]:
            raise RuntimeError("Falha no FETCH")

        raw = msgdata[0][1]
        msg = email.message_from_bytes(raw, policy=policy.default)

        subj = _decode_maybe_hdr(msg.get("Subject") or "")
        from_hdr = msg.get("From") or ""
        from_name, from_addr = parseaddr(from_hdr)
        body_txt = _extract_text_from_email(msg)

        # Chama modelo
        raw_out = call_openrouter(subj, from_addr, body_txt)

        # Tenta normalizar; em caso de falha, ESCALAR
        try:
            obj = normalize_model_output(raw_out)
        except Exception as e:
            preview = (raw_out or "")[:1200]
            logger.warning("Falha ao parsear JSON do modelo: %s | preview=%s", e, preview)
            obj = {"assunto": "", "corpo_markdown": "", "nivel_confianca": 0.5, "acao": "escalar"}

        acao = obj["acao"]
        conf = obj["nivel_confianca"]

        logger.info("Decisão do modelo: acao=%s conf=%.2f", acao, conf)

        if acao == "responder" and conf >= 0.80 and obj["corpo_markdown"]:
            try:
                reply = _build_reply(msg, obj["assunto"], obj["corpo_markdown"])
                smtp_send(reply)
                move_message(imap, seq, FOLDER_RESPONDIDOS)
                logger.info("E-mail respondido e movido para %s", FOLDER_RESPONDIDOS)
            except Exception as e_send:
                logger.error("Falha ao enviar resposta (%s). Escalando ...", e_send)
                move_message(imap, seq, FOLDER_ESCALAR)
        else:
            move_message(imap, seq, FOLDER_ESCALAR)
            logger.info("E-mail movido para %s", FOLDER_ESCALAR)

    except Exception as e:
        logger.error("Erro no ciclo IMAP: %s", e)
        logger.debug("Trace:\n%s", traceback.format_exc())
    finally:
        try:
            if imap is not None:
                imap.logout()
        except Exception:
            pass

# -----------------------------------------------------------------------------
# Worker loop (thread)
# -----------------------------------------------------------------------------
def worker_loop():
    if APP_PUBLIC_URL:
        logger.info("App público em: %s", APP_PUBLIC_URL)
    else:
        logger.info("App iniciando (APP_PUBLIC_URL não definido)")

    while True:
        t0 = time.time()
        process_unseen_once()
        elapsed = time.time() - t0
        sleep_for = max(1.0, POLL_INTERVAL_SEC - elapsed)
        time.sleep(sleep_for)

# -----------------------------------------------------------------------------
# Flask app
# -----------------------------------------------------------------------------
app = Flask(__name__)

@app.route("/")
def index():
    return Response("OK", mimetype="text/plain")

@app.route("/health")
def health():
    return jsonify({"ok": True, "ts": int(time.time())})

@app.route("/diag/openrouter-chat")
def diag_openrouter():
    # Diagnóstico simples
    if not OPENROUTER_API_KEY:
        # Sem API: retorna echo controlado
        return jsonify({"body_json": {"eco": "teste", "ok": True}, "model": OPENROUTER_MODEL, "ok": True})

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": APP_PUBLIC_URL or "https://cobol-support-agent-cloud-hostgator.onrender.com",
        "X-Title": "COBOL Support Agent",
    }
    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": "Teste rápido: responda apenas com o JSON minificado solicitado, acao=escalar, nivel_confianca=0.5."},
        ],
        "max_tokens": 60,
        "temperature": 0.0,
    }
    try:
        r = requests.post(OPENROUTER_URL, headers=headers, data=json.dumps(payload), timeout=30)
        r.raise_for_status()
        j = r.json()
        content = j["choices"][0]["message"]["content"]
        # tenta parse (apenas para validar)
        ok = True
        err = None
        try:
            _ = normalize_model_output(content)
        except Exception as e:
            ok = False
            err = str(e)
        return jsonify({"body_json": content, "ok": ok, "error": err, "model": OPENROUTER_MODEL})
    except Exception as e:
        return jsonify({"error": str(e), "ok": False, "model": OPENROUTER_MODEL})

# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    # Inicia worker em background
    threading.Thread(target=worker_loop, name="imap-worker", daemon=True).start()

    # Sobe o Flask
    port = int(os.getenv("PORT", "10000"))
    host = "0.0.0.0"
    logger.info("Watcher IMAP — envio via SMTP HostGator")
    if APP_PUBLIC_URL:
        logger.info("App público em: %s", APP_PUBLIC_URL)
    app.run(host=host, port=port, debug=False)
