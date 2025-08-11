#!/usr/bin/env python3 - v9.6 (atualizado)
# -*- coding: utf-8 -*-

"""
COBOL Support Agent — IMAP watcher + SMTP sender + OpenRouter
- Lê INBOX por IMAP em polling
- Classifica/gera ação via OpenRouter
- Responde por SMTP OU move para INBOX.Escalar/INBOX.Respondidos
- Exibe rotas / e /diag/* para health-check

Requisitos: apenas libs padrão + requests (Render já tem).
"""

import os
import re
import ssl
import time
import json
import hmac
import hashlib
import logging
import imaplib
import smtplib
from datetime import datetime, timezone
from email import policy
from email.parser import BytesParser
from email.message import EmailMessage
from email.header import decode_header, make_header
from email.utils import formatdate, make_msgid, parsedate_to_datetime

from flask import Flask, jsonify

# ==========================
# Config & Logging
# ==========================
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

# IMAP
IMAP_HOST = os.getenv("IMAP_HOST")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
MAIL_USER = os.getenv("MAIL_USER")
MAIL_PASS = os.getenv("MAIL_PASS")
FOLDER_PROCESSED = os.getenv("FOLDER_PROCESSED", "Respondidos")
FOLDER_ESCALATE = os.getenv("FOLDER_ESCALATE", "Escalar")
EXPUNGE_AFTER_COPY = os.getenv("EXPUNGE_AFTER_COPY", "true").lower() == "true"
SENT_FOLDER = os.getenv("SENT_FOLDER", "INBOX.Sent")

# SMTP
SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_TLS_MODE = os.getenv("SMTP_TLS_MODE", "starttls").lower()  # starttls|ssl
SMTP_DEBUG = int(os.getenv("SMTP_DEBUG", "0"))

# LLM / OpenRouter
LLM_BACKEND = os.getenv("LLM_BACKEND", "openrouter")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openrouter/auto")
OPENROUTER_MAX_TOKENS = int(os.getenv("OPENROUTER_MAX_TOKENS", "512"))
OPENROUTER_APP_NAME = os.getenv("OPENROUTER_APP_NAME", "COBOL Support Agent")
OPENROUTER_SITE_URL = os.getenv("OPENROUTER_SITE_URL", "")
CONFIDENCE_THRESHOLD = float(os.getenv("CONFIDENCE_THRESHOLD", "0.8"))

# App
CHECK_INTERVAL_SECONDS = int(os.getenv("CHECK_INTERVAL_SECONDS", "60"))
APP_PUBLIC_URL = os.getenv("APP_PUBLIC_URL", "")
APP_TITLE = os.getenv("APP_TITLE", "COBOL Support Agent")

# Assinatura / rodapé
SIGNATURE_NAME = os.getenv("SIGNATURE_NAME", "Equipe Aprenda COBOL — Suporte")
SIGNATURE_FOOTER = os.getenv(
    "SIGNATURE_FOOTER",
    (
        "Se precisar, responda este e-mail com mais detalhes ou anexe seu arquivo .COB/.CBL.\n"
        "Horário de atendimento: 9h–18h (ET), seg–sex.\n"
        "Conheça nossa Formação Completa de Programadores COBOL, com COBOL Avançado,\n"
        "JCL, Db2 e Bancos de Dados completo em:"
    ),
)
SIGNATURE_LINKS = os.getenv("SIGNATURE_LINKS", "https://aprendacobol.com.br/assinatura/")

# ==========================
# Prompt do sistema (com CTAs obrigatórios)
# ==========================
SYSTEM_PROMPT = (
    "Você é um assistente de suporte de um curso de COBOL. "
    "SEMPRE produza um JSON VÁLIDO e nada além disso. "
    "Formato do JSON (minificado, sem comentários, sem markdown, sem texto extra): "
    "{\"assunto\": \"...\", \"corpo_markdown\": \"...\", \"nivel_confianca\": 0.0, \"acao\": \"responder|escalar\"} "
    "Regras: "
    "1) NUNCA inclua crases ou ``` no output. "
    "2) NUNCA acrescente explicações fora do JSON. "
    "3) Sempre use chaves exatamente como no esquema. "
    "4) PT-BR no corpo. "
    "5) 'nivel_confianca' entre 0 e 1. "
    "6) Se pedido estiver claro e respondível, 'acao'='responder' com nivel_confianca>=0.8; "
    "   se ambíguo/incompleto, 'acao'='escalar' com nivel_confianca<=0.6. "
    "7) Assunto: defina 'assunto' EXATAMENTE como o assunto original do e-mail (não traduza, não resuma, não invente). "
    "   Se o original já tiver 'Re:' no início, mantenha como está. OBS: o sistema adicionará 'Re: ' no envio se faltar. "
    "8) Se houver arquivo anexo .COB/.CBL/.CPY com código COBOL, priorize analisar o código; cite elementos COBOL "
    "   (DIVISION, SECTION, PIC, níveis, I/O, SQLCA etc.). Identifique erros comuns e sugira correções objetivas. "
    "9) Não mude o tema da conversa. Responda ao que foi solicitado, de forma educada e objetiva. "
    "10) Se faltar informação para compilar/executar, peça os dados mínimos (ex.: amostras de entrada/saída, layout, JCL). "
    "11) No final do 'corpo_markdown', SEMPRE inclua exatamente estas duas linhas (URLs como texto puro, sem markdown de link): "
    "- Entre no nosso grupo no Telegram: https://t.me/aprendacobol "
    "- Conheça a Formação Completa de Programador COBOL: https://assinatura.aprendacobol.com.br "
)

# sha1 do prompt para diagnóstico
SYSTEM_PROMPT_SHA1 = hashlib.sha1(SYSTEM_PROMPT.encode("utf-8")).hexdigest()
log.info("SYSTEM_PROMPT_SHA1=%s (primeiros 120 chars): %s", SYSTEM_PROMPT_SHA1[:12], SYSTEM_PROMPT[:120])

# ==========================
# Utilitários de e-mail
# ==========================

def _safe_box(name: str) -> str:
    # se já vier com INBOX. mantemos; caso contrário, prefixamos.
    if name.upper().startswith("INBOX"):
        return name
    return f"INBOX.{name}"


def decode_mime_words(s):
    if not s:
        return ""
    try:
        return str(make_header(decode_header(s)))
    except Exception:
        return s


def extract_text_body(msg):
    # Prioriza text/plain; se não houver, converte text/html para texto
    text_parts = []
    html_parts = []
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get("Content-Disposition") or "").lower()
            if ctype == "text/plain" and "attachment" not in disp:
                try:
                    text_parts.append(part.get_content())
                except Exception:
                    pass
            elif ctype == "text/html" and "attachment" not in disp:
                try:
                    html_parts.append(part.get_content())
                except Exception:
                    pass
    else:
        ctype = msg.get_content_type()
        if ctype == "text/plain":
            try:
                text_parts.append(msg.get_content())
            except Exception:
                pass
        elif ctype == "text/html":
            try:
                html_parts.append(msg.get_content())
            except Exception:
                pass

    if text_parts:
        return "\n\n".join(t.strip() for t in text_parts if t)

    # fallback simples de HTML->texto
    if html_parts:
        html = "\n\n".join(html_parts)
        # remove tags básicas
        text = re.sub(r"<br\s*/?>", "\n", html, flags=re.I)
        text = re.sub(r"</p>", "\n\n", text, flags=re.I)
        text = re.sub(r"<[^>]+>", "", text)
        return re.sub(r"\n{3,}", "\n\n", text).strip()

    return ""


def extract_cobol_attachments(msg, max_bytes=80_000):
    cobol_files = []
    if msg.is_multipart():
        for part in msg.walk():
            disp = (part.get("Content-Disposition") or "").lower()
            if "attachment" in disp:
                filename = part.get_filename()
                filename = decode_mime_words(filename)
                if not filename:
                    continue
                lower = filename.lower()
                if lower.endswith((".cob", ".cbl", ".cpy")):
                    try:
                        data = part.get_payload(decode=True)
                        if not data:
                            continue
                        snippet = data[:max_bytes].decode("utf-8", errors="replace")
                        cobol_files.append((filename, snippet))
                    except Exception:
                        continue
    return cobol_files


# ==========================
# OpenRouter client
# ==========================
import requests

def call_openrouter(system_prompt: str, user_prompt: str) -> dict:
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": OPENROUTER_SITE_URL or APP_PUBLIC_URL or "",
        "X-Title": OPENROUTER_APP_NAME,
    }

    payload = {
        "model": OPENROUTER_MODEL,
        "max_tokens": OPENROUTER_MAX_TOKENS,
        "temperature": 0.1,
        "top_p": 0,
        "response_format": {"type": "json_object"},  # força JSON
        "messages": [
            {"role": "system", "content": system_prompt},  # apenas 1 system
            {"role": "user", "content": user_prompt},      # e 1 user
        ],
    }

    r = requests.post(url, headers=headers, data=json.dumps(payload), timeout=60)
    log.debug("OpenRouter status=%s", r.status_code)

    if r.status_code != 200:
        raise RuntimeError(f"OpenRouter HTTP {r.status_code}")

    data = r.json()
    content = (data.get("choices", [{}])[0].get("message", {}).get("content", "") or "")
    content = content.strip().strip("`")

    # caminho feliz: já veio JSON limpo
    try:
        return json.loads(content)
    except Exception:
        pass

    # fallback: extrair o primeiro objeto JSON balanceando chaves
    def _first_json_object(s: str):
        start = s.find("{")
        while start != -1:
            depth = 0
            in_str = False
            esc = False
            for i in range(start, len(s)):
                ch = s[i]
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
                        depth += 1
                    elif ch == "}":
                        depth -= 1
                        if depth == 0:
                            candidate = s[start:i+1]
                            try:
                                return json.loads(candidate)
                            except Exception:
                                break
            start = s.find("{", start + 1)
        return None

    obj = _first_json_object(content)
    if obj is None:
        raise ValueError("Resposta do LLM sem JSON reconhecível.")
    return obj

# ==========================
# Fluxo principal
# ==========================

def ensure_reply_prefix(subject: str) -> str:
    if not subject:
        return "Re:"
    s = subject.strip()
    if s.lower().startswith("re:"):
        return s
    return f"Re: {s}"


def build_user_prompt(original_subject: str, body_text: str, cobol_files: list) -> str:
    # Monta o prompt para o modelo
    parts = []
    parts.append(f"Assunto original: {original_subject}")
    parts.append("\nCorpo do e-mail do aluno:\n" + (body_text or "(vazio)"))

    if cobol_files:
        parts.append("\nAnexos COBOL (até 80KB cada, apenas resumo):")
        for (fn, snip) in cobol_files:
            parts.append(f"--- {fn} ---\n{snip}\n")

    parts.append(
        "\nSua tarefa: decida se dá para responder ou se deve escalar. "
        "Se der para responder, produza resposta objetiva e educada, com observações sobre o código COBOL quando houver. "
        "Use URLs em texto puro nas chamadas para ação."
    )

    return "\n".join(parts)


def build_outgoing_body(corpo_markdown: str) -> str:
    # adiciona assinatura institucional; não mexe nas URLs (texto puro)
    lines = [corpo_markdown.strip(), "", SIGNATURE_FOOTER.strip(), SIGNATURE_LINKS.strip(), "", f"— {SIGNATURE_NAME.strip()} "]
    # normaliza quebras
    out = "\n".join(lines)
    out = out.replace("\r\n", "\n").replace("\r", "\n")
    # elimina \n extras consecutivos
    out = re.sub(r"\n{3,}", "\n\n", out)
    return out


def send_email_reply(original_msg, to_addr: str, subject: str, body_text: str) -> bytes:
    msg = EmailMessage()
    msg["From"] = MAIL_USER
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)

    # threading headers
    orig_msgid = original_msg.get("Message-ID")
    if orig_msgid:
        msg["In-Reply-To"] = orig_msgid
        refs = original_msg.get_all("References", [])
        ref_line = " ".join(refs + [orig_msgid]) if refs else orig_msgid
        msg["References"] = ref_line

    msg.set_content(body_text)

    # Envio SMTP
    if SMTP_TLS_MODE == "ssl":
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as s:
            s.set_debuglevel(SMTP_DEBUG)
            s.login(MAIL_USER, MAIL_PASS)
            s.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.set_debuglevel(SMTP_DEBUG)
            s.ehlo()
            s.starttls(context=ssl.create_default_context())
            s.ehlo()
            s.login(MAIL_USER, MAIL_PASS)
            s.send_message(msg)

    # retorna bytes crus para APPEND no Sent
    return msg.as_bytes()


def ensure_mailbox(imap: imaplib.IMAP4_SSL, box: str):
    try:
        imap.create(box)
    except Exception:
        # alguns servidores retornam NO [ALREADYEXISTS]
        log.debug("Mailbox pode já existir: %s", box)


def move_message(imap: imaplib.IMAP4_SSL, msg_seq: bytes, dest_box: str):
    ensure_mailbox(imap, dest_box)
    imap.copy(msg_seq, dest_box)
    if EXPUNGE_AFTER_COPY:
        imap.store(msg_seq, "+FLAGS", r"(\Deleted)")
        imap.expunge()


def append_to_sent(raw_bytes: bytes):
    # Faz um login rápido só para APPEND (isola falhas)
    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
        imap.login(MAIL_USER, MAIL_PASS)
        ensure_mailbox(imap, SENT_FOLDER)
        # IMPORTANTE: deixe o servidor definir a data (None)
        imap.append(SENT_FOLDER, r"(\Seen)", None, raw_bytes)
        imap.logout()
        log.info("Mensagem copiada para a pasta de enviados: %s", SENT_FOLDER)


def decide_and_respond(imap: imaplib.IMAP4_SSL, msg_seq: bytes, msg_bytes: bytes):
    msg = BytesParser(policy=policy.default).parsebytes(msg_bytes)
    sender = decode_mime_words(msg.get("From"))
    from_addr = re.search(r"<([^>]+)>", sender)
    to_reply = from_addr.group(1) if from_addr else sender

    original_subject = decode_mime_words(msg.get("Subject"))
    body_text = extract_text_body(msg)
    cobol_files = extract_cobol_attachments(msg)

    log.debug(
        "Email de %s / subj='%s' / anexos=%d",
        to_reply,
        original_subject,
        len(cobol_files),
    )

    # Monta prompt e chama LLM
    user_prompt = build_user_prompt(original_subject, body_text, cobol_files)

    try:
        llm_json = call_openrouter(SYSTEM_PROMPT, user_prompt)
    except Exception as e:
        log.error("LLM error")
        log.exception(e)
        # Escala o e-mail original
        move_message(imap, msg_seq, _safe_box(FOLDER_ESCALATE))
        log.info("E-mail movido para %s", _safe_box(FOLDER_ESCALATE))
        return

    acao = llm_json.get("acao", "escalar")
    nivel = float(llm_json.get("nivel_confianca", 0) or 0)
    assunto_model = llm_json.get("assunto", original_subject or "")
    corpo_markdown = llm_json.get("corpo_markdown", "")

    # Política de confiança
    should_answer = (acao == "responder") and (nivel >= CONFIDENCE_THRESHOLD)

    if not should_answer:
        move_message(imap, msg_seq, _safe_box(FOLDER_ESCALATE))
        log.info("Decisão do modelo: acao=%s conf=%.2f → escalado", acao, nivel)
        return

    # Ajusta assunto para resposta (mantém original, apenas prefixo Re: se faltar)
    subject_to_send = ensure_reply_prefix(original_subject or assunto_model or "")

    # Monta corpo final com assinatura
    body_out = build_outgoing_body(corpo_markdown)

    # Envia
    raw_out = send_email_reply(msg, to_reply, subject_to_send, body_out)
    log.info("E-mail enviado para %s (Subject: %s)", to_reply, subject_to_send)

    # Copia para enviados
    try:
        ensure_mailbox(imap, SENT_FOLDER)
        # usar sessão separada ajuda, mas mantemos aqui por compat
        append_to_sent(raw_out)
    except Exception as e:
        log.error("Falha ao copiar para enviados")
        log.exception(e)

    # Move original para Respondidos
    move_message(imap, msg_seq, _safe_box(FOLDER_PROCESSED))
    log.info("E-mail movido para %s", _safe_box(FOLDER_PROCESSED))


# ==========================
# Watcher IMAP
# ==========================

def watch_imap_loop():
    log.info("Watcher IMAP — envio via SMTP %s", SMTP_HOST)
    log.info("App público em: %s", APP_PUBLIC_URL)
    while True:
        try:
            with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
                # Pré-login caps
                typ, caps = imap.capability()
                if typ == "OK" and caps:
                    try:
                        log.debug("CAPABILITIES: %s", " ".join([c.decode() if isinstance(c, bytes) else c for c in caps]))
                    except Exception:
                        pass

                imap.login(MAIL_USER, MAIL_PASS)
                imap.select("INBOX")

                typ, data = imap.search(None, "UNSEEN")
                unseen = data[0].split() if data and data[0] else []
                log.debug("UNSEEN: %s", unseen)

                for num in unseen:
                    typ, msg_data = imap.fetch(num, "(RFC822)")
                    if typ != "OK":
                        continue
                    msg_bytes = msg_data[0][1]
                    decide_and_respond(imap, num, msg_bytes)

                # limpeza opcional pós-processamento
                if EXPUNGE_AFTER_COPY:
                    try:
                        imap.expunge()
                    except Exception:
                        pass

                imap.logout()
        except Exception as e:
            log.exception(e)
        time.sleep(CHECK_INTERVAL_SECONDS)


# ==========================
# Flask app (diagnóstico)
# ==========================
app = Flask(__name__)

@app.route("/")
def index():
    return f"<h3>{APP_TITLE}</h3><p>OK</p>"

@app.route("/diag/prompt")
def diag_prompt():
    return jsonify({
        "sha1": SYSTEM_PROMPT_SHA1,
        "first120": SYSTEM_PROMPT[:120],
    })

@app.route("/diag/openrouter-chat")
def diag_openrouter():
    try:
        payload = {
            "model": OPENROUTER_MODEL,
            "max_tokens": 8,
            "temperature": 0.1,
            "top_p": 0,
            "response_format": {"type": "json_object"},
            "messages": [
                {\1'Reponda somente JSON: {"ok":true}'ok\\":true}"},
                {"role": "user", "content": "ping"},
            ],
        }
        headers = {
            "Authorization": f"Bearer {OPENROUTER_API_KEY}",
            "Content-Type": "application/json",
        }
        r = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, data=json.dumps(payload), timeout=20)
        status = r.status_code
        body = None
        try:
            body = r.json()
        except Exception:
            body = {"text": r.text[:200]}
        return jsonify({"status": status, "body": body})
    except Exception as e:
        return jsonify({"status": 500, "error": str(e)})


if __name__ == "__main__":
    # roda watcher em thread simples (processo único na Render)
    import threading

    t = threading.Thread(target=watch_imap_loop, daemon=True)
    t.start()

    port = int(os.getenv("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)


