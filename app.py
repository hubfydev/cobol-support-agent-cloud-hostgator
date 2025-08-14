#!/usr/bin/env python3 - v10.7
# -*- coding: utf-8 -*-

"""
COBOL Support Agent — IMAP watcher + SMTP sender + OpenRouter
"""

import os
import re
import ssl
import time
import json
import hashlib
import logging
import imaplib
import smtplib

from datetime import datetime, timedelta, timezone

from email import policy
from email.parser import BytesParser
from email.message import EmailMessage
from email.header import decode_header, make_header
from email.utils import formatdate, make_msgid

from flask import Flask, jsonify, request

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
IMAP_FOLDER_INBOX = os.getenv("IMAP_FOLDER_INBOX", "INBOX")
IMAP_FALLBACK_LAST_N = int(os.getenv("IMAP_FALLBACK_LAST_N", "0"))  # 0 = desliga

# >>> NOVOS CONTROLES DE BUSCA IMAP <<<
IMAP_STRICT_UNSEEN_ONLY = os.getenv("IMAP_STRICT_UNSEEN_ONLY", "true").lower() == "true"
IMAP_SINCE_DAYS = int(os.getenv("IMAP_SINCE_DAYS", "0"))  # 0 = sem filtro de data
IMAP_FALLBACK_WHEN_LLM_BLOCKED = os.getenv("IMAP_FALLBACK_WHEN_LLM_BLOCKED", "false").lower() == "true"

# SMTP
SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_TLS_MODE = os.getenv("SMTP_TLS_MODE", "starttls").lower()  # starttls|ssl
SMTP_DEBUG = int(os.getenv("SMTP_DEBUG", "0"))

# LLM / OpenRouter
LLM_BACKEND = os.getenv("LLM_BACKEND", "openrouter")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openrouter/auto")
OPENROUTER_MODEL_FALLBACK = os.getenv("OPENROUTER_MODEL_FALLBACK", "openrouter/auto")
OPENROUTER_MAX_TOKENS = int(os.getenv("OPENROUTER_MAX_TOKENS", "512"))
OPENROUTER_APP_NAME = os.getenv("OPENROUTER_APP_NAME", "COBOL Support Agent")
OPENROUTER_SITE_URL = os.getenv("OPENROUTER_SITE_URL", "")
CONFIDENCE_THRESHOLD = float(os.getenv("CONFIDENCE_THRESHOLD", "0.8"))

# --- LLM robustez extra ---
LLM_COOLDOWN_SECONDS = int(os.getenv("LLM_COOLDOWN_SECONDS", "900"))  # 15 min
LLM_DISABLE_ON_402 = os.getenv("LLM_DISABLE_ON_402", "true").lower() == "true"
LLM_HARD_DISABLE = os.getenv("LLM_HARD_DISABLE", "false").lower() == "true"

# Pré-triagem (regex para pular LLM)
AUTO_ESCALATE_FROM_REGEX = os.getenv(
    "AUTO_ESCALATE_FROM_REGEX",
    r"(?i)(^mailer-daemon@|^postmaster@|^no[-_\. ]?reply@|^2fa@|^noreply@|@hotmart\.com$)"
)
AUTO_ESCALATE_SUBJECT_REGEX = os.getenv(
    "AUTO_ESCALATE_SUBJECT_REGEX",
    r"(?i)(verification code|delivery status|failure notice|bounce|sale made|refund|payment|invoice|assinatura|pagamento atrasado|chargeback)"
)

# Estado global simples do LLM
_llm_block_until_ts = 0.0
_last_llm_error = ""

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
SIGNATURE_FOOTER = SIGNATURE_FOOTER.replace("\\n", "\n")
SIGNATURE_LINKS = os.getenv("SIGNATURE_LINKS", "https://aprendacobol.com.br/assinatura/")

# ==========================
# Prompt do sistema
# ==========================
SYSTEM_PROMPT = (
    "Você é um assistente do time de suporte de um curso de COBOL da Aprenda COBOL. "
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
    "9) Não mude o tema da conversa. Responda ao que foi solicitado, de forma educada e objetiva, sempre como perte de um time (nós). "
    "10) Se faltar informação para compilar/executar, peça os dados mínimos (ex.: amostras de entrada/saída, layout, JCL). "
    "11) No final do 'corpo_markdown', SEMPRE inclua exatamente estas duas linhas (URLs como texto puro, sem markdown de link): "
    "- Nossa Comunidade no Telegram: https://t.me/aprendacobol "
    "- Conheça a Formação Completa de Programador COBOL: https://assinatura.aprendacobol.com.br "
)

SYSTEM_PROMPT_SHA1 = hashlib.sha1(SYSTEM_PROMPT.encode("utf-8")).hexdigest()
log.info("SYSTEM_PROMPT_SHA1=%s (primeiros 120 chars): %s", SYSTEM_PROMPT_SHA1[:12], SYSTEM_PROMPT[:120])

# ==========================
# Helpers LLM
# ==========================
def _llm_is_blocked_now() -> bool:
    return LLM_HARD_DISABLE or (time.time() < _llm_block_until_ts)

def _llm_block(reason: str, seconds: int):
    global _llm_block_until_ts, _last_llm_error
    _last_llm_error = f"{reason} (cooldown {seconds}s)"
    _llm_block_until_ts = time.time() + max(0, seconds)
    log.warning("LLM bloqueado: %s", _last_llm_error)

# ==========================
# Utilitários de e-mail
# ==========================
def _safe_box(name: str) -> str:
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
    text_parts, html_parts = [], []
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

    if html_parts:
        html = "\n\n".join(html_parts)
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

def _post_openrouter(payload):
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": OPENROUTER_SITE_URL or APP_PUBLIC_URL or "",
        "X-Title": OPENROUTER_APP_NAME,
    }
    # timeout mais curto evita travamentos quando a rota está ruim
    return requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers=headers,
        data=json.dumps(payload),
        timeout=30,
    )

def call_openrouter(system_prompt: str, user_prompt: str) -> dict:
    if _llm_is_blocked_now():
        raise RuntimeError("LLM temporarily disabled by cooldown / hard-disable")

    def _make_payload(compat: bool):
        p = {
            "model": OPENROUTER_MODEL,
            "max_tokens": OPENROUTER_MAX_TOKENS,
            "temperature": 0.0,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        }
        # modo “rico” (preferido)
        if not compat:
            p["tools"] = [{
                "type": "function",
                "function": {
                    "name": "compose_email",
                    "description": "Retorne somente os campos exigidos no esquema.",
                    "parameters": {
                        "type": "object",
                        "properties": {
                            "assunto": {"type": "string"},
                            "corpo_markdown": {"type": "string"},
                            "nivel_confianca": {"type": "number"},
                            "acao": {"type": "string", "enum": ["responder", "escalar"]}
                        },
                        "required": ["assunto", "corpo_markdown", "nivel_confianca", "acao"],
                        "additionalProperties": False
                    }
                }
            }]
            p["tool_choice"] = "required"
            p["response_format"] = {"type": "json_object"}
            # evite top_p=0 em provedores que rejeitam 0: comente a linha abaixo se preferir não enviar top_p
            # p["top_p"] = 0
        return p

    # 1) tenta com tools/json_mode
    payload = _make_payload(compat=False)
    r = _post_openrouter(payload)
    log.debug("OpenRouter status=%s", r.status_code)

    # 402 → cooldown (já implementado no seu código)
    if r.status_code == 402 and LLM_DISABLE_ON_402:
        _llm_block("OpenRouter 402 (limite/rota indisponível)", LLM_COOLDOWN_SECONDS)
        raise RuntimeError(f"OpenRouter HTTP {r.status_code}")

    # 400 → refaz em modo compatível (sem tools/response_format/top_p)
    if r.status_code == 400:
        try:
            log.error("OpenRouter 400 body (primeiros 500 chars): %s", r.text[:500])
        except Exception:
            pass
        compat_payload = _make_payload(compat=True)
        r = _post_openrouter(compat_payload)
        log.debug("OpenRouter (compat) status=%s", r.status_code)

    # Fallback só para 404/429/500
    if r.status_code in (404, 429, 500):
        simple_payload = _make_payload(compat=True)
        simple_payload["model"] = OPENROUTER_MODEL_FALLBACK
        r = _post_openrouter(simple_payload)
        log.debug("OpenRouter fallback status=%s", r.status_code)
        if r.status_code == 402 and LLM_DISABLE_ON_402:
            _llm_block("OpenRouter 402 no fallback", LLM_COOLDOWN_SECONDS)
            raise RuntimeError(f"OpenRouter HTTP {r.status_code}")

    if r.status_code != 200:
        try:
            log.error("OpenRouter erro %s, corpo: %s", r.status_code, r.text[:500])
        except Exception:
            pass
        raise RuntimeError(f"OpenRouter HTTP {r.status_code}")

    data = r.json()
    choice = (data.get("choices") or [{}])[0]
    message = choice.get("message") or {}

    # caminho 1: function-calling
    tool_calls = message.get("tool_calls") or []
    if tool_calls:
        try:
            args_str = tool_calls[0]["function"]["arguments"]
            return json.loads(args_str)
        except Exception as e:
            log.debug("Falha ao parsear tool_call.arguments: %s", e)

    # caminho 2: conteúdo normal → extrai 1º objeto JSON
    content = (message.get("content") or "").strip()
    log.debug("LLM raw content (primeiros 400 chars): %s", content[:400])

    m = re.search(r"```(?:json)?\s*({.*})\s*```", content, flags=re.S)
    if m:
        try:
            return json.loads(m.group(1))
        except Exception:
            pass

    content = content.replace("```json", "```").strip("`").strip()

    def _first_json_object(s: str):
        start = s.find("{")
        while start != -1:
            depth, in_str, esc = 0, False, False
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
    if obj is not None:
        return obj

    raise ValueError("Resposta do LLM sem JSON reconhecível.")

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
    lines = [
        corpo_markdown.strip(),
        "",
        SIGNATURE_FOOTER.strip(),
        SIGNATURE_LINKS.strip(),
        "",
        f"— {SIGNATURE_NAME.strip()} ",
    ]
    out = "\n".join(lines)
    out = out.replace("\r\n", "\n").replace("\r", "\n")
    out = re.sub(r"\n{3,}", "\n\n", out)
    return out

def send_email_reply(original_msg, to_addr: str, subject: str, body_text: str) -> bytes:
    msg = EmailMessage()
    msg["From"] = MAIL_USER
    msg["Sender"] = MAIL_USER
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=(MAIL_USER.split("@", 1)[-1] if MAIL_USER and "@" in MAIL_USER else None))
    msg["X-Mailer"] = "COBOL Support Agent"

    if original_msg:
        try:
            orig_msgid = original_msg.get("Message-ID")
        except Exception:
            orig_msgid = None
        if orig_msgid:
            msg["In-Reply-To"] = orig_msgid
            refs = original_msg.get_all("References", [])
            ref_line = " ".join(refs + [orig_msgid]) if refs else orig_msgid
            msg["References"] = ref_line

    msg.set_content(body_text)

    def _send(smtp):
        smtp.set_debuglevel(SMTP_DEBUG)
        smtp.ehlo()
        if SMTP_TLS_MODE == "starttls":
            smtp.starttls(context=ssl.create_default_context())
            smtp.ehlo()
        smtp.login(MAIL_USER, MAIL_PASS)
        refused = smtp.sendmail(MAIL_USER, [to_addr], msg.as_string())
        if refused:
            log.error("SMTP recusou destinatários: %s", refused)
            raise RuntimeError(f"SMTP refused {refused}")

    if SMTP_TLS_MODE == "ssl":
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ssl.create_default_context()) as s:
            _send(s)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            _send(s)

    return msg.as_bytes()

def ensure_mailbox(imap: imaplib.IMAP4_SSL, box: str):
    try:
        imap.create(box)
    except Exception:
        log.debug("Mailbox pode já existir: %s", box)

def move_message_uid(imap: imaplib.IMAP4_SSL, msg_uid: bytes, dest_box: str):
    ensure_mailbox(imap, dest_box)
    imap.uid("COPY", msg_uid, dest_box)
    imap.uid("STORE", msg_uid, "+FLAGS", r"(\Deleted)")

def append_to_sent(raw_bytes: bytes):
    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
        imap.login(MAIL_USER, MAIL_PASS)
        ensure_mailbox(imap, SENT_FOLDER)
        imap.append(SENT_FOLDER, r"(\Seen)", None, raw_bytes)
        imap.logout()
        log.info("Mensagem copiada para a pasta de enviados: %s", SENT_FOLDER)

def decide_and_respond(imap: imaplib.IMAP4_SSL, msg_uid: bytes, msg_bytes: bytes):
    msg = BytesParser(policy=policy.default).parsebytes(msg_bytes)
    sender = decode_mime_words(msg.get("From"))
    from_addr = re.search(r"<([^>]+)>", sender)
    to_reply = from_addr.group(1) if from_addr else sender

    original_subject = decode_mime_words(msg.get("Subject"))
    body_text = extract_text_body(msg)
    cobol_files = extract_cobol_attachments(msg)

    # PRE-TRIAGEM (regex)
    try:
        if AUTO_ESCALATE_FROM_REGEX and re.search(AUTO_ESCALATE_FROM_REGEX, (sender or ""), re.I):
            move_message_uid(imap, msg_uid, _safe_box(FOLDER_ESCALATE))
            log.info("Pré-triagem: remetente operacional → escalado (%s)", sender)
            return
        if AUTO_ESCALATE_SUBJECT_REGEX and re.search(AUTO_ESCALATE_SUBJECT_REGEX, (original_subject or ""), re.I):
            move_message_uid(imap, msg_uid, _safe_box(FOLDER_ESCALATE))
            log.info("Pré-triagem: assunto operacional → escalado ('%s')", original_subject)
            return
    except Exception:
        log.debug("Pré-triagem ignorada (regex inválida?)", exc_info=True)

    # CURTO-CIRCUITO: LLM bloqueado
    if _llm_is_blocked_now():
        remaining = max(0, int(_llm_block_until_ts - time.time())) if not LLM_HARD_DISABLE else -1
        log.warning("LLM em cooldown/hard-disable (%ss). Pulando LLM e escalando. De: %s Assunto: %s",
                    remaining, sender, original_subject)
        move_message_uid(imap, msg_uid, _safe_box(FOLDER_ESCALATE))
        return

    log.debug("Email de %s / subj='%s' / anexos=%d", to_reply, original_subject, len(cobol_files))
    user_prompt = build_user_prompt(original_subject, body_text, cobol_files)

    try:
        llm_json = call_openrouter(SYSTEM_PROMPT, user_prompt)
    except Exception:
        log.error("LLM error", exc_info=True)
        move_message_uid(imap, msg_uid, _safe_box(FOLDER_ESCALATE))
        log.info("E-mail movido para %s", _safe_box(FOLDER_ESCALATE))
        return

    acao = llm_json.get("acao", "escalar")
    nivel = float(llm_json.get("nivel_confianca", 0) or 0)
    assunto_model = llm_json.get("assunto", original_subject or "")
    corpo_markdown = llm_json.get("corpo_markdown", "")

    should_answer = (acao == "responder") and (nivel >= CONFIDENCE_THRESHOLD)

    if not should_answer:
        move_message_uid(imap, msg_uid, _safe_box(FOLDER_ESCALATE))
        log.info("Decisão do modelo: acao=%s conf=%.2f → escalado", acao, nivel)
        return

    subject_to_send = ensure_reply_prefix(original_subject or assunto_model or "")
    body_out = build_outgoing_body(corpo_markdown)

    raw_out = send_email_reply(msg, to_reply, subject_to_send, body_out)
    log.info("E-mail enviado para %s (Subject: %s)", to_reply, subject_to_send)

    try:
        ensure_mailbox(imap, SENT_FOLDER)
        append_to_sent(raw_out)
    except Exception:
        log.error("Falha ao copiar para enviados", exc_info=True)

    move_message_uid(imap, msg_uid, _safe_box(FOLDER_PROCESSED))
    log.info("E-mail movido para %s", _safe_box(FOLDER_PROCESSED))

# ==========================
# Watcher IMAP (por UID)
# ==========================
def _imap_uid_search(imap, criteria: str):
    typ, data = imap.uid("search", None, criteria)
    return data[0].split() if typ == "OK" and data and data[0] else []

def _select_box(imap, box: str):
    typ, data = imap.select(box)
    if typ == "OK" and data:
        try:
            exists = int(data[0])
        except Exception:
            exists = data[0].decode() if isinstance(data[0], bytes) else data[0]
        log.debug("SELECT %s → EXISTS=%s", box, exists)
    else:
        log.warning("Falha ao selecionar caixa %s: %s %s", box, typ, data)

def _imap_since_date(days: int) -> str:
    """Retorna data IMAP 'DD-Mon-YYYY' em inglês (ex.: 13-Aug-2025)."""
    dt = datetime.now(timezone.utc) - timedelta(days=days)
    mon = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][dt.month - 1]
    return f"{dt.day:02d}-{mon}-{dt.year}"

def _imap_search_unseen(imap, since_days: int) -> list[bytes]:
    if since_days and since_days > 0:
        date_str = _imap_since_date(since_days)
        return _imap_uid_search(imap, f"UNSEEN SINCE {date_str}")
    return _imap_uid_search(imap, "UNSEEN")

def watch_imap_loop():
    log.info("Watcher IMAP — envio via SMTP %s", SMTP_HOST)
    log.info("App público em: %s", APP_PUBLIC_URL)
    log.info(
        "IMAP_STRICT_UNSEEN_ONLY=%s | IMAP_SINCE_DAYS=%d | IMAP_FALLBACK_LAST_N=%d | IMAP_FALLBACK_WHEN_LLM_BLOCKED=%s",
        IMAP_STRICT_UNSEEN_ONLY, IMAP_SINCE_DAYS, IMAP_FALLBACK_LAST_N, IMAP_FALLBACK_WHEN_LLM_BLOCKED
    )
    while True:
        try:
            with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
                typ, caps = imap.capability()
                if typ == "OK" and caps:
                    try:
                        log.debug("CAPABILITIES: %s", " ".join([c.decode() if isinstance(c, bytes) else c for c in caps]))
                    except Exception:
                        pass

                imap.login(MAIL_USER, MAIL_PASS)
                _select_box(imap, IMAP_FOLDER_INBOX)

                # 1) Sempre tente UNSEEN (com SINCE opcional)
                uids = _imap_search_unseen(imap, IMAP_SINCE_DAYS)

                llm_blocked = _llm_is_blocked_now()

                # 2) Se modo estrito: NÃO tenta NEW/RECENT/FALLBACK
                if not uids and not IMAP_STRICT_UNSEEN_ONLY:
                    # NEW/RECENT apenas se LLM não estiver bloqueado
                    if not uids and not llm_blocked:
                        recent = _imap_uid_search(imap, "NEW")
                        if recent:
                            log.debug("NEW (UIDs): %s", recent)
                            uids = recent

                    if not uids and not llm_blocked:
                        recent = _imap_uid_search(imap, "RECENT")
                        if recent:
                            log.debug("RECENT (UIDs): %s", recent)
                            uids = recent

                    # Fallback ALL → últimos N, respeitando o flag quando LLM bloqueado
                    if (not uids) and IMAP_FALLBACK_LAST_N > 0 and (not llm_blocked or IMAP_FALLBACK_WHEN_LLM_BLOCKED):
                        all_uids = _imap_uid_search(imap, "ALL")
                        tail = all_uids[-IMAP_FALLBACK_LAST_N:] if all_uids else []
                        if tail:
                            log.warning(
                                "UNSEEN/NEW/RECENT vazios — usando últimos %d UIDs: %s",
                                IMAP_FALLBACK_LAST_N, tail
                            )
                            uids = tail

                log.debug("UIDs a processar: %s", uids)

                # Processa do menor para o maior
                for uid in uids:
                    if not uid or not uid.strip():
                        continue
                    try:
                        # BODY.PEEK[] evita marcar \Seen durante o fetch
                        typ, msg_data = imap.uid("fetch", uid, "(BODY.PEEK[])")
                        if typ != "OK" or not msg_data or not msg_data[0]:
                            log.warning("Falha no FETCH UID=%s: %s %s", uid, typ, msg_data)
                            continue
                        # Alguns servidores retornam (b'UID ... RFC822 {len}', b'rawbytes')
                        raw_bytes = msg_data[0][1] if isinstance(msg_data[0], tuple) else msg_data[1]
                        decide_and_respond(imap, uid, raw_bytes)
                    except Exception:
                        log.exception("Erro ao processar UID=%s", uid)

                if EXPUNGE_AFTER_COPY:
                    try:
                        imap.expunge()
                    except Exception:
                        log.exception("Falha no expunge final")

                imap.logout()
        except Exception:
            log.exception("Loop IMAP falhou")
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

# Adicione perto das outras rotas /diag/*
@app.route("/diag/llm/unblock", methods=["POST"])
def diag_llm_unblock():
    global _llm_block_until_ts, _last_llm_error
    _llm_block_until_ts = 0.0
    _last_llm_error = ""
    return jsonify({"ok": True, "unblocked": True})

def _post_openrouter_diag(payload):
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": OPENROUTER_SITE_URL or APP_PUBLIC_URL or "",
        "X-Title": OPENROUTER_APP_NAME,
    }
    return requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers=headers, data=json.dumps(payload), timeout=20
    )

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
                {"role": "system", "content": "Responda SOMENTE JSON válido."},
                {"role": "user", "content": "Retorne {\"ok\": true}."},
            ],
        }
        r = _post_openrouter_diag(payload)
        status = r.status_code
        try:
            body = r.json()
        except Exception:
            body = {"text": r.text[:200]}
        return jsonify({"status": status, "body": body})
    except Exception as e:
        return jsonify({"status": 500, "error": str(e)})

@app.route("/diag/smtp")
def diag_smtp():
    to = request.args.get("to", MAIL_USER)
    subject = "Teste SMTP — COBOL Support Agent"
    body = "Olá! Este é um teste de envio SMTP direto do /diag/smtp.\n\n— Sistema"
    try:
        raw = send_email_reply(None, to, subject, body)
        append_to_sent(raw)
        return jsonify({"ok": True, "to": to})
    except Exception as e:
        log.exception("Falha no /diag/smtp")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.route("/diag/imap")
def diag_imap():
    box = request.args.get("box", IMAP_FOLDER_INBOX)
    n = int(request.args.get("n", "20"))
    try:
        with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
            imap.login(MAIL_USER, MAIL_PASS)
            typ, _ = imap.select(box)
            if typ != "OK":
                return jsonify({"error": f"não foi possível selecionar {box}"}), 500

            def _count(q):
                ids = _imap_uid_search(imap, q)
                tail = [i.decode() if isinstance(i, bytes) else i for i in ids[-n:]]
                return len(ids), tail

            c_unseen, ids_unseen = _count("UNSEEN")
            c_new, ids_new = _count("NEW")
            c_recent, ids_recent = _count("RECENT")
            c_all, ids_all = _count("ALL")

            imap.logout()

        return jsonify({
            "box": box,
            "counts": {"UNSEEN": c_unseen, "NEW": c_new, "RECENT": c_recent, "ALL": c_all},
            "tail_uids": {"UNSEEN": ids_unseen, "NEW": ids_new, "RECENT": ids_recent, "ALL": ids_all},
            "strict_unseen_only": IMAP_STRICT_UNSEEN_ONLY,
            "since_days": IMAP_SINCE_DAYS,
            "fallback_last_n": IMAP_FALLBACK_LAST_N,
            "fallback_when_llm_blocked": IMAP_FALLBACK_WHEN_LLM_BLOCKED,
        })
    except Exception as e:
        log.exception("diag/imap falhou")
        return jsonify({"error": str(e)}), 500

@app.route("/diag/llm")
def diag_llm():
    remaining = max(0, int(_llm_block_until_ts - time.time()))
    return jsonify({
        "blocked": _llm_is_blocked_now(),
        "block_expires_in_seconds": remaining if not LLM_HARD_DISABLE else None,
        "hard_disable": LLM_HARD_DISABLE,
        "last_error": _last_llm_error,
        "model": OPENROUTER_MODEL,
        "fallback": OPENROUTER_MODEL_FALLBACK,
        "cooldown_seconds": LLM_COOLDOWN_SECONDS,
        "disable_on_402": LLM_DISABLE_ON_402,
    })

if __name__ == "__main__":
    import threading
    t = threading.Thread(target=watch_imap_loop, daemon=True)
    t.start()

    port = int(os.getenv("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
