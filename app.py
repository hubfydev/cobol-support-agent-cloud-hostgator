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

# app.py
import os
import re
import ssl
import json
import time
import html
import email
import imaplib
import smtplib
import logging
import threading
import traceback
from typing import Optional, Tuple, Dict, Any
from email.message import EmailMessage
from email.policy import default as default_policy
from email.header import decode_header, make_header
from email.utils import parsedate_to_datetime, formatdate, make_msgid

from flask import Flask, jsonify

# =========================
# Config / Env
# =========================

def _get_log_level():
    lvl = os.getenv("LOG_LEVEL", "INFO").strip().upper()
    return getattr(logging, lvl, logging.INFO)

logging.basicConfig(
    level=_get_log_level(),
    format="%(asctime)s [%(levelname)s] %(message)s",
)

if logging.getLogger().level <= logging.DEBUG:
    try:
        imaplib.Debug = 4
    except Exception:
        pass

APP_TITLE = os.getenv("APP_TITLE", "COBOL Support Agent")
APP_PUBLIC_URL = os.getenv("APP_PUBLIC_URL", "http://localhost:10000")

PORT = int(os.getenv("PORT", "10000"))
CHECK_INTERVAL_SECONDS = int(os.getenv("CHECK_INTERVAL_SECONDS", "60"))

# IMAP/SMTP
IMAP_HOST = os.getenv("IMAP_HOST", "")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
MAIL_USER = os.getenv("MAIL_USER", "")
MAIL_PASS = os.getenv("MAIL_PASS", "")

SMTP_HOST = os.getenv("SMTP_HOST", IMAP_HOST)
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_TLS_MODE = os.getenv("SMTP_TLS_MODE", "starttls").strip().lower()  # ssl | starttls | plain

# Pastas
FOLDER_ESCALATE = os.getenv("FOLDER_ESCALATE", "Escalar")
FOLDER_PROCESSED = os.getenv("FOLDER_PROCESSED", "Respondidos")
SENT_FOLDER = os.getenv("SENT_FOLDER", "INBOX.Sent")

EXPUNGE_AFTER_COPY = os.getenv("EXPUNGE_AFTER_COPY", "true").strip().lower() in {"1", "true", "yes", "y"}

# LLM
LLM_BACKEND = os.getenv("LLM_BACKEND", "openrouter").strip().lower()
CONFIDENCE_THRESHOLD = float(os.getenv("CONFIDENCE_THRESHOLD", "0.5"))

# Re-ask config
REASK_ON_ESCALAR = os.getenv("REASK_ON_ESCALAR", "true").strip().lower() in {"1","true","yes","y"}
REASK_ESCALAR_MIN = float(os.getenv("REASK_ESCALAR_MIN", "0.55"))

# OpenRouter
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openrouter/auto")
OPENROUTER_MAX_TOKENS = int(os.getenv("OPENROUTER_MAX_TOKENS", "256"))
OPENROUTER_SITE_URL = os.getenv("OPENROUTER_SITE_URL", APP_PUBLIC_URL)
OPENROUTER_APP_NAME = os.getenv("OPENROUTER_APP_NAME", APP_TITLE)

# Ollama (opcional)
OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "llama3.1:8b")

# Assinatura
SIGNATURE_NAME = os.getenv("SIGNATURE_NAME", "Equipe Aprenda COBOL — Suporte")
SIGNATURE_FOOTER = os.getenv(
    "SIGNATURE_FOOTER",
    "Se precisar, responda este e-mail com mais detalhes ou anexe seu arquivo .COB/.CBL.\n"
    "Horário de atendimento: 9h–18h (ET), seg–sex. Conheça nossa Formação Completa de Programadores COBOL, "
    "com COBOL Avançado, JCL, Db2 e Bancos de Dados completo em:"
)
SIGNATURE_LINKS = os.getenv("SIGNATURE_LINKS", "https://aprendacobol.com.br/assinatura/")

# =========================
# Prompts
# =========================

SYSTEM_PROMPT_BASE = (
    "Você é um assistente de suporte de um curso de COBOL. "
    "SEMPRE produza um JSON VÁLIDO e nada além disso. "
    "Formato do JSON (minificado, sem comentários, sem markdown, sem texto extra): "
    '{"assunto":"...","corpo_markdown":"...","nivel_confianca":0.0,"acao":"responder|escalar"} '
    "Regras: "
    "1) NUNCA inclua crases ou ``` no output. "
    "2) NUNCA acrescente explicações fora do JSON. "
    "3) Use exatamente as chaves do esquema. "
    "4) PT-BR no corpo. "
    "5) 'nivel_confianca' entre 0 e 1. "
    "6) PRIORIZE 'responder' quando houver uma orientação útil; "
    "   use 'escalar' APENAS quando faltar contexto essencial, for cobrança/conta, ou for altamente técnico sem dados. "
    "7) Se pedido estiver claro e respondível, 'acao'='responder' com nivel_confianca>=0.8; "
    "   se ambíguo/incompleto a ponto de não dar para orientar, 'acao'='escalar' (nivel_confianca<=0.6). "
    "Inclua orientações de COBOL (DIVISION, SECTION, PIC, níveis, I/O, SQLCA etc.) quando fizer sentido. "
    "Sugira passos práticos. "
    "Se o e-mail pedir link do Telegram, inclua https://t.me/aprendacobol. "
    "Sempre ofereça a Formação Completa de Programador COBOL por assinatura em: https://assinatura.aprendacobol.com.br."
)

SYSTEM_PROMPT_NUDGE = (
    SYSTEM_PROMPT_BASE +
    " IMPORTANTE: Procure responder. Se conseguir montar orientação inicial, devolva 'acao'='responder'."
)

# =========================
# Markdown helpers (links sem duplicar + HTML)
# =========================

LINK_MD_RE   = re.compile(r'$begin:math:display$([^$end:math:display$]+)\]$begin:math:text$(https?://[^\\s)]+)$end:math:text$')
AUTO_LINK_RE = re.compile(r'<(https?://[^>]+)>')
RAW_URL_RE   = re.compile(r'(?<![">])(https?://[^\s<)]+)')

def _normalize_url(u: str) -> str:
    return u.strip().rstrip('/')

def md_links_to_text(md: str) -> str:
    if not md:
        return ""
    s = md
    def repl_link(m):
        txt = m.group(1).strip()
        url = m.group(2).strip()
        if _normalize_url(txt).lower() == _normalize_url(url).lower():
            return url
        return f"{txt} ({url})"
    s = LINK_MD_RE.sub(repl_link, s)
    s = AUTO_LINK_RE.sub(lambda m: m.group(1), s)
    s = s.replace("**", "").replace("__", "")
    s = re.sub(r"(^|[^`])`([^`]+)`", r"\1\2", s)
    return s

def _escape_outside_anchors(s: str) -> str:
    parts = re.split(r'(<a\b.*?>.*?</a>)', s, flags=re.I | re.S)
    out = []
    for i, part in enumerate(parts):
        out.append(part if i % 2 == 1 else html.escape(part))
    return ''.join(out)

def md_to_html(md: str) -> str:
    if not md:
        return "<!doctype html><html><body></body></html>"
    s = md
    def repl_link(m):
        txt = html.escape(m.group(1).strip())
        url = html.escape(m.group(2).strip(), quote=True)
        return f'<a href="{url}">{txt}</a>'
    s = LINK_MD_RE.sub(repl_link, s)
    s = AUTO_LINK_RE.sub(lambda m: f'<a href="{html.escape(m.group(1), quote=True)}">{html.escape(m.group(1))}</a>', s)
    def autolink_raw(segment: str) -> str:
        return RAW_URL_RE.sub(lambda m: f'<a href="{html.escape(m.group(1), quote=True)}">{html.escape(m.group(1))}</a>', segment)
    parts = re.split(r'(<a\b.*?>.*?</a>)', s, flags=re.I | re.S)
    s = ''.join(part if i % 2 == 1 else autolink_raw(part) for i, part in enumerate(parts))
    s = _escape_outside_anchors(s)
    s = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"__([^_]+)__", r"<strong>\1</strong>", s)
    s = re.sub(r"(?<!\*)\*([^*]+)\*(?!\*)", r"<em>\1</em>", s)
    s = re.sub(r"(?<!_)_([^_]+)_(?!_)", r"<em>\1</em>", s)
    s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "<br>")
    return f"<!doctype html><html><body>{s}</body></html>"

# =========================
# Utils
# =========================

def make_signature() -> str:
    # Converte "\n" literal em quebras de linha reais
    footer = (SIGNATURE_FOOTER or "").replace("\\n", "\n").strip()
    link = (SIGNATURE_LINKS or "").strip()

    # Se o footer terminar com "em:", coloca o link na MESMA linha
    if footer and link and footer.rstrip().endswith(":"):
        footer = footer.rstrip() + " " + link
        link = ""  # já incorporado

    lines = []
    if footer:
        lines.append(footer)
    if link:
        lines.append(link)
    if SIGNATURE_NAME:
        lines.append(f"\n— {SIGNATURE_NAME}")

    return "\n\n" + "\n".join(lines) + "\n"

def decode_str(v: Optional[str]) -> str:
    if not v:
        return ""
    try:
        return str(make_header(decode_header(v)))
    except Exception:
        return v

def msg_get_addresses(msg, header: str) -> str:
    return decode_str(msg.get(header, ""))

def is_our_own_message(msg) -> bool:
    from_addr = msg_get_addresses(msg, "From").lower()
    return MAIL_USER.lower() in from_addr

def find_first_text_part(msg) -> Tuple[str, str]:
    plain, htmlp = "", ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get("Content-Disposition") or "").lower()
            if ctype == "text/plain" and "attachment" not in disp and not plain:
                try:
                    plain = part.get_content().strip()
                except Exception:
                    payload = part.get_payload(decode=True) or b""
                    plain = payload.decode(part.get_content_charset() or "utf-8", "ignore").strip()
            elif ctype == "text/html" and "attachment" not in disp and not htmlp:
                try:
                    htmlp = part.get_content().strip()
                except Exception:
                    payload = part.get_payload(decode=True) or b""
                    htmlp = payload.decode(part.get_content_charset() or "utf-8", "ignore").strip()
    else:
        ctype = msg.get_content_type()
        if ctype == "text/plain":
            plain = msg.get_content().strip()
        elif ctype == "text/html":
            htmlp = msg.get_content().strip()
    return plain, htmlp

def clamp_conf(v: Any) -> float:
    try:
        x = float(v)
    except Exception:
        return 0.0
    return max(0.0, min(1.0, x))

def safe_json_extract(s: str) -> Optional[Dict[str, Any]]:
    if not s:
        return None
    s = s.strip()
    if s.startswith("```"):
        s = re.sub(r"^```(\w+)?", "", s).strip()
        if s.endswith("```"):
            s = s[:-3].strip()
    try:
        obj = json.loads(s)
        if isinstance(obj, dict):
            return obj
    except Exception:
        pass
    opens = [i for i, ch in enumerate(s) if ch == "{"]
    for i in opens:
        depth = 0
        for j in range(i, len(s)):
            if s[j] == "{": depth += 1
            elif s[j] == "}":
                depth -= 1
                if depth == 0:
                    chunk = s[i:j+1]
                    try:
                        obj = json.loads(chunk)
                        if isinstance(obj, dict):
                            return obj
                    except Exception:
                        break
    return None

# =========================
# Regras simples antes do LLM
# =========================

def rule_based_autoreply(subject: str, plain: str, htmlp: str) -> Optional[Dict[str, Any]]:
    text = " ".join([subject or "", plain or "", re.sub("<[^>]+>", " ", htmlp or "")]).lower()

    # Telegram
    if any(k in text for k in ["telegram", "grupo do telegram", "link do telegram", "t.me/aprendacobol"]):
        corpo = (
            "Claro! Para entrar no nosso grupo de alunos no Telegram, use o link: "
            "[https://t.me/aprendacobol](https://t.me/aprendacobol)\n\n"
            "Fique à vontade para postar dúvidas de exercícios e projetos lá. Também recomendo a "
            "[Formação Completa de Programador COBOL](https://assinatura.aprendacobol.com.br) para aprofundar."
        )
        return {"acao":"responder","nivel_confianca":0.95,"assunto":"Re: Grupo do Telegram","corpo_markdown":corpo}

    # Assinatura/curso
    if any(k in text for k in ["assinatura", "assinar", "formação completa", "formacao completa", "curso", "plano", "preço", "valor"]):
        corpo = (
            "Legal! Nossa Formação Completa de Programador COBOL está disponível por assinatura em "
            "[https://assinatura.aprendacobol.com.br](https://assinatura.aprendacobol.com.br).\n\n"
            "Você encontra COBOL básico ao avançado, JCL, Db2 e projetos práticos."
        )
        subj = "Re: Assinatura / Formação COBOL"
        return {"acao":"responder","nivel_confianca":0.9,"assunto":subj,"corpo_markdown":corpo}

    return None

# =========================
# LLM Backends
# =========================

def llm_decide(payload: Dict[str, Any], nudge: bool=False) -> Dict[str, Any]:
    if LLM_BACKEND == "openrouter":
        return llm_openrouter(payload, nudge=nudge)
    elif LLM_BACKEND == "ollama":
        return llm_ollama(payload, nudge=nudge)
    else:
        logging.warning("LLM_BACKEND desconhecido: %s; usando openrouter", LLM_BACKEND)
        return llm_openrouter(payload, nudge=nudge)

def _normalize_decision(obj: Dict[str, Any]) -> Dict[str, Any]:
    acao = obj.get("acao") if obj.get("acao") in ("responder","escalar") else "escalar"
    nivel = clamp_conf(obj.get("nivel_confianca", 0))
    assunto = str(obj.get("assunto","")).strip()
    corpo = str(obj.get("corpo_markdown","")).strip()
    return {"acao":acao,"nivel_confianca":nivel,"assunto":assunto,"corpo_markdown":corpo}

def llm_openrouter(payload: Dict[str, Any], nudge: bool=False) -> Dict[str, Any]:
    import requests
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "HTTP-Referer": OPENROUTER_SITE_URL,
        "X-Title": OPENROUTER_APP_NAME,
        "Content-Type": "application/json",
    }
    system = SYSTEM_PROMPT_NUDGE if nudge else SYSTEM_PROMPT_BASE
    messages = [
        {"role":"system","content":system},
        {"role":"user","content":payload.get("input","")}
    ]
    body = {
        "model": OPENROUTER_MODEL,
        "messages": messages,
        "max_tokens": OPENROUTER_MAX_TOKENS,
        "temperature": 0.2,
        "response_format": {"type":"json_object"},
    }
    try:
        resp = requests.post("https://openrouter.ai/api/v1/chat/completions",
                             headers=headers, json=body, timeout=60)
        logging.debug("OpenRouter status=%s", resp.status_code)
        data = resp.json()
        content = (data.get("choices") or [{}])[0].get("message",{}).get("content","")
        obj = safe_json_extract(content) or {}
        return _normalize_decision(obj)
    except Exception as e:
        logging.warning("Falha no OpenRouter: %s", e)
        return {"acao":"escalar","nivel_confianca":0.0,"assunto":"","corpo_markdown":""}

def llm_ollama(payload: Dict[str, Any], nudge: bool=False) -> Dict[str, Any]:
    import requests
    system = SYSTEM_PROMPT_NUDGE if nudge else SYSTEM_PROMPT_BASE
    prompt = (
        system + "\n\n---\nE-MAIL A SEGUIR:\n" +
        payload.get("input","") +
        "\n---\nResponda APENAS com o JSON minificado."
    )
    body = {"model": OLLAMA_MODEL, "prompt": prompt, "options":{"temperature":0.2}, "stream": False}
    try:
        resp = requests.post(f"{OLLAMA_HOST}/api/generate", json=body, timeout=60)
        data = resp.json()
        content = data.get("response","")
        obj = safe_json_extract(content) or {}
        return _normalize_decision(obj)
    except Exception as e:
        logging.warning("Falha no Ollama: %s", e)
        return {"acao":"escalar","nivel_confianca":0.0,"assunto":"","corpo_markdown":""}

# =========================
# IMAP helpers
# =========================

def imap_connect() -> imaplib.IMAP4_SSL:
    ctx = ssl.create_default_context()
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx)
    try:
        typ, data = imap.capability()
        if typ == 'OK' and data:
            caps = data[0].decode().upper().split()
            logging.debug("CAPABILITIES: %s", " ".join([c for c in caps if c != "CAPABILITY"]))
    except Exception:
        pass
    imap.login(MAIL_USER, MAIL_PASS)
    return imap

def ensure_mailbox(imap: imaplib.IMAP4_SSL, name: str):
    try:
        typ, _ = imap.create(name)
        if typ == "OK":
            logging.info("Mailbox criada: %s", name)
    except Exception:
        pass

def move_message(imap: imaplib.IMAP4_SSL, msg_id: bytes, target_box: str):
    dest = target_box
    if not target_box.upper().startswith("INBOX"):
        dest = f"INBOX.{target_box}"
    ensure_mailbox(imap, dest)
    imap.copy(msg_id, dest)
    imap.store(msg_id, "+FLAGS", r"(\Deleted)")
    if EXPUNGE_AFTER_COPY:
        imap.expunge()

# =========================
# SMTP + Append em Enviados
# =========================

def smtp_send_and_append(to_addr: str, subject: str, body_markdown: str, orig_msg: email.message.Message):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = MAIL_USER
    msg["To"] = to_addr

    orig_msgid = (orig_msg.get("Message-ID") or "").strip()
    if orig_msgid:
        msg["In-Reply-To"] = orig_msgid
        refs = (orig_msg.get("References") or "").strip()
        msg["References"] = (refs + " " + orig_msgid).strip()

    msg["Message-ID"] = make_msgid()
    msg["Date"] = formatdate(localtime=True)

    full_md = body_markdown.strip() + make_signature()
    plain_text = md_links_to_text(full_md)
    html_body  = md_to_html(full_md)

    msg.set_content(plain_text, subtype="plain", charset="utf-8")
    msg.add_alternative(html_body, subtype="html")

    try:
        if SMTP_TLS_MODE == "ssl":
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as s:
                s.login(MAIL_USER, MAIL_PASS)
                s.send_message(msg)
        else:
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
                s.ehlo()
                if SMTP_TLS_MODE == "starttls":
                    context = ssl.create_default_context()
                    s.starttls(context=context)
                    s.ehlo()
                s.login(MAIL_USER, MAIL_PASS)
                s.send_message(msg)
        logging.info("E-mail enviado para %s (Subject: %s)", to_addr, subject)
    except Exception as e:
        logging.error("Falha no envio SMTP: %s", e)
        return False, None

    try:
        imap = imap_connect()
        ensure_mailbox(imap, SENT_FOLDER)
        flags = r"(\Seen)"
        date = imaplib.Time2Internaldate(time.time())
        imap.append(SENT_FOLDER, flags, date, msg.as_bytes())
        imap.logout()
        logging.info("Mensagem copiada para a pasta de enviados: %s", SENT_FOLDER)
    except Exception as e:
        logging.warning("Falha ao APPEND em enviados (%s): %s", SENT_FOLDER, e)

    return True, msg["Message-ID"]

# =========================
# Core
# =========================

def assemble_llm_input(orig_msg: email.message.Message, plain: str, htmlp: str) -> str:
    subj = decode_str(orig_msg.get("Subject", ""))
    from_ = decode_str(orig_msg.get("From", ""))
    to_   = decode_str(orig_msg.get("To", ""))
    date_ = decode_str(orig_msg.get("Date", ""))
    lines = [
        f"De: {from_}",
        f"Para: {to_}",
        f"Assunto: {subj}",
        f"Data: {date_}",
        "",
        "==== Corpo (texto): ====",
        plain or "",
        "",
        "==== Corpo (HTML - texto bruto): ====",
        re.sub("<[^>]+>", " ", htmlp or ""),
    ]
    return "\n".join(lines).strip()

def decide_and_act(imap: imaplib.IMAP4_SSL, msg_id: bytes, msg: email.message.Message):
    if is_our_own_message(msg):
        logging.info("Ignorando e-mail enviado por nós mesmos.")
        imap.store(msg_id, "+FLAGS", r"(\Seen)")
        return

    plain, htmlp = find_first_text_part(msg)
    orig_subj = decode_str(msg.get("Subject", "")).strip()

    # 1) Regras simples antes do LLM
    rb = rule_based_autoreply(orig_subj, plain, htmlp)
    if rb:
        acao = "responder"
        nivel = rb["nivel_confianca"]
        assunto_resp = rb["assunto"] or (f"Re: {orig_subj}" if not orig_subj.lower().startswith("re:") else orig_subj)
        corpo_md = rb["corpo_markdown"]
        logging.info("Regra simples aplicada: acao=responder conf=%.2f", nivel)
    else:
        # 2) LLM pass 1
        llm_input = assemble_llm_input(msg, plain, htmlp)
        decision = llm_decide({"input": llm_input}, nudge=False)
        acao = decision.get("acao","escalar")
        nivel = decision.get("nivel_confianca",0.0)
        assunto_resp = decision.get("assunto","").strip()
        corpo_md = decision.get("corpo_markdown","").strip()

        # 3) Se veio 'escalar' porém confiança não é tão baixa, tentar 2ª passada incentivando resposta
        if acao == "escalar" and REASK_ON_ESCALAR and nivel >= REASK_ESCALAR_MIN:
            logging.info("Reask ativado (acao=escalar conf=%.2f >= %.2f). Tentando 2ª passada.",
                         nivel, REASK_ESCALAR_MIN)
            decision2 = llm_decide({"input": llm_input}, nudge=True)
            if decision2.get("acao") == "responder":
                acao = "responder"
                nivel = decision2.get("nivel_confianca", nivel)
                assunto_resp = decision2.get("assunto","").strip() or assunto_resp
                corpo_md = decision2.get("corpo_markdown","").strip() or corpo_md

    # Segurança: se pediu responder, mas confiança muito baixa, volta a escalar
    if acao == "responder" and nivel < CONFIDENCE_THRESHOLD:
        logging.info("Confiança baixa (%.2f < %.2f). Alterando ação para 'escalar'.", nivel, CONFIDENCE_THRESHOLD)
        acao = "escalar"

    if acao == "responder":
        if not assunto_resp:
            assunto_resp = f"Re: {orig_subj}" if not orig_subj.lower().startswith("re:") else orig_subj
        to_addr = (email.utils.parseaddr(msg_get_addresses(msg, "Reply-To"))[1]
                   or email.utils.parseaddr(msg_get_addresses(msg, "From"))[1]
                   or "")
        if not to_addr:
            logging.warning("Sem destinatário para responder; escalando.")
            move_message(imap, msg_id, FOLDER_ESCALATE)
            return
        ok, _ = smtp_send_and_append(to_addr, assunto_resp, corpo_md, msg)
        if ok:
            target = FOLDER_PROCESSED if FOLDER_PROCESSED else "Respondidos"
            move_message(imap, msg_id, target)
            logging.info("E-mail movido para INBOX.%s", target)
        else:
            move_message(imap, msg_id, FOLDER_ESCALATE)
            logging.info("Falha no envio; e-mail movido para INBOX.%s", FOLDER_ESCALATE)
    else:
        move_message(imap, msg_id, FOLDER_ESCALATE)
        logging.info("E-mail movido para INBOX.%s", FOLDER_ESCALATE)

def process_unseen_once():
    try:
        imap = imap_connect()
        imap.select("INBOX")
        typ, data = imap.search(None, "UNSEEN")
        ids = data[0].split() if typ == "OK" else []
        logging.debug("UNSEEN: %s", ids)
        for msg_id in ids:
            typ, msgdata = imap.fetch(msg_id, "(RFC822)")
            if typ != "OK" or not msgdata or not msgdata[0]:
                continue
            raw = msgdata[0][1]
            msg = email.message_from_bytes(raw, policy=default_policy)
            decide_and_act(imap, msg_id, msg)
        imap.expunge()
        imap.logout()
    except imaplib.IMAP4.error as e:
        logging.error("Erro no ciclo IMAP: %s", e)
        logging.debug(traceback.format_exc())
    except Exception as e:
        logging.error("Erro geral no ciclo IMAP: %s", e)
        logging.debug(traceback.format_exc())

def background_loop():
    logging.info("Watcher IMAP — envio via SMTP HostGator")
    logging.info("App público em: %s", APP_PUBLIC_URL)
    while True:
        process_unseen_once()
        time.sleep(CHECK_INTERVAL_SECONDS)

# =========================
# HTTP App (Flask)
# =========================

app = Flask(__name__)

@app.route("/")
def index():
    return f"{APP_TITLE} — online"

@app.route("/healthz")
def healthz():
    return "ok"

@app.route("/diag/openrouter-chat")
def diag_openrouter():
    sample = (
        "Assunto: Teste\n"
        "De: Aluno <aluno@exemplo.com>\n\n"
        "Oi, como entro no grupo do Telegram e a assinatura do curso?"
    )
    dec1 = llm_decide({"input": sample}, nudge=False)
    dec2 = llm_decide({"input": sample}, nudge=True)
    return jsonify({"ok": True, "model": OPENROUTER_MODEL, "first": dec1, "nudge": dec2})

# =========================
# Main
# =========================

if __name__ == "__main__":
    t = threading.Thread(target=background_loop, daemon=True)
    t.start()
    app.run(host="0.0.0.0", port=PORT)
