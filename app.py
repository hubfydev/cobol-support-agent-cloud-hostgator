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

import os
import sys
import re
import json
import time
import logging
import imaplib
import smtplib
import threading
import traceback
from email.header import decode_header, make_header
from email.message import EmailMessage
from email import policy
from email.parser import BytesParser

from flask import Flask, jsonify, request

# ------------------------------------------------------------------------------
# Logging
# ------------------------------------------------------------------------------
logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)

# Deixa o imaplib verboso como no seu log
try:
    imaplib.Debug = int(os.getenv("IMAP_DEBUG", "4"))
except Exception:
    imaplib.Debug = 4

# ------------------------------------------------------------------------------
# Util: leitura de env com fallback, normalizando espaços
# ------------------------------------------------------------------------------
def _get_env(*names, default=None):
    for n in names:
        v = os.getenv(n)
        if v is not None:
            v = v.strip()
            if v != "":
                return v
    return default

# ------------------------------------------------------------------------------
# Config — IMAP/SMTP
# ------------------------------------------------------------------------------
IMAP_HOST = _get_env("IMAP_HOST", default=_get_env("MAIL_HOST", default="mail.aprendacobol.com.br"))
IMAP_PORT = int(_get_env("IMAP_PORT", default="993"))
IMAP_SSL  = _get_env("IMAP_SSL", default="true").lower() in ("1", "true", "yes", "on")

SMTP_HOST = _get_env("SMTP_HOST", default=IMAP_HOST)
SMTP_PORT = int(_get_env("SMTP_PORT", default="465"))
SMTP_SSL  = _get_env("SMTP_SSL", default="true").lower() in ("1", "true", "yes", "on")

# Credenciais com fallbacks (IMAP_* -> MAIL_* -> SMTP_*)
IMAP_USER = _get_env("IMAP_USER", "MAIL_USER", "SMTP_USER")
IMAP_PASS = _get_env("IMAP_PASS", "MAIL_PASS", "SMTP_PASS")

SMTP_USER = _get_env("SMTP_USER", default=IMAP_USER)
SMTP_PASS = _get_env("SMTP_PASS", default=IMAP_PASS)

if not IMAP_USER or not IMAP_PASS:
    logger.critical("Variáveis ausentes: IMAP_USER/IMAP_PASS (ou MAIL_USER/MAIL_PASS).")
    sys.exit(1)

logger.info("Credenciais: IMAP_USER=%s | SMTP_USER=%s", IMAP_USER, SMTP_USER)

# Pastas
FOLDER_INBOX = _get_env("IMAP_FOLDER_INBOX", default="INBOX")
FOLDER_ESCALAR = _get_env("IMAP_FOLDER_ESCALAR", default="INBOX.Escalar")
FOLDER_RESPONDIDOS = _get_env("IMAP_FOLDER_RESPONDIDOS", default="INBOX.Respondidos")

# Intervalo do loop de leitura
POLL_SECONDS = int(_get_env("POLL_SECONDS", default="60"))

# ------------------------------------------------------------------------------
# OpenRouter
# ------------------------------------------------------------------------------
OPENROUTER_API_KEY   = _get_env("OPENROUTER_API_KEY")
OPENROUTER_BASE_URL  = _get_env("OPENROUTER_BASE_URL", default="https://openrouter.ai/api/v1")
OPENROUTER_MODEL     = _get_env("OPENROUTER_MODEL", default="openrouter/auto")
OPENROUTER_MAXTOKENS = int(_get_env("OPENROUTER_MAXTOKENS", default="700"))
OPENROUTER_TEMP      = float(_get_env("OPENROUTER_TEMPERATURE", default="0.3"))

APP_PUBLIC_URL       = _get_env("APP_PUBLIC_URL", default=_get_env("RENDER_EXTERNAL_URL"))

# ------------------------------------------------------------------------------
# Prompt do sistema (exatamente como solicitado)
# ------------------------------------------------------------------------------
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

# ------------------------------------------------------------------------------
# Flask
# ------------------------------------------------------------------------------
app = Flask(__name__)

if APP_PUBLIC_URL:
    logger.info("App público em: %s", APP_PUBLIC_URL)

# ------------------------------------------------------------------------------
# Util: helpers de texto/HTML/assunto
# ------------------------------------------------------------------------------
_TAG_RE = re.compile(r"<[^>]+>")

def html_to_text(html: str) -> str:
    if not html:
        return ""
    # remove tags simples
    text = _TAG_RE.sub("", html)
    # converte entidades básicas
    text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    # normaliza espaços
    return re.sub(r"[ \t]+", " ", text).strip()

def decode_mime_header(v: str) -> str:
    if not v:
        return ""
    try:
        return str(make_header(decode_header(v)))
    except Exception:
        return v

def build_reply_subject(original_subj: str, suggested: str | None) -> str:
    base = suggested or original_subj or "(sem assunto)"
    base = base.strip()
    if not base.lower().startswith("re:"):
        return f"Re: {base}"
    return base

# ------------------------------------------------------------------------------
# Parser JSON tolerante (extrai o 1º objeto de nível 0 e valida)
# ------------------------------------------------------------------------------
def extract_first_json_object(text: str) -> dict:
    """
    Varre o texto, encontra o 1º bloco {...} de nível 0 (ignorando strings e escapes)
    e tenta json.loads. Lança ValueError se falhar.
    """
    if not text:
        raise ValueError("vazio")

    # Se já começa com JSON válido
    s = text.strip()
    if s.startswith("{"):
        try:
            return json.loads(s)
        except Exception:
            pass

    in_str = False
    esc = False
    depth = 0
    start = -1
    for i, ch in enumerate(text):
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
            elif ch == "}":
                if depth > 0:
                    depth -= 1
                    if depth == 0 and start >= 0:
                        candidate = text[start : i + 1]
                        try:
                            return json.loads(candidate)
                        except Exception:
                            # continua procurando próximo bloco
                            start = -1
                else:
                    # chaves desbalanceadas: ignora
                    pass

    raise ValueError("nenhum JSON válido encontrado")

_ALLOWED_ACOES = {"responder", "escalar"}

def normalize_decision(obj: dict) -> dict:
    """
    Garante o esquema:
      {"assunto": str, "corpo_markdown": str, "nivel_confianca": float[0..1], "acao": responder|escalar}
    Qualquer erro ⇒ acao='escalar', nivel=0.5
    """
    try:
        acao = str(obj.get("acao", "")).strip().lower()
        if acao not in _ALLOWED_ACOES:
            acao = "escalar"
        assunto = str(obj.get("assunto", "") or "").strip()
        corpo = str(obj.get("corpo_markdown", "") or "").strip()
        try:
            nivel = float(obj.get("nivel_confianca", 0.0))
        except Exception:
            nivel = 0.0
        if not (0.0 <= nivel <= 1.0):
            nivel = 0.0

        # Ajuste de regras pedidas:
        if acao == "responder" and nivel < 0.8:
            # força escalar se resposta veio com confiança baixa
            acao = "escalar"
        if acao == "escalar" and nivel > 0.6:
            # mantém escalar mas limita a 0.6 no máximo
            nivel = min(nivel, 0.6)

        return {
            "assunto": assunto,
            "corpo_markdown": corpo,
            "nivel_confianca": nivel,
            "acao": acao,
        }
    except Exception:
        return {
            "assunto": "",
            "corpo_markdown": "",
            "nivel_confianca": 0.5,
            "acao": "escalar",
        }

# ------------------------------------------------------------------------------
# OpenRouter: chamada e decisão
# ------------------------------------------------------------------------------
import urllib.request
import urllib.error

def openrouter_chat(email_payload: dict) -> dict:
    """
    Envia para o modelo e retorna dict normalizado do esquema.
    Em qualquer falha, retorna escalar (0.5).
    """
    if not OPENROUTER_API_KEY:
        logger.warning("OPENROUTER_API_KEY ausente — escalando por segurança.")
        return {"assunto": "", "corpo_markdown": "", "nivel_confianca": 0.5, "acao": "escalar"}

    url = f"{OPENROUTER_BASE_URL.rstrip('/')}/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        # headers recomendados pela OpenRouter:
        "HTTP-Referer": APP_PUBLIC_URL or "http://localhost",
        "X-Title": "COBOL Support Agent",
    }

    user_txt = (
        "Dados do e-mail:\n"
        f"De: {email_payload.get('from','')}\n"
        f"Assunto: {email_payload.get('subject','')}\n"
        f"Corpo:\n{email_payload.get('body','')}\n"
    )

    body = {
        "model": OPENROUTER_MODEL,
        "temperature": OPENROUTER_TEMP,
        "max_tokens": OPENROUTER_MAXTOKENS,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_txt},
        ],
    }

    data = json.dumps(body).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")

    preview = ""
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
            # não confiar cegamente—pega apenas 'choices[0].message.content' se houver
            j = json.loads(raw)
            content = (
                j.get("choices", [{}])[0]
                .get("message", {})
                .get("content", "")
            )
            preview = content[:1200]
            # extrai JSON do content
            parsed = extract_first_json_object(content)
            decision = normalize_decision(parsed)
            return decision
    except urllib.error.HTTPError as e:
        err_body = e.read().decode("utf-8", errors="replace")
        logger.warning("Falha HTTP OpenRouter %s: %s", e.code, err_body)
    except Exception as e:
        logger.warning("Falha OpenRouter: %s", str(e))

    # fallback ⇒ escalar
    if preview:
        logger.warning("Falha ao parsear JSON do modelo: preview=%s", preview[:200].replace("\n", " ")[:200])
    return {"assunto": "", "corpo_markdown": "", "nivel_confianca": 0.5, "acao": "escalar"}

# ------------------------------------------------------------------------------
# IMAP helpers
# ------------------------------------------------------------------------------
def imap_connect():
    if IMAP_SSL:
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    else:
        imap = imaplib.IMAP4(IMAP_HOST, IMAP_PORT)
    imap.login(IMAP_USER, IMAP_PASS)
    return imap

def ensure_mailbox(imap, mailbox: str):
    try:
        typ, _ = imap.create(mailbox)
        # Se já existir, muitos servidores retornam NO [ALREADYEXISTS] — ok
    except Exception:
        pass

def move_message(imap, msg_seq: bytes, dest_mailbox: str):
    ensure_mailbox(imap, dest_mailbox)
    imap.copy(msg_seq, dest_mailbox)
    imap.store(msg_seq, "+FLAGS", r"(\Deleted)")
    imap.expunge()

# ------------------------------------------------------------------------------
# SMTP: envio de resposta
# ------------------------------------------------------------------------------
def send_reply(to_addr: str, subject: str, markdown_body: str, original_from: str):
    if not to_addr:
        to_addr = original_from

    # Conversão super simples de markdown para HTML (linhas/quebras)
    html = "<html><body>" + "<br>".join(markdown_body.splitlines()) + "</body></html>"

    msg = EmailMessage()
    msg["From"] = SMTP_USER
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content(markdown_body)
    msg.add_alternative(html, subtype="html")

    if SMTP_SSL:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as s:
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.starttls()
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)

# ------------------------------------------------------------------------------
# Processa novos emails
# ------------------------------------------------------------------------------
def extract_email_payload(msg):
    # remetente
    from_hdr = decode_mime_header(msg.get("From", ""))
    subj_hdr = decode_mime_header(msg.get("Subject", ""))

    body_txt = ""
    body_html = ""

    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get("Content-Disposition") or "").lower()
            if "attachment" in disp:
                continue
            try:
                payload = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                text = payload.decode(charset, errors="replace")
            except Exception:
                text = ""
            if ctype == "text/plain" and not body_txt:
                body_txt = text
            elif ctype == "text/html" and not body_html:
                body_html = text
    else:
        payload = msg.get_payload(decode=True) or b""
        charset = msg.get_content_charset() or "utf-8"
        try:
            text = payload.decode(charset, errors="replace")
        except Exception:
            text = ""
        if msg.get_content_type() == "text/html":
            body_html = text
        else:
            body_txt = text

    if not body_txt and body_html:
        body_txt = html_to_text(body_html)

    return {
        "from": from_hdr,
        "subject": subj_hdr,
        "body": body_txt.strip(),
    }

def process_unseen_once():
    try:
        imap = imap_connect()
        imap.select(FOLDER_INBOX)
        typ, data = imap.search(None, "UNSEEN")
        unseen = data[0].split()
        logger.debug("UNSEEN: %s", unseen)

        for seq in unseen:
            # Carrega mensagem
            typ, msgdata = imap.fetch(seq, "(BODY.PEEK[])")
            if typ != "OK" or not msgdata:
                continue
            raw = None
            for part in msgdata:
                if isinstance(part, tuple):
                    raw = part[1]
                    break
            if not raw:
                continue

            msg = BytesParser(policy=policy.default).parsebytes(raw)
            payload = extract_email_payload(msg)

            # Chama modelo
            decision = openrouter_chat(payload)
            acao = decision.get("acao")
            nivel = decision.get("nivel_confianca", 0.0)

            logger.info("Decisão do modelo: acao=%s conf=%.2f", acao, nivel)

            if acao == "responder":
                # monta assunto
                reply_subj = build_reply_subject(payload.get("subject",""), decision.get("assunto"))
                try:
                    send_reply(
                        to_addr=None,  # usa o próprio From
                        subject=reply_subj,
                        markdown_body=decision.get("corpo_markdown", ""),
                        original_from=payload.get("from",""),
                    )
                    # move para Respondidos
                    move_message(imap, seq, FOLDER_RESPONDIDOS)
                except Exception:
                    logger.exception("Falha ao enviar resposta; escalando mesmo assim.")
                    move_message(imap, seq, FOLDER_ESCALAR)
            else:
                # escalar
                move_message(imap, seq, FOLDER_ESCALAR)

        # limpeza de flags
        imap.expunge()
        imap.logout()
    except Exception as e:
        logger.error("Erro no ciclo IMAP: %s", e)
        logger.debug(traceback.format_exc())

def background_loop():
    logger.info("Watcher IMAP — envio via SMTP HostGator")
    while True:
        process_unseen_once()
        time.sleep(POLL_SECONDS)

# ------------------------------------------------------------------------------
# Rotas HTTP
# ------------------------------------------------------------------------------
@app.get("/")
def index():
    return jsonify(
        ok=True,
        service="COBOL Support Agent",
        inbox=FOLDER_INBOX,
        escalar=FOLDER_ESCALAR,
        respondidos=FOLDER_RESPONDIDOS,
        model=OPENROUTER_MODEL,
        public_url=APP_PUBLIC_URL,
    )

@app.get("/diag/env")
def diag_env():
    return jsonify(
        IMAP_HOST=IMAP_HOST,
        IMAP_PORT=IMAP_PORT,
        SMTP_HOST=SMTP_HOST,
        SMTP_PORT=SMTP_PORT,
        IMAP_SSL=IMAP_SSL,
        SMTP_SSL=SMTP_SSL,
        IMAP_USER_present=bool(IMAP_USER),
        IMAP_PASS_present=bool(IMAP_PASS),
        SMTP_USER_present=bool(SMTP_USER),
        SMTP_PASS_present=bool(SMTP_PASS),
    )

@app.get("/diag/openrouter")
def diag_openrouter():
    return jsonify(
        base_url=OPENROUTER_BASE_URL,
        model=OPENROUTER_MODEL,
        have_key=bool(OPENROUTER_API_KEY),
    )

@app.get("/diag/openrouter-chat")
def diag_openrouter_chat():
    """
    Por padrão, não chama a API (evita custo). Se passar ?live=1, faz uma chamada mínima.
    """
    live = request.args.get("live") == "1"
    if not live:
        return jsonify(ok=True, eco="teste", model=OPENROUTER_MODEL)

    if not OPENROUTER_API_KEY:
        return jsonify(ok=False, error="OPENROUTER_API_KEY ausente"), 400

    try:
        decision = openrouter_chat({
            "from": "teste@exemplo.com",
            "subject": "ping",
            "body": "Apenas um teste rápido.",
        })
        return jsonify(ok=True, decision=decision)
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 500

# ------------------------------------------------------------------------------
# Main
# ------------------------------------------------------------------------------
if __name__ == "__main__":
    # Thread do watcher de e-mails
    t = threading.Thread(target=background_loop, daemon=True)
    t.start()

    host = "0.0.0.0"
    port = int(_get_env("PORT", default="10000"))
    app.run(host=host, port=port, debug=False)
