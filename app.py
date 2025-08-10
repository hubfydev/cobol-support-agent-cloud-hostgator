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
import smtplib
import imaplib
import logging
import threading
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.message import EmailMessage
from flask import Flask, jsonify

# ==============================================================================
# Logging (robusto: aceita "debug", "INFO", "20", etc.)
# ==============================================================================
_raw_level = os.getenv("LOG_LEVEL", "INFO")
try:
    LOG_LEVEL = int(_raw_level)
except ValueError:
    LOG_LEVEL = getattr(logging, str(_raw_level).upper(), logging.INFO)

logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("cobol-support-agent")

# Verbose do imaplib (0..5). Deixa em 4 no Render para ver comandos.
try:
    imaplib.Debug = int(os.getenv("IMAP_DEBUG", "4"))
except Exception:
    imaplib.Debug = 0

# ==============================================================================
# Config / ENV helpers
# ==============================================================================
def env(name: str, default: str | None = None) -> str | None:
    v = os.getenv(name)
    if v is None or (isinstance(v, str) and v.strip() == ""):
        return default
    return v

APP_PUBLIC_URL = env("APP_PUBLIC_URL")
PORT = int(env("PORT", "10000"))

# IMAP / SMTP — prioriza MAIL_*, depois IMAP_*/SMTP_*
IMAP_HOST = env("IMAP_HOST", env("MAIL_IMAP_HOST", "mail.aprendacobol.com.br"))
IMAP_PORT = int(env("IMAP_PORT", env("MAIL_IMAP_PORT", "993")))
IMAP_SSL = env("IMAP_SSL", env("MAIL_IMAP_SSL", "true")).lower() in ("1", "true", "yes")

SMTP_HOST = env("SMTP_HOST", env("MAIL_SMTP_HOST", IMAP_HOST))
SMTP_PORT = int(env("SMTP_PORT", env("MAIL_SMTP_PORT", "465")))
SMTP_SSL = env("SMTP_SSL", env("MAIL_SMTP_SSL", "true")).lower() in ("1", "true", "yes")

# Credenciais (usa MAIL_* primeiro)
IMAP_USER = env("IMAP_USER", env("MAIL_USER", env("SMTP_USER")))
IMAP_PASS = env("IMAP_PASS", env("MAIL_PASS", env("SMTP_PASS")))

FROM_EMAIL = env("FROM_EMAIL", IMAP_USER)  # remetente das respostas

# OpenRouter
OPENROUTER_API_KEY = env("OPENROUTER_API_KEY")
OPENROUTER_MODEL = env("OPENROUTER_MODEL", "openrouter/auto")
OPENROUTER_ENDPOINT = "https://openrouter.ai/api/v1/chat/completions"
OPENROUTER_REFERRER = APP_PUBLIC_URL or "https://render.com"
OPENROUTER_TITLE = "COBOL Support Agent"

# Watcher
POLL_INTERVAL_SEC = int(env("POLL_INTERVAL_SEC", "60"))
MOVE_ESCALAR_FOLDER = env("FOLDER_ESCALAR", "INBOX.Escalar")
MOVE_RESPONDIDOS_FOLDER = env("FOLDER_RESPONDIDOS", "INBOX.Respondidos")

# ==============================================================================
# System prompt (versão solicitada)
# ==============================================================================
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
    "Sempre ofereça o curso de Formação Completa de Programador Aprenda COBOL disponível por assinatura em: https://assinatura.aprendacobol.com.br."
)

# ==============================================================================
# Flask
# ==============================================================================
app = Flask(__name__)

@app.route("/")
def home():
    return "OK - COBOL Support Agent"

@app.route("/diag/env")
def diag_env():
    return jsonify({
        "APP_PUBLIC_URL": APP_PUBLIC_URL,
        "imap": {
            "host": IMAP_HOST,
            "port": IMAP_PORT,
            "ssl": IMAP_SSL,
            "IMAP_USER_present": bool(IMAP_USER),
            "IMAP_PASS_present": bool(IMAP_PASS),
        },
        "smtp": {
            "host": SMTP_HOST,
            "port": SMTP_PORT,
            "ssl": SMTP_SSL,
            "FROM_EMAIL": FROM_EMAIL,
        },
        "openrouter": {
            "model": OPENROUTER_MODEL,
            "api_key_present": bool(OPENROUTER_API_KEY),
        },
        "poll_interval_sec": POLL_INTERVAL_SEC,
        "folders": {
            "escalar": MOVE_ESCALAR_FOLDER,
            "respondidos": MOVE_RESPONDIDOS_FOLDER,
        },
        "log_level": LOG_LEVEL,
    })

@app.route("/diag/openrouter-chat")
def diag_openrouter_chat():
    # Rota “eco” simples para checagem sem custo
    return jsonify({"ok": True, "eco": "teste", "model": OPENROUTER_MODEL})

# ==============================================================================
# Helpers — IMAP / SMTP
# ==============================================================================
def imap_connect() -> imaplib.IMAP4:
    if not IMAP_USER or not IMAP_PASS:
        raise RuntimeError("Variáveis ausentes: IMAP_USER/IMAP_PASS (ou MAIL_USER/MAIL_PASS).")

    if IMAP_SSL:
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    else:
        imap = imaplib.IMAP4(IMAP_HOST, IMAP_PORT)
        imap.starttls()

    # imaplib.Debug cuidará de logar o LOGIN; aqui só fazemos o login
    imap.login(IMAP_USER, IMAP_PASS)
    return imap

def smtp_connect() -> smtplib.SMTP:
    if SMTP_SSL or SMTP_PORT == 465:
        context = ssl.create_default_context()
        server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context)
    else:
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.starttls()
    server.login(IMAP_USER, IMAP_PASS)  # usa as mesmas credenciais
    return server

def ensure_mailbox(imap: imaplib.IMAP4, mbox: str):
    typ, data = imap.list()
    if typ != "OK":
        logger.warning("LIST falhou ao verificar caixas.")
        return
    names = []
    if data:
        for raw in data:
            # raw é bytes, ex: b'(\\HasNoChildren) "." INBOX.Escalar'
            s = raw.decode("utf-8", errors="ignore")
            parts = s.split(" ")
            if parts:
                name = s.split(' "." ')[-1].strip().strip('"')
                names.append(name)
    if mbox not in names:
        typ, _ = imap.create(mbox)
        if typ != "OK":
            logger.warning("CREATE %s falhou ou já existe.", mbox)

def move_message(imap: imaplib.IMAP4, msg_seq: bytes, dest_folder: str):
    ensure_mailbox(imap, dest_folder)
    typ, _ = imap.copy(msg_seq, dest_folder)
    if typ != "OK":
        logger.error("COPY falhou para %s -> %s", msg_seq, dest_folder)
        return
    imap.store(msg_seq, "+FLAGS", r"(\Deleted)")
    imap.expunge()

# ==============================================================================
# Extração de texto do e-mail
# ==============================================================================
def extract_text_from_email(msg) -> str:
    if msg.is_multipart():
        parts = []
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = part.get_content_disposition()
            if ctype == "text/plain" and disp != "attachment":
                try:
                    parts.append(part.get_content())
                except Exception:
                    payload = part.get_payload(decode=True)
                    if payload:
                        parts.append(payload.decode("utf-8", errors="ignore"))
        if parts:
            return "\n\n".join(parts)
    # fallback
    try:
        return msg.get_content()
    except Exception:
        payload = msg.get_payload(decode=True)
        if payload:
            return payload.decode("utf-8", errors="ignore")
    return ""

# ==============================================================================
# OpenRouter — chamada e parsing
# ==============================================================================
def call_openrouter(messages, max_tokens=700, temperature=0.2) -> str:
    """
    Retorna o texto bruto do modelo (pode vir com lixo). Não levanta exceção
    por status HTTP != 200, devolve string vazia nesses casos.
    """
    if not OPENROUTER_API_KEY:
        logger.warning("OPENROUTER_API_KEY ausente; retornando string vazia.")
        return ""

    import urllib.request
    import urllib.error

    payload = {
        "model": OPENROUTER_MODEL,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
        # Ajuda alguns modelos a priorizarem JSON:
        "response_format": {"type": "json_object"},
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        OPENROUTER_ENDPOINT,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {OPENROUTER_API_KEY}",
            "HTTP-Referer": OPENROUTER_REFERRER,
            "X-Title": OPENROUTER_TITLE,
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read()
            try:
                obj = json.loads(body.decode("utf-8", errors="ignore"))
            except Exception:
                logger.warning("OpenRouter retornou corpo não-JSON.")
                return body.decode("utf-8", errors="ignore")
            # padrão OpenRouter
            if isinstance(obj, dict) and "choices" in obj and obj["choices"]:
                return obj["choices"][0]["message"].get("content", "")
            # fallback: retorna o bruto
            return body.decode("utf-8", errors="ignore")
    except urllib.error.HTTPError as e:
        logger.warning("OpenRouter HTTPError %s: %s", e.code, e.read().decode("utf-8", errors="ignore"))
        return ""
    except Exception as e:
        logger.warning("OpenRouter erro: %s", e)
        return ""

def _strip_code_fences(text: str) -> str:
    # remove ```json ... ``` e ``` ... ```
    return re.sub(r"```[\s\S]*?```", "", text or "", flags=re.DOTALL)

def _find_first_balanced_json(text: str) -> str | None:
    """
    Varre o texto e retorna o primeiro trecho com chaves balanceadas
    respeitando aspas e escapes. None se não achar.
    """
    s = text
    in_str = False
    esc = False
    depth = 0
    start = -1
    for i, ch in enumerate(s):
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
                    if depth == 0 and start != -1:
                        return s[start:i+1]
    return None

def parse_model_json(raw_text: str) -> dict | None:
    """
    Tenta extrair um JSON válido com o esquema exigido.
    Retorna dict ou None.
    """
    if not raw_text:
        return None

    # 1) remove cercas de código
    cleaned = _strip_code_fences(raw_text).strip()

    # 2) tentativa direta
    if cleaned.startswith("{") and cleaned.endswith("}"):
        try:
            return json.loads(cleaned)
        except Exception:
            pass

    # 3) tentar achar primeiro bloco balanceado
    candidate = _find_first_balanced_json(cleaned)
    if candidate:
        try:
            return json.loads(candidate)
        except Exception:
            # tenta de novo removendo possíveis BOMs/char estranhos
            try:
                return json.loads(candidate.encode("utf-8", "ignore").decode("utf-8"))
            except Exception:
                pass

    # 4) heurística: pegar a MAIOR substring entre primeira "{" e última "}"
    try:
        first = cleaned.index("{")
        last = cleaned.rindex("}")
        big = cleaned[first:last+1]
        return json.loads(big)
    except Exception:
        return None

def decide_action_from_model(email_subject: str, email_from: str, email_body: str) -> dict:
    """
    Monta prompt, chama o modelo, aplica parser e valida o esquema.
    Em caso de falha, retorna decisão de 'escalar'.
    """
    # Conteúdo para o modelo (limita tamanho para evitar custos/estouro)
    body_trim = email_body.strip()
    if len(body_trim) > 7000:
        body_trim = body_trim[:7000] + "\n\n[TRUNCADO]"

    user_content = (
        "### E-mail recebido\n"
        f"Assunto: {email_subject}\n"
        f"Remetente: {email_from}\n\n"
        f"Corpo:\n{body_trim}\n"
    )

    messages = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]

    raw = call_openrouter(messages)
    preview = (raw or "")[:1200]
    if not raw:
        logger.warning("Falha no OpenRouter: resposta vazia.")
        return {"acao": "escalar", "nivel_confianca": 0.5, "assunto": "", "corpo_markdown": ""}

    data = parse_model_json(raw)
    if not data:
        logger.warning("Falha ao parsear JSON do modelo: preview=%s", preview)
        return {"acao": "escalar", "nivel_confianca": 0.5, "assunto": "", "corpo_markdown": ""}

    # Normalização / validação
    acao = str(data.get("acao", "")).lower().strip()
    nivel = data.get("nivel_confianca")
    assunto = str(data.get("assunto", "")).strip()
    corpo = str(data.get("corpo_markdown", "")).strip()

    if acao not in ("responder", "escalar"):
        logger.warning("Campo 'acao' inválido: %s", acao)
        return {"acao": "escalar", "nivel_confianca": 0.5, "assunto": "", "corpo_markdown": ""}

    try:
        nivel = float(nivel)
    except Exception:
        nivel = 0.0

    if not (0.0 <= nivel <= 1.0):
        nivel = 0.0

    if acao == "responder":
        # requisitos mínimos para responder
        if nivel < 0.8 or not corpo:
            logger.info("Modelo sugeriu responder, mas nivel/corpo insuficiente. Escalando.")
            return {"acao": "escalar", "nivel_confianca": 0.5, "assunto": "", "corpo_markdown": ""}

        if not assunto:
            assunto = f"Re: {email_subject or ''}".strip()

    return {"acao": acao, "nivel_confianca": nivel, "assunto": assunto, "corpo_markdown": corpo}

# ==============================================================================
# Resposta por e-mail
# ==============================================================================
def send_reply(to_addr: str, original_subject: str, reply_subject: str, body_md: str):
    msg = EmailMessage()
    msg["From"] = FROM_EMAIL
    msg["To"] = to_addr
    msg["Subject"] = reply_subject or f"Re: {original_subject}"

    # Envia como texto simples; se quiser HTML, converter markdown -> HTML aqui
    msg.set_content(body_md or "", subtype="plain", charset="utf-8")

    with smtp_connect() as smtp:
        smtp.send_message(msg)

# ==============================================================================
# Loop de processamento IMAP
# ==============================================================================
def process_unseen_once():
    imap = None
    try:
        imap = imap_connect()
        typ, _ = imap.select("INBOX")
        if typ != "OK":
            logger.error("SELECT INBOX falhou.")
            return

        typ, data = imap.search(None, "UNSEEN")
        if typ != "OK":
            logger.error("SEARCH UNSEEN falhou.")
            return

        unseen = data[0].split() if data and data[0] else []
        logger.debug("UNSEEN: %s", unseen)

        if not unseen:
            # housekeeping: EXPUNGE limpa deletados
            imap.expunge()
            return

        for seq in unseen:
            # fetch corpo inteiro
            typ, msg_data = imap.fetch(seq, "(BODY.PEEK[])")
            if typ != "OK" or not msg_data:
                logger.warning("FETCH falhou para %s", seq)
                continue

            raw_email = b""
            for part in msg_data:
                if isinstance(part, tuple) and len(part) == 2:
                    raw_email += part[1]

            msg = BytesParser(policy=policy.default).parsebytes(raw_email)
            from_addr = (msg.get("Reply-To") or msg.get("From") or "").strip()
            subject = (msg.get("Subject") or "").strip()
            body = extract_text_from_email(msg)

            # Decisão via modelo (com fail-safe)
            decision = decide_action_from_model(subject, from_addr, body)
            acao = decision["acao"]
            nivel = decision["nivel_confianca"]
            assunto_resp = decision.get("assunto", "")
            corpo_md = decision.get("corpo_markdown", "")

            logger.info("Ação=%s conf=%.2f", acao, nivel)

            if acao == "responder" and nivel >= 0.8 and corpo_md:
                try:
                    send_reply(from_addr, subject, assunto_resp, corpo_md)
                    move_message(imap, seq, MOVE_RESPONDIDOS_FOLDER)
                except Exception as e:
                    logger.exception("Falha ao responder; escalando. Erro: %s", e)
                    move_message(imap, seq, MOVE_ESCALAR_FOLDER)
            else:
                move_message(imap, seq, MOVE_ESCALAR_FOLDER)

    except Exception as e:
        logger.exception("Erro no ciclo IMAP: %s", e)
    finally:
        try:
            if imap is not None:
                imap.logout()
        except Exception:
            pass

def watcher_loop():
    logger.info("Watcher IMAP — envio via SMTP HostGator")
    if APP_PUBLIC_URL:
        logger.info("App público em: %s", APP_PUBLIC_URL)
    while True:
        process_unseen_once()
        time.sleep(POLL_INTERVAL_SEC)

# ==============================================================================
# Main
# ==============================================================================
if __name__ == "__main__":
    # Sobe watcher em thread separada
    t = threading.Thread(target=watcher_loop, daemon=True)
    t.start()

    # Flask web
    app.run(host="0.0.0.0", port=PORT, debug=False)
