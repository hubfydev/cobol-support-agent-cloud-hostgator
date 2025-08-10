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
import re
import json
import time
import imaplib
import smtplib
import logging
import threading
import traceback
from typing import Optional, Tuple, List, Dict
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from email.header import decode_header, Header
from email.utils import parsedate_to_datetime, make_msgid, formatdate

import requests
from flask import Flask, jsonify

# ------------------------------------------------------------------------------
# Config
# ------------------------------------------------------------------------------

APP_NAME = os.getenv("APP_NAME", "COBOL Support Agent")
APP_PUBLIC_URL = os.getenv("APP_PUBLIC_URL", "")
PORT = int(os.getenv("PORT", "10000"))
IMAP_HOST = os.getenv("IMAP_HOST", "")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
IMAP_USER = os.getenv("IMAP_USER", "")
IMAP_PASS = os.getenv("IMAP_PASS", "")

SMTP_HOST = os.getenv("SMTP_HOST", "")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", IMAP_USER)
SMTP_PASS = os.getenv("SMTP_PASS", IMAP_PASS)
SMTP_TLS = os.getenv("SMTP_TLS", "true").lower() in ("1", "true", "yes")

FOLDER_ESCALAR = os.getenv("FOLDER_ESCALAR", "INBOX.Escalar")
FOLDER_RESPONDIDOS = os.getenv("FOLDER_RESPONDIDOS", "INBOX.Respondidos")

POLL_SECONDS = int(os.getenv("POLL_SECONDS", "60"))
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openrouter/auto")
OPENROUTER_MODEL_WRITER = os.getenv("OPENROUTER_MODEL_WRITER", OPENROUTER_MODEL)
OPENROUTER_MAXTOK = int(os.getenv("OPENROUTER_MAXTOK", "700"))
OPENROUTER_WRITER_MAXTOK = int(os.getenv("OPENROUTER_WRITER_MAXTOK", "600"))
OPENROUTER_TEMPERATURE = float(os.getenv("OPENROUTER_TEMPERATURE", "0.2"))
CONF_MIN = float(os.getenv("CONF_MIN", "0.80"))

# ------------------------------------------------------------------------------
# Logging
# ------------------------------------------------------------------------------

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("cobol-support-agent")

# imprimir cabeçalho amigável
logger.info("Watcher IMAP — envio via SMTP HostGator")


# ------------------------------------------------------------------------------
# Utilitários de e-mail
# ------------------------------------------------------------------------------

def _decode_mime_words(s: Optional[str]) -> str:
    if not s:
        return ""
    try:
        decoded = ""
        for frag, enc in decode_header(s):
            if isinstance(frag, bytes):
                decoded += frag.decode(enc or "utf-8", errors="replace")
            else:
                decoded += frag
        return decoded
    except Exception:
        return s


def _msg_get_text_parts(msg) -> Tuple[str, str]:
    """
    Retorna (text/plain, text/html) como strings (podem estar vazias).
    """
    text = ""
    html = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = part.get_content_disposition()
            if disp == "attachment":
                continue
            try:
                payload = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                content = payload.decode(charset, errors="replace")
            except Exception:
                content = ""
            if ctype == "text/plain" and not text:
                text = content
            elif ctype == "text/html" and not html:
                html = content
    else:
        ctype = msg.get_content_type()
        payload = msg.get_payload(decode=True) or b""
        charset = msg.get_content_charset() or "utf-8"
        content = payload.decode(charset, errors="replace")
        if ctype == "text/html":
            html = content
        else:
            text = content

    return text, html


def _markdown_to_html(md: str) -> str:
    # Evita dependência externa: conversão mínima
    safe = (md or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    safe = safe.replace("\n", "<br>")
    return f"<div style='font-family: -apple-system,Segoe UI,Roboto,Arial,sans-serif; line-height:1.5; font-size:14px'>{safe}</div>"


# ------------------------------------------------------------------------------
# IMAP helpers
# ------------------------------------------------------------------------------

def imap_connect() -> imaplib.IMAP4_SSL:
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    tag = imap._new_tag()
    logger.debug(" %s > %r", time.strftime("%H:%M.%S"), f'{tag} LOGIN {IMAP_USER} "********"'.encode())
    imap.login(IMAP_USER, IMAP_PASS)
    return imap


def ensure_mailbox(imap: imaplib.IMAP4_SSL, mailbox: str):
    typ, data = imap.list("", "*")
    if typ == "OK":
        existing = [ln.decode(errors="ignore").split(' "." ')[-1].strip() for ln in (data or []) if ln]
        if f'"{mailbox}"' in existing or mailbox in existing:
            return
    tag = imap._new_tag()
    logger.debug(" %s > %r", time.strftime("%H:%M.%S"), f"{tag} CREATE {mailbox}".encode())
    imap.create(mailbox)


def move_message(imap: imaplib.IMAP4_SSL, uid: bytes, dest_mailbox: str):
    """
    MOVE por COPY+DELETE+EXPUNGE (compatível com servidores sem MOVE).
    """
    ensure_mailbox(imap, dest_mailbox)
    tag = imap._new_tag()
    logger.debug(" %s > %r", time.strftime("%H:%M.%S"), f"{tag} COPY {uid.decode()} {dest_mailbox}".encode())
    imap.uid("COPY", uid, dest_mailbox)
    tag = imap._new_tag()
    logger.debug(" %s > %r", time.strftime("%H:%M.%S"), f"{tag} STORE {uid.decode()} +FLAGS (\\Deleted)".encode())
    imap.uid("STORE", uid, "+FLAGS", r"(\Deleted)")
    tag = imap._new_tag()
    logger.debug(" %s > %r", time.strftime("%H:%M.%S"), f"{tag} EXPUNGE".encode())
    imap.expunge()


# ------------------------------------------------------------------------------
# SMTP
# ------------------------------------------------------------------------------

def send_email_reply(
    to_addr: str,
    from_addr: str,
    subject: str,
    body_md: str,
    in_reply_to: Optional[str] = None,
    references: Optional[str] = None,
):
    msg = EmailMessage()
    msg["From"] = from_addr
    msg["To"] = to_addr
    msg["Subject"] = Header(subject, "utf-8")
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=(from_addr.split("@")[-1] if "@" in from_addr else None))
    if in_reply_to:
        msg["In-Reply-To"] = in_reply_to
    if references:
        msg["References"] = references

    body_text = (body_md or "").replace("\r\n", "\n")
    body_html = _markdown_to_html(body_text)

    msg.set_content(body_text)
    msg.add_alternative(body_html, subtype="html")

    if not SMTP_HOST or not SMTP_USER or not SMTP_PASS:
        raise RuntimeError("SMTP não configurado.")

    if SMTP_TLS:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as s:
            s.ehlo()
            s.starttls()
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
    else:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as s:
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)


# ------------------------------------------------------------------------------
# OpenRouter
# ------------------------------------------------------------------------------

_OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"


def openrouter_chat(
    messages: List[Dict[str, str]],
    model: Optional[str] = None,
    max_tokens: Optional[int] = None,
    temperature: Optional[float] = None,
    top_p: Optional[float] = None,
) -> Optional[str]:
    """
    Faz uma chamada de chat ao OpenRouter e retorna o content do 1º choice.
    Em erros, retorna None.
    """
    if not OPENROUTER_API_KEY:
        logger.warning("OpenRouter sem API key. Pulei chamada.")
        return None

    payload = {
        "model": model or OPENROUTER_MODEL,
        "messages": messages,
    }
    if max_tokens is not None:
        payload["max_tokens"] = max_tokens
    if temperature is not None:
        payload["temperature"] = temperature
    if top_p is not None:
        payload["top_p"] = top_p

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "HTTP-Referer": APP_PUBLIC_URL or f"https://{APP_NAME.replace(' ', '').lower()}.local",
        "X-Title": APP_NAME,
        "Content-Type": "application/json",
    }

    try:
        resp = requests.post(_OPENROUTER_URL, headers=headers, data=json.dumps(payload), timeout=60)
        if not resp.ok:
            logger.warning("Falha OpenRouter: %s %s", resp.status_code, resp.text[:400])
            return None
        j = resp.json()
        content = j.get("choices", [{}])[0].get("message", {}).get("content")
        if not content:
            logger.warning("OpenRouter retornou sem content.")
            return None
        return content
    except Exception as e:
        logger.warning("Erro OpenRouter: %s", e)
        return None


# ------------------------------------------------------------------------------
# Parser robusto (tolerante a truncos / lixo)
# ------------------------------------------------------------------------------

_JSON_FENCE_RE = re.compile(r"^```(?:json)?\s*|\s*```$", re.MULTILINE)

def parse_ai_json_robusta(s: str) -> Optional[dict]:
    """
    Extrai o PRIMEIRO objeto JSON de uma saída possivelmente contaminada:
    - ignora lixo antes/depois
    - fecha o objeto no primeiro '}' de nível 0
    - se truncado, remove a ÚLTIMA propriedade aberta e fecha com '}'
    - fallback por regex para campos mínimos (acao, nivel_confianca, assunto)
    Retorna None apenas quando nada dá para aproveitar.
    """
    if not s:
        return None

    # remove cercas ```json ... ```
    s = _JSON_FENCE_RE.sub("", s).strip()

    # pega a partir do primeiro '{'
    i = s.find("{")
    if i == -1:
        return None
    tail = s[i:]

    # varredura para achar o primeiro JSON balanceado
    depth = 0
    in_str = False
    esc = False
    end_idx = -1
    last_comma_depth1 = -1  # para cortar a última prop se truncado

    for idx, ch in enumerate(tail):
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
                    end_idx = idx
                    break
            elif ch == "," and depth == 1:
                last_comma_depth1 = idx

    candidate = tail[: end_idx + 1] if end_idx != -1 else tail

    # 1) tenta carregar direto
    try:
        return json.loads(candidate)
    except Exception:
        pass

    # 2) se NÃO achou '}' (truncado), corta no último ',' de nível 1 e fecha
    if end_idx == -1 and last_comma_depth1 != -1:
        fixed = tail[:last_comma_depth1] + "}"
        try:
            return json.loads(fixed)
        except Exception:
            pass

    # 3) salvamento mínimo por regex
    salv: Dict[str, object] = {}

    m = re.search(r'"acao"\s*:\s*"([^"]+)"', s, flags=re.IGNORECASE)
    if m:
        salv["acao"] = m.group(1).strip().lower()

    m = re.search(r'"nivel_confianca"\s*:\s*([0-9]*\.?[0-9]+)', s, flags=re.IGNORECASE)
    if m:
        try:
            salv["nivel_confianca"] = float(m.group(1))
        except Exception:
            pass

    m = re.search(r'"assunto"\s*:\s*"([^"]+)"', s, flags=re.IGNORECASE)
    if m:
        salv["assunto"] = m.group(1)

    if salv:
        salv["_partial"] = True
        return salv

    return None


# ------------------------------------------------------------------------------
# Lógica de decisão + integração com OpenRouter
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

def decidir_acao_assistente(
    remetente: str,
    assunto: str,
    corpo_texto: str,
) -> Optional[dict]:
    user_prompt = (
        f"Remetente: {remetente}\n"
        f"Assunto: {assunto}\n\n"
        f"Corpo:\n{corpo_texto}\n"
    )

    content = openrouter_chat(
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        model=OPENROUTER_MODEL,
        max_tokens=OPENROUTER_MAXTOK,
        temperature=OPENROUTER_TEMPERATURE,
        top_p=0.9,
    )
    if content is None:
        # Falha de chamada — escalar
        logger.warning("[WARN] Falha no OpenRouter: retorno None")
        return None

    # Pré-visualização para debug
    preview = content if len(content) < 1200 else (content[:1200] + " …")
    logger.warning("[WARN] JSON inválido do modelo (preview 1.2KB): \n%s", preview)  # mantemos mesma mensagem do seu log

    data = parse_ai_json_robusta(content)
    return data


# ------------------------------------------------------------------------------
# Loop IMAP
# ------------------------------------------------------------------------------

def process_unseen_once():
    imap = None
    try:
        imap = imap_connect()
        imap.select("INBOX")
        typ, data = imap.search(None, "UNSEEN")
        if typ != "OK":
            return

        uids = (data[0] or b"").split()
        logger.debug("[DEBUG] UNSEEN: %s", uids)

        if not uids:
            # housekeeping leve
            logger.debug("[DEBUG] Executando EXPUNGE…")
            imap.expunge()
            return

        for uid in uids:
            # FETCH
            typ, msg_data = imap.fetch(uid, "(BODY.PEEK[])")
            if typ != "OK" or not msg_data:
                continue

            raw = b""
            for part in msg_data:
                if isinstance(part, tuple) and part[1]:
                    raw += part[1]

            msg = BytesParser(policy=policy.default).parsebytes(raw)
            sub = _decode_mime_words(msg.get("Subject"))
            frm = _decode_mime_words(msg.get("From"))
            in_reply_to = msg.get("Message-ID")
            refs = msg.get("References")

            text, html = _msg_get_text_parts(msg)
            body_text = text or html or ""
            body_text = body_text.strip()

            logger.debug("Lendo UID=%s | From=%s | Subject=%s", uid.decode(), frm, sub)

            # DECISÃO
            logger.debug("[DEBUG] >> tick: iniciando ciclo de leitura IMAP")
            data = decidir_acao_assistente(frm, sub, body_text)

            # Falha REAL (sem nada aproveitável) -> ESCALAR
            if data is None:
                logger.info("[INFO] Ação=escalar conf=0.5")
                move_message(imap, uid, FOLDER_ESCALAR)
                continue

            acao = str(data.get("acao", "")).lower().strip()
            conf = float(data.get("nivel_confianca", 0) or 0.0)
            assunto_resp = data.get("assunto") or f"Re: {sub or ''}"
            corpo_md = data.get("corpo_markdown") or ""

            if data.get("_partial"):
                logger.warning("[WARN] Parser parcial — salvando o possível (acao=%s, conf=%.2f)", acao, conf)

            # Segurança: se ação desconhecida, força escalar
            if acao not in ("responder", "escalar", "ignorar"):
                logger.info("[INFO] Ação desconhecida => escalar")
                move_message(imap, uid, FOLDER_ESCALAR)
                continue

            # Regra de confiança: se abaixo do mínimo e pediu responder, escalar
            if acao == "responder" and conf < CONF_MIN:
                logger.info("[INFO] Confiança baixa (%.2f < %.2f) => escalar", conf, CONF_MIN)
                move_message(imap, uid, FOLDER_ESCALAR)
                continue

            # Se veio responder mas ficou sem corpo (parser parcial), tentar completar com mini chamada
            if acao == "responder" and not corpo_md:
                mini_sys = "Escreva apenas o corpo de um e-mail de resposta em Markdown (PT-BR), objetivo e educado."
                mini_user = f"Assunto original: {sub}\nResumo do e-mail do cliente:\n{body_text[:2000]}"
                corpo_try = openrouter_chat(
                    messages=[{"role": "system", "content": mini_sys}, {"role": "user", "content": mini_user}],
                    model=OPENROUTER_MODEL_WRITER,
                    max_tokens=OPENROUTER_WRITER_MAXTOK,
                    temperature=0.3, top_p=0.9
                )
                if corpo_try:
                    corpo_md = corpo_try.strip()

            # Executar ação
            if acao == "escalar":
                logger.info("[INFO] Chamando move_message -> Escalar (ação=escalar, conf=%.2f)", conf or 0.0)
                move_message(imap, uid, FOLDER_ESCALAR)

            elif acao == "ignorar":
                logger.info("[INFO] Ignorar solicitado (marcar como lido).")
                # marca como \Seen apenas
                imap.uid("STORE", uid, "+FLAGS", r"(\Seen)")

            else:  # responder
                try:
                    to_addr = None
                    # extrair e-mail do From:
                    # formato simples: "Nome <mail@dom.com>" ou apenas "mail@dom.com"
                    m = re.search(r"<([^>]+)>", frm or "")
                    to_addr = m.group(1) if m else (frm or "").strip()

                    send_email_reply(
                        to_addr=to_addr,
                        from_addr=SMTP_USER or IMAP_USER,
                        subject=assunto_resp,
                        body_md=corpo_md or "(resposta gerada automaticamente)",
                        in_reply_to=in_reply_to,
                        references=refs,
                    )
                    logger.info("[INFO] Resposta enviada → movendo para Respondidos")
                    move_message(imap, uid, FOLDER_RESPONDIDOS)
                except Exception as e:
                    logger.warning("[WARN] Falha ao enviar resposta (%s) → Escalar", e)
                    move_message(imap, uid, FOLDER_ESCALAR)

    except Exception as e:
        logger.error("Erro no ciclo IMAP: %s\n%s", e, traceback.format_exc())
    finally:
        try:
            if imap is not None:
                imap.logout()
        except Exception:
            pass


def watch_loop_forever():
    while True:
        try:
            process_unseen_once()
        except Exception as e:
            logger.error("Falha no watcher: %s", e)
        time.sleep(POLL_SECONDS)


# ------------------------------------------------------------------------------
# Flask
# ------------------------------------------------------------------------------

app = Flask(__name__)

@app.route("/")
def index():
    return (
        f"<h1>{APP_NAME}</h1>"
        f"<p>Status: ativo</p>"
        f"<ul>"
        f"<li>/diag/openrouter — ping de configuração</li>"
        f"<li>/diag/openrouter-chat — eco de teste (não consome créditos)</li>"
        f"</ul>"
    )

@app.route("/diag/openrouter")
def diag_openrouter():
    ok = bool(OPENROUTER_API_KEY)
    return jsonify({
        "ok": ok,
        "model": OPENROUTER_MODEL,
        "public_url": APP_PUBLIC_URL or None,
        "msg": "API key presente" if ok else "Falta OPENROUTER_API_KEY"
    })

@app.route("/diag/openrouter-chat")
def diag_openrouter_chat():
    # Mantemos apenas um eco para não gastar créditos
    return jsonify({
        "ok": True,
        "model": OPENROUTER_MODEL,
        "headers_sent": {
            "HTTP-Referer": APP_PUBLIC_URL or f"https://{APP_NAME.replace(' ', '').lower()}.local",
            "X-Title": APP_NAME
        },
        "body_json": {"eco": "teste", "ok": True},
        "body_text": None
    })


# ------------------------------------------------------------------------------
# Main
# ------------------------------------------------------------------------------

if __name__ == "__main__":
    # Sobe o watcher em background
    t = threading.Thread(target=watch_loop_forever, name="imap-watcher", daemon=True)
    t.start()

    # dica de URL pública no log
    if APP_PUBLIC_URL:
        logger.info("App público em: %s", APP_PUBLIC_URL)

    # inicia Flask
    app.run(host="0.0.0.0", port=PORT, debug=False)
