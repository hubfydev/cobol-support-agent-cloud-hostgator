#!/usr/bin/env python3 - v9.5 (atualizado)
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
import email
import imaplib
import smtplib
import hashlib
import logging
import traceback
from datetime import datetime, timezone
from email import policy
from email.message import EmailMessage
from email.utils import formatdate, make_msgid, parsedate_to_datetime
from flask import Flask, jsonify

# ----------------------------
# Config e utilitários
# ----------------------------
APP_PUBLIC_URL = os.getenv("APP_PUBLIC_URL", "http://localhost:10000")
APP_TITLE = os.getenv("APP_TITLE", "COBOL Support Agent")
PORT = int(os.getenv("PORT", "10000"))

IMAP_HOST = os.getenv("IMAP_HOST", "")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
MAIL_USER = os.getenv("MAIL_USER", "")
MAIL_PASS = os.getenv("MAIL_PASS", "")

SMTP_HOST = os.getenv("SMTP_HOST", "")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_TLS_MODE = os.getenv("SMTP_TLS_MODE", "starttls").lower()  # starttls|ssl
SMTP_DEBUG = int(os.getenv("SMTP_DEBUG", "0"))

FOLDER_ESCALATE = os.getenv("FOLDER_ESCALATE", "Escalar")
FOLDER_PROCESSED = os.getenv("FOLDER_PROCESSED", "Respondidos")
SENT_FOLDER = os.getenv("SENT_FOLDER", "INBOX.Sent")
EXPUNGE_AFTER_COPY = os.getenv("EXPUNGE_AFTER_COPY", "true").lower() == "true"

CHECK_INTERVAL_SECONDS = int(os.getenv("CHECK_INTERVAL_SECONDS", "60"))
CONFIDENCE_THRESHOLD = float(os.getenv("CONFIDENCE_THRESHOLD", "0.5"))

LLM_BACKEND = os.getenv("LLM_BACKEND", "openrouter")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openrouter/auto")
OPENROUTER_MAX_TOKENS = int(os.getenv("OPENROUTER_MAX_TOKENS", "512"))
OPENROUTER_APP_NAME = os.getenv("OPENROUTER_APP_NAME", APP_TITLE)
OPENROUTER_SITE_URL = os.getenv("OPENROUTER_SITE_URL", APP_PUBLIC_URL)

SIGNATURE_NAME = os.getenv("SIGNATURE_NAME", "Equipe Aprenda COBOL — Suporte")
SIGNATURE_LINKS = os.getenv("SIGNATURE_LINKS", "https://aprendacobol.com.br/assinatura/").strip()
SIGNATURE_FOOTER = os.getenv(
    "SIGNATURE_FOOTER",
    "Se precisar, responda este e-mail com mais detalhes ou anexe seu arquivo .COB/.CBL.\n"
    "Horário de atendimento: 9h–18h (ET), seg–sex. \n"
    "Conheça nossa Formação Completa de Programadores COBOL, com COBOL Avançado, \n"
    "JCL, Db2 e Bancos de Dados completo em:"
)

TELEGRAM_URL = "https://t.me/aprendacobol"
FORMACAO_URL = "https://assinatura.aprendacobol.com.br"

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(level=getattr(logging, LOG_LEVEL, logging.INFO), format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger("cobol-support-agent")

app = Flask(__name__)

# ----------------------------
# PROMPT do sistema
# ----------------------------
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
    "   se ambíguo/incompleto, 'acao'='escalar' com nivel_confianca<=0.6. "
    "7) **Assunto**: defina o campo 'assunto' exatamente igual ao assunto original do e-mail "
    "(não traduza, não resuma, não invente). Se o original já tiver 'Re:' no início, mantenha como está. "
    "OBS: o sistema adicionará 'Re: ' automaticamente na hora do envio se faltar. "
    "8) Sugestões de comunidade/curso: "
    "   8.1) Só mencione o grupo no Telegram se o usuário mencionar explicitamente termos como 'telegram', 'grupo' ou 'canal', "
    "        OU se houver espaço natural no final da resposta. Se mencionar, inclua o link canônico passado no contexto. "
    "   8.2) Se fizer sentido, mencione em UMA FRASE FINAL a Formação Completa (um único link canônico passado no contexto). "
    "   8.3) Não repita links já presentes na mesma resposta. Uma única menção curta, sem exagero. "
    "9) Se houver arquivo anexo .COB/.CBL/.CPY com código COBOL, priorize analisar o código; cite elementos COBOL "
    "(DIVISION, SECTION, PIC, níveis, I/O, SQLCA etc.). Identifique erros comuns e sugira correções objetivas. "
    "10) Não mude o tema da conversa. Responda ao que foi solicitado, de forma educada e objetiva. "
    "11) Se faltar informação para compilar/executar, peça os dados mínimos (ex.: amostras de entrada/saída, layout, JCL). "
)

PROMPT_SHA1 = hashlib.sha1(SYSTEM_PROMPT.encode("utf-8")).hexdigest()[:12]
logger.info(f"SYSTEM_PROMPT_SHA1={PROMPT_SHA1} (primeiros 120 chars): {SYSTEM_PROMPT[:120].replace(chr(10),' ')}")

# ----------------------------
# Helpers de e-mail
# ----------------------------
def normalize_mailbox(name: str) -> str:
    name = name.strip()
    if not name.lower().startswith("inbox"):
        return f"INBOX.{name}"
    return name

def ensure_mailbox(imap: imaplib.IMAP4_SSL, mailbox: str):
    mailbox = normalize_mailbox(mailbox)
    code, _ = imap.create(mailbox)
    if code == "OK":
        logger.debug(f"Mailbox criada: {mailbox}")
    else:
        logger.debug(f"Mailbox pode já existir: {mailbox}")
    return mailbox

def make_reply_subject(original_subject: str) -> str:
    s = (original_subject or "").strip()
    low = s.lower()
    if low.startswith("re:") or low.startswith("fw:") or low.startswith("fwd:"):
        return s
    return f"Re: {s}"

def parse_email_bytes(data_bytes: bytes):
    msg = email.message_from_bytes(data_bytes, policy=policy.default)
    subject = msg.get("Subject", "")
    from_addr = email.utils.parseaddr(msg.get("From", ""))[1]
    to_addr = email.utils.parseaddr(msg.get("To", ""))[1]
    message_id = msg.get("Message-ID", make_msgid())
    in_reply_to = msg.get("In-Reply-To")
    references = msg.get_all("References", [])
    if in_reply_to:
        references.append(in_reply_to)

    text_parts = []
    html_parts = []
    attachments = []

    for part in msg.walk():
        cdispo = (part.get_content_disposition() or "").lower()
        ctype = part.get_content_type()

        if cdispo == "attachment":
            filename = part.get_filename() or "anexo"
            payload = part.get_payload(decode=True) or b""
            attachments.append((filename, payload))
        else:
            if ctype == "text/plain":
                text_parts.append(part.get_content())
            elif ctype == "text/html":
                html_parts.append(part.get_content())

    body_text = "\n\n".join(text_parts).strip()
    body_html = "\n\n".join(html_parts).strip()

    return {
        "subject": subject,
        "from": from_addr,
        "to": to_addr,
        "message_id": message_id,
        "references": " ".join(references) if references else None,
        "in_reply_to": in_reply_to,
        "body_text": body_text,
        "body_html": body_html,
        "attachments": attachments,
    }

def pick_body_text(parsed):
    if parsed["body_text"]:
        return parsed["body_text"]
    # fallback rudimentar de HTML → texto
    if parsed["body_html"]:
        txt = re.sub(r"<br\s*/?>", "\n", parsed["body_html"], flags=re.I)
        txt = re.sub(r"</p\s*>", "\n\n", txt, flags=re.I)
        txt = re.sub(r"<[^>]+>", "", txt)
        return email.header.decode_header(txt)[0][0] if isinstance(txt, bytes) else txt
    return ""

def has_cobol_attachment(attachments):
    for name, _ in attachments:
        if name and name.lower().endswith((".cob", ".cbl", ".cpy")):
            return True
    return False

def load_cobol_snippets(attachments, max_bytes=60000):
    out = []
    used = 0
    for name, payload in attachments:
        if not name.lower().endswith((".cob", ".cbl", ".cpy")):
            continue
        chunk = payload[: max(0, max_bytes - used)]
        try:
            text = chunk.decode("utf-8", errors="replace")
        except Exception:
            text = chunk.decode("latin-1", errors="replace")
        out.append({"filename": name, "content": text})
        used += len(chunk)
        if used >= max_bytes:
            break
    return out

def connect_imap():
    imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    imap.login(MAIL_USER, MAIL_PASS)
    return imap

def move_msg(imap: imaplib.IMAP4_SSL, msg_id: bytes, dest_folder: str):
    dest = ensure_mailbox(imap, dest_folder)
    imap.copy(msg_id, dest)
    imap.store(msg_id, "+FLAGS", r"(\Deleted)")
    if EXPUNGE_AFTER_COPY:
        imap.expunge()
    logger.info(f"E-mail movido para {dest}")

def append_to_sent(raw_bytes: bytes):
    try:
        imap = connect_imap()
        sent_box = ensure_mailbox(imap, SENT_FOLDER)
        datestr = datetime.now(timezone.utc).strftime("%d-%b-%Y %H:%M:%S +0000")
        imap.append(sent_box, r"(\Seen)", datestr, raw_bytes)
        imap.logout()
        logger.info(f"Mensagem copiada para a pasta de enviados: {sent_box}")
    except Exception:
        logger.error("Falha ao copiar para enviados")
        logger.error(traceback.format_exc())

def smtp_send(msg: EmailMessage):
    context = ssl.create_default_context()
    if SMTP_TLS_MODE == "ssl":
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as s:
            s.set_debuglevel(SMTP_DEBUG)
            s.login(MAIL_USER, MAIL_PASS)
            s.send_message(msg)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.set_debuglevel(SMTP_DEBUG)
            s.starttls(context=context)
            s.login(MAIL_USER, MAIL_PASS)
            s.send_message(msg)

# ----------------------------
# Formatação de corpo/assinatura
# ----------------------------
URL_REGEX = re.compile(r"https?://\S+", re.I)

def ensure_single_link(text: str, url: str) -> str:
    """Garante no máximo uma ocorrência 'texto simples (markdown não-forçado)'. Não duplica se já existe."""
    if not url:
        return text
    if re.search(re.escape(url), text, flags=re.I):
        return text
    # insere como texto simples (sem [link](link)) para evitar duplicação visual em alguns clientes
    if text and not text.endswith("\n"):
        text += "\n"
    return text + url.strip() + "\n"

def render_signature_block() -> str:
    # monta assinatura com quebras esperadas
    block = SIGNATURE_FOOTER.strip()
    # garante que o link de assinatura final apareça uma única vez
    block = ensure_single_link(block + ("\n" if not block.endswith("\n") else ""), SIGNATURE_LINKS)
    block += f"\n— {SIGNATURE_NAME}"
    return block

def finalize_corpo_markdown(body_md: str) -> str:
    body = (body_md or "").strip()

    # Evitar duplicar links canônicos
    for canonical in (TELEGRAM_URL, FORMACAO_URL, SIGNATURE_LINKS):
        # se corpo já menciona, não adicionamos de novo em assinatura
        pass

    # Adiciona assinatura padronizada
    sig = render_signature_block()
    if body and not body.endswith("\n"):
        body += "\n"
    body += "\n" + sig.strip() + "\n"
    return body

# ----------------------------
# OpenRouter
# ----------------------------
import requests

class OpenRouterError(Exception):
    def __init__(self, status, body):
        super().__init__(f"OpenRouter HTTP {status}")
        self.status = status
        self.body = body

def call_openrouter(system_prompt: str, user_prompt: str):
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "HTTP-Referer": OPENROUTER_SITE_URL,
        "X-Title": OPENROUTER_APP_NAME,
        "Content-Type": "application/json",
    }
    payload = {
        "model": OPENROUTER_MODEL,
        "max_tokens": OPENROUTER_MAX_TOKENS,
        "temperature": 0.2,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
    }
    resp = requests.post(url, headers=headers, data=json.dumps(payload), timeout=60)
    logger.debug(f"OpenRouter status={resp.status_code}")
    if resp.status_code != 200:
        raise OpenRouterError(resp.status_code, resp.text)

    data = resp.json()
    content = data["choices"][0]["message"]["content"]
    # extrai JSON estrito
    try:
        # tentativa direta
        parsed = json.loads(content)
        return parsed
    except Exception:
        # fallback: extrair bloco {...}
        m = re.search(r"\{.*\}", content, flags=re.S)
        if not m:
            raise ValueError("Resposta do LLM sem JSON reconhecível.")
        parsed = json.loads(m.group(0))
        return parsed

# ----------------------------
# Decisão e fluxo
# ----------------------------
def should_use_simple_rules(body_text: str, attachments) -> bool:
    """Apenas quando NÃO há COBOL anexo e o texto pede Telegram/grupo/canal."""
    if has_cobol_attachment(attachments):
        return False
    t = (body_text or "").lower()
    return any(k in t for k in ("telegram", "grupo", "canal"))

def build_user_prompt_for_llm(parsed, cobol_snippets):
    subject = parsed["subject"] or ""
    body_text = pick_body_text(parsed)
    # Monta contexto objetivo para o modelo
    ctx = {
        "assunto_original": subject,
        "remetente": parsed["from"],
        "conteudo_texto": body_text,
        "tem_anexo_cobol": bool(cobol_snippets),
        "anexos_cobol": cobol_snippets,  # lista de {filename, content}
        "links_canonicos": {
            "telegram": TELEGRAM_URL,
            "formacao": FORMACAO_URL
        },
        "instrucoes_envio": {
            "nao_mudar_assunto": True,
            "o_sistema_podera_prefixar_re": True
        }
    }
    return json.dumps(ctx, ensure_ascii=False)

def decide_and_respond(imap: imaplib.IMAP4_SSL, msg_id: bytes, raw_bytes: bytes):
    parsed = parse_email_bytes(raw_bytes)
    subject_in = parsed["subject"] or ""
    from_addr = parsed["from"]
    to_addr = parsed["to"]
    body_text = pick_body_text(parsed)
    attachments = parsed["attachments"]
    cobol_snips = load_cobol_snippets(attachments)

    logger.debug(f"Email de {from_addr} / subj='{subject_in}' / anexos={len(attachments)}")

    # 1) Tenta LLM primeiro
    llm_ok = False
    llm_json = None
    if LLM_BACKEND == "openrouter" and OPENROUTER_API_KEY:
        try:
            user_prompt = build_user_prompt_for_llm(parsed, cobol_snips)
            llm_json = call_openrouter(SYSTEM_PROMPT, user_prompt)
            llm_ok = True
        except OpenRouterError as e:
            logger.error(f"LLM falhou: {e.status}")
            logger.debug(e.body)
        except Exception:
            logger.error("LLM error")
            logger.error(traceback.format_exc())

    # 2) Se LLM falhou
    if not llm_ok:
        if has_cobol_attachment(attachments):
            # Não arriscar: escalamos para análise humana
            move_msg(imap, msg_id, FOLDER_ESCALATE)
            return "escalar"
        # Sem anexo COBOL → pode usar resposta simples apenas se pedir Telegram
        if should_use_simple_rules(body_text, attachments):
            logger.info("Regra simples aplicada: acao=responder conf=0.95")
            body_md = (
                f"Claro! Para entrar no nosso grupo de alunos no Telegram, use o link: {TELEGRAM_URL}\n\n"
                f"Fique à vontade para postar dúvidas de exercícios e projetos lá. "
                f"Também recomendo a Formação Completa de Programador COBOL: {FORMACAO_URL}"
            )
            send_reply_and_archive(imap, msg_id, parsed, body_md)
            return "responder"
        # Caso contrário, escalar
        move_msg(imap, msg_id, FOLDER_ESCALATE)
        return "escalar"

    # 3) Temos saída do LLM → validar e aplicar regras locais
    try:
        acao = (llm_json.get("acao") or "").lower().strip()
        nivel = float(llm_json.get("nivel_confianca", 0.0))
        corpo_md = llm_json.get("corpo_markdown") or ""
    except Exception:
        logger.error("JSON do LLM inválido - escalando.")
        move_msg(imap, msg_id, FOLDER_ESCALATE)
        return "escalar"

    # Se modelo mandou responder mas confiança baixa e há COBOL, ainda podemos escalar
    if acao == "responder" and (nivel >= 0.8 or not has_cobol_attachment(attachments)):
        send_reply_and_archive(imap, msg_id, parsed, corpo_md)
        return "responder"
    else:
        move_msg(imap, msg_id, FOLDER_ESCALATE)
        return "escalar"

def build_reply_message(parsed_in, body_md: str):
    reply = EmailMessage()
    reply["From"] = MAIL_USER
    reply["To"] = parsed_in["from"]
    reply["Date"] = formatdate(localtime=True)

    # assunto preservado
    reply["Subject"] = make_reply_subject(parsed_in["subject"])
    if parsed_in["message_id"]:
        reply["In-Reply-To"] = parsed_in["message_id"]
    if parsed_in["references"]:
        reply["References"] = parsed_in["references"]

    final_md = finalize_corpo_markdown(body_md)
    # envia como texto simples (markdown), e opcionalmente como html simples (conversão básica)
    reply.set_content(final_md)

    # HTML opcional minimalista (parágrafos por linha)
    html = "<br>".join(map(lambda l: l if l else "<br>", final_md.split("\n")))
    html = f"<html><body><pre style='white-space:pre-wrap;font-family:system-ui,Segoe UI,Roboto,Arial'>{html}</pre></body></html>"
    reply.add_alternative(html, subtype="html")

    return reply

def send_reply_and_archive(imap, msg_id, parsed_in, body_md):
    msg = build_reply_message(parsed_in, body_md)
    smtp_send(msg)

    # copiar para enviados
    try:
        raw_bytes = msg.as_bytes()
        append_to_sent(raw_bytes)
    except Exception:
        logger.error("Falha no append para enviados (pós-envio).")
        logger.error(traceback.format_exc())

    # mover original para Respondidos
    move_msg(imap, msg_id, FOLDER_PROCESSED)
    logger.info(f"E-mail enviado para {parsed_in['from']} (Subject: {msg['Subject']})")

# ----------------------------
# Loop IMAP
# ----------------------------
def watcher_loop():
    logger.info(f"Watcher IMAP — envio via SMTP {SMTP_HOST}")
    logger.info(f"App público em: {APP_PUBLIC_URL}")
    while True:
        try:
            imap = connect_imap()
            logger.debug("CAPABILITIES: (aferir após LOGIN)")
            typ, _ = imap.select("INBOX")
            if typ != "OK":
                raise RuntimeError("Falha ao selecionar INBOX")

            typ, data = imap.search(None, "UNSEEN")
            unseen = data[0].split()
            logger.debug(f"UNSEEN: {unseen}")

            for msg_id in unseen:
                typ, msgdata = imap.fetch(msg_id, "(RFC822)")
                if typ != "OK":
                    continue
                raw = msgdata[0][1]
                try:
                    decide_and_respond(imap, msg_id, raw)
                except Exception:
                    logger.error("Erro no processamento da mensagem:")
                    logger.error(traceback.format_exc())
                    # Em erro inesperado, não deletar; sai do loop dessa msg

            # limpeza e saída
            imap.expunge()
            imap.logout()
        except Exception:
            logger.error("Falha no loop IMAP:")
            logger.error(traceback.format_exc())

        time.sleep(CHECK_INTERVAL_SECONDS)

# ----------------------------
# Flask endpoints
# ----------------------------
@app.get("/")
def root():
    return f"{APP_TITLE} está rodando. PromptSHA1={PROMPT_SHA1}"

@app.get("/diag/prompt")
def diag_prompt():
    return jsonify({
        "sha1": PROMPT_SHA1,
        "len": len(SYSTEM_PROMPT),
        "preview": SYSTEM_PROMPT[:400],
    })

@app.get("/diag/openrouter-chat")
def diag_openrouter():
    # pequeno ping que não consome tokens demais (semelhante ao seu teste)
    if not OPENROUTER_API_KEY:
        return jsonify({"ok": False, "error": "OPENROUTER_API_KEY ausente"}), 400
    try:
        payload = {
            "model": OPENROUTER_MODEL,
            "max_tokens": 16,
            "temperature": 0.0,
            "messages": [
                {"role": "system", "content": "Responda com um JSON: {\"ok\":true}"},
                {"role": "user", "content": "Confirme status."},
            ],
        }
        headers = {
            "Authorization": f"Bearer {OPENROUTER_API_KEY}",
            "HTTP-Referer": OPENROUTER_SITE_URL,
            "X-Title": OPENROUTER_APP_NAME,
            "Content-Type": "application/json",
        }
        r = requests.post("https://openrouter.ai/api/v1/chat/completions",
                          headers=headers, data=json.dumps(payload), timeout=30)
        ok = r.status_code == 200
        return jsonify({"ok": ok, "status": r.status_code, "body": r.text[:400]})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500

# ----------------------------
# Main
# ----------------------------
if __name__ == "__main__":
    import threading
    t = threading.Thread(target=watcher_loop, daemon=True)
    t.start()
    # Flask
    app.run(host="0.0.0.0", port=PORT)
