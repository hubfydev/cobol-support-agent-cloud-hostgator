# app.py - Versão 6.2 (Render + OpenRouter + IMAP PEEK + diagnósticos extras)
import os
import ssl
import time
import json
import sqlite3
import smtplib
import email
import re
import imaplib
from email import policy
from email.parser import BytesParser
from email.message import EmailMessage
from markdown import markdown
from dotenv import load_dotenv

from flask import Flask, jsonify
from threading import Thread

# ======= PROMPTS/LLM (opcional) =======
try:
    from prompts import SYSTEM_PROMPT, USER_TEMPLATE
except Exception:
    SYSTEM_PROMPT = (
        "Você é um assistente de suporte de um curso de COBOL. "
        "Gere respostas úteis, educadas, e diretas; se houver código, aponte erros específicos. "
        "Analise o e-mail (texto e possível código). Responda educadamente. "
        "Cite pontos do COBOL (DIVISION, SECTION, PIC, níveis, I/O, SQLCA etc.). "
        "Sugira passos práticos. Se no e-mail tiver qualquer solicitação para entrar no grupo do Telegram, dê o link 'https://t.me/aprendacobol'. "
        "Sempre ofereça o curso de Formação Completa de Programadore Aprenda COBOl disponível por assinatura em: https://assinatura.aprendacobol.com.br. "
        "Retorne um JSON com as chaves: assunto, corpo_markdown, nivel_confianca (0..1), acao ('responder'|'escalar')."
    )
    USER_TEMPLATE = (
        "De: {from_addr}\nAssunto: {subject}\n"
        "Texto:\n{plain_text}\n\nCódigo:\n{code_block}\n"
        "Responda em PT-BR. Se o pedido for claro, 'acao'='responder' com nivel_confianca>=0.8.\n"
        "Se estiver ambíguo/faltando anexos, 'acao'='escalar' com nivel_confianca<=0.6."
    )

# LLM local via Ollama (se indisponível no Render, app segue com fallback)
try:
    from ollama_client import OllamaClient
except Exception:
    OllamaClient = None

# ========= Carrega .env =========
load_dotenv()

# -------- IMAP (leitura) --------
IMAP_HOST = os.getenv("IMAP_HOST", "mail.aprendacobol.com.br")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
IMAP_TLS_MODE = os.getenv("IMAP_TLS_MODE", "ssl").lower()  # ssl | starttls
MAIL_USER = os.getenv("MAIL_USER")   # suporte@aprendacobol.com.br
MAIL_PASS = os.getenv("MAIL_PASS")   # senha

# -------- SMTP (envio) --------
SMTP_HOST = os.getenv("SMTP_HOST", "mail.aprendacobol.com.br")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))  # 587(starttls) ou 465(ssl)
SMTP_TLS_MODE = os.getenv("SMTP_TLS_MODE", "starttls").lower()  # starttls|ssl
SMTP_DEBUG_ON = os.getenv("SMTP_DEBUG", "0") == "1"
SENT_FOLDER = os.getenv("SENT_FOLDER", "INBOX.Sent")  # tentaremos normalizar

# -------- LLM (OpenRouter / Ollama) --------
LLM_BACKEND = os.getenv("LLM_BACKEND", "openrouter").lower()  # openrouter|ollama

# OpenRouter
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "meta-llama/llama-3.1-8b-instruct:free")

# Nomes antigos mantidos por compatibilidade:
OPENROUTER_SITE_URL = os.getenv("OPENROUTER_SITE_URL", "")
OPENROUTER_APP_NAME = os.getenv("OPENROUTER_APP_NAME", "COBOL Support Agent")

# Nomes preferidos (com fallback nos antigos):
APP_PUBLIC_URL = os.getenv("APP_PUBLIC_URL", OPENROUTER_SITE_URL)
APP_TITLE = os.getenv("APP_TITLE", OPENROUTER_APP_NAME)

# Ollama (apenas se rodar fora do Render)
OLLAMA_HOST = os.getenv("OLLAMA_HOST", "http://127.0.0.1:11434")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "llama3.1:8b")

# -------- Comportamento --------
CHECK_INTERVAL = int(os.getenv("CHECK_INTERVAL_SECONDS", "60"))
FOLDER_PROCESSED = os.getenv("FOLDER_PROCESSED", "Respondidos")
FOLDER_ESCALATE = os.getenv("FOLDER_ESCALATE", "Escalar")
CONFIDENCE_THRESHOLD = float(os.getenv("CONFIDENCE_THRESHOLD", "0.65"))
EXPUNGE_AFTER_COPY = os.getenv("EXPUNGE_AFTER_COPY", "false").lower() == "true"
LOG_LEVEL = os.getenv("LOG_LEVEL", "info").lower()
PORT = int(os.getenv("PORT", "10000"))

# -------- Assinatura --------
SIGNATURE_NAME = os.getenv("SIGNATURE_NAME", "Equipe Aprenda COBOL — Suporte")
SIGNATURE_FOOTER = os.getenv(
    "SIGNATURE_FOOTER",
    "Se precisar, responda este e-mail com mais detalhes ou anexe seu arquivo .COB/.CBL.\nHorário de atendimento: 9h–18h (ET), seg–sex."
)
SIGNATURE_LINKS = os.getenv("SIGNATURE_LINKS", "")

DB_PATH = "state.db"

# ========= Utils =========
def log(level, *args):
    levels = {"debug": 0, "info": 1, "warn": 2, "error": 3}
    if levels.get(level, 1) >= levels.get(LOG_LEVEL, 1):
        print(f"[{level.upper()}]", *args)

def require_env():
    missing = []
    for k in ["IMAP_HOST", "MAIL_USER", "MAIL_PASS", "SMTP_HOST", "SMTP_PORT"]:
        if not globals().get(k):
            missing.append(k)
    if missing:
        raise RuntimeError("Faltam variáveis no ambiente: " + ", ".join(missing))

def db_init():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS processed (message_id TEXT PRIMARY KEY)")
    con.commit()
    con.close()

def already_processed(msgid: str) -> bool:
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("SELECT 1 FROM processed WHERE message_id=?", (msgid,))
    row = cur.fetchone()
    con.close()
    return row is not None

def mark_processed(msgid: str):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("INSERT OR IGNORE INTO processed(message_id) VALUES (?)", (msgid,))
    con.commit()
    con.close()

# ========= IMAP =========
def connect_imap():
    """
    Conecta e faz login no IMAP suportando SSL:993 e STARTTLS:143,
    com log detalhado quando LOG_LEVEL=debug.
    """
    ctx = ssl.create_default_context()
    if IMAP_TLS_MODE == "ssl":
        imap = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=ctx)
        if LOG_LEVEL == "debug":
            imap.debug = 4
        imap.login(MAIL_USER, MAIL_PASS)
        return imap
    elif IMAP_TLS_MODE == "starttls":
        imap = imaplib.IMAP4(IMAP_HOST, IMAP_PORT)
        if LOG_LEVEL == "debug":
            imap.debug = 4
        imap.starttls(ssl_context=ctx)
        imap.login(MAIL_USER, MAIL_PASS)
        return imap
    else:
        raise RuntimeError(f"IMAP_TLS_MODE inválido: {IMAP_TLS_MODE}")

def select_inbox(imap):
    typ, _ = imap.select("INBOX")
    if typ != "OK":
        raise RuntimeError("Não foi possível selecionar INBOX")

def fetch_unseen(imap):
    typ, data = imap.search(None, 'UNSEEN')
    if typ != "OK":
        return []
    return data[0].split()

def parse_message(raw_bytes):
    msg = BytesParser(policy=policy.default).parsebytes(raw_bytes)
    msgid = msg.get("Message-ID") or msg.get("Message-Id") or ""
    from_addr = email.utils.parseaddr(msg.get("From"))[1]
    subject = msg.get("Subject", "")
    plain_parts, code_chunks = [], []

    def walk(m):
        if m.is_multipart():
            for part in m.iter_parts():
                walk(part)
        else:
            ctype = m.get_content_type()
            filename = m.get_filename()
            payload = m.get_payload(decode=True) or b""
            try:
                text = payload.decode(m.get_content_charset() or "utf-8", errors="ignore")
            except:
                text = ""
            if filename and filename.lower().endswith((".cob", ".cbl", ".txt")):
                code_chunks.append(f"--- {filename} ---\n{text}")
            elif ctype == "text/plain":
                plain_parts.append(text)
            elif ctype == "text/html" and not plain_parts:
                import re as _re
                plain_parts.append(_re.sub("<[^<]+?>", "", text))
    walk(msg)
    plain_text = "\n".join(plain_parts).strip()
    code_block = ""
    if code_chunks:
        code_block = "```\n" + "\n\n".join(code_chunks) + "\n```"
    elif "IDENTIFICATION DIVISION" in plain_text.upper():
        code_block = "```cobol\n" + plain_text + "\n```"
    return msg, msgid, from_addr, subject, plain_text, code_block

def guess_first_name(from_addr: str) -> str:
    local = from_addr.split("@")[0]
    local = re.sub(r"[._\-]+", " ", local).strip()
    parts = local.split()
    name = parts[0].capitalize() if parts else ""
    if name.lower() in {"contato", "aluno", "suporte", "noreply", "no"}:
        return ""
    return name

# ========= LIST/parse helpers =========
def _parse_list_line(line: str):
    """
    Parseia uma linha de LIST IMAP no formato: (<flags>) "<delim>" <name>
    Retorna (flags, delim, name) ou (None, None, None) se falhar.
    """
    m = re.search(r'\((?P<flags>.*?)\)\s+"(?P<delim>[^"]+)"\s+(?P<name>.*)$', line.strip())
    if not m:
        return None, None, None
    flags = m.group("flags").strip()
    delim = m.group("delim")
    name = m.group("name").strip()
    if name.startswith('"') and name.endswith('"'):
        name = name[1:-1]
    return flags, delim, name

_listed_boxes_printed = False

def _list_mailboxes_once(imap):
    global _listed_boxes_printed
    boxes = {}
    if _listed_boxes_printed:
        return boxes
    try:
        typ, data = imap.list()
        if typ == "OK":
            print("[DEBUG] LIST mailboxes:")
            for raw in (data or []):
                line = raw.decode(errors="ignore")
                print("   ", line)
                flags, delim, name = _parse_list_line(line)
                if name:
                    boxes[name] = {"flags": flags, "delim": delim}
        else:
            print("[WARN] LIST não retornou OK:", data)
    except Exception as e:
        print("[WARN] Falha ao listar mailboxes:", e)
    _listed_boxes_printed = True
    return boxes

# ========= Mover robusto =========
def move_message(imap, num, dest_folder):
    """
    Move a mensagem para dest_folder:
    1) tenta COPY por número sequencial
    2) tenta UID COPY (pega UID antes)
    3) marca como \\Deleted (expurge ao final do ciclo se EXPUNGE_AFTER_COPY)
    (Como usamos BODY.PEEK[], a mensagem permanece UNSEEN ao mover)
    """
    existing = _list_mailboxes_once(imap)
    candidates = [dest_folder]
    for sep in ("/", "."):
        if not dest_folder.upper().startswith("INBOX"):
            candidates.append(f"INBOX{sep}{dest_folder}")

    existing_names = list(existing.keys())
    exact = [n for n in existing_names if n.lower() == dest_folder.lower()]
    ends = [n for n in existing_names if n.lower().endswith(dest_folder.lower())]
    ordered = []
    ordered += [n for n in exact if n not in ordered]
    ordered += [n for n in ends if n not in ordered]
    ordered += [c for c in candidates if c not in ordered]

    last_err = None

    for mb in ordered:
        try:
            imap.create(mb)
        except Exception:
            pass
        log("info", f"Tentando copiar para: {mb}")
        typ, resp = imap.copy(num, mb)
        log("debug", f"IMAP COPY -> typ={typ} resp={resp}")
        if typ == "OK":
            typ2, resp2 = imap.store(num, '+FLAGS', '\\Deleted')
            log("debug", f"IMAP STORE Deleted -> typ={typ2} resp={resp2}")
            if typ2 == "OK":
                return True
            last_err = (typ2, resp2)
        else:
            last_err = (typ, resp)

    # Fallback UID
    try:
        typ_uid, data_uid = imap.fetch(num, '(UID)')
        uid = None
        if typ_uid == "OK" and data_uid and data_uid[0]:
            m = re.search(rb'UID\s+(\d+)', data_uid[0])
            if m:
                uid = m.group(1).decode()
        if uid:
            for mb in ordered:
                log("info", f"Tentando UID COPY para: {mb} (uid={uid})")
                typ, resp = imap.uid('COPY', uid, mb)
                log("debug", f"IMAP UID COPY -> typ={typ} resp={resp}")
                if typ == "OK":
                    typ2, resp2 = imap.uid('STORE', uid, '+FLAGS', '(\\Deleted)')
                    log("debug", f"IMAP UID STORE Deleted -> typ={typ2} resp={resp2}")
                    if typ2 == "OK":
                        return True
                    last_err = (typ2, resp2)
                else:
                    last_err = (typ, resp)
    except Exception as e:
        log("warn", f"Falha no fallback UID COPY: {e}")

    log("warn", f"Falha ao mover para {dest_folder}. Último erro: {last_err}")
    return False

# ========= SMTP envio =========
def smtp_send(message: EmailMessage):
    if SMTP_TLS_MODE == "ssl":
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as smtp:
            if SMTP_DEBUG_ON:
                smtp.set_debuglevel(1)
            smtp.login(MAIL_USER, MAIL_PASS)
            smtp.send_message(message)
    else:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            if SMTP_DEBUG_ON:
                smtp.set_debuglevel(1)
            smtp.ehlo()
            smtp.starttls(context=ssl.create_default_context())
            smtp.ehlo()
            smtp.login(MAIL_USER, MAIL_PASS)
            smtp.send_message(message)

# ========= Append em Enviados =========
def append_to_sent(imap_host, imap_port, user, pwd, sent_folder_name, msg):
    try:
        ctx = ssl.create_default_context()
        im = imaplib.IMAP4_SSL(imap_host, imap_port, ssl_context=ctx)
        im.login(user, pwd)
        existing = {}
        try:
            typ, data = im.list()
            if typ == "OK":
                for raw in (data or []):
                    line = raw.decode(errors="ignore")
                    _, _, name = _parse_list_line(line)
                    if name:
                        existing[name] = True
        except Exception:
            pass
        candidates = []
        if sent_folder_name:
            candidates.append(sent_folder_name)
        for n in ("INBOX.Sent", "INBOX.Enviados", "Sent", "Enviados"):
            if n not in candidates:
                candidates.append(n)
        ordered = sorted(candidates, key=lambda x: (x not in existing, len(x)))
        dest = ordered[0]
        try:
            im.create(dest)
        except Exception:
            pass
        im.append(dest, "", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        im.logout()
        log("debug", f"Cópia enviada para pasta de enviados: {dest}")
    except Exception as e:
        log("warn", "Falha ao APPEND em Enviados:", e)

# ========= Assunto/assinatura =========
def make_reply_subject(original_subject: str) -> str:
    s = (original_subject or "").strip()
    if s[:3].lower() == "re:":
        return "Re:" + s[3:]
    if s.lower().startswith("re :"):
        return "Re:" + s[4:]
    return f"Re: {s}" if s else "Re:"

def wrap_with_signature(first_name: str, body_markdown: str) -> str:
    saud = f"Olá{', ' + first_name if first_name else ''}!\n\n"
    sig_lines = ["\n---", f"**{SIGNATURE_NAME}**"]
    if SIGNATURE_FOOTER:
        sig_lines.append(SIGNATURE_FOOTER)
    if SIGNATURE_LINKS:
        sig_lines.append(SIGNATURE_LINKS)
    return saud + body_markdown.strip() + "\n" + "\n".join(sig_lines) + "\n"

def send_reply(original_msg, to_addr, reply_subject, body_markdown):
    body_html = markdown(body_markdown)

    reply = EmailMessage()
    reply["Subject"] = reply_subject
    reply["From"] = MAIL_USER
    reply["To"] = to_addr
    if original_msg.get("Message-ID"):
        reply["In-Reply-To"] = original_msg["Message-ID"]
        reply["References"] = original_msg["Message-ID"]

    reply.set_content(body_markdown)
    reply.add_alternative(body_html, subtype="html")

    log("info", f"Enviando resposta (SMTP {SMTP_TLS_MODE.upper()}) para {to_addr}…")
    smtp_send(reply)
    log("info", "Resposta enviada com sucesso.")
    try:
        append_to_sent(IMAP_HOST, IMAP_PORT, MAIL_USER, MAIL_PASS, SENT_FOLDER, reply)
    except Exception as e:
        log("warn", "Não foi possível salvar cópia em Enviados:", e)

# ========= IA (OpenRouter / Ollama / Fallback) =========
def _ask_openrouter(system_prompt: str, user_prompt: str):
    import requests
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    if APP_PUBLIC_URL:
        headers["HTTP-Referer"] = APP_PUBLIC_URL
        headers["Referer"] = APP_PUBLIC_URL
    if APP_TITLE:
        headers["X-Title"] = APP_TITLE

    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.2,
        "response_format": {"type": "json_object"}
    }
    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code in (402, 429):
        raise RuntimeError(f"OpenRouter limit: {r.status_code} {r.text[:200]}")
    r.raise_for_status()
    data = r.json()
    content = data["choices"][0]["message"]["content"]
    return json.loads(content)

def call_agent_local(from_addr, subject, plain_text, code_block):
    user_prompt = USER_TEMPLATE.format(
        from_addr=from_addr, subject=subject,
        plain_text=plain_text[:8000], code_block=code_block[:8000]
    )

    # 1) OpenRouter (Render-friendly; tem tier grátis)
    if LLM_BACKEND == "openrouter" and OPENROUTER_API_KEY:
        try:
            data = _ask_openrouter(SYSTEM_PROMPT, user_prompt)
            return {
                "assunto": data.get("assunto", f"Re: {subject[:200]}"),
                "corpo_markdown": data.get("corpo_markdown", "Obrigado pelo contato!"),
                "nivel_confianca": float(data.get("nivel_confianca", 0.0)),
                "acao": data.get("acao", "escalar")
            }
        except Exception as e:
            log("warn", "Falha no OpenRouter:", e)

    # 2) Ollama local (para rodar no seu PC/servidor com Ollama)
    if LLM_BACKEND == "ollama" and OllamaClient is not None:
        try:
            client = OllamaClient(OLLAMA_HOST, OLLAMA_MODEL)
            data = client.generate_json(SYSTEM_PROMPT, user_prompt)
            return {
                "assunto": data.get("assunto", f"Re: {subject[:200]}"),
                "corpo_markdown": data.get("corpo_markdown", "Obrigado pelo contato!"),
                "nivel_confianca": float(data.get("nivel_confianca", 0.0)),
                "acao": data.get("acao", "escalar")
            }
        except Exception as e:
            log("warn", "Falha no Ollama:", e)

    # 3) Fallback (sem IA)
    body = (
        "- Obrigado por enviar seu código/erro.\n"
        "- Verifique divisões (IDENTIFICATION/DATA/PROCEDURE), níveis e PIC.\n"
        "- Para I/O, confirme OPEN/READ/WRITE/CLOSE e status codes.\n"
        "- Se puder, anexe seu .COB/.CBL para revisão pontual.\n"
    )
    return {
        "assunto": f"Re: {subject[:200]}",
        "corpo_markdown": body,
        "nivel_confianca": 0.5,
        "acao": "escalar"
    }

# ========= Loop principal =========
def main_loop():
    require_env()
    print("Watcher IMAP — envio via SMTP HostGator")
    db_init()
    while True:
        log("debug", ">> tick: iniciando ciclo de leitura IMAP")
        try:
            imap = connect_imap()
            select_inbox(imap)
            ids = fetch_unseen(imap)
            log("debug", f"UNSEEN: {ids}")
            for num in ids:
                # >>> NÃO marca como lido: mantém UNSEEN ao mover <<<
                typ, data = imap.fetch(num, '(BODY.PEEK[])')
                if typ != "OK" or not data or not data[0]:
                    continue
                raw = data[0][1]
                msg, msgid, from_addr, subject, plain_text, code_block = parse_message(raw)
                if not msgid:
                    msgid = f"no-id-{num.decode()}-{int(time.time())}"
                if already_processed(msgid):
                    continue

                ai = call_agent_local(from_addr, subject, plain_text, code_block)
                action = ai.get("acao", "escalar")
                confidence = float(ai.get("nivel_confianca", 0.0))
                log("info", f"Ação={action} conf={confidence}")

                if action == "responder" and confidence >= CONFIDENCE_THRESHOLD:
                    first = guess_first_name(from_addr)
                    full_body = wrap_with_signature(first, ai["corpo_markdown"])
                    reply_subject = make_reply_subject(subject)
                    log("info", f"Assunto final (reply): {reply_subject}")
                    send_reply(msg, from_addr, reply_subject, full_body)

                    log("info", f"Chamando move_message -> {FOLDER_PROCESSED}")
                    ok = move_message(imap, num, FOLDER_PROCESSED)
                    if not ok:
                        log("warn", f"Não consegui mover para {FOLDER_PROCESSED}. Fallback: {FOLDER_ESCALATE}")
                        move_message(imap, num, FOLDER_ESCALATE)
                else:
                    log("info", f"Chamando move_message -> {FOLDER_ESCALATE} (ação={action}, conf={confidence})")
                    move_message(imap, num, FOLDER_ESCALATE)

                mark_processed(msgid)

            if EXPUNGE_AFTER_COPY:
                log("debug", "Executando EXPUNGE…")
                imap.expunge()
            imap.logout()
        except Exception as e:
            log("error", "Erro no loop:", e)
        time.sleep(CHECK_INTERVAL)

# ========= HTTP (Render Free) =========
def imap_self_check():
    try:
        imap = connect_imap()
        select_inbox(imap)
        typ, data = imap.search(None, 'UNSEEN')
        unseen = (data[0].split() if typ == "OK" else [])
        count = len(unseen)
        imap.logout()
        return True, f"IMAP OK. UNSEEN={count}"
    except Exception as e:
        return False, f"IMAP FAIL: {repr(e)}"

def create_http_app():
    app = Flask(__name__)

    @app.get("/")
    def index():
        return "COBOL Support Agent v6.2 - online", 200

    @app.get("/health")
    def health():
        return jsonify({"status": "ok"}), 200

    @app.get("/status")
    def status():
        return jsonify({
            "imap_host": IMAP_HOST,
            "imap_port": IMAP_PORT,
            "imap_tls_mode": IMAP_TLS_MODE,
            "smtp_host": SMTP_HOST,
            "smtp_port": SMTP_PORT,
            "smtp_mode": SMTP_TLS_MODE,
            "backend": LLM_BACKEND,
            "model_openrouter": OPENROUTER_MODEL,
            "model_ollama": OLLAMA_MODEL,
            "processed_folder": FOLDER_PROCESSED,
            "escalate_folder": FOLDER_ESCALATE
        }), 200

    @app.get("/diag/imap")
    def diag_imap():
        ok, msg = imap_self_check()
        code = 200 if ok else 500
        return jsonify({"ok": ok, "msg": msg}), code

    @app.get("/diag/env")
    def diag_env():
        return jsonify({
            "imap_host": IMAP_HOST,
            "imap_port": IMAP_PORT,
            "imap_tls_mode": IMAP_TLS_MODE,
            "mail_user": MAIL_USER,
            "mail_user_hex": (MAIL_USER or "").encode("utf-8").hex(),
            "mail_pass_len": len(MAIL_PASS or ""),
            "smtp_host": SMTP_HOST,
            "smtp_port": SMTP_PORT,
            "smtp_tls_mode": SMTP_TLS_MODE,
            "llm_backend": LLM_BACKEND,
            "openrouter_model": OPENROUTER_MODEL if OPENROUTER_API_KEY else "(sem chave)",
            "app_public_url": APP_PUBLIC_URL,
            "app_title": APP_TITLE,
        }), 200

    @app.get("/diag/imap-auth")
    def diag_imap_auth():
        try:
            im = connect_imap()  # já faz login
            im.logout()
            return jsonify({"ok": True, "msg": "LOGIN OK"}), 200
        except Exception as e:
            return jsonify({"ok": False, "error": repr(e)}), 500

    # ===== OpenRouter: teste rápido de GET (autorização e headers) =====
    @app.get("/diag/openrouter")
    def diag_openrouter():
        try:
            import requests
            url = "https://openrouter.ai/api/v1/models"
            headers = {
                "Authorization": f"Bearer {OPENROUTER_API_KEY}",
                "Accept": "application/json",
            }
            if APP_PUBLIC_URL:
                headers["HTTP-Referer"] = APP_PUBLIC_URL
                headers["Referer"] = APP_PUBLIC_URL
            if APP_TITLE:
                headers["X-Title"] = APP_TITLE

            r = requests.get(url, headers=headers, timeout=30)
            data = None
            try:
                data = r.json()
            except Exception:
                pass
            return jsonify({
                "ok": r.ok,
                "status": r.status_code,
                "models_count": (len(data.get("data", [])) if isinstance(data, dict) and "data" in data else None),
                "headers_used": {k: headers[k] for k in headers if k in ["HTTP-Referer", "Referer", "X-Title"]},
                "error_preview": (r.text[:500] if not r.ok else None)
            }), (200 if r.ok else 500)
        except Exception as e:
            return jsonify({"ok": False, "error": str(e)}), 500

    # ===== OpenRouter: POST de chat completo (debug do 404) =====
    @app.get("/diag/openrouter-chat")
    def diag_openrouter_chat():
        try:
            import requests
            url = "https://openrouter.ai/api/v1/chat/completions"
            headers = {
                "Authorization": f"Bearer {OPENROUTER_API_KEY}",
                "Content-Type": "application/json",
                "Accept": "application/json",
            }
            if APP_PUBLIC_URL:
                headers["HTTP-Referer"] = APP_PUBLIC_URL
                headers["Referer"] = APP_PUBLIC_URL
            if APP_TITLE:
                headers["X-Title"] = APP_TITLE

            payload = {
                "model": OPENROUTER_MODEL,
                "messages": [
                    {"role": "system", "content": "Responda em JSON com {\"ok\": true}."},
                    {"role": "user", "content": "Ping?"}
                ],
                "temperature": 0.0,
                "response_format": {"type": "json_object"}
            }

            r = requests.post(url, headers=headers, json=payload, timeout=45)
            try:
                body_json = r.json()
            except Exception:
                body_json = None

            return jsonify({
                "ok": r.ok,
                "status": r.status_code,
                "model": OPENROUTER_MODEL,
                "headers_sent": {k: headers[k] for k in ["HTTP-Referer", "X-Title", "Referer"] if k in headers},
                "body_json": body_json,
                "body_text": (r.text[:800] if not body_json else None)
            }), (200 if r.ok else 500)
        except Exception as e:
            return jsonify({"ok": False, "error": str(e)}), 500

    return app

def run_watcher():
    try:
        main_loop()
    except BaseException as e:
        log("error", "Watcher encerrou com erro crítico:", e)
        raise

if __name__ == "__main__":
    t = Thread(target=run_watcher, daemon=True)
    t.start()
    app = create_http_app()
    app.run(host="0.0.0.0", port=PORT)

# depois dos outros envs:
OPENROUTER_MAX_TOKENS = int(os.getenv("OPENROUTER_MAX_TOKENS", "256"))

# ...

def _ask_openrouter(system_prompt: str, user_prompt: str):
    import requests
    url = "https://openrouter.ai/api/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    if APP_PUBLIC_URL:
        headers["HTTP-Referer"] = APP_PUBLIC_URL
        headers["Referer"] = APP_PUBLIC_URL
    if APP_TITLE:
        headers["X-Title"] = APP_TITLE

    payload = {
        "model": OPENROUTER_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 0.2,
        "response_format": {"type": "json_object"},
        "max_tokens": OPENROUTER_MAX_TOKENS,   # <<< limite de saída
    }

    r = requests.post(url, headers=headers, json=payload, timeout=60)
    if r.status_code in (402, 429):
        raise RuntimeError(f"OpenRouter limit: {r.status_code} {r.text[:200]}")
    r.raise_for_status()
    data = r.json()
    content = data["choices"][0]["message"]["content"]
    return json.loads(content)
