#!/usr/bin/env python3 - v10.15.1
# -*- coding: utf-8 -*-

"""
COBOL Support Agent — IMAP watcher + SMTP sender + OpenRouter
– v10.15.1:
  * (NOVO) IMAP_USER/IMAP_PASS para separar credenciais de leitura (IMAP) do MAIL_USER/MAIL_PASS.
  * (NOVO) /diag/imap/auth: testa autenticação IMAP (login/logout) e retorna o erro literal.
  * Ajustes internos para usar helper centralizado de credenciais IMAP.
– v10.15:
  * Envio primário por API da Mailgun (se MAILGUN_API_KEY/MAILGUN_DOMAIN presentes).
  * Fallback automático para SMTP (v10.14) com IPv4 preferível, probe e unblock.
  * Novo /diag/email (envio real via transporte primário + fallback) e /diag/transport/status.
  * Mantém /diag/smtp/probe, /diag/smtp/unblock, cooldown SMTP só para falhas temporárias.
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
import socket

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

# Credenciais “genéricas” (também usadas como From).
MAIL_USER = os.getenv("MAIL_USER")  # Ex.: suporte@aprendacobol.com.br
MAIL_PASS = os.getenv("MAIL_PASS")

# (NOVO) Overrides específicos para IMAP (se não definidos, usa MAIL_USER/MAIL_PASS)
IMAP_USER = os.getenv("IMAP_USER", "").strip() or None
IMAP_PASS = os.getenv("IMAP_PASS", "").strip() or None

def _imap_creds():
    """Retorna (user, pass) para login IMAP, respeitando overrides IMAP_USER/IMAP_PASS."""
    user = IMAP_USER if IMAP_USER else MAIL_USER
    pw   = IMAP_PASS if IMAP_PASS else MAIL_PASS
    return user, pw

FOLDER_PROCESSED = os.getenv("FOLDER_PROCESSED", "Respondidos")
FOLDER_ESCALATE = os.getenv("FOLDER_ESCALATE", "Escalar")
EXPUNGE_AFTER_COPY = os.getenv("EXPUNGE_AFTER_COPY", "true").lower() == "true"
SENT_FOLDER = os.getenv("SENT_FOLDER", "INBOX.Sent")
IMAP_FOLDER_INBOX = os.getenv("IMAP_FOLDER_INBOX", "INBOX")
IMAP_FALLBACK_LAST_N = int(os.getenv("IMAP_FALLBACK_LAST_N", "0"))
IMAP_STRICT_UNSEEN_ONLY = os.getenv("IMAP_STRICT_UNSEEN_ONLY", "true").lower() == "true"
IMAP_SINCE_DAYS = int(os.getenv("IMAP_SINCE_DAYS", "0"))
IMAP_FALLBACK_WHEN_LLM_BLOCKED = os.getenv("IMAP_FALLBACK_WHEN_LLM_BLOCKED", "false").lower() == "true"

# SMTP (fallback)
SMTP_HOSTS = os.getenv("SMTP_HOSTS", "")
SMTP_HOST = os.getenv("SMTP_HOST")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_TLS_MODE = os.getenv("SMTP_TLS_MODE", "starttls").lower()  # starttls|ssl
SMTP_DEBUG = int(os.getenv("SMTP_DEBUG", "0"))
SMTP_CONNECT_TIMEOUT = int(os.getenv("SMTP_CONNECT_TIMEOUT", os.getenv("SMTP_TIMEOUT", "12")))
SMTP_TIMEOUT = SMTP_CONNECT_TIMEOUT
SMTP_FALLBACKS = os.getenv("SMTP_FALLBACKS", "465:ssl,2525:starttls")
SMTP_PREFER_IPV4 = os.getenv("SMTP_PREFER_IPV4", os.getenv("SMTP_FORCE_IPV4", "false")).lower() == "true"
SMTP_COOLDOWN_SECONDS = int(os.getenv("SMTP_COOLDOWN_SECONDS", "900"))
_smtp_block_until_ts = 0.0
_last_smtp_error = ""

def _smtp_is_blocked_now() -> bool:
    return time.time() < _smtp_block_until_ts

def _smtp_block(reason: str, seconds: int):
    global _smtp_block_until_ts, _last_smtp_error
    _last_smtp_error = f"{reason} (cooldown {seconds}s)"
    _smtp_block_until_ts = time.time() + max(0, seconds)
    log.warning("SMTP bloqueado: %s", _last_smtp_error)

# Mailgun API (primário)
MAILGUN_API_KEY = os.getenv("MAILGUN_API_KEY", "").strip()
MAILGUN_DOMAIN = os.getenv("MAILGUN_DOMAIN", "").strip()       # ex.: mg.aprendacobol.com.br
MAILGUN_API_BASE = os.getenv("MAILGUN_API_BASE", "https://api.mailgun.net/v3").rstrip("/")
MAIL_PRIMARY = os.getenv("MAIL_PRIMARY", "").strip().lower()    # "mailgun_api" | "smtp" | "" (auto)

def _mail_primary_transport() -> str:
    # Se explicitado, respeita; senão, AUTO: usa Mailgun API se credenciais presentes
    if MAIL_PRIMARY in ("mailgun_api", "smtp"):
        return MAIL_PRIMARY
    return "mailgun_api" if (MAILGUN_API_KEY and MAILGUN_DOMAIN) else "smtp"

# LLM / OpenRouter
LLM_BACKEND = os.getenv("LLM_BACKEND", "openrouter")
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
OPENROUTER_MODEL = os.getenv("OPENROUTER_MODEL", "openrouter/auto")
OPENROUTER_MODEL_FALLBACK = os.getenv("OPENROUTER_MODEL_FALLBACK", "openrouter/auto")
OPENROUTER_MAX_TOKENS = int(os.getenv("OPENROUTER_MAX_TOKENS", "512"))
OPENROUTER_APP_NAME = os.getenv("OPENROUTER_APP_NAME", "COBOL Support Agent")
OPENROUTER_SITE_URL = os.getenv("OPENROUTER_SITE_URL", "")
OPENROUTER_TIMEOUT = int(os.getenv("OPENROUTER_TIMEOUT", "30"))
CONFIDENCE_THRESHOLD = float(os.getenv("CONFIDENCE_THRESHOLD", "0.8"))

LLM_COOLDOWN_SECONDS = int(os.getenv("LLM_COOLDOWN_SECONDS", "900"))
LLM_DISABLE_ON_402 = os.getenv("LLM_DISABLE_ON_402", "true").lower() == "true"
LLM_HARD_DISABLE = os.getenv("LLM_HARD_DISABLE", "false").lower() == "true"
_llm_block_until_ts = 0.0
_last_llm_error = ""

# Pré-triagem
AUTO_ESCALATE_FROM_REGEX = os.getenv(
    "AUTO_ESCALATE_FROM_REGEX",
    r"(?i)(^mailer-daemon@|^postmaster@|^no[-_\. ]?reply@|^noreply@|^2fa@|@hotmart\.com(\.br)?$)"
)
AUTO_ESCALATE_SUBJECT_REGEX = os.getenv(
    "AUTO_ESCALATE_SUBJECT_REGEX",
    r"(?i)(verification code|delivery status|failure notice|bounce|sale made|refund|payment|invoice|assinatura|pagamento atrasado|chargeback)"
)

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
).replace("\\n", "\n")
SIGNATURE_LINKS = os.getenv("SIGNATURE_LINKS", "https://aprendacobol.com.br/assinatura/")

# ==========================
# Prompt do sistema
# ==========================
SYSTEM_PROMPT = (
    "Você é um assistente do time de suporte de um curso de COBOL da Aprenda COBOL. "
    "E-mails da Hotmart ou originados com o remetente 'noreply' não devem ser respondidos, devem ser ignorados por você."
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
    "9) Não mude o tema da conversa. Responda ao que foi solicitado, de forma educada e objetiva, sempre como parte de um time (nós). "
    "10) Se faltar informação para compilar/executar, peça os dados mínimos (ex.: amostras de entrada/saída, layout, JCL). "
    "11) No final do 'corpo_markdown', SEMPRE inclua exatamente estas duas linhas: "
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
                try: text_parts.append(part.get_content())
                except Exception: pass
            elif ctype == "text/html" and "attachment" not in disp:
                try: html_parts.append(part.get_content())
                except Exception: pass
    else:
        ctype = msg.get_content_type()
        if ctype == "text/plain":
            try: text_parts.append(msg.get_content())
            except Exception: pass
        elif ctype == "text/html":
            try: html_parts.append(msg.get_content())
            except Exception: pass

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
                filename = decode_mime_words(part.get_filename())
                if not filename: continue
                lower = filename.lower()
                if lower.endswith((".cob", ".cbl", ".cpy")):
                    try:
                        data = part.get_payload(decode=True)
                        if not data: continue
                        snippet = data[:max_bytes].decode("utf-8", errors="replace")
                        cobol_files.append((filename, snippet))
                    except Exception:
                        continue
    return cobol_files

# ==========================
# OpenRouter client
# ==========================
import requests
from requests.exceptions import RequestException

def _post_openrouter(payload):
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
        "HTTP-Referer": OPENROUTER_SITE_URL or APP_PUBLIC_URL or "",
        "X-Title": OPENROUTER_APP_NAME,
    }
    return requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers=headers,
        data=json.dumps(payload),
        timeout=OPENROUTER_TIMEOUT,
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
        return p

    try:
        r = _post_openrouter(_make_payload(compat=False))
    except RequestException as e:
        _llm_block(f"OpenRouter network error: {e}", LLM_COOLDOWN_SECONDS)
        raise RuntimeError("OpenRouter network error")

    log.debug("OpenRouter status=%s", r.status_code)

    if r.status_code == 402 and LLM_DISABLE_ON_402:
        _llm_block("OpenRouter 402 (limite/rota indisponível)", LLM_COOLDOWN_SECONDS)
        raise RuntimeError(f"OpenRouter HTTP {r.status_code}")

    if r.status_code == 400:
        try: log.error("OpenRouter 400 body: %s", r.text[:500])
        except Exception: pass
        try:
            r = _post_openrouter(_make_payload(compat=True))
        except RequestException as e:
            _llm_block(f"OpenRouter network error (compat): {e}", LLM_COOLDOWN_SECONDS)
            raise RuntimeError("OpenRouter network error (compat)")
        log.debug("OpenRouter (compat) status=%s", r.status_code)

    if r.status_code in (404, 429, 500):
        payload = _make_payload(compat=True)
        payload["model"] = OPENROUTER_MODEL_FALLBACK
        try:
            r = _post_openrouter(payload)
        except RequestException as e:
            _llm_block(f"OpenRouter network error (fallback): {e}", LLM_COOLDOWN_SECONDS)
            raise RuntimeError("OpenRouter network error (fallback)")
        log.debug("OpenRouter fallback status=%s", r.status_code)
        if r.status_code == 402 and LLM_DISABLE_ON_402:
            _llm_block("OpenRouter 402 no fallback", LLM_COOLDOWN_SECONDS)
            raise RuntimeError(f"OpenRouter HTTP {r.status_code}")

    if r.status_code != 200:
        try: log.error("OpenRouter erro %s, corpo: %s", r.status_code, r.text[:500])
        except Exception: pass
        raise RuntimeError(f"OpenRouter HTTP {r.status_code}")

    try:
        data = r.json()
    except Exception as e:
        _llm_block(f"OpenRouter JSON parse error: {e}", LLM_COOLDOWN_SECONDS)
        raise RuntimeError("OpenRouter JSON parse error")

    choice = (data.get("choices") or [{}])[0]
    message = choice.get("message") or {}

    tool_calls = message.get("tool_calls") or []
    if tool_calls:
        try:
            args_str = tool_calls[0]["function"]["arguments"]
            return json.loads(args_str)
        except Exception as e:
            log.debug("Falha ao parsear tool_call.arguments: %s", e)

    content = (message.get("content") or "").strip()
    log.debug("LLM raw content (primeiros 400 chars): %s", content[:400])

    m = re.search(r"```(?:json)?\s*({.*})\s*```", content, flags=re.S)
    if m:
        try: return json.loads(m.group(1))
        except Exception: pass

    content = content.replace("```json", "```").strip("`").strip()

    def _first_json_object(s: str):
        start = s.find("{")
        while start != -1:
            depth, in_str, esc = 0, False, False
            for i in range(start, len(s)):
                ch = s[i]
                if in_str:
                    if esc: esc = False
                    elif ch == "\\": esc = True
                    elif ch == '"': in_str = False
                else:
                    if ch == '"': in_str = True
                    elif ch == "{": depth += 1
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
# SMTP helpers (fallback)
# ==========================
def _parse_smtp_hosts(env_value: str) -> list[str]:
    if not env_value:
        return []
    items = []
    raw = env_value.replace(';', ',').replace('\n', ' ')
    for token in raw.split(','):
        token = token.strip()
        if not token: continue
        for t in token.split():
            t = t.strip()
            if t: items.append(t)
    seen, uniq = set(), []
    for h in items:
        if h not in seen:
            uniq.append(h); seen.add(h)
    return uniq

def _effective_smtp_hosts() -> list[str]:
    hosts = _parse_smtp_hosts(SMTP_HOSTS)
    if not hosts and SMTP_HOST:
        hosts = [SMTP_HOST]
    return hosts

def _can_resolve_host(host: str) -> bool:
    try:
        socket.getaddrinfo(host, None)
        return True
    except socket.gaierror:
        return False

def _is_temporary_smtp_error(exc: Exception) -> bool:
    if isinstance(exc, (smtplib.SMTPServerDisconnected,
                        smtplib.SMTPConnectError,
                        smtplib.SMTPDataError,
                        smtplib.SMTPHeloError,
                        smtplib.SMTPAuthenticationError,
                        TimeoutError,
                        socket.timeout,
                        ConnectionRefusedError)):
        return True
    if isinstance(exc, socket.gaierror):
        return False
    return True

def _resolve_endpoints(host: str, port: int, prefer_ipv4: bool):
    families = [socket.AF_INET] if prefer_ipv4 else [socket.AF_UNSPEC]
    endpoints = []
    for fam in families:
        try:
            infos = socket.getaddrinfo(host, port, fam, socket.SOCK_STREAM)
            endpoints.extend(infos)
        except socket.gaierror:
            continue
    if prefer_ipv4:
        endpoints.sort(key=lambda x: 0 if x[0] == socket.AF_INET else 1)
    return endpoints

def _dial_smtp_endpoint(host: str, port: int, mode: str, timeout: int, ctx: ssl.SSLContext):
    endpoints = _resolve_endpoints(host, port, SMTP_PREFER_IPV4)
    if not endpoints:
        raise RuntimeError(f"getaddrinfo vazio para {host}:{port}")

    last_exc = None
    for family, socktype, proto, _, sockaddr in endpoints:
        sock = None
        try:
            sock = socket.socket(family, socktype, proto)
            sock.settimeout(timeout)
            t0 = time.time()
            sock.connect(sockaddr)
            t_ms = int((time.time() - t0) * 1000)

            if mode == "ssl":
                ssl_sock = ctx.wrap_socket(sock, server_hostname=host)
                s = smtplib.SMTP_SSL()
                s.timeout = timeout
                s.sock = ssl_sock
                s.file = s.sock.makefile("rb")
                s._host = host
                s.ehlo()
                log.info("[SMTP] TCP+SSL OK %s:%s em %dms", host, port, t_ms)
                return s
            else:
                s = smtplib.SMTP()
                s.timeout = timeout
                s.sock = sock
                s.file = s.sock.makefile("rb")
                s._host = host
                s.ehlo()
                s.starttls(context=ctx)
                s.ehlo()
                log.info("[SMTP] TCP+STARTTLS OK %s:%s em %dms", host, port, t_ms)
                return s

        except Exception as e:
            last_exc = e
            try:
                if sock: sock.close()
            except Exception:
                pass
            continue

    raise last_exc if last_exc else RuntimeError("Falha desconhecida no dial SMTP")

def _smtp_connect_with_fallback():
    if _smtp_is_blocked_now():
        remaining = int(_smtp_block_until_ts - time.time())
        raise RuntimeError(f"SMTP em cooldown ({remaining}s) — pulando envio")

    hosts = _effective_smtp_hosts()
    if not hosts:
        raise RuntimeError("Nenhum host SMTP definido. Configure SMTP_HOSTS ou SMTP_HOST.")

    attempts = []
    attempts.extend([(h, SMTP_PORT, SMTP_TLS_MODE) for h in hosts])
    for item in (SMTP_FALLBACKS or "").split(","):
        item = item.strip()
        if not item: continue
        try:
            port_str, mode = item.split(":", 1)
            port = int(port_str)
            mode = mode.strip().lower()
            for h in hosts:
                if (h, port, mode) not in attempts:
                    attempts.append((h, port, mode))
        except Exception:
            continue

    ctx = ssl.create_default_context()
    last_err = None
    had_temporary_error = False

    for host, port, mode in attempts:
        if not _can_resolve_host(host):
            log.warning("[SMTP] host sem DNS: %s — pulando (sem cooldown)", host)
            last_err = RuntimeError(f"DNS não resolve para {host}")
            continue
        try:
            log.info("[SMTP] tentando %s:%s (%s) ipv4_prefer=%s", host, port, mode, SMTP_PREFER_IPV4)
            s = _dial_smtp_endpoint(host, port, mode, SMTP_TIMEOUT, ctx)
            s.set_debuglevel(SMTP_DEBUG)
            s.login(MAIL_USER, MAIL_PASS)
            log.info("[SMTP] conectado e autenticado em %s:%s (%s)", host, port, mode)
            return s
        except Exception as e:
            last_err = e
            try:
                s.quit()
            except Exception:
                try: s.close()
                except Exception: pass
            if _is_temporary_smtp_error(e):
                had_temporary_error = True
                log.warning("[SMTP] falhou %s:%s (%s) — temporário: %s", host, port, mode, e)
            else:
                log.warning("[SMTP] falhou %s:%s (%s) — configuração/DNS (sem cooldown): %s", host, port, mode, e)
            continue

    if had_temporary_error:
        _smtp_block(f"Todas as tentativas SMTP falharam (erros temporários). Último erro: {last_err}", SMTP_COOLDOWN_SECONDS)
        raise RuntimeError(f"Todas as tentativas SMTP falharam (temporário). Cooldown {SMTP_COOLDOWN_SECONDS}s. Último erro: {last_err}")
    else:
        raise RuntimeError(f"Falha de configuração/DNS em SMTP: {last_err} (sem cooldown)")

# ==========================
# Construção do e-mail (com headers de thread)
# ==========================
def ensure_reply_prefix(subject: str) -> str:
    if not subject: return "Re:"
    s = subject.strip()
    return s if s.lower().startswith("re:") else f"Re: {s}"

def build_user_prompt(original_subject: str, body_text: str, cobol_files: list) -> str:
    parts = [
        f"Assunto original: {original_subject}",
        "\nCorpo do e-mail do aluno:\n" + (body_text or "(vazio)")
    ]
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
    out = "\n".join(lines).replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"\n{3,}", "\n\n", out)

def _thread_headers_for(original_msg):
    headers = {}
    if original_msg:
        try:
            orig_msgid = original_msg.get("Message-ID")
        except Exception:
            orig_msgid = None
        if orig_msgid:
            headers["In-Reply-To"] = orig_msgid
            refs = original_msg.get_all("References", [])
            ref_line = " ".join(refs + [orig_msgid]) if refs else orig_msgid
            headers["References"] = ref_line
    return headers

# ==========================
# Envio por Mailgun API (primário)
# ==========================
def _send_via_mailgun_api(to_addr: str, subject: str, body_text: str, extra_headers: dict) -> dict:
    if not (MAILGUN_API_KEY and MAILGUN_DOMAIN):
        raise RuntimeError("Mailgun API não configurada (defina MAILGUN_API_KEY e MAILGUN_DOMAIN).")

    url = f"{MAILGUN_API_BASE}/{MAILGUN_DOMAIN}/messages"
    auth = ("api", MAILGUN_API_KEY)
    data = {
        "from": MAIL_USER or f"postmaster@{MAILGUN_DOMAIN}",
        "to": [to_addr],
        "subject": subject,
        "text": body_text,
    }
    # Headers de thread (Mailgun aceita prefixo h:)
    for k, v in (extra_headers or {}).items():
        if v:
            data[f"h:{k}"] = v

    r = requests.post(url, auth=auth, data=data, timeout=20)
    if r.status_code // 100 != 2:
        raise RuntimeError(f"Mailgun API HTTP {r.status_code}: {r.text[:300]}")
    try:
        return r.json()
    except Exception:
        return {"ok": True, "raw": r.text[:300]}

# ==========================
# Envio (API primária + SMTP fallback)
# ==========================
def send_email_reply(original_msg, to_addr: str, subject: str, body_text: str) -> bytes:
    """
    Tenta 1) Mailgun API (se configurada), senão 2) SMTP com fallback.
    Retorna o MIME em bytes (para copiar para Sent).
    """
    # Monta MIME (independente do transporte)
    msg = EmailMessage()
    msg["From"] = MAIL_USER or (f"postmaster@{MAILGUN_DOMAIN}" if MAILGUN_DOMAIN else MAIL_USER)
    msg["Sender"] = msg["From"]
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=(MAIL_USER.split("@", 1)[-1] if MAIL_USER and "@" in MAIL_USER else None))
    msg["X-Mailer"] = "COBOL Support Agent"

    th = _thread_headers_for(original_msg)
    for k, v in th.items():
        msg[k] = v

    msg.set_content(body_text)

    transport = _mail_primary_transport()
    last_err = None

    # 1) Primário: Mailgun API
    if transport == "mailgun_api":
        try:
            _ = _send_via_mailgun_api(to_addr, subject, body_text, th)
            log.info("E-mail enviado via Mailgun API para %s (Subject: %s)", to_addr, subject)
            return msg.as_bytes()
        except Exception as e:
            last_err = e
            log.warning("Falha no envio via Mailgun API — tentando SMTP fallback: %s", e)

    # 2) Fallback: SMTP
    s = _smtp_connect_with_fallback()
    try:
        refused = s.sendmail(msg["From"], [to_addr], msg.as_string())
        if refused:
            raise RuntimeError(f"SMTP refused {refused}")
        log.info("E-mail enviado via SMTP para %s (Subject: %s)", to_addr, subject)
        return msg.as_bytes()
    finally:
        try: s.quit()
        except Exception:
            try: s.close()
            except Exception: pass

# ==========================
# IMAP helpers
# ==========================
def ensure_mailbox(imap: imaplib.IMAP4_SSL, box: str):
    try: imap.create(box)
    except Exception: log.debug("Mailbox pode já existir: %s", box)

def move_message_uid(imap: imaplib.IMAP4_SSL, msg_uid: bytes, dest_box: str):
    ensure_mailbox(imap, dest_box)
    imap.uid("COPY", msg_uid, dest_box)
    imap.uid("STORE", msg_uid, "+FLAGS", r"(\Deleted)")

def append_to_sent(raw_bytes: bytes):
    u, p = _imap_creds()
    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
        imap.login(u, p)
        ensure_mailbox(imap, SENT_FOLDER)
        imap.append(SENT_FOLDER, r"(\Seen)", None, raw_bytes)
        imap.logout()
        log.info("Mensagem copiada para a pasta de enviados: %s", SENT_FOLDER)

def _extract_raw_bytes_from_fetch(msg_data) -> bytes | None:
    if not msg_data:
        return None
    for part in msg_data:
        if isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], (bytes, bytearray)):
            return part[1]
    return None

# ==========================
# Fluxo principal
# ==========================
def decide_and_respond(imap: imaplib.IMAP4_SSL, msg_uid: bytes, msg_bytes: bytes):
    msg = BytesParser(policy=policy.default).parsebytes(msg_bytes)
    sender = decode_mime_words(msg.get("From"))
    from_addr = re.search(r"<([^>]+)>", sender or "")
    to_reply = from_addr.group(1) if from_addr else (sender or "")

    original_subject = decode_mime_words(msg.get("Subject"))
    body_text = extract_text_body(msg)
    cobol_files = extract_cobol_attachments(msg)

    # PRE-TRIAGEM
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

    # Se LLM bloqueado, escalamos sem chamar LLM
    if _llm_is_blocked_now():
        remaining = max(0, int(_llm_block_until_ts - time.time())) if not LLM_HARD_DISABLE else -1
        log.warning("LLM em cooldown/hard-disable (%ss). Pulando LLM e escalando. De: %s Assunto: %s",
                    remaining, sender, original_subject)
        move_message_uid(imap, msg_uid, _safe_box(FOLDER_ESCALATE))
        return

    user_prompt = build_user_prompt(original_subject or "", body_text, cobol_files)

    try:
        llm_json = call_openrouter(SYSTEM_PROMPT, user_prompt)
    except Exception:
        log.error("LLM error", exc_info=True)
        move_message_uid(imap, msg_uid, _safe_box(FOLDER_ESCALATE))
        log.info("E-mail movido para %s", _safe_box(FOLDER_ESCALATE))
        return

    acao = llm_json.get("acao", "escalar")
    try: nivel = float(llm_json.get("nivel_confianca", 0) or 0)
    except Exception: nivel = 0.0
    assunto_model = llm_json.get("assunto", original_subject or "")
    corpo_markdown = llm_json.get("corpo_markdown", "")

    should_answer = (acao == "responder") and (nivel >= CONFIDENCE_THRESHOLD)

    if not should_answer:
        move_message_uid(imap, msg_uid, _safe_box(FOLDER_ESCALATE))
        log.info("Decisão do modelo: acao=%s conf=%.2f → escalado", acao, nivel)
        return

    subject_to_send = ensure_reply_prefix(original_subject or assunto_model or "")
    body_out = build_outgoing_body(corpo_markdown)

    try:
        raw_out = send_email_reply(msg, to_reply, subject_to_send, body_out)
        log.info("E-mail enviado para %s (Subject: %s)", to_reply, subject_to_send)
        try:
            ensure_mailbox(imap, SENT_FOLDER)
            append_to_sent(raw_out)
        except Exception:
            log.error("Falha ao copiar para enviados", exc_info=True)
        move_message_uid(imap, msg_uid, _safe_box(FOLDER_PROCESSED))
        log.info("E-mail movido para %s", _safe_box(FOLDER_PROCESSED))
    except Exception:
        log.error("Falha no envio — movendo para %s (evita reprocessamento)",
                  _safe_box(FOLDER_ESCALATE), exc_info=True)
        move_message_uid(imap, msg_uid, _safe_box(FOLDER_ESCALATE))

# ==========================
# Watcher IMAP (por UID)
# ==========================
def _imap_uid_search(imap, criteria: str):
    typ, data = imap.uid("search", None, criteria)
    return data[0].split() if typ == "OK" and data and data[0] else []

def _select_box(imap, box: str):
    typ, data = imap.select(box)
    if typ == "OK" and data:
        try: exists = int(data[0])
        except Exception:
            exists = data[0].decode() if isinstance(data[0], bytes) else data[0]
        log.debug("SELECT %s → EXISTS=%s", box, exists)
    else:
        log.warning("Falha ao selecionar caixa %s: %s %s", box, typ, data)

def _imap_since_date(days: int) -> str:
    dt = datetime.now(timezone.utc) - timedelta(days=days)
    mon = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][dt.month - 1]
    return f"{dt.day:02d}-{mon}-{dt.year}"

def _imap_search_unseen(imap, since_days: int) -> list[bytes]:
    if since_days and since_days > 0:
        date_str = _imap_since_date(since_days)
        return _imap_uid_search(imap, f"UNSEEN SINCE {date_str}")
    return _imap_uid_search(imap, "UNSEEN")

def watch_imap_loop():
    hosts = _effective_smtp_hosts()
    log.info("Watcher IMAP — envio primário=%s | SMTP hosts=%s", _mail_primary_transport(), hosts or "[nenhum host configurado]")
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

                u, p = _imap_creds()
                imap.login(u, p)
                _select_box(imap, IMAP_FOLDER_INBOX)

                uids = _imap_search_unseen(imap, IMAP_SINCE_DAYS)
                llm_blocked = _llm_is_blocked_now()

                if not uids and not IMAP_STRICT_UNSEEN_ONLY:
                    if not uids and not llm_blocked:
                        recent = _imap_uid_search(imap, "NEW")
                        if recent: log.debug("NEW (UIDs): %s", recent); uids = recent
                    if not uids and not llm_blocked:
                        recent = _imap_uid_search(imap, "RECENT")
                        if recent: log.debug("RECENT (UIDs): %s", recent); uids = recent
                    if (not uids) and IMAP_FALLBACK_LAST_N > 0 and (not llm_blocked or IMAP_FALLBACK_WHEN_LLM_BLOCKED):
                        all_uids = _imap_uid_search(imap, "ALL")
                        tail = all_uids[-IMAP_FALLBACK_LAST_N:] if all_uids else []
                        if tail:
                            log.warning("UNSEEN/NEW/RECENT vazios — usando últimos %d UIDs: %s", IMAP_FALLBACK_LAST_N, tail)
                            uids = tail

                log.debug("UIDs a processar: %s", uids)

                for uid in uids:
                    if not uid or not uid.strip():
                        continue
                    try:
                        typ, msg_data = imap.uid("fetch", uid, "(BODY.PEEK[])")
                        if typ != "OK" or not msg_data:
                            log.warning("Falha no FETCH UID=%s: %s %s", uid, typ, msg_data)
                            continue
                        raw_bytes = _extract_raw_bytes_from_fetch(msg_data)
                        if not raw_bytes:
                            log.warning("FETCH sem corpo legível UID=%s — pulando", uid)
                            continue
                        try:
                            decide_and_respond(imap, uid, raw_bytes)
                        except Exception:
                            log.exception("Exceção em decide_and_respond UID=%s — movendo para Escalar", uid)
                            try:
                                move_message_uid(imap, uid, _safe_box(FOLDER_ESCALATE))
                            except Exception:
                                log.exception("Falha ao mover UID=%s para Escalar", uid)
                    except Exception:
                        log.exception("Erro ao processar UID=%s", uid)

                if EXPUNGE_AFTER_COPY:
                    try: imap.expunge()
                    except Exception: log.exception("Falha no expunge final")

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
    return jsonify({"sha1": SYSTEM_PROMPT_SHA1, "first120": SYSTEM_PROMPT[:120]})

@app.route("/diag/llm")
def diag_llm_status():
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

@app.route("/diag/llm/unblock", methods=["POST"])
def diag_llm_unblock():
    global _llm_block_until_ts, _last_llm_error
    _llm_block_until_ts = 0.0
    _last_llm_error = ""
    return jsonify({"ok": True, "unblocked": True})

@app.route("/diag/transport/status")
def diag_transport_status():
    remaining = max(0, int(_smtp_block_until_ts - time.time()))
    return jsonify({
        "primary_transport": _mail_primary_transport(),
        "mailgun": {
            "configured": bool(MAILGUN_API_KEY and MAILGUN_DOMAIN),
            "domain": MAILGUN_DOMAIN or None,
            "api_base": MAILGUN_API_BASE,
        },
        "smtp": {
            "blocked": _smtp_is_blocked_now(),
            "block_expires_in_seconds": remaining,
            "last_error": _last_smtp_error,
            "hosts": _effective_smtp_hosts(),
            "primary": {"host": SMTP_HOST or None, "port": SMTP_PORT, "mode": SMTP_TLS_MODE},
            "fallbacks": SMTP_FALLBACKS,
            "timeout": SMTP_TIMEOUT,
            "prefer_ipv4": SMTP_PREFER_IPV4,
        }
    })

@app.route("/diag/smtp/status")
def diag_smtp_status():
    remaining = max(0, int(_smtp_block_until_ts - time.time()))
    return jsonify({
        "blocked": _smtp_is_blocked_now(),
        "block_expires_in_seconds": remaining,
        "last_error": _last_smtp_error,
        "hosts": _effective_smtp_hosts(),
        "primary": {"host": SMTP_HOST or None, "port": SMTP_PORT, "mode": SMTP_TLS_MODE},
        "fallbacks": SMTP_FALLBACKS,
        "timeout": SMTP_TIMEOUT,
        "prefer_ipv4": SMTP_PREFER_IPV4,
    })

@app.route("/diag/smtp/unblock", methods=["POST"])
def diag_smtp_unblock():
    global _smtp_block_until_ts, _last_smtp_error
    _smtp_block_until_ts = 0.0
    _last_smtp_error = ""
    return jsonify({"ok": True, "unblocked": True})

@app.route("/diag/smtp/probe")
def diag_smtp_probe():
    """Prova todas as tentativas (sem enviar e-mail) e SEM cooldown."""
    hosts = _effective_smtp_hosts()
    attempts = []
    attempts.extend([(h, SMTP_PORT, SMTP_TLS_MODE) for h in hosts])
    for item in (SMTP_FALLBACKS or "").split(","):
        item = item.strip()
        if not item: continue
        try:
            port_str, mode = item.split(":", 1)
            port = int(port_str)
            mode = mode.strip().lower()
            for h in hosts:
                if (h, port, mode) not in attempts:
                    attempts.append((h, port, mode))
        except Exception:
            continue

    ctx = ssl.create_default_context()
    report = []
    for host, port, mode in attempts:
        entry = {"host": host, "port": port, "mode": mode, "ok": False}
        try:
            s = _dial_smtp_endpoint(host, port, mode, SMTP_TIMEOUT, ctx)
            try: s.close()
            except Exception: pass
            entry["ok"] = True
        except Exception as e:
            entry["error"] = str(e)
        report.append(entry)

    return jsonify({
        "prefer_ipv4": SMTP_PREFER_IPV4,
        "hosts": hosts,
        "attempts": report,
    })

@app.route("/diag/email")
def diag_email():
    """
    Envio real usando transporte primário (Mailgun API se configurada; se falhar, SMTP).
    Parâmetros: /diag/email?to=alguem@destino.com&subject=Teste&body=Texto
    """
    to = request.args.get("to", MAIL_USER)
    subject = request.args.get("subject", "Teste de envio — COBOL Support Agent")
    body = request.args.get("body", "Olá! Teste de envio com transporte primário + fallback.\n\n— Sistema")
    try:
        raw = send_email_reply(None, to, subject, body)
        # opcional: copia para Sent (IMAP)
        try:
            append_to_sent(raw)
        except Exception:
            log.exception("Falha ao copiar para enviados (opcional)")
        return jsonify({"ok": True, "to": to, "primary_transport": _mail_primary_transport()})
    except Exception as e:
        log.exception("Falha no /diag/email")
        return jsonify({"ok": False, "error": str(e), "primary_transport": _mail_primary_transport()}), 500

@app.route("/diag/imap")
def diag_imap():
    box = request.args.get("box", IMAP_FOLDER_INBOX)
    n = int(request.args.get("n", "20"))
    try:
        u, p = _imap_creds()
        with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as imap:
            imap.login(u, p)
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
        })
    except Exception as e:
        log.exception("diag/imap falhou")
        return jsonify({"error": str(e)}), 500

# (NOVO) Teste explícito de autenticação IMAP
@app.route("/diag/imap/auth")
def diag_imap_auth():
    """
    Testa somente o login IMAP e retorna erro literal do servidor.
    Parâmetros opcionais: host, port
    Ex.: /diag/imap/auth?host=mail.aprendacobol.com.br&port=993
    """
    host = request.args.get("host", IMAP_HOST)
    port = int(request.args.get("port", IMAP_PORT))
    u, p = _imap_creds()
    try:
        with imaplib.IMAP4_SSL(host, port) as imap:
            try:
                imap.login(u, p)
                imap.logout()
                return jsonify({"ok": True, "host": host, "port": port, "user": u})
            except imaplib.IMAP4.error as e:
                return jsonify({"ok": False, "host": host, "port": port, "user": u, "error": str(e)}), 401
    except Exception as e:
        return jsonify({"ok": False, "host": host, "port": port, "user": u, "error": str(e)}), 500

# ==========================
# Main
# ==========================
if __name__ == "__main__":
    import threading
    t = threading.Thread(target=watch_imap_loop, daemon=True)
    t.start()
    port = int(os.getenv("PORT", "10000"))
    log.info("Iniciando Flask em 0.0.0.0:%s", port)
    app.run(host="0.0.0.0", port=port)
