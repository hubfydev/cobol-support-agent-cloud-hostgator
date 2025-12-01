#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ****** COBOL Support Agent — v10.19 ********
# ****** Andre Richest                ********
# ****** Sun Nov 30 2025              ********

import os
import ssl
import time
import json
import socket
import logging
import threading
from typing import Optional, Tuple, List

import imaplib
import smtplib
from email.message import EmailMessage

import requests  # <-- Mailgun API

from flask import Flask, jsonify, request

# -------------------------------------------------------------
# Logging
# -------------------------------------------------------------
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

# -------------------------------------------------------------
# Env / Config
# -------------------------------------------------------------
APP_PUBLIC_URL = os.getenv("APP_PUBLIC_URL", "http://localhost:10000")
PORT = int(os.getenv("PORT", "10000"))

IMAP_HOST = os.getenv("IMAP_HOST", "")
IMAP_PORT = int(os.getenv("IMAP_PORT", "993"))
IMAP_USER = os.getenv("IMAP_USER", "")
IMAP_PASS = os.getenv("IMAP_PASS", "")
IMAP_TLS_MODE = os.getenv("IMAP_TLS_MODE", "ssl").lower()  # ssl | starttls | plain
IMAP_FOLDER_INBOX = os.getenv("IMAP_FOLDER_INBOX", "INBOX")
IMAP_STRICT_UNSEEN_ONLY = os.getenv("IMAP_STRICT_UNSEEN_ONLY", "True").lower() == "true"
IMAP_SINCE_DAYS = int(os.getenv("IMAP_SINCE_DAYS", "0"))
IMAP_FALLBACK_LAST_N = int(os.getenv("IMAP_FALLBACK_LAST_N", "0"))
IMAP_FALLBACK_WHEN_LLM_BLOCKED = os.getenv("IMAP_FALLBACK_WHEN_LLM_BLOCKED", "False").lower() == "true"

FOLDER_PROCESSED = os.getenv("FOLDER_PROCESSED", "Respondidos")
FOLDER_ESCALATE = os.getenv("FOLDER_ESCALATE", "Escalar")
EXPUNGE_AFTER_COPY = os.getenv("EXPUNGE_AFTER_COPY", "true").lower() == "true"

CHECK_INTERVAL_SECONDS = int(os.getenv("CHECK_INTERVAL_SECONDS", "60"))

# --- SMTP (mantido, mas hoje bloqueado pela Render) ---
SMTP_HOST = os.getenv("SMTP_HOST", "")
SMTP_HOSTS = [h.strip() for h in os.getenv("SMTP_HOSTS", SMTP_HOST).split(",") if h.strip()]
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))  # Mailgun recomenda 587
SMTP_TLS_MODE = os.getenv("SMTP_TLS_MODE", "starttls").lower()  # ssl | starttls
SMTP_USER = os.getenv("MAIL_USER", os.getenv("SMTP_USER", ""))
SMTP_PASS = os.getenv("SMTP_PASS", os.getenv("MAIL_PASS", ""))
SMTP_CONNECT_TIMEOUT = int(os.getenv("SMTP_CONNECT_TIMEOUT", "10"))
SMTP_TIMEOUT = int(os.getenv("SMTP_TIMEOUT", "20"))
SMTP_PREFER_IPV4 = os.getenv("SMTP_PREFER_IPV4", "true").lower() == "true"
SMTP_FALLBACKS = os.getenv("SMTP_FALLBACKS", "587:starttls,465:ssl,2525:starttls")
SMTP_COOLDOWN_SECONDS = int(os.getenv("SMTP_COOLDOWN_SECONDS", "900"))

SIGNATURE_NAME = os.getenv("SIGNATURE_NAME", "Equipe Aprenda COBOL — Suporte")
SIGNATURE_FOOTER = os.getenv("SIGNATURE_FOOTER", "")
SIGNATURE_LINKS = os.getenv("SIGNATURE_LINKS", "")

SMTP_FROM_EMAIL = os.getenv("SMTP_FROM_EMAIL", SMTP_USER or "")
SMTP_FROM_NAME = os.getenv("SMTP_FROM_NAME", SIGNATURE_NAME)
SMTP_REPLY_TO = os.getenv("SMTP_REPLY_TO", SMTP_FROM_EMAIL)

APP_TITLE = os.getenv("APP_TITLE", "COBOL Support Agent")

# --- Mailgun API ---
MAILGUN_API_KEY = os.getenv("MAILGUN_API_KEY", "")
MAILGUN_DOMAIN = os.getenv("MAILGUN_DOMAIN", "")
MAILGUN_API_BASE = os.getenv("MAILGUN_API_BASE", "https://api.mailgun.net/v3")

# -------------------------------------------------------------
# Helpers
# -------------------------------------------------------------

def _resolve_host(host: str) -> List[str]:
    """Resolve hostnames to IPs; opcionalmente prefere IPv4 (só para log)."""
    try:
        family = socket.AF_INET if SMTP_PREFER_IPV4 else socket.AF_UNSPEC
        infos = socket.getaddrinfo(host, None, family, socket.SOCK_STREAM)
        addrs = []
        for info in infos:
            addr = info[4][0]
            if addr not in addrs:
                addrs.append(addr)
        return addrs or [host]
    except Exception:
        return [host]


def _ssl_context() -> ssl.SSLContext:
    ctx = ssl.create_default_context()
    ctx.check_hostname = True
    ctx.verify_mode = ssl.CERT_REQUIRED
    return ctx


# -------------------------------------------------------------
# IMAP connect
# -------------------------------------------------------------

class ImapAuthError(Exception):
    pass


def imap_connect(host: str, port: int, user: str, password: str, mode: str):
    mode = (mode or "ssl").lower()
    log.info(f"IMAP endpoint: {host}:{port} (mode={mode})")
    if mode == "ssl":
        imap = imaplib.IMAP4_SSL(host, port, ssl_context=_ssl_context())
        try:
            imap.login(user, password)
        except imaplib.IMAP4.error as e:
            raise ImapAuthError(str(e))
        return imap
    elif mode == "starttls":
        imap = imaplib.IMAP4(host, port)
        imap.starttls(ssl_context=_ssl_context())
        try:
            imap.login(user, password)
        except imaplib.IMAP4.error as e:
            raise ImapAuthError(str(e))
        return imap
    elif mode == "plain":
        imap = imaplib.IMAP4(host, port)
        try:
            imap.login(user, password)
        except imaplib.IMAP4.error as e:
            raise ImapAuthError(str(e))
        return imap
    else:
        raise ValueError("IMAP mode must be one of: ssl, starttls, plain")


# -------------------------------------------------------------
# SMTP connect (+ fallback list and mode switching)
# -------------------------------------------------------------

class SmtpTempError(Exception):
    pass


_last_smtp_fail_ts: Optional[float] = None


def smtp_connect_once(host: str, port: int, mode: str) -> smtplib.SMTP:
    """
    Conecta em um único host/porta/mode usando o hostname para TLS/SNI.
    (Hoje deve falhar por bloqueio de porta na Render.)
    """
    mode = (mode or "ssl").lower()
    addrs = _resolve_host(host)
    log.info(f"SMTP tentativa — {host} -> {addrs}, port={port}, mode={mode}")

    try:
        if mode == "ssl":
            s = smtplib.SMTP_SSL(host, port, timeout=SMTP_CONNECT_TIMEOUT, context=_ssl_context())
        else:
            s = smtplib.SMTP(host, port, timeout=SMTP_CONNECT_TIMEOUT)
            if mode == "starttls":
                s.starttls(context=_ssl_context())
        s.login(SMTP_USER, SMTP_PASS)
        s.timeout = SMTP_TIMEOUT
        log.info(f"SMTP conectado via {host}:{port} ({mode})")
        return s
    except (smtplib.SMTPConnectError, smtplib.SMTPServerDisconnected, socket.timeout) as e:
        log.warning(f"SMTP connect falhou em {host}:{port} ({mode}) — {e}")
    except smtplib.SMTPAuthenticationError as e:
        log.error(f"SMTP AUTH falhou em {host}:{port} ({mode}) — {e}")
        raise
    except Exception as e:
        log.warning(f"SMTP erro em {host}:{port} ({mode}) — {e}")

    raise SmtpTempError("Todas as tentativas SMTP falharam (temporárias)")


def smtp_connect_with_fallback() -> smtplib.SMTP:
    global _last_smtp_fail_ts

    if _last_smtp_fail_ts is not None:
        remaining = int(SMTP_COOLDOWN_SECONDS - (time.time() - _last_smtp_fail_ts))
        if remaining > 0:
            raise RuntimeError(f"SMTP em cooldown ({remaining}s) — pulando envio")
        else:
            _last_smtp_fail_ts = None

    try:
        if SMTP_HOSTS:
            for host in SMTP_HOSTS:
                try:
                    return smtp_connect_once(host, SMTP_PORT, SMTP_TLS_MODE)
                except SmtpTempError:
                    continue
        else:
            return smtp_connect_once(SMTP_HOST, SMTP_PORT, SMTP_TLS_MODE)
    except smtplib.SMTPAuthenticationError:
        raise
    except Exception as e:
        log.warning(f"SMTP primário indisponível: {e}")

    for item in [x.strip() for x in SMTP_FALLBACKS.split(',') if x.strip()]:
        try:
            p, m = item.split(':', 1)
            p = int(p)
            m = m.strip().lower()
        except Exception:
            continue
        try:
            if SMTP_HOSTS:
                for host in SMTP_HOSTS:
                    try:
                        return smtp_connect_once(host, p, m)
                    except SmtpTempError:
                        continue
            else:
                return smtp_connect_once(SMTP_HOST, p, m)
        except smtplib.SMTPAuthenticationError:
            raise
        except Exception as e:
            log.warning(f"SMTP fallback {p}/{m} falhou: {e}")

    _last_smtp_fail_ts = time.time()
    raise RuntimeError("SMTP temporariamente indisponível: timed out")


# -------------------------------------------------------------
# Mailgun API send
# -------------------------------------------------------------

def _build_from_header() -> str:
    email = SMTP_FROM_EMAIL or SMTP_USER
    name = (SMTP_FROM_NAME or "").strip()
    if email and name:
        return f"{name} <{email}>"
    return email or ""


def send_via_mailgun_api(to_addr: str, subject: str, body: str) -> str:
    if not MAILGUN_API_KEY or not MAILGUN_DOMAIN:
        raise RuntimeError("Mailgun API não configurada (MAILGUN_API_KEY/MAILGUN_DOMAIN)")

    from_header = _build_from_header()
    text_body = body + f"\n\n{SIGNATURE_NAME}\n{SIGNATURE_FOOTER}\n{SIGNATURE_LINKS}"

    url = f"{MAILGUN_API_BASE.rstrip('/')}/{MAILGUN_DOMAIN}/messages"
    data = {
        "from": from_header,
        "to": [to_addr],
        "subject": subject,
        "text": text_body,
    }
    if SMTP_REPLY_TO:
        data["h:Reply-To"] = SMTP_REPLY_TO

    log.info(f"Mailgun API POST {url} -> to={to_addr}")
    resp = requests.post(
        url,
        auth=("api", MAILGUN_API_KEY),
        data=data,
        timeout=SMTP_TIMEOUT,
    )
    resp.raise_for_status()
    log.info(f"Mailgun API resposta {resp.status_code}: {resp.text[:200]}")
    return "ok"


# -------------------------------------------------------------
# Minimal mail actions (stub for reply flow)
# -------------------------------------------------------------

def send_test_email(to_addr: str, subject: str, body: str) -> str:
    """
    Envia e-mail. Prioridade:
    1) Mailgun API (porta 443, deve funcionar na Render)
    2) SMTP (mantido como fallback, mas hoje bloqueado)
    """
    if MAILGUN_API_KEY and MAILGUN_DOMAIN:
        return send_via_mailgun_api(to_addr, subject, body)

    # fallback SMTP (provavelmente não vai funcionar na Render, mas fica para compatibilidade)
    s = smtp_connect_with_fallback()
    try:
        msg = EmailMessage()
        from_header = _build_from_header()
        if from_header:
            msg["From"] = from_header
        if SMTP_REPLY_TO:
            msg["Reply-To"] = SMTP_REPLY_TO

        msg["To"] = to_addr
        msg["Subject"] = subject
        msg.set_content(
            body + f"\n\n{SIGNATURE_NAME}\n{SIGNATURE_FOOTER}\n{SIGNATURE_LINKS}"
        )
        s.send_message(msg)
        return "ok"
    finally:
        try:
            s.quit()
        except Exception:
            pass


# -------------------------------------------------------------
# Watcher (checks auth + SELECT INBOX)
# -------------------------------------------------------------

def watch_imap_loop():
    log.info(
        f"IMAP_STRICT_UNSEEN_ONLY={IMAP_STRICT_UNSEEN_ONLY} | "
        f"IMAP_SINCE_DAYS={IMAP_SINCE_DAYS} | "
        f"IMAP_FALLBACK_LAST_N={IMAP_FALLBACK_LAST_N} | "
        f"IMAP_FALLBACK_WHEN_LLM_BLOCKED={IMAP_FALLBACK_WHEN_LLM_BLOCKED}"
    )
    while True:
        try:
            log.info(
                f"IMAP tentando login como {IMAP_USER[:2]}***@ "
                f"em {IMAP_HOST}:{IMAP_PORT} (mode={IMAP_TLS_MODE})"
            )
            imap = imap_connect(IMAP_HOST, IMAP_PORT, IMAP_USER, IMAP_PASS, IMAP_TLS_MODE)
            try:
                typ, _ = imap.select(IMAP_FOLDER_INBOX)
                if typ != 'OK':
                    raise RuntimeError(f"SELECT {IMAP_FOLDER_INBOX} falhou: {typ}")
                log.info(f"IMAP autenticado e INBOX aberto — aguardando {CHECK_INTERVAL_SECONDS}s")
            finally:
                try:
                    imap.logout()
                except Exception:
                    pass
        except ImapAuthError as e:
            log.error(f"IMAP AUTH falhou: {e}")
        except Exception as e:
            log.error("Loop IMAP falhou", exc_info=e)
        time.sleep(CHECK_INTERVAL_SECONDS)


# -------------------------------------------------------------
# Flask app & diagnostics
# -------------------------------------------------------------
app = Flask(__name__)


@app.get("/")
def root():
    return jsonify({
        "title": APP_TITLE,
        "public_url": APP_PUBLIC_URL,
        "imap": {
            "host": IMAP_HOST,
            "port": IMAP_PORT,
            "mode": IMAP_TLS_MODE,
            "user": IMAP_USER,
        },
        "smtp": {
            "hosts": SMTP_HOSTS or [SMTP_HOST],
            "port": SMTP_PORT,
            "mode": SMTP_TLS_MODE,
            "user": SMTP_USER,
            "fallbacks": SMTP_FALLBACKS,
            "from_email": SMTP_FROM_EMAIL,
            "from_name": SMTP_FROM_NAME,
        },
        "mailgun": {
            "domain": MAILGUN_DOMAIN,
            "api_base": MAILGUN_API_BASE,
            "api_configured": bool(MAILGUN_API_KEY and MAILGUN_DOMAIN),
        }
    })


@app.get("/diag/imap/auth")
def diag_imap_auth():
    host = request.args.get("host", IMAP_HOST)
    port = int(request.args.get("port", IMAP_PORT))
    mode = request.args.get("mode", IMAP_TLS_MODE)
    user = request.args.get("user", IMAP_USER)
    password = request.args.get("pass", IMAP_PASS)

    try:
        imap = imap_connect(host, port, user, password, mode)
        try:
            typ, _ = imap.select(IMAP_FOLDER_INBOX)
            if typ != 'OK':
                raise RuntimeError(f"SELECT {IMAP_FOLDER_INBOX} falhou: {typ}")
        finally:
            try:
                imap.logout()
            except Exception:
                pass
        return jsonify({"ok": True, "host": host, "port": port, "mode": mode, "user": user})
    except ImapAuthError as e:
        return jsonify({
            "ok": False,
            "host": host,
            "port": port,
            "mode": mode,
            "user": user,
            "error": str(e)
        }), 401
    except Exception as e:
        return jsonify({
            "ok": False,
            "host": host,
            "port": port,
            "mode": mode,
            "user": user,
            "error": str(e)
        }), 500


@app.get("/diag/smtp/auth")
def diag_smtp_auth():
    # Mantido para debug, mas provavelmente vai continuar dando timeout na Render.
    host = request.args.get("host")
    port = request.args.get("port")
    mode = (request.args.get("mode") or SMTP_TLS_MODE).lower()
    user = request.args.get("user") or SMTP_USER
    password = request.args.get("pass") or SMTP_PASS

    if host and port:
        try:
            s = smtp_connect_once(host, int(port), mode)
            try:
                code = s.noop()[0]
            finally:
                try:
                    s.quit()
                except Exception:
                    pass
            return jsonify({
                "ok": True,
                "host": host,
                "port": int(port),
                "mode": mode,
                "user": user,
                "code": int(code),
            })
        except smtplib.SMTPAuthenticationError as e:
            return jsonify({
                "ok": False,
                "host": host,
                "port": int(port),
                "mode": mode,
                "user": user,
                "error": f"SMTP AUTH failed: {e}",
            }), 401
        except Exception as e:
            return jsonify({
                "ok": False,
                "host": host,
                "port": int(port),
                "mode": mode,
                "user": user,
                "error": str(e),
            }), 500

    try:
        s = smtp_connect_with_fallback()
        try:
            code = s.noop()[0]
        finally:
            try:
                s.quit()
            except Exception:
                pass
        return jsonify({
            "ok": True,
            "hosts": SMTP_HOSTS or [SMTP_HOST],
            "port": SMTP_PORT,
            "mode": SMTP_TLS_MODE,
            "fallbacks": SMTP_FALLBACKS,
            "user": user,
            "code": int(code),
        })
    except smtplib.SMTPAuthenticationError as e:
        return jsonify({
            "ok": False,
            "hosts": SMTP_HOSTS or [SMTP_HOST],
            "port": SMTP_PORT,
            "mode": SMTP_TLS_MODE,
            "fallbacks": SMTP_FALLBACKS,
            "user": user,
            "error": f"SMTP AUTH failed: {e}",
        }), 401
    except Exception as e:
        return jsonify({
            "ok": False,
            "hosts": SMTP_HOSTS or [SMTP_HOST],
            "port": SMTP_PORT,
            "mode": SMTP_TLS_MODE,
            "fallbacks": SMTP_FALLBACKS,
            "user": user,
            "error": str(e),
        }), 500


@app.get("/diag/smtp/ehlo")
def diag_smtp_ehlo():
    host = request.args.get("host") or (SMTP_HOSTS[0] if SMTP_HOSTS else SMTP_HOST)
    port = int(request.args.get("port") or SMTP_PORT)
    mode = (request.args.get("mode") or SMTP_TLS_MODE).lower()

    try:
        if mode == "ssl":
            s = smtplib.SMTP_SSL(host, port, timeout=SMTP_CONNECT_TIMEOUT, context=_ssl_context())
        else:
            s = smtplib.SMTP(host, port, timeout=SMTP_CONNECT_TIMEOUT)
            if mode == "starttls":
                s.starttls(context=_ssl_context())

        code, msg = s.ehlo()
        try:
            s.quit()
        except Exception:
            pass

        if isinstance(msg, bytes):
            msg_text = msg.decode(errors="ignore")
        else:
            msg_text = str(msg)

        features = [line.strip() for line in msg_text.splitlines() if line.strip()]

        return jsonify({
            "ok": True,
            "host": host,
            "port": port,
            "mode": mode,
            "code": int(code),
            "features": features,
        })
    except Exception as e:
        return jsonify({
            "ok": False,
            "host": host,
            "port": port,
            "mode": mode,
            "error": str(e),
        }), 500


@app.get("/diag/mailgun/api")
def diag_mailgun_api():
    """
    Envia um e-mail de teste via Mailgun API para validar conectividade HTTP.
    """
    to_addr = request.args.get("to") or (SMTP_FROM_EMAIL or IMAP_USER)
    try:
        send_via_mailgun_api(to_addr, "Teste Mailgun API", "Envio de teste via Mailgun API.")
        return jsonify({
            "ok": True,
            "to": to_addr,
            "domain": MAILGUN_DOMAIN,
        })
    except Exception as e:
        return jsonify({
            "ok": False,
            "to": to_addr,
            "domain": MAILGUN_DOMAIN,
            "error": str(e),
        }), 500


# -------------------------------------------------------------
# Boot
# -------------------------------------------------------------
if __name__ == "__main__":
    log.info(f"Watcher IMAP — envio primário=smtp | SMTP hosts={SMTP_HOSTS or [SMTP_HOST]}")
    log.info(f"App público em: {APP_PUBLIC_URL}")
    threading.Thread(target=watch_imap_loop, daemon=True).start()
    from werkzeug.serving import run_simple
    log.info("Iniciando Flask em 0.0.0.0:%s", PORT)
    run_simple("0.0.0.0", PORT, app, use_reloader=False)
