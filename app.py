#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# ****** COBOL Support Agent — v10.26 ********
# ****** Andre Richest                ********
# ****** Wed Dec 03 2025              ********

import os
import ssl
import time
import json
import socket
import logging
import threading
import hashlib
from typing import Optional, Tuple, List

import imaplib
import smtplib
from email.message import EmailMessage
from email import message_from_bytes
from email.utils import parseaddr
from datetime import datetime, timedelta

import requests  # <-- Mailgun API + LLM HTTP

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

# Pasta onde a resposta enviada deve aparecer (Sent Items)
IMAP_FOLDER_SENT = os.getenv("IMAP_FOLDER_SENT", "Enviados")

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

# --- LLM config (genérico via HTTP / OpenRouter) ---
# Mapeia automaticamente variáveis OPENROUTER_* se LLM_* não forem definidas
LLM_API_URL = (
    os.getenv("LLM_API_URL")
    or "https://openrouter.ai/api/v1/chat/completions"
)
LLM_API_KEY = os.getenv("LLM_API_KEY") or os.getenv("OPENROUTER_API_KEY", "")
LLM_MODEL = os.getenv("LLM_MODEL") or os.getenv("OPENROUTER_MODEL", "")
LLM_TIMEOUT = int(os.getenv("LLM_TIMEOUT", "60"))
LLM_MAX_RETRIES = int(os.getenv("LLM_MAX_RETRIES", "2"))
LLM_BLOCK_SECONDS = int(os.getenv("LLM_BLOCK_SECONDS", "300"))
LLM_HARD_DISABLE = os.getenv("LLM_HARD_DISABLE", "false").lower() == "true"

# Headers extras para OpenRouter (se aplicável)
OPENROUTER_SITE_URL = os.getenv("OPENROUTER_SITE_URL", APP_PUBLIC_URL)
OPENROUTER_APP_NAME = os.getenv("OPENROUTER_APP_NAME", APP_TITLE)

_llm_block_until_ts: float = 0.0

# ==========================
# Prompt do sistema
# ==========================
SYSTEM_PROMPT = (
    "Você é um assistente do time de suporte da Aprenda COBOL, ajudando alunos com dúvidas e correções de código em COBOL, JCL e Visualg (Portugol). "
    "E-mails da Hotmart ou originados com o remetente 'noreply' não devem ser respondidos; devem ser ignorados por você. "
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
    "8) REGRA CRÍTICA DE ANEXOS (SEMPRE VALE): Mesmo que o e-mail venha sem texto, sem contexto ou sem perguntas, "
    "   você DEVE ler todos os arquivos anexados, identificar a linguagem/propósito e fazer uma análise com recomendações objetivas. "
    "   Considera-se 'pedido suficiente' a mera presença de anexos de código: nesse caso, responda com 'acao'='responder' "
    "   e uma análise útil, a menos que os anexos estejam corrompidos/ilegíveis. "
    "9) Se houver anexos de código, priorize analisar o conteúdo do anexo e a linguagem correta: "
    "   9.1) COBOL: arquivos .COB/.CBL/.CPY. Cite elementos como DIVISION, SECTION, PIC, níveis (01/05/77/88), "
    "        WORKING-STORAGE, FILE SECTION, FD, SELECT/ASSIGN, ORGANIZATION, PERFORM, IF/EVALUATE, STRING/UNSTRING, "
    "        tabelas OCCURS, REDEFINES, e quando aplicável SQLCA/EXEC SQL. Identifique erros comuns (tipagem, pontuação, "
    "        colunas, incompatibilidade de PIC, falta de períodos, problemas em I/O) e sugira correções objetivas. "
    "   9.2) JCL: arquivos .JCL/.JOB/.TXT contendo JCL. Cite JOB/EXEC/DD, PROC/INCLUDE, PARM, COND, DISP, DSN, SPACE, "
    "        DCB, LRECL/RECFM/BLKSIZE, SYSIN, STEPLIB, e mensagens/abends quando fornecidas. Sugira melhorias de clareza, "
    "        padronização, reuso via PROC, e correções de parâmetros. "
    "   9.3) Visualg (Portugol): arquivos .ALG/.TXT contendo algoritmos. Analise sintaxe e lógica: inicio/fimalgoritmo, "
    "        var, tipos (inteiro/real/logico/caractere), leia/escreva, se/senao, escolha/caso, para, enquanto, repita, "
    "        vetores/matrizes, funções/procedimentos, operadores relacionais e lógicos. Aponte erros de digitação, variáveis "
    "        não declaradas/inicializadas, conversões de tipo, laços infinitos e melhorias de legibilidade. "
    "   9.4) Se a linguagem do código estiver incerta, deduza pelo conteúdo e responda focando na linguagem mais provável, "
    "        deixando isso explícito no corpo (ex.: \"Pelo seu código, parece Visualg...\"). "
    "10) Não mude o tema da conversa. Responda ao que foi solicitado, de forma educada e objetiva, sempre como parte de um time (nós). "
    "11) Se faltar informação para compilar/executar/testar, peça os dados mínimos: "
    "   - COBOL: compilador/ambiente (GnuCOBOL/Enterprise), SO, JCL (se batch), layouts/FD, amostras de entrada/saída, "
    "     mensagens de erro e comando de compilação. "
    "   - JCL: objetivo do job, PROC envolvidos, dataset names, mensagens/RC/abends, e trechos completos dos passos relevantes. "
    "   - Visualg: enunciado do exercício, entradas esperadas e saídas esperadas, e a versão do Visualg (se souber). "
    "   Observação: Mesmo quando pedir esses dados mínimos, ainda assim entregue a melhor análise possível do que estiver disponível no anexo. "
    "12) No final do 'corpo_markdown', SEMPRE inclua um parágrafo curto de incentivo, parabenizando o aluno por estar estudando "
    "    e motivando a continuar. Em seguida, inclua exatamente estas duas linhas (URLs como texto puro, sem markdown de link): "
    "- Nossa Comunidade no Telegram: https://t.me/aprendacobol "
    "- Conheça a Formação Completa de Programador COBOL: https://assinatura.aprendacobol.com.br "
)
SYSTEM_PROMPT_SHA1 = hashlib.sha1(SYSTEM_PROMPT.encode("utf-8")).hexdigest()
log.info(
    "SYSTEM_PROMPT_SHA1=%s (primeiros 120 chars): %s",
    SYSTEM_PROMPT_SHA1[:12],
    SYSTEM_PROMPT[:120],
)

# ==========================
# Helpers LLM
# ==========================
def _llm_is_blocked_now() -> bool:
    return LLM_HARD_DISABLE or (time.time() < _llm_block_until_ts)


def _llm_mark_blocked(seconds: int):
    global _llm_block_until_ts
    _llm_block_until_ts = time.time() + max(0, seconds)


def _parse_llm_json(raw: str) -> dict:
    """
    Tenta extrair JSON de um texto retornado pela LLM,
    mesmo que ela coloque algo antes/depois.
    """
    s = raw.strip()
    start = s.find("{")
    end = s.rfind("}")
    if start != -1 and end != -1 and end > start:
        s = s[start : end + 1]
    return json.loads(s)


def _call_llm(user_prompt: str) -> dict:
    """
    Chama o endpoint HTTP da LLM (estilo OpenAI / OpenRouter),
    usando SYSTEM_PROMPT no role=system e retornando o dict parseado.
    """
    if not LLM_API_URL or not LLM_MODEL:
        raise RuntimeError("LLM_API_URL/LLM_MODEL não configurados")

    headers = {"Content-Type": "application/json"}

    if LLM_API_KEY:
        headers["Authorization"] = f"Bearer {LLM_API_KEY}"

    # Headers específicos exigidos pela OpenRouter
    if "openrouter.ai" in LLM_API_URL:
        if OPENROUTER_SITE_URL:
            headers["HTTP-Referer"] = OPENROUTER_SITE_URL
        if OPENROUTER_APP_NAME:
            headers["X-Title"] = OPENROUTER_APP_NAME

    payload = {
        "model": LLM_MODEL,
        "messages": [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.2,
    }

    last_err = None
    for attempt in range(1, LLM_MAX_RETRIES + 1):
        try:
            log.info(f"Chamando LLM em {LLM_API_URL} (tentativa {attempt}/{LLM_MAX_RETRIES}, model={LLM_MODEL})")
            resp = requests.post(
                LLM_API_URL,
                headers=headers,
                json=payload,
                timeout=LLM_TIMEOUT,
            )
            if resp.status_code != 200:
                # Log detalhado para debug (inclui início do corpo da resposta)
                try:
                    snippet = resp.text[:500]
                except Exception:
                    snippet = "<não foi possível decodificar corpo>"
                log.warning(
                    f"LLM HTTP {resp.status_code} em {LLM_API_URL} — corpo parcial: {snippet}"
                )
                resp.raise_for_status()

            data = resp.json()
            # Tenta formato estilo chat.completions OpenAI/OpenRouter
            content = None
            try:
                content = data["choices"][0]["message"]["content"]
            except Exception:
                if "choices" in data and data["choices"]:
                    ch = data["choices"][0]
                    if "text" in ch:
                        content = ch["text"]
            if not content:
                raise RuntimeError("Resposta LLM em formato inesperado")
            return _parse_llm_json(content)
        except Exception as e:
            last_err = e
            log.warning(f"Chamada LLM falhou (tentativa {attempt}/{LLM_MAX_RETRIES}): {e}")
            time.sleep(1)

    raise RuntimeError(f"LLM indisponível após {LLM_MAX_RETRIES} tentativas: {last_err}")


def _extract_email_plaintext(msg) -> str:
    """
    Extrai o corpo em texto simples do e-mail, priorizando text/plain.
    """
    def decode_payload(part):
        payload = part.get_payload(decode=True) or b""
        charset = part.get_content_charset() or "utf-8"
        try:
            return payload.decode(charset, errors="replace")
        except Exception:
            return payload.decode("utf-8", errors="replace")

    if msg.is_multipart():
        # Primeiro tenta text/plain sem ser attachment
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get("Content-Disposition") or "").lower()
            if ctype == "text/plain" and "attachment" not in disp:
                return decode_payload(part)
        # Depois tenta text/html
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get("Content-Disposition") or "").lower()
            if ctype == "text/html" and "attachment" not in disp:
                return decode_payload(part)
        return ""
    else:
        return decode_payload(msg)


def _extract_cobol_attachments(msg) -> str:
    """
    Retorna string com conteúdo (truncado) de anexos COBOL (.cob, .cbl, .cpy).
    """
    texts = []
    for part in msg.walk():
        disp = (part.get("Content-Disposition") or "").lower()
        fname = part.get_filename()
        if "attachment" in disp and fname:
            lf = fname.lower()
            if lf.endswith((".cob", ".cbl", ".cpy")):
                payload = part.get_payload(decode=True) or b""
                charset = part.get_content_charset() or "utf-8"
                try:
                    content = payload.decode(charset, errors="replace")
                except Exception:
                    content = payload.decode("utf-8", errors="replace")
                if len(content) > 8000:
                    content = content[:8000]
                texts.append(f"Arquivo: {fname}\n{content}")
    return "\n\n".join(texts)


def _default_reply_dict(subject: str) -> dict:
    """
    Resposta de fallback caso LLM esteja indisponível.
    """
    corpo = (
        "Olá!\n\n"
        "Somos do time de suporte e confirmamos que recebemos o seu e-mail. Estamos com um backlog grande de mensagens, "
        "e vamos tratar cada uma com muita atenção, portanto agradecemos sua compreensão se demorar um pouco sua resposta.\n\n"
        "Queremos dar à você a nossa melhor atenção.\n"
        "\n"
        "- Nossa Comunidade no Telegram: https://t.me/aprendacobol\n"
        "- Conheça a Formação Completa de Programador COBOL: https://assinatura.aprendacobol.com.br"
    )
    return {
        "assunto": subject or "(sem assunto)",
        "corpo_markdown": corpo,
        "nivel_confianca": 0.0,
        "acao": "responder",
    }


def _generate_reply_for_message(msg) -> Optional[dict]:
    """
    Gera o dict {assunto, corpo_markdown, nivel_confianca, acao}
    usando a LLM; se bloqueada e fallback habilitado, usa resposta padrão;
    se bloqueada e fallback desabilitado, retorna None (vai para Escalar).
    """
    subject = msg.get("Subject", "(sem assunto)")

    if _llm_is_blocked_now():
        log.warning("LLM está bloqueada no momento.")
        if IMAP_FALLBACK_WHEN_LLM_BLOCKED:
            return _default_reply_dict(subject)
        else:
            return None

    # Se LLM não está configurada, trata como 'hard disable'
    if (not LLM_API_URL or not LLM_MODEL or not LLM_API_KEY) or LLM_HARD_DISABLE:
        log.info("LLM desativada ou não configurada; usando resposta padrão.")
        return _default_reply_dict(subject)

    from_addr = parseaddr(msg.get("From", ""))[1]
    body_txt = _extract_email_plaintext(msg)
    cobol_txt = _extract_cobol_attachments(msg)

    user_prompt = (
        f"Remetente: {from_addr}\n"
        f"Assunto original: {subject}\n\n"
        f"Corpo do e-mail (texto):\n{body_txt}\n\n"
    )
    if cobol_txt:
        user_prompt += f"Anexos COBOL:\n{cobol_txt}\n\n"

    try:
        reply = _call_llm(user_prompt)
        # Garantir chaves básicas
        for k in ("assunto", "corpo_markdown", "nivel_confianca", "acao"):
            reply.setdefault(k, "")
        return reply
    except Exception as e:
        log.error("Falha ao obter resposta da LLM; marcando como bloqueada.", exc_info=True)
        _llm_mark_blocked(LLM_BLOCK_SECONDS)
        if IMAP_FALLBACK_WHEN_LLM_BLOCKED:
            return _default_reply_dict(subject)
        else:
            return None


# -------------------------------------------------------------
# Helpers gerais
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


def _compose_full_text(body: str) -> str:
    """
    Garante que TODAS as saídas tenham a mesma assinatura padrão
    com nome, footer e links (cursos, etc.).
    """
    return body + f"\n\n{SIGNATURE_NAME}\n{SIGNATURE_FOOTER}\n{SIGNATURE_LINKS}"


def _build_from_header() -> str:
    email = SMTP_FROM_EMAIL or SMTP_USER
    name = (SMTP_FROM_NAME or "").strip()
    if email and name:
        return f"{name} <{email}>"
    return email or ""


def _append_to_sent_imap(imap, to_addr: str, subject: str, full_body: str):
    """
    Grava a cópia da resposta na pasta de itens enviados (IMAP_FOLDER_SENT),
    já como lida (\\Seen), para atender o requisito de aparecer em 'Sent Items'.
    Faz também fallback automático para INBOX.<pasta> se o servidor exigir namespace.
    """
    if not IMAP_FOLDER_SENT:
        return
    try:
        msg_out = EmailMessage()
        from_header = _build_from_header() or IMAP_USER
        if from_header:
            msg_out["From"] = from_header
        msg_out["To"] = to_addr
        msg_out["Subject"] = subject
        if SMTP_REPLY_TO:
            msg_out["Reply-To"] = SMTP_REPLY_TO
        msg_out.set_content(full_body)

        raw_out = msg_out.as_bytes()

        # Tentativa direta na pasta configurada
        try:
            typ, data = imap.append(IMAP_FOLDER_SENT, "\\Seen", None, raw_out)
        except Exception as e:
            log.warning(f"Falha ao gravar resposta em {IMAP_FOLDER_SENT}: {e}")
            typ, data = "ERR", None

        if typ == "OK":
            log.info(f"Resposta gravada em pasta de enviados: {IMAP_FOLDER_SENT}")
            return
        else:
            log.warning(
                f"Falha ao gravar resposta em {IMAP_FOLDER_SENT}: typ={typ}, data={data}"
            )

        # Fallback com namespace INBOX., caso IMAP_FOLDER_SENT não seja totalmente qualificado
        if not IMAP_FOLDER_SENT.upper().startswith("INBOX."):
            fallback_folder = f"INBOX.{IMAP_FOLDER_SENT}"
            try:
                typ2, data2 = imap.append(fallback_folder, "\\Seen", None, raw_out)
                if typ2 == "OK":
                    log.info(
                        f"Resposta gravada em pasta de enviados (fallback): {fallback_folder}"
                    )
                else:
                    log.warning(
                        f"Falha ao gravar resposta em {fallback_folder} (fallback): "
                        f"typ={typ2}, data={data2}"
                    )
            except Exception as e2:
                log.warning(
                    f"Erro ao gravar resposta em {fallback_folder} (fallback): {e2}"
                )
    except Exception as e:
        log.warning(f"Falha geral ao preparar resposta para {IMAP_FOLDER_SENT}: {e}")


def _copy_with_namespace_fallback(imap, msg_id: bytes, folder: str, context: str):
    """
    Tenta copiar para 'folder'. Se o servidor reclamar de namespace inexistente,
    tenta novamente com 'INBOX.' como prefixo (ex.: INBOX.Escalar).
    """
    if not folder:
        return

    # Tentativa direta
    try:
        typ, data = imap.copy(msg_id, folder)
        if typ == "OK":
            log.info(f"Mensagem ID={msg_id} copiada para {folder} ({context})")
            return
        else:
            log.warning(
                f"Falha ao copiar mensagem ID={msg_id} para {folder} ({context}): "
                f"typ={typ}, data={data}"
            )
    except Exception as e:
        log.warning(
            f"Erro ao copiar mensagem ID={msg_id} para {folder} ({context}): {e}"
        )

    # Fallback com namespace INBOX.
    fallback_folder = f"INBOX.{folder}"
    try:
        typ, data = imap.copy(msg_id, fallback_folder)
        if typ == "OK":
            log.info(f"Mensagem ID={msg_id} copiada para {fallback_folder} ({context})")
        else:
            log.warning(
                f"Falha ao copiar mensagem ID={msg_id} para {fallback_folder} ({context}): "
                f"typ={typ}, data={data}"
            )
    except Exception as e:
        log.warning(
            f"Erro ao copiar mensagem ID={msg_id} para {fallback_folder} ({context}): {e}"
        )


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

def send_via_mailgun_api(to_addr: str, subject: str, body: str) -> str:
    if not MAILGUN_API_KEY or not MAILGUN_DOMAIN:
        raise RuntimeError("Mailgun API não configurada (MAILGUN_API_KEY/MAILGUN_DOMAIN)")

    from_header = _build_from_header()
    text_body = _compose_full_text(body)

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
        msg.set_content(_compose_full_text(body))
        s.send_message(msg)
        return "ok"
    finally:
        try:
            s.quit()
        except Exception:
            pass


# -------------------------------------------------------------
# IMAP: busca e processamento básico de mensagens
# -------------------------------------------------------------

def _search_messages(imap) -> List[bytes]:
    """
    Retorna uma lista de IDs (em bytes) de mensagens candidatas
    para processamento, respeitando IMAP_STRICT_UNSEEN_ONLY e IMAP_SINCE_DAYS.
    """
    criteria: List[str] = []

    # Base: UNSEEN ou ALL
    if IMAP_STRICT_UNSEEN_ONLY:
        criteria.append("UNSEEN")
    else:
        criteria.append("ALL")

    # SINCE (opcional)
    if IMAP_SINCE_DAYS > 0:
        since_date = (datetime.utcnow() - timedelta(IMAP_SINCE_DAYS)).strftime("%d-%b-%Y")
        criteria.extend(["SINCE", since_date])

    try:
        typ, data = imap.search(None, *criteria)
        if typ != "OK" or not data or not data[0]:
            return []
        ids = data[0].split()
        return ids
    except Exception as e:
        log.error(f"IMAP SEARCH falhou: {e}", exc_info=True)
        return []


def _should_skip_message(msg) -> bool:
    """
    Regras simples para evitar loop de auto-resposta, etc.
    """
    from_addr = parseaddr(msg.get("From", ""))[1].lower()
    to_addr = parseaddr(msg.get("To", ""))[1].lower()

    support_addrs = set([
        (SMTP_USER or "").lower(),
        (SMTP_FROM_EMAIL or "").lower(),
        IMAP_USER.lower(),
    ])

    # Não responder e-mails que nós mesmos enviamos
    if from_addr in support_addrs:
        log.info(f"Pulando mensagem de {from_addr} (provavelmente nós mesmos)")
        return True

    # Evitar auto-resposta para notificações típicas
    subj = (msg.get("Subject", "") or "").lower()
    if "mailer-daemon" in from_addr or "postmaster@" in from_addr:
        log.info(f"Pulando bounce/mail daemon: {from_addr}")
        return True

    # Filtrar rápido notificações Hotmart/noreply aqui também (além do prompt)
    if "hotmart" in from_addr or "noreply" in from_addr:
        log.info(f"Pulando notificação automatizada: {from_addr}")
        return True

    # Exemplo de filtro simples por assunto (ajuste se quiser)
    if subj.startswith("re:") and from_addr == to_addr:
        log.info(f"Pulando potencial loop de resposta para {from_addr}")
        return True

    return False


def _process_single_message(imap, msg_id: bytes):
    """
    Processa UMA mensagem:
    - faz FETCH
    - parseia
    - gera resposta via LLM (ou fallback)
    - envia resposta (sempre que houver corpo e remetente, mesmo com acao='escalar')
    - grava cópia da resposta em IMAP_FOLDER_SENT (se houver resposta)
    - se escalável: copia APENAS para FOLDER_ESCALATE (como não lido lá),
      marca original como lida e opcionalmente remove da INBOX.
    - se não escalável: marca como lida, copia para FOLDER_PROCESSED e
      opcionalmente remove da INBOX.
    """
    try:
        typ, data = imap.fetch(msg_id, "(RFC822)")
        if typ != "OK" or not data or not data[0]:
            log.warning(f"FETCH falhou para ID {msg_id}")
            return

        raw = data[0][1]
        msg = message_from_bytes(raw)

        from_addr = parseaddr(msg.get("From", ""))[1]
        subject = msg.get("Subject", "(sem assunto)")

        log.info(f"Processando mensagem ID={msg_id} de={from_addr} assunto={subject!r}")

        if _should_skip_message(msg):
            return

        reply = _generate_reply_for_message(msg)

        # Se reply=None, significa que LLM está bloqueada e fallback desabilitado:
        # não responde, apenas manda para Escalar.
        if reply is None:
            log.info("Nenhuma resposta automática gerada. Encaminhando e-mail para Escalar.")
            try:
                imap.store(msg_id, "+FLAGS", "\\Seen")
            except Exception as e:
                log.warning(f"Falha ao marcar \\Seen para ID={msg_id}: {e}")

            if FOLDER_ESCALATE:
                _copy_with_namespace_fallback(imap, msg_id, FOLDER_ESCALATE, "escalar por falha LLM")

            if EXPUNGE_AFTER_COPY:
                try:
                    imap.store(msg_id, "+FLAGS", "\\Deleted")
                    log.info(f"Mensagem ID={msg_id} marcada como \\Deleted (após falha LLM)")
                except Exception as e:
                    log.warning(f"Falha ao marcar \\Deleted para ID={msg_id}: {e}")
            return

        # Normaliza resposta da LLM / fallback
        llm_subject = reply.get("assunto") or subject or "(sem assunto)"
        corpo_markdown = reply.get("corpo_markdown") or ""
        acao = (reply.get("acao") or "responder").lower().strip()
        nivel_confianca = reply.get("nivel_confianca")

        reply_subject = llm_subject
        if not reply_subject.lower().startswith("re:"):
            reply_subject = f"Re: {reply_subject}"

        # Decide se escalável
        escalate = (acao == "escalar")

        full_body = _compose_full_text(corpo_markdown)

        # Envia resposta para o remetente original (SEMPRE que houver corpo e remetente),
        # independentemente de 'acao' ser 'responder' ou 'escalar'
        if from_addr and corpo_markdown.strip():
            send_test_email(
                to_addr=from_addr,
                subject=reply_subject,
                body=corpo_markdown,
            )
            log.info(
                f"Resposta enviada para {from_addr} (acao={acao}, nivel_confianca={nivel_confianca})"
            )

            # Grava cópia na pasta de enviados (Sent Items)
            _append_to_sent_imap(imap, from_addr, reply_subject, full_body)
        elif not from_addr:
            log.warning(f"Mensagem ID={msg_id} sem remetente válido; não foi possível responder")
        else:
            log.info(
                f"Mensagem ID={msg_id} com resposta vazia; nada enviado, mas fluxo de pastas continuará (acao={acao})."
            )

        # MOVIMENTAÇÃO EM PASTAS
        if escalate:
            # 1) Copiar APENAS para FOLDER_ESCALATE, mantendo como 'não lido' lá.
            if FOLDER_ESCALATE:
                _copy_with_namespace_fallback(imap, msg_id, FOLDER_ESCALATE, "escalar")

            # 2) Agora marcar original como lida (para não reprocessar)
            try:
                imap.store(msg_id, "+FLAGS", "\\Seen")
            except Exception as e:
                log.warning(f"Falha ao marcar \\Seen para ID={msg_id}: {e}")

            # 3) Opcionalmente marcar para remoção da INBOX
            if EXPUNGE_AFTER_COPY:
                try:
                    imap.store(msg_id, "+FLAGS", "\\Deleted")
                    log.info(f"Mensagem ID={msg_id} marcada como \\Deleted (após escalar)")
                except Exception as e:
                    log.warning(f"Falha ao marcar \\Deleted para ID={msg_id}: {e}")
        else:
            # Fluxo normal: marcar como lida, copiar para Respondidos e opcionalmente remover da INBOX
            try:
                imap.store(msg_id, "+FLAGS", "\\Seen")
            except Exception as e:
                log.warning(f"Falha ao marcar \\Seen para ID={msg_id}: {e}")

            if FOLDER_PROCESSED:
                _copy_with_namespace_fallback(imap, msg_id, FOLDER_PROCESSED, "processado")

            if EXPUNGE_AFTER_COPY:
                try:
                    imap.store(msg_id, "+FLAGS", "\\Deleted")
                    log.info(f"Mensagem ID={msg_id} marcada como \\Deleted")
                except Exception as e:
                    log.warning(f"Falha ao marcar \\Deleted para ID={msg_id}: {e}")

    except Exception as e:
        log.error(f"Erro ao processar mensagem ID={msg_id}: {e}", exc_info=True)


def process_inbox_once():
    """
    Abre conexão IMAP, seleciona INBOX, busca mensagens candidatas e processa.
    É chamado em loop pelo watcher.
    """
    try:
        log.info(
            f"IMAP tentando login como {IMAP_USER[:2]}***@ "
            f"em {IMAP_HOST}:{IMAP_PORT} (mode={IMAP_TLS_MODE})"
        )
        imap = imap_connect(IMAP_HOST, IMAP_PORT, IMAP_USER, IMAP_PASS, IMAP_TLS_MODE)
        try:
            typ, _ = imap.select(IMAP_FOLDER_INBOX)
            if typ != "OK":
                raise RuntimeError(f"SELECT {IMAP_FOLDER_INBOX} falhou: {typ}")
            log.info("IMAP INBOX selecionada, iniciando busca de mensagens")

            ids = _search_messages(imap)
            if not ids:
                log.info("Nenhuma mensagem nova para processar")
            else:
                log.info(f"Encontradas {len(ids)} mensagens para processar: {ids}")
                for msg_id in ids:
                    _process_single_message(imap, msg_id)

            # Se deletamos algo, dá um expunge
            if EXPUNGE_AFTER_COPY:
                try:
                    imap.expunge()
                except Exception as e:
                    log.warning(f"EXPUNGE falhou: {e}")
        finally:
            try:
                imap.logout()
            except Exception:
                pass
    except ImapAuthError as e:
        log.error(f"IMAP AUTH falhou: {e}")
    except Exception as e:
        log.error("process_inbox_once falhou", exc_info=e)


# -------------------------------------------------------------
# Watcher (loop de processamento IMAP)
# -------------------------------------------------------------

def watch_imap_loop():
    log.info(
        f"IMAP_STRICT_UNSEEN_ONLY={IMAP_STRICT_UNSEEN_ONLY} | "
        f"IMAP_SINCE_DAYS={IMAP_SINCE_DAYS} | "
        f"IMAP_FALLBACK_LAST_N={IMAP_FALLBACK_LAST_N} | "
        f"IMAP_FALLBACK_WHEN_LLM_BLOCKED={IMAP_FALLBACK_WHEN_LLM_BLOCKED}"
    )
    while True:
        process_inbox_once()
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
        },
        "llm": {
            "api_url": LLM_API_URL,
            "model": LLM_MODEL,
            "hard_disable": LLM_HARD_DISABLE,
            "has_api_key": bool(LLM_API_KEY),
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
