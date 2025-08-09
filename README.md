# COBOL Support Agent — v6 (Render Free + HostGator)

**Leitura IMAP + Envio SMTP no HostGator**. WebService Flask (para Render Free)
e watcher IMAP em thread de fundo.

## Deploy (Render)
1. Suba estes arquivos no seu repositório (app.py, requirements.txt, render.yaml, prompts.py, ollama_client.py, .env.example).
2. No Render: **New → Blueprint** → selecione este repositório.
3. Em **Environment → Variables**, preencha:
   - MAIL_USER, MAIL_PASS (os dois estão com `sync:false` no render.yaml)
4. Deploy. Verifique `/health` e `/diag/imap`.

## Variáveis chave
- `IMAP_HOST/PORT`, `MAIL_USER/PASS` → leitura (INBOX).
- `SMTP_HOST/PORT`, `SMTP_TLS_MODE` → envio por HostGator.
- `FOLDER_PROCESSED`, `FOLDER_ESCALATE` → pastas destino (serão normalizadas p/ `INBOX.*`).
- `SENT_FOLDER` → cópia do e-mail enviado via APPEND (ex.: `INBOX.Sent`).

## Notas
- No Render Free não há Ollama: o fluxo usa fallback seguro (ação `escalar`).
  Para IA de verdade no Render, considere um provedor HTTP (ex.: OpenRouter) — posso te enviar uma variante quando quiser.
