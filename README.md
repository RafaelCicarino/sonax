App para integracao de lista de clientes em DOCX com WhatsApp via Sonax.

Execucao em deploy (sem interface grafica):
- Chrome/Chromium roda em modo headless no servidor.
- Flags Linux aplicadas: --headless=new, --no-sandbox, --disable-dev-shm-usage.

Deploy com autenticacao:
- O login do seu navegador local nao e compartilhado com o servidor.
- Para autenticar no deploy, configure credenciais no servidor:
  - SONAX_USERNAME e SONAX_PASSWORD (env vars), ou
  - st.secrets com as mesmas chaves (ou secao [sonax] com username e password).
