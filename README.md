# Casos Hemodinâmica

Formulário web para registrar casos de hemodinâmica direto no OneDrive pessoal (Microsoft Graph).
Nenhum dado de paciente fica neste repositório: tudo é gravado na pasta `Casos Hemodinâmica` do OneDrive de quem faz login.

- `config.js` — Client ID do registro no Azure e pasta no OneDrive
- `app.js` — lógica (login, envio de arquivos, casos.json, planilha)
- `lib/` — MSAL (login Microsoft) e ExcelJS, servidos localmente
