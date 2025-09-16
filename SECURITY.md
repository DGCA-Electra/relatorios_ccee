# Security Policy

## Reportar Vulnerabilidades

- Envie um e-mail para security@empresa.com com detalhes, PoC e impacto esperado.
- SLA de resposta inicial: 5 dias úteis.

## Controles Implementados

- Sanitização de HTML com bleach e StrictUndefined no Jinja2
- Normalização de placeholders e bloqueio de header injection no assunto
- Validação e higienização de e-mails e caminhos de anexos (path traversal)
- Limite de tamanho de anexos (25MB)
- Pré-visualização em tela (sem envio automático)
- CI com Bandit e pip-audit
- Pre-commit com detect-secrets, black, ruff e bandit

## Rodar Scans Localmente

```
pip install -r requirements.txt
pip install bandit pip-audit detect-secrets
bandit -q -r .
pip-audit -r requirements.txt
```

## Política de auto_send

- Envio automático só é permitido quando auto_send=True explicitamente.
- Recomenda-se proteger via variável de ambiente AUTO_SEND_ALLOWED=false e checar antes de enviar em produção.

## Gestão de Segredos

- Use .env local, não commitar.
- Habilite detect-secrets: pre-commit install.
- Caso um segredo seja cometido, siga o guia no security_audit.md para remoção/rotação.
