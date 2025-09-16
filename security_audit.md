# Auditoria de Segurança (Resumo)

## Descobertas

- Segredos no código: uso de EMAIL_USER/EMAIL_PASSWORD via dotenv (ok). Nenhum segredo hardcoded encontrado.
- Dependências: adicionar pip-audit no CI; versões fixadas `jinja2==3.1.4`, `bleach==6.1.0`.
- SAST: Bandit configurado no CI. Endurecimentos aplicados:
  - Sanitização de HTML (`utils/security_utils.sanitize_html`).
  - Sanitização de assunto (`sanitize_subject`).
  - Validação de e-mail e preview sem envio.
  - Validação de anexos (path traversal + tamanho).
- Engine Jinja2: `SandboxedEnvironment` + `StrictUndefined` e conversão de placeholders.

## Riscos e Mitigações

- XSS em previews HTML (alto) → mitigado com bleach e sandbox Jinja2.
- Email header injection (médio) → mitigado removendo CR/LF do assunto.
- Path traversal em anexos (alto) → mitigado com `is_safe_path` e limite de 25MB.
- Envio não intencional (médio) → preview não abre Outlook; envio só com `auto_send=True`.

## Ações Recomendadas

1. Ativar pre-commit: `pip install pre-commit && pre-commit install`.
2. Habilitar GitHub Actions e monitorar relatórios do workflow Security CI.
3. Rodar `pip-audit` localmente e aplicar upgrades quando necessário.
4. Se segredos forem expostos, rotacionar e remover do histórico.

## Remoção Segura de Segredos do Histórico (exemplo)

Usando git filter-repo:

```
pipx install git-filter-repo
git filter-repo --path .env --invert-paths
git push origin --force --all
```

Em seguida, rotacione credenciais (SMTP/Outlook, APIs) e invalide chaves antigas.
