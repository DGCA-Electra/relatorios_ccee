```markdown
# 🤖 RPA Envio de Emails - Streamlit & Microsoft Graph API

## 📋 Visão Geral

Este projeto é uma solução de **Automação de Processos Robóticos (RPA)** desenvolvida para otimizar o envio de relatórios da Câmara de Comercialização de Energia Elétrica (CCEE). Diferente de automações legadas baseadas em desktop, esta aplicação utiliza **Streamlit** para interface web e a **Microsoft Graph API** para integração direta com o Exchange Online, permitindo a geração de rascunhos de e-mail de forma segura, auditável e independente da máquina do usuário.

A aplicação foi projetada para suportar múltiplos analistas e diversos tipos de relatórios regulatórios (GFN, SUM, LFN, LFRES, RCAP, etc.).

---

## 🚀 Funcionalidades Principais

* **Autenticação Moderna**: Login via **Microsoft Azure AD (OAuth 2.0)** utilizando a biblioteca `MSAL`, garantindo que apenas usuários autorizados acessem a ferramenta.
* **Integração via API**: Criação de rascunhos diretamente na nuvem (pasta *Drafts* do usuário) via requisições REST à Microsoft Graph API, eliminando a necessidade do Outlook Desktop instalado.
* **Interface Web Amigável**: Painel desenvolvido em Streamlit para seleção de parâmetros (Mês, Ano, Analista) e visualização de status.
* **Multi-Relatório**: Suporte nativo e configurável para relatórios como:
    * `GFN001` e `SUM001` (Garantia Financeira e Sumário)
    * `LFN001` (Liquidação Financeira)
    * `LFRES001` (Energia de Reserva)
    * `LFRCAP001` e `RCAP002` (Reserva de Capacidade).
* **Templates Dinâmicos**: Utilização de **Jinja2** para renderização de corpos de e-mail HTML personalizados, com suporte a condicionais (ex: textos diferentes para Crédito vs. Débito).
* **Configuração Self-Service**: Interface dedicada para editar mapeamentos de Excel e templates JSON sem necessidade de alterar o código fonte.

---

## 🏗️ Arquitetura e Estrutura do Projeto

O projeto segue uma estrutura modular para facilitar a manutenção e testes:

```text
RPA-Envio-Emails-STREAMLIT/
├── .devcontainer/          # Configuração para desenvolvimento em Container
├── .github/workflows/      # Pipelines de CI (Segurança e Testes)
├── docs/                   # Documentação do projeto
├── logs/                   # Diretório de logs de execução (ex: app.log)
├── src/                    # Código fonte principal
│   ├── config/             # Gerenciamento de configurações JSON e caminhos
│   ├── handlers/           # Regras de negócio específicas por relatório
│   ├── utils/              # Utilitários de segurança, arquivos e dados
│   ├── view/               # Componentes de UI do Streamlit (Pages)
│   └── services.py         # Orquestrador de envio e comunicação com Graph API
├── static/                 # Assets estáticos (ícones, logos)
├── tests/                  # Testes unitários com Pytest
├── app.py                  # Ponto de entrada da aplicação
└── requirements.txt        # Dependências do Python
```

---

## 🛠️ Pré-requisitos e Instalação

### 1. Requisitos de Sistema

* **Python**: Versão 3.11 ou superior.
* **Acesso Azure**: Registro de Aplicativo (App Registration) no Azure AD.
* **Permissões API**: O app requer escopos `User.Read` e `Mail.ReadWrite`.

### 2. Configuração do Ambiente

Clone o repositório e instale as dependências:

```bash
git clone https://github.com/seu-repo/RPA-Envio-Emails-STREAMLIT.git
cd RPA-Envio-Emails-STREAMLIT

# Criar ambiente virtual
python -m venv venv

# Ativar ambiente (Windows)
.\venv\Scripts\Activate.ps1

# Instalar dependências
pip install -r requirements.txt
```

### 3. Variáveis de Ambiente (.env)

Crie um arquivo `.env` na raiz do projeto com as credenciais do Azure AD:

```ini
# Configurações do Azure Active Directory
AZURE_CLIENT_ID="seu_client_id_aqui"
AZURE_CLIENT_SECRET="seu_client_secret_aqui"
AZURE_TENANT_ID="seu_tenant_id_aqui"

# URI de Redirecionamento (Deve corresponder ao registrado no Azure)
# Para local: http://localhost:8501
# Para rede: https://SEU_IP:8501
AZURE_REDIRECT_URI="http://localhost:8501"
```

---

## 🖥️ Como Executar

### Execução Padrão (Localhost)

Para rodar a aplicação em sua máquina local:

```bash
streamlit run app.py
```

### Execução Segura em Rede (HTTPS)

O Azure AD exige HTTPS para URIs de redirecionamento que não sejam `localhost`. O projeto inclui um script para facilitar isso:

1. Gere certificados autoassinados (`cert.pem` e `key.pem`) com OpenSSL.
2. Execute via PowerShell:

```powershell
.\run_secure.ps1
```

### Execução via DevContainer (Docker)

Este projeto está configurado para VS Code DevContainers. Ao abrir a pasta no VS Code, aceite a sugestão para "Reopen in Container" para ter um ambiente Python 3.11 configurado automaticamente.

---

## ⚙️ Configuração de Relatórios

O sistema é altamente configurável através de arquivos JSON localizados em `src/config/`.

### Mapeamento de Dados (`config_relatorios.json`)

Define como o robô lê o Excel de dados. Exemplo para `GFN001`:

```json
"GFN001": {
    "sheet_dados": "GFN003 - Garantia Financeira po",
    "sheet_contatos": "Planilha1",
    "header_row": 30,
    "data_columns": "Agente:Empresa,Garantia Avulsa (R$):Valor",
    "path_template": {
        "excel_dados": "{sharepoint_root}/{ano}/{ano_mes}/Garantia...xlsx",
        "pdfs_dir": "{sharepoint_root}/{ano}/{ano_mes}/GFN001"
    }
}
```

### Templates de E-mail (`email_templates.json`)

Define o assunto e corpo do e-mail. Suporta variantes condicionais:

```json
"SUM001": {
    "subject_template": "SUM001 - Liquidação - {empresa}",
    "variants": {
        "credito": { "body_html": "<p>Prezado, informamos crédito de {valor}...</p>" },
        "debito": { "body_html": "<p>Prezado, informamos débito de {valor}...</p>" }
    },
    "logic": {
        "variant_selector": "situacao",
        "conditions": { "Crédito": "credito", "Débito": "debito" }
    }
}
```

---

## 🛡️ Segurança e Qualidade de Código

O projeto utiliza ferramentas robustas para garantir a segurança e padronização do código, configuradas via CI/CD:

* **Detect Secrets**: Impede o commit acidental de credenciais e chaves de API.
* **Bandit**: Análise estática de segurança (SAST) para Python.
* **Black**: Formatador de código automático.
* **Ruff**: Linter de alta performance.
* **Pip-Audit**: Verifica vulnerabilidades conhecidas nas dependências instaladas.

Para rodar as verificações localmente antes de um commit:

```bash
pre-commit run --all-files
```

---

## 🔍 Tratamento de Erros e Logs

* **Logs de Aplicação**: Armazenados em `logs/app.log`. O sistema registra todo o fluxo de processamento, incluindo falhas de autenticação, arquivos não encontrados e erros de renderização de template.
* **Interface**: Erros críticos são exibidos via `st.error` na interface do usuário para feedback imediato.
* **Sanitização**: Todo input HTML nos templates é sanitizado via biblioteca `bleach` para prevenir injeção de código (XSS).

---

## 🤝 Contribuição

1. Realize um Fork do projeto.
2. Crie uma Branch para sua Feature (`git checkout -b feature/NovaFeature`).
3. Commit suas mudanças (`git commit -m 'Adiciona Nova Feature'`).
4. Push para a Branch (`git push origin feature/NovaFeature`).
5. Abra um Pull Request.

---

**Desenvolvido por:** Malik Ribeiro Mourad  
**Licença:** Uso interno - Electra Energy
```