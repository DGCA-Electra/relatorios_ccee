# RPA-Envio-Emails-STREAMLIT

Sistema de automação para envio de relatórios CCEE via e-mail desenvolvido em Streamlit.

## 📋 Descrição

Este projeto automatiza o processo de envio de relatórios da Câmara de Comercialização de Energia Elétrica (CCEE) para clientes, gerando e-mails personalizados com anexos PDF baseados em dados de planilhas Excel.

## 🚀 Funcionalidades

- **Automação de E-mails**: Geração automática de e-mails com Outlook
- **Múltiplos Tipos de Relatório**: Suporte a GFN001, SUM001, LFN001, LFRES, LEMBRETE, LFRCAP, RCAP
- **Interface Web**: Interface amigável desenvolvida em Streamlit
- **Configuração Flexível**: Sistema de configuração via JSON
- **Envio Multi-Analista**: Possibilidade de enviar relatórios para qualquer analista
- **Tratamento de Erros**: Sistema robusto de tratamento de erros

## 🛠️ Tecnologias Utilizadas

- **Streamlit**: Interface web
- **Pandas**: Manipulação de dados
- **OpenPyXL**: Leitura de arquivos Excel
- **PyWin32**: Integração com Microsoft Outlook
- **Pathlib**: Manipulação de caminhos de arquivo

## 📦 Instalação

### Pré-requisitos

- Python 3.8 ou superior
- Windows (para integração com Outlook)
- Microsoft Outlook instalado

### Passos de Instalação

1. **Clone o repositório**:
   ```bash
   git clone <url-do-repositorio>
   cd RPA-Envio-Emails-STREAMLIT
   ```

2. **Crie um ambiente virtual**:
   ```bash
   python -m venv venv
   ```

3. **Ative o ambiente virtual**:
   ```bash
   # Windows (PowerShell)
   .\venv\Scripts\Activate.ps1
   
   # Windows (Command Prompt)
   .\venv\Scripts\activate.bat
   
   # Linux/macOS
   source venv/bin/activate
   ```

4. **Instale as dependências**:
   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Execução

1. **Ative o ambiente virtual** (se não estiver ativo):
   ```bash
   .\venv\Scripts\Activate.ps1
   ```

2. **Execute a aplicação**:
   ```bash
   streamlit run app.py
   ```

> **Nota de UI:** Os parâmetros de envio agora aparecem apenas no painel principal. A barra lateral (sidebar) contém apenas navegação e links rápidos.

3. **Acesse no navegador**:
   - A aplicação estará disponível em `http://localhost:8501`

## 📁 Estrutura do Projeto

```
RPA-Envio-Emails-STREAMLIT/
├── app.py                 # Aplicação principal Streamlit
├── services.py            # Lógica de negócio e handlers de e-mail
├── config.py              # Configurações e utilitários
├── config_relatorios.json # Configurações dos relatórios
├── requirements.txt       # Dependências do projeto
├── README.md             # Este arquivo
├── static/               # Arquivos estáticos (logo, ícones)
├── templates/            # Templates HTML (se aplicável)
└── venv/                # Ambiente virtual (não versionado)
```

## 🔧 Configuração

### Login do Usuário

- O sistema utiliza o login de rede do usuário para configurar automaticamente os caminhos dos arquivos
- Formato esperado: `nome.sobrenome`

### Estrutura de Arquivos Esperada

```
C:/Users/{login_usuario}/
└── ELECTRA COMERCIALIZADORA DE ENERGIA S.A/
    └── GE - ECE/
        ├── DGCA/
        │   ├── DGA/
        │   │   └── CCEE/
        │   │       └── Relatórios CCEE/
        │   │           └── {ano}/
        │   │               └── {ano_mes}/
        │   │                   ├── Garantia Financeira/
        │   │                   ├── Liquidação Financeira/
        │   │                   ├── Sumário/
        │   │                   └── ...
        │   └── DGC/
        │       └── Macro/
        │           └── Contatos de E-mail para Macros.xlsx
```

### Configuração de Relatórios

Cada tipo de relatório pode ser configurado através da interface web ou diretamente no arquivo `config_relatorios.json`:

```json
{
  "GFN001": {
    "sheet_dados": "GFN003 - Garantia Financeira po",
    "sheet_contatos": "Planilha1",
    "header_row": 30,
    "data_columns": "Agente:Empresa,Garantia Avulsa (R$):Valor"
  }
}
```

## 📊 Tipos de Relatório Suportados

| Tipo | Descrição | Arquivo de Dados |
|------|-----------|------------------|
| GFN001 | Garantia Financeira | GFN003 |
| SUM001 | Sumário da Liquidação Financeira | LFN004 |
| LFN001 | Liquidação Financeira | LFN004 |
| LFRES | Liquidação da Energia de Reserva | LFRES002 |
| LEMBRETE | Lembrete de Aporte | GFN003 |
| LFRCAP | Liquidação de Reserva de Capacidade | LFRCAP002 |
| RCAP | Reserva de Capacidade | RCAP002 |

## 🔍 Uso

### 1. Login
- Acesse a aplicação e faça login com seu usuário de rede
- O sistema configurará automaticamente os caminhos dos arquivos

### 2. Seleção de Parâmetros
- Escolha o tipo de relatório
- Selecione o mês e ano
- Clique em "Pré-visualizar Dados"

### 3. Processamento
- O sistema carregará os dados das planilhas
- Filtrará por analista responsável
- Gerará e-mails no Outlook para revisão

### 4. Envio Multi-Analista
- Qualquer usuário pode enviar relatórios para qualquer analista
- Útil durante férias ou ausências, quando um analista precisa enviar relatórios para outro

## ⚙️ Configurações Avançadas

### Personalização de Caminhos

Os caminhos são configurados automaticamente, mas podem ser personalizados editando `config.py`:

```python
PATH_CONFIGS = {
    "sharepoint_root": "ELECTRA COMERCIALIZADORA DE ENERGIA S.A/GE - ECE/DGCA/DGA/CCEE/Relatórios CCEE",
    "contatos_email": "ELECTRA COMERCIALIZADORA DE ENERGIA S.A/GE - ECE/DGCA/DGC/Macro/Contatos de E-mail para Macros.xlsx",
    "user_base": "C:/Users"
}
```

### Adicionando Novos Tipos de Relatório

1. Adicione a configuração em `config.py`:
```python
DEFAULT_CONFIGS["NOVO_TIPO"] = {
    "sheet_dados": "Nome da Aba",
    "sheet_contatos": "Planilha1",
    "header_row": 0,
    "data_columns": "Coluna1:Map1,Coluna2:Map2"
}
```

2. Crie o handler em `services.py`:
```python
def handle_novo_tipo(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Dict[str, Any]:
    # Lógica do handler
    pass

REPORT_HANDLERS['NOVO_TIPO'] = handle_novo_tipo
```

## 🐛 Tratamento de Erros

O sistema inclui tratamento robusto de erros:

- **Arquivos não encontrados**: Verificação de existência de arquivos
- **Configurações inválidas**: Validação de configurações
- **Dados ausentes**: Tratamento de dados faltantes
- **Erros de Outlook**: Tratamento de falhas na integração

## 📝 Logs

- Os logs são salvos em `app.log`
- Incluem informações de erro e processamento
- Útil para debugging e monitoramento

## 🔒 Segurança

- Login baseado em usuário de rede
- Validação de formatos de entrada
- Tratamento seguro de caminhos de arquivo
- Logs para auditoria

## 🤝 Contribuição

Para contribuir com o projeto:

1. Faça um fork do repositório
2. Crie uma branch para sua feature
3. Implemente as mudanças
4. Teste adequadamente
5. Submeta um pull request

## 📄 Licença

Este projeto é de uso interno da ELECTRA COMERCIALIZADORA DE ENERGIA S.A.

## 👥 Autores

- Desenvolvido para DGCA
- Mantido pela equipe de desenvolvimento

## 📞 Suporte

Para suporte técnico ou dúvidas, entre em contato com a equipe de desenvolvimento.

---

**Versão**: 1.0.0  
**Última atualização**: Julho 2025
