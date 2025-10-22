# 🤖 RPA-Envio-Emails-STREAMLIT

## Automação Inteligente para Envio de Relatórios CCEE via E-mail com Streamlit

Este projeto inovador oferece uma solução de **Automação de Processos Robóticos (RPA)** para otimizar o envio de relatórios da Câmara de Comercialização de Energia Elétrica (CCEE) a clientes. Desenvolvido com **Streamlit**, ele proporciona uma interface web intuitiva para a geração e envio automatizado de e-mails personalizados, acompanhados de anexos em PDF, com base em dados extraídos de planilhas Excel.

--- 

## ✨ Funcionalidades Principais

O sistema foi projetado para oferecer uma experiência robusta e flexível, destacando-se por:

-   **Automação de E-mails**: Geração e envio automático de e-mails através da integração com o Microsoft Outlook, permitindo a criação de rascunhos para revisão ou envio direto.
-   **Suporte a Múltiplos Relatórios CCEE**: Compatibilidade com diversos tipos de relatórios, incluindo GFN001, SUM001, LFN001, LFRES, LEMBRETE, LFRCAP e RCAP, garantindo cobertura abrangente das necessidades da CCEE.
-   **Interface Web Intuitiva (Streamlit)**: Uma aplicação web amigável que simplifica a interação do usuário, tornando o processo de envio de relatórios acessível mesmo para usuários não técnicos.
-   **Configuração Dinâmica**: Permite a configuração flexível de parâmetros via interface web ou arquivos JSON, adaptando-se facilmente a novas necessidades ou mudanças nos formatos de relatório.
-   **Envio Multi-Analista**: Capacidade de qualquer usuário enviar relatórios em nome de qualquer analista, crucial para cenários de férias, ausências ou delegação de tarefas.
-   **Tratamento de Erros Robusto**: Mecanismos avançados de tratamento de erros para garantir a resiliência do sistema, com logs detalhados para diagnóstico e monitoramento.
-   **Engine de Templates Jinja2**: Utilização de templates Jinja2 para a criação dinâmica de assuntos e corpos de e-mail, permitindo alta personalização e flexibilidade na comunicação.
-   **Validação de Anexos**: Verificação automática da existência e do tamanho dos arquivos anexados, prevenindo erros de envio e garantindo a conformidade.

--- 

## 🛠️ Tecnologias Utilizadas

Este projeto foi construído com uma pilha de tecnologias modernas e eficientes:

| Categoria         | Tecnologia         | Descrição                                                              |
| :---------------- | :----------------- | :--------------------------------------------------------------------- |
| **Framework Web** | Streamlit          | Para a construção da interface de usuário interativa e responsiva.     |
| **Dados**         | Pandas             | Essencial para manipulação e análise de dados de planilhas Excel.      |
| **Excel**         | OpenPyXL           | Biblioteca para leitura e escrita de arquivos `.xlsx`.                 |
| **Automação**     | PyWin32            | Integração com o Microsoft Outlook para automação de e-mails (apenas Windows). |
| **Caminhos**      | Pathlib            | Manipulação de caminhos de arquivo de forma orientada a objetos.       |
| **Templates**     | Jinja2             | Motor de templates para renderização dinâmica de e-mails.              |
| **Logging**       | `logging` (Python) | Para registro de eventos e depuração do sistema.                       |

--- 

## 📦 Instalação e Configuração

Para colocar o projeto em funcionamento, siga os passos abaixo:

### Pré-requisitos

-   **Python**: Versão 3.8 ou superior.
-   **Sistema Operacional**: Windows (obrigatório para a integração com o Microsoft Outlook via `PyWin32`).
-   **Microsoft Outlook**: Instalado e configurado no ambiente local.

### Passos de Instalação

1.  **Clone o repositório**: Abra seu terminal ou prompt de comando e execute:

    ```bash
    git clone https://github.com/malikribeiro/RPA-Envio-Emails-STREAMLIT.git
    cd RPA-Envio-Emails-STREAMLIT
    ```

2.  **Crie um ambiente virtual**: É altamente recomendável usar um ambiente virtual para gerenciar as dependências do projeto.

    ```bash
    python -m venv venv
    ```

3.  **Ative o ambiente virtual**:

    -   **Windows (PowerShell)**:
        ```bash
        .\venv\Scripts\Activate.ps1
        ```
    -   **Windows (Command Prompt)**:
        ```bash
        .\venv\Scripts\activate.bat
        ```
    -   **Linux/macOS** (apenas para desenvolvimento, Outlook não será funcional):
        ```bash
        source venv/bin/activate
        ```

4.  **Instale as dependências**: Com o ambiente virtual ativado, instale todas as bibliotecas necessárias:

    ```bash
    pip install -r requirements.txt
    ```

### Estrutura de Arquivos Esperada

O sistema espera uma estrutura de diretórios específica para localizar os arquivos de relatório e contatos. Esta estrutura é baseada no login de rede do usuário e pode ser personalizada em `config.py`.

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
        │   │                   ├── Garantia Financeira/  # PDFs GFN001
        │   │                   ├── Liquidação Financeira/ # PDFs LFN001
        │   │                   ├── Sumário/             # PDFs SUM001
        │   │                   └── ...
        │   └── DGC/
        │       └── Macro/
        │           └── Contatos de E-mail para Macros.xlsx # Planilha de contatos
```

--- 

## 🚀 Execução da Aplicação

Após a instalação, siga estes passos para executar o RPA:

1.  **Ative o ambiente virtual** (se ainda não estiver ativo).

2.  **Execute a aplicação Streamlit**:

    ```bash
    streamlit run app.py
    streamlit run app.py --server.sslCertFile=cert.pem --server.sslKeyFile=key.pem --server.port=8501
    ```

3.  **Acesse no navegador**: A aplicação estará disponível em `http://localhost:8501`.

4.  **Rodar modo fácil**: Inserir no terminal o comando `./run_secure.ps1`.

--- 

## 🖥️ Visão Geral da Interface e Navegação

A interface do usuário foi cuidadosamente projetada para ser clara e eficiente:

-   **Navegação Principal**: Localizada na barra lateral (sidebar), com opções como "Envio de Relatórios" e "Configurações".
-   **Parâmetros de Envio**: Todos os parâmetros essenciais (tipo de relatório, analista, mês, ano) estão centralizados no painel principal para facilitar o acesso.
-   **Pré-visualização de E-mail**: Uma funcionalidade de pré-visualização exibe o e-mail renderizado em HTML antes do envio, permitindo verificações.
-   **Visualização de Dados**: Dados e KPIs são apresentados em um layout limpo e responsivo, otimizado para a visualização.

--- 

## ⚙️ Configurações Avançadas

O projeto oferece opções de configuração para maior flexibilidade:

### Configuração de Relatórios

Cada tipo de relatório pode ser ajustado via interface web na seção "Configurações" ou diretamente no arquivo `config_relatorios.json`. As configurações incluem o nome da aba dos dados (`sheet_dados`), a aba de contatos (`sheet_contatos`), a linha do cabeçalho (`header_row`) e o mapeamento de colunas (`data_columns`).

Exemplo de `config_relatorios.json`:

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

### Templates de E-mail

Os templates de e-mail (assunto, corpo e anexos) são gerenciados via `config/email_templates.json` e podem ser editados através da interface de configurações. O sistema suporta variantes de templates para diferentes cenários, como no caso do relatório LFRES.

### Adicionando Novos Tipos de Relatório

Para estender o sistema com novos tipos de relatório:

1.  **Adicione a configuração** em `config.py` e `config_relatorios.json`.
2.  **Crie um handler** correspondente em `services.py` para definir a lógica de processamento e montagem do e-mail para o novo tipo.
3.  **Atualize `REPORT_HANDLERS`** em `services.py` para incluir o novo handler.

--- 

## 🐛 Tratamento de Erros e Logs

O sistema incorpora um tratamento de erros abrangente para garantir a estabilidade e a confiabilidade:

-   **Verificação de Arquivos**: Validação da existência de arquivos e permissões de acesso.
-   **Validação de Configurações**: Checagem de configurações inválidas ou incompletas.
-   **Tratamento de Dados**: Gerenciamento de dados ausentes ou inconsistentes.
-   **Integração Outlook**: Tratamento de falhas na comunicação com o Microsoft Outlook.

Todos os eventos e erros são registrados em `logs/app.log`, facilitando a depuração e o monitoramento do sistema.

--- 

## 🔒 Segurança

Aspectos de segurança foram considerados no desenvolvimento:

-   **Login de Rede**: Autenticação baseada no usuário de rede para acesso seguro.
-   **Validação de Entrada**: Sanitização e validação de formatos de entrada para prevenir vulnerabilidades.
-   **Caminhos Seguros**: Tratamento seguro de caminhos de arquivo para evitar acessos não autorizados.
-   **Auditoria**: Logs detalhados para fins de auditoria e rastreabilidade.

--- 

## 🤝 Contribuição

Contribuições são bem-vindas! Para contribuir com o projeto:

1.  Faça um fork do repositório.
2.  Crie uma nova branch para sua feature (`git checkout -b feature/minha-nova-feature`).
3.  Implemente suas mudanças e certifique-se de que os testes passem.
4.  Submeta um Pull Request detalhado.

--- 

## 📄 Licença

Este projeto é de uso interno da ELECTRA COMERCIALIZADORA DE ENERGIA S.A.

--- 

## 👥 Autores

-   **Desenvolvido para**: DGCA
-   **Mantido por**: Malik Ribeiro Mourad

--- 

**Versão**: 1.0.0  
**Última atualização**: Outubro 2025
