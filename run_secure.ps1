# Ativa o ambiente virtual (ajuste o caminho se necessário)
Write-Host "Ativando ambiente virtual..."
.\venv\Scripts\Activate.ps1

# Executa o Streamlit com SSL
Write-Host "Iniciando Streamlit com HTTPS..."
streamlit run app.py --server.sslCertFile=cert.pem --server.sslKeyFile=key.pem --server.port=8501

# Opcional: Pausa no final se você for dar duplo clique no script
# Read-Host -Prompt "Servidor Streamlit encerrado. Pressione Enter para sair."