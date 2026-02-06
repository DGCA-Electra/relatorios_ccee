# Ativa o ambiente virtual (ajuste o caminho se necessário)
Write-Host "Ativando ambiente virtual..."
.\venv\Scripts\Activate.ps1

# Executa o Streamlit com SSL
# Write-Host "Iniciando Streamlit com HTTPS..."
# streamlit run app.py --server.sslCertFile=cert.pem --server.sslKeyFile=key.pem --server.port=8501

# Executa o Streamlit NORMAL (sem SSL local, deixe o túnel cuidar disso)
Write-Host "Iniciando Streamlit..."
streamlit run app.py --server.port=8501