# seu_projeto/services.py

import pandas as pd
from pathlib import Path
import sys
import pythoncom
from datetime import timedelta
import re

try:
    import win32com.client as win32
    WIN32_AVAILABLE = sys.platform == "win32"
except ImportError:
    WIN32_AVAILABLE = False

from config import load_configs, MESES

class ReportProcessingError(Exception):
    """Exceção customizada para erros de processamento."""
    pass

# ==============================================================================
# FUNÇÕES AUXILIARES
# ==============================================================================

def _create_outlook_draft(recipient: str, subject: str, body: str, attachments: list):
    """Cria e exibe um rascunho de e-mail no Outlook."""
    if not WIN32_AVAILABLE:
        print("--- MODO DE SIMULAÇÃO ---")
        print(f"PARA: {recipient}")
        print(f"ASSUNTO: {subject}")
        print(f"ANEXOS: {[p.name for p in attachments if p and p.exists()]}")
        print("-------------------------")
        return

    pythoncom.CoInitialize()
    try:
        outlook = win32.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = recipient
        mail.Subject = subject
        mail.HTMLBody = body
        
        for attachment_path in attachments:
            if attachment_path and attachment_path.exists():
                mail.Attachments.Add(str(attachment_path.resolve()))
            else:
                print(f"AVISO: Anexo não encontrado e não será adicionado: {attachment_path}")
        
        mail.Display(True)
        print(f"Rascunho de e-mail para '{recipient}' exibido com sucesso.")
    except Exception as e:
        raise ReportProcessingError(f"Falha ao interagir com o Outlook: {e}")
    finally:
        pythoncom.CoUninitialize()

def _build_filename(company: str, report_type: str, month: str, year: str) -> str:
    """Constrói o nome do arquivo PDF padrão: EMPRESA_TIPO_MES_ANO.pdf"""
    company_clean = company.strip()
    company_part = re.sub(r'[\s_-]+', '_', company_clean).upper()
    
    report_part = report_type.upper()
    month_part = month.lower()[:3]
    year_part = year[-2:]
    
    return f"{company_part}_{report_part}_{month_part}_{year_part}.pdf"

def _format_currency(value) -> str:
    """Formata um valor numérico como moeda brasileira."""
    try:
        val = float(value)
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "R$ 0,00"

def _format_date(date_value) -> str:
    """Formata um valor de data para dd/mm/YYYY, tratando possíveis erros."""
    try:
        return pd.to_datetime(date_value).strftime('%d/%m/%Y')
    except (ValueError, TypeError):
        return "Data Inválida"

# ==============================================================================
# HANDLERS DE E-MAIL (LÓGICA E TEMPLATES)
# ==============================================================================

def handle_gfn001(row: pd.Series, cfg: dict, common: dict):
    if not row.get('Valor', 0) > 0:
        print(f"AVISO GFN001 para '{row['Empresa']}': Valor é zero ou negativo. E-mail não será criado.")
        return None
    
    filename_gfn = _build_filename(row['Empresa'], 'GFN001', common['month'], common['year'])
    filename_sum = _build_filename(row['Empresa'], 'SUM001', common['month'], common['year'])
    
    gfn_path = Path(cfg['pdfs_dir']) / filename_gfn
    
    all_configs = load_configs()
    sum001_cfg = all_configs.get('SUM001')
    if not sum001_cfg:
        print(f"AVISO GFN001: Configuração para 'SUM001' não encontrada. Impossível anexar.")
        return None
        
    sum_path = Path(sum001_cfg['pdfs_dir']) / filename_sum
    attachments = [gfn_path, sum_path]
    
    gfn_exists = gfn_path.exists()
    sum_exists = sum_path.exists()

    if not (gfn_exists and sum_exists):
        print("="*50)
        print(f"AVISO GFN001: Anexos não encontrados para '{row['Empresa']}'.")
        print(f"Nome da Empresa (Original): {row['Empresa']}")
        print(f"Nome do Arquivo GFN Procurado: {filename_gfn}")
        print(f"Caminho Completo GFN: {gfn_path}")
        print(f"Encontrado: {'SIM' if gfn_exists else 'NÃO'}")
        print("-" * 20)
        print(f"Nome do Arquivo SUM Procurado: {filename_sum}")
        print(f"Caminho Completo SUM: {sum_path}")
        print(f"Encontrado: {'SIM' if sum_exists else 'NÃO'}")
        print("="*50)
        return None

    subject = f"GFN001 - Aporte de Garantia Financeira à CCEE - {row['Empresa']} - {common['month_num']}/{common['year']}"
    body = f"""<p>Prezado(a),</p>
    <p>Seguem anexos os relatórios GFN001 e SUM001, que apresenta a Memória de Cálculo de Garantias Financeiras, divulgados pela Câmara de Comercialização de Energia Elétrica - CCEE, com os valores para aporte de Garantias Financeiras referentes à contabilização do mês de {common['month_long']}/{common['year']}.</p>
    <p>A data para realização do aporte é <strong>{common['report_date']}</strong>. Neste dia a CCEE irá verificar se o saldo na sua conta no Departamento de Ações e Custódia (DAWC) do Banco Bradesco comtempla o valor do aporte.</p>
    <p>O saldo necessário em sua conta deverá ser maior ou igual a <strong>{_format_currency(row['Valor'])}</strong>.</p>
    <p>Ressaltamos que os montantes de Garantias Financeiras refletem as premissas previstas na Resolução Normativa ANEEL 957/2021.</p>
    <p>O Quadro 3 - Valor  da Garantia Financeira Avulsa, do GFN001, representa o montante a ser aportado pelo agente na data mencionada, sendo sua composição: ((Total de Garantia Financeira Necessária Preliminar) - (Valor do Ajuste de Garantia Financeira)) * 5%.</p>
    <p>Estamos à disposição para mais informações.</p>"""
    return {'subject': subject, 'body': body, 'attachments': attachments}

def handle_sum001(row: pd.Series, cfg: dict, common: dict):
    valor = row.get('Valor', 0)
    if valor == 0: return None

    data_liquidacao = (_format_date(row['Data']))
    if valor > 0:
        texto1 = "crédito"
        texto2 = "ressaltamos que esse crédito está sujeito ao rateio da eventual inadimplência observada no processo de liquidação financeira da Câmara. Dessa forma, caso o valor não seja creditado na íntegra, o mesmo será incluído no próximo ciclo de contabilização e liquidação financeira, estando o agente sujeito a um novo rateio de inadimplência, conforme Resolução ANEEL nº 552, de 14/10/2002."
    else:
        texto1 = "débito"
        texto2 = "teoricamente a conta possui o saldo necessário, visto que o aporte financeiro foi solicitado anteriormente. No entanto, a fim de evitar qualquer penalidade junto à CCEE, orientamos a verificação do saldo e também que o aporte de qualquer diferença seja efetuado com 1 (um) dia útil de antecedência da data da liquidação financeira."

    filename = _build_filename(row['Empresa'], 'SUM001', common['month'], common['year'])
    attachment = Path(cfg['pdfs_dir']) / filename
    if not attachment.exists(): return None

    subject = f"SUM001 - Sumário da Liquidação Financeira - {row['Empresa']} - {common['month_num']}/{common['year']}"
    body = f"""<p>Prezado(a),</p>
    <p>Segue anexo o relatório SUM001, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente a liquidação financeira de <strong>{common['month_long']}/{common['year']}</strong>. No dia <strong>{data_liquidacao}</strong> há uma previsão de <strong>{texto1}</strong> na sua conta no Departamento de Ações e Custódia (DAC) do Banco Bradesco de <strong>{_format_currency(abs(valor))}</strong>.</p>
    <p>Sendo assim, {texto2}</p>
    <p>Estamos à disposição para mais informações.</p>"""
    return {'subject': subject, 'body': body, 'attachments': [attachment]}

def handle_lfn001(row: pd.Series, cfg: dict, common: dict):
    situacao = row.get('Situacao')
    body = ""
    if situacao == 'Crédito':
        body = f"""<p>Prezado (a),</p>
        <p>Segue anexo o relatório LFN001, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente ao resultado da liquidação financeira de <strong>{common['month_long']}/{common['year']}</strong>. Este relatório demonstra a redução ocorrida no crédito da liquidação financeira decorrente do rateio das inadimplências dos agentes devedores da Câmara.</p>
        <p>Ressaltamos que no próximo ciclo de contabilização e liquidação financeira serão incluídos no resultado do agente todo e qualquer crédito não recebido, estando o agente sujeito a um novo rateio de inadimplência, conforme Resolução ANEEL nº 552, de 14/10/2002.</p>
        <p>Valor a Liquidar do Agente: <strong>{_format_currency(row['ValorLiquidacao'])}</strong>.<br>
        Valor Liquidado do Agente: <strong>{_format_currency(row['ValorLiquidado'])}</strong>.<br>
        Participação do agente no rateio de inadimplências: <strong>{_format_currency(row['ValorInadimplencia'])}</strong>.</p>
        <p>Estamos à disposição para mais informações.</p>"""
    elif situacao == 'Débito':
        body = f"""<p>Prezado (a),</p>
        <p>Segue anexo o relatório LFN001, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente ao resultado da liquidação financeira de <strong>{common['month_long']}/{common['year']}</strong>. Este relatório demonstra o valor debitado na liquidação financeira da CCEE.</p>
        <p>Valor a Liquidar do Agente: <strong>{_format_currency(row['ValorLiquidacao'])}</strong>.<br>
        Valor Liquidado do Agente: <strong>{_format_currency(row['ValorLiquidado'])}</strong>.<br>
        Inadimplência: <strong>{_format_currency(row['ValorInadimplencia'])}</strong>.</p>
        <p>Estamos à disposição para mais informações.</p>"""
    else: return None
    
    filename = _build_filename(row['Empresa'], 'LFN001', common['month'], common['year'])
    attachment = Path(cfg['pdfs_dir']) / filename
    if not attachment.exists(): return None
    
    subject = f"LFN001 - Resultado da Liquidação Financeira - {row['Empresa']} - {common['month_num']}/{common['year']}"
    return {'subject': subject, 'body': body, 'attachments': [attachment]}

def handle_lfres(row: pd.Series, cfg: dict, common: dict):
    valor, data_debito, tipo_agente = row.get('Valor', 0), _format_date(row['Data']), row.get('TipoAgente')
    body = ""
    if valor != 0:
        texto_base = "referente ao pagamento de ressarcimento pela energia contratada não entregue." if tipo_agente == 'Gerador-EER' else f"referente a Liquidação de Energia de Reserva de {common['month_long']}/{common['year']}."
        body = f"""<p>Prezado(a),</p>
        <p>Segue anexo o relatório LFRES0{common['month_num']}, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, {texto_base}</p>
        <p>O valor do Encargo de Energia de Reserva - EER a ser debitado da sua conta no Departamento de Ações e Custódia (DAC) do Banco Bradesco é de <strong>{_format_currency(valor)}</strong> e deverá estar disponível independentemente do valor do Aporte de Garantia Financeira.</p>
        <p>A data do débito será no dia <strong>{data_debito}</strong>. Recomendamos que o valor seja disponibilizado com 1 (um) dia útil de antecedência.</p>
        <p>Estamos à disposição para mais informações.</p>"""
    else:
        if tipo_agente == 'Gerador-EER': return None
        body = f"""<p>Prezado(a),</p>
        <p>Segue anexo o relatório LFRES0{common['month_num']}, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente a Liquidação Financeira de Energia de Reserva de <strong>{common['month_long']}/{common['year']}</strong>.</p>
        <p>Para esse mês os recursos disponíveis na Conta de Energia de Reserva - CONER são suficientes para o pagamento de todas as obrigações vinculadas à energia de reserva, portanto, não será realizada a cobrança do Encargo de Energia de Reserva - EER no dia <strong>{data_debito}</strong>.</p>
        <p>Estamos à disposição para mais informações.</p>"""
    
    filename = _build_filename(row['Empresa'], f"LFRES0{common['month_num']}", common['month'], common['year'])
    attachment = Path(cfg['pdfs_dir']) / filename
    if not attachment.exists(): return None
    
    subject = f"LFRES0{common['month_num']} - Liquidação energia de reserva à CCEE - {row['Empresa']} - {common['month_num']}/{common['year']}"
    return {'subject': subject, 'body': body, 'attachments': [attachment]}

def handle_lembrete(row: pd.Series, cfg: dict, common: dict):
    if not row.get('Valor', 0) > 0: return None
    subject = f"Atenção hoje é o dia do Aporte de Garantia Financeira à CCEE - {row['Empresa']}"
    body = f"""<p>Prezado(a),</p>
    <p>Conforme informado anteriormente, hoje <strong>{common['report_date']}</strong> é a data para o Aporte de Garantia Financeira à CCEE.</p>
    <p>O saldo necessário em sua conta deverá ser maior ou igual a <strong>{_format_currency(row['Valor'])}</strong>.</p>
    <p>A CCEE irá verificar se o saldo na sua conta no Departamento de Ações e Custódia (DAC) do Banco Bradesco contempla o valor do aporte. Os Agentes vendedores que não realizarem o aporte de Garantia Financeira do MCP, estão sujeitos aos ajustes/cortes dos seus contratos de venda na proporção do valor não aportado, essa regra vale inclusive para os Consumidores com Cessão de Contratos. Além dos ajustes dos contratos, poderá ser instaurado o processo de desligamento do agente da CCEE.</p>
    <p>Estamos à disposição para mais informações.</p>"""
    return {'subject': subject, 'body': body, 'attachments': []}

def handle_lfrcap(row: pd.Series, cfg: dict, common: dict):
    data_debito = _format_date(row['Data'])
    filename = _build_filename(row['Empresa'], 'LFRCAP001', common['month'], common['year'])
    attachment = Path(cfg['pdfs_dir']) / filename
    if not attachment.exists(): return None
    
    subject = f"LFRCAP001 - Liquidação de Reserva de Capacidade - {row['Empresa']} - {common['month_num']}/{common['year']}"
    body = f"""<p>Prezado (a),</p>
    <p>Segue anexo o relatório LFRCAP001, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente ao resultado da Liquidação Financeira de Reserva de Capacidade de <strong>{common['month_long']}/{common['year']}</strong>.</p>
    <p>O valor do Encargo de Reserva de Capacidade - ERCAP a ser debitado da sua conta no Departamento de Ações e Custódia (DAC) do Banco Bradesco é de <strong>{_format_currency(row['Valor'])}</strong> e deverá estar disponível independentemente do valor do Aporte de Garantia Financeira.</p>
    <p>A data do débito será no dia <strong> {data_debito}</strong>. Recomendamos que o valor seja disponibilizado com 1 (um) dia útil de antecedência.</p>
    <p>Estamos à disposição para mais informações.</p>"""
    return {'subject': subject, 'body': body, 'attachments': [attachment]}

def handle_rcap(row: pd.Series, cfg: dict, common: dict):
    data_debito = _format_date(row['Data'])
    filename = _build_filename(row['Empresa'], 'RCAP002', common['month'], common['year'])
    attachment = Path(cfg['pdfs_dir']) / filename
    if not attachment.exists(): return None

    subject = f"RCAP002 - Reserva de Capacidade - {row['Empresa']} - {common['month_num']}/{common['year']}"
    body = f"""<p>Prezado (a),</p>
    <p>Segue anexo o relatório RCAP002, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente ao resultado da Reserva de Capacidade de <strong>{common['month_long']}/{common['year']}</strong>.</p>
    <p>O valor do Encargo de Reserva de Capacidade - ERCAP a ser debitado da sua conta no Departamento de Ações e Custódia (DAC) do Banco Bradesco é de <strong>{_format_currency(row['Valor'])}</strong> e deverá estar disponível independentemente do valor do Aporte de Garantia Financeira.</p>
    <p>A data do débito será no dia <strong>{data_debito}</strong>. Recomendamos que o valor seja disponibilizado com 1 (um) dia útil de antecedência.</p>
    <p>Estamos à disposição para mais informações.</p>"""
    return {'subject': subject, 'body': body, 'attachments': [attachment]}

REPORT_HANDLERS = {
    'GFN001': handle_gfn001, 'SUM001': handle_sum001, 'LFN001': handle_lfn001,
    'LFRES': handle_lfres, 'GFN - LEMBRETE': handle_lembrete, 'LFRCAP': handle_lfrcap, 'RCAP': handle_rcap
}

# ==============================================================================
# FUNÇÃO PRINCIPAL DE PROCESSAMENTO
# ==============================================================================

def process_reports(report_type: str, analyst: str, month: str, year: str) -> list:
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    if not cfg: raise ReportProcessingError(f"'{report_type}' não encontrado nas configs.")

    try:
        header = int(cfg.get('header_row', 0))
        df_dados = pd.read_excel(Path(cfg['excel_dados']), sheet_name=cfg['sheet_dados'], header=header)
        df_contatos = pd.read_excel(Path(cfg['excel_contatos']), sheet_name=cfg['sheet_contatos'])
    except Exception as e:
        raise ReportProcessingError(f"Erro ao ler as planilhas. Verifique os caminhos, nomes das abas e linha de cabeçalho. Erro: {e}")

    try:
        column_mapping = dict(item.split(':') for item in cfg['data_columns'].split(','))
    except ValueError:
        raise ReportProcessingError(f"Formato inválido em 'Mapeamento de Colunas' para {report_type}. Use 'NomeNoExcel:NomePadrão,OutraColuna:OutroNome'.")
        
    df_dados.rename(columns=column_mapping, inplace=True)
    df_contatos.rename(columns={'AGENTE': 'Empresa', 'ANALISTA': 'Analista', 'E-MAILS RELATÓRIOS CCEE': 'Email'}, inplace=True)
    
    if 'Empresa' not in df_dados.columns: raise ReportProcessingError(f"Coluna 'Empresa' não encontrada no arquivo de dados de {report_type} após mapeamento.")
    if 'Empresa' not in df_contatos.columns: raise ReportProcessingError("Coluna 'AGENTE' não encontrada no arquivo de contatos.")
    if 'Analista' not in df_contatos.columns: raise ReportProcessingError("Coluna 'ANALISTA' não encontrada no arquivo de contatos.")

    df_merged = pd.merge(df_dados, df_contatos, on='Empresa', how='left')
    df_filtered = df_merged[df_merged['Analista'] == analyst].copy()
    if df_filtered.empty: raise ReportProcessingError(f"Nenhum registro para o analista '{analyst}'.")
    
    df_filtered['Email'] = df_filtered['Email'].fillna('EMAIL_NAO_ENCONTRADO')

    meses_map = {m.upper(): f"{i+1:02d}" for i, m in enumerate(MESES)}
    report_date = ""
    if report_type in ['GFN001', 'GFN - LEMBRETE']:
        try:
            df_data_raw = pd.read_excel(Path(cfg['excel_dados']), sheet_name=cfg['sheet_dados'], header=None, usecols=[0])
            report_date = _format_date(str(df_data_raw.iloc[23, 0]).replace("Data do Aporte de Garantias:", "").strip())
        except Exception:
            report_date = "Data não encontrada"

    common_data = {
        'analyst': analyst,
        'month_long': month.title(),
        'month_num': meses_map.get(month.upper(), '00'),
        'month': month,
        'year': year,
        'year_2_digits': year[-2:],
        'report_date': report_date
    }

    results, created_count = [], 0
    for _, row in df_filtered.iterrows():
        handler = REPORT_HANDLERS.get(report_type)
        email_data, anexos = None, 0
        if handler: email_data = handler(row, cfg, common_data)
        
        if email_data:
            created_count += 1
            assinatura = f"<br><br><p>Atenciosamente,</p><p><strong>{analyst}</strong></p>"
            email_data['body'] += assinatura
            _create_outlook_draft(row['Email'], **email_data)
            anexos = len(email_data.get('attachments', []))
        
        results.append({
            'empresa': row['Empresa'],
            'data': report_date or _format_date(row.get('Data')),
            'valor': _format_currency(row.get('Valor')),
            'email': row['Email'], 'anexos_count': anexos, 'created_count': created_count,
        })
    return results