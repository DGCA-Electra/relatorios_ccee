# seu_projeto/services.py

import pandas as pd
from pathlib import Path
import sys
import pythoncom
from datetime import timedelta, datetime
import re
import os
from typing import Dict, List, Any, Optional, Tuple

try:
    import win32com.client as win32
    WIN32_AVAILABLE = sys.platform == "win32"
except ImportError:
    WIN32_AVAILABLE = False

from config import load_configs, MESES, build_report_paths, get_user_paths

class ReportProcessingError(Exception):
    """Exceção customizada para erros de processamento."""
    pass

class PathNotFoundError(Exception):
    """Exceção para quando um caminho não é encontrado."""
    pass

class ConfigurationError(Exception):
    """Exceção para erros de configuração."""
    pass

# ==============================================================================
# FUNÇÕES AUXILIARES
# ==============================================================================

def _create_outlook_draft(recipient: str, subject: str, body: str, attachments: List[Path]) -> None:
    """
    Cria e exibe um rascunho de e-mail no Outlook.
    
    Args:
        recipient: Destinatário do e-mail
        subject: Assunto do e-mail
        body: Corpo do e-mail em HTML
        attachments: Lista de caminhos para anexos
    """
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
    """
    Constrói o nome do arquivo PDF padrão: EMPRESA_TIPO_MES_ANO.pdf
    
    Args:
        company: Nome da empresa
        report_type: Tipo do relatório
        month: Mês
        year: Ano
        
    Returns:
        Nome do arquivo PDF
    """
    company_clean = company.strip()
    company_part = re.sub(r'[\s_-]+', '_', company_clean).upper()
    
    report_part = report_type.upper()
    month_part = month.lower()[:3]
    year_part = year[-2:]
    
    return f"{company_part}_{report_part}_{month_part}_{year_part}.pdf"

def _format_currency(value: Any) -> str:
    """
    Formata um valor numérico como moeda brasileira.
    
    Args:
        value: Valor a ser formatado
        
    Returns:
        String formatada como moeda brasileira
    """
    try:
        val = float(value)
        return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "R$ 0,00"

def _format_date(date_value: Any) -> str:
    """
    Formata um valor de data para dd/mm/YYYY, tratando possíveis erros.
    
    Args:
        date_value: Valor de data a ser formatado
        
    Returns:
        String formatada da data ou mensagem de erro
    """
    try:
        if date_value is None or pd.isna(date_value):
            return "Data não informada"
        return pd.to_datetime(date_value).strftime('%d/%m/%Y')
    except (ValueError, TypeError):
        return "Data Inválida"

def _create_warning_html(warnings: List[str]) -> str:
    """
    Cria um bloco de alerta HTML a partir de uma lista de avisos.
    
    Args:
        warnings: Lista de avisos
        
    Returns:
        HTML formatado dos avisos
    """
    if not warnings:
        return ""
    
    items_html = "".join(f"<li>{w}</li>" for w in warnings)
    return f"""
    <p style='background-color: #FFF3CD; border-left: 6px solid #FFC107; color: #856404; padding: 10px; margin-bottom: 15px;'>
    <strong>ATENÇÃO:</strong> Foram encontrados os seguintes problemas: {items_html}
    </p>
    """

def _load_excel_data(excel_path: str, sheet_name: str, header_row: int) -> pd.DataFrame:
    """
    Carrega dados de uma planilha Excel com tratamento de erros específico.
    
    Args:
        excel_path: Caminho para o arquivo Excel
        sheet_name: Nome da aba
        header_row: Linha do cabeçalho
        
    Returns:
        DataFrame com os dados carregados
        
    Raises:
        FileNotFoundError: Se o arquivo não for encontrado
        ValueError: Se a aba não for encontrada
        ReportProcessingError: Para outros erros de leitura
    """
    try:
        if not Path(excel_path).exists():
            raise FileNotFoundError(f"Arquivo não encontrado: {excel_path}")
        
        df = pd.read_excel(Path(excel_path), sheet_name=sheet_name, header=header_row)
        return df
    except FileNotFoundError:
        raise
    except ValueError as e:
        if "sheet" in str(e).lower():
            raise ValueError(f"Aba '{sheet_name}' não encontrada no arquivo {excel_path}")
        raise
    except Exception as e:
        raise ReportProcessingError(f"Erro ao ler planilha {excel_path}: {e}")

def _find_attachment(pdf_dir: str, filename: str, alternative_paths: Optional[List[str]] = None) -> Tuple[Optional[Path], List[str]]:
    """
    Busca um anexo em múltiplos caminhos possíveis.
    
    Args:
        pdf_dir: Diretório principal para busca
        filename: Nome do arquivo
        alternative_paths: Lista de caminhos alternativos
        
    Returns:
        Tupla com (caminho encontrado, lista de avisos)
    """
    warnings = []
    attachment_path = Path(pdf_dir) / filename
    
    if attachment_path.exists():
        return attachment_path, warnings
    
    warnings.append(f"Anexo não encontrado no caminho principal: {attachment_path}")
    
    # Tentar caminhos alternativos
    if alternative_paths:
        for alt_path in alternative_paths:
            alt_attachment = Path(alt_path) / filename
            if alt_attachment.exists():
                warnings.append(f"Anexo encontrado em caminho alternativo: {alt_attachment}")
                return alt_attachment, warnings
            else:
                warnings.append(f"Anexo não encontrado em caminho alternativo: {alt_attachment}")
    
    return None, warnings

# ==============================================================================
# HANDLERS DE E-MAIL (LÓGICA E TEMPLATES COMPLETOS)
# ==============================================================================

def handle_gfn001(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any], all_configs: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handler para relatórios GFN001.
    
    Args:
        row: Linha de dados da empresa
        cfg: Configuração do relatório
        common: Dados comuns (mês, ano, etc.)
        all_configs: Todas as configurações carregadas
        
    Returns:
        Dicionário com dados do e-mail
    """
    warnings = []
    if not row.get('Valor', 0) > 0:
        warnings.append("O valor do aporte é zero ou negativo. Verifique a planilha de dados.")
    
    filename_gfn = _build_filename(row['Empresa'], 'GFN001', common['month'], common['year'])
    filename_sum = _build_filename(row['Empresa'], 'SUM001', common['month'], common['year'])
    
    # Buscar anexo GFN001
    gfn_path, gfn_warnings = _find_attachment(cfg['pdfs_dir'], filename_gfn)
    warnings.extend(gfn_warnings)
    
    # Buscar anexo SUM001 (memória de cálculo)
    sum001_cfg = all_configs.get('SUM001')
    sum_path = None
    sum_warnings = []
    
    if sum001_cfg and 'pdfs_dir' in sum001_cfg:
        sum_path, sum_warnings = _find_attachment(
            sum001_cfg['pdfs_dir'], 
            filename_sum,
            # Caminho alternativo para memória de cálculo
            [f"{cfg['pdfs_dir'].replace('SUM001', 'SUM001 - Memória_de_Cálculo')}"]
        )
        warnings.extend(sum_warnings)
    else:
        warnings.append("Configuração do SUM001 não encontrada para anexar PDF da memória de cálculo")
    
    # Montar lista de anexos
    attachments = []
    if gfn_path:
        attachments.append(gfn_path)
    if sum_path:
        attachments.append(sum_path)

    warning_html = _create_warning_html(warnings)

    subject = f"GFN001 - Aporte de Garantia Financeira à CCEE - {row['Empresa']} - {common['month_num']}/{common['year']}"
    body = f"""
    {warning_html}
    <p>Prezado(a),</p>
    <p>Seguem anexos os relatórios GFN001 e SUM001, que apresenta a Memória de Cálculo de Garantias Financeiras, divulgados pela Câmara de Comercialização de Energia Elétrica - CCEE, com os valores para aporte de Garantias Financeiras referentes à contabilização do mês de {common['month_long']}/{common['year']}.</p>
    <p>A data para realização do aporte é <strong>{common['report_date']}</strong>. Neste dia a CCEE irá verificar se o saldo na sua conta no Departamento de Ações e Custódia (DAWC) do Banco Bradesco comtempla o valor do aporte.</p>
    <p>O saldo necessário em sua conta deverá ser maior ou igual a <strong>{_format_currency(row['Valor'])}</strong>.</p>
    <p>Ressaltamos que os montantes de Garantias Financeiras refletem as premissas previstas na Resolução Normativa ANEEL 957/2021.</p>
    <p>O Quadro 3 - Valor  da Garantia Financeira Avulsa, do GFN001, representa o montante a ser aportado pelo agente na data mencionada, sendo sua composição: ((Total de Garantia Financeira Necessária Preliminar) - (Valor do Ajuste de Garantia Financeira)) * 5%.</p>
    <p>Estamos à disposição para mais informações.</p>
    """
    return {'subject': subject, 'body': body, 'attachments': attachments}

def handle_sum001(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handler para relatórios SUM001.
    
    Args:
        row: Linha de dados da empresa
        cfg: Configuração do relatório
        common: Dados comuns (mês, ano, etc.)
        
    Returns:
        Dicionário com dados do e-mail
    """
    warnings = []
    valor = row.get('Valor', 0)
    situacao = row.get('Situacao', '')
    
    if valor == 0:
        warnings.append("O valor a liquidar é zero. O e-mail foi gerado para conferência.")

    filename = _build_filename(row['Empresa'], 'SUM001', common['month'], common['year'])
    attachment, attachment_warnings = _find_attachment(cfg['pdfs_dir'], filename)
    warnings.extend(attachment_warnings)

    warning_html = _create_warning_html(warnings)
    
    # Extrair datas do Quadro 1 (linha 24, colunas A e B)
    try:
        df_raw = pd.read_excel(Path(cfg['excel_dados']), sheet_name=cfg['sheet_dados'], header=None)
        data_debito_quadro1 = df_raw.iloc[23, 0]  # Linha 24, Coluna A (índice 23, 0)
        data_credito_quadro1 = df_raw.iloc[23, 1]  # Linha 24, Coluna B (índice 23, 1)
    except Exception as e:
        warnings.append(f"Erro ao extrair datas do Quadro 1: {e}")
        data_debito_quadro1 = None
        data_credito_quadro1 = None
    
    # Determinar se é débito ou crédito baseado na coluna Situacao
    if situacao == 'Crédito':
        texto1 = "crédito"
        # Usar Data do Crédito do Quadro 1
        if data_credito_quadro1 is not None and not pd.isna(data_credito_quadro1):
            data_liquidacao = _format_date(data_credito_quadro1)
        else:
            # Fallback: usar data atual
            data_liquidacao = datetime.now().strftime('%d/%m/%Y')
        texto2 = "ressaltamos que esse crédito está sujeito ao rateio da eventual inadimplência observada no processo de liquidação financeira da Câmara. Dessa forma, caso o valor não seja creditado na íntegra, o mesmo será incluído no próximo ciclo de contabilização e liquidação financeira, estando o agente sujeito a um novo rateio de inadimplência, conforme Resolução ANEEL nº 552, de 14/10/2002."
    elif situacao == 'Débito':
        texto1 = "débito"
        # Usar Data do Débito do Quadro 1
        if data_debito_quadro1 is not None and not pd.isna(data_debito_quadro1):
            data_liquidacao = _format_date(data_debito_quadro1)
        else:
            # Fallback: usar data atual
            data_liquidacao = datetime.now().strftime('%d/%m/%Y')
        texto2 = "teoricamente a conta possui o saldo necessário, visto que o aporte financeiro foi solicitado anteriormente. No entanto, a fim de evitar qualquer penalidade junto à CCEE, orientamos a verificação do saldo e também que o aporte de qualquer diferença seja efetuado com 1 (um) dia útil de antecedência da data da liquidação financeira."
    else:
        warnings.append(f"Situação '{situacao}' não reconhecida. Usando data atual como fallback.")
        data_liquidacao = datetime.now().strftime('%d/%m/%Y')
        texto1 = "transação"
        texto2 = "verifique os dados na planilha."

    subject = f"SUM001 - Sumário da Liquidação Financeira - {row['Empresa']} - {common['month_num']}/{common['year']}"
    body = f"""{warning_html}
    <p>Prezado(a),</p>
    <p>Segue anexo o relatório SUM001, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente a liquidação financeira de <strong>{common['month_long']}/{common['year']}</strong>. No dia <strong>{data_liquidacao}</strong> há uma previsão de <strong>{texto1}</strong> na sua conta no Departamento de Ações e Custódia (DAC) do Banco Bradesco de <strong>{_format_currency(abs(valor))}</strong>.</p>
    <p>Sendo assim, {texto2}</p>
    <p>Estamos à disposição para mais informações.</p>"""
    
    return {'subject': subject, 'body': body, 'attachments': [attachment] if attachment else []}

def handle_lfn001(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handler para relatórios LFN001.
    
    Args:
        row: Linha de dados da empresa
        cfg: Configuração do relatório
        common: Dados comuns (mês, ano, etc.)
        
    Returns:
        Dicionário com dados do e-mail
    """
    warnings = []
    situacao = row.get('Situacao')
    # Tratamento dos valores para garantir que não sejam nan/None
    valor_liquidacao = row.get('ValorLiquidacao', 0)
    valor_liquidado = row.get('ValorLiquidado', 0)
    valor_inadimplencia = row.get('ValorInadimplencia', 0)
    def safe_val(val):
        if pd.isna(val) or val in [None, 0, 0.0, "0", "0.0", "nan", "None"]:
            return 0
        return val
    valor_liquidacao = safe_val(valor_liquidacao)
    valor_liquidado = safe_val(valor_liquidado)
    valor_inadimplencia = safe_val(valor_inadimplencia)

    if situacao == 'Crédito':
        body = f"""
        <p>Prezado (a),</p>
        <p>Segue anexo o relatório LFN001, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente ao resultado da liquidação financeira de <strong>{common['month_long']}/{common['year']}</strong>. Este relatório demonstra a redução ocorrida no crédito da liquidação financeira decorrente do rateio das inadimplências dos agentes devedores da Câmara.</p>
        <p>Ressaltamos que no próximo ciclo de contabilização e liquidação financeira serão incluídos no resultado do agente todo e qualquer crédito não recebido, estando o agente sujeito a um novo rateio de inadimplência, conforme Resolução ANEEL nº 552, de 14/10/2002.</p>
        <p>Valor a Liquidar do Agente: <strong>{_format_currency(valor_liquidacao)}</strong>.<br>
        Valor Liquidado do Agente: <strong>{_format_currency(valor_liquidado)}</strong>.<br>
        Participação do agente no rateio de inadimplências: <strong>{_format_currency(valor_inadimplencia)}</strong>.</p>
        <p>Estamos à disposição para mais informações.</p>"""
    elif situacao == 'Débito':
        body = f"""
        <p>Prezado (a),</p>
        <p>Segue anexo o relatório LFN001, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente ao resultado da liquidação financeira de <strong>{common['month_long']}/{common['year']}</strong>. Este relatório demonstra o valor debitado na liquidação financeira da CCEE.</p>
        <p>Valor a Liquidar do Agente: <strong>{_format_currency(valor_liquidacao)}</strong>.<br>
        Valor Liquidado do Agente: <strong>{_format_currency(valor_liquidado)}</strong>.<br>
        Inadimplência: <strong>{_format_currency(valor_inadimplencia)}</strong>.</p>
        <p>Estamos à disposição para mais informações.</p>"""
    else:
        warnings.append(f"A situação ('{situacao}') não é 'Crédito' ou 'Débito'. Verifique a planilha.")
        body = "<p>Não foi possível gerar o corpo do e-mail. Verifique a coluna 'Situação' na planilha de dados.</p>"

    filename = _build_filename(row['Empresa'], 'LFN001', common['month'], common['year'])
    attachment, attachment_warnings = _find_attachment(cfg['pdfs_dir'], filename)
    warnings.extend(attachment_warnings)
    
    warning_html = _create_warning_html(warnings)
    
    subject = f"LFN001 - Resultado da Liquidação Financeira - {row['Empresa']} - {common['month_num']}/{common['year']}"
    return {'subject': subject, 'body': warning_html + body, 'attachments': [attachment] if attachment else []}

def handle_lfres(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handler para relatórios LFRES.
    
    Args:
        row: Linha de dados da empresa
        cfg: Configuração do relatório
        common: Dados comuns (mês, ano, etc.)
        
    Returns:
        Dicionário com dados do e-mail
    """
    warnings = []
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
        if tipo_agente == 'Gerador-EER':
             warnings.append("O valor para 'Gerador-EER' é zero. O e-mail foi gerado para conferência, mas normalmente não seria enviado.")
        body = f"""<p>Prezado(a),</p>
        <p>Segue anexo o relatório LFRES0{common['month_num']}, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente a Liquidação Financeira de Energia de Reserva de <strong>{common['month_long']}/{common['year']}</strong>.</p>
        <p>Para esse mês os recursos disponíveis na Conta de Energia de Reserva - CONER são suficientes para o pagamento de todas as obrigações vinculadas à energia de reserva, portanto, não será realizada a cobrança do Encargo de Energia de Reserva - EER no dia <strong>{data_debito}</strong>.</p>
        <p>Estamos à disposição para mais informações.</p>"""
    
    filename = _build_filename(row['Empresa'], f"LFRES0{common['month_num']}", common['month'], common['year'])
    attachment, attachment_warnings = _find_attachment(cfg['pdfs_dir'], filename)
    warnings.extend(attachment_warnings)
    
    warning_html = _create_warning_html(warnings)
    
    subject = f"LFRES0{common['month_num']} - Liquidação energia de reserva à CCEE - {row['Empresa']} - {common['month_num']}/{common['year']}"
    return {'subject': subject, 'body': warning_html + body, 'attachments': [attachment] if attachment else []}

def handle_lembrete(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """
    Handler para lembretes.
    
    Args:
        row: Linha de dados da empresa
        cfg: Configuração do relatório
        common: Dados comuns (mês, ano, etc.)
        
    Returns:
        Dicionário com dados do e-mail ou None se não deve ser enviado
    """
    if not row.get('Valor', 0) > 0:
        print(f"AVISO LEMBRETE para '{row['Empresa']}': Valor é zero ou negativo. E-mail NÃO será criado.")
        return None
        
    subject = f"Atenção hoje é o dia do Aporte de Garantia Financeira à CCEE - {row['Empresa']}"
    body = f"""<p>Prezado(a),</p>
    <p>Conforme informado anteriormente, hoje <strong>{common['report_date']}</strong> é a data para o Aporte de Garantia Financeira à CCEE.</p>
    <p>O saldo necessário em sua conta deverá ser maior ou igual a <strong>{_format_currency(row['Valor'])}</strong>.</p>
    <p>A CCEE irá verificar se o saldo na sua conta no Departamento de Ações e Custódia (DAC) do Banco Bradesco contempla o valor do aporte. Os Agentes vendedores que não realizarem o aporte de Garantia Financeira do MCP, estão sujeitos aos ajustes/cortes dos seus contratos de venda na proporção do valor não aportado, essa regra vale inclusive para os Consumidores com Cessão de Contratos. Além dos ajustes dos contratos, poderá ser instaurado o processo de desligamento do agente da CCEE.</p>
    <p>Estamos à disposição para mais informações.</p>"""
    return {'subject': subject, 'body': body, 'attachments': []}

def handle_lfrcap(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handler para relatórios LFRCAP.
    
    Args:
        row: Linha de dados da empresa
        cfg: Configuração do relatório
        common: Dados comuns (mês, ano, etc.)
        
    Returns:
        Dicionário com dados do e-mail
    """
    warnings = []
    data_debito = _format_date(row['Data'])
    filename = _build_filename(row['Empresa'], 'LFRCAP001', common['month'], common['year'])
    attachment, attachment_warnings = _find_attachment(cfg['pdfs_dir'], filename)
    warnings.extend(attachment_warnings)
    
    warning_html = _create_warning_html(warnings)

    subject = f"LFRCAP001 - Liquidação de Reserva de Capacidade - {row['Empresa']} - {common['month_num']}/{common['year']}"
    body = f"""{warning_html}
    <p>Prezado (a),</p>
    <p>Segue anexo o relatório LFRCAP001, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente ao resultado da Liquidação Financeira de Reserva de Capacidade de <strong>{common['month_long']}/{common['year']}</strong>.</p>
    <p>O valor do Encargo de Reserva de Capacidade - ERCAP a ser debitado da sua conta no Departamento de Ações e Custódia (DAC) do Banco Bradesco é de <strong>{_format_currency(row['Valor'])}</strong> e deverá estar disponível independentemente do valor do Aporte de Garantia Financeira.</p>
    <p>A data do débito será no dia <strong> {data_debito}</strong>. Recomendamos que o valor seja disponibilizado com 1 (um) dia útil de antecedência.</p>
    <p>Estamos à disposição para mais informações.</p>"""
    return {'subject': subject, 'body': body, 'attachments': [attachment] if attachment else []}

def handle_rcap(row: pd.Series, cfg: Dict[str, Any], common: Dict[str, Any]) -> Dict[str, Any]:
    """
    Handler para relatórios RCAP.
    
    Args:
        row: Linha de dados da empresa
        cfg: Configuração do relatório
        common: Dados comuns (mês, ano, etc.)
        
    Returns:
        Dicionário com dados do e-mail
    """
    warnings = []
    data_debito = _format_date(row['Data'])
    filename = _build_filename(row['Empresa'], 'RCAP002', common['month'], common['year'])
    attachment, attachment_warnings = _find_attachment(cfg['pdfs_dir'], filename)
    warnings.extend(attachment_warnings)

    warning_html = _create_warning_html(warnings)

    subject = f"RCAP002 - Reserva de Capacidade - {row['Empresa']} - {common['month_num']}/{common['year']}"
    body = f"""{warning_html}
    <p>Prezado (a),</p>
    <p>Segue anexo o relatório RCAP002, divulgado pela Câmara de Comercialização de Energia Elétrica - CCEE, referente ao resultado da Reserva de Capacidade de <strong>{common['month_long']}/{common['year']}</strong>.</p>
    <p>O valor do Encargo de Reserva de Capacidade - ERCAP a ser debitado da sua conta no Departamento de Ações e Custódia (DAC) do Banco Bradesco é de <strong>{_format_currency(row['Valor'])}</strong> e deverá estar disponível independentemente do valor do Aporte de Garantia Financeira.</p>
    <p>A data do débito será no dia <strong>{data_debito}</strong>. Recomendamos que o valor seja disponibilizado com 1 (um) dia útil de antecedência.</p>
    <p>Estamos à disposição para mais informações.</p>"""
    return {'subject': subject, 'body': body, 'attachments': [attachment] if attachment else []}

# Mapeamento de handlers
REPORT_HANDLERS = {
    'GFN001': handle_gfn001, 
    'SUM001': handle_sum001, 
    'LFN001': handle_lfn001,
    'LFRES': handle_lfres, 
    'LEMBRETE': handle_lembrete, 
    'LFRCAP': handle_lfrcap, 
    'RCAP': handle_rcap
}

# ==============================================================================
# FUNÇÃO PRINCIPAL DE PROCESSAMENTO
# ==============================================================================

def _calculate_sum001_dates(cfg: Dict[str, Any], situacao: str) -> str:
    """
    Calcula a data de débito ou crédito para SUM001 baseado na situação.
    
    Args:
        cfg: Configuração do relatório
        situacao: Situação (Débito/Crédito)
        
    Returns:
        Data formatada ou mensagem de erro
    """
    try:
        df_raw = pd.read_excel(Path(cfg['excel_dados']), sheet_name=cfg['sheet_dados'], header=None)
        # Linha 24: Data do Débito (Coluna A) e Data do Crédito (Coluna B)
        data_debito_quadro1 = df_raw.iloc[23, 0]  # Linha 24, Coluna A (índice 23, 0)
        data_credito_quadro1 = df_raw.iloc[23, 1]  # Linha 24, Coluna B (índice 23, 1)
        
        if situacao == 'Crédito':
            if data_credito_quadro1 is not None and not pd.isna(data_credito_quadro1):
                return _format_date(data_credito_quadro1)
        elif situacao == 'Débito':
            if data_debito_quadro1 is not None and not pd.isna(data_debito_quadro1):
                return _format_date(data_debito_quadro1)
        
        # Fallback: usar data atual
        return datetime.now().strftime('%d/%m/%Y')
    except Exception as e:
        return f"Erro ao calcular data: {e}"

def _load_and_process_data(cfg: Dict[str, Any], login_usuario: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Carrega e processa dados das planilhas.
    
    Args:
        cfg: Configuração do relatório
        login_usuario: Login do usuário
        
    Returns:
        Tupla com (DataFrame de dados, DataFrame de contatos)
        
    Raises:
        ReportProcessingError: Se houver erro no processamento
    """
    try:
        header = int(cfg.get('header_row', 0))
        df_dados = _load_excel_data(cfg['excel_dados'], cfg['sheet_dados'], header)
        df_contatos = _load_excel_data(cfg['excel_contatos'], cfg['sheet_contatos'], 0)
    except (FileNotFoundError, ValueError) as e:
        raise ReportProcessingError(f"Erro ao ler as planilhas: {e}")

    try:
        column_mapping = dict(item.split(':') for item in cfg['data_columns'].split(','))
    except ValueError:
        raise ReportProcessingError(f"Formato inválido em 'Mapeamento de Colunas'. Use 'NomeNoExcel:NomePadrão,OutraColuna:OutroNome'.")
        
    df_dados.rename(columns=column_mapping, inplace=True)
    df_contatos.rename(columns={'AGENTE': 'Empresa', 'ANALISTA': 'Analista', 'E-MAILS RELATÓRIOS CCEE': 'Email'}, inplace=True)
    
    if 'Empresa' not in df_dados.columns: 
        raise ReportProcessingError(f"Coluna 'Empresa' não encontrada nos dados após mapeamento.")
    if 'Empresa' not in df_contatos.columns: 
        raise ReportProcessingError("Coluna 'AGENTE' não encontrada nos contatos.")
    if 'Analista' not in df_contatos.columns: 
        raise ReportProcessingError("Coluna 'ANALISTA' não encontrada nos contatos.")

    return df_dados, df_contatos

def process_reports(report_type: str, analyst: str, month: str, year: str, login_usuario: str) -> List[Dict[str, Any]]:
    """
    Processa relatórios e gera e-mails.
    
    Args:
        report_type: Tipo do relatório
        analyst: Analista responsável
        month: Mês do relatório
        year: Ano do relatório
        login_usuario: Login do usuário
        
    Returns:
        Lista com resultados do processamento
        
    Raises:
        ReportProcessingError: Se houver erro no processamento
    """
    if not login_usuario:
        raise ReportProcessingError("Login do usuário não informado.")
    
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    if not cfg: 
        raise ReportProcessingError(f"'{report_type}' não encontrado nas configurações.")

    # Construir caminhos dinamicamente
    try:
        report_paths = build_report_paths(report_type, year, month, login_usuario)
        cfg.update(report_paths)
    except ValueError as e:
        raise ReportProcessingError(f"Erro ao construir caminhos: {e}")

    # Carregar e processar dados
    df_dados, df_contatos = _load_and_process_data(cfg, login_usuario)
    
    # Filtrar dados por analista
    df_merged = pd.merge(df_dados, df_contatos, on='Empresa', how='left')
    df_filtered = df_merged[df_merged['Analista'] == analyst].copy()
    if df_filtered.empty: 
        raise ReportProcessingError(f"Nenhum registro encontrado para o analista '{analyst}'.")
    
    df_filtered['Email'] = df_filtered['Email'].fillna('EMAIL_NAO_ENCONTRADO')

    # Preparar dados comuns
    meses_map = {m.upper(): f"{i+1:02d}" for i, m in enumerate(MESES)}
    report_date = ""
    
    if report_type in ['GFN001', 'LEMBRETE']:
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

    # Processar cada linha e gerar e-mails
    results, created_count = [], 0
    handler = REPORT_HANDLERS.get(report_type)
    
    if not handler:
        raise ReportProcessingError(f"Handler não encontrado para o tipo de relatório '{report_type}'")
    
    for _, row in df_filtered.iterrows():
        email_data = None
        anexos = 0
        
        try:
            if report_type == 'GFN001':
                email_data = handler(row, cfg, common_data, all_configs)
            else:
                email_data = handler(row, cfg, common_data)
        except Exception as e:
            print(f"Erro ao processar linha para {row.get('Empresa', 'Empresa desconhecida')}: {e}")
            continue
        
        if email_data:
            created_count += 1
            assinatura = f"<br><br><p>Atenciosamente,</p><p><strong>{analyst}</strong></p>"
            email_data['body'] += assinatura
            _create_outlook_draft(row['Email'], **email_data)
            anexos = sum(1 for p in email_data.get('attachments', []) if p and p.exists())
        
        results.append({
            'empresa': row['Empresa'],
            'data': report_date or _format_date(row.get('Data')),
            'valor': _format_currency(row.get('Valor')),
            'email': row['Email'], 
            'anexos_count': anexos, 
            'created_count': created_count,
        })
    
    return results

def preview_dados(report_type: str, analyst: str, month: str, year: str, login_usuario: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Pré-visualiza dados de um relatório.
    
    Args:
        report_type: Tipo do relatório
        analyst: Analista responsável
        month: Mês do relatório
        year: Ano do relatório
        login_usuario: Login do usuário
        
    Returns:
        Tupla com (DataFrame completo, DataFrame de preview)
        
    Raises:
        ReportProcessingError: Se houver erro no processamento
    """
    if not login_usuario:
        raise ReportProcessingError("Login do usuário não informado.")
    
    all_configs = load_configs()
    cfg = all_configs.get(report_type)
    if not cfg:
        raise ReportProcessingError(f"'{report_type}' não encontrado nas configurações.")
    
    # Construir caminhos dinamicamente
    try:
        report_paths = build_report_paths(report_type, year, month, login_usuario)
        cfg.update(report_paths)
    except ValueError as e:
        raise ReportProcessingError(f"Erro ao construir caminhos: {e}")

    # Carregar e processar dados
    df_dados, df_contatos = _load_and_process_data(cfg, login_usuario)
    
    # Filtrar dados por analista
    df_merged = pd.merge(df_dados, df_contatos, on='Empresa', how='left')
    df_filtered = df_merged[df_merged['Analista'] == analyst].copy()
    if df_filtered.empty:
        raise ReportProcessingError(f"Nenhum registro encontrado para o analista '{analyst}'.")
    
    df_filtered['Email'] = df_filtered['Email'].fillna('EMAIL_NAO_ENCONTRADO')
    
    # Para SUM001, calcular as datas de débito/crédito
    if report_type == 'SUM001':
        df_filtered['Data_Debito_Credito'] = df_filtered.apply(
            lambda row: _calculate_sum001_dates(cfg, row.get('Situacao', '')),
            axis=1
        )
    
    preview_df = df_filtered.head(20)  # Mostra as 20 primeiras linhas para preview
    
    return df_filtered, preview_df