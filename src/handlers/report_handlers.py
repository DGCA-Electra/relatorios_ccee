import pandas as pd
import logging
from src.utils.file_utils import load_excel_data
from src.utils.data_utils import parse_brazilian_number 

def prepare_lfres_context(context, row, cfg, report_type, **kwargs):
    situacao = str(row.get("Situacao", "")).strip()
    data_linha = row.get("Data")
    if data_linha is not None and not pd.isna(data_linha) and str(data_linha).strip() != "":
        context["data"] = data_linha
    else:
        try:
            df_raw_data_lfres = load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], -1)
            data_debito = df_raw_data_lfres.iloc[26, 0]
            data_credito = df_raw_data_lfres.iloc[26, 1]
            context["data"] = data_credito if situacao == "Crédito" else data_debito
        except Exception as e:
            logging.warning(f"LFRES: Não foi possível extrair a data do Excel: {e}")
            context["data"] = None
    logging.info(f"LFRES: TipoAgente='{context.get('TipoAgente')}', Valor={context.get('valor')}, Situacao='{situacao}'")
    return context

def prepare_lfn001_context(context, row, cfg, report_type, **kwargs):
    from src.utils.data_utils import parse_brazilian_number
    context["ValorLiquidacao"] = parse_brazilian_number(row.get("ValorLiquidacao", 0))
    context["ValorLiquidado"] = parse_brazilian_number(row.get("ValorLiquidado", 0))
    context["ValorInadimplencia"] = parse_brazilian_number(row.get("ValorInadimplencia", 0))
    return context

def prepare_gfn_context(context, row, cfg, report_type, **kwargs):
    try:
        df_raw_gfn = load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], -1)
        data_aporte = df_raw_gfn.iloc[23, 0]
        context["dataaporte"] = data_aporte
    except Exception as e:
        logging.warning(f"GFN: Não foi possível extrair a data do aporte do Excel: {e}")
        context["dataaporte"] = None
    return context

def prepare_sum001_context(context, row, cfg, report_type, parsed_valor, **kwargs):
    if report_type == "SUM001":
        try:
            df_raw_sum = load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], -1)
            data_debito, data_credito = df_raw_sum.iloc[23, 0], df_raw_sum.iloc[23, 1]
        except Exception:
            data_debito, data_credito = None, None
        
        situacao = str(row.get("Situacao", "")).strip()
        if situacao == "Crédito":
            context["texto1"] = "crédito"
            context["texto2"] = "Ressaltamos que esse crédito está sujeito ao rateio de inadimplência dos agentes devedores da Câmara, conforme Resolução ANEEL nº 552, de 14/10/2002."
            context["data_liquidacao"] = data_credito
        elif situacao == "Débito":
            context["texto1"] = "débito"
            context["texto2"] = "Teoricamente a conta possui o saldo necessário, mas recomendamos verificar e disponibilizar o valor com 1 (um) dia útil de antecedência."
            context["data_liquidacao"] = data_debito
        else:
            context["texto1"], context["texto2"] = "transação", "verifique os dados na planilha."
        context["valor"] = abs(parsed_valor)

    if report_type in ["LFRCAP001", "RCAP002"]:

        if report_type == "LFRCAP001":
            try:
                df_raw_lfrcap = load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], -1)
                data_aporte = df_raw_lfrcap.iloc[34, 0]
                context["dataaporte"] = data_aporte
            except Exception as e:
                logging.warning(f"LFRCAP001: Não foi possível extrair a data do aporte do Excel: {e}")
                context["dataaporte"] = None
        else:
             context["dataaporte"] = row.get("Data")
    return context

def prepare_lfrcap_context(context, row, cfg, report_type, **kwargs):
    if report_type == "LFRCAP001":
        try:
            df_raw_lfrcap = load_excel_data(cfg["excel_dados"], cfg["sheet_dados"], -1)
            data_aporte = df_raw_lfrcap.iloc[34, 0]
            context["dataaporte"] = data_aporte
        except Exception as e:
            logging.warning(f"LFRCAP001: Não foi possível extrair a data do Excel: {e}")
            context["dataaporte"] = None
    else: # RCAP002
         context["dataaporte"] = row.get("Data")
    return context

REPORT_HANDLERS = {
    "LFRES001": prepare_lfres_context,
    "LFN001": prepare_lfn001_context,
    "GFN001": prepare_gfn_context,
    "GFN - LEMBRETE": prepare_gfn_context,
    "SUM001": prepare_sum001_context,
    "LFRCAP001": prepare_lfrcap_context,
    "RCAP002": prepare_lfrcap_context,
}