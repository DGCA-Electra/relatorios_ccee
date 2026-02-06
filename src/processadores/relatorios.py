import pandas as pd
import logging
from src.utilitarios.arquivos import ler_dados_excel
from src.utilitarios.dados import converter_numero_br 

def preparar_contexto_lfres(context, row, config, tipo_relatorio, **kwargs):

    situacao = str(row.get("Situacao", "")).strip()
    data_linha = row.get("Data")
    if data_linha is not None and not pd.isna(data_linha) and str(data_linha).strip() != "":
        context["data"] = data_linha
    else:
        try:
            df_raw_data_lfres = ler_dados_excel(config["excel_dados"], config["planilha_dados"], -1)
            data_debito = df_raw_data_lfres.iloc[26, 0]
            data_credito = df_raw_data_lfres.iloc[26, 1]
            context["data"] = data_credito if situacao == "Crédito" else data_debito
        except Exception as e:
            logging.warning(f"LFRES: Não foi possível extrair a data do Excel: {e}")
            context["data"] = None

    logging.info(f"LFRES: TipoAgente='{context.get('TipoAgente')}', Valor={context.get('valor')}, Situacao='{situacao}'")
    return context

def preparar_contexto_lfn001(context, row, config, tipo_relatorio, **kwargs):

    from src.utilitarios.dados import converter_numero_br
    context["ValorLiquidacao"] = converter_numero_br(row.get("ValorLiquidacao", 0))
    context["ValorLiquidado"] = converter_numero_br(row.get("ValorLiquidado", 0))
    context["ValorInadimplencia"] = converter_numero_br(row.get("ValorInadimplencia", 0))
    return context

def preparar_contexto_gfn(context, row, config, tipo_relatorio, **kwargs):

    try:
        df_raw_gfn = ler_dados_excel(config["excel_dados"], config["planilha_dados"], -1)
        data_aporte = df_raw_gfn.iloc[23, 0]
        context["dataaporte"] = data_aporte
    except Exception as e:
        logging.warning(f"GFN: Não foi possível extrair a data do aporte do Excel: {e}")
        context["dataaporte"] = None
    return context

def preparar_contexto_sum(context, row, config, tipo_relatorio, parsed_valor, **kwargs):

    if tipo_relatorio == "SUM001":
        try:
            df_raw_sum = ler_dados_excel(config["excel_dados"], config["planilha_dados"], -1)
            data_debito, data_credito = df_raw_sum.iloc[23, 0], df_raw_sum.iloc[23, 1]
        except Exception:
            data_debito, data_credito = None, None
        situacao = str(row.get("Situacao", "")).strip()
        # Define a situação para selecionar a variante do template
        context["situacao"] = situacao
        # Define a data com base na situação
        context["data_liquidacao"] = data_credito if situacao == "Crédito" else data_debito
        # Para débito, garante valor negativo; para crédito, valor positivo
        context["valor"] = abs(parsed_valor) if situacao == "Crédito" else -abs(parsed_valor)
    if tipo_relatorio in ["LFRCAP001", "RCAP002"]:
        if tipo_relatorio == "LFRCAP001":
            try:
                df_raw_lfrcap = ler_dados_excel(config["excel_dados"], config["planilha_dados"], -1)
                data_aporte = df_raw_lfrcap.iloc[34, 0]
                context["dataaporte"] = data_aporte
            except Exception as e:
                logging.warning(f"LFRCAP001: Não foi possível extrair a data do aporte do Excel: {e}")
                context["dataaporte"] = None
        else:
             context["dataaporte"] = row.get("Data")
    return context
def preparar_contexto_lfrcap(context, row, config, tipo_relatorio, **kwargs):
    if tipo_relatorio == "LFRCAP001":
        try:
            df_raw_lfrcap = ler_dados_excel(config["excel_dados"], config["planilha_dados"], -1)
            data_aporte = df_raw_lfrcap.iloc[34, 0]
            context["dataaporte"] = data_aporte
        except Exception as e:
            logging.warning(f"LFRCAP001: Não foi possível extrair a data do Excel: {e}")
            context["dataaporte"] = None
    else: # RCAP002
         context["dataaporte"] = row.get("Data")
    return context

def processador_generico_relatorio(context, row, config, tipo_relatorio, **kwargs):
    """
    Handler universal que processa extrações baseadas puramente no JSON de configuração.
    Permite definir: 'extra_fields': [{'name': 'data_venc', 'row': 23, 'col': 1}]
    """
    extra_fields = config.get("extra_fields", [])
    if extra_fields:
        try:
            df_raw = ler_dados_excel(config["excel_dados"], config["planilha_dados"], -1)
            for field in extra_fields:
                field_name = field.get("name")
                r = int(field.get("row", 0))
                c = int(field.get("col", 0))
                try:
                    val = df_raw.iloc[r, c]
                    context[field_name] = val
                    logging.info(f"[{tipo_relatorio}] Extraído '{field_name}' da celula ({r},{c}): {val}")
                except IndexError:
                    logging.warning(f"[{tipo_relatorio}] Erro ao extrair '{field_name}': Coordenada ({r},{c}) inválida.")
                    context[field_name] = "N/D"
        except Exception as e:
            logging.error(f"[{tipo_relatorio}] Erro ao carregar Excel para extração genérica: {e}")
    return context

PROCESSADORES_RELATORIO = {

    "LFRES001": preparar_contexto_lfres,
    "LFN001": preparar_contexto_lfn001,
    "GFN001": preparar_contexto_gfn,
    "GFN - LEMBRETE": preparar_contexto_gfn,
    "SUM001": preparar_contexto_sum,
    "LFRCAP001": preparar_contexto_lfrcap,
    "RCAP002": preparar_contexto_lfrcap,
}