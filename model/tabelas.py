import pandas as pd
from apps.relatorios_ccee.model.utils_dados import formatar_moeda, formatar_data

def tratar_valores_df(df: pd.DataFrame, colunas_moeda=None, colunas_data=None, mapa_preenchimento=None):

    colunas_moeda = colunas_moeda or []
    colunas_data = colunas_data or []
    mapa_preenchimento = mapa_preenchimento or {}

    for col in df.columns:
        if col in colunas_moeda or any(v in col.lower() for v in ["valor", "inadimplência"]):
            df[col] = df[col].apply(lambda x: "R$ 0,00" if pd.isna(x) or x in [None, 0, 0.0, "0", "0.0", "nan", "None"] else formatar_moeda(x))
        if col in colunas_data or "data" in col.lower():
            df[col] = df[col].apply(lambda x: "Data não informada" if pd.isna(x) else formatar_data(x))
        if col in mapa_preenchimento:
            df[col] = df[col].fillna(mapa_preenchimento[col])
    return df