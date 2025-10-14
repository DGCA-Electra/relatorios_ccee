import pandas as pd
import src.services as services

def tratar_valores_df(df: pd.DataFrame, currency_cols=None, date_cols=None, fill_map=None):
    currency_cols = currency_cols or []
    date_cols = date_cols or []
    fill_map = fill_map or {}
    for col in df.columns:
        if col in currency_cols or any(v in col.lower() for v in ["valor", "inadimplência"]):
            df[col] = df[col].apply(lambda x: "R$ 0,00" if pd.isna(x) or x in [None, 0, 0.0, "0", "0.0", "nan", "None"] else services._format_currency(x))
        if col in date_cols or "data" in col.lower():
            df[col] = df[col].apply(lambda x: "Data não informada" if pd.isna(x) else services._format_date(x))
        if col in fill_map:
            df[col] = df[col].fillna(fill_map[col])
    return df
