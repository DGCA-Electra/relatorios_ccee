"""
tests/test_dataframe_utils.py
Testes unitários para funções utilitárias de DataFrame.
"""
import pandas as pd
from utils.dataframe_utils import tratar_valores_df

def test_tratar_valores_df():
    df = pd.DataFrame({
        'Valor': [None, 0, 10.5, 'nan'],
        'Data': [None, '2025-09-12', 'nan', '']
    })
    df_tratado = tratar_valores_df(df)
    assert all(isinstance(x, str) for x in df_tratado['Valor'])
    assert all(isinstance(x, str) for x in df_tratado['Data'])
    assert df_tratado['Valor'][0] == 'R$ 0,00'
    assert df_tratado['Data'][0] == 'Data não informada'
