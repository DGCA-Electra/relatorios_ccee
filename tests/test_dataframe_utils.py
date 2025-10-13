import pandas as pd
from utils.dataframe_utils import tratar_valores_df
from services import render_email_from_template

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

def test_variant_lfres_selection():
    common = {'month_long': 'Janeiro', 'month_num': '01', 'year': '2025'}
    row = {'Empresa': 'ACME', 'Valor': 0, 'Email': 'a@b.com', 'TipoAgente': 'Gerador-EER', 'dataaporte': '2025-01-10'}
    res = render_email_from_template('LFRES', row, common, auto_send=False)
    assert res['variant'].startswith('LFRES0')
