import pytest

from src.servicos import renderizar_email_modelo


def test_render_placeholders_sum001():

    common = {'month_long': 'Janeiro', 'month_num': '01', 'year': '2025'}

    row = {'Empresa': 'ACME', 'Valor': 123.45, 'Email': 'a@b.com', 'Situacao': 'Débito', 'dataaporte': '2025-01-12'}

    res = renderizar_email_modelo('SUM001', row, common, auto_send=False)

    assert 'SUM001' in res['subject']

    assert res['missing_placeholders'] == []


TEST_TEMPLATE = """

<p>Prezado(a) {{ analista }},</p>

<p>Empresa: {{ empresa }}</p>

<p>Valor: {{ valor }}</p>
"""


def test_montar_corpo_email_jinja2():

    dados = {

        'Analista': 'João',

        'Empresa': 'Eletrobras',

        'Valor': 'R$ 1000,00'

    }

    corpo = montar_corpo_email(dados, TEST_TEMPLATE)

    assert "Prezado(a) João" in corpo

    assert "Empresa: Eletrobras" in corpo

    assert "Valor: R$ 1000,00" in corpo



def test_montar_corpo_email_campos_faltantes():

    dados = {

        'Analista': 'Maria'

    }

    corpo = montar_corpo_email(dados, TEST_TEMPLATE)

    assert "Prezado(a) Maria" in corpo
    assert "Empresa: " in corpo  

    assert "Valor: " in corpo  

