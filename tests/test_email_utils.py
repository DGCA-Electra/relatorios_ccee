import pytest
from mail.email_utils import montar_corpo_email

# Template Jinja2 simplificado para teste
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

# Teste de campos faltantes

def test_montar_corpo_email_campos_faltantes():
    dados = {
        'Analista': 'Maria'
    }
    corpo = montar_corpo_email(dados, TEST_TEMPLATE)
    assert "Prezado(a) Maria" in corpo
    assert "Empresa: " in corpo  # Campo vazio
    assert "Valor: " in corpo    # Campo vazio
