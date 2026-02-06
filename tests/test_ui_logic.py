import pytest
import streamlit as st
from src.interface.principal import iniciar_estado_sessao as init_state


def test_init_state_sets_defaults():

    st.session_state.clear()

    init_state()

    assert st.session_state["report_type"] == "GFN001"

    assert st.session_state["analyst"] == "Artur Bello Rodrigues"

    assert st.session_state["month"] == "JANEIRO"

    assert st.session_state["year"] == 2025


from src.interface.principal import unir_emails_seguro as safe_join_emails


def test_safe_join_emails_list():

    emails = ["a@b.com", "c@d.com ", None, " "]

    result = safe_join_emails(emails)

    assert result == "a@b.com; c@d.com"


def test_safe_join_emails_str():

    emails = "a@b.com; c@d.com; ;"

    result = safe_join_emails(emails)

    assert result == "a@b.com; c@d.com"


def test_safe_join_emails_none():

    assert safe_join_emails(None) == ""


from jinja2 import Environment, FileSystemLoader


def test_render_email_template(tmp_path):

    # Cria template temporário

    template_dir = tmp_path

    template_file = template_dir / "email_template.html"

    template_file.write_text("<p>Olá {{ nome }}</p>")

    env = Environment(loader=FileSystemLoader(str(template_dir)))

    template = env.get_template("email_template.html")

    html = template.render(nome="Teste")

    assert "Olá Teste" in html


def test_display_mode_default(monkeypatch):

    # Não envia de fato; apenas verifica que função existe e não quebra

    from src.servicos import renderizar_email_modelo

    common = {'month_long': 'Janeiro', 'month_num': '01', 'year': '2025'}

    row = {'Empresa': 'ACME', 'Valor': 10, 'Email': 'a@b.com', 'Situacao': 'Débito', 'dataaporte': '2025-01-12'}
    res = renderizar_email_modelo('GFN001', row, common, auto_send=False)

    assert 'GFN001' in res['subject']