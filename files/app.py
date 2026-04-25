"""
Responsabilidade: interface com o usuário via Streamlit.

Fluxo:
  1. Recebe upload do PDF e do XLSX
  2. Chama cada módulo em sequência
  3. Exibe avisos e erros encontrados
  4. Disponibiliza o arquivo de saída para download
"""

import os
import tempfile

import streamlit as st

from matcher import match_records
from pdf_parser import extract_negative_commission_records
from report_builder import build_report
from xlsx_parser import extract_client_vendor_pairs


# ── Configuração da página ────────────────────────────────────────────────────

st.set_page_config(
    page_title="Controle de Estornos",
    page_icon="📋",
    layout="centered",
)

st.title("📋 Gerador de Controle de Estornos")
st.markdown(
    "Faça o upload do **relatório PDF** e da **planilha de comissões (XLSX)** "
    "para gerar automaticamente o controle de estornos."
)

# ── Upload dos arquivos ───────────────────────────────────────────────────────

col_pdf, col_xlsx = st.columns(2)

with col_pdf:
    pdf_file = st.file_uploader("Relatório de repasse (PDF)", type=["pdf"])

with col_xlsx:
    xlsx_file = st.file_uploader("Planilha de comissões (XLSX)", type=["xlsx", "xls"])

# ── Botão de processamento ────────────────────────────────────────────────────

if st.button("Gerar relatório", type="primary", disabled=not (pdf_file and xlsx_file)):

    with tempfile.TemporaryDirectory() as tmp_dir:
        pdf_path = os.path.join(tmp_dir, "input.pdf")
        xlsx_path = os.path.join(tmp_dir, "input.xlsx")
        output_path = os.path.join(tmp_dir, "controle_estornos.xlsx")

        with open(pdf_path, "wb") as f:
            f.write(pdf_file.read())
        with open(xlsx_path, "wb") as f:
            f.write(xlsx_file.read())

        # ── Etapa 1: leitura do PDF ───────────────────────────────────────────
        with st.spinner("Lendo PDF..."):
            try:
                pdf_records = extract_negative_commission_records(pdf_path)
            except Exception as e:
                st.error(f"Erro ao ler o PDF: {e}")
                st.stop()

        if not pdf_records:
            st.warning("Nenhum registro com comissão negativa encontrado no PDF.")
            st.stop()

        st.info(f"✅ {len(pdf_records)} registro(s) com comissão negativa encontrado(s) no PDF.")

        # ── Etapa 2: leitura do XLSX ──────────────────────────────────────────
        with st.spinner("Lendo planilha..."):
            try:
                xlsx_records, xlsx_warnings = extract_client_vendor_pairs(xlsx_path)
            except Exception as e:
                st.error(f"Erro ao ler a planilha: {e}")
                st.stop()

        if xlsx_warnings:
            with st.expander("⚠️ Problemas encontrados nas abas da planilha", expanded=True):
                for w in xlsx_warnings:
                    st.warning(w)

        # ── Etapa 3: cruzamento PDF × XLSX ───────────────────────────────────
        with st.spinner("Cruzando dados..."):
            matched_records, not_found = match_records(pdf_records, xlsx_records)

        if not_found:
            with st.expander(
                f"❌ {len(not_found)} segurado(s) NÃO encontrado(s) na planilha",
                expanded=True,
            ):
                st.markdown(
                    "Os seguintes nomes estão no PDF com comissão negativa, "
                    "mas **não foram localizados** em nenhuma aba da planilha:"
                )
                for name in not_found:
                    st.markdown(f"- `{name}`")

        if not matched_records:
            st.error("Nenhum registro pôde ser cruzado. Relatório não gerado.")
            st.stop()

        st.info(f"✅ {len(matched_records)} linha(s) gerada(s) no relatório.")

        # ── Etapa 4: geração do relatório ─────────────────────────────────────
        with st.spinner("Gerando XLSX..."):
            try:
                build_report(matched_records, output_path)
            except Exception as e:
                st.error(f"Erro ao gerar o relatório: {e}")
                st.stop()

        # ── Download ──────────────────────────────────────────────────────────
        with open(output_path, "rb") as f:
            output_bytes = f.read()

        st.success("Relatório gerado com sucesso!")
        st.download_button(
            label="⬇️ Baixar controle_estornos.xlsx",
            data=output_bytes,
            file_name="controle_estornos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
