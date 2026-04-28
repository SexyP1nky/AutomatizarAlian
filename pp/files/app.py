"""
Responsabilidade: interface com o usuário via Streamlit.

Fluxo:
  1. Recebe upload de um ou mais PDFs e um ou mais XLSX
  2. Consolida dados de todos os arquivos
  3. Chama cada módulo em sequência
  4. Exibe avisos e erros encontrados
  5. Disponibiliza o arquivo de saída para download
"""

import os
import tempfile

import streamlit as st

from matcher import match_records
from pdf_parser import extract_negative_commission_records
from report_builder import build_report
from vendor_sales_counter import count_sales_per_vendor_month
from xlsx_parser import extract_client_vendor_pairs


# ── Configuração da página ────────────────────────────────────────────────────

st.set_page_config(
    page_title="Controle de Estornos",
    page_icon="📋",
    layout="centered",
)

st.title("📋 Gerador de Controle de Estornos")
st.markdown("**Versão: 2.1.0** — *Com suporte a apólices e fuzzy matching*")
st.markdown(
    "Faça o upload do(s) **relatório(s) PDF** e da(s) **planilha(s) de comissões (XLSX)** "
    "para gerar automaticamente o controle de estornos."
)

# ── Aviso de contato ──────────────────────────────────────────────────────────

st.info(
    "Em caso de mudança no formato dos arquivos ou problemas encontrados, "
    "contacte o estagiário responsável."
)

# ── Inicialização do session_state ────────────────────────────────────────────

if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None
if "output_messages" not in st.session_state:
    st.session_state.output_messages = []

# ── Upload dos arquivos ───────────────────────────────────────────────────────

col_pdf, col_xlsx = st.columns(2)

with col_pdf:
    pdf_files = st.file_uploader(
        "Relatório(s) de repasse (PDF)",
        type=["pdf"],
        accept_multiple_files=True,
    )

with col_xlsx:
    xlsx_files = st.file_uploader(
        "Planilha(s) de comissões (XLSX)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
    )

# Limpar resultado anterior quando o usuário troca os arquivos
if not pdf_files or not xlsx_files:
    st.session_state.output_bytes = None
    st.session_state.output_messages = []

# ── Botão de processamento ────────────────────────────────────────────────────

if st.button(
    "Gerar relatório",
    type="primary",
    disabled=not (pdf_files and xlsx_files),
):

    # Limpar resultado anterior ao reprocessar
    st.session_state.output_bytes = None
    st.session_state.output_messages = []

    with tempfile.TemporaryDirectory() as tmp_dir:
        output_path = os.path.join(tmp_dir, "controle_estornos.xlsx")

        # ── Etapa 1: leitura de todos os PDFs ─────────────────────────────────
        all_pdf_records = []
        seen_pdf_keys = set()

        with st.spinner(f"Lendo {len(pdf_files)} PDF(s)..."):
            for idx, pdf_file in enumerate(pdf_files):
                pdf_path = os.path.join(tmp_dir, f"input_{idx}.pdf")
                with open(pdf_path, "wb") as f:
                    f.write(pdf_file.read())

                try:
                    records = extract_negative_commission_records(pdf_path)
                except Exception as e:
                    st.error(f"Erro ao ler o PDF '{pdf_file.name}': {e}")
                    st.stop()

                # Deduplicar registros entre múltiplos PDFs
                for rec in records:
                    # Verifica se já existe a mesma pessoa/data em all_pdf_records
                    existing_idx = -1
                    for i, e in enumerate(all_pdf_records):
                        if e.segurado.upper() == rec.segurado.upper() and e.inicio_vig == rec.inicio_vig:
                            if e.apolice == rec.apolice:
                                existing_idx = i
                                break
                            elif not e.apolice:
                                # Já existe um vazio, e o atual pode ter (ou ser vazio também)
                                existing_idx = i
                                break
                            elif not rec.apolice:
                                # O novo é vazio e já existe um com apólice
                                existing_idx = i
                                break
                    
                    if existing_idx == -1:
                        # Não existe ninguém com essa data/nome (ou existem, mas as apólices são diferentes e válidas)
                        all_pdf_records.append(rec)
                    else:
                        # Já existe. Vamos ver se o existente é vazio e o novo tem apólice
                        e = all_pdf_records[existing_idx]
                        if not e.apolice and rec.apolice:
                            all_pdf_records[existing_idx] = rec

        if not all_pdf_records:
            st.warning("Nenhum registro com comissão negativa encontrado nos PDFs.")
            st.stop()

        st.info(
            f"✅ {len(all_pdf_records)} registro(s) com comissão negativa "
            f"encontrado(s) em {len(pdf_files)} PDF(s)."
        )

        # ── Etapa 2: leitura de todos os XLSX ─────────────────────────────────
        all_xlsx_records = []
        all_xlsx_warnings = []

        with st.spinner(f"Lendo {len(xlsx_files)} planilha(s)..."):
            for idx, xlsx_file in enumerate(xlsx_files):
                xlsx_path = os.path.join(tmp_dir, f"input_{idx}.xlsx")
                with open(xlsx_path, "wb") as f:
                    f.write(xlsx_file.read())

                try:
                    records, warnings = extract_client_vendor_pairs(xlsx_path)
                except Exception as e:
                    st.error(f"Erro ao ler a planilha '{xlsx_file.name}': {e}")
                    st.stop()

                all_xlsx_records.extend(records)
                all_xlsx_warnings.extend(warnings)

        if all_xlsx_warnings:
            with st.expander(
                "⚠️ Problemas encontrados nas abas das planilhas",
                expanded=True,
            ):
                for w in all_xlsx_warnings:
                    st.warning(w)

        # ── Etapa 3: cruzamento PDF × XLSX ───────────────────────────────────
        with st.spinner("Cruzando dados..."):
            matched_records, not_found = match_records(
                all_pdf_records, all_xlsx_records
            )

        # ── Etapa 3b: contagem de vendas por vendedor/mês (baseado no XLSX) ───
        vendor_sales_counts = count_sales_per_vendor_month(all_xlsx_records)

        if not_found:
            with st.expander(
                f"❌ {len(not_found)} segurado(s) NÃO encontrado(s) na planilha",
                expanded=True,
            ):
                st.markdown(
                    "Os seguintes seguros estão no PDF com comissão negativa, "
                    "mas **não foram localizados** em nenhuma aba da planilha:"
                )
                for rec in not_found:
                    st.markdown(f"- `{rec.segurado}` (Início: {rec.inicio_vig} | Apólice: {rec.apolice})")

        if not matched_records:
            st.error("Nenhum registro pôde ser cruzado. Relatório não gerado.")
            st.stop()

        st.info(f"✅ {len(matched_records)} linha(s) gerada(s) no relatório.")

        # ── Etapa 4: geração do relatório ─────────────────────────────────────
        with st.spinner("Gerando XLSX..."):
            try:
                build_report(
                    matched_records,
                    output_path,
                    not_found=not_found,
                    vendor_sales_counts=vendor_sales_counts,
                )
            except Exception as e:
                st.error(f"Erro ao gerar o relatório: {e}")
                st.stop()

        # ── Armazenar bytes no session_state para persistir entre re-runs ─────
        with open(output_path, "rb") as f:
            st.session_state.output_bytes = f.read()

        st.success("Relatório gerado com sucesso!")

# ── Download (renderizado fora do bloco de processamento) ─────────────────────

if st.session_state.output_bytes is not None:
    st.download_button(
        label="⬇️ Baixar controle_estornos.xlsx",
        data=st.session_state.output_bytes,
        file_name="controle_estornos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
