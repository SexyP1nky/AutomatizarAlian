# Controle de Estornos — Gerador Automático

## O que o sistema faz

1. Lê um PDF de relatório de repasse e extrai todos os segurados com **comissão negativa**
2. Cruza esses nomes com a planilha de comissões (XLSX) para encontrar o **vendedor** associado
3. Gera um arquivo XLSX formatado pronto para uso

---

## Pré-requisitos

- Python 3.11 ou superior
- pip

---

## Instalação (apenas na primeira vez)

```bash
pip install -r requirements.txt
```

---

## Como rodar localmente

```bash
streamlit run app.py
```

O navegador abrirá automaticamente em `http://localhost:8501`.

---

## Como hospedar gratuitamente no Streamlit Cloud

1. Suba esta pasta para um repositório no GitHub (pode ser privado)
2. Acesse https://streamlit.io/cloud e faça login com sua conta GitHub
3. Clique em **"New app"**, selecione o repositório e o arquivo `app.py`
4. Clique em **Deploy** — pronto, a URL fica pública e permanente

> O Streamlit Cloud é gratuito para um app público. Para app privado (acesso restrito),
> é necessário plano pago. Para uso interno sem restrição de privacidade, o plano gratuito basta.

---

## Estrutura dos arquivos

```
estornos_app/
├── app.py              → Interface Streamlit (UI e orquestração)
├── pdf_parser.py       → Extração de registros negativos do PDF
├── xlsx_parser.py      → Leitura de pares cliente/vendedor do XLSX
├── name_normalizer.py  → Normalização e comparação de nomes
├── matcher.py          → Cruzamento PDF × XLSX
├── report_builder.py   → Geração do XLSX de saída
└── requirements.txt    → Dependências Python
```

---

## Regras de negócio implementadas

- **Comissão negativa**: apenas registros com valor < 0 na coluna COMISSÃO do PDF são processados
- **Normalização de nomes**: acentos removidos exceto Ç; comparação exata por token; Ç ≈ C permitido apenas no token que contém Ç
- **Múltiplos vendedores**: se o mesmo cliente aparecer com vendedores diferentes em abas distintas, uma linha é gerada para cada vendedor
- **Valor de estorno**:
  - Vendedor aparece 1–3 vezes na tabela → R$ 30,00 por linha
  - Vendedor aparece 4+ vezes na tabela → R$ 50,00 por linha
- **Não encontrados**: listados na tela antes do download; não bloqueiam a geração do relatório para os encontrados
