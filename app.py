import streamlit as st
import pandas as pd
import os
from tabela_auto import header, body, footer

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Gestão de Caixa Automática",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CSS CUSTOMIZADO ---
st.markdown("""
<style>
    /* Reset e fundo */
    .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
    }

    /* Estilo da sidebar */
    section[data-testid="stSidebar"] {
        background-color: #1B2432;
    }
    section[data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] .stRadio > label {
        font-size: 16px !important;
        font-weight: 500 !important;
        color: #E0E0E0 !important;
    }
    section[data-testid="stSidebar"] .stRadio > div {
        gap: 0.5rem !important;
    }

    /* Cards de métricas */
    .metric-card {
        background: linear-gradient(135deg, #1B2432, #2C3E50);
        border-radius: 12px;
        padding: 1.3rem;
        text-align: center;
        border: 1px solid rgba(255, 255, 255, 0.08);
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.15);
        min-height: 120px;
    }
    .metric-card .icon {
        font-size: 32px;
        margin-bottom: 4px;
    }
    .metric-card .label {
        color: #A0AEC0;
        font-size: 13px;
        margin-bottom: 4px;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .metric-card .value {
        color: #FFFFFF;
        font-size: 22px;
        font-weight: 700;
    }

    /* Cards de seções */
    .section-card {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 1.5rem;
        border: 1px solid #E2E8F0;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
        margin-bottom: 1.2rem;
    }
    .section-title {
        color: #1B2432;
        font-size: 18px;
        font-weight: 700;
        margin: 0;
        padding-bottom: 12px;
        border-bottom: 2px solid #F7C948;
        margin-bottom: 1rem;
    }

    /* Tabela estilizada */
    .stDataFrame {
        border-radius: 12px !important;
        overflow: hidden !important;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
    }
    div[data-testid="stDataframe"] thead tr th {
        background-color: #1B2432 !important;
        color: #FFFFFF !important;
    }

    /* Botões de ação */
    .stButton > button {
        font-weight: 600;
        border-radius: 8px;
        padding: 0.5rem 1.5rem;
    }

    /* Formulário de movimentação */
    .form-row {
        display: flex;
        gap: 0.8rem;
        align-items: end;
        flex-wrap: wrap;
    }

    /* Download button */
    .download-wrapper {
        text-align: center;
        padding: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# 1. Inicializar estados se não existirem
if 'lista_itens' not in st.session_state:
    st.session_state.lista_itens = []

# Inicializar os valores dos campos para evitar erro de primeira execução
if "input_data" not in st.session_state:
    st.session_state["input_data"] = ""
if "input_desc" not in st.session_state:
    st.session_state["input_desc"] = ""
if "input_ent" not in st.session_state:
    st.session_state["input_ent"] = 0.0
if "input_sai" not in st.session_state:
    st.session_state["input_sai"] = 0.0

# --- FUNÇÃO DISPARADA PELO BOTÃO (CALLBACK) ---
def processar_submissao():
    data = st.session_state["input_data"]
    desc = st.session_state["input_desc"]
    ent = st.session_state["input_ent"]
    sai = st.session_state["input_sai"]

    if data and desc:
        smp_val = st.session_state.get("val_smp", 0.0)
        des_val = st.session_state.get("val_des", 0.0)
        saldo_base = smp_val - des_val

        s_anterior = st.session_state.lista_itens[-1][5] if st.session_state.lista_itens else saldo_base
        novo_saldo = s_anterior + ent - sai

        # Formatação da moeda no saldo
        st.session_state.lista_itens.append([
            len(st.session_state.lista_itens) + 1, data, desc, ent, sai, novo_saldo
        ])

        st.session_state["input_data"] = ""
        st.session_state["input_desc"] = ""
        st.session_state["input_ent"] = 0.0
        st.session_state["input_sai"] = 0.0

        st.toast("Item adicionado com sucesso!", icon="✅")
    else:
        st.warning("Preencha **Data** e **Designação** antes de adicionar.", icon="⚠️")

# --- SIDEBAR ---
with st.sidebar:
    st.markdown("# 🏦")
    st.markdown("###### Menu")
    st.markdown("")
    opcao = st.radio("", ["Nova Tabela", "Sobre o App"], label_visibility="collapsed")
    st.markdown("")
    st.divider()
    st.caption("Desenvolvido por **Cristiano Glória**")

# --- HEADER ---
st.markdown("## 📊 Gestão de Caixa Automática")
st.subheader("Diário de movimentações de entrada e saída")
st.markdown("---")

if opcao == "Nova Tabela":
    # ========== SEÇÃO 1: CONFIGURAÇÃO ==========
    with st.container():
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<p class="section-title">1. Configuração do Diário</p>', unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.text_input("Ano", "2026", key="conf_ano")
        with c2:
            st.text_input("Mês", "MARÇO", key="conf_mes")
        with c3:
            st.text_input("Moeda", "AKZ", key="conf_moeda")

        c4, c5 = st.columns(2)
        with c4:
            st.session_state["val_smp"] = st.number_input("Saldo Mês Anterior", value=0.0, format="%.2f")
        with c5:
            st.session_state["val_des"] = st.number_input("Despesa Mês Anterior", value=0.0, format="%.2f")

        st.markdown("</div>", unsafe_allow_html=True)

    # ========== SEÇÃO 2: MOVIMENTAÇÃO ==========
    with st.container():
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown('<p class="section-title">2. Adicionar Movimentação</p>', unsafe_allow_html=True)

        st.text_input("Data (dd-mm-aaaa)", key="input_data")
        st.text_input("Designação", key="input_desc")

        c_ent, c_sai = st.columns(2)
        with c_ent:
            st.number_input("Entrada (+)", key="input_ent", step=500.0, format="%.2f")
        with c_sai:
            st.number_input("Saída (-)", key="input_sai", step=500.0, format="%.2f")

        st.button("➕  Adicionar Item", on_click=processar_submissao, type="primary", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    # ========== CARDS DE RESUMO ==========
    if st.session_state.lista_itens:
        total_ent = sum(x[3] for x in st.session_state.lista_itens)
        total_sai = sum(x[4] for x in st.session_state.lista_itens)
        saldo_base = st.session_state["val_smp"] - st.session_state["val_des"]
        s_anterior = st.session_state.lista_itens[-1][5]

        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="icon">📈</div>
                    <div class="label">Total Entradas</div>
                    <div class="value">AKZ {total_ent:,.2f}</div>
                </div>
            """, unsafe_allow_html=True)
        with mc2:
            st.markdown(f"""
                <div class="metric-card">
                    <div class="icon">📉</div>
                    <div class="label">Total Saídas</div>
                    <div class="value">AKZ {total_sai:,.2f}</div>
                </div>
            """, unsafe_allow_html=True)
        with mc3:
            cor = "#48BB78" if s_anterior >= 0 else "#F56565"
            st.markdown(f"""
                <div class="metric-card">
                    <div class="icon">💰</div>
                    <div class="label">Saldo Atual</div>
                    <div class="value" style="color: {cor}">AKZ {s_anterior:,.2f}</div>
                </div>
            """, unsafe_allow_html=True)

        # ========== SEÇÃO 3: TABELA E EDIÇÃO ==========
        with st.container():
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<p class="section-title">3. Movimentações Registadas</p>', unsafe_allow_html=True)

            df = pd.DataFrame(st.session_state.lista_itens, columns=["Nº", "Data", "Designação", "Entrada", "Saída", "Saldo"])
            st.dataframe(df, use_container_width=True, height=300)

            tamanho = len(st.session_state.lista_itens)
            c_ed, c_rm = st.columns(2)

            with c_ed:
                n_edit = st.number_input("Número do item para editar", min_value=1, max_value=tamanho, step=1, key="n_edit_input")
                if st.button("📝 Carregar para Editar", use_container_width=True):
                    item = st.session_state.lista_itens.pop(int(n_edit) - 1)
                    st.session_state["input_data"] = item[1]
                    st.session_state["input_desc"] = item[2]
                    st.session_state["input_ent"] = item[3]
                    st.session_state["input_sai"] = item[4]

                    saldo_base = st.session_state["val_smp"] - st.session_state["val_des"]
                    for i, it in enumerate(st.session_state.lista_itens):
                        it[0] = i + 1
                        it[5] = saldo_base + it[3] - it[4]
                        saldo_base = it[5]
                    st.rerun()

            with c_rm:
                n_rem = st.number_input("Número do item para remover", min_value=1, max_value=tamanho, step=1, key="n_rem_input")
                if st.button("🗑️ Remover Definitivamente", use_container_width=True, type="secondary"):
                    st.session_state.lista_itens.pop(int(n_rem) - 1)
                    saldo_base = st.session_state["val_smp"] - st.session_state["val_des"]
                    for i, it in enumerate(st.session_state.lista_itens):
                        it[0] = i + 1
                        it[5] = saldo_base + it[3] - it[4]
                        saldo_base = it[5]
                    st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)

        # ========== SEÇÃO 4: ASSINATURA E DOWNLOAD ==========
        with st.container():
            st.markdown('<div class="section-card">', unsafe_allow_html=True)
            st.markdown('<p class="section-title">4. Gerar Documento Excel</p>', unsafe_allow_html=True)

            cnome, csexo = st.columns(2)
            with cnome:
                nome_tes = st.text_input("Nome do Tesoureiro/a")
            with csexo:
                sexo = st.selectbox("Sexo", ["M", "F"], key="sexo_select")

            if nome_tes:
                try:
                    saldo_base = st.session_state["val_smp"] - st.session_state["val_des"]

                    with st.spinner("A gerar o ficheiro Excel..."):
                        header(ano=st.session_state.get("conf_ano", "2026"),
                               mes=st.session_state.get("conf_mes", "MARÇO").upper(),
                               moeda=st.session_state.get("conf_moeda", "AKZ"),
                               smp=saldo_base)
                        body(st.session_state.lista_itens, st.session_state.get("conf_mes", "MARÇO").upper(),
                             st.session_state.get("conf_ano", "2026"))
                        footer(st.session_state.lista_itens,
                               st.session_state.get("conf_mes", "MARÇO").upper(),
                               st.session_state.get("conf_ano", "2026"),
                               sexo, nome_tes,
                               st.session_state.lista_itens[-1][5],
                               total_sai,
                               st.session_state["val_smp"],
                               st.session_state["val_des"],
                               total_ent)

                    st.success("Ficheiro gerado com sucesso!", icon="✅")
                    st.markdown('<div class="download-wrapper">', unsafe_allow_html=True)
                    with open("Custo_Entrada&Saida.xlsx", "rb") as f:
                        st.download_button(
                            "📥  DESCARREGAR EXCEL",
                            f.read(),
                            f"DIARIO_{st.session_state.get('conf_mes', 'MARÇO').upper()}_{st.session_state.get('conf_ano', '2026')}.xlsx",
                            type="primary",
                            use_container_width=False
                        )
                    st.markdown("</div>", unsafe_allow_html=True)
                except Exception as e:
                    st.error(f"Erro ao gerar Excel: {e}")

            st.markdown("</div>", unsafe_allow_html=True)

elif opcao == "Sobre o App":
    with st.container():
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.markdown("## ℹ️ Sobre o App")
        st.markdown("""
**Gestão de Caixa Automática** — ferramenta para automatizar a criação de tabelas
de custo de entrada e saída, para a empresa **Logotipo Oficial | Marketing & Vendas**.

Esta aplicação permite:
- Registo diário de movimentações de caixa
- Cálculo automático de saldos
- Geração de ficheiros Excel formatados prontos para uso

_Desenvolvido por **Cristiano Glória**_
        """)
        st.markdown("</div>", unsafe_allow_html=True)