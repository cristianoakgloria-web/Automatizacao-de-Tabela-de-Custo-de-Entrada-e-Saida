import streamlit as st
import pandas as pd
import os
from tabela_auto import header, body, footer

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão de Caixa Automática", layout="wide")

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
    # Pegamos os valores atuais dos widgets via session_state
    data = st.session_state["input_data"]
    desc = st.session_state["input_desc"]
    ent = st.session_state["input_ent"]
    sai = st.session_state["input_sai"]
    
    if data and desc:
        # Cálculo do saldo base
        smp_val = st.session_state.get("val_smp", 0.0)
        des_val = st.session_state.get("val_des", 0.0)
        saldo_base = smp_val - des_val
        
        s_anterior = st.session_state.lista_itens[-1][5] if st.session_state.lista_itens else saldo_base
        novo_saldo = s_anterior + ent - sai
        
        st.session_state.lista_itens.append([
            len(st.session_state.lista_itens) + 1, data, desc, ent, sai, novo_saldo
        ])
        
        # LIMPAR OS CAMPOS APÓS ADICIONAR
        st.session_state["input_data"] = ""
        st.session_state["input_desc"] = ""
        st.session_state["input_ent"] = 0.0
        st.session_state["input_sai"] = 0.0
    else:
        st.warning("Preencha Data e Designação antes de adicionar.")

# --- INTERFACE ---
st.sidebar.title("MENU")
opcao = st.sidebar.radio("Ir para:", ["Nova Tabela", "Sobre o App"])

if opcao == "Nova Tabela":
    st.header("1. Configuração do Diário")
    c1, c2, c3 = st.columns(3)
    ano = c1.text_input("Ano", "2026")
    mes = c2.text_input("Mês", "MARÇO").upper()
    moeda = c3.text_input("Moeda", "AKZ")
    
    c4, c5 = st.columns(2)
    # Guardamos os valores iniciais no session_state para a função de cálculo acessar
    st.session_state["val_smp"] = c4.number_input("Saldo Mês Anterior", value=0.0)
    st.session_state["val_des"] = c5.number_input("Despesa Mês Anterior", value=0.0)
    saldo_base = st.session_state["val_smp"] - st.session_state["val_des"]

    st.divider()
    st.header("2. Adicionar/Editar Movimentação")
    
    col_a, col_b, col_c, col_d = st.columns([1, 2, 1, 1])
    
    # Widgets vinculados ao session_state
    st.text_input("Data (dd-mm-aaaa)", key="input_data")
    st.text_input("Designação", key="input_desc")
    st.number_input("Entrada", key="input_ent", step=500.0)
    st.number_input("Saída", key="input_sai", step=500.0)

    # O segredo é usar o on_click
    st.button("➕ Adicionar Item", on_click=processar_submissao)

    if st.session_state.lista_itens:
        st.divider()
        st.subheader("3. Visualização e Edição")
        df = pd.DataFrame(st.session_state.lista_itens, columns=["Nº", "Data", "Designação", "Entrada", "Saída", "Saldo"])
        st.dataframe(df, use_container_width=True)

        tamanho = len(st.session_state.lista_itens)
        c_ed, c_rm = st.columns(2)

        with c_ed:
            n_edit = st.number_input("Editar item Nº:", min_value=1, max_value=tamanho, step=1, key="n_edit_input")
            if st.button("📝 Carregar para Editar"):
                item = st.session_state.lista_itens.pop(int(n_edit) - 1)
                # Joga os valores de volta para os campos de cima
                st.session_state["input_data"] = item[1]
                st.session_state["input_desc"] = item[2]
                st.session_state["input_ent"] = item[3]
                st.session_state["input_sai"] = item[4]
                
                # Recalcular saldos
                s_base_temp = saldo_base
                for i, it in enumerate(st.session_state.lista_itens):
                    it[0] = i + 1
                    it[5] = s_base_temp + it[3] - it[4]
                    s_base_temp = it[5]
                st.rerun()

        with c_rm:
            n_rem = st.number_input("Remover item Nº:", min_value=1, max_value=tamanho, step=1, key="n_rem_input")
            if st.button("🗑️ Remover Definitivamente"):
                st.session_state.lista_itens.pop(int(n_rem) - 1)
                s_base_temp = saldo_base
                for i, it in enumerate(st.session_state.lista_itens):
                    it[0] = i + 1
                    it[5] = s_base_temp + it[3] - it[4]
                    s_base_temp = it[5]
                st.rerun()

        st.divider()
        cnome, csexo = st.columns(2)
        nome_tes = cnome.text_input("Nome do Tesoureiro/a")
        sexo = csexo.selectbox("Sexo", ["M", "F"])

        if nome_tes:
            try:
                header(ano, mes, moeda, saldo_base)
                body(st.session_state.lista_itens, mes, ano)
                total_ent = sum(x[3] for x in st.session_state.lista_itens)
                total_sai = sum(x[4] for x in st.session_state.lista_itens)
                footer(st.session_state.lista_itens, mes, ano, sexo, nome_tes, st.session_state.lista_itens[-1][5], total_sai, st.session_state["val_smp"], st.session_state["val_des"], total_ent)
                
                with open("Custo_Entrada&Saida.xlsx", "rb") as f:
                    st.download_button("📥 DESCARREGAR EXCEL", f.read(), f"DIARIO_{mes}_{ano}.xlsx", type="primary")
            except Exception as e:
                st.error(f"Erro ao gerar Excel: {e}")

elif opcao == "Sobre o App":
    st.write("Desenvolvido por Cristiano Glória.")