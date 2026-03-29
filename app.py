import customtkinter as ctk
from tabela_auto import header, body, footer
from tkinter import messagebox
import webbrowser

# --- CONFIGURAÇÃO GLOBAL ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Variáveis globais para armazenar os dados da sessão
lista_itens = []
widgets_edicao = {}

# --- FUNÇÕES DE NAVEGAÇÃO E LÓGICA ---

def limpar_palco():
    """Remove todos os widgets do frame principal para trocar de tela."""
    for widget in palco_conteudo.winfo_children():
        widget.destroy()

def mostrar_sobre():
    """Desenha a tela 'Sobre' com informações e guia de instalação."""
    limpar_palco()

    # Título Principal
    ctk.CTkLabel(palco_conteudo, text="Sobre o Sistema", font=("Arial", 24, "bold")).pack(pady=(20, 10))
    
    # Frame de Informações do Autor
    frame_info = ctk.CTkFrame(palco_conteudo)
    frame_info.pack(fill="x", padx=30, pady=10)
    
    ctk.CTkLabel(frame_info, text="Desenvolvedor: Cristiano Glória", font=("Arial", 14, "bold")).pack(pady=5)
    ctk.CTkLabel(frame_info, text="Suporte Técnico: cristianoakgloria@gmail.com", font=("Arial", 12)).pack()
    
    # Link do GitHub funcional
    link_github = ctk.CTkLabel(frame_info, text="Código Fonte: github.com/cristianoakgloria-web/Automatizacao...", 
                               text_color="#1f538d", cursor="hand2", font=("Arial", 12, "underline"))
    link_github.pack(pady=5)
    link_github.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/cristianoakgloria-web/Automatizacao-de-Tabela-de-Custo-de-Entrada-e-Saida"))

    # Instruções de Uso
    ctk.CTkLabel(palco_conteudo, text="Instruções de Uso:", font=("Arial", 16, "bold")).pack(pady=(15, 5))
    txt_uso = ctk.CTkTextbox(palco_conteudo, height=100, width=600)
    txt_uso.pack(padx=30, pady=5)
    txt_uso.insert("0.0", "1. Clique em 'Nova Tabela' e preencha os dados de configuração.\n"
                          "2. Adicione os itens. Os campos limpam-se automaticamente após adicionar.\n"
                          "3. Se precisar editar, altere os valores na lista e clique em 'Salvar' na linha correspondente.\n"
                          "4. Clique em 'GERAR EXCEL FINAL' para criar o ficheiro .xlsx.")
    txt_uso.configure(state="disabled")

    # Guia de Instalação (Estilizado como Terminal: Fundo Preto e Letras Verdes)
    ctk.CTkLabel(palco_conteudo, text="Como Gerar Executável (PyInstaller):", font=("Arial", 16, "bold")).pack(pady=(15, 5))
    txt_install = ctk.CTkTextbox(palco_conteudo, height=130, width=600, fg_color="black", text_color="#00FF00")
    txt_install.pack(padx=30, pady=5)
    
    comando_pyinst = (
        "Use o PyInstaller no terminal para Mac, Windows ou Linux:\n\n"
        "1. Instalar: pip install pyinstaller\n"
        "2. Gerar: pyinstaller --noconsole --onefile --collect-all customtkinter app.py\n\n"
        "O executável será criado na pasta 'dist'."
    )
    txt_install.insert("0.0", comando_pyinst)
    txt_install.configure(state="disabled")

def mostrar_novaTabela():
    """Desenha a interface de criação de tabela."""
    global widgets_edicao
    widgets_edicao = {}
    limpar_palco()

    # 1. CONFIGURAÇÃO DO CABEÇALHO
    ctk.CTkLabel(palco_conteudo, text="1. Configuração do Diário", font=("Arial", 16, "bold")).pack(pady=(10, 5))
    frame_config = ctk.CTkFrame(palco_conteudo)
    frame_config.pack(fill="x", padx=20, pady=5)

    ent_ano = ctk.CTkEntry(frame_config, placeholder_text="Ano (ex: 2026)"); ent_ano.grid(row=0, column=0, padx=5, pady=5)
    ent_mes = ctk.CTkEntry(frame_config, placeholder_text="Mês (ex: MARÇO)"); ent_mes.grid(row=0, column=1, padx=5, pady=5)
    ent_moeda = ctk.CTkEntry(frame_config, placeholder_text="Moeda (ex: AKZ)"); ent_moeda.grid(row=0, column=2, padx=5, pady=5)
    ent_smp = ctk.CTkEntry(frame_config, placeholder_text="Saldo Mês Anterior"); ent_smp.grid(row=1, column=0, padx=5, pady=5)
    ent_des_ant = ctk.CTkEntry(frame_config, placeholder_text="Despesa Mês Anterior"); ent_des_ant.grid(row=1, column=1, padx=5, pady=5)

    # 2. ADIÇÃO DE ITENS
    ctk.CTkLabel(palco_conteudo, text="2. Adicionar Movimentação", font=("Arial", 14, "bold")).pack(pady=(10, 5))
    frame_add = ctk.CTkFrame(palco_conteudo)
    frame_add.pack(fill="x", padx=20, pady=5)

    e_data = ctk.CTkEntry(frame_add, placeholder_text="Data (dd-mm)"); e_data.grid(row=0, column=0, padx=5, pady=5)
    e_desc = ctk.CTkEntry(frame_add, placeholder_text="Designação"); e_desc.grid(row=0, column=1, padx=5, pady=5)
    e_ent = ctk.CTkEntry(frame_add, placeholder_text="Valor Entrada"); e_ent.grid(row=1, column=0, padx=5, pady=5)
    e_sai = ctk.CTkEntry(frame_add, placeholder_text="Valor Saída"); e_sai.grid(row=1, column=1, padx=5, pady=5)

    # 3. LISTA COM SCROLL (VISUALIZAÇÃO E EDIÇÃO)
    scroll_frame = ctk.CTkScrollableFrame(palco_conteudo, height=180)
    scroll_frame.pack(fill="both", expand=True, padx=20, pady=10)

    def atualizar_tabela_visual():
        """Atualiza os widgets dentro do frame de scroll."""
        for w in scroll_frame.winfo_children(): 
            w.destroy()
        
        headers = ["Nº", "Data", "Designação", "Entrada", "Saída", "Ações"]
        for i, h in enumerate(headers): 
            ctk.CTkLabel(scroll_frame, text=h, font=("Arial", 11, "bold")).grid(row=0, column=i, padx=10)

        for idx, item in enumerate(lista_itens):
            row = idx + 1
            ctk.CTkLabel(scroll_frame, text=str(item[0])).grid(row=row, column=0)
            
            ed_dat = ctk.CTkEntry(scroll_frame, width=80); ed_dat.insert(0, item[1]); ed_dat.grid(row=row, column=1, padx=2)
            ed_des = ctk.CTkEntry(scroll_frame, width=150); ed_des.insert(0, item[2]); ed_des.grid(row=row, column=2, padx=2)
            ed_ent = ctk.CTkEntry(scroll_frame, width=80); ed_ent.insert(0, str(item[3])); ed_ent.grid(row=row, column=3, padx=2)
            ed_sai = ctk.CTkEntry(scroll_frame, width=80); ed_sai.insert(0, str(item[4])); ed_sai.grid(row=row, column=4, padx=2)
            
            widgets_edicao[idx] = [ed_dat, ed_des, ed_ent, ed_sai]
            ctk.CTkButton(scroll_frame, text="Salvar", width=60, fg_color="gray", 
                          command=lambda i=idx: salvar_edicao_linha(i)).grid(row=row, column=5, padx=2)

    def salvar_edicao_linha(i):
        """Atualiza a lista global com os dados das entries editadas."""
        try:
            w = widgets_edicao[i]
            lista_itens[i][1] = w[0].get()
            lista_itens[i][2] = w[1].get()
            lista_itens[i][3] = float(w[2].get() or 0)
            lista_itens[i][4] = float(w[3].get() or 0)
            recalcular_tudo()
            messagebox.showinfo("Sucesso", f"Linha {i+1} atualizada e saldos recalculados.")
        except ValueError:
            messagebox.showerror("Erro", "Valores numéricos inválidos na edição.")

    def recalcular_tudo():
        """Refaz os saldos acumulados de toda a lista."""
        try:
            saldo_anterior = float(ent_smp.get() or 0) - float(ent_des_ant.get() or 0)
            for it in lista_itens:
                it[5] = saldo_anterior + it[3] - it[4]
                saldo_anterior = it[5]
            atualizar_tabela_visual()
        except ValueError: pass

    def adicionar_item():
        """Lógica para adicionar novo item e limpar campos."""
        try:
            s_ini = float(ent_smp.get() or 0) - float(ent_des_ant.get() or 0)
            s_acumulado = lista_itens[-1][5] if lista_itens else s_ini
            v_ent, v_sai = float(e_ent.get() or 0), float(e_sai.get() or 0)
            
            novo_saldo = s_acumulado + v_ent - v_sai
            lista_itens.append([len(lista_itens)+1, e_data.get(), e_desc.get(), v_ent, v_sai, novo_saldo])
            
            # LIMPEZA DOS CAMPOS
            e_data.delete(0, 'end'); e_desc.delete(0, 'end'); e_ent.delete(0, 'end'); e_sai.delete(0, 'end')
            
            atualizar_tabela_visual()
        except ValueError:
            messagebox.showerror("Erro", "Por favor, insira números válidos nos valores.")

    ctk.CTkButton(palco_conteudo, text="Adicionar Item à Lista", command=adicionar_item).pack(pady=5)

    # 4. FINALIZAÇÃO E ASSINATURA
    frame_fim = ctk.CTkFrame(palco_conteudo)
    frame_fim.pack(fill="x", padx=20, pady=10)
    e_nome = ctk.CTkEntry(frame_fim, placeholder_text="Nome do Tesoureiro/a", width=200); e_nome.grid(row=0, column=0, padx=5)
    e_sexo = ctk.CTkEntry(frame_fim, placeholder_text="M/F", width=60); e_sexo.grid(row=0, column=1, padx=5)

    def finalizar_e_gerar():
        """Chama as funções do tabela_auto para salvar o Excel."""
        if not lista_itens:
            messagebox.showwarning("Aviso", "A lista está vazia.")
            return
        try:
            recalcular_tudo()
            smp_val = float(ent_smp.get() or 0)
            des_val = float(ent_des_ant.get() or 0)
            
            # Chamar funções do tabela_auto.py
            header(ent_ano.get(), ent_mes.get().upper(), ent_moeda.get().upper(), float(smp_val - des_val))
            body(lista_itens, ent_mes.get().upper(), ent_ano.get())
            footer(lista_itens, ent_mes.get().upper(), ent_ano.get(), e_sexo.get(), e_nome.get(), 
                   lista_itens[-1][5], sum(x[4] for x in lista_itens), smp_val, des_val, sum(x[3] for x in lista_itens))
            
            messagebox.showinfo("Sucesso", f"O ficheiro 'Custo_Entrada&Saida.xlsx' foi gerado para {ent_mes.get()}!")
        except Exception as err:
            messagebox.showerror("Erro Crítico", f"Falha ao gerar Excel: {err}")

    btn_gerar = ctk.CTkButton(palco_conteudo, text="GERAR EXCEL FINAL", fg_color="green", 
                               hover_color="darkgreen", font=("Arial", 14, "bold"), command=finalizar_e_gerar)
    btn_gerar.pack(pady=10)
    
    atualizar_tabela_visual()

# --- JANELA PRINCIPAL ---
app = ctk.CTk()
app.title("Sistema de Automatização de Tabelas - Cristiano Glória")
app.geometry("1000x800")

app.grid_columnconfigure(1, weight=1)
app.grid_rowconfigure(0, weight=1)

# Sidebar (Menu Lateral)
sidebar = ctk.CTkFrame(app, width=200, corner_radius=0)
sidebar.grid(row=0, column=0, sticky="nsew")

ctk.CTkLabel(sidebar, text="MENU", font=("Arial", 20, "bold")).pack(pady=30)
ctk.CTkButton(sidebar, text="Nova Tabela", command=mostrar_novaTabela).pack(pady=10, padx=15)
ctk.CTkButton(sidebar, text="Sobre o App", command=mostrar_sobre).pack(pady=10, padx=15)

# Palco de Conteúdo (Área Dinâmica)
palco_conteudo = ctk.CTkFrame(app, corner_radius=15)
palco_conteudo.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")

# Inicia na tela de Nova Tabela por padrão
mostrar_novaTabela()

app.mainloop()