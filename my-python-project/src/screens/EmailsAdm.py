import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import tkinter as tk
from tkinter import ttk, messagebox
from database.connection import get_connection
from datetime import datetime
import glob

class EmailsAdm:
    def __init__(self, root):
        self.root = root
        self.root.title("Cadastro de Emails ADM")
        self.root.geometry("1500x950")

        # Variáveis de controle para edição
        self.editing = False
        self.current_codigo = None

        # Definindo paleta de cores
        self.verde_escuro = "#0B6E4F"
        self.verde_claro = "#6FB98F"
        self.verde_agua = "#E7F2F8"
        self.branco = "#FFFFFF"
        self.preto = "#212121"
        self.cinza_claro = "#F5F5F5"
        self.cinza_medio = "#E0E0E0"
        self.vermelho = "#FF5252"
        self.azul = "#2196F3"
        self.amarelo = "#FFC107"

        self.configurar_estilo()

        main_frame = ttk.Frame(self.root, style="Main.TFrame")
        main_frame.pack(expand=True, fill="both", padx=15, pady=10)

        # Cabeçalho com design melhorado
        header_frame = ttk.Frame(main_frame, style="Header.TFrame")
        header_frame.pack(fill="x", pady=(0, 15))
        ttk.Label(header_frame, text="Cadastro de Emails ADM", style="Header.TLabel").pack(pady=10)
        
        # Substituindo Combobox por botões de alternância modernos para escolher o tipo de e-mail
        tipo_frame = ttk.Frame(main_frame, style="Content.TFrame")
        tipo_frame.pack(fill="x", padx=5, pady=10)
        
        ttk.Label(tipo_frame, text="Tipo de E-mail:", style="Bold.TLabel").pack(side=tk.LEFT, padx=10)
        
        # Frame para os botões de alternância
        toggle_frame = ttk.Frame(tipo_frame, style="Toggle.TFrame")
        toggle_frame.pack(side=tk.LEFT, padx=5)
        
        self.email_type = tk.StringVar(value="Enviados")
        
        # Criando botões de alternância estilizados
        self.btn_enviados = ttk.Button(toggle_frame, text="Enviados", style="Toggle.TButton.Active" if self.email_type.get() == "Enviados" else "Toggle.TButton", 
                                       command=lambda: self.toggle_email_type("Enviados"))
        self.btn_enviados.pack(side=tk.LEFT)
        
        self.btn_recebidos = ttk.Button(toggle_frame, text="Recebidos", style="Toggle.TButton" if self.email_type.get() == "Enviados" else "Toggle.TButton.Active", 
                                        command=lambda: self.toggle_email_type("Recebidos"))
        self.btn_recebidos.pack(side=tk.LEFT)

        # Frame para botões de ação com ícones e cores modernas
        buttons_action_frame = ttk.Frame(main_frame, style="Content.TFrame")
        buttons_action_frame.pack(fill="x", pady=10)
        
        # Botões de ação com estilos modernos e específicos
        ttk.Button(buttons_action_frame, text="➕ Incluir", style="Action.TButton.Add", 
                   command=self.novo_cadastro).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(buttons_action_frame, text="✏️ Alterar", style="Action.TButton.Edit", 
                   command=self.alterar_cadastro).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(buttons_action_frame, text="❌ Excluir", style="Action.TButton.Delete", 
                   command=self.excluir_cadastro).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(buttons_action_frame, text="↩️ Cancelar", style="Action.TButton.Cancel", 
                   command=self.cancelar_cadastro).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(buttons_action_frame, text="💾 Salvar", style="Action.TButton.Save", 
                   command=self.salvar_email).pack(side=tk.LEFT, padx=10, pady=5)

        # Frame para campos de cadastro ("Dados do E-mail")
        cadastro_frame = ttk.LabelFrame(main_frame, text="Dados do E-mail", style="Verde.TLabelframe")
        cadastro_frame.pack(expand=True, fill="both", padx=5, pady=10)
        # Linha 0: Código e Assunto
        ttk.Label(cadastro_frame, text="Código:", style="Bold.TLabel").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.codigo_var = tk.StringVar()
        self.entry_codigo = ttk.Entry(cadastro_frame, textvariable=self.codigo_var, state='disabled', style="TEntry")
        self.entry_codigo.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Label(cadastro_frame, text="Assunto:", style="Bold.TLabel").grid(row=0, column=2, sticky="e", padx=5, pady=5)
        self.assunto_var = tk.StringVar()
        self.entry_assunto = ttk.Entry(cadastro_frame, textvariable=self.assunto_var, style="TEntry")
        self.entry_assunto.grid(row=0, column=3, columnspan=3, sticky="ew", padx=5, pady=5)

        # Linha 1: Data
        ttk.Label(cadastro_frame, text="Data (DD/MM/AAAA):", style="Bold.TLabel").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.data_var = tk.StringVar()
        self.entry_data = ttk.Entry(cadastro_frame, textvariable=self.data_var, style="TEntry")
        self.entry_data.grid(row=1, column=1, sticky="ew", padx=5, pady=5)

        # Linha 2: Código do Cliente e Cliente
        ttk.Label(cadastro_frame, text="Código do Cliente:", style="Bold.TLabel").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.cod_cliente_var = tk.StringVar()
        self.combo_cod_cliente = ttk.Combobox(cadastro_frame, textvariable=self.cod_cliente_var, style="TCombobox")
        self.combo_cod_cliente.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
        self.combo_cod_cliente.bind("<<ComboboxSelected>>", self.preencher_cliente_nome)
        ttk.Label(cadastro_frame, text="Cliente:", style="Bold.TLabel").grid(row=2, column=2, sticky="e", padx=5, pady=5)
        self.cliente_var = tk.StringVar()
        self.entry_cliente = ttk.Entry(cadastro_frame, textvariable=self.cliente_var, style="TEntry")
        self.entry_cliente.grid(row=2, column=3, columnspan=3, sticky="ew", padx=5, pady=5)

        # Linha 3: Para
        ttk.Label(cadastro_frame, text="Para:", style="Bold.TLabel").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.para_var = tk.StringVar()
        self.entry_para = ttk.Entry(cadastro_frame, textvariable=self.para_var, style="TEntry")
        self.entry_para.grid(row=3, column=1, columnspan=5, sticky="ew", padx=5, pady=5)

        # Linha 4: De (somente para recebidos)
        ttk.Label(cadastro_frame, text="De:", style="Bold.TLabel").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        self.de_var = tk.StringVar()
        self.entry_de = ttk.Entry(cadastro_frame, textvariable=self.de_var, style="TEntry", state="disabled")
        self.entry_de.grid(row=4, column=1, columnspan=5, sticky="ew", padx=5, pady=5)

        # Linha 5: Atenciosamente
        ttk.Label(cadastro_frame, text="Atenciosamente:", style="Bold.TLabel").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        self.atenciosamente_var = tk.StringVar()
        self.combo_atenciosamente = ttk.Combobox(cadastro_frame, textvariable=self.atenciosamente_var,
                                                 values=["SRA. TÂNIA CARDOSO", "ENG BRUNELLI"], style="TCombobox")
        self.combo_atenciosamente.grid(row=5, column=1, sticky="ew", padx=5, pady=5)

        # Linha 6: Observação
        ttk.Label(cadastro_frame, text="Observação:", style="Bold.TLabel").grid(row=6, column=0, sticky="ne", padx=5, pady=5)
        self.obs_var = tk.StringVar()
        self.text_obs = tk.Text(cadastro_frame, height=3)
        self.text_obs.grid(row=6, column=1, columnspan=5, sticky="ew", padx=5, pady=5)
        self.text_obs.config(background=self.branco, foreground=self.preto, borderwidth=1, relief="solid")

        # Linha 7: CC
        ttk.Label(cadastro_frame, text="CC:", style="Bold.TLabel").grid(row=7, column=0, sticky="e", padx=5, pady=5)
        self.cc_var = tk.StringVar()
        self.entry_cc = ttk.Entry(cadastro_frame, textvariable=self.cc_var, style="TEntry")
        self.entry_cc.grid(row=7, column=1, columnspan=5, sticky="ew", padx=5, pady=5)

        for i in range(6):
            cadastro_frame.columnconfigure(i, weight=1)

        # Filtro com aparência modernizada
        filtro_frame = ttk.Frame(main_frame, style="Filter.TFrame")
        filtro_frame.pack(fill="x", padx=5, pady=(15, 5))
        
        ttk.Label(filtro_frame, text="🔍", style="Icon.TLabel").pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(filtro_frame, text="Filtro:", style="Bold.TLabel").pack(side=tk.LEFT, padx=5)
        
        self.filtro_var = tk.StringVar()
        filtro_entry = ttk.Entry(filtro_frame, textvariable=self.filtro_var, style="Modern.TEntry", width=30)
        filtro_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(filtro_frame, text="Pesquisar por:", style="Bold.TLabel").pack(side=tk.LEFT, padx=5)
        
        self.filtro_categoria_var = tk.StringVar()
        self.filtro_categoria_var.set("Todos")
        
        # Combobox com estilo moderno para categorias de busca
        self.combo_filtro_categoria = ttk.Combobox(filtro_frame, textvariable=self.filtro_categoria_var,
                                                   values=["Todos", "Código", "Assunto", "Cliente"],
                                                   state="readonly", width=10, style="Modern.TCombobox")
        self.combo_filtro_categoria.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(filtro_frame, text="Buscar", style="Action.TButton.Search", 
                   command=self.filtrar_registros).pack(side=tk.LEFT, padx=10)

        # Frame para exibição dos registros com visual moderno
        self.list_frame = ttk.LabelFrame(main_frame, text="Registros", style="Modern.TLabelframe")
        self.list_frame.pack(expand=True, fill="both", padx=5, pady=(10, 15))
        
        columns = ("Código", "Assunto", "Data", "Cliente")
        self.tree = ttk.Treeview(self.list_frame, columns=columns, show="headings", height=6)
        for col in self.tree["columns"]:
            if col == "Código":
                self.tree.heading(col, text=col, command=lambda c=col: self.treeview_sort_column(self.tree, c, False))
            else:
                self.tree.heading(col, text=col)
            self.tree.column(col, width=150)
        self.tree.pack(side=tk.LEFT, fill="both", expand=True)
        scrollbar = ttk.Scrollbar(self.list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        # Bind para seleção e duplo clique
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        self.tree.bind("<Double-1>", self.abrir_email)

        # Barra de status
        status_bar = ttk.Frame(main_frame, style="Footer.TFrame")
        status_bar.pack(fill="x", side=tk.BOTTOM)
        self.status_label = ttk.Label(status_bar, text="Pronto", style="Footer.TLabel")
        self.status_label.pack(pady=5)

        # Inicializações
        self.update_fields_state()  # Atualiza estados de acordo com tipo
        self.update_filtro_categoria()
        self.carregar_clientes()
        self.novo_cadastro()
        self.carregar_registros()

    def configurar_estilo(self):
        style = ttk.Style(self.root)
        style.theme_use('clam')
        
        # Copia o layout base do TButton para os estilos personalizados
        base_layout = style.layout("TButton")
        style.layout("Toggle.TButton", base_layout)
        style.layout("Toggle.TButton.Active", base_layout)
        
        # Definindo layout para os Action.TButton
        style.layout("Action.TButton.Add", base_layout)
        style.layout("Action.TButton.Edit", base_layout)
        style.layout("Action.TButton.Delete", base_layout)
        style.layout("Action.TButton.Cancel", base_layout)
        style.layout("Action.TButton.Save", base_layout)
        style.layout("Action.TButton.Search", base_layout)
        
        # Estilos base
        style.configure("Main.TFrame", background=self.branco)
        style.configure("Content.TFrame", background=self.verde_agua)
        style.configure("Filter.TFrame", background=self.verde_agua, relief="groove", borderwidth=1, padding=10)
        style.configure("Header.TFrame", background=self.verde_escuro)
        style.configure("Footer.TFrame", background=self.verde_escuro)
        
        # Rótulos
        style.configure("TLabel", background=self.verde_agua, foreground=self.preto)
        style.configure("Bold.TLabel", background=self.verde_agua, foreground=self.preto, font=("Segoe UI", 9, "bold"))
        style.configure("Header.TLabel", background=self.verde_escuro, foreground=self.branco, font=("Segoe UI", 16, "bold"))
        style.configure("Footer.TLabel", background=self.verde_escuro, foreground=self.branco, font=("Segoe UI", 9))
        style.configure("Icon.TLabel", background=self.verde_agua, font=("Segoe UI", 12))
        
        # Entradas e campos
        style.configure("TEntry", fieldbackground=self.branco, foreground=self.preto, borderwidth=1)
        style.configure("Modern.TEntry", fieldbackground=self.branco, foreground=self.preto, 
                        borderwidth=1, relief="flat", padding=5)
        
        # Combobox moderno
        style.configure("TCombobox", fieldbackground=self.branco, foreground=self.preto, 
                        borderwidth=1, relief="flat", arrowsize=12)
        style.configure("Modern.TCombobox", padding=5, relief="flat", borderwidth=0)
        style.map("Modern.TCombobox", fieldbackground=[("readonly", self.branco)])
        style.map("Modern.TCombobox", selectbackground=[("readonly", self.verde_claro)])
        style.map("Modern.TCombobox", selectforeground=[("readonly", self.preto)])
        
        # Área de texto
        style.configure("TText", background=self.branco, foreground=self.preto, borderwidth=1, relief="flat")
        
        # Botões toggle para seleção de tipo
        style.configure("Toggle.TButton", background=self.branco, foreground=self.preto, 
                        relief="flat", padding=8, font=("Segoe UI", 9))
        style.configure("Toggle.TButton.Active", background=self.verde_escuro, foreground=self.branco, 
                        relief="flat", padding=8, font=("Segoe UI", 9, "bold"))
        style.map("Toggle.TButton", 
                  background=[("active", self.cinza_medio)],
                  foreground=[("active", self.preto)])
        style.map("Toggle.TButton.Active", 
                  background=[("active", self.verde_claro)],
                  foreground=[("active", self.branco)])
        
        # Botões de ação coloridos
        style.configure("Action.TButton.Add", background=self.verde_claro, foreground=self.branco, 
                        relief="flat", padding=8, borderwidth=0, font=("Segoe UI", 9))
        style.configure("Action.TButton.Edit", background=self.azul, foreground=self.branco, 
                        relief="flat", padding=8, borderwidth=0, font=("Segoe UI", 9))
        style.configure("Action.TButton.Delete", background=self.vermelho, foreground=self.branco, 
                        relief="flat", padding=8, borderwidth=0, font=("Segoe UI", 9))
        style.configure("Action.TButton.Cancel", background=self.cinza_medio, foreground=self.preto, 
                        relief="flat", padding=8, borderwidth=0, font=("Segoe UI", 9))
        style.configure("Action.TButton.Save", background=self.verde_escuro, foreground=self.branco, 
                        relief="flat", padding=8, borderwidth=0, font=("Segoe UI", 9, "bold"))
        style.configure("Action.TButton.Search", background=self.amarelo, foreground=self.preto, 
                        relief="flat", padding=8, borderwidth=0, font=("Segoe UI", 9))
        
        # Mapeamentos de hover para botões de ação
        style.map("Action.TButton.Add", 
                  background=[("active", "#5DA980")],
                  foreground=[("active", self.branco)])
        style.map("Action.TButton.Edit", 
                  background=[("active", "#1976D2")],
                  foreground=[("active", self.branco)])
        style.map("Action.TButton.Delete", 
                  background=[("active", "#D32F2F")],
                  foreground=[("active", self.branco)])
        style.map("Action.TButton.Cancel", 
                  background=[("active", "#BDBDBD")],
                  foreground=[("active", self.preto)])
        style.map("Action.TButton.Save", 
                  background=[("active", "#095E42")],
                  foreground=[("active", self.branco)])
        style.map("Action.TButton.Search", 
                  background=[("active", "#FFA000")],
                  foreground=[("active", self.preto)])
        
        # Treeview (tabela de registros)
        style.configure("Treeview", 
                      background=self.branco,
                      fieldbackground=self.branco,
                      foreground=self.preto,
                      rowheight=25,
                      borderwidth=0)
        style.configure("Treeview.Heading", 
                      font=("Segoe UI", 9, "bold"),
                      background=self.verde_escuro,
                      foreground=self.branco,
                      relief="flat")
        style.map("Treeview.Heading", 
                background=[("active", self.verde_claro)],
                foreground=[("active", self.branco)])
        style.map("Treeview", 
                background=[("selected", self.verde_claro)],
                foreground=[("selected", self.branco)])
        
        # Frames agrupadores
        style.configure("Modern.TLabelframe", 
                      background=self.verde_agua, 
                      borderwidth=1, 
                      relief="groove")
        style.configure("Modern.TLabelframe.Label", 
                      foreground=self.verde_escuro, 
                      background=self.verde_agua, 
                      font=("Segoe UI", 10, "bold"))
        style.configure("Verde.TLabelframe", 
                      background=self.verde_agua, 
                      borderwidth=1, 
                      relief="groove")
        style.configure("Verde.TLabelframe.Label", 
                      foreground=self.verde_escuro, 
                      background=self.verde_agua, 
                      font=("Segoe UI", 10, "bold"))

    # Adiciona nova função para alternar entre tipos de e-mail
    def toggle_email_type(self, tipo):
        if self.email_type.get() != tipo:
            self.email_type.set(tipo)
            
            # Atualiza estilos dos botões
            if tipo == "Enviados":
                self.btn_enviados.configure(style="Toggle.TButton.Active")
                self.btn_recebidos.configure(style="Toggle.TButton")
            else:
                self.btn_enviados.configure(style="Toggle.TButton")
                self.btn_recebidos.configure(style="Toggle.TButton.Active")
            
            # Atualiza os campos e registros
            self.update_fields_state()
            self.update_filtro_categoria()
            self.carregar_registros()

    def carregar_clientes(self):
        conn = get_connection()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT cliente FROM cliente")
                clientes = cursor.fetchall()
                cliente_nomes = [cliente[0] for cliente in clientes]
                self.combo_cod_cliente['values'] = cliente_nomes
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar clientes: {e}")
            finally:
                conn.close()
        else:
            messagebox.showerror("Erro", "Falha na conexão com o banco de dados.")

    def preencher_cliente_nome(self, event=None):
        nome_cliente_selecionado = self.combo_cod_cliente.get()
        conn = get_connection()
        if conn:
            try:
                cursor = conn.cursor()
                cursor.execute("SELECT codigo FROM cliente WHERE cliente = %s", (nome_cliente_selecionado,))
                result = cursor.fetchone()
                if result:
                    self.cod_cliente_var.set(result[0])
                    self.cliente_var.set(nome_cliente_selecionado)
                else:
                    self.cod_cliente_var.set("")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao buscar código do cliente: {e}")
            finally:
                conn.close()

    def novo_cadastro(self):
        self.limpar_campos()
        self.gerar_codigo_email()
        self.editing = False
        self.current_codigo = None
        self.entry_assunto.focus_set()
        self.status_label.config(text="Novo cadastro iniciado")

    def alterar_cadastro(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Atenção", "Selecione um registro para alterar.")
            return
        self.editing = True
        self.status_label.config(text="Modo de alteração ativado.")

    def excluir_cadastro(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Atenção", "Selecione um registro para excluir.")
            return
        resposta = messagebox.askyesno("Confirmação", "Tem certeza que deseja excluir o registro?")
        if resposta:
            codigo = self.tree.item(selected[0])['values'][0]
            conn = get_connection()
            if conn:
                try:
                    cursor = conn.cursor()
                    tabela = "enviadoadm" if self.email_type.get() == "Enviados" else "recebidoadm"
                    cursor.execute(f"DELETE FROM {tabela} WHERE codigo = %s", (codigo,))
                    conn.commit()
                    messagebox.showinfo("Sucesso", "Registro excluído com sucesso!")
                    self.carregar_registros()
                    self.limpar_campos()
                except Exception as e:
                    conn.rollback()
                    messagebox.showerror("Erro", f"Erro ao excluir registro: {e}")
                finally:
                    conn.close()
            else:
                messagebox.showerror("Erro", "Falha na conexão com o banco de dados.")

    def cancelar_cadastro(self):
        if messagebox.askyesno("Cancelar", "Deseja cancelar o cadastro e limpar os campos?"):
            self.limpar_campos()
            self.status_label.config(text="Cadastro cancelado")

    def limpar_campos(self):
        self.codigo_var.set("")
        self.assunto_var.set("")
        self.data_var.set("")
        self.cod_cliente_var.set("")
        self.cliente_var.set("")
        self.para_var.set("")
        self.atenciosamente_var.set("")
        self.obs_var.set("")
        self.cc_var.set("")
        self.de_var.set("")
        self.text_obs.delete("1.0", tk.END)

    def gerar_codigo_email(self):
        conn = get_connection()
        if conn:
            try:
                cursor = conn.cursor()
                ano_abreviado = datetime.now().strftime('%y')
                tabela = "enviadoadm" if self.email_type.get() == "Enviados" else "recebidoadm"
                sufixo = "E" if tabela == "enviadoadm" else "R"
                prefixo_codigo = f"EM-ADM/{ano_abreviado}/"  # ex: EM-ADM/25/
                prefixo_len = len(prefixo_codigo)
                cursor.execute(f"SELECT MAX(codigo) FROM {tabela} WHERE codigo LIKE %s", (f"{prefixo_codigo}%/{sufixo}",))
                ultimo_codigo = cursor.fetchone()[0]
                if ultimo_codigo:
                    ultimo_sequencial_str = ultimo_codigo[prefixo_len:prefixo_len+4]
                    try:
                        ultimo_sequencial = int(ultimo_sequencial_str)
                        novo_sequencial = ultimo_sequencial + 1
                    except ValueError:
                        novo_sequencial = 1
                else:
                    novo_sequencial = 1
                novo_codigo = f"{prefixo_codigo}{novo_sequencial:04d}/{sufixo}"
                print(f"Código gerado: {novo_codigo}")
                self.codigo_var.set(novo_codigo)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar código: {e}")
                print(f"Erro ao gerar código: {e}")
            finally:
                conn.close()
        else:
            messagebox.showerror("Erro", "Falha na conexão com o banco de dados ao gerar código.")

    def salvar_email(self):
        conn = get_connection()
        if not conn:
            messagebox.showerror("Erro", "Falha ao conectar com o banco de dados!")
            return

        codigo = self.codigo_var.get()
        assunto = self.assunto_var.get()
        data_str = self.data_var.get()
        cod_cli = self.cod_cliente_var.get()
        cliente = self.cliente_var.get()
        para = self.para_var.get()
        de = self.de_var.get()
        atenciosamente = self.atenciosamente_var.get()
        obs = self.text_obs.get("1.0", tk.END).strip()
        cc = self.cc_var.get()

        if not assunto or not data_str or not para:
            messagebox.showerror("Erro", "Assunto, Data e Para são campos obrigatórios.")
            return

        try:
            data_db = datetime.strptime(data_str, '%d/%m/%Y').strftime('%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA.")
            return

        conn = get_connection()
        if conn:
            try:
                cursor = conn.cursor()
                tabela = "enviadoadm" if self.email_type.get() == "Enviados" else "recebidoadm"
                if tabela == "recebidoadm":
                    hora_atual = datetime.now().strftime('%H:%M:%S')
                if self.editing and self.current_codigo:
                    if tabela == "enviadoadm":
                        sql = """
                            UPDATE enviadoadm 
                            SET assunto=%s, data=%s, cod_cli=%s, cliente=%s, para=%s, atenciosamente=%s, obs=%s, cc=%s, cod_emp=%s 
                            WHERE codigo=%s
                        """
                        values = (assunto, data_db, cod_cli, cliente, para, atenciosamente, obs, cc, '1', self.current_codigo)
                    else:
                        sql = """
                            UPDATE recebidoadm 
                            SET assunto=%s, data=%s, cod_cli=%s, cliente=%s, para=%s, de=%s, obs=%s, cod_emp=%s 
                            WHERE codigo=%s
                        """
                        values = (assunto, data_db, cod_cli, cliente, para, de, obs, '1', self.current_codigo)
                    cursor.execute(sql, values)
                    conn.commit()
                    messagebox.showinfo("Sucesso", "E-mail atualizado com sucesso!")
                    self.editing = False
                    self.current_codigo = None
                else:
                    if tabela == "enviadoadm":
                        sql = """
                            INSERT INTO enviadoadm (codigo, assunto, data, cod_cli, cliente, para, atenciosamente, obs, cc, cod_emp)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """
                        values = (codigo, assunto, data_db, cod_cli, cliente, para, atenciosamente, obs, cc, '1')
                    else:
                        sql = """
                            INSERT INTO recebidoadm (codigo, assunto, data, hora, cod_cli, cliente, para, de, obs, cod_emp)
                            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        """
                        values = (codigo, assunto, data_db, hora_atual, cod_cli, cliente, para, de, obs, '1')
                    cursor.execute(sql, values)
                    conn.commit()
                    messagebox.showinfo("Sucesso", "E-mail salvo com sucesso!")
                self.status_label.config(text="Operação realizada com sucesso")
                self.novo_cadastro()
                self.carregar_registros()
            except Exception as e:
                conn.rollback()
                messagebox.showerror("Erro", f"Erro ao salvar e-mail: {e}")
                print(f"Erro ao salvar e-mail: {e}")
                self.status_label.config(text="Erro ao salvar e-mail")
            finally:
                conn.close()
        else:
            messagebox.showerror("Erro", "Falha na conexão com o banco de dados!")

    def carregar_registros(self):
        conn = get_connection()
        if conn:
            try:
                cursor = conn.cursor()
                tabela = "enviadoadm" if self.email_type.get() == "Enviados" else "recebidoadm"
                if tabela == "enviadoadm":
                    sql = "SELECT codigo, assunto, data, cliente FROM enviadoadm ORDER BY codigo DESC"
                    cursor.execute(sql)
                    registros = cursor.fetchall()
                    self.tree["columns"] = ("Código", "Assunto", "Data", "Cliente")
                else:
                    sql = "SELECT codigo, assunto, data, de, cliente FROM recebidoadm ORDER BY codigo DESC"
                    cursor.execute(sql)
                    registros = cursor.fetchall()
                    self.tree["columns"] = ("Código", "Assunto", "Data", "De", "Cliente")
                for col in self.tree["columns"]:
                    if col == "Código":
                        self.tree.heading(col, text=col, command=lambda c=col: self.treeview_sort_column(self.tree, c, False))
                    else:
                        self.tree.heading(col, text=col)
                    self.tree.column(col, width=150)
                self.tree.delete(*self.tree.get_children())
                for reg in registros:
                    try:
                        data_formatada = datetime.strptime(str(reg[2]), '%Y-%m-%d').strftime('%d/%m/%Y')
                    except ValueError:
                        data_formatada = reg[2]
                    if tabela == "enviadoadm":
                        self.tree.insert("", tk.END, values=(reg[0], reg[1], data_formatada, reg[3]))
                    else:
                        self.tree.insert("", tk.END, values=(reg[0], reg[1], data_formatada, reg[3], reg[4]))
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar registros: {e}")
            finally:
                conn.close()
        else:
            messagebox.showerror("Erro", "Falha na conexão com o banco de dados.")

    def filtrar_registros(self):
        filtro = self.filtro_var.get().strip()
        categoria = self.filtro_categoria_var.get()
        conn = get_connection()
        if conn:
            try:
                cursor = conn.cursor()
                tabela = "enviadoadm" if self.email_type.get() == "Enviados" else "recebidoadm"
                if filtro:
                    if tabela == "enviadoadm":
                        if categoria == "Código":
                            sql = "SELECT codigo, assunto, data, cliente FROM enviadoadm WHERE codigo LIKE %s"
                            params = (f"%{filtro}%",)
                        elif categoria == "Assunto":
                            sql = "SELECT codigo, assunto, data, cliente FROM enviadoadm WHERE assunto LIKE %s"
                            params = (f"%{filtro}%",)
                        elif categoria == "Cliente":
                            sql = "SELECT codigo, assunto, data, cliente FROM enviadoadm WHERE cliente LIKE %s"
                            params = (f"%{filtro}%",)
                        else:  # "Todos"
                            sql = "SELECT codigo, assunto, data, cliente FROM enviadoadm WHERE codigo LIKE %s OR cliente LIKE %s OR assunto LIKE %s"
                            params = (f"%{filtro}%", f"%{filtro}%", f"%{filtro}%")
                    else:
                        if categoria == "Código":
                            sql = "SELECT codigo, assunto, data, de, cliente FROM recebidoadm WHERE codigo LIKE %s"
                            params = (f"%{filtro}%",)
                        elif categoria == "Assunto":
                            sql = "SELECT codigo, assunto, data, de, cliente FROM recebidoadm WHERE assunto LIKE %s"
                            params = (f"%{filtro}%",)
                        elif categoria == "Cliente":
                            sql = "SELECT codigo, assunto, data, de, cliente FROM recebidoadm WHERE cliente LIKE %s"
                            params = (f"%{filtro}%",)
                        elif categoria == "De":
                            sql = "SELECT codigo, assunto, data, de, cliente FROM recebidoadm WHERE de LIKE %s"
                            params = (f"%{filtro}%",)
                        else:  # "Todos"
                            sql = "SELECT codigo, assunto, data, de, cliente FROM recebidoadm WHERE codigo LIKE %s OR cliente LIKE %s OR assunto LIKE %s OR de LIKE %s"
                            params = (f"%{filtro}%", f"%{filtro}%", f"%{filtro}%", f"%{filtro}%")
                    cursor.execute(sql, params)
                else:
                    if tabela == "enviadoadm":
                        sql = "SELECT codigo, assunto, data, cliente FROM enviadoadm"
                        cursor.execute(sql)
                    else:
                        sql = "SELECT codigo, assunto, data, de, cliente FROM recebidoadm"
                        cursor.execute(sql)
                registros = cursor.fetchall()
                if tabela == "enviadoadm":
                    self.tree["columns"] = ("Código", "Assunto", "Data", "Cliente")
                else:
                    self.tree["columns"] = ("Código", "Assunto", "Data", "De", "Cliente")
                for col in self.tree["columns"]:
                    if col == "Código":
                        self.tree.heading(col, text=col, command=lambda c=col: self.treeview_sort_column(self.tree, c, False))
                    else:
                        self.tree.heading(col, text=col)
                    self.tree.column(col, width=150)
                self.tree.delete(*self.tree.get_children())
                for reg in registros:
                    try:
                        data_formatada = datetime.strptime(str(reg[2]), '%Y-%m-%d').strftime('%d/%m/%Y')
                    except ValueError:
                        data_formatada = reg[2]
                    if tabela == "enviadoadm":
                        self.tree.insert("", tk.END, values=(reg[0], reg[1], data_formatada, reg[3]))
                    else:
                        self.tree.insert("", tk.END, values=(reg[0], reg[1], data_formatada, reg[3], reg[4]))
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao filtrar registros: {e}")
            finally:
                conn.close()
        else:
            messagebox.showerror("Erro", "Falha na conexão com o banco de dados.")

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])
            codigo = item['values'][0]
            conn = get_connection()
            if conn:
                try:
                    cursor = conn.cursor()
                    tabela = "enviadoadm" if self.email_type.get() == "Enviados" else "recebidoadm"
                    if tabela == "enviadoadm":
                        cursor.execute("""
                            SELECT codigo, assunto, DATE_FORMAT(data, '%d/%m/%Y'), cod_cli, cliente, para, atenciosamente, obs, cc 
                            FROM enviadoadm 
                            WHERE codigo = %s
                        """, (codigo,))
                    else:
                        cursor.execute("""
                            SELECT codigo, assunto, data, cod_cli, cliente, para, de, obs 
                            FROM recebidoadm 
                            WHERE codigo = %s
                        """, (codigo,))
                    registro = cursor.fetchone()
                    if registro:
                        self.codigo_var.set(registro[0])
                        self.assunto_var.set(registro[1])
                        if isinstance(registro[2], datetime):
                            data_formatada = registro[2].strftime('%d/%m/%Y')
                        else:
                            try:
                                data_formatada = datetime.strptime(str(registro[2]), '%Y-%m-%d').strftime('%d/%m/%Y')
                            except ValueError:
                                data_formatada = registro[2]
                        self.data_var.set(data_formatada)
                        self.cod_cliente_var.set(registro[3] if registro[3] else "")
                        self.cliente_var.set(registro[4] if registro[4] else "")
                        self.para_var.set(registro[5] if registro[5] else "")
                        if tabela == "enviadoadm":
                            self.atenciosamente_var.set(registro[6] if registro[6] else "")
                            self.text_obs.delete("1.0", tk.END)
                            self.text_obs.insert("1.0", registro[7] if registro[7] else "")
                            self.cc_var.set(registro[8] if registro[8] else "")
                            self.de_var.set("")
                        else:
                            self.atenciosamente_var.set("")
                            self.text_obs.delete("1.0", tk.END)
                            self.text_obs.insert("1.0", registro[7] if registro[7] else "")
                            self.cc_var.set("")
                            self.de_var.set(registro[6] if registro[6] else "")
                        self.current_codigo = registro[0]
                        self.editing = True
                        self.status_label.config(text="Registro carregado para alteração.")
                    else:
                        messagebox.showerror("Erro", "Registro não encontrado no banco de dados.")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao carregar registro: {e}")
                finally:
                    conn.close()
            else:
                messagebox.showerror("Erro", "Falha na conexão com o banco de dados.")

    def update_fields_state(self):
        if self.email_type.get() == "Recebidos":
            self.entry_de.configure(state="normal")
            self.combo_atenciosamente.configure(state="disabled")
            self.entry_cc.configure(state="disabled")
        else:
            self.entry_de.configure(state="disabled")
            self.combo_atenciosamente.configure(state="readonly")
            self.entry_cc.configure(state="normal")

    def update_filtro_categoria(self):
        if self.email_type.get() == "Recebidos":
            self.combo_filtro_categoria['values'] = ["Todos", "Código", "Assunto", "Cliente", "De"]
        else:
            self.combo_filtro_categoria['values'] = ["Todos", "Código", "Assunto", "Cliente"]
        self.filtro_categoria_var.set("Todos")

    def on_email_type_change(self, event):
        # Esta função será mantida para compatibilidade, mas agora delegamos para toggle_email_type
        tipo = self.email_type.get()
        self.toggle_email_type(tipo)

    def treeview_sort_column(self, tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        try:
            l.sort(key=lambda t: int(t[0]), reverse=reverse)
        except ValueError:
            l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)
        tv.heading(col, command=lambda: self.treeview_sort_column(tv, col, not reverse))

    def abrir_email(self, event):
        selected = self.tree.selection()
        if not selected:
            return
        item = self.tree.item(selected[0])
        codigo = item['values'][0]  # Exemplo: "EM-ADM/25/0282/R"
        data_str = self.data_var.get()  # Assumindo formato "DD/MM/AAAA"
        try:
            dt = datetime.strptime(data_str, '%d/%m/%Y')
        except Exception as ex:
            messagebox.showerror("Erro", f"Data inválida: {ex}")
            return
        ano = dt.strftime('%Y')
        mes_num = dt.month
        meses_pt = {
            1: "JANEIRO",
            2: "FEVEREIRO",
            3: "MARÇO",
            4: "ABRIL",
            5: "MAIO",
            6: "JUNHO",
            7: "JULHO",
            8: "AGOSTO",
            9: "SETEMBRO",
            10: "OUTUBRO",
            11: "NOVEMBRO",
            12: "DEZEMBRO"
        }
        mes_extenso = meses_pt.get(mes_num, "")
        
        # Define a pasta base conforme o tipo de e-mail
        if self.email_type.get() == "Enviados":
            base_folder = r"L:\ADM\EMAIL\ENVIADO"
        else:
            base_folder = r"L:\ADM\EMAIL\RECEBIDO"
        
        # Converter o código para o padrão de nome do arquivo
        # Exemplo: "EM-ADM/25/0282/R" → "EM-ADM-25-0282-R"
        file_name_pattern = codigo.replace("/", "-") + "*"
        folder_path = os.path.join(base_folder, ano, mes_extenso)
        file_path_pattern = os.path.join(folder_path, file_name_pattern)
        
        # Usar glob para buscar o arquivo que começa com o código
        matching_files = glob.glob(file_path_pattern)
        
        if matching_files:
            # Abrir o primeiro arquivo encontrado
            os.startfile(matching_files[0])
        else:
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{file_path_pattern}")

if __name__ == '__main__':
    root = tk.Tk()
    app = EmailsAdm(root)
    root.mainloop()
