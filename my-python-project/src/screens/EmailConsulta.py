import sys
import os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import tkinter as tk
from tkinter import ttk, font, messagebox
from database.connection import get_connection
from datetime import datetime

class EmailConsultaScreen:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta Email ADM")
        self.root.geometry("1000x600")

        # Paleta de cores
        self.verde_escuro = "#0B6E4F"
        self.verde_claro = "#6FB98F"
        self.verde_agua = "#E7F2F8"
        self.branco = "#FFFFFF"
        self.preto = "#212121"

        self.configurar_estilo()

        main_frame = ttk.Frame(self.root, style="Main.TFrame")
        main_frame.pack(expand=True, fill="both")

        header_frame = ttk.Frame(main_frame, style="Header.TFrame")
        header_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(header_frame, text="Sistema de Consulta de E-mails", style="Header.TLabel").pack(pady=10)

        frame = ttk.Frame(main_frame, style="Content.TFrame")
        frame.pack(expand=True, fill="both", padx=15, pady=10)

        radio_frame = ttk.LabelFrame(frame, text="Tipo de E-mail", style="Verde.TLabelframe")
        radio_frame.grid(row=0, column=0, columnspan=8, sticky="ew", pady=(0, 10), padx=5)
        self.email_type = tk.StringVar(value="Enviado")
        ttk.Radiobutton(radio_frame, text="Enviado", variable=self.email_type, value="Enviado", style="TRadiobutton").pack(side=tk.LEFT, padx=20, pady=5)
        ttk.Radiobutton(radio_frame, text="Recebido", variable=self.email_type, value="Recebido", style="TRadiobutton").pack(side=tk.LEFT, padx=20, pady=5)

        search_frame = ttk.LabelFrame(frame, text="Critérios de Pesquisa", style="Verde.TLabelframe")
        search_frame.grid(row=1, column=0, columnspan=8, sticky="ew", pady=(0, 10), padx=5)

        self.search_entries = []

        ttk.Label(search_frame, text="Código Email:", style="Bold.TLabel").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.entry_codigo_email = ttk.Entry(search_frame, style="TEntry")
        self.entry_codigo_email.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        self.search_entries.append(self.entry_codigo_email)

        ttk.Label(search_frame, text="Período de:", style="Bold.TLabel").grid(row=0, column=2, sticky="e", padx=5, pady=5)
        self.entry_periodo_de = ttk.Entry(search_frame, style="TEntry")
        self.entry_periodo_de.grid(row=0, column=3, sticky="ew", padx=5, pady=5)
        self.search_entries.append(self.entry_periodo_de)

        ttk.Label(search_frame, text="Período até:", style="Bold.TLabel").grid(row=0, column=4, sticky="e", padx=5, pady=5)
        self.entry_periodo_ate = ttk.Entry(search_frame, style="TEntry")
        self.entry_periodo_ate.grid(row=0, column=5, sticky="ew", padx=5, pady=5)
        self.search_entries.append(self.entry_periodo_ate)

        ttk.Label(search_frame, text="Assunto:", style="Bold.TLabel").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.entry_assunto = ttk.Entry(search_frame, style="TEntry")
        self.entry_assunto.grid(row=1, column=1, columnspan=7, sticky="ew", padx=5, pady=5)
        self.search_entries.append(self.entry_assunto)

        ttk.Label(search_frame, text="Cliente:", style="Bold.TLabel").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.entry_cliente = ttk.Entry(search_frame, style="TEntry")
        self.entry_cliente.grid(row=2, column=1, columnspan=7, sticky="ew", padx=5, pady=5)
        self.search_entries.append(self.entry_cliente)

        buttons_frame = ttk.Frame(frame, style="Content.TFrame")
        buttons_frame.grid(row=3, column=0, columnspan=8, pady=(5, 10), sticky="ew")

        ttk.Button(buttons_frame, text="Pesquisar", style="TButton", command=self.pesquisar).pack(side=tk.LEFT, padx=10)
        ttk.Button(buttons_frame, text="Limpar", style="TButton", command=self.limpar).pack(side=tk.LEFT, padx=10)

        verifica_frame = ttk.LabelFrame(buttons_frame, text="Verificação de Arquivos", style="Verde.TLabelframe")
        verifica_frame.pack(side=tk.RIGHT, padx=10)
        ttk.Label(verifica_frame, text="Ano:", style="Bold.TLabel").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        self.entry_ano = ttk.Entry(verifica_frame, width=6, style="TEntry")
        self.entry_ano.grid(row=0, column=1, sticky="ew", padx=10, pady=5)
        self.search_entries.append(self.entry_ano)
        ttk.Button(verifica_frame, text="Verificar", style="TButton").grid(row=0, column=2, padx=10, pady=5)

        result_frame = ttk.LabelFrame(frame, text="Resultados", style="Verde.TLabelframe")
        result_frame.grid(row=4, column=0, columnspan=8, sticky="nsew", pady=5, padx=5)

        style = ttk.Style()
        style.configure("Treeview", background=self.branco, fieldbackground=self.branco, foreground=self.preto)
        style.map("Treeview", background=[("selected", self.verde_claro)], foreground=[("selected", self.preto)])

        columns = ("Código", "Data", "Assunto", "Cliente")
        self.tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=7)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)
        self.tree.column("Código", width=70)
        self.tree.column("Data", width=100)
        self.tree.column("Assunto", width=300)
        self.tree.column("Cliente", width=200)

        # Make "Data" column header clickable for sorting
        self.sort_ascending_date = True # Initial sort order for date
        self.tree.heading("Data", text="Data", command=self.on_data_header_click)
        self.data_table = [] # To store table data for sorting

        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for i in range(8):
            search_frame.columnconfigure(i, weight=1)
        frame.columnconfigure(7, weight=1)
        frame.rowconfigure(4, weight=1)

        status_bar = ttk.Frame(main_frame, style="Footer.TFrame")
        status_bar.pack(fill="x", side=tk.BOTTOM)
        ttk.Label(status_bar, text="Sistema de Gestão de E-mails © 2024", style="Footer.TLabel").pack(pady=5)

    def configurar_estilo(self):
        style = ttk.Style()
        style.configure("Main.TFrame", background=self.branco)
        style.configure("Content.TFrame", background=self.verde_agua)
        style.configure("Header.TFrame", background=self.verde_escuro)
        style.configure("Footer.TFrame", background=self.verde_escuro)
        style.configure("TLabel", background=self.verde_agua, foreground=self.preto)
        style.configure("Bold.TLabel", background=self.verde_agua, foreground=self.preto, font=("Segoe UI", 9, "bold"))
        style.configure("Header.TLabel", background=self.verde_escuro, foreground=self.branco, font=("Segoe UI", 14, "bold"))
        style.configure("Footer.TLabel", background=self.verde_escuro, foreground=self.branco, font=("Segoe UI", 9))
        style.configure("TButton", background=self.branco, foreground=self.preto, borderwidth=1)
        style.configure("Verde.TButton", background=self.verde_escuro, foreground=self.branco)
        style.map("Verde.TButton",
                  background=[("active", self.verde_claro), ("pressed", self.verde_escuro)],
                  foreground=[("active", self.preto), ("pressed", self.branco)])
        style.configure("Verde.TLabelframe", background=self.verde_agua)
        style.configure("Verde.TLabelframe.Label", foreground=self.verde_escuro, background=self.verde_agua, font=("Segoe UI", 9, "bold"))
        style.configure("TRadiobutton", background=self.verde_agua, foreground=self.preto)
        style.configure("TEntry", fieldbackground=self.branco, foreground=self.preto)

    def format_date_to_dd_mm_yyyy(self, date_str):
        if not date_str:
            return ""
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            return date_obj.strftime('%d/%m/%Y')
        except ValueError:
            return date_str

    def convert_date_from_dd_mm_yyyy_to_yyyy_mm_dd(self, date_str):
        if not date_str:
            return None
        try:
            date_obj = datetime.strptime(date_str, '%d/%m/%Y')
            return date_obj.strftime('%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Erro de Data", "Formato de data inválido. Use DD/MM/YYYY.")
            return None

    def pesquisar(self):
        conn = get_connection()
        if conn is None:
            messagebox.showerror("Erro", "Falha ao conectar com o banco de dados!")
            return
        try:
            cursor = conn.cursor()
            email_type = self.email_type.get()
            if email_type == "Enviado":
                table_name = "enviadoadm"
            elif email_type == "Recebido":
                table_name = "recebidoadm"
            else:
                messagebox.showinfo("Informação", "Tipo de e-mail não reconhecido.")
                return

            query = f"SELECT codigo, data, assunto, cliente FROM {table_name}"
            filtros = []
            params = []

            codigo_email = self.entry_codigo_email.get().strip()
            if codigo_email:
                filtros.append("codigo LIKE %s")
                params.append(f"%{codigo_email}%")

            periodo_de_str = self.entry_periodo_de.get().strip()
            periodo_de_converted = self.convert_date_from_dd_mm_yyyy_to_yyyy_mm_dd(periodo_de_str)
            if periodo_de_converted:
                filtros.append("data >= %s")
                params.append(periodo_de_converted)
            elif periodo_de_str:
                return

            periodo_ate_str = self.entry_periodo_ate.get().strip()
            periodo_ate_converted = self.convert_date_from_dd_mm_yyyy_to_yyyy_mm_dd(periodo_ate_str)
            if periodo_ate_converted:
                filtros.append("data <= %s")
                params.append(periodo_ate_converted)
            elif periodo_ate_str:
                return

            assunto = self.entry_assunto.get().strip()
            if assunto:
                filtros.append("assunto LIKE %s")
                params.append(f"%{assunto}%")

            cliente = self.entry_cliente.get().strip()
            if cliente:
                filtros.append("cliente LIKE %s")
                params.append(f"%{cliente}%")

            if filtros:
                query += " WHERE " + " AND ".join(filtros)

            print("Executando consulta:", query)
            print("Parâmetros:", params)
            cursor.execute(query, params)
            resultados = cursor.fetchall()
            print("Registros retornados:", len(resultados))

            self.data_table = [] # Clear existing data
            self.tree.delete(*self.tree.get_children()) # Clear treeview

            for row in resultados:
                formatted_date = self.format_date_to_dd_mm_yyyy(str(row[1]))
                self.tree.insert("", tk.END, values=(row[0], formatted_date, row[2], row[3]))
                self.data_table.append(list(row)) # Store data for sorting

            print("Tabela atualizada com sucesso.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante a consulta: {str(e)}")
            print("Erro durante a consulta:", e)
        finally:
            conn.close()

    def limpar(self):
        for entry in self.search_entries:
            entry.delete(0, tk.END)
        self.tree.delete(*self.tree.get_children())
        self.data_table = [] # Clear stored data

    def sort_by_date(self):
        def parse_date_for_sort(row):
            date_str = str(row[1]) # Date is the second column (index 1)
            try:
                return datetime.strptime(date_str, '%Y-%m-%d')
            except ValueError:
                return datetime.min # Put invalid dates at the beginning when ascending, end when descending

        reverse_order = not self.sort_ascending_date
        self.data_table.sort(key=parse_date_for_sort, reverse=reverse_order)
        self.sort_ascending_date = not self.sort_ascending_date # Toggle sort order

    def update_treeview_with_data(self):
        self.tree.delete(*self.tree.get_children()) # Clear current table
        for row_data in self.data_table:
            formatted_date = self.format_date_to_dd_mm_yyyy(str(row_data[1])) # Format date for display
            self.tree.insert("", tk.END, values=(row_data[0], formatted_date, row_data[2], row_data[3]))

    def on_data_header_click(self):
        if self.data_table: # Only sort if there is data
            self.sort_by_date()
            self.update_treeview_with_data()


if __name__ == "__main__":
    root = tk.Tk()
    EmailConsultaScreen(root)
    root.mainloop()