import tkinter as tk
from tkinter import ttk
from PIL import Image, ImageTk
from src.screens.EmailConsulta import EmailConsultaScreen
from src.screens.EmailsAdm import EmailsAdm
from src.screens.AvisoDescricao import AvisoDescricaoScreen
from src.screens.AvisoGrupo import AvisoGrupoScreen
from src.screens.AvisoVencimento import AvisoVencimentoScreen

class MenuScreen:
    def __init__(self, root):
        self.root = root
        self.root.title("Menu - Empral Piracicaba") 
        window_width = 1800
        window_height = 200
        self.center_window(window_width, window_height)
        self.root.configure(bg='white')

        # Estilos atualizados: base branco e verde, com padding reduzido
        style = ttk.Style(self.root)
        style.theme_use('clam')
        style.configure('TFrame', background='white')
        style.configure('TLabel', background='white', foreground='#333', font=('Segoe UI', 12))
        style.configure('Header.TLabel', font=('Segoe UI', 20, 'bold'), foreground='#4CAF50', background='white')
        style.configure('Menu.TButton', font=('Segoe UI', 11, 'bold'),
                        background='white', foreground='#4CAF50', relief='flat', padding=5)  # padding reduzido de 10 para 5
        style.map('Menu.TButton',
                  background=[('active', '#4CAF50')],
                  foreground=[('active', 'white')])
        
        self.frame = ttk.Frame(self.root, style='TFrame')
        self.frame.pack(expand=True, fill='both', padx=20, pady=20)

        self.create_header()
        self.create_menu()

    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def create_header(self):
        header_frame = ttk.Frame(self.frame, style='TFrame')
        header_frame.pack(pady=(10, 20), fill='x')

        # Logo (ajuste o caminho se necessário)
        try:
            logo_path = r"L:\Público\Brian\Logo empral\logo.png"
            image = Image.open(logo_path)
            image = image.resize((80, 80), Image.LANCZOS)
            self.logo = ImageTk.PhotoImage(image)
            logo_label = ttk.Label(header_frame, image=self.logo, background='white')
            logo_label.pack(side='left', padx=(0, 10))
        except Exception as e:
            print("Erro ao carregar a logo:", e)

        title_label = ttk.Label(header_frame, text="Menu Principal Empral Piracicaba", style='Header.TLabel')  # Alterado de "Painel de Controle" para "Menu Principal"
        title_label.pack(side='left', fill='x', expand=True)

    def create_menu(self):
        # Criar uma área rolável para o menu principal com menor altura
        menu_container = ttk.Frame(self.frame, style='TFrame')
        menu_container.pack(pady=(0, 5), fill='x')  
        
        canvas = tk.Canvas(menu_container, height=40, bg='white', highlightthickness=0)  
        canvas.pack(side='left', fill='x', expand=True)

        h_scroll = ttk.Scrollbar(menu_container, orient='horizontal', command=canvas.xview)
        h_scroll.pack(side='bottom', fill='x')
        canvas.configure(xscrollcommand=h_scroll.set)

        # Frame interno que conterá os botões do menu
        menu_frame = ttk.Frame(canvas, style='TFrame')
        canvas.create_window((0, 0), window=menu_frame, anchor='nw')

        # Lista de todas as opções do menu principal
        menu_options = [
             "Consulta", "Cadastro", "Catálogo", "Controle", "Engenharia",
            "Preparo", "Relatórios", "Área do Colaborador", "Verifica Equipamento",
            "Procedimentos Empral", "Serviços", "Log Off"
        ]

        # Itera e cria os botões do menu com espaçamento reduzido (padx=2, pady=2)
        for option in menu_options:
            btn = ttk.Button(menu_frame, text=option, style='Menu.TButton')
            btn.pack(side='left', padx=2, pady=2)  # espaçamento reduzido
            if option == "Cadastro":
                admin_menu = tk.Menu(self.root, tearoff=0, bg='white', fg='#4CAF50', 
                                     activebackground='#4CAF50', activeforeground='white')
                admin_options = [
                    "Administrativo", "Agenda", "CC", "Desenho", "Despesas", "Email", "Empresa",
                    "Equipamento", "FEP", "Funcionários", "Impressoras", "ITE",
                    "MTE", "MC", "Ocorrências", "OS", "Orçamento", "Pastas", "PT",
                    "RVT", "RT", "ME", "MEX", "RTE", "SE", "ICE", "SC", "PTB",
                    "PTT", "PE", "Controle de Orçamento Terceiros"
                ]
                for item in admin_options:
                    if item == "Administrativo":
                        admin_submenu = tk.Menu(self.root, tearoff=0, bg='white', fg='#4CAF50', 
                                                activebackground='#4CAF50', activeforeground='white')
                        admin_sub_options = ["aviso de vencimento", "e-mail"]
                        for sub_item in admin_sub_options:
                            if sub_item == "aviso de vencimento":
                                aviso_submenu = tk.Menu(self.root, tearoff=0, bg='white', fg='#4CAF50',
                                                        activebackground='#4CAF50', activeforeground='white')
                                aviso_options = ["descrição", "grupo", "vencimento"]
                                for option_aviso in aviso_options:
                                    if option_aviso == "descrição":
                                        aviso_submenu.add_command(label=option_aviso, command=self.open_aviso_descricao)
                                    elif option_aviso == "grupo":
                                        aviso_submenu.add_command(label=option_aviso, command=self.open_aviso_grupo)
                                    elif option_aviso == "vencimento":
                                        aviso_submenu.add_command(label=option_aviso, command=self.open_aviso_vencimento)
                                admin_submenu.add_cascade(label="aviso de vencimento", menu=aviso_submenu)
                            elif sub_item == "e-mail":
                                admin_submenu.add_command(label="e-mail", command=self.open_email_enviado)
                            else:
                                admin_submenu.add_command(label=sub_item, command=lambda opt=sub_item: self.menu_item_clicked(opt))
                        admin_menu.add_cascade(label="Administrativo", menu=admin_submenu)
                    else:
                        admin_menu.add_command(label=item, command=lambda opt=item: self.menu_item_clicked(opt))
                btn.bind("<Button-1>", lambda event, menu=admin_menu: self.show_menu(menu, event))
            elif option == "Consulta":
                consulta_menu = tk.Menu(self.root, tearoff=0, bg='white', fg='#4CAF50',
                                        activebackground='#4CAF50', activeforeground='white')
                consulta_options = [
                    "OS", "Desenho", "Email", "CC / FEP / MC / ITE / RT / RTE / ICE / SC / PTB / PTT",
                    "Ocorrência", "Pasta", "Remessa", "Controle de Orçamento de Terceiros"
                ]
                for item in consulta_options:
                    consulta_menu.add_command(label=item, command=lambda opt=item: self.menu_item_clicked(opt))
                btn.bind("<Button-1>", lambda event, menu=consulta_menu: self.show_menu(menu, event))
            else:
                btn.config(command=lambda opt=option: self.menu_item_clicked(opt))

        # Atualiza a área rolável com base no conteúdo
        menu_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def show_menu(self, menu, event):
        """Exibe o menu suspenso na posição do mouse."""
        menu.post(event.x_root, event.y_root)

    def menu_item_clicked(self, option):
        """Ação ao clicar em um item do menu ou submenu."""
        print(f"Menu item '{option}' clicked")
        # Implementar lógica para cada item

    def open_email_consulta(self):
        new_window = tk.Toplevel(self.root)
        EmailConsultaScreen(new_window)

    def open_email_enviado(self):
        new_window = tk.Toplevel(self.root)
        EmailsAdm(new_window)

    def open_aviso_descricao(self):
        new_window = tk.Toplevel(self.root)
        AvisoDescricaoScreen(new_window)

    def open_aviso_grupo(self):
        new_window = tk.Toplevel(self.root)
        AvisoGrupoScreen(new_window)

    def open_aviso_vencimento(self):
        new_window = tk.Toplevel(self.root)
        AvisoVencimentoScreen(new_window)

    def logout(self):
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = MenuScreen(root)
    root.mainloop()