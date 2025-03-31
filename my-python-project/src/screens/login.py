import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from PIL import Image, ImageTk
from src.database.connection import get_connection
from src.screens.Menu import MenuScreen

class LoginScreen:
    def __init__(self, root):
        self.root = root
        self.root.title("Login - Empral Piracicaba")
        # Configura o estilo para esta janela
        self.configurar_estilo()
        window_width = 400
        window_height = 550
        self.center_window(window_width, window_height)
        self.root.configure(bg='#f0f0f0')
        
        # Frame principal
        self.frame = ttk.Frame(self.root, padding=30)
        self.frame.pack(expand=True, fill='both')
        
        # Exibição da logo
        try:
            logo_path = r"L:\Público\Brian\Logo empral\logo.png"
            image = Image.open(logo_path)
            image = image.resize((200, 200), Image.LANCZOS)
            self.logo = ImageTk.PhotoImage(image)
            lbl_logo = ttk.Label(self.frame, image=self.logo, background='#f0f0f0')
            lbl_logo.pack(pady=(0, 30))
        except Exception as e:
            print("Erro ao carregar a logo:", e)
        
        # Campo de usuário
        lbl_user = ttk.Label(self.frame, text="Usuário:")
        lbl_user.pack(pady=(10, 5), anchor='w')
        self.entry_user = ttk.Entry(self.frame, width=30, font=('Segoe UI', 12))
        self.entry_user.pack(pady=5, fill='x')
        
        # Campo de senha
        lbl_pass = ttk.Label(self.frame, text="Senha:")
        lbl_pass.pack(pady=(15, 5), anchor='w')
        self.entry_pass = ttk.Entry(self.frame, show="*", width=30, font=('Segoe UI', 12))
        self.entry_pass.pack(pady=5, fill='x')
        
        # Botão de login
        btn = ttk.Button(self.frame, text="Acessar", command=self.authenticate, padding=10)
        btn.pack(pady=30, fill='x')
        
        # Link "Esqueceu sua senha?"
        lbl_forgot_pass = ttk.Label(self.frame, text="Esqueceu sua senha?", cursor="hand2",
                                    foreground="#4CAF50", background="#f0f0f0",
                                    font=('Segoe UI', 10, 'underline'))
        lbl_forgot_pass.pack(pady=(10, 0), anchor='e')
        lbl_forgot_pass.bind("<Button-1>", lambda e: messagebox.showinfo("Esqueceu a Senha", "Funcionalidade 'Esqueceu a senha' não implementada."))
    
    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = int((screen_width / 2) - (width / 2))
        y = int((screen_height / 2) - (height / 2))
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        
    def configurar_estilo(self):
        style = ttk.Style(self.root)
        style.theme_use('clam')
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", background="#f0f0f0", foreground="black", font=("Segoe UI", 12))
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground="#212121", borderwidth=1, relief="solid", font=("Segoe UI", 12))
        style.configure("TButton", font=("Segoe UI", 12), background="#4CAF50", foreground="white")
        style.map("TButton", background=[("active", "#45a049")])
        style.configure("TCombobox", fieldbackground="#FFFFFF", foreground="#212121", borderwidth=1, relief="flat", font=("Segoe UI", 12))
        
    def authenticate(self):
        username = self.entry_user.get()
        password = self.entry_pass.get()
        
        connection = get_connection()
        if connection is None:
            messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados!")
            return
        
        cursor = connection.cursor()
        query = "SELECT * FROM funcionario WHERE usuario = %s AND senha = %s"
        cursor.execute(query, (username, password))
        result = cursor.fetchone()
        connection.close()
        
        if result:
            messagebox.showinfo("Login", "Bem-vindo à Empral Piracicaba!")
            self.root.withdraw()
            dashboard_window = tk.Toplevel(self.root)
            style = ttk.Style(dashboard_window)
            style.theme_use('clam')
            MenuScreen(dashboard_window)
            dashboard_window.protocol("WM_DELETE_WINDOW", self.root.destroy)
        else:
            messagebox.showerror("Login", "Credenciais inválidas!")
            
if __name__ == "__main__":
    root = tk.Tk()
    app = LoginScreen(root)
    root.mainloop()