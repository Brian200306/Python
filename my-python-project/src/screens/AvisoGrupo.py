import tkinter as tk
from tkinter import ttk

class AvisoGrupoScreen:
    def __init__(self, root):
        self.root = root
        self.root.title("Aviso - Grupo")
        self.root.geometry("800x600")
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(expand=True, fill='both')
        label = ttk.Label(frame, text="Tela de Aviso - Grupo", font=("Segoe UI", 16))
        label.pack()

if __name__ == "__main__":
    root = tk.Tk()
    AvisoGrupoScreen(root)
    root.mainloop()