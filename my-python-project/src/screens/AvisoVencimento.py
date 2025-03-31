import tkinter as tk
from tkinter import ttk

class AvisoVencimentoScreen:
    def __init__(self, root):
        self.root = root
        self.root.title("Aviso - Vencimento")
        self.root.geometry("800x600")
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(expand=True, fill='both')
        label = ttk.Label(frame, text="Tela de Aviso - Vencimento", font=("Segoe UI", 16))
        label.pack()

if __name__ == "__main__":
    root = tk.Tk()
    AvisoVencimentoScreen(root)
    root.mainloop()