import tkinter as tk
from tkinter import ttk

class AvisoDescricaoScreen:
    def __init__(self, root):
        self.root = root
        self.root.title("Aviso - Descrição")
        self.root.geometry("800x600")
        frame = ttk.Frame(self.root, padding=20)
        frame.pack(expand=True, fill='both')
        label = ttk.Label(frame, text="Tela de Aviso - Descrição", font=("Segoe UI", 16))
        label.pack()

if __name__ == "__main__":
    root = tk.Tk()
    AvisoDescricaoScreen(root)
    root.mainloop()