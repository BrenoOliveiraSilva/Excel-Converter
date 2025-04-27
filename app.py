import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import random

class Bubble:
    def __init__(self, canvas):
        self.canvas = canvas
        size = random.randint(10, 30)
        x = random.randint(0, 400)
        y = 600 + size
        color = '#90EE90'
        self.shape = canvas.create_oval(x, y, x + size, y + size, fill=color, outline=color)
        self.speed = random.uniform(0.5, 2)
        
    def move(self):
        self.canvas.move(self.shape, 0, -self.speed)
        pos = self.canvas.coords(self.shape)
        if pos[1] < -30:
            x = random.randint(0, 400)
            self.canvas.coords(self.shape, x, 600, x + (pos[2] - pos[0]), 630)
        self.canvas.after(50, self.move)

class ConverterApp:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Conversor CC → Excel")
        self.window.geometry("400x600")
        self.setup_ui()
        
    def convert_txt_to_excel(self):
        path_txt = filedialog.askopenfilename(title="Escolha o arquivo TXT", filetypes=[("Arquivos de texto", "*.txt;*.TXT")])
        if path_txt:
            try:
                colspecs = [(1, 10), (11, 22), (23, 64), (65, 70), (71, 80), (81, 112), (113, 144), (145, 176), (177, 203), (204, 215), (216, 242), (243, 286), (287, 295)]
                df = pd.read_fwf(path_txt, colspecs=colspecs, encoding='latin1')
                # Preserva o caso original do nome do arquivo
                if path_txt.endswith('.TXT'):
                    path_excel = path_txt.replace('.TXT', '.xlsx')
                else:
                    path_excel = path_txt.replace('.txt', '.xlsx')
                df.to_excel(path_excel, index=False)
                messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso! Salvo como: {path_excel}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar o arquivo: {e}")

    def setup_ui(self):
        background_color = "#f5f5f5"
        self.window.configure(bg=background_color)

        # Canvas para as bolhas (agora em tela cheia)
        self.canvas = tk.Canvas(self.window, bg=background_color, width=400, height=600, highlightthickness=0)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)

        # Frame principal com fundo transparente
        frame = tk.Frame(self.window, bg=background_color, highlightthickness=0)
        frame.place(relx=0.5, rely=0.5, anchor="center")

        # Criar bolhas
        self.bubbles = []
        for _ in range(15):
            bubble = Bubble(self.canvas)
            bubble.move()
            self.bubbles.append(bubble)

        # Logo
        try:
            logo_image = Image.open('assets/images/logo.png')
            width = 200
            wpercent = (width / float(logo_image.size[0]))
            hsize = int((float(logo_image.size[1]) * float(wpercent)))
            logo_image = logo_image.resize((width, hsize), Image.Resampling.LANCZOS)
            self.logo = ImageTk.PhotoImage(logo_image)
            logo_label = tk.Label(frame, image=self.logo, bg=background_color)
            logo_label.pack(pady=(20, 30))
        except Exception as e:
            print(f"Erro ao carregar a logo: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar a logo: {e}")

        # Título
        titulo_label = tk.Label(
            frame,
            text="Conversor de Relatórios",
            font=("Helvetica", 16, "bold"),
            bg=background_color,
            fg="#16733b"
        )
        titulo_label.pack(pady=(0, 10))

        # Texto explicativo
        explanation_text = """
Este aplicativo converte relatórios do Centro Cirúrgico
para o formato Excel (.xlsx), facilitando a análise
e manipulação dos dados.

Clique no botão abaixo para selecionar o arquivo
e iniciar a conversão.
"""
        explanation_label = tk.Label(
            frame,
            text=explanation_text,
            font=("Helvetica", 11),
            bg=background_color,
            fg="#555555",
            justify="center",
            wraplength=350
        )
        explanation_label.pack(pady=(0, 30))

        # Botão
        button_style = {
            'bg': '#16733b',
            'fg': 'white',
            'font': ('Helvetica', 12, 'bold'),
            'padx': 25,
            'pady': 12,
            'bd': 0,
            'activebackground': '#1a8847'
        }

        button = tk.Button(
            frame,
            text="Selecionar Arquivo e Converter",
            command=self.convert_txt_to_excel,
            cursor="hand2",
            **button_style
        )
        button.pack(pady=20)

        # Versão
        versao_label = tk.Label(
            frame,
            text="v1.0.0",
            font=("Helvetica", 8),
            bg=background_color,
            fg="#999999"
        )
        versao_label.pack(side="bottom", pady=10)

    def on_closing(self):
        if messagebox.askokcancel("Sair", "Deseja realmente sair?"):
            self.window.quit()
            self.window.destroy()

    def run(self):
        try:
            self.window.mainloop()
        except KeyboardInterrupt:
            self.on_closing()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro inesperado: {e}")
            self.window.destroy()

if __name__ == "__main__":
    app = ConverterApp()
    app.run()