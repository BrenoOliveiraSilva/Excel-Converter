import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd

class ConverterApp:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Conversor CC → Excel")
        self.window.geometry("400x600")
        self.setup_ui()
        
    def convert_txt_to_excel(self):
        path_txt = filedialog.askopenfilename(title="Escolha o arquivo TXT", filetypes=[("Arquivos de texto", "*.txt")])

        if path_txt:
            try:
                colspecs = [(0, 10), (10, 20), (20, 30)]
                df = pd.read_fwf(path_txt, colspecs=colspecs)
                path_excel = path_txt.replace('.txt', '.xlsx')
                df.to_excel(path_excel, index=False)
                messagebox.showinfo("Sucesso", f"Arquivo convertido com sucesso! Salvo como: {path_excel}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar o arquivo: {e}")

    def setup_ui(self):
        # Configuração do fundo
        background_color = "#f5f5f5"
        self.window.configure(bg=background_color)

        # Criando um frame principal
        frame = tk.Frame(self.window, bg=background_color)
        frame.pack(expand=True, fill='both', padx=20, pady=20)

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

        # Título principal
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

        # Botão de conversão estilizado
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

        # Rodapé com versão
        versao_label = tk.Label(
            frame,
            text="v1.0.0",
            font=("Helvetica", 8),
            bg=background_color,
            fg="#999999"
        )
        versao_label.pack(side="bottom", pady=10)

        # Configurar manipulador de fechamento da janela
        self.window.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """Handle window closing event"""
        if messagebox.askokcancel("Sair", "Deseja realmente sair?"):
            self.window.quit()
            self.window.destroy()

    def run(self):
        """Start the application"""
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