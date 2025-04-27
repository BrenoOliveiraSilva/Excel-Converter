import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
import math
import time

class LEDEffect:
    def __init__(self, canvas, width, height):
        self.canvas = canvas
        self.width = width
        self.height = height
        self.margin = 20  # Margem para mover o LED para dentro
        self.angle = 0
        self.trail_points = []
        self.max_trail = 15  # Aumentei o comprimento do rastro
        self.led_size = 3  # Diminui um pouco o tamanho do LED
        self.speed = 1  # Diminui a velocidade
        
    def move(self):
        # Calcula a posição atual do LED
        self.angle = (self.angle + self.speed) % 360
        progress = self.angle / 360.0
        
        # Calcula a posição baseada no perímetro do retângulo, considerando a margem
        perimeter = 2 * ((self.width - 2*self.margin) + (self.height - 2*self.margin))
        distance = progress * perimeter
        
        # Determina em qual lado do retângulo o LED está, considerando a margem
        if distance < (self.width - 2*self.margin):  # Topo
            x = distance + self.margin
            y = self.margin
        elif distance < (self.width - 2*self.margin) + (self.height - 2*self.margin):  # Lado direito
            x = self.width - self.margin
            y = (distance - (self.width - 2*self.margin)) + self.margin
        elif distance < 2*(self.width - 2*self.margin) + (self.height - 2*self.margin):  # Base
            x = (self.width - self.margin) - (distance - ((self.width - 2*self.margin) + (self.height - 2*self.margin)))
            y = self.height - self.margin
        else:  # Lado esquerdo
            x = self.margin
            y = (self.height - self.margin) - (distance - (2*(self.width - 2*self.margin) + (self.height - 2*self.margin)))
            
        # Adiciona a nova posição à lista de rastros
        self.trail_points.append((x, y))
        if len(self.trail_points) > self.max_trail:
            self.trail_points.pop(0)
            
        # Limpa o canvas
        self.canvas.delete("led")
        
        # Desenha o rastro como uma linha suave
        if len(self.trail_points) > 1:
            for i in range(len(self.trail_points) - 1):
                # Calcula a opacidade para cada segmento
                opacity = int(155 * (i / len(self.trail_points)))  # Reduzido para um brilho mais suave
                color = f'#{opacity:02x}ff{opacity:02x}'
                
                # Desenha uma linha entre pontos consecutivos
                x1, y1 = self.trail_points[i]
                x2, y2 = self.trail_points[i + 1]
                
                # Largura da linha diminui gradualmente
                width = 2 * (i / len(self.trail_points))
                
                self.canvas.create_line(
                    x1, y1, x2, y2,
                    fill=color,
                    width=width,
                    smooth=True,
                    tags="led"
                )
        
        # Desenha o LED principal com um brilho suave
        self.canvas.create_oval(
            x - self.led_size - 2, y - self.led_size - 2,
            x + self.led_size + 2, y + self.led_size + 2,
            fill='#80ff80', outline='#80ff80', tags="led"  # Cor mais clara para o brilho externo
        )
        self.canvas.create_oval(
            x - self.led_size, y - self.led_size,
            x + self.led_size, y + self.led_size,
            fill='#00ff00', outline='#00ff00', tags="led"  # LED principal
        )
        
        self.canvas.after(20, self.move)

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
                column_names = ['Registro', 'Data', 'Paciente', 'Idade', 'Sexo', 'Cidade', 'Cirurgião', 
                              'Auxiliar', 'Anestesista', 'Anestesia', 'Convênio', 'Cirurgia', 'Porte']
                
                df = pd.read_fwf(path_txt, colspecs=colspecs, encoding='latin1', header=None, names=column_names)
                df = df[df['Sexo'].str.strip().isin(['Masculino', 'Feminino'])]
                
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

        # Canvas para o efeito LED
        self.led_canvas = tk.Canvas(self.window, width=400, height=600, bg=background_color, highlightthickness=0)
        self.led_canvas.place(x=0, y=0)
        
        # Inicia o efeito LED
        self.led_effect = LEDEffect(self.led_canvas, 400, 600)
        self.led_effect.move()

        # Frame principal
        frame = tk.Frame(self.window, bg=background_color)
        frame.place(relx=0.5, rely=0.5, anchor="center")

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