import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
import pandas as pd
import math
import time

class DataTableWindow:
    def __init__(self, df):
        self.window = tk.Toplevel()
        self.window.title("Dados Convertidos")
        self.window.geometry("1200x600")
        
        # Criar frame principal
        main_frame = tk.Frame(self.window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Criar frame para a tabela com scrollbar
        table_frame = tk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # Criar scrollbars
        y_scrollbar = ttk.Scrollbar(table_frame)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scrollbar = ttk.Scrollbar(table_frame, orient='horizontal')
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Criar Treeview
        self.tree = ttk.Treeview(table_frame, 
                                yscrollcommand=y_scrollbar.set,
                                xscrollcommand=x_scrollbar.set)
        
        # Configurar scrollbars
        y_scrollbar.config(command=self.tree.yview)
        x_scrollbar.config(command=self.tree.xview)
        
        # Definir colunas
        self.tree['columns'] = list(df.columns)
        self.tree['show'] = 'headings'
        
        # Definir larguras e cabeçalhos das colunas
        column_widths = {
            'Registro': 80,
            'Data': 100,
            'Paciente': 200,
            'Idade': 60,
            'Sexo': 80,
            'Cidade': 150,
            'Cirurgião': 150,
            'Auxiliar': 150,
            'Anestesista': 150,
            'Anestesia': 100,
            'Convênio': 150,
            'Cirurgia': 200,
            'Porte': 60
        }
        
        # Configurar as colunas
        for column in df.columns:
            width = column_widths.get(column, 100)
            self.tree.column(column, width=width, minwidth=50)
            self.tree.heading(column, text=column, anchor=tk.W)
        
        # Inserir dados
        for idx, row in df.iterrows():
            values = ['' if pd.isna(val) or str(val).lower() == 'nan' else str(val) for val in row]
            self.tree.insert("", tk.END, values=values)
        
        # Posicionar a tabela
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Estilizar a tabela
        style = ttk.Style()
        style.theme_use("clam")
        
        style.configure("Treeview",
                       background="#ffffff",
                       foreground="#333333",
                       rowheight=25,
                       fieldbackground="#ffffff")
        
        style.configure("Treeview.Heading",
                       background="#16733b",
                       foreground="white",
                       relief="flat")
        
        style.map("Treeview",
                  background=[("selected", "#16733b")],
                  foreground=[("selected", "white")])
        
        # Frame para os botões
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Estilo comum para os botões
        button_style = {
            'bg': '#16733b',
            'fg': 'white',
            'font': ('Helvetica', 10, 'bold'),
            'padx': 15,
            'pady': 5,
            'cursor': 'hand2'
        }
        
        # Botões
        copy_button = tk.Button(button_frame, text="Copiar Dados",
                              command=self.copy_to_clipboard, **button_style)
        copy_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        export_button = tk.Button(button_frame, text="Exportar para Excel",
                                command=self.export_to_excel, **button_style)
        export_button.pack(side=tk.RIGHT, padx=5)
        
        # Label para feedback
        self.feedback_label = tk.Label(button_frame, text="",
                                     fg="#16733b",
                                     bg=self.window.cget('bg'),
                                     font=('Helvetica', 10))
        self.feedback_label.pack(side=tk.LEFT, padx=5)
        
        self.df = df
        self.center_window()

    def copy_to_clipboard(self):
        try:
            lines = []
            headers = [self.tree.heading(col)['text'] for col in self.tree['columns']]
            lines.append('\t'.join(headers))
            
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                values = ['' if v is None or str(v).lower() == 'nan' else str(v).replace('\t', ' ') for v in values]
                lines.append('\t'.join(values))
            
            clipboard_text = '\n'.join(lines)
            self.window.clipboard_clear()
            self.window.clipboard_append(clipboard_text)
            self.window.update()
            
            self.show_feedback("Dados copiados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao copiar dados: {e}")
    
    def show_feedback(self, message):
        self.feedback_label.config(text=message)
        self.window.after(3000, lambda: self.feedback_label.config(text=""))
    
    def center_window(self):
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')

    def export_to_excel(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Salvar como"
        )
        if file_path:
            data = []
            columns = self.tree['columns']
            
            for item in self.tree.get_children():
                values = self.tree.item(item)['values']
                values = ['' if v is None or str(v).lower() == 'nan' else v for v in values]
                data.append(values)
            
            df_to_save = pd.DataFrame(data, columns=columns)
            df_to_save.to_excel(file_path, index=False)
            self.show_feedback("Arquivo exportado com sucesso!")

class LEDEffect:
    def __init__(self, canvas):
        self.canvas = canvas
        self.margin = 20
        self.angle = 0
        self.trail_points = []
        self.max_trail = 15
        self.led_size = 3
        self.speed = 1
        
    def move(self):
        # Atualiza as dimensões atuais do canvas
        width = self.canvas.winfo_width()
        height = self.canvas.winfo_height()
        
        # Calcula a posição atual do LED
        self.angle = (self.angle + self.speed) % 360
        progress = self.angle / 360.0
        
        # Calcula a posição baseada no perímetro do retângulo
        perimeter = 2 * ((width - 2*self.margin) + (height - 2*self.margin))
        distance = progress * perimeter
        
        # Determina a posição do LED
        if distance < (width - 2*self.margin):  # Topo
            x = distance + self.margin
            y = self.margin
        elif distance < (width - 2*self.margin) + (height - 2*self.margin):  # Lado direito
            x = width - self.margin
            y = (distance - (width - 2*self.margin)) + self.margin
        elif distance < 2*(width - 2*self.margin) + (height - 2*self.margin):  # Base
            x = (width - self.margin) - (distance - ((width - 2*self.margin) + (height - 2*self.margin)))
            y = height - self.margin
        else:  # Lado esquerdo
            x = self.margin
            y = (height - self.margin) - (distance - (2*(width - 2*self.margin) + (height - 2*self.margin)))
            
        # Atualiza o rastro
        self.trail_points.append((x, y))
        if len(self.trail_points) > self.max_trail:
            self.trail_points.pop(0)
            
        # Limpa o canvas
        self.canvas.delete("led")
        
        # Desenha o rastro
        if len(self.trail_points) > 1:
            for i in range(len(self.trail_points) - 1):
                opacity = int(155 * (i / len(self.trail_points)))
                color = f'#{opacity:02x}ff{opacity:02x}'
                
                x1, y1 = self.trail_points[i]
                x2, y2 = self.trail_points[i + 1]
                
                width = 2 * (i / len(self.trail_points))
                
                self.canvas.create_line(
                    x1, y1, x2, y2,
                    fill=color,
                    width=width,
                    smooth=True,
                    tags="led"
                )
        
        # Desenha o LED
        self.canvas.create_oval(
            x - self.led_size - 2, y - self.led_size - 2,
            x + self.led_size + 2, y + self.led_size + 2,
            fill='#80ff80', outline='#80ff80', tags="led"
        )
        self.canvas.create_oval(
            x - self.led_size, y - self.led_size,
            x + self.led_size, y + self.led_size,
            fill='#00ff00', outline='#00ff00', tags="led"
        )
        
        self.canvas.after(20, self.move)

class ConverterApp:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Grand Converter")
        self.window.geometry("500x700")
        self.setup_ui()
        
    def convert_txt_to_excel(self):
        path_txt = filedialog.askopenfilename(
            title="Escolha o arquivo TXT",
            filetypes=[("Arquivos de texto", "*.txt;*.TXT")]
        )
        if path_txt:
            try:
                colspecs = [(1, 10), (11, 22), (23, 64), (65, 70), (71, 80),
                           (81, 112), (113, 144), (145, 176), (177, 203),
                           (204, 215), (216, 242), (243, 286), (287, 295)]
                column_names = ['Registro', 'Data', 'Paciente', 'Idade', 'Sexo',
                              'Cidade', 'Cirurgião', 'Auxiliar', 'Anestesista',
                              'Anestesia', 'Convênio', 'Cirurgia', 'Porte']
                
                df = pd.read_fwf(path_txt, colspecs=colspecs, encoding='latin1',
                               header=None, names=column_names)
                df = df[df['Sexo'].str.strip().isin(['Masculino', 'Feminino'])]
                df = df.fillna('')
                
                DataTableWindow(df)
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar o arquivo: {e}")

    def setup_ui(self):
        background_color = "#f5f5f5"
        self.window.configure(bg=background_color)

        # Canvas para o efeito LED
        self.led_canvas = tk.Canvas(self.window, bg=background_color,
                                  highlightthickness=0)
        self.led_canvas.pack(fill=tk.BOTH, expand=True)
        
        # Inicia o efeito LED
        self.led_effect = LEDEffect(self.led_canvas)
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
            text="v1.0.2",
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
            self.window.protocol("WM_DELETE_WINDOW", self.on_closing)
            self.window.mainloop()
        except KeyboardInterrupt:
            self.on_closing()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro inesperado: {e}")
            self.window.destroy()

if __name__ == "__main__":
    app = ConverterApp()
    app.run()