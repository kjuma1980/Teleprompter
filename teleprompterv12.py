import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
import traceback
import time
import re

class TeleprompterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Teleprompter")

        # Detectar el tamaño de la pantalla
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Establecer tamaño de la ventana basado en el tamaño de la pantalla
        self.root.geometry(f"{screen_width}x{screen_height}")
        
        # Crear widget de texto
        self.text_widget = tk.Text(self.root, font=("Arial", int(screen_width // 40)), wrap=tk.WORD, height=4, spacing3=10)
        self.text_widget.pack(expand=True, fill='both')
        self.text_widget.tag_configure("center", justify='center')  # Centrar el texto
        self.text_widget.tag_configure("highlight", background="yellow")  # Resaltar texto en amarillo

        # Control de velocidad
        self.speed_slider = tk.Scale(self.root, from_=0.5, to=5, resolution=0.1, orient='horizontal', label='Velocidad de scroll')
        self.speed_slider.pack(fill='x')

        # Botones para cargar archivos y controlar
        self.load_button = tk.Button(self.root, text="Cargar Guion (TXT o Word)", command=self.load_file)
        self.load_button.pack(side='left')

        self.reset_button = tk.Button(self.root, text="Reset", command=self.reset_text)
        self.reset_button.pack(side='left')

        self.start_button = tk.Button(self.root, text="Iniciar", command=self.start_scroll)
        self.start_button.pack(side='left')

        self.pause_button = tk.Button(self.root, text="Pausar", command=self.pause_scroll)
        self.pause_button.pack(side='left')

        self.restart_button = tk.Button(self.root, text="Reiniciar", command=self.restart_scroll)
        self.restart_button.pack(side='left')

        self.is_running = False
        self.current_segment_index = 0  # Índice para iterar sobre los segmentos
        self.segments = []  # Aquí almacenaremos los segmentos de las oraciones largas y divididas

        # Barra espaciadora para pausar/reanudar
        self.root.bind('<space>', self.toggle_pause)

        # Vincular el scroll del mouse para avanzar o retroceder
        self.root.bind("<MouseWheel>", self.mouse_scroll)

    def reset_text(self):
        """Limpia el texto del teleprompter."""
        self.text_widget.delete(1.0, tk.END)
        self.current_segment_index = 0
        self.segments = []

    def load_file(self):
        """Cargar archivos .txt y .docx con manejo mejorado de excepciones."""
        file_path = filedialog.askopenfilename(filetypes=[("Archivos de texto y Word", "*.txt *.docx")])
        if not file_path:
            messagebox.showwarning("Advertencia", "No se seleccionó ningún archivo.")
            return
        try:
            self.reset_text()
            text = ""
            if file_path.endswith('.txt'):
                with open(file_path, 'r', encoding='utf-8') as file:
                    text = file.read()
            elif file_path.endswith('.docx'):
                text = self.load_word_file(file_path)  # Cargar archivos de Word con función separada
            self.segments = self.split_into_segments(text)  # Dividir el texto en segmentos
            self.display_segments()  # Mostrar los primeros segmentos
        except Exception as e:
            error_message = f"Error al abrir el archivo: {str(e)}\n\nDetalles:\n{traceback.format_exc()}"
            messagebox.showerror("Error", error_message)

    def load_word_file(self, file_path):
        """Cargar y convertir un archivo .docx en texto."""
        try:
            document = Document(file_path)
            text = ""
            for para in document.paragraphs:
                text += para.text + " "
            return text
        except Exception as e:
            raise RuntimeError(f"Error al leer el archivo de Word: {str(e)}")

    def split_into_segments(self, text):
        """Divide el texto en segmentos basados en comas, puntos, y divide oraciones largas en 4 partes."""
        segments = []
        sentences = re.split(r'([,.!?])', text)  # Dividir en fragmentos con puntuación

        for i in range(0, len(sentences), 2):
            segment = sentences[i].strip()
            
            if len(segment.split()) > 9:  # Oración larga sin comas
                words_in_segment = segment.split()
                parts = 4
                chunk_size = len(words_in_segment) // parts
                for j in range(parts):
                    start = j * chunk_size
                    end = (j + 1) * chunk_size if j != parts - 1 else len(words_in_segment)
                    part = " ".join(words_in_segment[start:end])
                    segments.append(part)
            else:
                segments.append(segment)

            if i + 1 < len(sentences):
                segments[-1] += sentences[i + 1]

        return segments

    def display_segments(self):
        """Muestra los segmentos y resalta el primero."""
        self.text_widget.delete(1.0, tk.END)

        # Mostrar máximo 4 segmentos en pantalla simultáneamente
        max_segments_per_display = 4
        segments_to_display = self.segments[self.current_segment_index:self.current_segment_index + max_segments_per_display]

        # Mostrar los segmentos
        segment_text = "\n".join(segments_to_display)
        self.text_widget.insert(tk.END, segment_text)
        self.text_widget.tag_add("center", "1.0", "end")

        # Resaltar el primer renglón
        self.text_widget.tag_add("highlight", "1.0", "2.0")

    def start_scroll(self):
        if not self.segments:
            messagebox.showwarning("Advertencia", "Cargue un guion antes de iniciar.")
            return
        self.is_running = True
        self.scroll_text()

    def pause_scroll(self):
        self.is_running = False

    def restart_scroll(self):
        """Reinicia el scroll desde el principio."""
        self.current_segment_index = 0
        self.start_scroll()

    def toggle_pause(self, event=None):
        """Función que permite iniciar o pausar el scroll con la barra espaciadora."""
        self.is_running = not self.is_running
        if self.is_running:
            self.scroll_text()
        else:
            self.pause_scroll()

    def mouse_scroll(self, event):
        """Permite avanzar o retroceder con el scroll del ratón incluso mientras el texto está corriendo."""
        if event.delta > 0 and self.current_segment_index > 0:  # Retroceder con scroll hacia arriba
            self.current_segment_index -= 1
        elif event.delta < 0 and self.current_segment_index < len(self.segments) - 4:  # Avanzar con scroll hacia abajo
            self.current_segment_index += 1

        self.display_segments()  # Actualizar la pantalla inmediatamente después del movimiento

        # Si el texto está corriendo, reiniciar el desplazamiento automático
        if self.is_running:
            self.pause_scroll()  # Pausa temporal para manejar el scroll manual
            self.start_scroll()  # Reiniciar el scroll automático después del ajuste

    def scroll_text(self):
        """Desplaza el texto hacia arriba progresivamente, manteniendo el resaltado en el primer renglón."""
        max_segments_per_display = 4
        while self.is_running:
            if self.current_segment_index < len(self.segments):
                # Muestra y resalta el primer renglón
                self.display_segments()

                # Avanzar el índice de los segmentos
                self.current_segment_index += 1
                self.root.update()
                time.sleep(1 / self.speed_slider.get())
            else:
                self.is_running = False
                break

# Crear la ventana principal
root = tk.Tk()
app = TeleprompterApp(root)
root.mainloop()
