import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading

from .logger import setup_logger
from .processor import FormProcessor

class FormAutomationApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Automatizador de Formularios PDF")
        self.root.geometry("600x450")
        self.root.minsize(550, 400)
        
        self.logger = setup_logger()
        self.processor = FormProcessor(self.logger)
        
        self.template_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        
        self._setup_ui()
        
    def _setup_ui(self):
        # Estilos
        style = ttk.Style()
        style.configure('TLabel', font=('Helvetica', 10))
        style.configure('TButton', font=('Helvetica', 10))
        style.configure('Header.TLabel', font=('Helvetica', 12, 'bold'))
        
        # Contenedor principal
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        ttk.Label(main_frame, text="Generador de PDFs", style='Header.TLabel').pack(pady=(0, 10))
        
        # --- 1. Plantilla Excel ---
        template_frame = ttk.LabelFrame(main_frame, text="1. Plantilla Base (Excel)", padding="10")
        template_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Entry(template_frame, textvariable=self.template_path_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(template_frame, text="Examinar...", command=self._browse_template).pack(side=tk.RIGHT)
        
        # --- 2. Carpeta de Salida ---
        output_frame = ttk.LabelFrame(main_frame, text="2. Carpeta de Salida", padding="10")
        output_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Entry(output_frame, textvariable=self.output_dir_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        ttk.Button(output_frame, text="Examinar...", command=self._browse_output).pack(side=tk.RIGHT)
        
        # --- Botón Iniciar ---
        self.btn_start = ttk.Button(main_frame, text="Iniciar Proceso", command=self._start_process)
        self.btn_start.pack(pady=(0, 10), ipady=5)
        
        # --- Progreso y Logs ---
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        
        self.status_label = ttk.Label(progress_frame, text="Listo para comenzar.")
        self.status_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Consola de logs
        self.log_text = tk.Text(progress_frame, height=6, state=tk.DISABLED, font=('Consolas', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
    def _log_to_ui(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _browse_template(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar plantilla de Excel",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if filename:
            self.template_path_var.set(filename)

    def _browse_output(self):
        dirname = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if dirname:
            self.output_dir_var.set(dirname)

    def _update_progress(self, current, total):
        percentage = (current / total) * 100
        self.progress_var.set(percentage)
        self.status_label.config(text=f"Procesando... {current}/{total} completados ({percentage:.1f}%)")
        self.root.update_idletasks()

    def _start_process(self):
        template_path = self.template_path_var.get()
        output_dir = self.output_dir_var.get()

        if not template_path or not output_dir:
            messagebox.showwarning("Campos faltantes", "Asegúrese de seleccionar la plantilla y la carpeta de salida.")
            return

        self.btn_start.config(state=tk.DISABLED)
        self.progress_var.set(0)
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.status_label.config(text="Iniciando...")

        # Ejecutar en un hilo separado
        thread = threading.Thread(target=self._run_processing_thread, args=(template_path, output_dir))
        thread.daemon = True
        thread.start()

    def _run_processing_thread(self, template_path, output_dir):
        try:
            self._log_to_ui("Conectando a la base de datos y cargando datos...\n")
            
            df = self.processor.load_data_sql()
            
            self._log_to_ui(f"Datos cargados exitosamente. Registros a procesar: {len(df)}\n")
            
            if len(df) == 0:
                self._log_to_ui("No se encontraron registros.\n")
                self.status_label.config(text="Cero registros encontrados.")
                return

            success_count, total = self.processor.process_records(
                df, 
                template_path, 
                output_dir, 
                progress_callback=self._update_progress,
                log_callback=self._log_to_ui
            )
            
            messagebox.showinfo("Proceso Completado", f"Se han generado {success_count} PDFs correctamente en:\n{output_dir}")
            self.status_label.config(text="Proceso completado exitosamente.")
            
        except Exception as e:
            self._log_to_ui(f"ERROR CRÍTICO: {str(e)}\n")
            messagebox.showerror("Error", f"Ha ocurrido un error durante el proceso:\n{str(e)}")
            self.status_label.config(text="Proceso detenido con errores.")
        finally:
            self.btn_start.config(state=tk.NORMAL)

    def run(self):
        self.root.mainloop()
