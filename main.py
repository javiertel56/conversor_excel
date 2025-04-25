import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import os

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("📊 Transformador Excel Contable")
        self.geometry("600x400")
        self.configure(bg="#e8f0fe")
        self.iconbitmap(default='icono.ico') if os.path.exists("icono.ico") else None

        self.crear_widgets()
        self.crear_statusbar()

    def crear_widgets(self):
        # Notebook con tabs
        notebook = ttk.Notebook(self)
        notebook.pack(expand=1, fill='both', padx=10, pady=10)

        # Estilo general
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 12), padding=10)
        style.configure("TLabel", font=("Segoe UI", 14))

        # Pestaña RM
        frame_rm = ttk.Frame(notebook)
        notebook.add(frame_rm, text="💼 RM")

        label_rm = ttk.Label(frame_rm, text="Transformación para RM")
        label_rm.pack(pady=40)

        btn_rm = ttk.Button(frame_rm, text="Ejecutar Transformación RM", command=self.ejecutar_rm)
        btn_rm.pack()

        # Pestaña Tcomunicamos
        frame_tcom = ttk.Frame(notebook)
        notebook.add(frame_tcom, text="🌐 Tcomunicamos")

        label_tcom = ttk.Label(frame_tcom, text="Transformación para Tcomunicamos")
        label_tcom.pack(pady=40)

        btn_tcom = ttk.Button(frame_tcom, text="Ejecutar Transformación Tcomunicamos", command=self.ejecutar_tcomunicamos)
        btn_tcom.pack()

    def crear_statusbar(self):
        self.status = tk.StringVar()
        self.status.set("Listo")
        status_bar = tk.Label(self, textvariable=self.status, relief=tk.SUNKEN, anchor='w', bg="#dce6f2", font=("Segoe UI", 10))
        status_bar.pack(fill='x', side='bottom')

    def ejecutar_rm(self):
        ruta_script = os.path.join('rm', 'index.py')
        self.ejecutar_script(ruta_script, "RM")

    def ejecutar_tcomunicamos(self):
        ruta_script = os.path.join('tcomunicamos', 'indext.py')
        self.ejecutar_script(ruta_script, "Tcomunicamos")

    def ejecutar_script(self, ruta, nombre):
        if os.path.exists(ruta):
            try:
                self.status.set(f"Ejecutando {nombre}...")
                self.update_idletasks()
                subprocess.run(['python', ruta], check=True)
                self.status.set(f"{nombre} ejecutado correctamente.")
                messagebox.showinfo("Éxito", f"Transformación {nombre} completada.")
            except subprocess.CalledProcessError as e:
                self.status.set("Error en la ejecución.")
                messagebox.showerror("Error", f"Ocurrió un error al ejecutar {nombre}.\n{e}")
        else:
            self.status.set("Archivo no encontrado.")
            messagebox.showerror("Error", f"No se encontró el archivo {ruta}")

if __name__ == "__main__":
    app = App()
    app.mainloop()
