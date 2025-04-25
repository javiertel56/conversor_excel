import tkinter as tk
from tkinter import filedialog, messagebox
import os
import subprocess

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Transformador Excel Contable")
        self.root.geometry("700x420")
        self.root.configure(bg='#181a20')
        self.ruta_entrada = ''
        self.ruta_salida = ''
        self.build_ui()

    def build_ui(self):
        # Título
        title = tk.Label(
            self.root,
            text="💼 Transformador Excel - T Comunicamos",
            font=("Segoe UI", 22, "bold"),
            bg="#181a20",
            fg="#f1f2f6"
        )
        title.pack(pady=(18, 8))

        # Marco principal con borde y padding
        main_frame = tk.Frame(self.root, bg="#23272e", bd=0, relief="flat", highlightbackground="#444", highlightthickness=2)
        main_frame.pack(padx=28, pady=8, fill="both", expand=True)

        # Botones y etiquetas
        btn_frame = tk.Frame(main_frame, bg="#23272e")
        btn_frame.pack(pady=18)

        def on_enter(e): e.widget['bg'] = '#3b4252'
        def on_leave(e): e.widget['bg'] = e.widget.default_bg

        btn_cargar = tk.Button(
            btn_frame, text="📂 Cargar archivo Excel", font=("Segoe UI", 13, "bold"),
            bg="#0984e3", fg="#f1f2f6", activebackground="#74b9ff", activeforeground="#23272e",
            width=20, height=1, command=self.cargar_archivo, bd=0, relief="ridge", cursor="hand2"
        )
        btn_cargar.default_bg = "#0984e3"
        btn_cargar.bind("<Enter>", on_enter)
        btn_cargar.bind("<Leave>", on_leave)
        btn_cargar.grid(row=0, column=0, padx=10, pady=8, sticky="ew")

        btn_guardar = tk.Button(
            btn_frame, text="💾 Generar y guardar archivo", font=("Segoe UI", 13, "bold"),
            bg="#00b894", fg="#f1f2f6", activebackground="#55efc4", activeforeground="#23272e",
            width=24, height=1, command=self.guardar_archivo, bd=0, relief="ridge", cursor="hand2"
        )
        btn_guardar.default_bg = "#00b894"
        btn_guardar.bind("<Enter>", on_enter)
        btn_guardar.bind("<Leave>", on_leave)
        btn_guardar.grid(row=0, column=1, padx=10, pady=8, sticky="ew")

        self.btn_abrir = tk.Button(
            main_frame, text="📊 Abrir archivo generado", font=("Segoe UI", 13, "bold"),
            bg="#fdcb6e", fg="#23272e", activebackground="#ffeaa7", activeforeground="#23272e",
            width=28, height=1, command=self.abrir_archivo, bd=0, relief="ridge", cursor="hand2", state=tk.DISABLED
        )
        self.btn_abrir.default_bg = "#fdcb6e"
        self.btn_abrir.bind("<Enter>", on_enter)
        self.btn_abrir.bind("<Leave>", on_leave)
        self.btn_abrir.pack(side="bottom", pady=(22, 0), fill="x", padx=10)

        # Área de mensajes tipo consola
        msg_frame = tk.Frame(main_frame, bg="#181a20", bd=0, relief="flat")
        msg_frame.pack(fill="x", padx=18, pady=(8, 0))

        self.label_archivo = tk.Label(
            msg_frame, text="Archivo cargado: Ninguno",
            font=("Segoe UI", 10, "bold"), bg="#181a20", fg="#b2bec3", anchor="w", wraplength=600, pady=4
        )
        self.label_archivo.pack(fill="x")

        self.label_guardado = tk.Label(
            msg_frame, text="Archivo generado: Ninguno",
            font=("Segoe UI", 10, "bold"), bg="#181a20", fg="#b2bec3", anchor="w", wraplength=600, pady=4
        )
        self.label_guardado.pack(fill="x")

        # Pie de página
        footer = tk.Label(
            self.root, text="© 2025 T Comunicamos", font=("Segoe UI", 9),
            bg="#181a20", fg="#636e72"
        )
        footer.pack(side="bottom", pady=5)

    def cargar_archivo(self):
        ruta = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if ruta:
            self.ruta_entrada = ruta
            self.label_archivo.config(
                text=f"✅ Archivo cargado: {os.path.basename(ruta)}",
                fg="#00e676"
            )
            self.btn_abrir.config(state=tk.DISABLED)
        else:
            self.label_archivo.config(
                text="Archivo cargado: Ninguno",
                fg="#b2bec3"
            )
            self.btn_abrir.config(state=tk.DISABLED)

    def guardar_archivo(self):
        if not self.ruta_entrada:
            self.label_guardado.config(
                text="⚠️ Selecciona un archivo primero.",
                fg="#fdcb6e"
            )
            messagebox.showwarning("⚠️ Selecciona un archivo primero.")
            return
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Archivos Excel","*.xlsx")])
        if ruta:
            try:
                # Aquí deberías llamar a tu función de transformación, por ejemplo:
                # transformar_excel(self.ruta_entrada, ruta)
                self.ruta_salida = ruta
                self.label_guardado.config(
                    text=f"✅ Generado: {os.path.basename(ruta)}",
                    fg="#00e676"
                )
                self.btn_abrir.config(state=tk.NORMAL)
                messagebox.showinfo("✅ Éxito","Archivo creado y estilizado.")
            except Exception as e:
                self.label_guardado.config(
                    text=f"❌ Error: {str(e)}",
                    fg="#ff7675"
                )
                self.btn_abrir.config(state=tk.DISABLED)
                messagebox.showerror("❌ Error", str(e))

    def abrir_archivo(self):
        if self.ruta_salida and os.path.exists(self.ruta_salida):
            subprocess.Popen(['start','',self.ruta_salida], shell=True)
        else:
            messagebox.showwarning("⚠️","Archivo no encontrado.")

if __name__ == '__main__':
    root = tk.Tk()
    App(root)
    root.mainloop()