
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import json
import traceback

from motor_microcurriculos import procesar_pdf


class MicrocurriculosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatizador de Microcurrículos")
        self.root.geometry("720x430")

        self.template_path = None
        self.pdf_paths = []

        title = tk.Label(root, text="Automatizador de Microcurrículos", font=("Arial", 18, "bold"))
        title.pack(pady=15)

        self.template_label = tk.Label(root, text="Plantilla Word: no seleccionada", anchor="w")
        self.template_label.pack(fill="x", padx=25)

        btn_template = tk.Button(root, text="Seleccionar plantilla Word (.docx)", command=self.select_template, height=2)
        btn_template.pack(fill="x", padx=25, pady=6)

        self.pdf_label = tk.Label(root, text="PDF seleccionados: 0", anchor="w")
        self.pdf_label.pack(fill="x", padx=25)

        btn_pdf = tk.Button(root, text="Seleccionar PDF(s)", command=self.select_pdfs, height=2)
        btn_pdf.pack(fill="x", padx=25, pady=6)

        btn_generate = tk.Button(root, text="Generar microcurrículos", command=self.generate, height=2, bg="#2e7d32", fg="white")
        btn_generate.pack(fill="x", padx=25, pady=15)

        self.log = tk.Text(root, height=9)
        self.log.pack(fill="both", expand=True, padx=25, pady=10)

    def log_msg(self, msg):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.root.update_idletasks()

    def select_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        if path:
            self.template_path = Path(path)
            self.template_label.config(text=f"Plantilla Word: {self.template_path.name}")
            self.log_msg(f"Plantilla seleccionada: {self.template_path}")

    def select_pdfs(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if paths:
            self.pdf_paths = [Path(p) for p in paths]
            self.pdf_label.config(text=f"PDF seleccionados: {len(self.pdf_paths)}")
            for p in self.pdf_paths:
                self.log_msg(f"PDF seleccionado: {p.name}")

    def generate(self):
        if not self.template_path:
            messagebox.showerror("Error", "Selecciona primero la plantilla Word.")
            return

        if not self.pdf_paths:
            messagebox.showerror("Error", "Selecciona al menos un PDF.")
            return

        output_dir = filedialog.askdirectory(title="Selecciona carpeta de salida")
        if not output_dir:
            return

        output_dir = Path(output_dir)
        diagnosticos = []

        self.log_msg("Iniciando generación...")

        try:
            for pdf in self.pdf_paths:
                self.log_msg(f"Procesando: {pdf.name}")
                salida, diag = procesar_pdf(self.template_path, pdf, output_dir)
                diagnosticos.append(diag)
                self.log_msg(f"Generado: {salida.name}")

            diag_path = output_dir / "diagnostico_microcurriculos.json"
            diag_path.write_text(json.dumps(diagnosticos, ensure_ascii=False, indent=2), encoding="utf-8")

            self.log_msg("Proceso terminado.")
            messagebox.showinfo("Listo", f"Documentos generados en:\n{output_dir}")

        except Exception as e:
            self.log_msg("ERROR:")
            self.log_msg(str(e))
            self.log_msg(traceback.format_exc())
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = MicrocurriculosApp(root)
    root.mainloop()
