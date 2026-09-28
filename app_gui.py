"""
Interfaz Gráfica de Usuario (GUI) moderna para el Generador DGII 606.
Permite seleccionar archivos, configurar parámetros y generar los reportes con un clic.
"""

from decimal import Decimal
import os
from pathlib import Path
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from generador_606.config import AppConfig
from generador_606.processor import load_providers, load_transactions, process_records
from generador_606.writers import (
    write_606_csv,
    write_606_excel,
    write_excel_agrupacion,
    write_incidences_report,
)


class AppGUI(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Generador DGII 606 y Compras sin Comprobante")
        self.geometry("820x680")
        self.minsize(750, 600)

        # Configurar colores y estilos
        self.configure(bg="#F4F6F9")
        self.style = ttk.Style(self)
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self._init_variables()
        self._build_ui()

    def _init_variables(self):
        cwd = Path.cwd()
        
        # Detectar archivo de transacciones por defecto
        tx_default = "muestra.xlsx"
        if (cwd / "grande.xlsx").exists():
            tx_default = "grande.xlsx"
        elif not (cwd / "muestra.xlsx").exists():
            # Buscar cualquier xlsx que no sea Proveedores
            xlsx_files = [f.name for f in cwd.glob("*.xlsx") if "proveedor" not in f.name.lower()]
            if xlsx_files:
                tx_default = xlsx_files[0]

        self.var_transacciones = tk.StringVar(value=str(cwd / tx_default) if (cwd / tx_default).exists() else tx_default)
        self.var_proveedores = tk.StringVar(value=str(cwd / "Proveedores.xlsx") if (cwd / "Proveedores.xlsx").exists() else "Proveedores.xlsx")
        self.var_salida = tk.StringVar(value=str(cwd / "salida"))
        self.var_tasa = tk.StringVar(value="10")
        self.var_periodo = tk.StringVar(value="")

    def _build_ui(self):
        # -------------------------------------------------------------
        # 1. Encabezado
        # -------------------------------------------------------------
        header_frame = tk.Frame(self, bg="#1F4E78", height=80)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)

        title_lbl = tk.Label(
            header_frame,
            text="Generador DGII - Formato 606 & Compras sin Comprobante",
            font=("Segoe UI", 15, "bold"),
            bg="#1F4E78",
            fg="#FFFFFF"
        )
        title_lbl.pack(pady=(12, 2))

        sub_lbl = tk.Label(
            header_frame,
            text="Separación automática de comprobantes válidos, retención de ISR y gastos B13",
            font=("Segoe UI", 9),
            bg="#1F4E78",
            fg="#D9E2EC"
        )
        sub_lbl.pack()

        # -------------------------------------------------------------
        # 2. Contenedor de formulario
        # -------------------------------------------------------------
        content = tk.Frame(self, bg="#F4F6F9", padx=20, pady=15)
        content.pack(fill="both", expand=True)

        card_files = tk.LabelFrame(
            content,
            text=" Archivos de Entrada y Salida ",
            font=("Segoe UI", 10, "bold"),
            bg="#FFFFFF",
            fg="#1F4E78",
            padx=15,
            pady=12,
            relief="solid",
            bd=1
        )
        card_files.pack(fill="x", pady=(0, 10))

        # Transacciones
        self._create_file_row(
            card_files,
            row=0,
            label="Transacciones (Excel):",
            variable=self.var_transacciones,
            cmd_browse=self._browse_transacciones
        )

        # Proveedores
        self._create_file_row(
            card_files,
            row=1,
            label="Maestro Proveedores:",
            variable=self.var_proveedores,
            cmd_browse=self._browse_proveedores
        )

        # Carpeta de salida
        self._create_file_row(
            card_files,
            row=2,
            label="Carpeta de Salida:",
            variable=self.var_salida,
            cmd_browse=self._browse_salida,
            is_folder=True
        )

        # Parámetros fiscales
        card_params = tk.LabelFrame(
            content,
            text=" Parámetros de Cálculo ",
            font=("Segoe UI", 10, "bold"),
            bg="#FFFFFF",
            fg="#1F4E78",
            padx=15,
            pady=10,
            relief="solid",
            bd=1
        )
        card_params.pack(fill="x", pady=(0, 10))

        # Tasa de retención
        lbl_tasa = tk.Label(card_params, text="Tasa Retención Renta (%):", font=("Segoe UI", 9), bg="#FFFFFF")
        lbl_tasa.grid(row=0, column=0, sticky="w", pady=4)
        ent_tasa = tk.Entry(card_params, textvariable=self.var_tasa, width=8, font=("Segoe UI", 9))
        ent_tasa.grid(row=0, column=1, sticky="w", padx=10, pady=4)
        lbl_tasa_hint = tk.Label(card_params, text="(Por defecto 10% sobre servicios sin comprobante)", font=("Segoe UI", 8), fg="#718096", bg="#FFFFFF")
        lbl_tasa_hint.grid(row=0, column=2, sticky="w")

        # Período fiscal
        lbl_per = tk.Label(card_params, text="Período Fiscal (opcional):", font=("Segoe UI", 9), bg="#FFFFFF")
        lbl_per.grid(row=1, column=0, sticky="w", pady=4)
        ent_per = tk.Entry(card_params, textvariable=self.var_periodo, width=12, font=("Segoe UI", 9))
        ent_per.grid(row=1, column=1, sticky="w", padx=10, pady=4)
        lbl_per_hint = tk.Label(card_params, text="(Formato AAAAMM, ej. 202608 para auditar fechas)", font=("Segoe UI", 8), fg="#718096", bg="#FFFFFF")
        lbl_per_hint.grid(row=1, column=2, sticky="w")

        # Botones de acción
        btn_frame = tk.Frame(content, bg="#F4F6F9")
        btn_frame.pack(fill="x", pady=(5, 10))

        self.btn_process = tk.Button(
            btn_frame,
            text=" ▶ PROCESAR Y GENERAR ARCHIVOS",
            font=("Segoe UI", 10, "bold"),
            bg="#107C41",  # Verde Excel
            fg="#FFFFFF",
            activebackground="#0E6B37",
            activeforeground="#FFFFFF",
            padx=18,
            pady=8,
            relief="flat",
            cursor="hand2",
            command=self._start_processing_thread
        )
        self.btn_process.pack(side="left", padx=(0, 10))

        self.btn_open_folder = tk.Button(
            btn_frame,
            text=" 📂 Abrir Carpeta de Salida",
            font=("Segoe UI", 9),
            bg="#E2E8F0",
            fg="#2D3748",
            padx=14,
            pady=8,
            relief="flat",
            cursor="hand2",
            command=self._open_output_folder
        )
        self.btn_open_folder.pack(side="left")

        # Área de registro / incidencias
        log_frame = tk.LabelFrame(
            content,
            text=" Registro de Ejecución e Incidencias ",
            font=("Segoe UI", 10, "bold"),
            bg="#FFFFFF",
            fg="#1F4E78",
            padx=10,
            pady=8,
            relief="solid",
            bd=1
        )
        log_frame.pack(fill="both", expand=True)

        self.txt_log = tk.Text(
            log_frame,
            wrap="word",
            font=("Consolas", 9),
            bg="#1E1E1E",
            fg="#D4D4D4",
            insertbackground="#FFFFFF",
            padx=8,
            pady=8
        )
        self.txt_log.pack(side="left", fill="both", expand=True)

        scroll_log = tk.Scrollbar(log_frame, command=self.txt_log.yview)
        scroll_log.pack(side="right", fill="y")
        self.txt_log.config(yscrollcommand=scroll_log.set)

        # Configurar tags de color para el texto
        self.txt_log.tag_config("SUCCESS", foreground="#4EC9B0")
        self.txt_log.tag_config("ERROR", foreground="#F48771")
        self.txt_log.tag_config("WARN", foreground="#DCDCAA")
        self.txt_log.tag_config("INFO", foreground="#9CDCFE")
        self.txt_log.tag_config("TITLE", foreground="#FFFFFF", font=("Consolas", 9, "bold"))

        self._log("Listo para iniciar. Seleccione los archivos y presione 'PROCESAR'.\n", "INFO")

    def _create_file_row(self, parent, row, label, variable, cmd_browse, is_folder=False):
        lbl = tk.Label(parent, text=label, font=("Segoe UI", 9, "bold"), bg="#FFFFFF", width=20, anchor="w")
        lbl.grid(row=row, column=0, sticky="w", pady=4)

        entry = tk.Entry(parent, textvariable=variable, font=("Segoe UI", 9))
        entry.grid(row=row, column=1, sticky="ew", padx=8, pady=4)

        btn = tk.Button(
            parent,
            text="Examinar...",
            font=("Segoe UI", 8),
            bg="#EDF2F7",
            relief="groove",
            command=cmd_browse
        )
        btn.grid(row=row, column=2, padx=4, pady=4)
        parent.columnconfigure(1, weight=1)

    def _browse_transacciones(self):
        f = filedialog.askopenfilename(
            title="Seleccionar archivo de transacciones",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if f:
            self.var_transacciones.set(f)

    def _browse_proveedores(self):
        f = filedialog.askopenfilename(
            title="Seleccionar maestro de proveedores",
            filetypes=[("Archivos de Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        )
        if f:
            self.var_proveedores.set(f)

    def _browse_salida(self):
        d = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if d:
            self.var_salida.set(d)

    def _log(self, text: str, tag: str = "INFO"):
        self.txt_log.insert(tk.END, text + "\n", tag)
        self.txt_log.see(tk.END)

    def _open_output_folder(self):
        out_dir = Path(self.var_salida.get())
        out_dir.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform == "win32":
                os.startfile(str(out_dir))
            else:
                subprocess.Popen(["xdg-open", str(out_dir)])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir la carpeta: {e}")

    def _start_processing_thread(self):
        self.btn_process.config(state="disabled", text=" ⏳ Procesando...")
        self.txt_log.delete("1.0", tk.END)
        threading.Thread(target=self._process_pipeline, daemon=True).start()

    def _process_pipeline(self):
        try:
            tx_path = Path(self.var_transacciones.get().strip())
            pr_path = Path(self.var_proveedores.get().strip())
            out_dir = Path(self.var_salida.get().strip())

            if not tx_path.exists():
                raise FileNotFoundError(f"No existe el archivo de transacciones: {tx_path}")
            if not pr_path.exists():
                raise FileNotFoundError(f"No existe el maestro de proveedores: {pr_path}")

            try:
                tasa_val = float(self.var_tasa.get().strip()) / 100.0
            except ValueError:
                tasa_val = 0.10

            periodo_val = self.var_periodo.get().strip() or None

            self._log("=" * 65, "TITLE")
            self._log("INICIANDO PROCESO DE GENERACIÓN DGII 606", "TITLE")
            self._log("=" * 65, "TITLE")
            self._log(f"Transacciones : {tx_path.name}")
            self._log(f"Proveedores   : {pr_path.name}")
            self._log(f"Tasa Retención: {tasa_val * 100:.1f}%")
            if periodo_val:
                self._log(f"Período Fiscal: {periodo_val}")
            self._log("-" * 65)

            # Cargar configuración
            cfg = AppConfig.load(
                config_path="config.yaml",
                transacciones_path=str(tx_path),
                proveedores_path=str(pr_path),
                salida_dir=str(out_dir),
                tasa_retencion=tasa_val,
                periodo=periodo_val
            )

            # 1. Cargar proveedores
            self._log("Cargando maestro de proveedores...")
            providers = load_providers(cfg.proveedores_path, hoja_name=cfg.hoja_proveedores)
            self._log(f"  ✓ {len(providers):,} proveedores cargados con éxito.", "SUCCESS")

            # 2. Cargar transacciones
            self._log("Cargando transacciones...")
            transactions = load_transactions(cfg.transacciones_path)
            self._log(f"  ✓ {len(transactions):,} transacciones leídas.", "SUCCESS")

            # 3. Procesar
            self._log("Evaluando reglas fiscales y comprobantes...")
            records, incidences = process_records(transactions, providers, cfg)

            rec_606 = [r for r in records if r.is_606]
            rec_agrupados = [r for r in records if not r.is_606 and r.destino == "Agrupado"]
            rec_b13 = [r for r in records if not r.is_606 and r.destino == "B13 Gastos menores"]

            self._log(f"  ✓ Comprobantes válidos para Formato 606 : {len(rec_606)}", "SUCCESS")
            self._log(f"  ✓ Sin comprobante (Agrupados Proveedor) : {len(rec_agrupados)}", "SUCCESS")
            self._log(f"  ✓ Sin comprobante (Gastos Menores B13)  : {len(rec_b13)}", "SUCCESS")

            # 4. Escribir salidas
            out_dir.mkdir(parents=True, exist_ok=True)
            p_606_xlsx = out_dir / (cfg.archivo_606 if cfg.archivo_606.endswith(".xlsx") else "formato_606.xlsx")
            p_606_csv = out_dir / (cfg.archivo_606_csv or "formato_606.csv")
            p_excel = out_dir / cfg.archivo_agrupacion
            p_inc = out_dir / cfg.archivo_incidencias

            self._log(f"\nGenerando {p_606_xlsx.name} (Excel con columnas y formato)...")
            write_606_excel(records, providers, p_606_xlsx, cfg)

            # También generamos el CSV oficial
            write_606_csv(records, providers, p_606_csv, cfg)

            self._log(f"Generando {p_excel.name}...")
            write_excel_agrupacion(records, p_excel, cfg)

            self._log(f"Generando {p_inc.name}...")
            write_incidences_report(incidences, len(transactions), len(rec_606), len(rec_agrupados) + len(rec_b13), p_inc)

            self._log("\n" + "=" * 65, "SUCCESS")
            self._log("¡PROCESO FINALIZADO CON ÉXITO!", "SUCCESS")
            self._log("=" * 65, "SUCCESS")

            errores = [i for i in incidences if i.severidad == "ERROR"]
            advertencias = [i for i in incidences if i.severidad == "ADVERTENCIA"]

            if errores:
                self._log(f"\n[!] Se detectaron {len(errores)} ERRORES:", "ERROR")
                for e in errores:
                    self._log(f"  • {e}", "ERROR")

            if advertencias:
                self._log(f"\n[i] Advertencias ({len(advertencias)}):", "WARN")
                for w in advertencias:
                    self._log(f"  • {w}", "WARN")

            messagebox.showinfo(
                "Proceso Exitoso",
                f"Se generaron los archivos correctamente en:\n{out_dir}\n\n"
                f"• 606: {len(rec_606)} registros\n"
                f"• Agrupados: {len(rec_agrupados)} registros\n"
                f"• Gastos B13: {len(rec_b13)} registros"
            )

        except Exception as e:
            self._log(f"\n[ERROR] Ocurrió un error: {e}", "ERROR")
            messagebox.showerror("Error", str(e))

        finally:
            self.btn_process.config(state="normal", text=" ▶ PROCESAR Y GENERAR ARCHIVOS")


def launch_gui():
    app = AppGUI()
    app.mainloop()


if __name__ == "__main__":
    launch_gui()
