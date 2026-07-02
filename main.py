import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
import datetime
import os
import time
import json

from api_siigo import SiigoClient
from reconciliation import Reconciler


def resolve_transaction_date(row, year, month, fallback_day=28):
    raw_date = None
    for field in ("Fecha_C3", "Fecha", "Fecha_Emision", "Fecha Emisión", "Fecha_Emisión", "Fecha_Transaccion"):
        if field in row:
            raw_date = row.get(field)
            break

    if raw_date is None or (isinstance(raw_date, float) and pd.isna(raw_date)):
        return f"{year}-{month:02d}-{fallback_day:02d}"

    if isinstance(raw_date, str):
        raw_date = raw_date.strip()
    if not raw_date:
        return f"{year}-{month:02d}-{fallback_day:02d}"

    try:
        parsed_date = pd.to_datetime(raw_date, errors="coerce")
        if pd.isna(parsed_date):
            return f"{year}-{month:02d}-{fallback_day:02d}"
        return parsed_date.strftime("%Y-%m-%d")
    except Exception:
        return f"{year}-{month:02d}-{fallback_day:02d}"


# Configuración básica de UI
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class BridgeApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Bridge-Finance Siigo-C3")
        self.geometry("900x650")
        self.minsize(800, 600)

        # Variables
        self.c3_file_path = tk.StringVar()
        self.current_result_df = None
        self.client = SiigoClient()

        # Leer tolerance desde config.json con fallback seguro
        tolerance = 10.0
        try:
            config_path = 'config.json'
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    cfg = json.load(f)
                    tolerance = float(cfg.get('settings', {}).get('tolerance_pesos', tolerance))
        except Exception as e:
            # Registrar el error y continuar con el fallback
            try:
                self.client._log_error(f"Error leyendo tolerance desde config.json: {str(e)}")
            except Exception:
                pass

        self.reconciler = Reconciler(tolerance=tolerance)

        self._build_ui()

    def _build_ui(self):
        # Frame Superior: Controles
        self.top_frame = ctk.CTkFrame(self)
        self.top_frame.pack(pady=10, padx=10, fill="x")

        # --- Fila 1: Archivo C3 ---
        self.lbl_c3 = ctk.CTkLabel(self.top_frame, text="Archivo C3:", font=("Arial", 12, "bold"))
        self.lbl_c3.grid(row=0, column=0, padx=5, pady=10, sticky="w")

        self.entry_c3 = ctk.CTkEntry(self.top_frame, textvariable=self.c3_file_path, width=400, state="readonly")
        self.entry_c3.grid(row=0, column=1, padx=5, pady=10, sticky="ew")

        self.btn_browse = ctk.CTkButton(self.top_frame, text="Subir Archivo", command=self.browse_file)
        self.btn_browse.grid(row=0, column=2, padx=5, pady=10)

        # --- Fila 2: Filtro Siigo ---
        self.lbl_periodo = ctk.CTkLabel(self.top_frame, text="Periodo Siigo:", font=("Arial", 12, "bold"))
        self.lbl_periodo.grid(row=1, column=0, padx=5, pady=10, sticky="w")

        period_frame = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        period_frame.grid(row=1, column=1, sticky="w")

        current_year = datetime.datetime.now().year
        self.combo_year = ctk.CTkComboBox(period_frame, values=[str(y) for y in range(current_year-2, current_year+2)], width=100)
        self.combo_year.set(str(current_year))
        self.combo_year.pack(side="left", padx=5)

        meses = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
        current_month = datetime.datetime.now().month
        self.combo_month = ctk.CTkComboBox(period_frame, values=meses, width=80)
        self.combo_month.set(f"{current_month:02d}")
        self.combo_month.pack(side="left", padx=5)

        # --- Fila 3: Botones Ejecutar ---
        buttons_frame = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        buttons_frame.grid(row=2, column=0, columnspan=3, pady=15)
        
        self.btn_run = ctk.CTkButton(buttons_frame, text="Conciliar", command=self.run_reconciliation_thread, fg_color="green", hover_color="darkgreen")
        self.btn_run.pack(side="left", padx=10)

        self.btn_mass_invoice = ctk.CTkButton(buttons_frame, text="Subir información y Conciliar", command=self.run_mass_invoice_thread, fg_color="#1f538d", hover_color="#14375e")
        self.btn_mass_invoice.pack(side="left", padx=10)

        # Frame Central: Búsqueda y Tabla
        self.mid_frame = ctk.CTkFrame(self)
        self.mid_frame.pack(pady=5, padx=10, fill="both", expand=True)

        # Buscador
        search_frame = ctk.CTkFrame(self.mid_frame, fg_color="transparent")
        search_frame.pack(fill="x", pady=5, padx=5)
        
        self.lbl_search = ctk.CTkLabel(search_frame, text="Buscar (DNI o Nombre):")
        self.lbl_search.pack(side="left", padx=5)
        
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_table)
        self.entry_search = ctk.CTkEntry(search_frame, textvariable=self.search_var, width=250)
        self.entry_search.pack(side="left", padx=5)

        self.btn_export = ctk.CTkButton(search_frame, text="Exportar a Excel", command=self.export_excel, state="disabled")
        self.btn_export.pack(side="right", padx=5)

        # Tabla (Treeview)
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", rowheight=25, font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 11, "bold"))

        tree_scroll_y = ttk.Scrollbar(self.mid_frame)
        tree_scroll_y.pack(side="right", fill="y")
        tree_scroll_x = ttk.Scrollbar(self.mid_frame, orient="horizontal")
        tree_scroll_x.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(self.mid_frame, yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        self.tree.pack(fill="both", expand=True)
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)

        # Barra de estado y progreso
        status_frame = ctk.CTkFrame(self, fg_color="transparent")
        status_frame.pack(side="bottom", fill="x", padx=10, pady=5)
        
        self.status_var = tk.StringVar()
        self.status_var.set("Listo")
        self.status_bar = ctk.CTkLabel(status_frame, textvariable=self.status_var, anchor="w", font=("Arial", 11, "italic"))
        self.status_bar.pack(side="left")
        
        self.percent_var = tk.StringVar()
        self.percent_var.set("0%")
        self.percent_label = ctk.CTkLabel(status_frame, textvariable=self.percent_var, font=("Arial", 11, "bold"))
        self.percent_label.pack(side="right", padx=(0, 10))
        
        self.progress_bar = ctk.CTkProgressBar(status_frame, width=200)
        self.progress_bar.pack(side="right", padx=10)
        self.progress_bar.set(0.0)

    def browse_file(self):
        filetypes = (
            ('Archivos Excel/CSV', '*.xlsx *.xls *.csv'),
            ('Todos los archivos', '*.*')
        )
        filename = filedialog.askopenfilename(title='Seleccionar reporte C3', filetypes=filetypes)
        if filename:
            self.c3_file_path.set(filename)

    def update_status(self, message):
        self.status_var.set(message)
        self.update_idletasks()

    def run_reconciliation_thread(self):
        if not self.c3_file_path.get():
            messagebox.showwarning("Atención", "Por favor selecciona un archivo C3 primero.")
            return

        # Deshabilitar botones durante ejecución
        self.btn_run.configure(state="disabled")
        self.btn_mass_invoice.configure(state="disabled")
        self.btn_export.configure(state="disabled")
        self.tree.delete(*self.tree.get_children())
        
        thread = threading.Thread(target=self._execute_reconciliation)
        thread.start()

    def _execute_reconciliation(self):
        try:
            year = int(self.combo_year.get())
            month = int(self.combo_month.get())
            
            self.update_status(f"Cargando archivo C3 desde {self.c3_file_path.get()}...")
            df_c3 = self.reconciler.load_c3_data(self.c3_file_path.get())
            
            self.update_status(f"Obteniendo datos de Siigo para el periodo {month:02d}/{year}...")
            time.sleep(0.5)
            df_siigo = self.client.get_invoices(year, month)
            
            self.update_status("Ejecutando motor de conciliación...")
            self.current_result_df = self.reconciler.reconcile(df_siigo, df_c3)
            
            self.update_status("Generando vista de tabla...")
            # Llenar la tabla de forma segura para Tkinter desde el thread
            self.after(0, self.populate_table, self.current_result_df)
            self.after(0, lambda: self.update_status(f"Conciliación finalizada. Total registros: {len(self.current_result_df)}"))
            
        except ValueError as ve:
            self.after(0, lambda: messagebox.showerror("Error de Formato", str(ve)))
            self.after(0, lambda: self.update_status("Error en formato de archivo."))
        except Exception as e:
            self.client._log_error(f"Error inesperado en main: {str(e)}")
            self.after(0, lambda: messagebox.showerror("Error Inesperado", f"Ocurrió un error: {str(e)}\nRevisa log_errores.txt"))
            self.after(0, lambda: self.update_status("Error durante la conciliación."))
        finally:
            self.after(0, lambda: self.btn_run.configure(state="normal"))
            self.after(0, lambda: self.btn_mass_invoice.configure(state="normal"))
            if self.current_result_df is not None:
                self.after(0, lambda: self.btn_export.configure(state="normal"))

    def run_mass_invoice_thread(self):
        if not self.c3_file_path.get():
            messagebox.showwarning("Atención", "Por favor selecciona un archivo C3 primero.")
            return

        try:
            df_c3 = self.reconciler.load_c3_data(self.c3_file_path.get())
        except ValueError as ve:
            messagebox.showerror("Error de Formato", str(ve))
            return
        except Exception as e:
            self.client._log_error(f"Error cargando C3 para confirmación de facturación: {str(e)}")
            messagebox.showerror("Error Inesperado", f"Ocurrió un error al cargar el archivo C3: {str(e)}")
            return

        record_count = len(df_c3)
        confirmed = messagebox.askyesno(
            "Confirmar facturación masiva",
            f"Se van a crear facturas/recibos reales en Siigo para {record_count} registros.\n\nEsta acción es irreversible. ¿Desea continuar?"
        )
        if not confirmed:
            return

        self.btn_run.configure(state="disabled")
        self.btn_mass_invoice.configure(state="disabled")
        self.btn_export.configure(state="disabled")
        self.tree.delete(*self.tree.get_children())
        
        thread = threading.Thread(target=self._execute_mass_invoicing)
        thread.start()

    def _execute_mass_invoicing(self):
        try:
            self.update_status(f"Cargando archivo C3 para facturación...")
            df_c3 = self.reconciler.load_c3_data(self.c3_file_path.get())
            
            total_records = len(df_c3)
            success_count = 0
            error_count = 0
            
            year = int(self.combo_year.get())
            month = int(self.combo_month.get())

            total_transactions = 0
            for _, r in df_c3.iterrows():
                valor_cobrado = float(r.get('Valor_Cobrado', 0))
                valor_recibido = float(r.get('Valor_Recibido', 0))
                if valor_cobrado > 0:
                    total_transactions += 1
                if valor_recibido > 0:
                    total_transactions += 1
            if total_transactions == 0:
                self.after(0, lambda: messagebox.showinfo("Sin datos", "No hay valores a facturar o registrar recibos."))
                self.after(0, lambda: self.btn_run.configure(state="normal"))
                self.after(0, lambda: self.btn_mass_invoice.configure(state="normal"))
                return

            processed = 0
            for idx, row in df_c3.iterrows():
                transaction_date = resolve_transaction_date(row, year, month)
                dni = str(row.get('DNI_Estudiante', '')).strip()
                name = str(row.get('Nombre_Estudiante', '')).strip()
                acudiente = str(row.get('Nombre_Acudiente', '')).strip()
                nit = str(row.get('NIT', '')).strip()
                num_factura = str(row.get('Numero_Factura', '')).strip()
                concepto = str(row.get('Concepto', self.client.default_product_code)).strip()
                if not concepto:
                    concepto = self.client.default_product_code

                customer_ident = nit if nit and nit.lower() != 'nan' else dni
                customer_name = acudiente if acudiente and acudiente.lower() != 'nan' else name

                valor_cobrado = float(row.get('Valor_Cobrado', 0))
                valor_recibido = float(row.get('Valor_Recibido', 0))

                # Crear Factura de Venta (Débito) si hay valor cobrado
                if valor_cobrado > 0:
                    invoice_data = {
                        "document": {"id": self.client.document_id},
                        "date": transaction_date,
                        "customer": {
                            "identification": customer_ident,
                            "name": [customer_name]
                        },
                        "cost_center": self.client.cost_center,
                        "seller": self.client.seller_id,
                        "mail": {"send": False},
                        "observations": f"Estudiante: {name} | Cod Perm: {dni} | Factura Origen: {num_factura}",
                        "items": [{
                            "code": concepto,
                            "description": concepto,
                            "quantity": 1,
                            "price": valor_cobrado
                        }]
                    }
                    self.update_status(f"Facturando {idx+1}/{len(df_c3)}: {name} (cobro)...")
                    try:
                        idempotency = str(row.get('Idempotency_Key', '')) if row.get('Idempotency_Key', '') is not None else None
                        invoice_resp = self.client.create_invoice(invoice_data, idempotency_key=idempotency)
                        success_count += 1
                    except Exception as ex:
                        error_count += 1
                        self.client._log_error(f"Error creando factura para DNI {dni}: {str(ex)}")
                    processed += 1
                    self.after(0, lambda p=processed/total_transactions, ps=f"{int((processed/total_transactions)*100)}%": (self.progress_bar.set(p), self.percent_var.set(ps)))

                # Crear Recibo de Caja (Crédito) si hay valor recibido — independiente de la factura
                if valor_recibido > 0:
                    receipt_data = {
                        "date": transaction_date,
                        "customer": {
                            "identification": customer_ident,
                            "name": [customer_name]
                        },
                        "cost_center": self.client.cost_center,
                        "seller": self.client.seller_id,
                        "observations": f"Pago recibido: {name} | Cod Perm: {dni}",
                        "items": [{
                            "code": concepto,
                            "description": concepto,
                            "quantity": 1,
                            "price": valor_recibido
                        }]
                    }
                    self.update_status(f"Registrando recibido {idx+1}/{len(df_c3)}: {name} (pago)...")
                    try:
                        idempotency = str(row.get('Idempotency_Key', '')) if row.get('Idempotency_Key', '') is not None else None
                        self.client.create_receipt(receipt_data, idempotency_key=idempotency)
                        success_count += 1
                    except Exception as ex:
                        error_count += 1
                        self.client._log_error(f"Error creando recibo para DNI {dni}: {str(ex)}")
                    processed += 1
                    self.after(0, lambda p=processed/total_transactions, ps=f"{int((processed/total_transactions)*100)}%": (self.progress_bar.set(p), self.percent_var.set(ps)))

                time.sleep(0.5)
            
            self.after(0, lambda: self.progress_bar.set(0.0))
            self.after(0, lambda: self.percent_var.set("0%"))
            self.after(0, lambda sc=success_count, ec=error_count: messagebox.showinfo(
                "Resumen de Facturación",
                f"Proceso finalizado.\n\nTransacciones exitosas: {sc}\nErrores: {ec}\n\nIniciando conciliación automática..."))
            
            # Disparar conciliación al terminar
            self._execute_reconciliation()
            
        except ValueError as ve:
            self.after(0, lambda: messagebox.showerror("Error de Formato", str(ve)))
            self.after(0, lambda: self.update_status("Error en formato de archivo."))
            self.after(0, lambda: self.btn_run.configure(state="normal"))
            self.after(0, lambda: self.btn_mass_invoice.configure(state="normal"))
        except Exception as e:
            self.client._log_error(f"Error inesperado en facturación: {str(e)}")
            self.after(0, lambda: messagebox.showerror("Error Inesperado", f"Ocurrió un error: {str(e)}\nRevisa log_errores.txt"))
            self.after(0, lambda: self.update_status("Error durante la facturación."))
            self.after(0, lambda: self.btn_run.configure(state="normal"))
            self.after(0, lambda: self.btn_mass_invoice.configure(state="normal"))

    def populate_table(self, df):
        self.tree.delete(*self.tree.get_children())
        
        if df is None or df.empty:
            return

        # Configurar columnas
        columns = list(df.columns)
        self.tree["columns"] = columns
        self.tree["show"] = "headings"
        
        for col in columns:
            self.tree.heading(col, text=col)
            # Ajustar ancho estimado
            self.tree.column(col, width=150, minwidth=100)
            
        # Insertar filas
        for idx, row in df.iterrows():
            values = [str(val) if not pd.isna(val) else "" for val in row]
            # Etiquetas para colorear filas
            tags = ()
            if "ANOMALÍA" in str(row.get("Estado", "")).upper() or "ERROR" in str(row.get("Estado", "")).upper():
                tags = ('anomaly',)
                
            self.tree.insert("", "end", values=values, tags=tags)
            
        # Configurar color de tags
        self.tree.tag_configure('anomaly', background='#FFC7CE', foreground='#9C0006')

    def filter_table(self, *args):
        if self.current_result_df is None:
            return
            
        search_term = self.search_var.get().lower()
        if not search_term:
            self.populate_table(self.current_result_df)
            return
            
        # Filtrar por DNI o Nombres
        # Combinamos los nombres de C3 y Siigo para buscar
        mask = self.current_result_df.apply(lambda row: 
            search_term in str(row.get('DNI_Estudiante', '')).lower() or
            search_term in str(row.get('Nombre_Estudiante_C3', '')).lower() or
            search_term in str(row.get('Nombre_Estudiante_Siigo', '')).lower()
        , axis=1)
        
        filtered_df = self.current_result_df[mask]
        self.populate_table(filtered_df)

    def export_excel(self):
        if self.current_result_df is None:
            return
            
        default_name = f"Conciliacion_Siigo_C3_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx")],
            title="Guardar Reporte"
        )
        
        if save_path:
            try:
                self.update_status("Exportando a Excel...")
                self.reconciler.export_to_excel(self.current_result_df, save_path)
                self.update_status(f"Exportado exitosamente a {os.path.basename(save_path)}")
                messagebox.showinfo("Éxito", "Archivo exportado correctamente.")
            except Exception as e:
                self.update_status("Error exportando a Excel.")
                messagebox.showerror("Error", f"No se pudo exportar: {str(e)}")

if __name__ == "__main__":
    import pandas as pd # Asegurar que pandas está en scope principal para chequeos rápidos
    app = BridgeApp()
    app.mainloop()
