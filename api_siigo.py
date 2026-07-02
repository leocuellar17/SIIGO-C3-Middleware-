import json
import os
import pandas as pd
from datetime import datetime
import time
import requests
import calendar  # Integrado para el control hermético de fin de mes

class SiigoClient:
    def __init__(self, config_path="config.json"):
        """
        Inicializa el cliente de Siigo.
        """
        self.config_path = config_path
        self.base_url = "https://api.siigo.com/v1"
        self.auth_url = "https://api.siigo.com/auth"
        self.username = ""
        self.access_key = ""
        self.token = None
        
        # Parámetros de facturación (Fase 2) - Valores fallback por defecto
        self.document_id = 1
        self.seller_id = 1
        self.cost_center = 1
        self.default_product_code = "DEFAULT"
        self.default_payment_id = 1
        
        self._load_config()

    def _load_config(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    siigo_conf = config.get("siigo_api", {})
                    self.base_url = siigo_conf.get("base_url", "https://api.siigo.com/v1")
                    self.username = siigo_conf.get("username", "")
                    self.access_key = siigo_conf.get("access_key", "")
                    
                    # Sincronización estricta y segura de las variables de la Fase 2
                    self.document_id = siigo_conf.get("document_id", self.document_id)
                    self.seller_id = siigo_conf.get("seller_id", self.seller_id)
                    self.cost_center = siigo_conf.get("cost_center", self.cost_center)
                    self.default_product_code = siigo_conf.get("default_product_code", self.default_product_code)
                    self.default_payment_id = siigo_conf.get("default_payment_id", self.default_payment_id)
        except Exception as e:
            self._log_error(f"Error cargando config.json: {str(e)}")

    def _log_error(self, message):
        """Registra errores en log_errores.txt"""
        try:
            with open("log_errores.txt", "a", encoding="utf-8") as f:
                f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
        except:
            pass

    def _format_request_error(self, error):
        err_msg = str(error)
        if getattr(error, "response", None) is not None:
            err_msg = f"{error.response.status_code} - {error.response.text}"
        if isinstance(error, requests.exceptions.Timeout):
            return f"Timeout de 30 segundos: {err_msg}"
        return err_msg

    def authenticate(self):
        """
        Obtiene el token Bearer utilizando el usuario y access_key.
        """
        if not self.username or not self.access_key:
            raise ValueError("Credenciales de Siigo no configuradas en config.json")
            
        payload = {
            "username": self.username,
            "access_key": self.access_key
        }
        
        try:
            response = requests.post(self.auth_url, json=payload, headers={"Content-Type": "application/json"}, timeout=30)
            response.raise_for_status()
            data = response.json()
            self.token = data.get("access_token")
            if not self.token:
                raise ValueError("Respuesta exitosa pero sin access_token")
        except requests.exceptions.RequestException as e:
            err_msg = self._format_request_error(e)
            self._log_error(f"Error de autenticación con Siigo: {err_msg}")
            raise Exception(f"Fallo en la autenticación: {err_msg}")
        except Exception as e:
            self._log_error(f"Error de autenticación con Siigo: {str(e)}")
            raise Exception(f"Fallo en la autenticación: {str(e)}")

    def get_headers(self):
        if not self.token:
            self.authenticate()
        return {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.token}"
        }

    def create_invoice(self, data: dict, idempotency_key: str = None) -> dict:
        """
        Realiza la petición POST real a Siigo para crear la factura.
        """
        url = f"{self.base_url}/invoices"
        try:
            headers = self.get_headers()
            if idempotency_key:
                headers["Idempotency-Key"] = str(idempotency_key)

            response = requests.post(url, json=data, headers=headers, timeout=30)

            # Si el token expiró (401), reautenticamos y reintentamos
            if response.status_code == 401:
                self.authenticate()
                retry_headers = self.get_headers()
                if idempotency_key:
                    retry_headers["Idempotency-Key"] = str(idempotency_key)
                response = requests.post(url, json=data, headers=retry_headers, timeout=30)

            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            err_msg = self._format_request_error(e)
            self._log_error(f"Fallo al crear factura (API Error): {err_msg}")
            raise Exception(f"Error creando factura en Siigo: {err_msg}")

    def create_receipt(self, data: dict, idempotency_key: str = None) -> dict:
        """
        Realiza la petición POST a Siigo para crear un recibo de caja.
        """
        url = f"{self.base_url}/receipts"
        
        if "payments" not in data:
            total_value = 0
            if "items" in data:
                for item in data["items"]:
                    price = float(item.get("price", 0))
                    qty = float(item.get("quantity", 1))
                    total_value += price * qty
            data["payments"] = [
                {
                    "id": self.default_payment_id,
                    "value": total_value
                }
            ]

        try:
            headers = self.get_headers()
            if idempotency_key:
                headers["Idempotency-Key"] = str(idempotency_key)

            response = requests.post(url, json=data, headers=headers, timeout=30)
            # Reintentar autenticación si token expiró
            if response.status_code == 401:
                self.authenticate()
                retry_headers = self.get_headers()
                if idempotency_key:
                    retry_headers["Idempotency-Key"] = str(idempotency_key)
                response = requests.post(url, json=data, headers=retry_headers, timeout=30)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.RequestException as e:
            err_msg = self._format_request_error(e)
            self._log_error(f"Fallo al crear recibo (API Error): {err_msg}")
            raise Exception(f"Error creando recibo en Siigo: {err_msg}")

    def get_invoices(self, year: int, month: int) -> pd.DataFrame:
        """
        Llama a la API de Siigo para obtener las facturas, manejando paginación.
        Filtra por año y mes de forma exacta.
        """
        # Corrección Edge Case 3: Filtro de fechas estrictamente acotado al mes consultado
        start_date = f"{year}-{month:02d}-01"
        _, last_day = calendar.monthrange(year, month)
        end_date = f"{year}-{month:02d}-{last_day:02d}"
            
        url = f"{self.base_url}/invoices"
        all_invoices = []
        page = 1
        
        while True:
            params = {
                "date_start": start_date,
                "date_end": end_date,
                "page": page,
                "page_size": 100
            }
            
            try:
                response = requests.get(url, params=params, headers=self.get_headers(), timeout=30)
                
                # Manejo de expiración de token
                if response.status_code == 401:
                    self.authenticate()
                    response = requests.get(url, params=params, headers=self.get_headers(), timeout=30)
                    
                response.raise_for_status()
                data = response.json()
                
                results = data.get("results", [])
                if not results:
                    break
                    
                all_invoices.extend(results)
                
                # Corrección Edge Case 2: Control defensivo de paginación contra nulos (NoneType)
                pagination = data.get("pagination", {})
                total_pages = pagination.get("total_pages")
                if total_pages is None:
                    break

                total_pages = int(total_pages)
                if page >= total_pages:
                    break

                time.sleep(0.3)
                page += 1
                
            except requests.exceptions.RequestException as e:
                err_msg = self._format_request_error(e)
                self._log_error(f"Fallo al obtener facturas (API Error): {err_msg}")
                raise Exception(f"Error obteniendo facturas de Siigo: {err_msg}")

        # Procesar los datos en formato Pandas
        processed_data = []
        for invoice in all_invoices:
            customer = invoice.get("customer", {})
            identification = str(customer.get("identification", ""))
            
            # Corrección Edge Case 1: Validación estricta para evitar quiebre por listas de nombres nulas
            names = customer.get("name")
            nombre_completo = " ".join(names) if isinstance(names, list) else str(names or "Sin Nombre")

            total = float(invoice.get("total", 0))
            fecha = invoice.get("date", "")

            # Captura segura de identificadores de documentos de la API
            invoice_id = invoice.get("id") or invoice.get("invoice_id") or invoice.get("number")
            invoice_number = invoice.get("number") or invoice.get("invoice_number") or ""

            items = invoice.get("items", [])
            if isinstance(items, list):
                try:
                    items_str = "; ".join([f"{it.get('code','')}: {it.get('description','')} ({it.get('quantity',1)} x {it.get('price',0)})" for it in items])
                except Exception:
                    items_str = str(items)
            else:
                items_str = str(items)

            processed_data.append({
                "Invoice_Id": invoice_id,
                "Invoice_Number": invoice_number,
                "DNI_Estudiante": identification,
                "Nombre_Estudiante": nombre_completo,
                "Valor_Facturado": total,
                "Fecha": fecha,
                "Items": items_str
            })

        df_siigo = pd.DataFrame(processed_data)
        return df_siigo

if __name__ == "__main__":
    # Test de autenticación si se ejecuta directo
    client = SiigoClient()
    try:
        client.authenticate()
        print("Autenticación exitosa. Token obtenido.")
    except Exception as e:
        print(f"No se pudo autenticar: {e}")    
