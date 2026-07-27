"""
main.py
Punto de entrada del Automatizador de Cartas. Crea la ventana de
escritorio (pywebview) y expone una API Python que el frontend
(web/index.html + app.js) consume mediante window.pywebview.api.*
"""

import os
import sys
import traceback

import webview

from docx_parser import detect_fields, generate_letter, extract_letter_name
from excel_loader import read_excel, auto_map_columns, build_rows_for_fields
from storage import Storage, STATUS_OPTIONS

try:
    import pythoncom
    import win32com
    import win32com.client
except ImportError:
    pass


def resource_path(relative_path):
    """Devuelve la ruta correcta tanto en desarrollo como empaquetado
    con PyInstaller (usa sys._MEIPASS cuando corre como ejecutable)."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class Api:
    def __init__(self):
        self.storage = Storage()

    # ---------- estado ----------
    def get_state(self):
        return self.storage.get()

    def reset_state(self):
        self.storage.reset()
        return self.storage.get()

    # ---------- plantilla ----------
    def choose_template(self):
        window = webview.windows[0]
        result = window.create_file_dialog(
            webview.OPEN_DIALOG, file_types=("Documentos Word (*.docx)",)
        )
        if not result:
            return {"ok": False, "error": "No se seleccionó ningún archivo."}

        template_path = result[0]
        try:
            fields = detect_fields(template_path)
        except Exception as e:
            return {"ok": False, "error": f"No se pudo leer la plantilla: {e}"}

        if not fields:
            return {
                "ok": False,
                "error": "No se encontraron campos variables (formatos: {campo}, «campo» o <<campo>>) en la plantilla. "
                "Verifica que la plantilla de Word los contenga.",
            }

        # Asegurar que existan campos especiales para correo
        if "destinatario_correo" not in fields:
            fields.append("destinatario_correo")
        if "copia_correo" not in fields:
            fields.append("copia_correo")

        # conserva los datos ya ingresados para campos que sigan existiendo
        existing_rows = self.storage.get().get("rows", [])
        new_rows = []
        for r in existing_rows:
            new_data = {f: r.get("data", {}).get(f, "") for f in fields}
            new_rows.append({"data": new_data, "status": r.get("status", "Pendiente"), "selected": r.get("selected", True)})

        self.storage.update(
            {
                "template_path": template_path,
                "fields": fields,
                "rows": new_rows,
                "excel_path": None,
                "excel_mapping": {},
            }
        )
        return {"ok": True, "state": self.storage.get()}

    # ---------- excel ----------
    def choose_excel(self):
        window = webview.windows[0]
        result = window.create_file_dialog(
            webview.OPEN_DIALOG, file_types=("Hojas de cálculo (*.xlsx;*.xls)",)
        )
        if not result:
            return {"ok": False, "error": "No se seleccionó ningún archivo."}

        excel_path = result[0]
        state = self.storage.get()
        fields = state.get("fields", [])
        if not fields:
            return {
                "ok": False,
                "error": "Primero carga la plantilla Word para detectar los campos.",
            }

        try:
            columns, excel_rows = read_excel(excel_path)
        except Exception as e:
            return {"ok": False, "error": f"No se pudo leer el Excel: {e}"}

        if not columns:
            return {"ok": False, "error": "El Excel parece estar vacío."}

        mapping = auto_map_columns(columns, fields)
        mapped_rows = build_rows_for_fields(fields, columns, excel_rows, mapping)
        
        to_template = state.get("email_to_template", "{destinatario_correo}")
        new_rows = []
        for row in mapped_rows:
            recipient = to_template
            for key, val in row.items():
                placeholder = "{" + key + "}"
                val_str = "" if val is None else str(val)
                recipient = recipient.replace(placeholder, val_str)
            is_valid = bool(recipient and "@" in recipient and recipient.strip())
            new_rows.append({"data": row, "status": "Pendiente", "selected": is_valid})

        self.storage.update(
            {
                "excel_path": excel_path,
                "excel_mapping": mapping,
                "rows": new_rows,
            }
        )
        return {"ok": True, "state": self.storage.get(), "excel_columns": columns}

    def update_mapping(self, field, column):
        state = self.storage.get()
        mapping = state.get("excel_mapping", {})
        mapping[field] = column or None
        self.storage.update({"excel_mapping": mapping})
        return {"ok": True, "state": self.storage.get()}

    # ---------- filas ----------
    def add_row(self):
        state = self.storage.get()
        empty = {f: "" for f in state.get("fields", [])}
        rows = state.get("rows", [])
        rows.append({"data": empty, "status": "Pendiente", "selected": True})
        self.storage.update({"rows": rows})
        return self.storage.get()

    def delete_row(self, index):
        state = self.storage.get()
        rows = state.get("rows", [])
        if 0 <= index < len(rows):
            rows.pop(index)
            self.storage.update({"rows": rows})
        return self.storage.get()

    def update_cell(self, index, field, value):
        state = self.storage.get()
        rows = state.get("rows", [])
        if 0 <= index < len(rows):
            rows[index]["data"][field] = value
            self.storage.update({"rows": rows})
        return {"ok": True}

    def update_status(self, index, status):
        state = self.storage.get()
        rows = state.get("rows", [])
        if 0 <= index < len(rows) and status in STATUS_OPTIONS:
            rows[index]["status"] = status
            self.storage.update({"rows": rows})
        return {"ok": True}

    def toggle_row_selection(self, index, selected):
        state = self.storage.get()
        rows = state.get("rows", [])
        if 0 <= index < len(rows):
            rows[index]["selected"] = bool(selected)
            self.storage.update({"rows": rows})
        return {"ok": True, "state": self.storage.get()}

    def toggle_all_selection(self, selected):
        state = self.storage.get()
        rows = state.get("rows", [])
        for row in rows:
            row["selected"] = bool(selected)
        self.storage.update({"rows": rows})
        return {"ok": True, "state": self.storage.get()}

    def select_only_first(self):
        state = self.storage.get()
        rows = state.get("rows", [])
        for i, row in enumerate(rows):
            row["selected"] = (i == 0)
        self.storage.update({"rows": rows})
        return {"ok": True, "state": self.storage.get()}

    def select_only_valid_emails(self):
        state = self.storage.get()
        rows = state.get("rows", [])
        to_template = state.get("email_to_template", "{destinatario_correo}")
        
        count_selected = 0
        for row in rows:
            data = row.get("data", {})
            recipient = to_template
            for key, val in data.items():
                placeholder = "{" + key + "}"
                val_str = "" if val is None else str(val)
                recipient = recipient.replace(placeholder, val_str)
            
            is_valid = bool(recipient and "@" in recipient and recipient.strip())
            row["selected"] = is_valid
            if is_valid:
                count_selected += 1
                
        self.storage.update({"rows": rows})
        return {"ok": True, "count": count_selected, "state": self.storage.get()}

    # ---------- salida ----------
    def choose_output_folder(self):
        window = webview.windows[0]
        result = window.create_file_dialog(webview.FOLDER_DIALOG)
        if not result:
            return {"ok": False}
        folder = result[0]
        self.storage.update({"output_folder": folder})
        return {"ok": True, "state": self.storage.get()}

    def generate_letters(self):
        state = self.storage.get()
        template_path = state.get("template_path")
        output_folder = state.get("output_folder")
        rows = state.get("rows", [])

        if not template_path:
            return {"ok": False, "error": "No hay plantilla cargada."}
        if not output_folder:
            return {"ok": False, "error": "Selecciona primero una carpeta de salida."}
        if not rows:
            return {"ok": False, "error": "No hay datos para generar cartas."}

        generated = 0
        errors = []
        for i, row in enumerate(rows):
            data = row.get("data", {})
            base_name = None
            for candidate_key in ("empresa", "nombre_empresa", "compania", "cliente"):
                if candidate_key in data and data[candidate_key]:
                    base_name = str(data[candidate_key])
                    break
            if not base_name:
                base_name = f"registro_{i + 1}"

            safe_name = "".join(
                c if c.isalnum() or c in " _-" else "_" for c in base_name
            ).strip()
            output_path = os.path.join(output_folder, f"Carta_{safe_name}.docx")

            try:
                generate_letter(template_path, data, output_path)
                row["status"] = "Generado"
                generated += 1
            except Exception as e:
                errors.append(f"Fila {i + 1} ({base_name}): {e}")

        self.storage.update({"rows": rows})
        return {
            "ok": True,
            "generated": generated,
            "errors": errors,
            "state": self.storage.get(),
        }

    def save_email_template(self, subject, body, to_template=None, cc_template=None):
        update_dict = {
            "email_subject_template": subject,
            "email_body_template": body
        }
        if to_template is not None:
            update_dict["email_to_template"] = to_template
        if cc_template is not None:
            update_dict["email_cc_template"] = cc_template

        self.storage.update(update_dict)
        return {"ok": True, "state": self.storage.get()}

    def get_signature_for_mail(self, mail):
        try:
            initial_body = mail.HTMLBody
            if initial_body and len(initial_body.strip()) > 100:
                return initial_body
        except Exception:
            pass
            
        try:
            import os
            appdata = os.environ.get('APPDATA', '')
            sig_dir = os.path.join(appdata, 'Microsoft', 'Signatures')
            if os.path.exists(sig_dir):
                files = [f for f in os.listdir(sig_dir) if f.endswith('.htm') or f.endswith('.html')]
                if files:
                    files.sort(key=lambda x: os.path.getmtime(os.path.join(sig_dir, x)), reverse=True)
                    sig_path = os.path.join(sig_dir, files[0])
                    
                    for encoding in ('utf-16', 'utf-8', 'latin-1'):
                        try:
                            with open(sig_path, 'r', encoding=encoding, errors='ignore') as f:
                                content = f.read()
                                if content:
                                    return content
                        except Exception:
                            continue
        except Exception:
            pass
        return ""

    def generate_outlook_drafts(self, send_directly=False):
        state = self.storage.get()
        template_path = state.get("template_path")
        output_folder = state.get("output_folder")
        rows = state.get("rows", [])
        subject_template = state.get("email_subject_template", "Carta para {empresa}")
        body_template = state.get("email_body_template", "")
        to_template = state.get("email_to_template", "{destinatario_correo}")
        cc_template = state.get("email_cc_template", "{copia_correo}")

        if not template_path:
            return {"ok": False, "error": "No hay plantilla Word cargada."}
        if not output_folder:
            import tempfile
            output_folder = tempfile.gettempdir()
        if not rows:
            return {"ok": False, "error": "No hay datos en la tabla."}

        try:
            import pythoncom
            import win32com.client
        except ImportError:
            return {
                "ok": False,
                "error": "No se encontró el módulo pywin32 para interactuar con Outlook. "
                "Por favor, instala la biblioteca ejecutando: pip install pywin32"
            }

        # Inicializar COM para el hilo actual
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
        except Exception:
            pass

        word_app = None
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
        except Exception as e:
            if com_initialized:
                pythoncom.CoUninitialize()
            return {
                "ok": False,
                "error": f"No se pudo conectar a Outlook o Word. Asegúrate de que estén abiertos/instalados. Error: {e}"
            }

        processed = 0
        errors = []
        
        import tempfile
        import shutil
        try:
            official_letter_name = extract_letter_name(template_path)
        except Exception:
            official_letter_name = "Carta_Generada"
        temp_dir = tempfile.gettempdir()

        selected_rows = [(i, r) for i, r in enumerate(rows) if r.get("selected", True)]
        if not selected_rows:
            return {"ok": False, "error": "No hay ningún registro seleccionado para enviar. Marca al menos una casilla."}

        for i, row in selected_rows:
            data = row.get("data", {})
            
            # --- Formatear Número de Cuenta ---
            for key in list(data.keys()):
                if "cuenta" in key.lower():
                    acc = str(data[key]).replace("-", "").replace(" ", "").strip()
                    if len(acc) == 10 and acc.isdigit():
                        data[key] = f"0{acc[0]}-{acc[1:4]}-{acc[4:]}"
                    elif len(acc) == 11 and acc.startswith("0") and acc.isdigit():
                        data[key] = f"0{acc[1]}-{acc[2:5]}-{acc[5:]}"
            
            # 1. Determinar el nombre base del archivo de salida
            cliente_name = None
            for candidate_key in ("empresa", "nombre_empresa", "compania", "cliente", "razon_social"):
                if candidate_key in data and data[candidate_key]:
                    cliente_name = str(data[candidate_key])
                    break
            if not cliente_name:
                cliente_name = f"registro_{i + 1}"

            carta_num = ""
            for candidate_key in ("n_carta", "numero_carta", "correlativo"):
                if candidate_key in data and data[candidate_key]:
                    carta_num = str(data[candidate_key]) + "_"
                    break

            base_name = f"{carta_num}{cliente_name}"

            safe_name = "".join(
                c if c.isalnum() or c in " _-" else "_" for c in base_name
            ).strip()
            output_path = os.path.join(output_folder, f"Carta_{safe_name}.docx")

            # 2. Generar el documento de Word
            try:
                generate_letter(template_path, data, output_path)
            except Exception as e:
                errors.append(f"Fila {i + 1} ({base_name}) - Error Word: {e}")
                continue

            # 3. Formatear asunto, cuerpo y destinatarios del correo reemplazando variables
            subject = subject_template
            body = body_template
            recipient = to_template
            cc = cc_template

            for key, val in data.items():
                placeholder = "{" + key + "}"
                val_str = "" if val is None else str(val)
                subject = subject.replace(placeholder, val_str)
                body = body.replace(placeholder, val_str)
                recipient = recipient.replace(placeholder, val_str)
                cc = cc.replace(placeholder, val_str)

            # 4. Crear correo en Outlook
            try:
                mail = outlook.CreateItem(0)
                mail.Subject = subject
                
                # Intentar inyectar firma
                signature = self.get_signature_for_mail(mail)
                if signature:
                    if "<body>" in signature.lower():
                        import re
                        parts = re.split(r'(<body>)', signature, flags=re.IGNORECASE)
                        mail.HTMLBody = parts[0] + parts[1] + body + "<br><br>" + parts[2]
                    else:
                        mail.HTMLBody = body + "<br><br>" + signature
                else:
                    mail.HTMLBody = body

                mail.To = recipient.strip()
                if cc.strip():
                    mail.CC = cc.strip()

                # Adjuntar archivo como PDF de forma robusta
                temp_pdf_path = os.path.join(temp_dir, f"{official_letter_name}.pdf")
                pdf_success = False
                
                try:
                    doc_for_pdf = word_app.Documents.Open(os.path.abspath(output_path), ReadOnly=True, ConfirmConversions=False)
                    doc_for_pdf.SaveAs2(os.path.abspath(temp_pdf_path), FileFormat=17) # 17 = wdFormatPDF
                    doc_for_pdf.Close(SaveChanges=0)
                    pdf_success = True
                except Exception as pdf_err:
                    # Reintento: matar instancia corrupta y reiniciar Word COM
                    try:
                        word_app.Quit()
                    except Exception:
                        pass
                    try:
                        word_app = win32com.client.Dispatch("Word.Application")
                        word_app.Visible = False
                        word_app.DisplayAlerts = 0
                        doc_for_pdf = word_app.Documents.Open(os.path.abspath(output_path), ReadOnly=True, ConfirmConversions=False)
                        doc_for_pdf.SaveAs2(os.path.abspath(temp_pdf_path), FileFormat=17)
                        doc_for_pdf.Close(SaveChanges=0)
                        pdf_success = True
                    except Exception as retry_err:
                        errors.append(f"Fila {i + 1} ({base_name}) - Error PDF (sin envío): {retry_err}")
                        continue  # NO envía el correo si no pudo generar PDF
                
                if pdf_success:
                    mail.Attachments.Add(os.path.abspath(temp_pdf_path))
                    try:
                        os.remove(temp_pdf_path)
                    except Exception:
                        pass

                if send_directly:
                    mail.Send()
                    row["status"] = "Enviado"
                else:
                    mail.Save()
                    row["status"] = "Generado"

                processed += 1
            except Exception as e:
                errors.append(f"Fila {i + 1} ({base_name}) - Error Outlook: {e}")

        # Guardar el estado actualizado de las filas
        self.storage.update({"rows": rows})

        if word_app:
            try:
                word_app.Quit()
            except Exception:
                pass

        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

        return {
            "ok": True,
            "processed": processed,
            "errors": errors,
            "state": self.storage.get()
        }
    def add_contact(self, email):
        state = self.storage.get()
        book = state.get("address_book", [])
        if email and email not in book:
            book.append(email)
            self.storage.update({"address_book": book})
        return {"ok": True, "state": self.storage.get()}

    def remove_contact(self, email):
        state = self.storage.get()
        book = state.get("address_book", [])
        if email in book:
            book.remove(email)
            self.storage.update({"address_book": book})
        return {"ok": True, "state": self.storage.get()}

    def generate_single_pdf(self, row_index):
        state = self.storage.get()
        template_path = state.get("template_path")
        rows = state.get("rows", [])
        
        if not template_path:
            return {"ok": False, "error": "No hay plantilla cargada."}
        if row_index < 0 or row_index >= len(rows):
            return {"ok": False, "error": "Índice de fila inválido."}
        
        data = dict(rows[row_index].get("data", {}))
        
        # Formatear cuenta
        for key in list(data.keys()):
            if "cuenta" in key.lower():
                acc = str(data[key]).replace("-", "").replace(" ", "").strip()
                if len(acc) == 10 and acc.isdigit():
                    data[key] = f"0{acc[0]}-{acc[1:4]}-{acc[4:]}"
                elif len(acc) == 11 and acc.startswith("0") and acc.isdigit():
                    data[key] = f"0{acc[1]}-{acc[2:5]}-{acc[5:]}"
        
        import tempfile
        temp_dir = tempfile.gettempdir()
        temp_docx = os.path.join(temp_dir, f"auditoria_fila_{row_index + 1}.docx")
        
        try:
            generate_letter(template_path, data, temp_docx)
        except Exception as e:
            return {"ok": False, "error": f"Error generando carta: {e}"}
        
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = 0
            
            letter_name = extract_letter_name(template_path)
            pdf_path = os.path.join(temp_dir, f"{letter_name}.pdf")
            
            doc = word_app.Documents.Open(os.path.abspath(temp_docx), ReadOnly=True, ConfirmConversions=False)
            doc.SaveAs2(os.path.abspath(pdf_path), FileFormat=17)
            doc.Close(SaveChanges=0)
            word_app.Quit()
            pythoncom.CoUninitialize()
            
            os.startfile(pdf_path)
            return {"ok": True, "pdf_path": pdf_path}
        except Exception as e:
            return {"ok": False, "error": f"Error convirtiendo a PDF: {e}"}
        finally:
            try:
                os.remove(temp_docx)
            except Exception:
                pass


def main():
    api = Api()
    html_path = resource_path(os.path.join("web", "index.html"))
    webview.create_window(
        "Automatizador de Cartas",
        html_path,
        js_api=api,
        width=1200,
        height=750,
        min_size=(900, 600),
    )
    webview.start(debug=False)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        input("Ocurrió un error. Presiona Enter para salir...")
