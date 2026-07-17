"""
storage.py
Persistencia en disco (JSON) del progreso del usuario: plantilla cargada,
campos detectados, datos de la tabla y estado de cada carta (Pendiente /
Generado / Enviado / Respondido). Permite cerrar el programa y continuar
después sin perder el trabajo ya avanzado.
"""

import json
import os

STATUS_OPTIONS = ["Pendiente", "Generado", "Enviado", "Respondido"]

DEFAULT_STATE = {
    "template_path": None,
    "fields": [],
    "excel_path": None,
    "excel_mapping": {},
    "output_folder": None,
    "rows": [],  # cada fila: {"data": {campo: valor, ...}, "status": "Pendiente"}
    "email_subject_template": "Carta para {empresa}",
    "email_body_template": "<p>Estimados,</p><p>Adjunto la carta de presentación correspondiente.</p><p>Atentamente,</p>",
    "email_to_template": "{destinatario_correo}",
    "email_cc_template": "{copia_correo}",
    "email_recipient_field": None,
    "email_cc_field": None,
}


class Storage:
    def __init__(self, state_file="carta_automator_state.json"):
        self.state_file = os.path.abspath(state_file)
        self.state = self._load()

    def _load(self):
        if os.path.exists(self.state_file):
            try:
                with open(self.state_file, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                state = dict(DEFAULT_STATE)
                state.update(loaded)
                return state
            except (json.JSONDecodeError, OSError):
                pass
        return dict(DEFAULT_STATE)

    def save(self):
        with open(self.state_file, "w", encoding="utf-8") as f:
            json.dump(self.state, f, ensure_ascii=False, indent=2)

    def reset(self):
        self.state = dict(DEFAULT_STATE)
        self.save()

    def get(self):
        return self.state

    def update(self, partial):
        self.state.update(partial)
        self.save()
