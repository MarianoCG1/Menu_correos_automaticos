# Automatizador de Cartas

Programa de escritorio que genera cartas Word personalizadas (mail merge)
a partir de una plantilla con campos `{como_este}` y una tabla de datos
(cargada desde Excel y/o editada manualmente). Guarda el progreso en
disco automáticamente para poder continuar en otra sesión.

## 1. Instalar dependencias

```bash
pip install -r requirements.txt
```

## 2. Ejecutar en modo desarrollo

```bash
python main.py
```

## 3. Preparar la plantilla Word

En Word, escribe el texto de la carta y marca los campos variables entre
llaves, por ejemplo:

```
Estimados señores de {empresa},

Por medio de la presente informamos que el número de cuenta
{numero_cuenta} registra un monto pendiente de {monto}.
```

El programa detecta automáticamente `empresa`, `numero_cuenta` y `monto`
como campos a rellenar.

## 4. Flujo de uso

1. **Cargar plantilla Word** → el programa detecta los campos y arma la
   tabla editable.
2. **Cargar Excel** (opcional) → intenta mapear automáticamente cada
   columna del Excel al campo correspondiente de la plantilla (por
   nombre). Los campos que no encuentren columna coincidente quedan en
   blanco y se pueden completar a mano en la tabla.
3. Editar la tabla directamente en pantalla (clic sobre una celda) para
   agregar o corregir datos, y usar el selector de **Estado** por fila
   (Pendiente / Generado / Enviado / Respondido) para hacer seguimiento
   manual mientras las empresas responden.
4. Elegir **Carpeta de salida**.
5. **Generar cartas**: crea un `.docx` por cada fila con el nombre
   `Carta_{empresa}.docx`.

El progreso (plantilla, campos, mapeo, filas y estados) se guarda
automáticamente en `carta_automator_state.json`, junto al ejecutable.
Cerrar y volver a abrir el programa retoma el trabajo donde quedó.

## 5. Generar el ejecutable (.exe / binario)

```bash
pip install pyinstaller
pyinstaller build.spec
```

El ejecutable queda en `dist/AutomatizadorCartas/` (o `dist/AutomatizadorCartas.exe`
en Windows si se usa `--onefile`, aunque el spec incluido usa modo carpeta
por ser más estable con pywebview).

**Importante:** compila en el mismo sistema operativo donde se va a usar
el programa — PyInstaller no genera ejecutables multiplataforma. En
Windows, pywebview usa el motor Edge WebView2 (viene preinstalado en
Windows 10/11 actualizados); en Linux requiere tener instalado GTK
(`python3-gi`, `gir1.2-webkit2-4.0`) o Qt WebEngine.

## Estructura del proyecto

```
carta_automator/
├── main.py            # Ventana pywebview + API expuesta al frontend
├── docx_parser.py      # Detección de campos y generación de cartas
├── excel_loader.py     # Lectura de Excel y mapeo automático de columnas
├── storage.py          # Persistencia en JSON del progreso
├── build.spec          # Spec de PyInstaller
├── requirements.txt
└── web/
    ├── index.html
    ├── style.css
    └── app.js
```
