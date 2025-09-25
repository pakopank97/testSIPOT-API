import pandas as pd
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import json
import re
import os
import uuid
import threading
import openpyxl

app = Flask(__name__)
CORS(app)

# --- Carpetas temporales ---
UPLOAD_FOLDER = 'temp_uploads'
DOWNLOAD_FOLDER = 'temp_downloads'
for folder in [UPLOAD_FOLDER, DOWNLOAD_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

tasks = {}

# --- Funciones de Validación ---
def es_numero(valor):
    try: float(valor); return True
    except (ValueError, TypeError): return False
def es_fecha(valor):
    try: pd.to_datetime(str(valor), errors='raise'); return True
    except (ValueError, TypeError): return False
def es_hora(valor):
    return bool(re.match(r'^([01]\d|2[0-3]):([0-5]\d)(:([0-5]\d))?$', str(valor).strip()))
def es_url(valor):
    return str(valor).strip().lower().startswith(('http://', 'https://'))
def es_anio(valor):
    return es_numero(valor) and len(str(valor).strip().split('.')[0]) == 4

VALIDADORES = {
    '3': {'func': es_numero, 'nombre': 'Número'}, '4': {'func': es_fecha, 'nombre': 'Fecha'},
    '5': {'func': es_hora, 'nombre': 'Hora (HH:MM)'}, '6': {'func': es_numero, 'nombre': 'Moneda (formato numérico)'},
    '7': {'func': es_url, 'nombre': 'Página web (URL)'}, '12': {'func': es_anio, 'nombre': 'Año (4 dígitos)'},
    '13': {'func': es_fecha, 'nombre': 'Fecha'}
}

def obtener_coordenada_excel(fila_idx, col_idx):
    col_str = ""
    col_idx_temp = col_idx
    while col_idx_temp >= 0:
        col_str = chr(ord('A') + col_idx_temp % 26) + col_str
        col_idx_temp = col_idx_temp // 26 - 1
    return f"{col_str}{fila_idx + 1}"

# --- FUNCIÓN PRINCIPAL ---
def procesar_archivo_en_segundo_plano(filepath, task_id):
    global tasks
    try:
        df = pd.read_excel(filepath, sheet_name=0, header=None, dtype=str)

        lista_de_errores = []
        filas_invalidas = set()
        reglas = df.iloc[3].fillna('0').tolist() if len(df) > 3 else []
        headers = list(df.iloc[6].fillna('')) if len(df) > 6 else []
        datos_a_validar = df.iloc[7:]

        # --- Validación ---
        for fila_idx, fila in datos_a_validar.iterrows():
            for col_idx, valor in enumerate(fila):
                if col_idx < len(headers) and str(headers[col_idx]).strip() != '':
                    header_name = headers[col_idx]
                    if pd.isna(valor) or str(valor).strip() == '':
                        coordenada = obtener_coordenada_excel(fila_idx, col_idx)
                        lista_de_errores.append(f"Fila {fila_idx+1} eliminada. Celda {coordenada} bajo '{header_name}' estaba vacía.")
                        filas_invalidas.add(fila_idx)
                        break
                    else:
                        regla = str(reglas[col_idx]).strip().split('.')[0] if col_idx < len(reglas) else '0'
                        if regla in VALIDADORES:
                            validador = VALIDADORES[regla]
                            if not validador['func'](valor):
                                coordenada = obtener_coordenada_excel(fila_idx, col_idx)
                                lista_de_errores.append(f"Fila {fila_idx+1} eliminada. Celda {coordenada} ('{valor}') inválida. Se esperaba: {validador['nombre']}.")
                                filas_invalidas.add(fila_idx)
                                break

        if lista_de_errores:
            # Mantener primeras 7 filas
            df_header = df.iloc[:7]
            df_datos = df.iloc[7:]

            # Eliminar filas inválidas
            filas_invalidas_absolutas = [i for i in filas_invalidas if i >= 7]
            df_corregido_datos = df_datos.drop(index=filas_invalidas_absolutas)

            # Reconstruir Excel
            df_corregido = pd.concat([df_header, df_corregido_datos])

            base, ext = os.path.splitext(os.path.basename(filepath))
            nombre_archivo_corregido = f"{base}_Formato_Valido_Permitido_Subir.xlsx"
            ruta_archivo_corregido = os.path.join(DOWNLOAD_FOLDER, nombre_archivo_corregido)

            df_corregido.to_excel(ruta_archivo_corregido, index=False, header=False)

            # Ocultar filas 1, 4 y 5
            wb = openpyxl.load_workbook(ruta_archivo_corregido)
            ws = wb.active
            for fila in [1, 4, 5]:
                ws.row_dimensions[fila].hidden = True
            wb.save(ruta_archivo_corregido)

            tasks[task_id]['status'] = 'complete'
            tasks[task_id]['result'] = {
                'status': 'error',
                'message': "El archivo fue limpiado y las filas inválidas eliminadas.",
                'errors': lista_de_errores,
                'download_file': nombre_archivo_corregido
            }
            return

        # Caso sin errores → JSON
        formato = df.iloc[0, 0] if not df.empty else "Formato no encontrado"
        headers_backend = df.iloc[4].fillna('header_desconocido').astype(str).tolist() if len(df) > 4 else []
        datos_backend_df = df.iloc[7:]
        data_intercalada = []
        if not datos_backend_df.empty and headers_backend:
            datos_backend_df.columns = headers_backend[:len(datos_backend_df.columns)]
            data_intercalada = datos_backend_df.fillna('').to_dict(orient='records')
        jsonBACKEND = {"id_formato": str(formato), "data": data_intercalada}

        nombre_archivo_unico = f"{task_id}.json"
        ruta_archivo = os.path.join(DOWNLOAD_FOLDER, nombre_archivo_unico)
        with open(ruta_archivo, 'w', encoding='utf-8') as f:
            json.dump(jsonBACKEND, f, ensure_ascii=False, indent=2)

        tasks[task_id]['status'] = 'complete'
        tasks[task_id]['result'] = {'status': 'success', 'download_file': nombre_archivo_unico}

    except Exception as e:
        tasks[task_id] = {'status': 'failed', 'error': str(e)}

# --- RUTAS ---
@app.route('/')
def servir_pagina_principal():
    return send_from_directory('.', 'index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'archivo' not in request.files: return jsonify({"error": "No se encontró el archivo"}), 400
    archivo = request.files['archivo']
    if archivo.filename == '': return jsonify({"error": "No se seleccionó archivo"}), 400
    filename = f"{uuid.uuid4()}_{archivo.filename}"
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    archivo.save(filepath)
    task_id = str(uuid.uuid4())
    tasks[task_id] = {'status': 'processing', 'original_file': filename}
    thread = threading.Thread(target=procesar_archivo_en_segundo_plano, args=(filepath, task_id))
    thread.start()
    return jsonify({'task_id': task_id})

@app.route('/status/<task_id>')
def task_status(task_id):
    task = tasks.get(task_id, None)
    if not task: return jsonify({'status': 'not_found'}), 404
    return jsonify(task)

@app.route('/download/<filename>')
def descargar_archivo(filename):
    if os.path.exists(os.path.join(DOWNLOAD_FOLDER, filename)):
        return send_from_directory(DOWNLOAD_FOLDER, filename, as_attachment=True)
    elif os.path.exists(os.path.join(UPLOAD_FOLDER, filename)):
        return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)
    else:
        return jsonify({"error": "Archivo no encontrado"}), 404

if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1')