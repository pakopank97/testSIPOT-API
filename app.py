import pandas as pd
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import json
import re

app = Flask(__name__)
CORS(app)

# --- Funciones de Validación (sin cambios) ---
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

@app.route('/')
def servir_pagina_principal():
    return send_from_directory('.', 'index.html')

@app.route('/upload', methods=['POST'])
def procesar_excel():
    if 'archivo' not in request.files: return jsonify({"error": "No se encontró el archivo"}), 400
    archivo = request.files['archivo']
    if archivo.filename == '': return jsonify({"error": "No se seleccionó archivo"}), 400

    try:
        df = pd.read_excel(archivo, sheet_name=0, header=None, dtype=str)

        # --- LÓGICA DE VALIDACIÓN Y ADVERTENCIAS ---
        lista_de_errores = []
        lista_de_advertencias = [] # <-- NUEVA LISTA
        
        reglas = df.iloc[3].fillna('0').tolist() if len(df) > 3 else []
        headers = list(df.iloc[6].fillna('')) if len(df) > 6 else [] # Usamos '' para cabeceras vacías
        datos_a_validar = df.iloc[7:]

        for fila_idx, fila in datos_a_validar.iterrows():
            for col_idx, valor in enumerate(fila):
                
                # --- NUEVO: LÓGICA DE ADVERTENCIAS PARA CELDAS VACÍAS ---
                # Condición: La celda está vacía PERO su cabecera no lo está.
                if (pd.isna(valor) or str(valor).strip() == '') and (col_idx < len(headers) and str(headers[col_idx]).strip() != ''):
                    coordenada = obtener_coordenada_excel(fila_idx, col_idx)
                    header_name = headers[col_idx]
                    advertencia = f"La celda {coordenada} (columna '{header_name}') está vacía."
                    lista_de_advertencias.append(advertencia)

                # --- LÓGICA DE ERRORES DE FORMATO (para celdas no vacías) ---
                elif pd.notna(valor) and str(valor).strip() != '':
                    regla = str(reglas[col_idx]).strip().split('.')[0] if col_idx < len(reglas) else '0'
                    if regla in VALIDADORES:
                        validador = VALIDADORES[regla]
                        if not validador['func'](valor):
                            coordenada = obtener_coordenada_excel(fila_idx, col_idx)
                            mensaje = f"Error de formato en celda {coordenada} ('{valor}'). Se esperaba: {validador['nombre']}."
                            lista_de_errores.append(mensaje)

        # --- EXTRACCIÓN DE DATOS Y JSONs (sin cambios) ---
        formato = df.iloc[0, 0] if not df.empty else "Formato no encontrado"
        datos_filas = df.iloc[7:].fillna('').values.tolist()
        
        headers_backend = df.iloc[4].fillna('header_desconocido').astype(str).tolist() if len(df) > 4 else []
        datos_backend_df = df.iloc[7:]
        data_intercalada = []
        if not datos_backend_df.empty and headers_backend:
            datos_backend_df.columns = headers_backend[:len(datos_backend_df.columns)]
            data_intercalada = datos_backend_df.fillna('').to_dict(orient='records')
        jsonBACKEND = {"id_formato": str(formato), "data": data_intercalada}
        
        print("\n[INFO] ======== jsonBACKEND (para el sistema) ========")
        print(json.dumps(jsonBACKEND, indent=2, ensure_ascii=False))
        print("=====================================================")

        # --- RESPUESTA PARA EL FRONTEND (ahora incluye las advertencias) ---
        respuesta_final = { 
            "formato": str(formato), 
            "headers": headers, 
            "rows": datos_filas, 
            "errors": lista_de_errores,
            "warnings": lista_de_advertencias, # <-- NUEVO
            "json_backend": jsonBACKEND
        }
        return jsonify(respuesta_final)

    except Exception as e:
        return jsonify({"error": f"Error al procesar el archivo: {e}"}), 500

if __name__ == '__main__':
    app.run(debug=True)