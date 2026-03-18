from flask import Flask, render_template, request, send_file, session
import pandas as pd
import re
import io
import uuid

app = Flask(__name__)
app.secret_key = "conciliacion_secret_key" # Necesario para guardar datos temporales

# Diccionario temporal para guardar los archivos generados (en un entorno real usarías Redis o DB)
cache_archivos = {}

# --- CONFIGURACIÓN DE COLUMNAS ---
COLUMNAS_SAP = ('Número de Cheque', 'CUIT Librador', 'Imp. moneda local')
COLUMNAS_BANCO = ('Nro', 'CUIT-CUIL CDI', 'CUIT Endosante', 'Monto')

def normalizar_cuit(cuit):
    if pd.isna(cuit): return ""
    return re.sub(r'\D', '', str(cuit))

def normalizar_cheque(nro):
    if pd.isna(nro): return ""
    s = str(nro).split('.')[0].strip()
    return s.lstrip('0')

def normalizar_monto(val):
    try:
        if isinstance(val, str):
            val = val.replace('$', '').replace('.', '').replace(',', '.')
        return round(float(val), 2)
    except:
        return 0.0

def validar_columnas(df, columnas_requeridas, nombre_archivo):
    columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
    if columnas_faltantes:
        return f"Error en '{nombre_archivo}': faltan las columnas {columnas_faltantes}"
    return None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file_sap = request.files.get('file_sap')
    file_bank = request.files.get('file_bank')

    if not file_sap or not file_bank:
        return render_template('index.html', error="Debes subir ambos archivos.")

    try:
        df_sap = pd.read_excel(file_sap)
        df_bank = pd.read_excel(file_bank)

        # 1. Verificación de formato (Columnas)
        error_sap = validar_columnas(df_sap, COLUMNAS_SAP, "Excel SAP")
        error_bank = validar_columnas(df_bank, COLUMNAS_BANCO, "Excel Banco")

        if error_sap or error_bank:
            return render_template('index.html', error=(error_sap or "") + " " + (error_bank or ""))

        # 2. Procesamiento

        nro_sap, cuit_sap, monto_sap = COLUMNAS_SAP
        df_sap['nro_clean'] = df_sap[nro_sap].apply(normalizar_cheque)
        df_sap['cuit_clean'] = df_sap[cuit_sap].apply(normalizar_cuit)
        df_sap['monto_clean'] = df_sap[monto_sap].apply(normalizar_monto)
        df_sap['key'] = df_sap['nro_clean'] + "_" + df_sap['cuit_clean']

        nro_bk, cuit_bk, cuit_endo, monto_bk = COLUMNAS_BANCO
        df_bank['nro_clean'] = df_bank[nro_bk].apply(normalizar_cheque)

        def elegir_cuit(fila):
            # pd.notna() detecta NaNs, y strip() != "" detecta celdas con espacios
            if pd.notna(fila[cuit_endo]) and str(fila[cuit_endo]).strip() not in ["", "-"]:
                return normalizar_cuit(fila[cuit_endo])
            return normalizar_cuit(fila[cuit_bk])

        df_bank['cuit_clean'] = df_bank.apply(elegir_cuit, axis=1)
        df_bank['monto_clean'] = df_bank[monto_bk].apply(normalizar_monto)
        df_bank['key'] = df_bank['nro_clean'] + "_" + df_bank['cuit_clean']

        df_merge = pd.merge(df_sap, df_bank, on='key', suffixes=('_SAP', '_BANCO'))

        conciliados = df_merge[abs(df_merge['monto_clean_SAP'] - df_merge['monto_clean_BANCO']) <= 0.01].copy()
        solo_sap = df_sap[~df_sap['key'].isin(df_bank['key'])].copy()
        solo_banco = df_bank[~df_bank['key'].isin(df_sap['key'])].copy()
        diff_monto = df_merge[abs(df_merge['monto_clean_SAP'] - df_merge['monto_clean_BANCO']) > 0.01].copy()

        pendientes = pd.concat([solo_banco, diff_monto], ignore_index=True)

        # 3. Guardar resumen para mostrar
        resumen = {
            'total_sap': len(df_sap),
            'total_banco': len(df_bank),
            'conciliados': len(conciliados),
            'solo_banco': len(solo_banco),
            'dif_monto': len(diff_monto)
        }

        # 4. Generar Excel y guardar en "cache" temporal
        file_id = str(uuid.uuid4())
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            conciliados.to_excel(writer, sheet_name='Conciliados', index=False)
            pendientes.to_excel(writer, sheet_name='Pendientes', index=False)
        output.seek(0)
        cache_archivos[file_id] = output.read()

        return render_template('index.html', resumen=resumen, file_id=file_id)

    except Exception as e:
        return render_template('index.html', error=f"Error procesando archivos: {str(e)}")

@app.route('/download/<file_id>')
def download(file_id):
    if file_id in cache_archivos:
        return send_file(
            io.BytesIO(cache_archivos[file_id]),
            as_attachment=True,
            download_name="Resultado_Conciliacion.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return "Archivo no encontrado o expirado", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
