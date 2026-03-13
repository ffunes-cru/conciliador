import pandas as pd
import re
from flask import Flask, render_template, request, send_file
import os
import io
#import argparse

app = Flask(__name__)

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

def exportar_a_excel(df_conciliados, df_pendientes, nombre_archivo="Resultado_Conciliacion.xlsx"):
    """Crea un Excel con dos hojas: Conciliados y Pendientes."""
    try:
        with pd.ExcelWriter(nombre_archivo, engine='openpyxl') as writer:
            df_conciliados.to_excel(writer, sheet_name='Conciliados', index=False)
            df_pendientes.to_excel(writer, sheet_name='Pendientes de Conciliar', index=False)
        print(f"\n✅ Archivo exportado con éxito: {nombre_archivo}")
    except Exception as e:
        print(f"\n❌ Error al exportar Excel: {e}")

def procesar_archivos(path_sap, path_banco):
    print(f"\nLeyendo archivos...")
    df_sap = pd.read_excel(path_sap)
    df_bank = pd.read_excel(path_banco)

    # Nombres de columnas según tus archivos
    col_nro_sap, col_cuit_sap, col_monto_sap = 'Número de Cheque', 'CUIT Librador', 'Imp. moneda local'
    col_nro_bank, col_cuit_bank, col_monto_bank = 'Nro', 'CUIT-CUIL CDI', 'Monto'

    # Normalización
    for df, nro, cuit, monto in [(df_sap, col_nro_sap, col_cuit_sap, col_monto_sap), 
                                 (df_bank, col_nro_bank, col_cuit_bank, col_monto_bank)]:
        df['nro_clean'] = df[nro].apply(normalizar_cheque)
        df['cuit_clean'] = df[cuit].apply(normalizar_cuit)
        df['monto_clean'] = df[monto].apply(normalizar_monto)
        df['key'] = df['nro_clean'] + "_" + df['cuit_clean']
    print(df)
    # 1. Identificar Conciliados (Coinciden en Key y Monto)
    df_merge = pd.merge(df_sap, df_bank, on='key', suffixes=('_SAP', '_BANCO'))
    print(df_merge)

    conciliados = df_merge[abs(df_merge['monto_clean_SAP'] - df_merge['monto_clean_BANCO']) <= 0.01].copy()
    conciliados['Estado_Conciliacion'] = 'OK - Conciliado'

    # 2. Identificar Pendientes
    # Solo en SAP
    solo_sap = df_sap[~df_sap['key'].isin(df_bank['key'])].copy()
    solo_sap['Estado_Conciliacion'] = 'Falta en BANCO'
    
    # Solo en BANCO
    solo_banco = df_bank[~df_bank['key'].isin(df_sap['key'])].copy()
    solo_banco['Estado_Conciliacion'] = 'Falta en SAP'
    
    # Diferencia de Monto (Misma key, distinto importe)
    diff_monto = df_merge[abs(df_merge['monto_clean_SAP'] - df_merge['monto_clean_BANCO']) > 0.01].copy()
    diff_monto['Estado_Conciliacion'] = 'Diferencia de Importe'

    # Unir todos los pendientes en un solo DataFrame
    pendientes = pd.concat([solo_banco, diff_monto], ignore_index=True)

    # Ejecutar Exportación
    #exportar_a_excel(conciliados, pendientes)

    # Crear el archivo Excel en memoria (buffer)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        conciliados.to_excel(writer, sheet_name='Conciliados', index=False)
        pendientes.to_excel(writer, sheet_name='Pendientes', index=False)
    output.seek(0)

    return output

    # Resumen rápido en consola
    #print(f"--- Resumen ---")
    #print(f"Conciliados: {len(conciliados)}")
    #print(f"Pendientes:  {len(solo_banco)}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file_sap = request.files.get('file_sap')
    file_bank = request.files.get('file_bank')

    if not file_sap or not file_bank:
        return "Error: Debes subir ambos archivos.", 400

    # Procesar y obtener el archivo resultante
    excel_stream = procesar_archivos(file_sap, file_bank)

    return send_file(
        excel_stream,
        as_attachment=True,
        download_name="Resultado_Conciliacion.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == '__main__':
    app.run(debug=True)

#if __name__ == "__main__":
#    parser = argparse.ArgumentParser(description="Programa de conciliación de cheques SAP vs Banco.")
#
#    # Definición de banderas
#    parser.add_argument("--sap", required=True, help="Ruta al archivo Excel de SAP")
#    parser.add_argument("--banco", required=True, help="Ruta al archivo Excel del Banco")
#    #parser.add_argument("--out", default="Resultado_Conciliacion.xlsx", help="Nombre del archivo de salida (opcional)")
#
#    args = parser.parse_args()
#
#    if not os.path.exists(args.sap):
#        print(f"❌ El archivo SAP no existe: {args.sap}")
#    elif not os.path.exists(args.banco):
#        print(f"❌ El archivo Banco no existe: {args.banco}")
#    else:
#        procesar_archivos(args.sap, args.banco)
