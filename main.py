import pandas as pd
import re
import os

def normalizar_cuit(cuit):
    """Elimina guiones y espacios de los CUITs."""
    if pd.isna(cuit): return ""
    return re.sub(r'\D', '', str(cuit))

def normalizar_cheque(nro):
    """Limpia el número de cheque quitando ceros a la izquierda y decimales."""
    if pd.isna(nro): return ""
    s = str(nro).split('.')[0].strip()
    return s.lstrip('0')

def normalizar_monto(val):
    """Convierte strings de moneda a float limpio."""
    try:
        if isinstance(val, str):
            val = val.replace('$', '').replace('.', '').replace(',', '.')
        return round(float(val), 2)
    except:
        return 0.0

def procesar_archivos(path_sap, path_banco):
    print(f"\n Cargando archivos...")
    
    # Leer SAP (Hoja 1 por defecto) y Banco (Hoja 1)
    df_sap = pd.read_excel(path_sap)
    df_bank = pd.read_excel(path_banco)

    # Definición de columnas (Ajustar nombres si varían en el archivo real)
    # SAP
    col_nro_sap = 'Número de Cheque'
    col_cuit_sap = 'CUIT Librador'
    col_monto_sap = 'Imp. moneda local'
    
    # BANCO
    col_nro_bank = 'Nro'
    col_cuit_bank = 'CUIT-CUIL CDI'
    col_monto_bank = 'Monto'

    # 1. Limpieza y Normalización
    for df, nro, cuit, monto in [
        (df_sap, col_nro_sap, col_cuit_sap, col_monto_sap),
        (df_bank, col_nro_bank, col_cuit_bank, col_monto_bank)
    ]:
        df['nro_clean'] = df[nro].apply(normalizar_cheque)
        df['cuit_clean'] = df[cuit].apply(normalizar_cuit)
        df['monto_clean'] = df[monto].apply(normalizar_monto)
        df['key'] = df['nro_clean'] + "_" + df['cuit_clean']

    # 2. Búsqueda de Duplicados Internos
    print(df)
    sap_dupes = df_sap[df_sap.duplicated(subset=['key'], keep=False)]
    bank_dupes = df_bank[df_bank.duplicated(subset=['key'], keep=False)]

    # 3. Comparación entre listas
    sap_keys = set(df_sap['key'])
    bank_keys = set(df_bank['key'])

    solo_en_sap = df_sap[~df_sap['key'].isin(bank_keys)]
    solo_en_banco = df_bank[~df_bank['key'].isin(sap_keys)]
    
    # Coincidencias para validar montos
    comunes = pd.merge(
        df_sap[['key', 'nro_clean', 'monto_clean']], 
        df_bank[['key', 'monto_clean']], 
        on='key', suffixes=('_sap', '_bank')
    ).drop_duplicates()

    diff_monto = comunes[abs(comunes['monto_clean_sap'] - comunes['monto_clean_bank']) > 0.01]
    conciliados = comunes[abs(comunes['monto_clean_sap'] - comunes['monto_clean_bank']) <= 0.01]

    # --- REPORTE CLI ---
    print("\n" + "="*40)
    print("      REPORTE DE CONCILIACIÓN")
    print("="*40)
    print(f"Total registros SAP:   {len(df_sap)}")
    print(f"Total registros Banco: {len(df_bank)}")
    print("-" * 40)
    print(f"✅ Conciliados OK:      {len(conciliados)}")
    print(f"❌ Diferencia Monto:    {len(diff_monto)}")
    print(f"⚠️ Solo en SAP:         {len(solo_en_sap)}")
    print(f"⚠️ Solo en Banco:       {len(solo_en_banco)}")
    print("-" * 40)
    
    if len(sap_dupes) > 0:
        print(f"❗ Alerta: SAP tiene {len(sap_dupes)} registros con Nro/CUIT duplicado.")
    if len(bank_dupes) > 0:
        print(f"❗ Alerta: Banco tiene {len(bank_dupes)} registros con Nro/CUIT duplicado.")

    if not diff_monto.empty:
        print("\n--- DETALLE DIFERENCIA DE MONTOS ---")
        print(diff_monto[['key', 'monto_clean_sap', 'monto_clean_bank']])

    if not solo_en_banco.empty:
        print("\n--- CHEQUES EN BANCO QUE NO ESTÁN EN SAP (Primeros 5) ---")
        print(solo_en_banco[['Nro', 'CUIT-CUIL CDI', 'Monto']].head())

if __name__ == "__main__":
    # Interfaz CLI simple
    print("Conciliador de E-Cheques")
    file_sap = input("Ingrese nombre o ruta del archivo SAP (ej: sap.xlsx): ")
    file_bank = input("Ingrese nombre o ruta del archivo BANCO (ej: banco.xlsx): ")
    
    if os.path.exists(file_sap) and os.path.exists(file_bank):
        procesar_archivos(file_sap, file_bank)
    else:
        print("Error: Uno o ambos archivos no existen.")
