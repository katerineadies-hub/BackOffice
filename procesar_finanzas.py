import pandas as pd
import json
import gspread
import os
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

# --- 1. CONFIGURACIÓN DE SEGURIDAD ---
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def obtener_cliente_gc():
    # El robot buscará la llave secreta que guardaste en GitHub
    if "GOOGLE_CREDS" in os.environ:
        creds_dict = json.loads(os.environ["GOOGLE_CREDS"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    else:
        # Esto solo se activa si intentas correrlo en tu PC localmente
        raise Exception("No se encontró la variable GOOGLE_CREDS")
    
    return gspread.authorize(creds)

# --- 2. LÓGICA DE LIMPIEZA DE MONTOS ---
def limpiar_monto_pro(valor):
    if pd.isna(valor) or str(valor).strip() in ["", "0", "0.0", "-"]:
        return 0.0
    s = str(valor).replace('$', '').replace('S/.', '').replace(' ', '').strip()
    if '.' in s and ',' in s:
        if s.rfind('.') < s.rfind(','): s = s.replace('.', '').replace(',', '.')
        else: s = s.replace(',', '')
    elif ',' in s: s = s.replace(',', '.')
    try: return round(float(s), 2)
    except: return 0.0

# --- 3. MOTOR DE ANÁLISIS (SIN ELEMENTOS VISUALES) ---
def ejecutar_analisis_bi():
    print("🚀 Iniciando proceso de análisis...")
    gc = obtener_cliente_gc()
    analisis_web = {}

    try:
        sh = gc.open("Financial Position")
        pestañas = {
            "Resumen Posición Financiera CF": "Bold_CF",
            "Resumen Posición Financiera CO SAS": "Bold_SAS",
            "Bold Perú": "Bold_Peru"
        }

        for p_nombre, p_id in pestañas.items():
            ws = sh.worksheet(p_nombre)
            data = ws.get_all_values()
            
            # Encontrar encabezado
            idx_h = next(i for i, row in enumerate(data) if "Fecha" in row)
            df = pd.DataFrame.from_records(data[idx_h+1:], columns=data[idx_h])
            df.columns = [str(c).strip() for c in df.columns]
            df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce')
            df = df.dropna(subset=['Fecha']).sort_values('Fecha')

            # Columnas de bancos
            cols_ignore = ['Fecha', 'Total', 'Saldo Total Soles', 'Variación', 'Saldo Total', '']
            cols_bancos = [c for c in df.columns if c not in cols_ignore and not c.startswith('Unnamed')]

            for c in cols_bancos:
                df[c] = df[c].apply(limpiar_monto_pro)

            df['Total_Empresa'] = df[cols_bancos].sum(axis=1)
            df_real = df[df['Total_Empresa'] > 0].copy()

            if not df_real.empty:
                serie = df_real['Total_Empresa']
                hoy_e = serie.iloc[-1]
                ayer_e = serie.iloc[-2] if len(serie) > 1 else hoy_e
                hace_7_e = serie.iloc[-8] if len(serie) > 7 else ayer_e

                analisis_web[p_id] = {
                    "saldo_actual": float(hoy_e),
                    "d_1_variacion": float(((hoy_e - ayer_e) / ayer_e * 100)) if ayer_e != 0 else 0,
                    "d_7_variacion": float(((hoy_e - hace_7_e) / hace_7_e * 100)) if hace_7_e != 0 else 0,
                    "promedio_30d": float(serie.tail(30).mean()),
                    "min_30d": float(serie.tail(30).min()),
                    "max_30d": float(serie.tail(30).max())
                }
                print(f"✅ Datos procesados para: {p_nombre}")

        # Guardar el JSON que consumirá tu página web
        with open('financial_data.json', 'w') as f:
            json.dump(analisis_web, f, indent=4)
        
        print("📊 Archivo financial_data.json generado con éxito.")

    except Exception as e:
        print(f"❌ Error: {e}")
        exit(1) # Reporta el error a GitHub

if __name__ == "__main__":
    ejecutar_analisis_bi()
