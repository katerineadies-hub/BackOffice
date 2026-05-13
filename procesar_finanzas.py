import pandas as pd
import json
import gspread
import os
from google.oauth2.service_account import Credentials
from datetime import datetime

# --- 1. CONFIGURACIÓN DE SEGURIDAD ---
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def obtener_cliente_gc():
    if "GOOGLE_CREDS" in os.environ:
        creds_dict = json.loads(os.environ["GOOGLE_CREDS"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    raise Exception("No se encontró la variable GOOGLE_CREDS en los Secrets de GitHub")

# --- 2. LIMPIADOR DE MONEDA ROBUSTO ---
def limpiar_monto_final(valor):
    if pd.isna(valor) or str(valor).strip() in ["", "0", "0.0", "-", "None"]:
        return 0.0
    
    # Eliminar símbolos de moneda y espacios
    s = str(valor).replace('$', '').replace('S/.', '').replace(' ', '').strip()
    
    # Lógica para detectar si el Excel usa comas para decimales o para miles
    if ',' in s and '.' in s:
        if s.rfind('.') < s.rfind(','): # Formato 1.234,56
            s = s.replace('.', '').replace(',', '.')
        else: # Formato 1,234.56
            s = s.replace(',', '')
    elif ',' in s: # Solo tiene comas
        # Si tiene más de una coma o la coma no está al final, es separador de miles
        if s.count(',') > 1 or (len(s.split(',')[1]) != 2):
            s = s.replace(',', '')
        else: # Es decimal 1234,56
            s = s.replace(',', '.')
            
    try:
        return round(float(s), 2)
    except:
        return 0.0

# --- 3. MOTOR PRINCIPAL ---
def ejecutar_analisis_bi():
    print("🚀 Iniciando extracción de datos detallada...")
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
            print(f"📊 Procesando pestaña: {p_nombre}")
            ws = sh.worksheet(p_nombre)
            data = ws.get_all_values()
            
            # Localizar la fila de encabezado
            idx_h = next(i for i, row in enumerate(data) if "Fecha" in row)
            df = pd.DataFrame.from_records(data[idx_h+1:], columns=data[idx_h])
            
            # Limpiar nombres de columnas
            df.columns = [str(c).strip() for c in df.columns]

            # Identificar columnas de bancos reales
            ignorar = ['Fecha', 'Total', 'Saldo Total Soles', 'Variación', 'Saldo Total', '', 'None']
            cols_bancos = [c for c in df.columns if c and c not in ignorar and "Unnamed" not in c]

            bancos_detalle = []
            for col in cols_bancos:
                # Obtener los últimos dos registros válidos de este banco
                serie_banco = df[col].replace('', pd.NA).dropna()
                
                if not serie_banco.empty:
                    monto_hoy = limpiar_monto_final(serie_banco.iloc[-1])
                    monto_ayer = limpiar_monto_final(serie_banco.iloc[-2]) if len(serie_banco) > 1 else monto_hoy
                    
                    variacion = ((monto_hoy - monto_ayer) / monto_ayer * 100) if monto_ayer != 0 else 0
                    
                    if monto_hoy != 0: # Solo agregar si tiene saldo
                        bancos_detalle.append({
                            "nombre": col,
                            "saldo": monto_hoy,
                            "variacion": variacion
                        })

            total_empresa = sum(b['saldo'] for b in bancos_detalle)
            
            analisis_web[p_id] = {
                "saldo_actual": total_empresa,
                "bancos": sorted(bancos_detalle, key=lambda x: x['saldo'], reverse=True)
            }

        # --- 4. GUARDADO DE ARCHIVOS (ACTUAL + HISTORIAL) ---
        
        # Guardar archivo actual para la web
        with open('financial_data.json', 'w') as f:
            json.dump(analisis_web, f, indent=4)
        
        # Guardar copia en historial
        fecha_hoy = datetime.now().strftime('%Y-%m-%d')
        if not os.path.exists('history'):
            os.makedirs('history')
            
        with open(f'history/data_{fecha_hoy}.json', 'w') as f:
            json.dump(analisis_web, f, indent=4)
        
        print(f"✅ Proceso terminado. Archivo de hoy y copia en historial ({fecha_hoy}) generados.")

    except Exception as e:
        print(f"❌ ERROR CRÍTICO: {str(e)}")
        exit(1)

if __name__ == "__main__":
    ejecutar_analisis_bi()
