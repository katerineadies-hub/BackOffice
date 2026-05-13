import pandas as pd
import json
import gspread
import os
from google.oauth2.service_account import Credentials

# --- CONFIGURACIÓN ---
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def obtener_cliente_gc():
    if "GOOGLE_CREDS" in os.environ:
        creds_dict = json.loads(os.environ["GOOGLE_CREDS"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    raise Exception("Llave secreta no encontrada")

def limpiar_monto_ultra(valor):
    """Limpia montos manejando cualquier formato de miles/decimales"""
    if pd.isna(valor) or str(valor).strip() in ["", "0", "0.0", "-", "None"]:
        return 0.0
    
    # Convertir a string y limpiar símbolos de moneda
    s = str(valor).replace('$', '').replace('S/.', '').replace(' ', '').strip()
    
    # Lógica inteligente de separadores
    if ',' in s and '.' in s:
        if s.rfind('.') < s.rfind(','): # Formato 1.234,56
            s = s.replace('.', '').replace(',', '.')
        else: # Formato 1,234.56
            s = s.replace(',', '')
    elif ',' in s: # Solo comas
        # Si hay más de una o está en posición de miles
        if s.count(',') > 1 or (len(s.split(',')[1]) != 2):
            s = s.replace(',', '')
        else: # Es decimal 1234,56
            s = s.replace(',', '.')
            
    try:
        return round(float(s), 2)
    except:
        return 0.0

def ejecutar_analisis_bi():
    print("🚑 Iniciando reparación de datos...")
    gc = obtener_cliente_gc()
    analisis_web = {}
    sh = gc.open("Financial Position")
    
    pestañas = {
        "Resumen Posición Financiera CF": "Bold_CF",
        "Resumen Posición Financiera CO SAS": "Bold_SAS",
        "Bold Perú": "Bold_Peru"
    }

    for p_nombre, p_id in pestañas.items():
        print(f"Procesando {p_nombre}...")
        ws = sh.worksheet(p_nombre)
        data = ws.get_all_values()
        
        # Encontrar cabecera
        idx_h = next(i for i, row in enumerate(data) if "Fecha" in row)
        df = pd.DataFrame.from_records(data[idx_h+1:], columns=data[idx_h])
        
        # Identificar columnas de bancos (todo lo que no sea Fecha/Total/etc)
        ignorar = ['Fecha', 'Total', 'Saldo Total Soles', 'Variación', 'Saldo Total', '']
        cols_bancos = [c for c in df.columns if c and c.strip() not in ignorar and "Unnamed" not in c]

        bancos_detalle = []
        for col in cols_bancos:
            # Tomar el último valor no vacío de la columna
            col_data = df[col].replace('', pd.NA).dropna()
            if not col_data.empty:
                monto_raw = col_data.iloc[-1]
                monto_clean = limpiar_monto_ultra(monto_raw)
                
                # Calcular variación simple vs penúltimo registro
                monto_ayer = limpiar_monto_ultra(col_data.iloc[-2]) if len(col_data) > 1 else monto_clean
                var = ((monto_clean - monto_ayer) / monto_ayer * 100) if monto_ayer != 0 else 0
                
                if monto_clean != 0:
                    bancos_detalle.append({
                        "nombre": col.strip(),
                        "saldo": monto_clean,
                        "variacion": var
                    })

        total = sum(b['saldo'] for b in bancos_detalle)
        analisis_web[p_id] = {
            "saldo_actual": total,
            "bancos": sorted(bancos_detalle, key=lambda x: x['saldo'], reverse=True)
        }

    with open('financial_data.json', 'w') as f:
        json.dump(analisis_web, f, indent=4)
    print("✅ Reparación completada.")

if __name__ == "__main__":
    ejecutar_analisis_bi()
