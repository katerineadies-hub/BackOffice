import pandas as pd
import json
import gspread
import os
from google.oauth2.service_account import Credentials

scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def obtener_cliente_gc():
    if "GOOGLE_CREDS" in os.environ:
        creds_dict = json.loads(os.environ["GOOGLE_CREDS"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    else:
        raise Exception("No se encontró la variable GOOGLE_CREDS")
    return gspread.authorize(creds)

def limpiar_monto_pro(valor):
    if pd.isna(valor) or str(valor).strip() in ["", "0", "0.0", "-"]:
        return 0.0
    s = str(valor).replace('$', '').replace('S/.', '').replace(' ', '').replace(',', '').strip()
    try: return round(float(s), 2)
    except: return 0.0

def ejecutar_analisis_bi():
    gc = obtener_cliente_gc()
    analisis_web = {}
    sh = gc.open("Financial Position")
    
    pestañas = {
        "Resumen Posición Financiera CF": "Bold_CF",
        "Resumen Posición Financiera CO SAS": "Bold_SAS",
        "Bold Perú": "Bold_Peru"
    }

    for p_nombre, p_id in pestañas.items():
        ws = sh.worksheet(p_nombre)
        data = ws.get_all_values()
        idx_h = next(i for i, row in enumerate(data) if "Fecha" in row)
        df = pd.DataFrame.from_records(data[idx_h+1:], columns=data[idx_h])
        df['Fecha'] = pd.to_datetime(df['Fecha'], dayfirst=True, errors='coerce')
        df = df.dropna(subset=['Fecha']).sort_values('Fecha')

        cols_ignore = ['Fecha', 'Total', 'Saldo Total Soles', 'Variación', 'Saldo Total', '']
        cols_bancos = [c for c in df.columns if c not in cols_ignore and not c.startswith('Unnamed') and c.strip() != ""]

        detalle_bancos = []
        for col in cols_bancos:
            df[col] = df[col].apply(limpiar_monto_pro)
            serie = df[col]
            if len(serie) >= 2:
                hoy = serie.iloc[-1]
                ayer = serie.iloc[-2]
                var = ((hoy - ayer) / ayer * 100) if ayer != 0 else 0
                if hoy != 0: # Solo mostrar bancos con saldo
                    detalle_bancos.append({
                        "nombre": col,
                        "saldo": float(hoy),
                        "variacion": float(var)
                    })

        total_hoy = sum(b['saldo'] for b in detalle_bancos)
        # Ordenar bancos de mayor a menor saldo
        detalle_bancos = sorted(detalle_bancos, key=lambda x: x['saldo'], reverse=True)

        analisis_web[p_id] = {
            "saldo_actual": total_hoy,
            "bancos": detalle_bancos
        }

    with open('financial_data.json', 'w') as f:
        json.dump(analisis_web, f, indent=4)

if __name__ == "__main__":
    ejecutar_analisis_bi()
