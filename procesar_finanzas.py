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
    s = str(valor).replace('$', '').replace('S/.', '').replace(' ', '').strip()
    if ',' in s and '.' in s:
        if s.rfind('.') < s.rfind(','): s = s.replace('.', '').replace(',', '.')
        else: s = s.replace(',', '')
    elif ',' in s:
        if s.count(',') > 1 or (len(s.split(',')[1]) != 2): s = s.replace(',', '')
        else: s = s.replace(',', '.')
    try: return round(float(s), 2)
    except: return 0.0

# --- 3. MOTOR PRINCIPAL ---
def ejecutar_analisis_bi():
    print("🚀 Iniciando extracción de datos de Tesorería...")
    gc = obtener_cliente_gc()
    analisis_web = {}
    
    try:
        # --- PARTE A: POSICIÓN FINANCIERA COMPAÑÍAS (Archivo Maestro) ---
        print("🏢 Procesando saldos consolidados de compañías...")
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
            df.columns = [str(c).strip() for c in df.columns]

            ignorar = ['Fecha', 'Total', 'Saldo Total Soles', 'Variación', 'Saldo Total', '', 'None']
            cols_bancos = [c for c in df.columns if c and c not in ignorar and "Unnamed" not in c]

            bancos_detalle = []
            for col in cols_bancos:
                serie_banco = df[col].replace('', pd.NA).dropna()
                if not serie_banco.empty:
                    monto_hoy = limpiar_monto_final(serie_banco.iloc[-1])
                    monto_ayer = limpiar_monto_final(serie_banco.iloc[-2]) if len(serie_banco) > 1 else monto_hoy
                    variacion = ((monto_hoy - monto_ayer) / monto_ayer * 100) if monto_ayer != 0 else 0
                    if monto_hoy != 0:
                        bancos_detalle.append({"nombre": col, "saldo": monto_hoy, "variacion": variacion})

            total_empresa = sum(b['saldo'] for b in bancos_detalle)
            analisis_web[p_id] = {
                "saldo_actual": total_empresa,
                "bancos": sorted(bancos_detalle, key=lambda x: x['saldo'], reverse=True)
            }

        # --- PARTE B: AUTOMATISMO MONITOR CUD (Archivo Independiente) ---
        print("🏛️ Conectando con archivo independiente MONITOREO CUD BANREP...")
        try:
            # Abrimos usando el ID único extraído de tu enlace para máxima precisión
            sh_cud = gc.open_by_key("1Xtq1YR9PVBVmztSt6y8Oi72QTe_TDE6cVQc4il0itVA")
            ws_cud = sh_cud.sheet1 # Toma automáticamente la primera pestaña (gid=0)
            data_cud = ws_cud.get_all_values()
            
            # Crear DataFrame (la fila 1 contiene los encabezados: Fecha, Concepto, Entradas, Salidas, Saldo)
            df_cud = pd.DataFrame(data_cud[1:], columns=data_cud[0])
            df_cud.columns = [str(c).strip() for c in df_cud.columns]
            
            # Limpiar y estandarizar montos numéricos
            df_cud['Entradas'] = df_cud['Entradas'].apply(limpiar_monto_final)
            df_cud['Salidas'] = df_cud['Salidas'].apply(limpiar_monto_final)
            df_cud['Saldo'] = df_cud['Saldo'].apply(limpiar_monto_final)
            
            # Filtrar registros que no tengan fecha válida
            df_cud = df_cud[df_cud['Fecha'].str.strip() != ""]
            
            if not df_cud.empty:
                # El saldo actual se ubica en la última línea registrada
                ultima_fila = df_cud.iloc[-1]
                ultima_fecha = ultima_fila['Fecha']
                
                # Consolidar los flujos totales únicamente de la última jornada cargada
                df_ultimo_dia = df_cud[df_cud['Fecha'] == ultima_fecha]
                total_entradas_dia = float(df_ultimo_dia['Entradas'].sum())
                total_salidas_dia = float(df_ultimo_dia['Salidas'].sum())
                
                # Extraer los últimos 10 movimientos para la visualización de auditoría
                ultimos_movimientos = []
                for _, row in df_cud.tail(10).iterrows():
                    ultimos_movimientos.append({
                        "fecha": row['Fecha'],
                        "concepto": row['Concepto'],
                        "entradas": float(row['Entradas']),
                        "salidas": float(row['Salidas']),
                        "saldo": float(row['Saldo'])
                    })
                
                # Inyectar estructura limpia al diccionario de la web
                analisis_web["CUD"] = {
                    "saldo_actual": float(ultima_fila['Saldo']),
                    "entradas_hoy": total_entradas_dia,
                    "salidas_hoy": total_salidas_dia,
                    "fecha_actualizacion": ultima_fecha,
                    "historial": list(reversed(ultimos_movimientos))
                }
            else:
                analisis_web["CUD"] = {"saldo_actual": 0.0, "entradas_hoy": 0.0, "salidas_hoy": 0.0, "historial": []}
        except Exception as e_cud:
            print(f"⚠️ Alerta CUD (No se pudo procesar el archivo CUD separado): {str(e_cud)}")
            analisis_web["CUD"] = {"saldo_actual": 0.0, "entradas_hoy": 0.0, "salidas_hoy": 0.0, "historial": []}

        # --- 4. GUARDADO Y SINCRONIZACIÓN ---
        with open('financial_data.json', 'w') as f:
            json.dump(analisis_web, f, indent=4)
        
        fecha_hoy = datetime.now().strftime('%Y-%m-%d')
        if not os.path.exists('history'): 
            os.makedirs('history')
        with open(f'history/data_{fecha_hoy}.json', 'w') as f:
            json.dump(analisis_web, f, indent=4)
        
        print("✅ Consolidación de fuentes completada con éxito.")

    except Exception as e:
        print(f"❌ ERROR CRÍTICO GENERAL: {str(e)}")
        exit(1)

if __name__ == "__main__":
    ejecutar_analisis_bi()
