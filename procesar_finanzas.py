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

def limpiar_monto_final(valor):
    if pd.isna(valor) or str(valor).strip() in ["", "0", "0.0", "-", "None"]:
        return 0.0
    s = str(valor).replace('$', '').replace('S/.', '').replace(' ', '').replace('.', '').strip()
    if ',' in s:
        s = s.replace(',', '.') # Convertir decimales a formato Python standard
    try: return round(float(s), 2)
    except: return 0.0

# --- 2. MOTOR DE PARSEO DE EXCEL CUD (REGLAS MAYA) ---
def procesar_excel_cud_raw(ruta_archivo):
    print("🧠 Ejecutando Motor Maya para el análisis del extracto CUD...")
    try:
        # Leer el archivo sin cabeceras iniciales para evaluar casos
        df_raw = pd.read_excel(ruta_archivo, header=None)
        
        # Determinar si es Caso A o Caso B buscando palabras clave en las primeras filas
        contenido_primeras_filas = " ".join(df_raw.iloc[0:20, 0].dropna().astype(str).tolist())
        
        if "consultas de movimientos" in contenido_primeras_filas.lower() or len(df_raw) > 56:
            print("📋 Detectado Caso A: Reporte consultas de movimientos")
            # Eliminar filas de abajo (37-56 en escala relativa desde el final si aplica, o por rangos)
            # Para evitar errores de desborde, localizamos dinámicamente donde termina la data real o limpiamos índices fijados
            df_parsed = df_raw.copy()
            if len(df_parsed) > 56:
                df_parsed = df_parsed.drop(range(36, min(56, len(df_parsed))))
            df_parsed = df_parsed.drop(range(0, 18)).reset_index(drop=True)
            
            # Asignar cabecera de la primera fila útil
            df_parsed.columns = df_parsed.iloc[0]
            df_parsed = df_parsed[1:].reset_index(drop=True)
            
        else:
            print("📑 Detectado Caso B: Reporte extracto cuenta")
            # Buscar la fila de cabecera que contiene "SECUEN."
            idx_cabecera = next(i for i, row in df_raw.iterrows() if row.astype(str).str.contains("SECUEN.").any())
            df_parsed = pd.read_excel(ruta_archivo, skiprows=idx_cabecera)
            
        # Normalizar nombres de columnas a minúsculas y quitar espacios
        df_parsed.columns = [str(c).strip().lower() for c in df_parsed.columns]
        
        # Filtrar solo filas con secuencia numérica válida
        col_secuencia = [c for c in df_parsed.columns if "secuen" in c][0]
        df_parsed = df_parsed[pd.to_numeric(df_parsed[col_secuencia], errors='coerce').notna()]
        
        # Columnas obligatorias del protocolo Maya
        columnas_finales = [
            'secuen.', 'fecha valor', 'fecha liq.', 'suc.', 'trans.', 'port.', 
            'valor débito', 'valor crédito', 'cuenta contraparte', 'portafolio contraparte', 
            'referencia', 'pormenor', 'id. tro. origen', 'id. tro destino', 'nombre tro.', 'usuario de aprobación'
        ]
        
        # Ajustar dinámicamente columnas faltantes con valores vacíos
        for col in columnas_finales:
            if col not in df_parsed.columns:
                # Mapeos lógicos alternativos (ej: concepto -> pormenor)
                if col == 'pormenor' and 'concepto' in df_parsed.columns:
                    df_parsed['pormenor'] = df_parsed['concepto']
                elif col == 'pormenor' and 'descripción transacción' in df_parsed.columns:
                    df_parsed['pormenor'] = df_parsed['descripción transacción']
                else:
                    df_parsed[col] = ""
                    
        df_final = df_parsed[columnas_finales].copy()
        
        # Limpieza de textos y GMF según protocolo
        df_final['pormenor'] = df_final['pormenor'].astype(str).apply(lambda x: "GMF" if "GMF" in x else x)
        
        # Limpieza y conversión contable de montos
        df_final['valor débito'] = df_final['valor débito'].apply(limpiar_monto_final)
        df_final['valor crédito'] = df_final['valor crédito'].apply(limpiar_monto_final)
        
        return df_final
    except Exception as e:
        print(f"❌ Error en el Motor Maya al parsear el Excel: {str(e)}")
        return pd.DataFrame()

# --- 3. PROCESAMIENTO GENERAL ---
def ejecutar_analisis_bi():
    print("🚀 Iniciando extracción de datos de Tesorería...")
    gc = obtener_cliente_gc()
    analisis_web = {}
    
    try:
        # --- PARTE A: POSICIÓN FINANCIERA COMPAÑÍAS ---
        sh = gc.open("Financial Position")
        pestañas = {"Resumen Posición Financiera CF": "Bold_CF", "Resumen Posición Financiera CO SAS": "Bold_SAS", "Bold Perú": "Bold_Peru"}

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

            analisis_web[p_id] = {"saldo_actual": sum(b['saldo'] for b in bancos_detalle), "bancos": sorted(bancos_detalle, key=lambda x: x['saldo'], reverse=True)}

        # --- PARTE B: AUTOMATISMO PARSER CUD (SIN DEPENDER DE GOOGLE SHEETS) ---
        ruta_cud_excel = "cud_raw.xlsx"
        if os.path.exists(ruta_cud_excel):
            df_cud = procesar_excel_cud_raw(ruta_cud_excel)
            if not df_cud.empty:
                total_debitos = float(df_cud['valor débito'].sum())
                total_creditos = float(df_cud['valor crédito'].sum())
                
                # Asumimos cálculo acumulado (si no hay saldo inicial explícito lo tomamos base)
                saldo_calculado = total_creditos - total_debitos
                fecha_reporte = datetime.now().strftime('%d/%m/%Y')
                
                historial_movimientos = []
                for _, row in df_cud.tail(15).iterrows():
                    historial_movimientos.append({
                        "fecha": str(row['fecha valor']),
                        "concepto": str(row['pormenor']),
                        "entradas": float(row['valor crédito']),
                        "salidas": float(row['valor débito']),
                        "saldo": float(row['valor crédito'] - row['valor débito']) 
                    })
                
                analisis_web["CUD"] = {
                    "saldo_actual": saldo_calculado,
                    "entradas_hoy": total_creditos,
                    "salidas_hoy": total_debitos,
                    "fecha_actualizacion": fecha_reporte,
                    "historial": historial_movimientos
                }
                
                # --- CONTROL DE SALDOS HISTÓRICOS EN LA NUBE ---
                ruta_historico_central = "history/cud_historical.json"
                historico_data = []
                if os.path.exists(ruta_historico_central):
                    with open(ruta_historico_central, 'r') as h_f:
                        try: historico_data = json.load(h_f)
                        except: historico_data = []
                
                # Evitar duplicar la misma fecha
                historico_data = [h for h in historico_data if h['fecha'] != fecha_reporte]
                historico_data.append({
                    "fecha": fecha_reporte,
                    "saldo": saldo_calculado,
                    "entradas": total_creditos,
                    "salidas": total_debitos
                })
                
                # Guardar el histórico centralizado para proyecciones futuras
                with open(ruta_historico_central, 'w') as h_f:
                    json.dump(historico_data, h_f, indent=4)
                    
            else:
                analisis_web["CUD"] = {"saldo_actual": 0.0, "entradas_hoy": 0.0, "salidas_hoy": 0.0, "historial": []}
        else:
            print("ℹ️ No se detectó archivo 'cud_raw.xlsx' en la raíz. Usando valores en cero.")
            analisis_web["CUD"] = {"saldo_actual": 0.0, "entradas_hoy": 0.0, "salidas_hoy": 0.0, "historial": []}

        # --- 4. GUARDADO DE RESULTADOS ---
        with open('financial_data.json', 'w') as f:
            json.dump(analisis_web, f, indent=4)
        
        fecha_hoy = datetime.now().strftime('%Y-%m-%d')
        if not os.path.exists('history'): os.makedirs('history')
        with open(f'history/data_{fecha_hoy}.json', 'w') as f:
            json.dump(analisis_web, f, indent=4)
        
        print("✅ Consolidación terminada exitosamente.")
    except Exception as e:
        print(f"❌ ERROR CRÍTICO GENERAL: {str(e)}")
        exit(1)

if __name__ == "__main__":
    ejecutar_analisis_bi()
