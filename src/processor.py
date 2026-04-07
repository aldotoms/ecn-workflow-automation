import pandas as pd
import re
import os

def clean_ecn_data(input_path, output_path):
    """Limpia los datos del CSV crudo extraído de Outlook."""
    df = pd.read_csv(input_path)
    
    # Regex para ECN y Job (Versión ultra-flexible)
    df['ECN_Number'] = df['Subject'].str.extract(r'ECN\s*#?\s*([\d-]+)', flags=re.IGNORECASE)
    df['Job_Number'] = df['Subject'].str.extract(r'Job\s+([A-Z0-9-/]+)', flags=re.IGNORECASE)
    
    # Cambia la parte de las fechas por esta versión más "amigable" con Excel:
    df['Received_Date'] = pd.to_datetime(df['Received_Date'], errors='coerce').dt.tz_localize(None)
    df['Month'] = df['Received_Date'].dt.strftime('%Y-%m')
    
    # KPI: Días transcurridos (ahora sin problemas de zona horaria)
    now = pd.Timestamp.now()
    df['Days_Open'] = (now - df['Received_Date']).dt.days
        
    # Limpieza de nulos y duplicados en el set nuevo
    df_clean = df.dropna(subset=['ECN_Number']).copy()
    df_clean = df_clean.sort_values('Received_Date', ascending=False).drop_duplicates('ECN_Number')
    
    # Guardamos el CSV procesado (como respaldo técnico)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df_clean.to_csv(output_path, index=False)
    
    return df_clean

def update_excel_tracker(df_new, excel_path):
    """Une los nuevos ECNs con el Excel existente sin borrar notas manuales."""
    if not os.path.exists(excel_path):
        # Si no existe el Excel, creamos las columnas de seguimiento manual
        df_new['Classification'] = "BOM_Rev"
        df_new['Urgency'] = "1_Green"
        df_new['Actions_Required'] = "Regenerate_PKs"
        df_new['Status'] = "Pending"
        df_new.to_excel(excel_path, index=False)
        print(f"-> Archivo Maestro creado: {excel_path}")
    else:
        # Leemos el Excel actual que ya tiene tus notas
        df_existing = pd.read_excel(excel_path)
        
        # Filtramos: Solo tomamos ECNs que NO estén ya en el Excel
        new_records = df_new[~df_new['ECN_Number'].isin(df_existing['ECN_Number'])].copy()
        
        if not new_records.empty:
            new_records['Classification'] = "BOM_Rev"
            new_records['Urgency'] = "1_Green"
            new_records['Actions_Required'] = "Regenerate_PKs"
            new_records['Status'] = "Pending"
            
            # Concatenamos y guardamos
            df_final = pd.concat([df_existing, new_records], ignore_index=True)
            df_final.to_excel(excel_path, index=False)
            print(f"-> ¡Éxito! Se agregaron {len(new_records)} ECNs nuevos al Tracker.")
        else:
            print("-> No se detectaron ECNs nuevos. El Tracker está al día.")

if __name__ == "__main__":
    # Rutas de archivos
    raw_csv = r"O:\11-SFM_Level_2_Planning\ECN_Project\data\raw\ecn_raw_data.csv"
    processed_csv = r"O:\11-SFM_Level_2_Planning\ECN_Project\data\processed\ecn_cleaned_data.csv"
    master_excel = r"O:\11-SFM_Level_2_Planning\ECN_Project\ECN_Tracker_Master.xlsx"
    
    # Ejecución del Pipeline
    print("Iniciando procesamiento de datos...")
    df_cleaned = clean_ecn_data(raw_csv, processed_csv)
    
    print("Actualizando Tracker Maestro de Excel...")
    update_excel_tracker(df_cleaned, master_excel)
    
    print("\nProceso finalizado correctamente.")