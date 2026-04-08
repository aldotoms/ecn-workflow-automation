import pandas as pd
import os
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils.dataframe import dataframe_to_rows

def update_excel_tracker(df_new, excel_path):
    """Actualiza o crea el Excel asegurando la existencia de la Tabla oficial."""
    sheet_name = "ECNs"
    table_name = "Tbl_ECN_Master"

    # Preparar los datos nuevos (si el archivo no existe, estos serán los primeros)
    if not os.path.exists(excel_path):
        df_to_add = df_new.copy()
        df_to_add['Classification'] = "BOM_Rev"
        df_to_add['Urgency'] = "1_Green"
        df_to_add['Actions_Required'] = "Regenerate_PKs"
        df_to_add['Status'] = "Pending"
        is_new_file = True
    else:
        df_existing = pd.read_excel(excel_path, sheet_name=sheet_name)
        df_to_add = df_new[~df_new['ECN_Number'].isin(df_existing['ECN_Number'])].copy()
        if not df_to_add.empty:
            df_to_add['Classification'] = "BOM_Rev"
            df_to_add['Urgency'] = "1_Green"
            df_to_add['Actions_Required'] = "Regenerate_PKs"
            df_to_add['Status'] = "Pending"
        is_new_file = False

    if df_to_add.empty and not is_new_file:
        print("-> No hay ECNs nuevos. El archivo se mantuvo intacto.")
        return

    # --- MANEJO CON OPENPYXL ---
    if is_new_file:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        # Escribir encabezados y datos
        for r in dataframe_to_rows(df_to_add, index=False, header=True):
            ws.append(r)
    else:
        wb = load_workbook(excel_path)
        ws = wb[sheet_name]
        # Escribir solo los datos nuevos (sin encabezados)
        for r in dataframe_to_rows(df_to_add, index=False, header=False):
            ws.append(r)

    # ACTUALIZAR O CREAR LA TABLA (El corazón de la automatización)
    last_col_letter = ws.cell(row=1, column=ws.max_column).column_letter
    new_range = f"A1:{last_col_letter}{ws.max_row}"

    if table_name in ws.tables:
        # Si la tabla ya existe, solo actualizamos su rango
        ws.tables[table_name].ref = new_range
    else:
        # Si es archivo nuevo, creamos la tabla desde cero
        tab = Table(displayName=table_name, ref=new_range)
        # Agregamos un estilo visual por defecto (puedes cambiarlo)
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        ws.add_table(tab)

    wb.save(excel_path)
    print(f"-> Proceso completado: Archivo {'creado' if is_new_file else 'actualizado'} y tabla '{table_name}' lista.")