import pandas as pd
import win32com.client
import os

def connect_to_outlook():
    # Conectamos con la aplicación Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    
    # 1. Accedemos específicamente al "Online Archive" que se ve en tu imagen
    # Debes usar el nombre exacto que aparece arriba de 'Inbox' en Outlook
    archive_name = "Online Archive - AOrduna@flowserve.com"
    
    try:
        # Buscamos el almacén de datos del archivo en línea
        store = outlook.Folders.Item(archive_name)
        
        # 2. Ahora buscamos la carpeta 'ECN_Releases' que está en la raíz de ese archivo
        ecn_folder = store.Folders.Item("ECN_Releases")
        
        print(f"Conexión exitosa a la carpeta: {ecn_folder.Name}")
        print(f"Mensajes sin leer en esta carpeta: {ecn_folder.UnReadItemCount}")
        return ecn_folder
        
    except Exception as e:
        print(f"Error: No se pudo encontrar la carpeta en el Online Archive. {e}")
        return None

def list_emails(folder):
    messages = folder.Items
    print(f"Total de correos encontrados: {len(messages)}")
    
    for message in messages:
        # Imprimimos el asunto para probar
        print(f"Subject: {message.Subject} | Recibido: {message.ReceivedTime}")

def get_data_from_emails(folder):
    messages = folder.Items
    data_list = []
    
    print(f"Extrayendo datos de {len(messages)} correos...")
    
    for message in messages:
        # Creamos un diccionario con la info de cada correo
        email_info = {
            "Subject": message.Subject,
            "Sender": message.SenderName,
            "Received_Date": str(message.ReceivedTime), # Convertimos a string para facilitar manejo
            "Body_Preview": message.Body[:200].strip(), # Tomamos los primeros 200 caracteres
            "Is_Read": not message.UnRead
        }
        data_list.append(email_info)
    
    # Convertimos la lista a un DataFrame de Pandas
    df = pd.DataFrame(data_list)
    return df

if __name__ == "__main__":
    folder = connect_to_outlook()
    if folder:
        df_ecn = get_data_from_emails(folder)
        
         # Definimos la ruta (usando la 'r' o '/')
        output_path = r"O:\11-SFM_Level_2_Planning\ECN_Project\data\raw\ecn_raw_data.csv"

        # ESTA LÍNEA ES CLAVE: Crea las carpetas si no existen
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        # Guardamos una copia en tu carpeta data/raw para el portafolio
        # output_path = r"O:\11.- SFM Level 2 Planning\ECN_Project\data\raw\ecn_raw_data.csv"
        
        # Ahora sí, guardamos el archivo
        df_ecn.to_csv(output_path, index=False)
                
        print(f"¡Listo! Dataset creado con {len(df_ecn)} filas en: {output_path}")
        print(df_ecn.head()) # Muestra las primeras 5 filas en la terminal
        print(f"¡Listo! Dataset guardado exitosamente.")