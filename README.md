# 📧 ECN Workflow Automation: Python, Outlook & Power BI
Este proyecto automatiza la extracción, procesamiento y visualización de correos de Notificación de Cambios de Ingeniería (ECN) desde Microsoft Outlook. Transforma un proceso manual y propenso a errores en un flujo de datos estructurado y visual.


## 🚀 Impacto del Proyecto
En entornos de manufactura y planeación de la demanda, la gestión de ECNs es crítica. Esta herramienta:
- Elimina la extracción manual de datos desde el cuerpo de los correos electrónicos.
- Centraliza la información en una base de datos estructurada.
- Facilita la toma de decisiones mediante un dashboard interactivo que muestra el estatus de los cambios en tiempo real.

## 🛠️ Stack Tecnológico
Lenguaje: Python 3.x | Automatización: pywin32 (para la integración nativa con Outlook). | Procesamiento de Datos: Pandas (limpieza y transformación de datos), Feature Enginnering (creación de métricos)
Visualización: Power BI (Dashboard dinámico). | Almacenamiento: CSV / Excel (como fuente de datos para el reporte).


## 📁 Estructura del Repositorio
- data/: Archivos procesados (ejemplos anonimizados).
- src/: Contiene el script de Python para la conexión y extracción de Outlook.
- dashboard/: Archivo .pbix de Power BI o capturas de pantalla del reporte.
- requirements.txt: Librerías necesarias para ejecutar el entorno.


## ⚙️ Funcionamiento
Extracción: El script accede a la carpeta específica de ECN en Outlook.
Parsing: Utiliza expresiones regulares o lógica de Python para extraer campos clave (Número de ECN, Fecha, Responsable, Estatus).
Carga: Los datos se exportan de forma incremental a un archivo maestro.
Visualización: Power BI se conecta al archivo maestro para actualizar los KPIs de cumplimiento y carga de trabajo.


## 📊 Visualización (Dashboard)
(Aquí podrías insertar una imagen de tu dashboard de Power BI)
Nota: El dashboard permite filtrar por departamento, prioridad y tiempo de respuesta desde que se recibe el correo hasta que se procesa el cambio.


##📄 Licencia
Este proyecto está bajo la Licencia MIT. Siéntete libre de usarlo o adaptarlo.
