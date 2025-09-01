# Automatización de Operaciones del Agregador


![Python](https://img.shields.io/badge/Python-3.10-blue?logo=python)
![Power BI](https://img.shields.io/badge/Power%20BI-Dashboard-yellow?logo=powerbi)
![Excel](https://img.shields.io/badge/Excel-Macros-green?logo=microsoft-excel)


Este proyecto tiene como objetivo **automatizar la gestión de operaciones de un agregador** (terminales de punto de venta, usuarios y transacciones).  
La solución combina **Excel, Python y Power BI** para:

- Centralizar información.  
- Mejorar la trazabilidad.  
- Calcular comisiones.  
- Analizar transacciones.  
- Optimizar la toma de decisiones en las áreas financiera, comercial y operativa.  


## Estructura del repositorio

Automatizacion-de-Operaciones-del-Agregador/
│
├── data/ # Datos de entrada y salida
│ ├── BD_TRANSACCIONES_PRUEBA.xlsx
│ ├── Transacciones_limpias.xlsx
│ ├── Usuarios.xlsx
│ ├── Detalles.xlsx
│ └── Macros.bas # Macros para actualizar automáticamente la tabla de hechos
│
├── notebook/ # Código en Jupyter/Colab
│ ├── Limpiar_transacciones.ipynb
│ └── limpiar_transacciones.py
│
├── dashboards/ # Dashboards en Excel/Power BI
│ ├── Operaciones TPV v3.xlsm
│ └── BD_OPERACIONES_V4.pbix
│
├── videos/ # Demostraciones en video
│ ├── POWER BI.zip
│ └── POWER BI2.zip
│
└── README.md # Documentación principal


## Instalación y Requisitos

### Programas necesarios
- Excel Desktop (para macros de inventario y gestión de terminales).  
- Python 3.x (Google Colab o ejecución local).  
- Power BI Desktop.  

### Librerías de Python
```bash
pip install pandas
pip install openpyxl


Uso
1. Excel:

- Buscar una terminal en la hoja Menú ingresando el ID de TPV.

- Actualizar ventas y pagos desde la macro en la hoja Detalles.

- Generar consolidado de operaciones con el botón Actualizar.

2. Python (Google Colab o local):

- Subir archivo BD_TRANSACCIONES_PRUEBA.xlsx.

- Ejecutar Limpiar_transacciones.ipynb.

- Generar un archivo limpio Transacciones_limpias.xlsx para usar en Power BI.

3. Power BI:

- Importar Detalles.xlsx y el archivo limpio generado por Python.

- Dar clic en Actualizar para refrescar datos.

Consultar dashboards:

- Transacciones por cliente/operativa.

- Ventas de TPVs por mes/vendedor.

- Crecimiento mensual.

- Control de comisiones.

Dashboards:

Los dashboards están en la carpeta dashboards/.
Ábrelos con Power BI Desktop para ver las visualizaciones interactivas.

Monitoreo de terminales:

<img width="1319" height="707" alt="image" src="https://github.com/user-attachments/assets/bdf025cf-3b28-4374-b16f-dfa93e29456e" />

Operaciones por usuario y ubicación:

<img width="1282" height="725" alt="image" src="https://github.com/user-attachments/assets/14355723-edd4-46bb-b46d-f630291a7636" />

Informe de Ventas por Operativa:

<img width="1231" height="653" alt="image" src="https://github.com/user-attachments/assets/09ab10b7-7c1d-4c02-8426-fd9c0d3f9511" /> <img width="1231" height="653" alt="image" src="https://github.com/user-attachments/assets/78ee55ad-cbf2-4847-8f67-ee5d978b72d3" />

Ranking de usuarios operando:
<img width="1224" height="677" alt="image" src="https://github.com/user-attachments/assets/068ad4a4-cd8c-45cc-b5fc-b9e8e592f3a8" />
Videos

En la carpeta videos/ encontrarás demostraciones en formato .zip con las interacciones en Power BI.

Resultados e Impacto:

- Control total de inventario de terminales.

- Seguimiento de pagos y comisiones de vendedores.

- Análisis de clientes activos/inactivos y estrategias de reactivación.

- Reducción del tiempo en la manipulación de datos masivos.

- Automatización con una sola persona a cargo.

Tecnologías utilizadas:

- Excel → Macros para inventario y gestión de TPVs.

- Python (Pandas, os, zipfile) → Limpieza y transformación de datos.

- Power BI → Dashboards interactivos y análisis financiero-operativo.

Próximos pasos:

  - Integrar base de datos SQL para almacenamiento centralizado.

  - Automatizar carga de datos en Power BI.

  - Generar alertas automáticas sobre clientes inactivos o inventario bajo.

Autor:

- Nombre: Mirna Alaniz

- Contacto: mirna.yt@gmail.com

- LinkedIn: Mirna Alaniz | www.linkedin.com/in/mirna-alaniz-0b979214b
