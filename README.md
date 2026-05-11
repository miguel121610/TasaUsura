# Automatizador de Formularios PDF

Esta aplicación en Python automatiza el llenado de formularios para múltiples usuarios basándose en un origen de datos (exportado como CSV/Excel o desde base de datos) y una plantilla en Excel, generando un PDF individual para cada registro.

## Requisitos Previos

- Windows OS (Requerido para la conversión directa a PDF mediante `win32com`).
- Microsoft Excel instalado en el sistema (Es utilizado en segundo plano para generar los PDF nativamente).
- Python 3.8 o superior.

## Instalación

1. Clona o descarga este proyecto.
2. Abre una terminal en la carpeta del proyecto (`FormAutoFill`).
3. Se recomienda crear un entorno virtual:
   ```cmd
   python -m venv venv
   venv\Scripts\activate
   ```
4. Instala las dependencias necesarias:
   ```cmd
   pip install -r requirements.txt
   ```

## Estructura del Proyecto

```text
FormAutoFill/
├── main.py                     # Punto de entrada de la aplicación
├── requirements.txt            # Dependencias del proyecto
├── README.md                   # Instrucciones
├── src/                        # Código fuente
│   ├── gui.py                  # Interfaz gráfica de usuario (Frontend)
│   ├── processor.py            # Lógica de procesamiento de datos y PDF
│   └── logger.py               # Configuración de logs
└── data_examples/              # Ejemplos de uso
    └── setup_db_example.py     # Script para generar base de datos SQLite y CSV de prueba
```

## Uso de la Aplicación

1. **Preparar la Plantilla (Excel):**
   Crea un archivo Excel que será tu plantilla de formulario. En las celdas donde desees que se reemplace la información del usuario, utiliza llaves `{}` o corchetes `[]` con el nombre exacto de la columna del origen de datos.
   _Ejemplo: Si tu base de datos tiene una columna `nombre`, pon en el Excel `{nombre}` o `[nombre]`._
2. **Preparar los Datos:**
   Si tienes una base de datos SQL, exporta el resultado de tu consulta a formato CSV o Excel (`.xlsx`). La primera fila debe contener los nombres de las columnas.

   _Opcional:_ Puedes ejecutar `python data_examples/setup_db_example.py` para generar una base de datos SQLite de ejemplo y un archivo CSV listo para probar.

3. **Ejecutar la App:**

   ```cmd
   python main.py
   ```

4. **Interfaz de Usuario:**
   - **Paso 1:** Carga tu archivo CSV o Excel con los datos de los usuarios.
   - **Paso 2:** Selecciona tu plantilla Excel.
   - **Paso 3:** Selecciona en qué carpeta quieres que se guarden los PDFs.
   - Haz clic en **"Iniciar Proceso"** y observa la barra de progreso y los logs en pantalla.

## Características Principales

- **Mapeo Automático:** La aplicación cruza dinámicamente las columnas del CSV/Base de datos con las variables indicadas en el Excel (`{variable}`).
- **Nombres Dinámicos:** Los PDFs generados adoptan nombres dinámicos combinando `nombre_apellido_identificacion.pdf` para fácil identificación.
- **Validación de Datos:** Maneja correctamente campos nulos o vacíos sin interrumpir la ejecución.
- **Progreso y Logs:** Barra de progreso en tiempo real y logs integrados (tanto en UI como guardados en archivo local `/logs`).
- **Separación de Lógica y Frontend:** Interfaz Tkinter desacoplada de la lógica de procesamiento (pandas + openpyxl + win32com).
