# PythonAutomation

## Descripción general
PythonAutomation es una aplicación que toma datos de un archivo de Excel y convierte cada fila en un PDF con un formato específico para presentar un proceso tributario en España.

## Requisitos previos
- **Python**: Se recomienda usar la versión más reciente de Python disponible a la fecha.
- **Librerías**: Las librerías necesarias están listadas en el archivo `requirements.txt`.

## Instrucciones de instalación
1. Clona este repositorio o descarga los archivos.
2. Asegúrate de tener Python instalado. Puedes descargarlo desde [python.org](https://www.python.org/).
3. Navega a la carpeta del proyecto donde se encuentra `app.py` y `requirements.txt`.
4. Instala las dependencias usando el siguiente comando:
```bash
pip install -r requirements.txt
```

## Configuración inicial
Antes de ejecutar la aplicación, configura las rutas de las carpetas a usar en `app.py` y en los archivos de los bots (`bot247.py`, `bot_i129s.py` y `read_excels.py`) según tu entorno local.

## Instrucciones de ejecución
Para iniciar la aplicación Flask, ejecuta el siguiente comando:
```bash
python app.py
```

## Rutas y funcionalidades principales
### Rutas disponibles
- **/**: Muestra la página principal con opciones para iniciar los bots.
- **/start_bot247**: Inicia el bot que procesa el archivo Excel para generar PDFs según el modelo 247.
- **/start_i129s**: Inicia el bot que procesa el archivo Excel para generar PDFs según el modelo i129s.
- **/start_excel_i129s**: Inicia el proceso de generación del archivo Excel necesario para el bot i129s.

## Estructura del proyecto
```
PYTHONAUTOMATION
│
├── Bot_247
│   └── bot247.py
│
├── Bot_i129s
│   ├── bot_i129s.py
│   └── read_excels.py
│
├── templates
│   └── index.html
│
├── app.bat
├── app.py
└── requirements.txt
```

### Archivos importantes
- **app.py**: Archivo principal que contiene la aplicación Flask.
- **requirements.txt**: Lista de dependencias necesarias para ejecutar la aplicación.
- **Bot_247/bot247.py**: Contiene la lógica para generar PDFs del modelo 247.
- **Bot_i129s/bot_i129s.py**: Contiene la lógica para generar PDFs del modelo i129s.
- **Bot_i129s/read_excels.py**: Contiene la lógica para generar el archivo Excel necesario para el bot i129s.
- **templates/index.html**: Plantilla HTML para la interfaz de usuario.
