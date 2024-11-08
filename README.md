# Visualizador de Registro de Eventos
![image](https://github.com/user-attachments/assets/d3cc15ad-4b57-4d31-be14-731759003ce5)

## Descripción

Este proyecto es una aplicación de escritorio desarrollada en Python que permite visualizar y analizar eventos del registro de Windows. Está diseñado para ayudar a identificar eventos de seguridad, administración del sistema, y otros eventos importantes del sistema operativo, como:

- Creación de usuarios y servicios.
- Conexión de dispositivos USB.
- Fallos en el inicio o la terminación de servicios.
- Cambios en políticas de seguridad.
- Eventos de inicio o cierre de sesión.

La aplicación también ofrece funcionalidades para exportar los eventos filtrados a formatos como TXT, PDF y Excel.

## Características

- Visualización de eventos del registro de Windows en una interfaz gráfica de usuario (GUI).
- Filtros de búsqueda en tiempo real.
- Exportación de los eventos filtrados a formatos TXT, PDF y Excel.
- Compatible con eventos de seguridad y administración de sistema, como la creación de usuarios, fallos en servicios, cambios en políticas de seguridad, etc.
- Soporte para eventos adicionales de análisis forense.

## Requisitos

- Python 3.x
- Bibliotecas Python:
  - `tkinter`: Para la interfaz gráfica de usuario.
  - `win32evtlog`: Para acceder a los registros de eventos de Windows.
  - `fpdf`: Para generar archivos PDF.
  - `pandas`: Para exportar los datos a archivos Excel.

### Instalación

1. **Clona el repositorio**:

   ```bash
   git clone https://github.com/bbaezb/event-log-viewer.git
   cd visualizador-eventos
   
2. **Crea un entorno virtual**:

   ```bash
   python -m venv venv
   source venv/bin/activate

3. **Crea un entorno virtual**:

    ```bash
    pip install -r requirements.txt

4. **Instala las dependencias**:

    ```bash
    python eventos-logs.py
