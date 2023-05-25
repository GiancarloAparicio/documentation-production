
Generación de los documentos de producción RPA para Sura.
* ✅ Queries para subir el Config.xlsx a la base de datos
* ✅ `Production document`
* ✅ `PR`: Generacion de scripts
    * Te genera los scripts base, pero dependiendo del proceso se deberán agregar mas scripts
* ✅ `REV`: Revision de pares
* ✅ `PO`: Plan de orden de trabajo.
    * Generación parcial, actualizar la pestaña "General" con las fechas de tu estimación
* ✅ `LR`: Lista de requerimiento (funcionales y no funcionales)
    * Pendiente
* ✅ `CP`: Casos de prueba y su detalle
    * Pendiente

# Dependencias
Instalar Python 3.9.2: [Descarga](https://www.python.org/ftp/python/3.9.2/python-3.9.2-amd64.exe)

![Python](./images/python.png)

Instalar dependencias, ejecutar en Powershell.
```sh
pip install -r requirements.txt
```

# Guia de uso
Para que el script funcione correctamente se debe seguir el [estándar especificado](https://gitlab.com/grupo_rpa/rpa_002_manualprocedimientodesarrollorpa#estructura-codigo-fuente-del-proyecto).

Clonar o descargar el repositorio.

![Gitlab](./images/repo-gitlab.png)

Copiar los siguientes documentos y archivos a la [raiz de su proyecto o repositorio](https://gitlab.com/grupo_rpa/rpa_002_manualprocedimientodesarrollorpa#creaci%C3%B3n-estructura-de-carpetas-seg%C3%BAn-est%C3%A1ndar-rpa-area-de-trabajo-y-fuente-proyecto)

![base](./images/base-script.png)

Configurar el script usando el archivo `Generate\Generate.xlsx`

![config](./images/config-generate.png)

Ejecutar la funcionalidad deseada.
* 1. Genera las queries para subir el **Config.xlsx** a la base de datos
* 2. Genera el documento word **"Production Document.docx"**
* 3. Generar documentos PR, REV y PO
* 99. Ejecuta todas las funcionalidades anteriores
* 0. Salir del menu

![run](./images/run-script.png)
