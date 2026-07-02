# Bridge-Finance Siigo-C3

Bridge-Finance Siigo-C3 es una herramienta de escritorio desarrollada en Python con interfaz gráfica basada en CustomTkinter. Su propósito es facilitar la conciliación entre reportes C3 y facturas/recibos registrados en Siigo, además de permitir la creación masiva de documentos cuando se requiere cargar información automáticamente.

Este proyecto está pensado para operaciones de control, contabilidad y administración escolar o institucional, donde se necesita comparar reportes financieros con transacciones ya registradas y, en casos específicos, generar documentos de forma masiva con los datos del reporte C3.

## Qué hace el proyecto

El flujo general del sistema es el siguiente:

1. El usuario selecciona un archivo C3 en la interfaz.
2. La aplicación carga el reporte y lo normaliza para que el contenido quede listo para ser procesado.
3. Se consulta la información correspondiente en Siigo para el periodo seleccionado.
4. Se ejecuta la conciliación entre ambos conjuntos de datos.
5. El resultado puede visualizarse en la tabla de la interfaz y exportarse a Excel.
6. Si el usuario lo requiere, puede lanzar un proceso de facturación masiva que crea facturas y recibos reales en Siigo.

## Componentes principales del proyecto

- main.py
  - Contiene la interfaz gráfica y el flujo principal de la aplicación.
  - Gestiona la selección de archivos, el periodo, los botones de ejecución y los hilos para no bloquear la UI.
  - Orquesta la conciliación y la facturación masiva.

- api_siigo.py
  - Encapsula la comunicación con la API de Siigo.
  - Maneja autenticación, generación de headers, creación de facturas, creación de recibos y paginación de consultas.
  - Implementa reintentos y control de errores para que la integración sea más robusta.

- reconciliation.py
  - Contiene la lógica de carga y transformación del reporte C3.
  - Normaliza columnas, prepara los datos y realiza la conciliación con los datos de Siigo.
  - No se modifica en este README, pero es el motor principal de comparación de información.

- config.json
  - Archivo de configuración runtime con credenciales y parámetros utilizados por la aplicación.
  - No debe compartirse públicamente si contiene datos sensibles.

- config.example.json
  - Plantilla base para crear un archivo de configuración real sin exponer secretos.

## Requisitos previos

Antes de ejecutar la aplicación asegúrate de tener instalado:

- Python 3.9 o superior
- Pip actualizado
- Acceso válido a la API de Siigo con credenciales funcionales
- Un archivo C3 en formato Excel o CSV compatible con el procesador del proyecto

## Dependencias

El proyecto usa las siguientes librerías, listadas en requirements.txt:

- pandas
- requests
- customtkinter
- openpyxl

Puedes instalarlas con:

```bash
pip install -r requirements.txt
```

## Configuración del proyecto

El archivo de configuración real debe llamarse config.json y debe basarse en la plantilla config.example.json.

### Opción 1: Linux / macOS

```bash
cp config.example.json config.json
```

### Opción 2: Windows (cmd)

```cmd
copy config.example.json config.json
```

### Qué debe contener config.json

El archivo de configuración debe incluir, al menos, los valores de:

- username
- access_key
- base_url
- document_id
- seller_id
- cost_center
- default_product_code
- default_payment_id
- tolerance_pesos

Aunque el archivo de ejemplo ya trae valores de referencia, debes reemplazar las credenciales con tus datos reales antes de usar la aplicación contra Siigo.

> Importante: si no completas las credenciales correctas, la autenticación contra Siigo fallará y no podrás crear ni consultar documentos desde la app.

## Cómo ejecutar la aplicación

Desde la raíz del proyecto:

```bash
python main.py
```

En algunos entornos Windows puede funcionar también con:

```cmd
python .\main.py
```

## Flujo de uso recomendado

### 1. Cargar el archivo C3

- Haz clic en "Subir Archivo".
- Selecciona el reporte C3 en formato Excel o CSV.
- El archivo se procesa automáticamente al iniciar la conciliación.

### 2. Definir el periodo de consulta

- Elige el año y el mes para consultar datos en Siigo.
- La aplicación usará ese periodo como referencia para obtener facturas y compararlas con las entradas del reporte C3.

### 3. Ejecutar conciliación

- Pulsa "Conciliar".
- El sistema carga los datos del reporte y los datos de Siigo, los compara y genera una tabla de resultados.
- En la tabla podrás ver coincidencias, diferencias y posibles anomalías.

### 4. Exportar resultados

- Si la conciliación se completó correctamente, puedes exportar el resultado a Excel desde el botón correspondiente.

### 5. Facturación masiva (solo cuando aplica)

- El botón "Subir información y Conciliar" permite iniciar un flujo que crea facturas y recibos reales en Siigo.
- Antes de hacerlo, la aplicación muestra una confirmación explícita porque la acción puede generar documentos reales y no es reversible de forma trivial.
- Este flujo está pensado para cuando se necesita cargar un lote grande de información de forma automática.

## Datos y formatos esperados

Aunque el proyecto está preparado para manejar reportes C3 de forma robusta, es importante que el contenido del archivo cumpla con las columnas esperadas por la lógica de carga y conciliación.

El flujo principal asume que el reporte trae campos como:

- Cod Perm
- Nombre del Estudiante
- Vlr Concepto
- Fecha Emisión
- Nombre Concepto
- Nombre Adquiriente
- No Factura
- Nit/Cédula

La aplicación intenta normalizar esos campos para que puedan ser comparados con los documentos de Siigo.

## Integración con Siigo

La integración se realiza a través de api_siigo.py. Entre las responsabilidades de este módulo se encuentran:

- autenticarse con Siigo
- consultar facturas por periodo
- crear facturas de venta
- crear recibos de caja
- manejar paginación de resultados
- registrar errores en el log de la aplicación

Por defecto, las llamadas HTTP tienen timeout configurable y el cliente está preparado para manejar fallas de red, respuestas inesperadas y problemas de autenticación.

## Logs y manejo de errores

El proyecto registra errores en un archivo llamado log_errores.txt.

Ese archivo es útil para revisar:

- fallas de autenticación
- errores de red
- respuestas inesperadas de la API
- problemas de formato en archivos cargados
- errores durante facturación o conciliación

Si algo falla durante la ejecución, revisa primero el contenido del log antes de repetir la operación.

## Recomendaciones operativas

- Usa siempre un archivo limpio y bien formado antes de iniciar la conciliación.
- Verifica que las credenciales de Siigo estén correctas antes de lanzar procesos masivos.
- Si vas a facturar muchos registros, haz pruebas primero con un conjunto pequeño.
- Revisa los resultados en la interfaz antes de aceptar que los datos sean correctos.
- Mantén el archivo config.json fuera de repositorios públicos o compartidos.

## Estructura de carpetas relevante

En la raíz del proyecto encontrarás normalmente:

- main.py
- api_siigo.py
- reconciliation.py
- config.json (no versionado, ver .gitignore)
- config.example.json
- requirements.txt
- log_errores.txt (no versionado, ver .gitignore)

## Pruebas

Actualmente el proyecto no cuenta con pruebas automatizadas. Se recomienda
validar manualmente con un archivo C3 de prueba y un periodo con pocos
registros antes de ejecutar facturación masiva contra datos reales.

## Notas importantes

- No se debe modificar la lógica central de conciliación sin validar primero el impacto en los resultados.
- La facturación masiva genera documentos reales en Siigo; debe usarse con cuidado.
- El archivo config.json es sensible y debe manejarse de forma segura.
- Si cambias parámetros de negocio como tolerancias, fechas o campos esperados, revisa el comportamiento completo antes de usarlo en producción.

## Resumen rápido

En una frase: este proyecto sirve para cargar reportes C3, compararlos contra Siigo, revisar diferencias y, cuando sea necesario, crear documentos de forma automática con un flujo guiado y controlado.
