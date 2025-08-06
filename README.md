# Conversor XML a Excel - Plugin IneoXlsx

Plugin para convertir archivos XML estructurados a archivos Excel (.xlsx) con soporte para estilos, múltiples hojas y preservación de datos existentes.

## Características

- Conversión de XML a Excel con estilos personalizados
- Soporte para múltiples hojas de cálculo
- Preservación de archivos Excel existentes
- Reutilización inteligente de hojas con el mismo nombre
- Aplicación de fuentes, colores, fondos y alineación
- Ajuste automático del ancho de columnas
- Interfaz de línea de comandos simple

## Uso

### Sintaxis básica
```bash
python ineoXlsxCmdLine.exe <archivo_xml> [archivo_excel]
```

### Ejemplos de uso

#### Uso básico (estructura simple)
```bash
# Convertir XML a Excel con configuración básica
python ineoXlsxCmdLine.py datos.xml

# Especificar archivo de salida personalizado
python ineoXlsxCmdLine.py datos.xml resultado.xlsx
```

#### Uso avanzado (estructura completa con logging)
```bash
# Procesar configuración completa con logging detallado
python ineoXlsxCmdLine.py configuracion_completa.xml
```

**Ejemplo de archivo de configuración completa:**
```xml
<?xml version="1.0" encoding="UTF-8"?>
<ineoDoc task="updateXlsx" task_id="conversion_001">
    <data>
        <dataIn>FILE://datos/empleados.xml</dataIn>
        <dataOut>FILE://output/empleados_procesados.xlsx</dataOut>
    </data>
    <log>
        <logLevel>DEBUG</logLevel>
        <logFile>FILE://logs/conversion.log</logFile>
        <logFormat>%(asctime)s - %(levelname)s - %(message)s</logFormat>
        <logConsole>true</logConsole>
    </log>
    <workbooks>
        <!-- Datos del Excel aquí -->
    </workbooks>
</ineoDoc>
```

**Ejemplo sin dataIn (usando configuración como datos):**
```xml
<?xml version="1.0" encoding="UTF-8"?>
<ineoDoc task="updateXlsx" task_id="simple_001">
    <data>
        <dataOut>FILE://tmp/resultado.xlsx</dataOut>
    </data>
    <log>
        <logLevel>INFO</logLevel>
        <logConsole>true</logConsole>
    </log>
    <workbooks>
        <styles>
            <style id="header">
                <font>Arial</font>
                <bold>true</bold>
                <color>#FFFFFF</color>
                <background>#4472C4</background>
            </style>
        </styles>
        <workbook name="Datos">
            <cell row="1" column="A" text="Título" style="header"/>
            <cell row="2" column="A" text="Contenido"/>
        </workbook>
    </workbooks>
</ineoDoc>
```

## Comportamiento con archivos existentes

- **Archivo Excel no existe**: Se crea un nuevo archivo con los datos del XML
- **Archivo Excel existe**: Se carga el archivo existente y se preservan todos los datos
  - **Hoja existe**: Se reutiliza la hoja existente, agregando/actualizando celdas
  - **Hoja no existe**: Se crea una nueva hoja en el archivo existente

## Estructura XML requerida

### Estructura completa (recomendada)
```xml
<?xml version="1.0" encoding="UTF-8"?>
<ineoDoc task="updateXlsx" task_id="identificador_unico">
    <data>
        <dataIn>FILE://ruta/archivo_datos.xml</dataIn>
        <dataOut>FILE://ruta/archivo_salida.xlsx</dataOut>
    </data>
    <responseOut>FILE://./respuesta.log</responseOut>
    <log>
        <logLevel>DEBUG</logLevel>
        <logFile>FILE://./conversion.log</logFile>
        <logFormat>%(asctime)s - %(name)s - %(levelname)s - %(message)s</logFormat>
        <logDateFormat>%Y-%m-%d %H:%M:%S</logDateFormat>
        <logConsole>true</logConsole>
    </log>
    <options>
        <option name="configuracion" value="valor"/>
    </options>
    <workbooks>
        <styles>
            <!-- Definición de estilos -->
        </styles>
        <workbook name="NombreHoja">
            <!-- Celdas del contenido -->
        </workbook>
    </workbooks>
</ineoDoc>
```

### Estructura básica (compatibilidad)
```xml
<?xml version="1.0" encoding="UTF-8"?>
<workbooks>
    <styles>
        <!-- Definición de estilos -->
    </styles>
    
    <workbook name="NombreHoja">
        <!-- Celdas del contenido -->
    </workbook>
</workbooks>
```

### Configuración de datos (dataIn y dataOut)

#### Tipos de fuentes soportadas

**dataIn** (archivo de entrada):
- `FILE://ruta/archivo.xml` - Archivo local
- `BASE64://contenido_codificado` - Contenido XML codificado en BASE64
- `ruta/archivo.xml` - Archivo local (por defecto)

**dataOut** (archivo de salida):
- `FILE://ruta/archivo.xlsx` - Archivo local
- `URL://https://servidor.com/api/upload` - URL para envío
- `ruta/archivo.xlsx` - Archivo local (por defecto)

#### Configuración opcional
- **dataIn opcional**: Si no se especifica, se usará el propio archivo de configuración como datos
- **Validación automática**: 
  - Para **dataIn**: Se verifica la existencia de archivos y decodificación de BASE64
  - Para **dataOut**: Se crean automáticamente los directorios padre si no existen
  - Para **URLs**: Se valida la accesibilidad del endpoint

#### Gestión automática de directorios
El sistema crea automáticamente los directorios necesarios para los archivos de salida:
- **Detección**: Verifica si el directorio padre del archivo de salida existe
- **Creación**: Crea automáticamente la estructura de directorios necesaria
- **Logging**: Registra la creación de directorios en el log
- **Manejo de errores**: Si no puede crear el directorio, registra el error y detiene el procesamiento

### Configuración de logging

La sección `<log>` permite configurar el sistema de registros:

```xml
<log>
    <logLevel>DEBUG|INFO|WARNING|ERROR</logLevel>
    <logFile>FILE://ruta/archivo.log</logFile>
    <logFormat>%(asctime)s - %(name)s - %(levelname)s - %(message)s</logFormat>
    <logDateFormat>%Y-%m-%d %H:%M:%S</logDateFormat>
    <logConsole>true|false</logConsole>
</log>
```

**Propiedades de logging**:
- `logLevel`: Nivel de detalle de los logs (DEBUG, INFO, WARNING, ERROR)
- `logFile`: Archivo donde guardar los logs (soporta prefijo FILE://)
- `logFormat`: Formato de los mensajes de log
- `logDateFormat`: Formato de fecha y hora
- `logConsole`: Mostrar logs en consola (true/false)

**Comportamiento del logging**:
- **Configuración automática**: Al inicio se registra la configuración aplicada
- **Validación de archivos**: Se registra la existencia/creación de archivos y directorios
- **Creación de directorios**: Se documenta cuando se crean directorios automáticamente
- **Procesamiento de datos**: Se registra el origen de los datos (dataIn o archivo de configuración)
- **Manejo de errores**: Todos los errores se registran con detalles específicos

**Ejemplos de logs generados**:
```
INFO - Iniciando conversión XML a Excel
INFO - Configuración de logging aplicada: {'logLevel': 'DEBUG', 'logFile': 'FILE://./conversion.log', 'logConsole': 'true'}
INFO - dataIn especificado: FILE://datos.xml
INFO - Archivo encontrado: datos.xml
INFO - Directorio creado: tmp/output
INFO - Archivo de salida configurado: tmp/output/resultado.xlsx
INFO - Archivo Excel creado exitosamente: tmp/output/resultado.xlsx
```

### Definición de estilos
Los estilos se definen una sola vez y se reutilizan mediante el atributo `style`:

```xml
<styles>
    <style id="1">
        <font>Arial</font>
        <size>12</size>
        <bold>true</bold>
        <color>#000000</color>
        <background>#E6E6FA</background>
        <alignment>center</alignment>
    </style>
    <style id="2">
        <font>Calibri</font>
        <size>10</size>
        <bold>false</bold>
        <color>#333333</color>
        <alignment>left</alignment>
    </style>
</styles>
```

#### Propiedades de estilo soportadas
- `font`: Nombre de la fuente (ej: Arial, Calibri, Times New Roman)
- `size`: Tamaño de fuente en puntos
- `bold`: `true` o `false` para texto en negrita
- `color`: Color del texto en formato hexadecimal (ej: #000000, #FF0000)
- `background`: Color de fondo de la celda en formato hexadecimal
- `alignment`: Alineación horizontal (`left`, `center`, `right`)

### Definición de hojas y celdas
```xml
<workbook name="Empleados">
    <cell row="1" column="A" text="Nombre Completo" style="1"/>
    <cell row="1" column="B" text="Edad" style="1"/>
    <cell row="1" column="C" text="Departamento" style="1"/>
    
    <cell row="2" column="A" text="Juan Pérez" style="2"/>
    <cell row="2" column="B" text="30" style="2"/>
    <cell row="2" column="C" text="Ventas" style="2"/>
</workbook>
```

#### Atributos de celda
- `row`: Número de fila (empezando desde 1)
- `column`: Letra de columna (A, B, C, ...)
- `text`: Contenido de la celda
- `style`: ID del estilo a aplicar (opcional)

### Ejemplo completo
```xml
<?xml version="1.0" encoding="UTF-8"?>
<workbooks>
    <styles>
        <style id="header">
            <font>Arial</font>
            <size>12</size>
            <bold>true</bold>
            <color>#FFFFFF</color>
            <background>#4472C4</background>
            <alignment>center</alignment>
        </style>
        <style id="data">
            <font>Arial</font>
            <size>10</size>
            <bold>false</bold>
            <color>#000000</color>
            <alignment>left</alignment>
        </style>
    </styles>
    
    <workbook name="Empleados">
        <!-- Encabezados -->
        <cell row="1" column="A" text="Nombre" style="header"/>
        <cell row="1" column="B" text="Edad" style="header"/>
        <cell row="1" column="C" text="Departamento" style="header"/>
        
        <!-- Datos -->
        <cell row="2" column="A" text="Ana García" style="data"/>
        <cell row="2" column="B" text="28" style="data"/>
        <cell row="2" column="C" text="Marketing" style="data"/>
        
        <cell row="3" column="A" text="Carlos López" style="data"/>
        <cell row="3" column="B" text="35" style="data"/>
        <cell row="3" column="C" text="IT" style="data"/>
    </workbook>
    
    <workbook name="Ventas">
        <cell row="1" column="A" text="Mes" style="header"/>
        <cell row="1" column="B" text="Importe" style="header"/>
        
        <cell row="2" column="A" text="Enero" style="data"/>
        <cell row="2" column="B" text="15000" style="data"/>
    </workbook>
</workbooks>
```

## Características avanzadas

### Validación automática de esquema XSD
El plugin valida automáticamente la estructura del XML contra un esquema XSD antes del procesamiento, proporcionando mensajes de error específicos si el formato no es correcto.

### Múltiples hojas con el mismo nombre
Si el XML contiene varios elementos `<workbook>` con el mismo `name`, todas las celdas se escribirán en la misma hoja Excel, permitiendo agregar contenido de forma incremental.

### Reutilización de hojas existentes
Al procesar un archivo Excel existente, el plugin detecta automáticamente las hojas presentes y las reutiliza si coinciden con los nombres en el XML.

### Procesamiento de contenido BASE64
Soporte para contenido XML embebido como BASE64 en el campo `dataIn`, útil para integraciones con APIs o sistemas que envían datos codificados.

### Envío a URLs
Capacidad de enviar el archivo Excel resultante directamente a una URL mediante `dataOut` con prefijo `URL://`.

### Sistema de logging configurable
Logging completo y configurable que permite rastrear todo el proceso de conversión, errores y advertencias según la configuración especificada. El sistema registra automáticamente su propia configuración y todas las operaciones realizadas.

### Ajuste automático de columnas
El plugin ajusta automáticamente el ancho de las columnas basándose en el contenido, con un ancho máximo de 50 caracteres.

## Mensajes informativos y logging

El plugin proporciona información detallada durante la ejecución a través de dos mecanismos:

### Salida estándar
Mensajes básicos mostrados en consola:
- Validación de esquema XSD
- Carga de archivos existentes vs creación de nuevos
- Reutilización vs creación de hojas
- Tiempo de procesamiento
- Estado de la conversión

### Sistema de logging detallado
Información completa registrada según configuración:
- **Configuración aplicada**: Muestra toda la configuración de logging al inicio
- **Procesamiento de datos**: 
  - Detección y validación de `dataIn` y `dataOut`
  - Creación automática de directorios
  - Validación de archivos y URLs
- **Operaciones de archivo**: Creación, carga y guardado de archivos Excel
- **Manejo de errores**: Errores detallados con contexto específico
- **Rendimiento**: Timestamps y duración de operaciones

### Error: "El archivo XML no es válido según el esquema XSD"
El XML no cumple con la estructura requerida. Verificar:
- Elementos obligatorios presentes
- Tipos de datos correctos (números, colores hexadecimales)
- Estructura de etiquetas válida

### Error: "El archivo X no existe"
Verificar que las rutas especificadas en `dataIn` sean correctas y los archivos existan.

### Error: "URL no accesible"
Para `dataOut` con `URL://`, verificar:
- La URL es correcta y accesible
- El servidor acepta conexiones
- No hay problemas de red o firewall

### Error: "Error decodificando BASE64"
Para `dataIn` con `BASE64://`, verificar:
- El contenido está correctamente codificado en BASE64
- No hay caracteres extraños o saltos de línea

### Caracteres especiales no se muestran correctamente
Asegurar que el archivo XML esté codificado en UTF-8.

### El logging no funciona
Verificar:
- La ruta especificada en `logFile` es escribible
- El directorio del archivo de log existe (se crea automáticamente)
- El nivel de log es correcto (DEBUG, INFO, WARNING, ERROR)
- La configuración de `logConsole` está en "true" o "false"

### No se crean los directorios automáticamente
El sistema debería crear directorios automáticamente. Si no lo hace:
- Verificar permisos de escritura en el directorio padre
- Revisar los logs para mensajes de error específicos
- Asegurarse de que la ruta no contenga caracteres inválidos

### Los logs no muestran información detallada
Para obtener información completa:
- Configurar `logLevel` en "DEBUG" para máximo detalle
- Verificar que `logConsole` esté en "true" para ver logs en pantalla
- El archivo de log siempre contiene información completa independientemente del nivel

## Soporte técnico

Para reportar problemas o solicitar funcionalidades adicionales, contactar con el equipo de desarrollo.