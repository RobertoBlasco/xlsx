# Conversor XML a Excel - Plugin IneoXlsx v2.2.1

Plugin avanzado para convertir archivos XML estructurados a archivos Excel (.xlsx) con soporte completo para estilos, formato de tablas, configuración de columnas/filas y múltiples funcionalidades profesionales.

**Versión 2.2.0** (Revisión: 20250829) incluye características avanzadas como ajuste automático configurable, 60 estilos de tablas predefinidos, gestión inteligente de archivos existentes y sistema de precedencia completo.

**Desarrollado por**: Ineo Solutions S.L. | info@ineosolutions.es | 2025

## Funcionalidades Principales

### Nuevas Características v2.2.0
- **Formato de tablas**: Creación de tablas Excel con filtros automáticos y estilos profesionales (60 estilos predefinidos)
- **Configuración de columnas y filas**: Aplicación de propiedades a columnas y filas completas
- **Sistema de precedencia**: Cell > Row > Column > Default para todas las propiedades
- **Altura de filas**: Control completo de alturas personalizadas
- **Ancho de columnas**: Configuración precisa de anchos
- **Ajuste automático configurable**: Control granular del ajuste automático de columnas con opciones globales y por hoja
- **Alineación completa**: Horizontal y vertical siguiendo estándares Excel
- **Estilos de texto avanzados**: Bold, italic, underline, strikethrough
- **Gestión inteligente de archivos existentes**: Trabajo no destructivo con preservación total de contenido
- **Sistema de logging mejorado**: Logging detallado con diferentes niveles y transparencia completa

### Características Base
- Conversión de XML a Excel con estilos personalizados
- Soporte para múltiples hojas de cálculo
- Preservación de archivos Excel existentes
- Validación automática con esquema XSD
- **Detección automática de encoding**: Compatible con UTF-8, ISO-8859-1, Windows-1252 y otros
- Ejecutable independiente (.exe) sin dependencias de Python

## Uso

### Sintaxis básica
```bash
ineoXlsxCmdLine.exe <archivo_xml> [archivo_excel]
```

### Ejemplos de uso

#### Uso básico
```bash
# Convertir XML a Excel
ineoXlsxCmdLine.exe datos.xml

# Especificar archivo de salida
ineoXlsxCmdLine.exe datos.xml resultado.xlsx
```

#### Uso avanzado con configuración completa
```bash
# Procesar con logging detallado
ineoXlsxCmdLine.exe configuracion_avanzada.xml
```

## Estructura XML

### Estructura Completa (Recomendada)
```xml
<?xml version="1.0" encoding="UTF-8"?>
<ineoDoc task="updateXlsx" task_id="conversion_001">
    <data>
        <dataIn>FILE://datos/empleados.xml</dataIn>
        <dataOut>FILE://output/empleados_procesados.xlsx</dataOut>
    </data>
    <responseOut>FILE://./respuesta.log</responseOut>
    <log>
        <logLevel>DEBUG</logLevel>
        <logFile>FILE://./conversion.log</logFile>
        <logFormat>%(asctime)s - %(name)s - %(levelname)s - %(message)s</logFormat>
        <logDateFormat>%Y-%m-%d %H:%M:%S</logDateFormat>
        <logConsole>true</logConsole>
    </log>
    <workbooks>
        <styles>
            <!-- Definición de estilos -->
        </styles>
        <workbook name="NombreHoja">
            <columnSettings>
                <!-- Configuración de columnas -->
            </columnSettings>
            <rowSettings>
                <!-- Configuración de filas -->
            </rowSettings>
            <table name="MiTabla" ref="A1:E10" style="TableStyleMedium9" 
                   showRowStripes="true" showColumnStripes="false" />
            <!-- Celdas del contenido -->
        </workbook>
    </workbooks>
</ineoDoc>
```

## Estilos y Formato

### Definición de Estilos
```xml
<styles>
    <style id="header">
        <font>Arial</font>
        <size>12</size>
        <bold>true</bold>
        <italic>false</italic>
        <underline>false</underline>
        <strikethrough>false</strikethrough>
        <color>#FFFFFF</color>
        <background>#4472C4</background>
    </style>
    <style id="data">
        <font>Calibri</font>
        <size>10</size>
        <bold>false</bold>
        <italic>true</italic>
        <color>#000000</color>
    </style>
    <style id="highlight">
        <font>Times New Roman</font>
        <size>11</size>
        <bold>true</bold>
        <italic>true</italic>
        <underline>true</underline>
        <strikethrough>true</strikethrough>
        <color>#C00000</color>
        <background>#FFFF99</background>
    </style>
</styles>
```

#### Propiedades de Estilo Soportadas
- `font`: Nombre de la fuente (Arial, Calibri, Times New Roman, etc.)
- `size`: Tamaño de fuente en puntos
- `bold`: `true`/`false` para texto en negrita
- `italic`: `true`/`false` para texto en cursiva
- `underline`: `true`/`false` para texto subrayado
- `strikethrough`: `true`/`false` para texto tachado
- `color`: Color del texto en formato hexadecimal (#000000, #FF0000, etc.)
- `background`: Color de fondo en formato hexadecimal

## Configuración de Columnas y Filas

### Configuración de Columnas
```xml
<columnSettings>
    <column name="A" width="20" horizontalAlignment="left" 
            verticalAlignment="center" format="General" style="header" />
    <column name="B" width="8" horizontalAlignment="right" 
            verticalAlignment="center" format="0" style="numeric" />
    <column name="C" width="15" horizontalAlignment="center" 
            verticalAlignment="bottom" format="dd/mm/yyyy" style="date" />
    <column name="D" width="12" horizontalAlignment="right" 
            verticalAlignment="center" format="€#,##0.00" style="currency" />
</columnSettings>
```

### Configuración de Filas
```xml
<rowSettings>
    <row number="1" height="25" horizontalAlignment="center" 
         verticalAlignment="center" style="header" />
    <row number="2" height="18" horizontalAlignment="left" 
         verticalAlignment="top" format="General" />
    <row number="5" height="30" style="highlight" />
</rowSettings>
```

#### Propiedades de Columnas/Filas
- `width`/`height`: Ancho de columna o alto de fila en puntos
- `horizontalAlignment`: `left`, `center`, `right`, `justify`
- `verticalAlignment`: `top`, `center`, `bottom`, `justify`
- `format`: Formato de número/fecha (General, 0, dd/mm/yyyy, €#,##0.00, 0.00%, etc.)
- `style`: ID del estilo a aplicar

## Tablas Profesionales

### Definición de Tablas
```xml
<table name="TablaEmpleados" ref="A1:G11" style="TableStyleMedium9" 
       showFirstColumn="false" showLastColumn="false" 
       showRowStripes="true" showColumnStripes="false" />
```

#### Propiedades de Tablas
- `name`: Nombre único de la tabla (requerido)
- `ref`: Rango de celdas (ej: A1:G11) (requerido)
- `style`: Estilo de tabla predefinido (TableStyleMedium9, TableStyleLight1, etc.)
- `showFirstColumn`: Destacar primera columna (`true`/`false`)
- `showLastColumn`: Destacar última columna (`true`/`false`)
- `showRowStripes`: Filas alternadas (`true`/`false`)
- `showColumnStripes`: Columnas alternadas (`true`/`false`)

#### Estilos de Tabla Predefinidos

Excel proporciona **60 estilos de tabla predeterminados** organizados en tres categorías:

##### Serie TableStyleLight (21 estilos)
Estilos con colores claros y sutiles:
```
TableStyleLight1, TableStyleLight2, TableStyleLight3, TableStyleLight4, TableStyleLight5,
TableStyleLight6, TableStyleLight7, TableStyleLight8, TableStyleLight9, TableStyleLight10,
TableStyleLight11, TableStyleLight12, TableStyleLight13, TableStyleLight14, TableStyleLight15,
TableStyleLight16, TableStyleLight17, TableStyleLight18, TableStyleLight19, TableStyleLight20,
TableStyleLight21
```

##### Serie TableStyleMedium (28 estilos)
Estilos con colores de intensidad media (más populares):
```
TableStyleMedium1, TableStyleMedium2, TableStyleMedium3, TableStyleMedium4, TableStyleMedium5,
TableStyleMedium6, TableStyleMedium7, TableStyleMedium8, TableStyleMedium9, TableStyleMedium10,
TableStyleMedium11, TableStyleMedium12, TableStyleMedium13, TableStyleMedium14, TableStyleMedium15,
TableStyleMedium16, TableStyleMedium17, TableStyleMedium18, TableStyleMedium19, TableStyleMedium20,
TableStyleMedium21, TableStyleMedium22, TableStyleMedium23, TableStyleMedium24, TableStyleMedium25,
TableStyleMedium26, TableStyleMedium27, TableStyleMedium28
```

##### Serie TableStyleDark (11 estilos)
Estilos con colores oscuros y contrastantes:
```
TableStyleDark1, TableStyleDark2, TableStyleDark3, TableStyleDark4, TableStyleDark5,
TableStyleDark6, TableStyleDark7, TableStyleDark8, TableStyleDark9, TableStyleDark10,
TableStyleDark11
```

##### Ejemplos de Uso por Categoría
```xml
<!-- Estilo claro y sutil -->
<table name="TablaLight" ref="A1:D10" style="TableStyleLight15" 
       showRowStripes="true" showColumnStripes="false" />

<!-- Estilo medio - más popular para uso general -->
<table name="TablaMedium" ref="A1:D10" style="TableStyleMedium9" 
       showRowStripes="true" showColumnStripes="false" />

<!-- Estilo oscuro para mayor contraste -->
<table name="TablaDark" ref="A1:D10" style="TableStyleDark5" 
       showRowStripes="true" showColumnStripes="false" />
```

##### Recomendaciones de Uso
- **TableStyleLight**: Ideal para documentos formales y presentaciones profesionales
- **TableStyleMedium**: Perfecto para uso general, balance entre visibilidad y elegancia
- **TableStyleDark**: Excelente para datos que requieren alto contraste y visibilidad

#### Características Automáticas de Tablas
- **Filtros automáticos**: Se aplican automáticamente en los headers
- **Estilos profesionales**: 60 estilos predefinidos de Excel
- **Bandas alternadas**: Configurables por filas o columnas
- **Integración completa**: Compatible con todos los estilos y formatos

## Definición de Celdas

### Sintaxis de Celdas
```xml
<cell row="1" column="A" value="Nombre Completo" 
      style="header" format="General"
      width="20" 
      horizontalAlignment="center" verticalAlignment="center" />
```

#### Atributos de Celda
- `row`: Número de fila (empezando desde 1) **[requerido]**
- `column`: Letra de columna (A, B, C, ...) **[requerido]**
- `value`: Contenido de la celda **[requerido]**
- `style`: ID del estilo a aplicar (opcional)
- `format`: Formato de número específico (opcional)
- `width`: Ancho de columna específico (opcional)
- `horizontalAlignment`: Alineación horizontal específica (opcional)
- `verticalAlignment`: Alineación vertical específica (opcional)

## Sistema de Precedencia

El sistema aplica propiedades siguiendo este orden de precedencia:

**Cell > Row > Column > Default**

### Ejemplo de Precedencia
```xml
<!-- Configuración de columna B: width=10, format="0" -->
<column name="B" width="10" format="0" style="numeric" />

<!-- Configuración de fila 3: height=25, style="highlight" -->
<row number="3" height="25" style="highlight" />

<!-- Celda B3: width=15 (sobrescribe columna), height=25 (hereda fila) -->
<cell row="3" column="B" value="42" width="15" format="0.00" />
```

**Resultado**: La celda B3 tendrá:
- `width`: 15 (de la celda, sobrescribe columna)
- `height`: 25 (de la fila)
- `format`: "0.00" (de la celda, sobrescribe columna)
- `style`: "highlight" (de la fila)

## Configuración de Datos

### Tipos de Fuentes Soportadas

#### dataIn (archivo de entrada)
- `FILE://ruta/archivo.xml` - Archivo local
- `BASE64://contenido_codificado` - Contenido XML codificado en BASE64
- `ruta/archivo.xml` - Archivo local (por defecto)

#### dataOut (archivo de salida)
- `FILE://ruta/archivo.xlsx` - Archivo local
- `URL://https://servidor.com/api/upload` - URL para envío
- `ruta/archivo.xlsx` - Archivo local (por defecto)

### Sistema de Opciones Globales

El sistema permite configurar opciones globales que afectan el comportamiento de conversión:

```xml
<options>
    <option name="autoAdjustColumnWidth" value="true"/>
    <option name="maxColumnWidth" value="50"/>
    <option name="minColumnWidth" value="8"/>
</options>
```

#### Opciones Disponibles

##### Ajuste Automático de Columnas
- **autoAdjustColumnWidth**: `true`/`false` (por defecto: `true`)
  - Controla si las columnas se ajustan automáticamente según su contenido
  - Se puede sobrescribir a nivel de workbook individual
  - Las columnas con `width` explícito se excluyen del ajuste automático

- **maxColumnWidth**: Número entero (por defecto: `50`)
  - Ancho máximo permitido para el ajuste automático
  - Previene columnas excesivamente anchas

- **minColumnWidth**: Número entero (por defecto: `8`)
  - Ancho mínimo garantizado para el ajuste automático
  - Asegura que las columnas no sean demasiado estrechas

##### Configuración por Workbook

Cada hoja puede sobrescribir la configuración global:

```xml
<!-- Usa configuración global -->
<workbook name="HojaGlobal">
    <!-- ... contenido ... -->
</workbook>

<!-- Sobrescribe: desactiva ajuste automático -->
<workbook name="HojaManual" autoAdjustColumnWidth="false">
    <!-- ... contenido ... -->
</workbook>

<!-- Sobrescribe: activa ajuste automático -->
<workbook name="HojaAuto" autoAdjustColumnWidth="true">
    <!-- ... contenido ... -->
</workbook>
```

##### Precedencia del Ajuste Automático

1. **Columnas con width explícito**: Nunca se ajustan automáticamente
2. **Configuración por workbook**: Sobrescribe la configuración global
3. **Configuración global**: Valor por defecto para todas las hojas

##### Ejemplo Completo de Configuración

```xml
<?xml version='1.0' encoding='utf-8'?>
<ineoDoc task="updateXlsx" task_id="auto_adjust_example">
    <!-- Configuración global -->
    <options>
        <option name="autoAdjustColumnWidth" value="true"/>
        <option name="maxColumnWidth" value="30"/>
        <option name="minColumnWidth" value="10"/>
    </options>
    
    <workbooks>
        <!-- Hoja con ajuste automático habilitado -->
        <workbook name="Datos">
            <columnSettings>
                <!-- Esta columna tiene width fijo, NO se ajustará -->
                <column name="A" width="25" />
            </columnSettings>
            
            <cell row="1" column="A" value="Width fijo: 25" />
            <cell row="1" column="B" value="Se ajustará automáticamente" />
            <cell row="1" column="C" value="También automático" />
        </workbook>
        
        <!-- Hoja con ajuste automático deshabilitado -->
        <workbook name="Manual" autoAdjustColumnWidth="false">
            <cell row="1" column="A" value="Sin ajuste automático" />
            <cell row="1" column="B" value="Contenido muy largo que NO se ajustará" />
        </workbook>
    </workbooks>
</ineoDoc>
```

##### Logging del Ajuste Automático

El sistema proporciona información detallada sobre las decisiones de ajuste:

```
DEBUG - Ajuste automático columna B: width=30 (contenido max: 85)
DEBUG - Ajuste automático columna C: width=10 (contenido max: 5)
INFO - Ajuste automático aplicado a hoja 'Datos' (excluidas 1 columnas con width explícito)
INFO - Ajuste automático deshabilitado para hoja 'Manual'
```

### Sistema de Logging Avanzado

```xml
<log>
    <logLevel>DEBUG</logLevel>
    <logFile>FILE://./conversion.log</logFile>
    <logFormat>%(asctime)s - %(name)s - %(levelname)s - %(message)s</logFormat>
    <logDateFormat>%Y-%m-%d %H:%M:%S</logDateFormat>
    <logConsole>true</logConsole>
</log>
```

#### Niveles de Log
- `DEBUG`: Información detallada (aplicación de estilos, precedencias, etc.)
- `INFO`: Información general del proceso
- `WARNING`: Advertencias no críticas
- `ERROR`: Errores que impiden la conversión

## Ejemplo Completo Avanzado

```xml
<?xml version='1.0' encoding='utf-8'?>
<ineoDoc task="updateXlsx" task_id="ejemplo_completo_001">
    <data>
        <dataOut>FILE://./output/empleados_completo.xlsx</dataOut>
    </data>
    <log>
        <logLevel>DEBUG</logLevel>
        <logFile>FILE://./conversion.log</logFile>
        <logFormat>%(asctime)s - %(name)s - %(levelname)s - %(message)s</logFormat>
        <logConsole>true</logConsole>   
    </log>
    
    <!-- Configuración de opciones globales -->
    <options>
        <option name="autoAdjustColumnWidth" value="true"/>
        <option name="maxColumnWidth" value="35"/>
        <option name="minColumnWidth" value="12"/>
    </options>
    
    <workbooks>
        <styles>
            <style id="header">
                <font>Arial</font>
                <size>12</size>
                <bold>true</bold>
                <color>#FFFFFF</color>
                <background>#4472C4</background>
            </style>
            <style id="data">
                <font>Arial</font>
                <size>10</size>
                <italic>true</italic>
                <color>#333333</color>
            </style>
            <style id="currency">
                <font>Arial</font>
                <size>10</size>
                <bold>true</bold>
                <color>#006400</color>
            </style>
        </styles>
        
        <workbook name="Empleados">
            <!-- Configuración de columnas -->
            <columnSettings>
                <column name="A" width="20" horizontalAlignment="left" 
                        verticalAlignment="center" style="data" />
                <column name="B" width="8" horizontalAlignment="right" 
                        verticalAlignment="center" format="0" />
                <column name="C" width="15" horizontalAlignment="center" 
                        verticalAlignment="center" format="dd/mm/yyyy" />
                <column name="D" width="12" horizontalAlignment="right" 
                        verticalAlignment="center" format="€#,##0.00" style="currency" />
            </columnSettings>
            
            <!-- Configuración de filas -->
            <rowSettings>
                <row number="1" height="25" horizontalAlignment="center" 
                     verticalAlignment="center" style="header" />
                <row number="2" height="18" />
            </rowSettings>
            
            <!-- Definición de tabla -->
            <table name="TablaEmpleados" ref="A1:D5" style="TableStyleMedium9" 
                   showFirstColumn="false" showLastColumn="false" 
                   showRowStripes="true" showColumnStripes="false" />
            
            <!-- Headers (heredan configuración de fila 1) -->
            <cell row="1" column="A" value="Nombre Completo" />
            <cell row="1" column="B" value="Edad" />
            <cell row="1" column="C" value="Fecha Ingreso" />
            <cell row="1" column="D" value="Salario" />
            
            <!-- Datos (heredan configuraciones de columnas) -->
            <cell row="2" column="A" value="Ana García" />
            <cell row="2" column="B" value="28" />
            <cell row="2" column="C" value="2023-01-15" />
            <cell row="2" column="D" value="45000" />
            
            <cell row="3" column="A" value="Carlos López" />
            <cell row="3" column="B" value="35" />
            <cell row="3" column="C" value="2022-03-22" />
            <cell row="3" column="D" value="52000" />
            
            <cell row="4" column="A" value="María Fernández" />
            <cell row="4" column="B" value="31" />
            <cell row="4" column="C" value="2021-07-10" />
            <cell row="4" column="D" value="48000" />
            
            <cell row="5" column="A" value="José Martín" />
            <cell row="5" column="B" value="29" />
            <cell row="5" column="C" value="2023-11-05" />
            <cell row="5" column="D" value="46500" />
        </workbook>
    </workbooks>
</ineoDoc>
```

## Validación y Errores

### Validación Automática XSD
El sistema valida automáticamente la estructura XML antes del procesamiento, proporcionando errores específicos.

### Errores Comunes

#### "El archivo XML no es válido según el esquema XSD"
- Verificar elementos obligatorios
- Comprobar tipos de datos (números, colores hexadecimales)
- Validar estructura de etiquetas

#### "Error creando tabla"
- Verificar que el rango `ref` sea válido (ej: A1:G11)
- Asegurar que el nombre de tabla sea único
- Comprobar que las celdas del rango existan

#### Problemas de precedencia
- Revisar configuraciones de columnas, filas y celdas
- Verificar que los IDs de estilos existan
- Comprobar formatos de alineación y formato de números

## Características Avanzadas

### Gestión Inteligente de Archivos Excel

#### Trabajo con Archivos Existentes
El sistema está diseñado para trabajar de forma **no destructiva** con archivos Excel existentes:

- **Preservación total**: Todo el contenido existente se mantiene intacto
- **Adición incremental**: Solo se añade o actualiza el contenido especificado en el XML
- **Reutilización de hojas**: Las hojas existentes se reutilizan automáticamente
- **Creación selectiva**: Solo crea nuevas hojas cuando es necesario

#### Comportamientos Específicos

**Cuando el archivo Excel existe**:
```
Cargando archivo Excel existente: ./reportes/informe_mensual.xlsx
  Utilizando hoja existente: Ventas
  Creando nueva hoja: AnalisisNuevo
  Utilizando hoja existente: Resumen
```

**Cuando el archivo no existe**:
```
Creando nuevo archivo Excel: ./reportes/informe_nuevo.xlsx
  Creando nueva hoja: Datos
  Creando nueva hoja: Graficos
```

#### Casos de Uso Reales

##### 1. Actualización de Reportes Existentes
```xml
<!-- Añadir datos del mes actual a reporte existente -->
<workbook name="ReporteMensual">
    <cell row="13" column="A" value="Febrero 2025" />
    <cell row="13" column="B" value="145000" />
    <cell row="13" column="C" value="12%" />
</workbook>
```

##### 2. Añadir Nuevas Secciones
```xml
<!-- Crear nueva hoja de análisis en workbook existente -->
<workbook name="TendenciasQ1" autoAdjustColumnWidth="true">
    <table name="TendenciasVentas" ref="A1:D10" style="TableStyleMedium12" 
           showRowStripes="true" showColumnStripes="false" />
    <cell row="1" column="A" value="Mes" />
    <cell row="1" column="B" value="Ventas" />
    <!-- ... más datos ... -->
</workbook>
```

##### 3. Reconfiguración de Hojas Existentes
```xml
<!-- Cambiar configuración de ajuste automático en hoja existente -->
<workbook name="HojaExistente" autoAdjustColumnWidth="false">
    <!-- Añadir nuevas celdas sin reajustar columnas -->
    <cell row="20" column="A" value="Datos adicionales" />
</workbook>
```

#### Integración con Todas las Funcionalidades

**Todas las características funcionan sobre archivos existentes**:
- **Ajuste automático configurable**: Recalcula considerando contenido nuevo + existente
- **Tablas profesionales**: Se pueden añadir a hojas existentes
- **Configuraciones de columnas/filas**: Se aplican respetando contenido previo
- **Estilos avanzados**: Se añaden sin afectar formatos existentes
- **Sistema de precedencia**: Funciona con contenido nuevo y existente

#### Ventajas para Automatización

- **Reportes incrementales**: Ideal para automatización de reportes que se actualizan regularmente
- **No destructivo**: Nunca se pierde información existente
- **Selectivo**: Solo actualiza lo especificado en el XML
- **Eficiente**: No recrea todo el archivo, solo modifica lo necesario
- **Flexible**: Permite tanto actualizaciones menores como reestructuraciones mayores

#### Directorios y Rutas
- **Creación automática**: Los directorios necesarios se crean automáticamente
- **Validación de rutas**: Verifica accesibilidad antes del procesamiento
- **Manejo de errores**: Logging detallado en caso de problemas de acceso

### Procesamiento BASE64 y URLs
- **BASE64**: Soporte completo para contenido XML embebido
- **URLs**: Envío directo de resultados a endpoints
- **Validación**: Verificación automática de accesibilidad

### Optimizaciones de Rendimiento
- **Ajuste automático configurable**: Ancho de columnas basado en contenido con control granular
- **Respeto de configuración manual**: Las columnas con width explícito se preservan
- **Configuración por hoja**: Control individual del ajuste automático
- **Límites configurables**: Anchos mínimos y máximos personalizables
- **Logging eficiente**: Sistema de logging con múltiples niveles
- **Validación temprana**: Detección de errores antes del procesamiento

## Rendimiento y Logs

### Información de Rendimiento
El sistema proporciona métricas detalladas:
- Tiempo total de conversión
- Número de celdas procesadas
- Configuraciones aplicadas
- Tablas creadas

### Ejemplo de Logs Detallados
```
INFO - Iniciando conversión XML a Excel
INFO - Configuración de logging aplicada: {'logLevel': 'DEBUG', 'logFile': './conversion.log', 'logConsole': 'true'}
INFO - No se especificó dataIn, usando archivo de configuración como datos
INFO - Archivo Excel de salida: ./output/empleados.xlsx
INFO - Workbook 'Empleados': Hoja creada, procesando 20 celdas
INFO - Configuraciones de columna aplicadas: ['A', 'B', 'C', 'D']
INFO - Configuraciones de fila aplicadas: [1, 2]
DEBUG - Aplicando width '20' a columna A
DEBUG - Aplicando height '25' a fila 1
DEBUG - Aplicando alineamiento horizontal=center, vertical=center a celda A1
INFO - Tabla 'TablaEmpleados' creada en rango A1:D5 con estilo TableStyleMedium9
INFO - Archivo Excel creado exitosamente: ./output/empleados.xlsx
```

## Compatibilidad

### Estructura Básica (Retrocompatible)
```xml
<?xml version="1.0" encoding="UTF-8"?>
<workbooks>
    <styles>
        <style id="1">
            <font>Arial</font>
            <bold>true</bold>
            <color>#000000</color>
        </style>
    </styles>
    <workbook name="Datos">
        <cell row="1" column="A" value="Título" style="1"/>
    </workbook>
</workbooks>
```

## Changelog

### v2.2.1 (Revisión: 20250926)
**Parche de Compatibilidad:**
- **Detección automática de encoding**: Soporte inteligente para múltiples codificaciones XML
  - Detecta automáticamente encoding basándose en la declaración XML (`<?xml encoding="..." ?>`)
  - Soporte para UTF-8, ISO-8859-1, Windows-1252 y otras codificaciones estándar
  - Fallback automático a UTF-8 si no se encuentra declaración de encoding
  - Manejo robusto de archivos XML en diferentes formatos (ANSI, Unicode, etc.)
- **Mejora en validación XSD**: Procesamiento más confiable de archivos con diferentes encodings
- **Logging mejorado**: Información sobre encoding detectado para transparencia en el procesamiento

### v2.2.0 (Revisión: 20250829)
**Funcionalidades Principales Añadidas:**
- **Ajuste automático configurable**: Control granular del ajuste de columnas con opciones globales y por workbook
- **60 estilos de tablas predefinidos**: Lista completa de estilos TableStyleLight, TableStyleMedium y TableStyleDark
- **Gestión inteligente de archivos existentes**: Trabajo no destructivo con preservación total de contenido
- **Sistema de precedencia completo**: Cell > Row > Column > Default para todas las propiedades
- **Altura de filas configurable**: Control completo con sistema de herencia
- **Estilos de texto avanzados**: Bold, italic, underline, strikethrough completamente implementados
- **Logging detallado mejorado**: Transparencia completa de decisiones y configuraciones aplicadas
- **Configuraciones de columnas/filas**: Todas las propiedades (width, height, alignment, format, style)
- **Tablas profesionales**: Filtros automáticos, bandas alternadas y estilos personalizables

**Mejoras de Rendimiento:**
- Optimización del ajuste automático para respetar configuraciones manuales
- Validación temprana de errores antes del procesamiento
- Logging eficiente con múltiples niveles de detalle

**Compatibilidad:**
- Retrocompatibilidad completa con XMLs existentes
- Estructura básica mantenida para migración sin cambios

### Versiones Anteriores
- **v2.1.1**: Funcionalidades base con estilos y formato básico
- **v2.0.0**: Primera versión con soporte de tablas y configuraciones avanzadas

## Soporte

Para reportar problemas, solicitar funcionalidades o obtener soporte técnico, contactar con el equipo de desarrollo.

**Ineo Solutions S.L.**  
Email: info@ineosolutions.es  
Año: 2025

---
**IneoXlsx v2.2.0** - Conversor XML a Excel Profesional