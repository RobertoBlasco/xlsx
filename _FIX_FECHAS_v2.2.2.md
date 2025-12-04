# Fix: Reconocimiento de Fechas y Números como Tipos Nativos en Excel

## Problema Reportado

Cuando se especificaba `format="dd/mm/yyyy"` en el XML, Excel reconocía la celda como TEXTO en lugar de como FECHA. Esto impedía realizar operaciones de fecha y ordenamiento correcto en Excel.

## Causa Raíz

El código anterior asignaba todos los valores directamente como strings desde el XML.

Para que Excel reconozca un valor como fecha, número o porcentaje, **debe recibir el tipo de dato correspondiente** junto con el `number_format`, no un string.

## Solución Implementada (v2.2.2)

Se agregó conversión automática de tipos basada en el formato especificado:

Ahora se analiza el `format` especificado y convierte el valor string al tipo apropiado:

#### Para Fechas
- **Detecta formatos de fecha**: `dd/mm/yyyy`, `dd-mm-yyyy`, `yyyy-mm-dd`, `dd.mm.yyyy`, etc.
- **Convierte el valor string a objeto `datetime`** de Python
- **Soporta múltiples formatos de entrada**:
  - ISO: `2023-01-15`
  - Europeo: `15/01/2023`, `15-01-2023`
  - Compacto: `20230115`
  - Y más...

#### Para Números
- **Detecta formatos numéricos**: `0`, `0.00`, `#,##0`, `€#,##0.00`, `0.00%`, etc.
- **Convierte a `int` o `float`** según corresponda
- **Limpia separadores de miles** antes de convertir

## Formatos Soportados

### Fechas
- `dd/mm/yyyy`, `dd-mm-yyyy`, `dd.mm.yyyy`
- `mm/dd/yyyy`, `mm-dd-yyyy`
- `yyyy/mm/dd`, `yyyy-mm-dd`
- `d/m/yy`, `dd/mm/yy`

### Números
- Enteros: `0`
- Decimales: `0.00`, `0.000`
- Con separadores: `#,##0`, `#,##0.00`
- Monedas: `€#,##0.00`, `$#,##0.00`, `£#,##0.00`, `¥#,##0.00`
- Porcentajes: `0%`, `0.00%`

### Formatos de Entrada de Valores

Los valores en el XML pueden venir en estos formatos:

**Para fechas:**
- ISO: `2023-01-15`
- Europeo: `15/01/2023`, `15-01-2023`
- Punto: `15.01.2023`
- Compacto: `20230115`
- Corto: `15/01/23`, `15-01-23`

**Para números:**
- Enteros: `100`, `1500`
- Decimales: `45.50`, `125.75`
- Con separadores: `1,500`, `1 500`

## Logging y Depuración

Con `logLevel=DEBUG`, ahora se muestra información sobre las conversiones:

```
DEBUG - Valor '2023-01-15' convertido a fecha usando formato '%Y-%m-%d'
DEBUG - Aplicando formato 'dd/mm/yyyy' a celda B2
DEBUG - Valor '1500' convertido a int: 1500
DEBUG - Aplicando formato '€#,##0.00' a celda D2
WARNING - No se pudo convertir '15/99/2023' a fecha. Se guardará como texto.
```

## Ejemplo de Uso Completo

```xml
<?xml version='1.0' encoding='utf-8'?>
<ineoDoc task="updateXlsx" task_id="ejemplo_fechas">
    <data>
        <dataOut>FILE://./resultado.xlsx</dataOut>
    </data>
    <log>
        <logLevel>DEBUG</logLevel>
        <logConsole>true</logConsole>
    </log>

    <workbooks>
        <workbook name="Ventas">
            <columnSettings>
                <column name="A" width="20" />
                <column name="B" width="15" format="dd/mm/yyyy" />
                <column name="C" width="12" format="0" />
                <column name="D" width="12" format="€#,##0.00" />
            </columnSettings>

            <!-- Headers -->
            <cell row="1" column="A" value="Producto" />
            <cell row="1" column="B" value="Fecha Venta" />
            <cell row="1" column="C" value="Cantidad" />
            <cell row="1" column="D" value="Importe" />

            <!-- Datos - las fechas y números serán reconocidos correctamente -->
            <cell row="2" column="A" value="Laptop" />
            <cell row="2" column="B" value="2023-01-15" />    <!-- Fecha reconocida -->
            <cell row="2" column="C" value="10" />            <!-- Número reconocido -->
            <cell row="2" column="D" value="1500.50" />       <!-- Número reconocido -->

            <cell row="3" column="A" value="Mouse" />
            <cell row="3" column="B" value="20/03/2023" />    <!-- Fecha reconocida -->
            <cell row="3" column="C" value="25" />            <!-- Número reconocido -->
            <cell row="3" column="D" value="625.75" />        <!-- Número reconocido -->
        </workbook>
    </workbooks>
</ineoDoc>
```

## Resultado en Excel

 **Columna B (Fechas)**: Excel reconoce como tipo FECHA
   - Ordenamiento cronológico funciona correctamente
   - Operaciones de fecha disponibles (sumar días, restar fechas, etc.)
   - Formato de visualización: `dd/mm/yyyy`

 **Columna C (Cantidades)**: Excel reconoce como tipo NÚMERO
   - Operaciones matemáticas funcionan correctamente
   - Alineación automática a la derecha
   - Formato de visualización: sin decimales

 **Columna D (Importes)**: Excel reconoce como tipo NÚMERO
   - Operaciones matemáticas funcionan correctamente
   - Formato de visualización: `€#,##0.00`
   - Muestra símbolo de euro y 2 decimales

## Retrocompatibilidad

✅ La solución es **100% retrocompatible**:
- Si no se especifica formato, se usa `'General'` y el valor se guarda como string
- XMLs antiguos sin formatos específicos siguen funcionando igual
- Solo afecta cuando se especifica un formato de fecha o número

## Manejo de Errores

Si el valor no puede convertirse según el formato especificado:
1. Se registra un WARNING en el log (si está habilitado)
2. Se guarda el valor como string (comportamiento anterior)
3. La conversión continúa sin interrupciones

Ejemplo:
```
WARNING - No se pudo convertir '32/13/2023' a fecha. Formatos intentados: [...]. Se guardará como texto.
```

## Testing
Se incluye archivo de test: `test_fechas.xml`

Para probar:
```bash
ineoXlsxCmdLine.exe test_fechas.xml
```

Verifica en el Excel generado que:
- Las fechas están reconocidas como tipo FECHA
- Los números están reconocidos como tipo NÚMERO
- El formato de visualización se aplica correctamente

## Recomendaciones

1. **Usar formatos estándar**: `dd/mm/yyyy` para fechas, `0` o `0.00` para números
2. **Valores de entrada consistentes**: Preferir ISO (`2023-01-15`) para fechas en el XML
3. **Activar logging DEBUG**: Para ver el proceso de conversión y detectar problemas
4. **Validar datos**: Asegurarse que los valores en el XML sean válidos

## Soporte

Para cualquier problema con la conversión de tipos:
1. Activar `logLevel=DEBUG` en el XML
2. Revisar el archivo de log generado
3. Verificar que el formato especificado sea correcto
4. Comprobar que el valor de entrada sea válido

---

**IneoXlsx v2.2.2** - Fix de reconocimiento de tipos de datos nativos
**Ineo Solutions S.L.** | info@ineosolutions.es | 2025
