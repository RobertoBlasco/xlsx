#!/usr/bin/env python3
"""
Programa optimizado para convertir XML a Excel usando xlsxwriter
Lee el archivo ineo_xlsx.xml y genera salida.xlsx
"""

import xml.etree.ElementTree as ET
import xlsxwriter
import time
from collections import defaultdict

def parse_styles(styles_element):
    """Parsea los estilos del XML y retorna un diccionario"""
    styles_dict = {}
    
    for style in styles_element.findall('style'):
        style_id = style.get('id')
        style_data = {}
        
        # Parsear propiedades del estilo
        font_elem = style.find('font')
        if font_elem is not None:
            style_data['font_name'] = font_elem.text
            
        size_elem = style.find('size')
        if size_elem is not None:
            style_data['font_size'] = int(size_elem.text)
            
        bold_elem = style.find('bold')
        if bold_elem is not None:
            style_data['bold'] = bold_elem.text.lower() == 'true'
            
        color_elem = style.find('color')
        if color_elem is not None:
            style_data['font_color'] = color_elem.text
            
        background_elem = style.find('background')
        if background_elem is not None:
            style_data['bg_color'] = background_elem.text
            
        alignment_elem = style.find('alignment')
        if alignment_elem is not None:
            style_data['align'] = alignment_elem.text
            
        styles_dict[style_id] = style_data
    
    return styles_dict

def parse_column_widths(columns_element):
    """Parsea las configuraciones de ancho de columna del XML"""
    column_widths = {}
    
    if columns_element is None:
        return column_widths
    
    for column in columns_element.findall('column'):
        col_letter = column.get('name')
        width = column.get('width')
        
        if col_letter and width:
            try:
                column_widths[col_letter] = float(width)
            except ValueError:
                print(f"Advertencia: Ancho inválido para columna {col_letter}: {width}")
    
    return column_widths

def create_xlsxwriter_formats(workbook, styles_dict):
    """Pre-compila todos los formatos de xlsxwriter"""
    formats = {}
    
    for style_id, style_data in styles_dict.items():
        format_props = {}
        
        # Configurar propiedades del formato
        if 'font_name' in style_data:
            format_props['font_name'] = style_data['font_name']
        if 'font_size' in style_data:
            format_props['font_size'] = style_data['font_size']
        if 'bold' in style_data:
            format_props['bold'] = style_data['bold']
        if 'font_color' in style_data:
            format_props['font_color'] = style_data['font_color']
        if 'bg_color' in style_data:
            format_props['bg_color'] = style_data['bg_color']
        if 'align' in style_data:
            format_props['align'] = style_data['align']
            
        # Crear formato
        formats[style_id] = workbook.add_format(format_props)
    
    return formats

def xml_to_excel_optimized(xml_file, excel_file):
    """Convierte el XML a Excel de forma optimizada"""
    try:
        print("Parseando XML...")
        # Parsear XML
        tree = ET.parse(xml_file)
        root = tree.getroot()
        
        # Obtener estilos
        styles_element = root.find('styles')
        styles_dict = parse_styles(styles_element) if styles_element is not None else {}
        
        # Obtener configuración de columnas
        columns_element = root.find('columns')
        column_widths = parse_column_widths(columns_element)
        
        print("Creando workbook Excel...")
        # Crear workbook con xlsxwriter
        workbook = xlsxwriter.Workbook(excel_file, {
            'constant_memory': True,  # Optimización para archivos grandes
            'tmpdir': '/tmp',         # Usar directorio temporal
        })
        
        # Pre-compilar todos los formatos
        print("Pre-compilando estilos...")
        formats = create_xlsxwriter_formats(workbook, styles_dict)
        
        # Procesar cada workbook
        for workbook_elem in root.findall('workbook'):
            sheet_name = workbook_elem.get('name', 'Hoja1')
            print(f"Procesando hoja: {sheet_name}")
            
            worksheet = workbook.add_worksheet(sheet_name)
            
            # Agrupar celdas por fila para escritura eficiente
            rows_data = defaultdict(dict)
            
            # Recopilar todas las celdas y agrupar por fila
            print("Agrupando celdas por fila...")
            for cell_elem in workbook_elem.findall('cell'):
                row = int(cell_elem.get('row')) - 1  # xlsxwriter usa índices base 0
                column = cell_elem.get('column')
                text = cell_elem.get('text', '')
                style_id = cell_elem.get('style')
                
                # Convertir columna a índice numérico
                col_index = ord(column) - ord('A')
                
                rows_data[row][col_index] = {
                    'value': text,
                    'format': formats.get(style_id) if style_id else None
                }
            
            # Escribir datos por filas (más eficiente)
            print("Escribiendo datos en Excel...")
            for row_num in sorted(rows_data.keys()):
                row_data = rows_data[row_num]
                for col_num in sorted(row_data.keys()):
                    cell_data = row_data[col_num]
                    
                    # Intentar convertir a número si es posible
                    value = cell_data['value']
                    try:
                        # Si es un número, convertirlo
                        if value.isdigit():
                            value = int(value)
                        elif value.replace('.', '').isdigit():
                            value = float(value)
                    except:
                        pass  # Mantener como string
                    
                    if cell_data['format']:
                        worksheet.write(row_num, col_num, value, cell_data['format'])
                    else:
                        worksheet.write(row_num, col_num, value)
            
            # Aplicar anchos de columna si están configurados
            if column_widths:
                print("Configurando anchos de columna...")
                for col_letter, width in column_widths.items():
                    if width == -1:
                        # Auto-calcular ancho basado en el contenido
                        col_index = ord(col_letter) - ord('A')
                        max_length = 0
                        
                        # Buscar el contenido más largo en esta columna
                        for row_num in rows_data.keys():
                            if col_index in rows_data[row_num]:
                                cell_value = str(rows_data[row_num][col_index]['value'])
                                max_length = max(max_length, len(cell_value))
                        
                        # Calcular ancho con padding (mínimo 8, máximo 50)
                        calculated_width = min(max(max_length + 2, 8), 50)
                        worksheet.set_column(f'{col_letter}:{col_letter}', calculated_width)
                        print(f"  Columna {col_letter}: ancho auto-calculado = {calculated_width}")
                    else:
                        worksheet.set_column(f'{col_letter}:{col_letter}', width)
                        print(f"  Columna {col_letter}: ancho fijo = {width}")
        
        print("Cerrando workbook...")
        workbook.close()
        print(f"Archivo Excel creado exitosamente: {excel_file}")
        return True
        
    except ET.ParseError as e:
        print(f"Error parsing XML: {e}")
        return False
    except Exception as e:
        print(f"Error creando Excel: {e}")
        return False

def main():
    """Función principal"""
    xml_file = "ineo_xlsx.xml"
    excel_file = "salida_optimized.xlsx"
    
    print(f"Convirtiendo {xml_file} a {excel_file} (versión optimizada)...")
    
    # Timestamp de inicio
    start_time = time.time()
    print(f"Inicio: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(start_time))}")
    
    if xml_to_excel_optimized(xml_file, excel_file):
        # Timestamp de fin
        end_time = time.time()
        elapsed_time = end_time - start_time
        
        print(f"Fin: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(end_time))}")
        print(f"Tiempo transcurrido: {elapsed_time:.2f} segundos")
        print("Conversión completada exitosamente")
    else:
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"Fin: {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(end_time))}")
        print(f"Tiempo transcurrido: {elapsed_time:.2f} segundos")
        print("Error en la conversión")

if __name__ == "__main__":
    main()