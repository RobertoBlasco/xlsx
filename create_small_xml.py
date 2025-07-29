#!/usr/bin/env python3
"""
Script para crear ineo_xlsx_small.xml con solo 10 registros de datos
basado en ineo_xlsx.xml
"""

import xml.etree.ElementTree as ET

def create_small_xml():
    """Crea un XML pequeño con solo 10 registros"""
    
    try:
        # Parsear el XML original
        tree = ET.parse('ineo_xlsx.xml')
        root = tree.getroot()
        
        # Crear nuevo root
        new_root = ET.Element('workbooks')
        
        # Copiar la sección de estilos
        styles_elem = root.find('styles')
        if styles_elem is not None:
            new_root.append(styles_elem)
        
        # Procesar workbooks
        for workbook_elem in root.findall('workbook'):
            workbook_name = workbook_elem.get('name')
            new_workbook = ET.SubElement(new_root, 'workbook', name=workbook_name)
            
            # Obtener todas las celdas
            cells = workbook_elem.findall('cell')
            
            # Filtrar para obtener cabecera + 10 registros de datos
            selected_cells = []
            
            # Primero agregar la cabecera (fila 1)
            for cell in cells:
                if cell.get('row') == '1':
                    selected_cells.append(cell)
            
            # Luego agregar las primeras 10 filas de datos (filas 2-11)
            for row_num in range(2, 12):  # filas 2 a 11
                for cell in cells:
                    if cell.get('row') == str(row_num):
                        selected_cells.append(cell)
            
            # Agregar las celdas seleccionadas al nuevo workbook
            for cell in selected_cells:
                new_workbook.append(cell)
        
        # Crear el árbol y guardar
        new_tree = ET.ElementTree(new_root)
        
        # Formatear y escribir el XML
        ET.indent(new_tree, space="    ")
        new_tree.write('ineo_xlsx_small.xml', encoding='utf-8', xml_declaration=True)
        
        print("Archivo ineo_xlsx_small.xml creado exitosamente")
        print("Contiene: 1 fila de cabecera + 10 registros de datos")
        
    except Exception as e:
        print(f"Error creando XML pequeño: {e}")

if __name__ == "__main__":
    create_small_xml()