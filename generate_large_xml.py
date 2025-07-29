#!/usr/bin/env python3
"""
Script para generar un XML con 10000 filas de datos inventados
"""

import random

def generate_large_xml():
    """Genera un XML con 10000 filas de datos"""
    
    # Datos para generar contenido aleatorio
    nombres = ["Ana", "Juan", "María", "Carlos", "Laura", "Pedro", "Carmen", "Miguel", "Isabel", "José", 
               "Lucía", "Antonio", "Elena", "Francisco", "Patricia", "Manuel", "Rosa", "David", "Cristina", "Javier"]
    
    apellidos = ["García", "Rodríguez", "González", "Fernández", "López", "Martínez", "Sánchez", "Pérez", 
                 "Gómez", "Martín", "Jiménez", "Ruiz", "Hernández", "Díaz", "Moreno", "Muñoz", "Álvarez", 
                 "Romero", "Alonso", "Gutiérrez"]
    
    ciudades = ["Madrid", "Barcelona", "Valencia", "Sevilla", "Zaragoza", "Málaga", "Murcia", "Palma", 
                "Las Palmas", "Bilbao", "Alicante", "Córdoba", "Valladolid", "Vigo", "Gijón", "Hospitalet", 
                "Granada", "Vitoria", "Coruña", "Elche"]
    
    departamentos = ["Ventas", "Marketing", "IT", "RRHH", "Finanzas", "Operaciones", "Logística", 
                     "Atención al Cliente", "Desarrollo", "Calidad"]
    
    # Crear el XML
    xml_content = '''<?xml version="1.0" encoding="UTF-8"?>
<workbooks>
    <styles>
        <style id="1">
            <font>Arial</font>
            <size>12</size>
            <bold>true</bold>
            <color>#000000</color>
            <background>#E6E6FA</background>
        </style>
        <style id="2">
            <font>Arial</font>
            <size>10</size>
            <bold>false</bold>
            <color>#333333</color>
        </style>
        <style id="3">
            <font>Arial</font>
            <size>10</size>
            <bold>false</bold>
            <color>#000000</color>
            <alignment>center</alignment>
        </style>
        <style id="4">
            <font>Arial</font>
            <size>10</size>
            <bold>false</bold>
            <color>#006400</color>
            <alignment>right</alignment>
        </style>
        <style id="5">
            <font>Arial</font>
            <size>10</size>
            <bold>false</bold>
            <color>#8B4513</color>
        </style>
    </styles>
    
    <columns>
        <column name="A" width="-1"/>
        <column name="B" width="8"/>
        <column name="C" width="-1"/>
        <column name="D" width="-1"/>
        <column name="E" width="12"/>
    </columns>
    
    <workbook name="Empleados">
        <cell row="1" column="A" text="Nombre Completo" style="1"/>
        <cell row="1" column="B" text="Edad" style="1"/>
        <cell row="1" column="C" text="Ciudad" style="1"/>
        <cell row="1" column="D" text="Departamento" style="1"/>
        <cell row="1" column="E" text="Salario" style="1"/>
'''
    
    print("Generando 100000 filas de datos...")
    
    # Generar 100000 filas de datos
    for i in range(2, 100002):  # Filas 2 a 100001 (100000 filas de datos)
        nombre_completo = f"{random.choice(nombres)} {random.choice(apellidos)}"
        edad = random.randint(22, 65)
        ciudad = random.choice(ciudades)
        departamento = random.choice(departamentos)
        salario = random.randint(25000, 80000)
        
        xml_content += f'''        <cell row="{i}" column="A" text="{nombre_completo}" style="2"/>
        <cell row="{i}" column="B" text="{edad}" style="3"/>
        <cell row="{i}" column="C" text="{ciudad}" style="5"/>
        <cell row="{i}" column="D" text="{departamento}" style="2"/>
        <cell row="{i}" column="E" text="{salario}" style="4"/>
'''
        
        if i % 10000 == 0:
            print(f"Generadas {i-1} filas...")
    
    xml_content += '''    </workbook>
</workbooks>'''
    
    # Escribir el archivo
    with open('ineo_xlsx.xml', 'w', encoding='utf-8') as f:
        f.write(xml_content)
    
    print("XML generado exitosamente: ineo_xlsx.xml")
    print(f"Total de filas: 100001 (1 cabecera + 100000 datos)")

if __name__ == "__main__":
    generate_large_xml()