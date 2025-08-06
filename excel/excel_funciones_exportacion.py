
def xml_to_excel(config_file, output_file=None):
    """Convierte el XML a Excel usando configuración del archivo"""
    logger = None
    
    try:
        # Validar XML contra XSD antes de procesarlo
        if not validate_xml_against_xsd(config_file):
            return False
        
        tree = ET.parse(config_file)
        root = tree.getroot()
        
        # Configurar logging primero
        log_element = root.find('log')
        logger = setup_logging(log_element)
        logger.info("Iniciando conversión XML a Excel")
        
        # Log de la configuración de logging
        if log_element is not None:
            log_config = {}
            for child in log_element:
                log_config[child.tag] = child.text
            logger.info(f"Configuración de logging aplicada: {log_config}")
        else:
            logger.info("No se especificó configuración de logging, usando configuración por defecto")
        
        # Extraer configuración de datos
        data_element = root.find('data')
        if data_element is not None:
            data_in_element = data_element.find('dataIn')
            data_out_element = data_element.find('dataOut')
            
            if data_in_element is not None:
                logger.info(f"dataIn especificado: {data_in_element.text}")
                xml_file = validate_and_get_data_source(data_in_element.text, logger, is_data_in=True)
                if xml_file is None:
                    return False
                logger.info(f"Archivo XML de datos procesado: {xml_file}")
            else:
                # Si no hay dataIn, usar el archivo de configuración como datos
                xml_file = config_file
                logger.info("No se especificó dataIn, usando archivo de configuración como datos")
                
            if data_out_element is not None:
                excel_file = validate_and_get_data_source(data_out_element.text, logger, is_data_in=False)
                if excel_file is None:
                    return False
                logger.info(f"Archivo Excel de salida: {excel_file}")
            elif output_file:
                excel_file = output_file
                logger.info(f"Usando archivo de salida del parámetro: {excel_file}")
            else:
                logger.error("No se encontró dataOut en la configuración ni se especificó archivo de salida")
                return False
        else:
            # Fallback: usar el archivo de configuración como XML de datos
            xml_file = config_file
            excel_file = output_file if output_file else "salida.xlsx"
            if logger:
                logger.warning("No se encontró sección <data>, usando modo compatibilidad")
            else:
                print("Advertencia: No se encontró sección <data>, usando modo compatibilidad")
        
        # Si xml_file es diferente al config_file, cargar el archivo de datos
        if xml_file != config_file:
            if not os.path.exists(xml_file):
                print(f"Error: El archivo de datos {xml_file} no existe")
                return False
            data_tree = ET.parse(xml_file)
            data_root = data_tree.getroot()
        else:
            data_root = root
        
        # Buscar estilos y workbooks en el archivo de datos
        styles_element = data_root.find('styles')
        if styles_element is None:
            # Si no hay estilos en datos, buscar en workbooks
            workbooks_element = data_root.find('workbooks')
            if workbooks_element is not None:
                styles_element = workbooks_element.find('styles')
        
        styles_dict = parse_styles(styles_element) if styles_element is not None else {}
        # Verificar si el archivo Excel ya existe
        if os.path.isfile(excel_file):
            wb = load_workbook(excel_file)
            print(f"Cargando archivo Excel existente: {excel_file}")
        else:
            wb = Workbook()
            wb.remove(wb.active)
            print(f"Creando nuevo archivo Excel: {excel_file}")
        
        created_sheets = {}  # Diccionario para rastrear hojas creadas
        
        # Si se cargó un archivo existente, indexar las hojas ya presentes
        for existing_sheet in wb.worksheets:
            created_sheets[existing_sheet.title] = existing_sheet
        
        # Buscar workbooks en el lugar correcto
        workbook_elements = data_root.findall('workbook')
        if not workbook_elements:
            # Si no hay workbooks directos, buscar en sección workbooks
            workbooks_element = data_root.find('workbooks')
            if workbooks_element is not None:
                workbook_elements = workbooks_element.findall('workbook')
        
        for workbook_elem in workbook_elements:
            sheet_name = workbook_elem.get('name', 'Hoja1')
            
            # Contar las celdas en este workbook
            cell_elements = workbook_elem.findall('cell')
            cell_count = len(cell_elements)
            
            # Verificar si la hoja ya existe (en archivo cargado o ya procesada)
            if sheet_name in created_sheets:
                ws = created_sheets[sheet_name]
                print(f"  Utilizando hoja existente: {sheet_name}")
                if logger:
                    logger.info(f"Workbook '{sheet_name}': Utilizando hoja existente, procesando {cell_count} celdas")
            else:
                ws = wb.create_sheet(title=sheet_name)
                created_sheets[sheet_name] = ws
                print(f"  Creando nueva hoja: {sheet_name}")
                if logger:
                    logger.info(f"Workbook '{sheet_name}': Hoja creada, procesando {cell_count} celdas")
            
            for cell_elem in cell_elements:
                row = int(cell_elem.get('row'))
                column = cell_elem.get('column')
                value = cell_elem.get('value', '')
                format_attr = cell_elem.get('format', 'General')
                style_id = cell_elem.get('style')
                ws[f"{column}{row}"] = value
                cell = ws[f"{column}{row}"]
                
                # Aplicar formato de número si está especificado
                if format_attr and format_attr != 'General':
                    cell.number_format = format_attr
                
                # Aplicar estilos si están especificados
                if style_id and style_id in styles_dict:
                    style_data = styles_dict[style_id]
                    font, fill, alignment = create_openpyxl_style(style_data)
                    cell.font = font
                    if fill:
                        cell.fill = fill
                    if alignment:
                        cell.alignment = alignment
        for sheet in wb.worksheets:
            for column in sheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                sheet.column_dimensions[column_letter].width = adjusted_width
        wb.save(excel_file)
        if logger:
            logger.info(f"Archivo Excel creado exitosamente: {excel_file}")
        else:
            print(f"Archivo Excel creado exitosamente: {excel_file}")
        return True
        
    except ET.ParseError as e:
        error_msg = f"Error parsing XML: {e}"
        if logger:
            logger.error(error_msg)
        else:
            print(error_msg)
        return False
    except Exception as e:
        error_msg = f"Error creando Excel: {e}"
        if logger:
            logger.error(error_msg)
        else:
            print(error_msg)
        return False