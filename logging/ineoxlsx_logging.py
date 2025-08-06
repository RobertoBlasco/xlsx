import logging
import os
from ineoXlsxGlobales import EXECUTION_TIMESTAMP

# Configurar logging dinámico según logOut
def setup_logging(log_out: str = None, identifier: str = None):
    """
    Configurar logging según el valor de logOut
    
    Args:
        log_out: Destino del log - None (auto FILE://), 'FILE://ruta.log', 'URL://endpoint', 'NOLOG'
        identifier: Identificador para usar en nombres de archivo (task_id o timestamp)
    """
    handlers = []
    
    if log_out == 'NOLOG':
        # No configurar ningún handler - sin logs
        pass
    elif log_out is None:
        # Por defecto: crear archivo con identificador (task_id o timestamp)
        if not identifier:
            identifier = ineoxlsxGlobales.EXECUTION_TIMESTAMP
        log_file_path = f"ineoDocxCmdLine_{identifier}.log"
        handlers.append(logging.FileHandler(log_file_path, mode='w'))
    
    elif log_out.startswith('FILE://'):
        # Logs a archivo específico
        log_file_path = log_out.replace('FILE://', '')
        
        # Crear directorio padre si no existe
        log_dir = os.path.dirname(log_file_path)
        if log_dir and not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        handlers.append(logging.FileHandler(log_file_path, mode='w'))
    
    elif log_out.startswith('URL://'):
        # Para URL:// se mantiene archivo por compatibilidad pero se podría implementar HTTPHandler
        # Por ahora usar archivo por defecto y notificar
        log_dir = os.path.join(os.path.dirname(__file__), 'logs')
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        log_file = os.path.join(log_dir, 'ineoDocx.log')
        handlers.append(logging.FileHandler(log_file, mode='w'))
        print(f"AVISO: URL:// para logs no implementado aún, usando archivo: {log_file}", file=sys.stderr)
    
    else:
        # Fallback: crear archivo con identificador
        if not identifier:
            identifier = EXECUTION_TIMESTAMP
        log_file_path = f"ineoDocxCmdLine_{identifier}.log"
        handlers.append(logging.FileHandler(log_file_path, mode='w'))
        print(f"AVISO: logOut '{log_out}' no reconocido, usando FILE://{log_file_path}", file=sys.stderr)
    
    # Configurar logging con nivel DEBUG para capturar todos los mensajes
    if handlers:
        # Solo configurar si hay handlers (no es NOLOG)
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=handlers,
            force=True  # Reconfigurar si ya estaba configurado
        )
    else:
        # NOLOG: configurar con nivel crítico para suprimir todos los mensajes
        logging.basicConfig(
            level=logging.CRITICAL + 1,  # Nivel más alto que CRITICAL para suprimir todo
            handlers=[],
            force=True
        )