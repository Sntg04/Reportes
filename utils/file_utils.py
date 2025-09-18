import logging
import os
from flask import current_app

def setup_logging():
    """Configura logging para la aplicación Flask."""
    log_level = os.getenv('LOG_LEVEL', 'INFO').upper()
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s %(levelname)s %(name)s %(message)s',
    )
    logging.getLogger('werkzeug').setLevel(logging.WARNING)


def allowed_file(filename, allowed_extensions):
    """Valida la extensión de un archivo subido."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions
