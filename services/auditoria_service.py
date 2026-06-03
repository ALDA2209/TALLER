from datetime import datetime
from flask_login import current_user
from models import db, Auditoria


def registrar_auditoria(accion, modulo):
    """
    Registra acciones importantes del sistema.
    Ejemplo:
    - Registro de trámite
    - Cambio de estado
    - Evaluación de currículo
    - Creación de usuario
    """

    try:
        if current_user and current_user.is_authenticated:
            usuario = current_user.nombre
            rol = current_user.rol
        else:
            usuario = "Sistema"
            rol = "sistema"

        registro = Auditoria(
            usuario=usuario,
            rol=rol,
            accion=accion,
            modulo=modulo,
            fecha=datetime.utcnow()
        )

        db.session.add(registro)
        db.session.commit()

        return True

    except Exception as e:
        db.session.rollback()
        print(f"Error registrando auditoría: {e}")
        return False