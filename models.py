from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()

class Usuario(UserMixin, db.Model):
    __tablename__ = 'usuarios'
    id              = db.Column(db.Integer, primary_key=True)
    nombre          = db.Column(db.String(100), nullable=False)
    username        = db.Column(db.String(50), unique=True, nullable=False)
    password        = db.Column(db.String(200), nullable=False)
    rol             = db.Column(db.String(20), nullable=False)  # admin | empleado | ciudadano
    dni             = db.Column(db.String(8), unique=True, nullable=True)
    email           = db.Column(db.String(120), unique=True, nullable=True)
    telefono        = db.Column(db.String(15), nullable=True)
    activo          = db.Column(db.Boolean, default=True)
    fecha_registro  = db.Column(db.DateTime, default=datetime.utcnow)
    ultimo_acceso   = db.Column(db.DateTime, nullable=True)

class Tramite(db.Model):
    __tablename__ = 'tramites'
    id              = db.Column(db.Integer, primary_key=True)
    ciudadano_id    = db.Column(db.Integer, db.ForeignKey('usuarios.id'), nullable=False)
    empleado_id     = db.Column(db.Integer, db.ForeignKey('usuarios.id'), nullable=True)
    tipo            = db.Column(db.String(100), nullable=False)
    descripcion     = db.Column(db.Text, nullable=False)
    urgencia        = db.Column(db.String(20), nullable=False)
    prioridad_ml    = db.Column(db.String(20), default='pendiente')
    estado          = db.Column(db.String(20), default='pendiente')
    fecha_registro  = db.Column(db.DateTime, default=datetime.utcnow)
    fecha_actualizacion = db.Column(db.DateTime, nullable=True, onupdate=datetime.utcnow)
    dias_resolucion = db.Column(db.Integer, default=0)
    observaciones   = db.Column(db.Text, nullable=True)
    confianza_ml    = db.Column(db.Float, default=0.0)
    explicacion_ml  = db.Column(db.Text, nullable=True)
    ciudadano       = db.relationship('Usuario', foreign_keys=[ciudadano_id], backref='tramites')
    empleado        = db.relationship('Usuario', foreign_keys=[empleado_id], backref='tramites_asignados')

class Documento(db.Model):
    __tablename__ = 'documentos'
    id              = db.Column(db.Integer, primary_key=True)
    tramite_id      = db.Column(db.Integer, db.ForeignKey('tramites.id'), nullable=False)
    usuario_id      = db.Column(db.Integer, db.ForeignKey('usuarios.id'), nullable=False)
    nombre_original = db.Column(db.String(255), nullable=False)
    nombre_archivo  = db.Column(db.String(255), nullable=False)   # nombre en disco
    tipo_mime       = db.Column(db.String(100), nullable=True)
    tamanio         = db.Column(db.Integer, default=0)             # bytes
    fecha_subida    = db.Column(db.DateTime, default=datetime.utcnow)
    tramite         = db.relationship('Tramite', backref='documentos')
    usuario         = db.relationship('Usuario', backref='documentos')

class Postulante(db.Model):
    __tablename__ = 'postulantes'
    id              = db.Column(db.Integer, primary_key=True)
    usuario_id      = db.Column(db.Integer, db.ForeignKey('usuarios.id'), nullable=True)
    nombre          = db.Column(db.String(100), nullable=False)
    dni             = db.Column(db.String(8), nullable=True)
    email           = db.Column(db.String(120), nullable=True)
    puesto          = db.Column(db.String(100), nullable=False)
    experiencia     = db.Column(db.Integer, nullable=False)
    educacion       = db.Column(db.String(50), nullable=False)
    habilidades     = db.Column(db.Integer, nullable=False)
    idiomas         = db.Column(db.String(100), nullable=True)
    puntaje_ml      = db.Column(db.Float, default=0.0)
    resultado_ml    = db.Column(db.String(20), default='pendiente')
    probabilidad    = db.Column(db.Float, default=0.0)
    fecha_registro  = db.Column(db.DateTime, default=datetime.utcnow)
    archivo_cv      = db.Column(db.String(255), nullable=True)   # nombre en disco
    texto_cv        = db.Column(db.Text, nullable=True)         # texto extraído del CV
    usuario         = db.relationship('Usuario', backref='postulaciones')

class Alerta(db.Model):
    __tablename__ = 'alertas'
    id              = db.Column(db.Integer, primary_key=True)
    tramite_id      = db.Column(db.Integer, db.ForeignKey('tramites.id'), nullable=False)
    ciudadano_id    = db.Column(db.Integer, db.ForeignKey('usuarios.id'), nullable=False)
    mensaje         = db.Column(db.Text, nullable=False)
    tipo            = db.Column(db.String(20), default='info')     # info | warning | danger | success
    leido           = db.Column(db.Boolean, default=False)
    enviado_email   = db.Column(db.Boolean, default=False)
    fecha           = db.Column(db.DateTime, default=datetime.utcnow)
    tramite         = db.relationship('Tramite', backref='alertas')
    ciudadano       = db.relationship('Usuario', backref='alertas')

class LogActividad(db.Model):
    __tablename__ = 'log_actividad'
    id              = db.Column(db.Integer, primary_key=True)
    usuario_id      = db.Column(db.Integer, db.ForeignKey('usuarios.id'), nullable=False)
    accion          = db.Column(db.String(200), nullable=False)
    entidad         = db.Column(db.String(50), nullable=True)
    entidad_id      = db.Column(db.Integer, nullable=True)
    ip              = db.Column(db.String(45), nullable=True)
    fecha           = db.Column(db.DateTime, default=datetime.utcnow)
    usuario         = db.relationship('Usuario', backref='actividad')
    
    
class Auditoria(db.Model):
    __tablename__ = 'auditoria'

    id = db.Column(db.Integer, primary_key=True)
    usuario = db.Column(db.String(120), nullable=False)
    rol = db.Column(db.String(50), nullable=False)
    accion = db.Column(db.String(255), nullable=False)
    modulo = db.Column(db.String(100), nullable=False)
    fecha = db.Column(db.DateTime, default=datetime.utcnow)
    