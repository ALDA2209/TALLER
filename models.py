from flask_sqlalchemy import SQLAlchemy
from flask_login import UserMixin
from datetime import datetime

db = SQLAlchemy()

class Usuario(UserMixin, db.Model):
    __tablename__ = 'usuarios'
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    username = db.Column(db.String(50), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    rol = db.Column(db.String(20), nullable=False)
    fecha_registro = db.Column(db.DateTime, default=datetime.utcnow)

class Tramite(db.Model):
    __tablename__ = 'tramites'
    id = db.Column(db.Integer, primary_key=True)
    ciudadano_id = db.Column(db.Integer, db.ForeignKey('usuarios.id'), nullable=False)
    tipo = db.Column(db.String(100), nullable=False)
    descripcion = db.Column(db.Text, nullable=False)
    urgencia = db.Column(db.String(20), nullable=False)
    prioridad_ml = db.Column(db.String(20), default='pendiente')
    estado = db.Column(db.String(20), default='pendiente')
    fecha_registro = db.Column(db.DateTime, default=datetime.utcnow)
    dias_resolucion = db.Column(db.Integer, default=0)
    ciudadano = db.relationship('Usuario', backref='tramites')

class Postulante(db.Model):
    __tablename__ = 'postulantes'
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    puesto = db.Column(db.String(100), nullable=False)
    experiencia = db.Column(db.Integer, nullable=False)
    educacion = db.Column(db.String(50), nullable=False)
    habilidades = db.Column(db.Integer, nullable=False)
    puntaje_ml = db.Column(db.Float, default=0.0)
    resultado_ml = db.Column(db.String(20), default='pendiente')
    fecha_registro = db.Column(db.DateTime, default=datetime.utcnow)

class Alerta(db.Model):
    __tablename__ = 'alertas'
    id = db.Column(db.Integer, primary_key=True)
    tramite_id = db.Column(db.Integer, db.ForeignKey('tramites.id'), nullable=False)
    ciudadano_id = db.Column(db.Integer, db.ForeignKey('usuarios.id'), nullable=False)
    mensaje = db.Column(db.Text, nullable=False)
    leido = db.Column(db.Boolean, default=False)
    fecha = db.Column(db.DateTime, default=datetime.utcnow)
    tramite = db.relationship('Tramite', backref='alertas')
    ciudadano = db.relationship('Usuario', backref='alertas')