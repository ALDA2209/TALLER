# GestiMuni Huánuco

Sistema automatizado de gestión municipal con Machine Learning para la Municipalidad Provincial de Huánuco.

## Descripción
GestiMuni es una web app desarrollada con Flask y scikit-learn que automatiza la gestión de trámites municipales y selección de currículos mediante algoritmos de Machine Learning.

## Características
- Página de inicio pública con trámites disponibles
- Clasificación automática de trámites con ML (Crítico / Normal / Bajo)
- Evaluación de currículos con IA (Apto / No apto)
- Sistema de alertas en tiempo real
- Login ciudadano con registro
- Panel administrativo con autenticación de 2 pasos (PIN por correo)
- Base de datos SQLite

## Tecnologías
- Python 3.12
- Flask + Flask-Login + Flask-Mail + SQLAlchemy
- scikit-learn + pandas + numpy
- SQLite
- Bootstrap 5

## Instalación
```bash
git clone https://github.com/ALDA2209/TALLER.git
cd TALLER
python3 -m venv venv
source venv/bin/activate
pip install flask flask-login flask-sqlalchemy flask-mail scikit-learn pandas numpy
python app.py
```

## Accesos
- Ciudadano: `http://localhost:5000/login`
- Admin: `http://localhost:5000/admin/login`

## Usuarios de prueba
- Admin: `admin` / `admin123`
- Ciudadano: `ciudadano1` / `1234`

## Curso
Taller de Desarrollo de Aplicaciones con Machine Learning — SENATI 2026
