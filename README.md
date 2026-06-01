# GestiMuni Huánuco

Sistema automatizado de gestión municipal con Machine Learning para la Municipalidad Provincial de Huánuco.

## Descripción
GestiMuni es una web app desarrollada con Flask y scikit-learn que automatiza la gestión de trámites municipales y selección de currículos mediante algoritmos de Machine Learning.

## Características
- Clasificación automática de trámites con ML (Crítico / Normal / Bajo)
- Evaluación de currículos con IA (Apto / No apto)
- Sistema de alertas en tiempo real
- Panel de administración con 3 roles (Admin, Empleado, Ciudadano)
- Autenticación con PIN de verificación para administradores

## Tecnologías
- Python 3.12
- Flask + Flask-Login + SQLAlchemy
- scikit-learn + pandas + numpy
- SQLite
- Bootstrap 5

## Instalación
```bash
git clone https://github.com/ALDA2209/TALLER.git
cd TALLER
python3 -m venv venv
source venv/bin/activate
pip install flask flask-login flask-sqlalchemy scikit-learn pandas numpy
python app.py
```

## Usuarios de prueba
- Admin: `admin` / `admin123` → http://localhost:5000/admin/login
- Empleado: `empleado1` / `user123`
- Ciudadano: `ciudadano1` / `1234`

## Curso
Taller de Desarrollo de Aplicaciones con Machine Learning — SENATI 2026
