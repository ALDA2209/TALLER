from flask import Flask, render_template, redirect, url_for, request, flash, session
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from flask_mail import Mail, Message
from werkzeug.security import generate_password_hash, check_password_hash
from models import db, Usuario, Tramite, Postulante, Alerta
from ml import clasificar_tramite, evaluar_cv
from datetime import datetime
from functools import wraps
import random

app = Flask(__name__)
app.config['SECRET_KEY'] = 'gestimuni2024'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///gestimuni.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'pasivi22@gmail.com'
app.config['MAIL_PASSWORD'] = 'lvvfijzzlrbqvjwc'
app.config['MAIL_DEFAULT_SENDER'] = 'pasivi22@gmail.com'

db.init_app(app)
mail = Mail(app)
login_manager = LoginManager(app)
login_manager.login_view = 'login'

@login_manager.user_loader
def load_user(user_id):
    return Usuario.query.get(int(user_id))

def rol_requerido(*roles):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            if current_user.rol not in roles:
                flash('No tienes permiso para acceder.', 'danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return decorated
    return decorator

# ── Inicio público ───────────────────────────────────────────────────
@app.route('/')
def index():
    return redirect(url_for('inicio'))

@app.route('/inicio')
def inicio():
    return render_template('inicio.html')

# ── Login ciudadano ──────────────────────────────────────────────────
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = Usuario.query.filter_by(username=username).first()
        if user and check_password_hash(user.password, password):
            if user.rol == 'ciudadano':
                login_user(user)
                return redirect(url_for('dashboard'))
            else:
                flash('Usa el acceso de administrador.', 'danger')
        else:
            flash('Usuario o contraseña incorrectos.', 'danger')
    return render_template('login.html')

# ── Registro ciudadano ───────────────────────────────────────────────
@app.route('/registro', methods=['GET', 'POST'])
def registro():
    if request.method == 'POST':
        nombre = request.form['nombre']
        username = request.form['username']
        password = request.form['password']
        if Usuario.query.filter_by(username=username).first():
            flash('El usuario ya existe.', 'danger')
            return redirect(url_for('registro'))
        usuario = Usuario(
            nombre=nombre,
            username=username,
            password=generate_password_hash(password),
            rol='ciudadano'
        )
        db.session.add(usuario)
        db.session.commit()
        flash('Cuenta creada correctamente. Inicia sesión.', 'success')
        return redirect(url_for('login'))
    return render_template('registro.html')

# ── Login admin con PIN por correo ───────────────────────────────────
@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        paso = request.form.get('paso')
        if paso == '1':
            username = request.form['username']
            password = request.form['password']
            user = Usuario.query.filter_by(username=username).first()
            if user and check_password_hash(user.password, password):
                if user.rol == 'admin':
                    pin = str(random.randint(100000, 999999))
                    session['admin_pin'] = pin
                    session['admin_username'] = username
                    try:
                        msg = Message(
                            subject='GestiMuni Huánuco — Código de verificación',
                            recipients=['pasivi22@gmail.com'],
                            html=f'''
                            <div style="font-family:Arial,sans-serif; max-width:400px; margin:0 auto; padding:20px; border:1px solid #dee2e6; border-radius:8px;">
                                <div style="background:#2E7D4F; padding:15px; border-radius:6px 6px 0 0; text-align:center;">
                                    <h2 style="color:white; margin:0;">GestiMuni Huánuco</h2>
                                </div>
                                <div style="padding:20px; text-align:center;">
                                    <p style="color:#555;">Tu código de verificación para acceder al panel administrativo es:</p>
                                    <div style="font-size:2.5rem; font-weight:700; letter-spacing:10px; color:#1B5E35; padding:15px; background:#E8F5EE; border-radius:6px;">
                                        {pin}
                                    </div>
                                    <p style="color:#888; font-size:12px; margin-top:15px;">Este código expira en 5 minutos. No lo compartas con nadie.</p>
                                </div>
                            </div>
                            '''
                        )
                        mail.send(msg)
                        flash('Código enviado a tu correo.', 'success')
                    except Exception as e:
                        flash('Error al enviar correo. Intenta de nuevo.', 'danger')
                        return render_template('admin_login.html', paso=1)
                    return render_template('admin_login.html', paso=2)
                else:
                    flash('No tienes permiso de administrador.', 'danger')
            else:
                flash('Usuario o contraseña incorrectos.', 'danger')
        elif paso == '2':
            pin_ingresado = request.form['pin']
            pin_correcto = session.get('admin_pin')
            username = session.get('admin_username')
            if pin_ingresado == pin_correcto:
                user = Usuario.query.filter_by(username=username).first()
                login_user(user)
                session.pop('admin_pin', None)
                session.pop('admin_username', None)
                return redirect(url_for('dashboard'))
            else:
                flash('Código incorrecto. Intenta de nuevo.', 'danger')
                return render_template('admin_login.html', paso=2)
    return render_template('admin_login.html', paso=1)

# ── Logout ───────────────────────────────────────────────────────────
@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('inicio'))

# ── Dashboard ────────────────────────────────────────────────────────
@app.route('/dashboard')
@login_required
def dashboard():
    if current_user.rol == 'admin':
        total_tramites = Tramite.query.count()
        criticos = Tramite.query.filter_by(prioridad_ml='critico').count()
        postulantes = Postulante.query.count()
        alertas = Alerta.query.filter_by(leido=False).count()
        tramites = Tramite.query.order_by(Tramite.fecha_registro.desc()).limit(5).all()
        return render_template('dashboard.html',
            total_tramites=total_tramites, criticos=criticos,
            postulantes=postulantes, alertas=alertas, tramites=tramites)
    else:
        mis_tramites = Tramite.query.filter_by(ciudadano_id=current_user.id).all()
        mis_alertas = Alerta.query.filter_by(ciudadano_id=current_user.id, leido=False).all()
        return render_template('ciudadano.html',
            tramites=mis_tramites, alertas=mis_alertas)

# ── Trámites ─────────────────────────────────────────────────────────
@app.route('/tramites')
@login_required
def tramites():
    if current_user.rol == 'admin':
        lista = Tramite.query.order_by(Tramite.fecha_registro.desc()).all()
    else:
        lista = Tramite.query.filter_by(ciudadano_id=current_user.id).all()
    return render_template('tramites.html', tramites=lista)

@app.route('/tramites/nuevo', methods=['GET', 'POST'])
@login_required
def nuevo_tramite():
    if request.method == 'POST':
        tipo = request.form['tipo']
        descripcion = request.form['descripcion']
        urgencias_map = {
            'denuncia_peligro': 'alta',
            'licencia_construccion': 'media',
            'permiso_negocio': 'media',
            'licencia_funcionamiento': 'media',
            'certificado_residencia': 'baja',
            'partida_nacimiento': 'baja'
        }
        urgencia = urgencias_map.get(tipo, 'media')
        prioridad = clasificar_tramite(tipo, urgencia)
        tramite = Tramite(
            ciudadano_id=current_user.id,
            tipo=tipo, descripcion=descripcion,
            urgencia=urgencia, prioridad_ml=prioridad
        )
        db.session.add(tramite)
        db.session.flush()
        alerta = Alerta(
            tramite_id=tramite.id,
            ciudadano_id=current_user.id,
            mensaje=f'Tu trámite "{tipo}" fue registrado con prioridad {prioridad.upper()}.'
        )
        db.session.add(alerta)
        db.session.commit()
        flash(f'Trámite registrado. Prioridad asignada: {prioridad.upper()}', 'success')
        return redirect(url_for('tramites'))
    return render_template('nuevo_tramite.html')

@app.route('/tramites/estado/<int:id>', methods=['POST'])
@login_required
@rol_requerido('admin')
def cambiar_estado(id):
    tramite = Tramite.query.get_or_404(id)
    tramite.estado = request.form['estado']
    alerta = Alerta(
        tramite_id=tramite.id,
        ciudadano_id=tramite.ciudadano_id,
        mensaje=f'Tu trámite "{tramite.tipo}" cambió a estado: {tramite.estado.upper()}.'
    )
    db.session.add(alerta)
    db.session.commit()
    flash('Estado actualizado.', 'success')
    return redirect(url_for('tramites'))

# ── Currículos ───────────────────────────────────────────────────────
@app.route('/curriculos')
@login_required
@rol_requerido('admin')
def curriculos():
    lista = Postulante.query.order_by(Postulante.puntaje_ml.desc()).all()
    return render_template('curriculos.html', postulantes=lista)

@app.route('/curriculos/nuevo', methods=['GET', 'POST'])
@login_required
@rol_requerido('admin')
def nuevo_curriculo():
    if request.method == 'POST':
        nombre = request.form['nombre']
        puesto = request.form['puesto']
        experiencia = int(request.form['experiencia'])
        educacion = request.form['educacion']
        habilidades = int(request.form['habilidades'])
        resultado, puntaje = evaluar_cv(experiencia, educacion, habilidades)
        postulante = Postulante(
            nombre=nombre, puesto=puesto,
            experiencia=experiencia, educacion=educacion,
            habilidades=habilidades, resultado_ml=resultado,
            puntaje_ml=puntaje
        )
        db.session.add(postulante)
        db.session.commit()
        flash(f'CV evaluado: {resultado.upper()} con puntaje {puntaje}%', 'success')
        return redirect(url_for('curriculos'))
    return render_template('nuevo_curriculo.html')

# ── Alertas ──────────────────────────────────────────────────────────
@app.route('/alertas')
@login_required
def alertas():
    if current_user.rol == 'admin':
        lista = Alerta.query.order_by(Alerta.fecha.desc()).all()
    else:
        lista = Alerta.query.filter_by(ciudadano_id=current_user.id).all()
    return render_template('alertas.html', alertas=lista)

# ── Usuarios (solo admin) ────────────────────────────────────────────
@app.route('/usuarios')
@login_required
@rol_requerido('admin')
def usuarios():
    lista = Usuario.query.filter_by(rol='ciudadano').all()
    return render_template('usuarios.html', usuarios=lista)

@app.route('/usuarios/nuevo', methods=['GET', 'POST'])
@login_required
@rol_requerido('admin')
def nuevo_usuario():
    if request.method == 'POST':
        usuario = Usuario(
            nombre=request.form['nombre'],
            username=request.form['username'],
            password=generate_password_hash(request.form['password']),
            rol='ciudadano'
        )
        db.session.add(usuario)
        db.session.commit()
        flash('Usuario creado correctamente.', 'success')
        return redirect(url_for('usuarios'))
    return render_template('nuevo_usuario.html')

# ── Inicializar BD ───────────────────────────────────────────────────
def inicializar_bd():
    with app.app_context():
        db.create_all()
        if not Usuario.query.first():
            usuarios_default = [
                Usuario(nombre='Administrador', username='admin',
                    password=generate_password_hash('admin123'), rol='admin'),
                Usuario(nombre='Juan Pérez', username='ciudadano1',
                    password=generate_password_hash('1234'), rol='ciudadano'),
            ]
            db.session.add_all(usuarios_default)
            db.session.commit()
            print('Base de datos inicializada con usuarios por defecto.')

if __name__ == '__main__':
    inicializar_bd()
    app.run(debug=True)