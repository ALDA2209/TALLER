from flask import Flask, render_template, redirect, url_for, request, flash
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from werkzeug.security import generate_password_hash, check_password_hash
from models import db, Usuario, Tramite, Postulante, Alerta
from ml import clasificar_tramite, evaluar_cv
from datetime import datetime
from functools import wraps

app = Flask(__name__)
app.config['SECRET_KEY'] = 'gestimuni2024'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///gestimuni.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db.init_app(app)
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
                flash('Usa el acceso de personal municipal.', 'danger')
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

# ── Login admin/empleado ─────────────────────────────────────────────
@app.route('/admin/login', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = Usuario.query.filter_by(username=username).first()
        if user and check_password_hash(user.password, password):
            if user.rol in ['admin', 'empleado']:
                login_user(user)
                return redirect(url_for('dashboard'))
            else:
                flash('No tienes permiso para acceder por esta vía.', 'danger')
        else:
            flash('Usuario o contraseña incorrectos.', 'danger')
    return render_template('admin_login.html')

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
    if current_user.rol in ['admin', 'empleado']:
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
    if current_user.rol in ['admin', 'empleado']:
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
        urgencia = request.form['urgencia']
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
        flash(f'Trámite registrado con prioridad: {prioridad.upper()}', 'success')
        return redirect(url_for('tramites'))
    return render_template('nuevo_tramite.html')

@app.route('/tramites/estado/<int:id>', methods=['POST'])
@login_required
@rol_requerido('admin', 'empleado')
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
@rol_requerido('admin', 'empleado')
def curriculos():
    lista = Postulante.query.order_by(Postulante.puntaje_ml.desc()).all()
    return render_template('curriculos.html', postulantes=lista)

@app.route('/curriculos/nuevo', methods=['GET', 'POST'])
@login_required
@rol_requerido('admin', 'empleado')
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
    if current_user.rol in ['admin', 'empleado']:
        lista = Alerta.query.order_by(Alerta.fecha.desc()).all()
    else:
        lista = Alerta.query.filter_by(ciudadano_id=current_user.id).all()
    return render_template('alertas.html', alertas=lista)

# ── Usuarios (solo admin) ────────────────────────────────────────────
@app.route('/usuarios')
@login_required
@rol_requerido('admin')
def usuarios():
    lista = Usuario.query.all()
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
            rol=request.form['rol']
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
                Usuario(nombre='Empleado Municipal', username='empleado1',
                    password=generate_password_hash('user123'), rol='empleado'),
                Usuario(nombre='Juan Pérez', username='ciudadano1',
                    password=generate_password_hash('1234'), rol='ciudadano'),
            ]
            db.session.add_all(usuarios_default)
            db.session.commit()
            print('Base de datos inicializada con usuarios por defecto.')

if __name__ == '__main__':
    inicializar_bd()
    app.run(debug=True)