"""
app.py — GestiMuni: Sistema de Gestión Municipal con ML
Roles: admin | empleado | ciudadano
Incluye: subida de docs, exportación Excel, alertas por correo al ciudadano,
         búsqueda por DNI, CRUD completo de usuarios, dashboard ML
"""

from flask import (Flask, render_template, redirect, url_for, request,
                   flash, session, jsonify, send_file, abort)
from flask_login import (LoginManager, login_user, logout_user,
                         login_required, current_user)
from flask_mail import Mail, Message
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from models import db, Usuario, Tramite, Postulante, Alerta, Documento, LogActividad
from ml import (clasificar_tramite, evaluar_cv,
                obtener_stats_tramites, obtener_stats_cvs, reentrenar)
from datetime import datetime
from functools import wraps
import random, os, io, uuid

# ── ExcelJS equivalente en Python: openpyxl ──────────────────────────
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment,
                              Border, Side, GradientFill)
from openpyxl.utils import get_column_letter
from services.auditoria_service import registrar_auditoria
from models import db, Usuario, Tramite, Postulante, Alerta, Auditoria
# ── dotenv ───────────────────────────────────────────────────────────
from dotenv import load_dotenv
load_dotenv()

# ════════════════════════════════════════════════════════════════════
app = Flask(__name__)
app.config['SECRET_KEY']                  = os.getenv('SECRET_KEY', 'gestimuni_dev_2024')
app.config['SQLALCHEMY_DATABASE_URI']     = os.getenv('SQLALCHEMY_DATABASE_URI',
                                                       'sqlite:///gestimuni.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
app.config['MAIL_SERVER']                 = os.getenv('MAIL_SERVER', 'smtp.gmail.com')
app.config['MAIL_PORT']                   = int(os.getenv('MAIL_PORT', 587))
app.config['MAIL_USE_TLS']                = os.getenv('MAIL_USE_TLS', 'True') == 'True'
app.config['MAIL_USERNAME']               = os.getenv('MAIL_USERNAME', '')
app.config['MAIL_PASSWORD']               = os.getenv('MAIL_PASSWORD', '')
app.config['MAIL_DEFAULT_SENDER']         = os.getenv('MAIL_USERNAME', '')
app.config['UPLOAD_FOLDER']               = os.path.join(os.path.dirname(__file__), 'uploads')
app.config['MAX_CONTENT_LENGTH']          = 16 * 1024 * 1024   # 16 MB

EXTENSIONES_PERMITIDAS = {'pdf', 'doc', 'docx', 'xls', 'xlsx',
                           'png', 'jpg', 'jpeg', 'gif', 'txt'}

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

db.init_app(app)
mail        = Mail(app)
login_mgr   = LoginManager(app)
login_mgr.login_view = 'login'

# ── Helpers ──────────────────────────────────────────────────────────
@login_mgr.user_loader
def load_user(uid):
    return Usuario.query.get(int(uid))


def extension_permitida(filename):
    return ('.' in filename and
            filename.rsplit('.', 1)[1].lower() in EXTENSIONES_PERMITIDAS)


def rol_requerido(*roles):
    def decorator(f):
        @wraps(f)
        def wrapped(*args, **kwargs):
            if current_user.rol not in roles:
                flash('No tienes permiso para acceder.', 'danger')
                return redirect(url_for('dashboard'))
            return f(*args, **kwargs)
        return wrapped
    return decorator


def registrar_log(accion, entidad=None, entidad_id=None):
    log = LogActividad(
        usuario_id=current_user.id,
        accion=accion,
        entidad=entidad,
        entidad_id=entidad_id,
        ip=request.remote_addr
    )
    db.session.add(log)


def enviar_correo_ciudadano(usuario, asunto, cuerpo_html):
    """Envía correo al email real del ciudadano si lo tiene registrado."""

    if not usuario.email:
        return False

    try:
        msg = Message(
            subject=f'GestiMuni — {asunto}',
            sender=f"GestiMuni <{app.config['MAIL_USERNAME']}>",
            recipients=[usuario.email],
            html=cuerpo_html
        )

        mail.send(msg)
        print(f"Correo enviado correctamente a {usuario.email}")
        return True

    except Exception as e:
        app.logger.error(f'Error enviando correo a {usuario.email}: {e}')
        print(f"Error enviando correo a {usuario.email}: {e}")
        return False

def _html_alerta_tramite(ciudadano_nombre, tramite_tipo, tramite_estado,
                          tramite_id, mensaje_extra=''):
    colores = {
        'aprobado':  ('#1D9E75', '#E1F5EE'),
        'rechazado': ('#A32D2D', '#FCEBEB'),
        'en_proceso':('#185FA5', '#E6F1FB'),
        'pendiente': ('#854F0B', '#FAEEDA'),
    }
    color_btn, color_bg = colores.get(tramite_estado, ('#444', '#f5f5f5'))
    return f'''
    <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;
                border:1px solid #dee2e6;border-radius:10px;overflow:hidden;">
      <div style="background:#2E7D4F;padding:18px 24px;text-align:center;">
        <h2 style="color:#fff;margin:0;font-size:20px;">🏛 GestiMuni Huánuco</h2>
        <p style="color:#a8d5ba;margin:4px 0 0;font-size:13px;">
          Municipalidad Provincial de Yau</p>
      </div>
      <div style="padding:24px;">
        <p style="color:#333;font-size:15px;">Hola <strong>{ciudadano_nombre}</strong>,</p>
        <p style="color:#555;font-size:14px;">
          Tu trámite <strong>#{tramite_id}</strong>
          (<em>{tramite_tipo.replace('_',' ').title()}</em>)
          ha sido actualizado.</p>
        <div style="background:{color_bg};border-radius:8px;padding:14px 18px;
                    margin:16px 0;text-align:center;">
          <span style="color:{color_btn};font-size:18px;font-weight:700;
                        letter-spacing:1px;">
            {tramite_estado.upper().replace('_',' ')}
          </span>
        </div>
        {f'<p style="color:#555;font-size:13px;">{mensaje_extra}</p>' if mensaje_extra else ''}
        <p style="color:#888;font-size:12px;margin-top:20px;">
          Puedes ingresar al sistema para más detalles.<br>
          Este correo es generado automáticamente, no respondas a él.</p>
      </div>
      <div style="background:#f8f9fa;padding:12px 24px;text-align:center;">
        <p style="color:#aaa;font-size:11px;margin:0;">
          © 2024 GestiMuni — Sistema de Gestión Municipal con IA</p>
      </div>
    </div>'''


# ════════════════════════════════════════════════════════════════════
# RUTAS PÚBLICAS
# ════════════════════════════════════════════════════════════════════
@app.route('/')
def index():
    return redirect(url_for('inicio'))


@app.route('/inicio')
def inicio():
    return render_template('inicio.html')


# ── Registro ciudadano ───────────────────────────────────────────────
@app.route('/registro', methods=['GET', 'POST'])
def registro():
    if request.method == 'POST':
        dni      = request.form.get('dni', '').strip()
        nombre   = request.form['nombre'].strip()
        username = request.form['username'].strip()
        email    = request.form.get('email', '').strip() or None
        telefono = request.form.get('telefono', '').strip() or None
        password = request.form['password']

        if Usuario.query.filter_by(username=username).first():
            flash('El usuario ya existe.', 'danger')
            return redirect(url_for('registro'))
        if dni and Usuario.query.filter_by(dni=dni).first():
            flash('El DNI ya está registrado.', 'danger')
            return redirect(url_for('registro'))

        usuario = Usuario(nombre=nombre, username=username,
                          password=generate_password_hash(password),
                          rol='ciudadano', dni=dni or None,
                          email=email, telefono=telefono)
        db.session.add(usuario)
        db.session.commit()
        flash('Cuenta creada. Inicia sesión.', 'success')
        return redirect(url_for('login'))
    return render_template('registro.html')


# ── Login ciudadano ──────────────────────────────────────────────────
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        user = Usuario.query.filter_by(username=username).first()
        if user and check_password_hash(user.password, password) and user.activo:
            if user.rol in ('ciudadano', 'empleado'):
                login_user(user)
                user.ultimo_acceso = datetime.utcnow()
                db.session.commit()
                registrar_auditoria(
                    f'Inicio de sesión de {user.nombre}',
                    'Seguridad'
                )
                return redirect(url_for('dashboard'))
            else:
                flash('Usa el acceso de administrador.', 'danger')
        else:
            flash('Usuario o contraseña incorrectos.', 'danger')
    return render_template('login.html')


# ── Login admin con PIN ──────────────────────────────────────────────
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
                    session['admin_pin']      = pin
                    session['admin_username'] = username
                    try:
                        msg = Message(
                            subject='GestiMuni — Código de verificación',
                            recipients=[user.email or app.config['MAIL_USERNAME']],
                            html=f'''
                            <div style="font-family:Arial;max-width:400px;margin:auto;
                                        border:1px solid #dee2e6;border-radius:8px;">
                              <div style="background:#2E7D4F;padding:16px;
                                          text-align:center;border-radius:8px 8px 0 0;">
                                <h2 style="color:#fff;margin:0;">GestiMuni Huánuco</h2>
                              </div>
                              <div style="padding:24px;text-align:center;">
                                <p style="color:#555;">Código de verificación admin:</p>
                                <div style="font-size:2.5rem;font-weight:700;
                                            letter-spacing:10px;color:#1B5E35;
                                            padding:15px;background:#E8F5EE;
                                            border-radius:6px;">{pin}</div>
                                <p style="color:#888;font-size:12px;margin-top:12px;">
                                  Expira en 5 minutos.</p>
                              </div>
                            </div>''')
                        mail.send(msg)
                        flash('Código enviado a tu correo.', 'success')
                    except Exception as e:
                        flash(f'Error al enviar correo: {e}', 'danger')
                        return render_template('admin_login.html', paso=1)
                    return render_template('admin_login.html', paso=2)
                else:
                    flash('No tienes permiso de administrador.', 'danger')
            else:
                flash('Usuario o contraseña incorrectos.', 'danger')

        elif paso == '2':
            pin_ok   = session.get('admin_pin')
            username = session.get('admin_username')
            if request.form['pin'] == pin_ok:
                user = Usuario.query.filter_by(username=username).first()
                login_user(user)
                user.ultimo_acceso = datetime.utcnow()
                db.session.commit()
                session.pop('admin_pin', None)
                session.pop('admin_username', None)
                registrar_log('login_admin')
                db.session.commit()
                registrar_auditoria(
                    f'Inicio de sesión administrador de {user.nombre}',
                    'Seguridad'
                )
                return redirect(url_for('dashboard'))
            else:
                flash('Código incorrecto.', 'danger')
                return render_template('admin_login.html', paso=2)

    return render_template('admin_login.html', paso=1)


@app.route('/logout')
@login_required
def logout():
    registrar_auditoria(
        f'Cierre de sesión de {current_user.nombre}',
        'Seguridad'
    )
    registrar_log('logout')
    db.session.commit()
    logout_user()
    return redirect(url_for('inicio'))


# ════════════════════════════════════════════════════════════════════
# DASHBOARD
# ════════════════════════════════════════════════════════════════════
@app.route('/dashboard')
@login_required
def dashboard():
    if current_user.rol == 'admin':
        total   = Tramite.query.count()
        criticos= Tramite.query.filter_by(prioridad_ml='critico').count()
        postulantes = Postulante.query.count()
        alertas_n= Alerta.query.filter_by(leido=False).count()
        tramites = Tramite.query.order_by(Tramite.fecha_registro.desc()).limit(5).all()
        return render_template('dashboard.html',
            total_tramites=total, criticos=criticos,
            postulantes=postulantes, alertas=alertas_n, tramites=tramites)

    elif current_user.rol == 'empleado':
        mis   = Tramite.query.filter_by(empleado_id=current_user.id).all()
        pend  = Tramite.query.filter_by(estado='pendiente').count()
        return render_template('dashboard_empleado.html',
            tramites=mis, pendientes=pend)

    else:  # ciudadano
        mis_tramites = Tramite.query.filter_by(ciudadano_id=current_user.id).all()
        mis_alertas  = Alerta.query.filter_by(
            ciudadano_id=current_user.id, leido=False).all()
        return render_template('ciudadano.html',
            tramites=mis_tramites, alertas=mis_alertas)


# ════════════════════════════════════════════════════════════════════
# TRÁMITES
# ════════════════════════════════════════════════════════════════════
@app.route('/tramites')
@login_required
def tramites():
    if current_user.rol == 'admin':
        lista = Tramite.query.order_by(Tramite.fecha_registro.desc()).all()
    elif current_user.rol == 'empleado':
        lista = Tramite.query.order_by(Tramite.fecha_registro.desc()).all()
    else:
        lista = Tramite.query.filter_by(
            ciudadano_id=current_user.id).order_by(Tramite.fecha_registro.desc()).all()
    return render_template('tramites.html', tramites=lista)


@app.route('/tramites/nuevo', methods=['GET', 'POST'])
@login_required
def nuevo_tramite():
    if request.method == 'POST':
        tipo        = request.form['tipo']
        descripcion = request.form['descripcion']
        urgencias_map = {
            'denuncia_peligro':        'alta',
            'licencia_construccion':   'media',
            'licencia_funcionamiento': 'media',
            'permiso_negocio':         'media',
            'certificado_residencia':  'baja',
            'partida_nacimiento':      'baja'
        }
        urgencia  = urgencias_map.get(tipo, 'media')
        prioridad, confianza = clasificar_tramite(tipo, urgencia)

        tramite = Tramite(
            ciudadano_id=current_user.id,
            tipo=tipo, descripcion=descripcion,
            urgencia=urgencia, prioridad_ml=prioridad,
            confianza_ml=confianza
        )
        db.session.add(tramite)
        db.session.flush()

        # Alerta interna
        alerta = Alerta(
            tramite_id=tramite.id,
            ciudadano_id=current_user.id,
            mensaje=(f'Tu trámite "{tipo.replace("_"," ").title()}" fue registrado '
                     f'con prioridad {prioridad.upper()} '
                     f'(confianza ML: {round(confianza*100,1)}%).'),
            tipo='info'
        )
        db.session.add(alerta)

        # Subida de documentos adjuntos
        archivos = request.files.getlist('documentos')
        for archivo in archivos:
            if archivo and extension_permitida(archivo.filename):
                nombre_orig = secure_filename(archivo.filename)
                ext         = nombre_orig.rsplit('.', 1)[1].lower()
                nombre_disco= f"{uuid.uuid4().hex}.{ext}"
                ruta        = os.path.join(app.config['UPLOAD_FOLDER'], nombre_disco)
                archivo.save(ruta)
                tam = os.path.getsize(ruta)
                doc = Documento(
                    tramite_id=tramite.id,
                    usuario_id=current_user.id,
                    nombre_original=nombre_orig,
                    nombre_archivo=nombre_disco,
                    tipo_mime=archivo.content_type,
                    tamanio=tam
                )
                db.session.add(doc)

        registrar_log('nuevo_tramite', 'Tramite', tramite.id)
        db.session.commit()
        registrar_auditoria(
            f"Registró trámite #{tramite.id}: {tipo.replace('_', ' ').title()} con prioridad {prioridad.upper()}",
            'Trámites'
        )

        # Correo al ciudadano
        html_c = _html_alerta_tramite(
            current_user.nombre, tipo, 'pendiente', tramite.id,
            f'Prioridad asignada por IA: {prioridad.upper()}'
        )
        enviado = enviar_correo_ciudadano(
            current_user, 'Trámite registrado', html_c)
        if enviado:
            alerta.enviado_email = True
            db.session.commit()

        flash(f'Trámite registrado. Prioridad IA: {prioridad.upper()} '
              f'(confianza {round(confianza*100,1)}%)', 'success')
        return redirect(url_for('tramites'))
    return render_template('nuevo_tramite.html')


@app.route('/tramites/estado/<int:id>', methods=['POST'])
@login_required
@rol_requerido('admin', 'empleado')
def cambiar_estado(id):
    tramite = Tramite.query.get_or_404(id)
    nuevo_estado  = request.form['estado']
    observaciones = request.form.get('observaciones', '')
    tramite.estado        = nuevo_estado
    tramite.observaciones = observaciones
    tramite.fecha_actualizacion = datetime.utcnow()

    alerta = Alerta(
        tramite_id=tramite.id,
        ciudadano_id=tramite.ciudadano_id,
        mensaje=(f'Tu trámite "{tramite.tipo.replace("_"," ").title()}" '
                 f'cambió a: {nuevo_estado.upper().replace("_"," ")}.'
                 + (f' Nota: {observaciones}' if observaciones else '')),
        tipo='success' if nuevo_estado == 'aprobado' else
             'danger'  if nuevo_estado == 'rechazado' else 'info'
    )
    db.session.add(alerta)
    registrar_log('cambiar_estado', 'Tramite', id)
    db.session.commit()
    registrar_auditoria(
        f"Cambió el trámite #{tramite.id} al estado {nuevo_estado.upper().replace('_', ' ')}",
        'Trámites'
    )

    # Correo al ciudadano
    html_c = _html_alerta_tramite(
        tramite.ciudadano.nombre, tramite.tipo,
        nuevo_estado, tramite.id, observaciones
    )
    enviado = enviar_correo_ciudadano(tramite.ciudadano,
                                      f'Estado actualizado: {nuevo_estado}', html_c)
    if enviado:
        alerta.enviado_email = True
        db.session.commit()

    flash('Estado actualizado.', 'success')
    return redirect(url_for('tramites'))


@app.route('/tramites/<int:id>')
@login_required
def detalle_tramite(id):
    tramite = Tramite.query.get_or_404(id)
    if (current_user.rol == 'ciudadano' and
            tramite.ciudadano_id != current_user.id):
        abort(403)
    return render_template('detalle_tramite.html', tramite=tramite)


# ════════════════════════════════════════════════════════════════════
# DOCUMENTOS
# ════════════════════════════════════════════════════════════════════
@app.route('/documentos/ver/<int:doc_id>')
@login_required
def ver_documento(doc_id):
    doc   = Documento.query.get_or_404(doc_id)
    ruta  = os.path.join(app.config['UPLOAD_FOLDER'], doc.nombre_archivo)
    if not os.path.exists(ruta):
        abort(404)
    return send_file(ruta, mimetype=doc.tipo_mime,
                     download_name=doc.nombre_original)


@app.route('/documentos/descargar/<int:doc_id>')
@login_required
def descargar_documento(doc_id):
    doc  = Documento.query.get_or_404(doc_id)
    ruta = os.path.join(app.config['UPLOAD_FOLDER'], doc.nombre_archivo)
    if not os.path.exists(ruta):
        abort(404)
    return send_file(ruta, as_attachment=True,
                     download_name=doc.nombre_original)


@app.route('/documentos/eliminar/<int:doc_id>', methods=['POST'])
@login_required
@rol_requerido('admin', 'empleado')
def eliminar_documento(doc_id):
    doc  = Documento.query.get_or_404(doc_id)
    ruta = os.path.join(app.config['UPLOAD_FOLDER'], doc.nombre_archivo)
    if os.path.exists(ruta):
        os.remove(ruta)
    db.session.delete(doc)
    registrar_log('eliminar_doc', 'Documento', doc_id)
    db.session.commit()
    registrar_auditoria(
        f'Eliminó documento {doc.nombre_original}',
        'Documentos'
    )
    flash('Documento eliminado.', 'success')
    return redirect(request.referrer or url_for('tramites'))


# ════════════════════════════════════════════════════════════════════
# CURRÍCULOS / POSTULANTES
# ════════════════════════════════════════════════════════════════════
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
        nombre     = request.form['nombre']
        dni        = request.form.get('dni', '').strip() or None
        email      = request.form.get('email', '').strip() or None
        puesto     = request.form['puesto']
        experiencia= int(request.form['experiencia'])
        educacion  = request.form['educacion']
        habilidades= int(request.form['habilidades'])
        idiomas    = int(request.form.get('idiomas', 0))

        resultado, puntaje, prob = evaluar_cv(
            experiencia, educacion, habilidades, idiomas)

        postulante = Postulante(
            nombre=nombre, dni=dni, email=email, puesto=puesto,
            experiencia=experiencia, educacion=educacion,
            habilidades=habilidades, idiomas=str(idiomas),
            resultado_ml=resultado, puntaje_ml=puntaje,
            probabilidad=prob
        )
        db.session.add(postulante)
        registrar_log('nuevo_curriculo', 'Postulante')
        db.session.commit()
        registrar_auditoria(
            f'Evaluó CV de {nombre} para el puesto {puesto}: {resultado.upper()} con puntaje {puntaje}%',
            'Currículos'
        )
        flash(f'CV evaluado: {resultado.upper()} — Puntaje: {puntaje}%', 'success')
        return redirect(url_for('curriculos'))
    return render_template('nuevo_curriculo.html')


@app.route('/curriculos/eliminar/<int:id>', methods=['POST'])
@login_required
@rol_requerido('admin')
def eliminar_postulante(id):
    p = Postulante.query.get_or_404(id)
    db.session.delete(p)
    registrar_log('eliminar_postulante', 'Postulante', id)
    db.session.commit()
    registrar_auditoria(
        f'Eliminó postulante {p.nombre}',
        'Currículos'
    )
    flash('Postulante eliminado.', 'success')
    return redirect(url_for('curriculos'))


# ════════════════════════════════════════════════════════════════════
# USUARIOS — CRUD COMPLETO
# ════════════════════════════════════════════════════════════════════
@app.route('/usuarios')
@login_required
@rol_requerido('admin')
def usuarios():
    q    = request.args.get('q', '').strip()
    rol  = request.args.get('rol', '')
    query= Usuario.query
    if q:
        query = query.filter(
            db.or_(Usuario.nombre.ilike(f'%{q}%'),
                   Usuario.username.ilike(f'%{q}%'),
                   Usuario.dni.ilike(f'%{q}%')))
    if rol:
        query = query.filter_by(rol=rol)
    lista = query.order_by(Usuario.fecha_registro.desc()).all()
    return render_template('usuarios.html', usuarios=lista,
                           q=q, rol_filtro=rol)


@app.route('/usuarios/buscar-dni')
@login_required
@rol_requerido('admin', 'empleado')
def buscar_por_dni():
    dni = request.args.get('dni', '').strip()
    if not dni:
        return jsonify({'error': 'DNI vacío'}), 400
    user = Usuario.query.filter_by(dni=dni).first()
    if not user:
        return jsonify({'error': 'No encontrado'}), 404
    tramites = Tramite.query.filter_by(ciudadano_id=user.id).all()
    return jsonify({
        'id':       user.id,
        'nombre':   user.nombre,
        'dni':      user.dni,
        'email':    user.email,
        'telefono': user.telefono,
        'rol':      user.rol,
        'tramites': [{'id': t.id, 'tipo': t.tipo,
                      'estado': t.estado, 'prioridad': t.prioridad_ml}
                     for t in tramites]
    })


@app.route('/usuarios/nuevo', methods=['GET', 'POST'])
@login_required
@rol_requerido('admin')
def nuevo_usuario():
    if request.method == 'POST':
        dni      = request.form.get('dni', '').strip() or None
        nombre   = request.form['nombre'].strip()
        username = request.form['username'].strip()
        email    = request.form.get('email', '').strip() or None
        telefono = request.form.get('telefono', '').strip() or None
        password = request.form['password']
        rol      = request.form.get('rol', 'ciudadano')

        if Usuario.query.filter_by(username=username).first():
            flash('El usuario ya existe.', 'danger')
            return redirect(url_for('nuevo_usuario'))
        if dni and Usuario.query.filter_by(dni=dni).first():
            flash('El DNI ya está registrado.', 'danger')
            return redirect(url_for('nuevo_usuario'))

        usuario = Usuario(nombre=nombre, username=username,
                          password=generate_password_hash(password),
                          rol=rol, dni=dni, email=email, telefono=telefono)
        db.session.add(usuario)
        registrar_log('crear_usuario', 'Usuario')
        db.session.commit()
        registrar_auditoria(
            f'Creó usuario {nombre} con rol {rol}',
            'Usuarios'
        )
        flash('Usuario creado correctamente.', 'success')
        return redirect(url_for('usuarios'))
    return render_template('nuevo_usuario.html')


@app.route('/usuarios/editar/<int:id>', methods=['GET', 'POST'])
@login_required
@rol_requerido('admin')
def editar_usuario(id):
    usuario = Usuario.query.get_or_404(id)
    if request.method == 'POST':
        usuario.nombre   = request.form['nombre'].strip()
        usuario.email    = request.form.get('email', '').strip() or None
        usuario.telefono = request.form.get('telefono', '').strip() or None
        usuario.rol      = request.form.get('rol', usuario.rol)
        usuario.dni      = request.form.get('dni', '').strip() or None
        nuevo_pass       = request.form.get('password', '').strip()
        if nuevo_pass:
            usuario.password = generate_password_hash(nuevo_pass)
        registrar_log('editar_usuario', 'Usuario', id)
        db.session.commit()
        registrar_auditoria(
            f'Editó usuario {usuario.nombre}',
            'Usuarios'
        )
        flash('Usuario actualizado.', 'success')
        return redirect(url_for('usuarios'))
    return render_template('editar_usuario.html', usuario=usuario)


@app.route('/usuarios/toggle/<int:id>', methods=['POST'])
@login_required
@rol_requerido('admin')
def toggle_usuario(id):
    usuario = Usuario.query.get_or_404(id)
    if usuario.id == current_user.id:
        flash('No puedes desactivar tu propia cuenta.', 'danger')
        return redirect(url_for('usuarios'))
    usuario.activo = not usuario.activo
    accion = 'activar_usuario' if usuario.activo else 'desactivar_usuario'
    registrar_log(accion, 'Usuario', id)
    db.session.commit()
    registrar_auditoria(
        f'{"Activó" if usuario.activo else "Desactivó"} usuario {usuario.nombre}',
        'Usuarios'
    )
    flash(f'Usuario {"activado" if usuario.activo else "desactivado"}.', 'success')
    return redirect(url_for('usuarios'))


@app.route('/usuarios/eliminar/<int:id>', methods=['POST'])
@login_required
@rol_requerido('admin')
def eliminar_usuario(id):
    usuario = Usuario.query.get_or_404(id)
    if usuario.id == current_user.id:
        flash('No puedes eliminar tu propia cuenta.', 'danger')
        return redirect(url_for('usuarios'))
    # Reasignar trámites a anónimo antes de eliminar
    Tramite.query.filter_by(ciudadano_id=id).update({'ciudadano_id': current_user.id})
    db.session.delete(usuario)
    registrar_log('eliminar_usuario', 'Usuario', id)
    db.session.commit()
    registrar_auditoria(
        f'Eliminó usuario {usuario.nombre}',
        'Usuarios'
    )
    flash('Usuario eliminado.', 'success')
    return redirect(url_for('usuarios'))


# ════════════════════════════════════════════════════════════════════
# ALERTAS
# ════════════════════════════════════════════════════════════════════
@app.route('/alertas')
@login_required
def alertas():
    if current_user.rol == 'admin':
        lista = Alerta.query.order_by(Alerta.fecha.desc()).all()
    else:
        lista = Alerta.query.filter_by(
            ciudadano_id=current_user.id).order_by(Alerta.fecha.desc()).all()
    return render_template('alertas.html', alertas=lista)


@app.route('/alertas/marcar-leido/<int:id>', methods=['POST'])
@login_required
def marcar_leido(id):
    alerta = Alerta.query.get_or_404(id)
    alerta.leido = True
    db.session.commit()
    return jsonify({'ok': True})


@app.route('/alertas/marcar-todas', methods=['POST'])
@login_required
def marcar_todas_leidas():
    q = Alerta.query.filter_by(leido=False)
    if current_user.rol != 'admin':
        q = q.filter_by(ciudadano_id=current_user.id)
    q.update({'leido': True})
    db.session.commit()
    flash('Todas las alertas marcadas como leídas.', 'success')
    return redirect(url_for('alertas'))


# ════════════════════════════════════════════════════════════════════
# DASHBOARD DE ML — MÉTRICAS
# ════════════════════════════════════════════════════════════════════
@app.route('/ml/stats')
@login_required
@rol_requerido('admin')
def ml_stats():
    st_tramites = obtener_stats_tramites()
    st_cvs      = obtener_stats_cvs()
    return render_template('ml_stats.html',
                           st_tramites=st_tramites,
                           st_cvs=st_cvs)


@app.route('/ml/stats/json')
@login_required
@rol_requerido('admin')
def ml_stats_json():
    return jsonify({
        'tramites': obtener_stats_tramites(),
        'cvs':      obtener_stats_cvs()
    })


@app.route('/ml/reentrenar', methods=['POST'])
@login_required
@rol_requerido('admin')
def ml_reentrenar():
    resultado = reentrenar()
    registrar_log('reentrenar_modelos', 'ML')
    db.session.commit()
    registrar_auditoria(
        'Reentrenó los modelos de Machine Learning',
        'Machine Learning'
    )
    flash(f'Modelos reentrenados. Tramites accuracy='
          f'{resultado["tramites"].get("accuracy","?")} | '
          f'CVs accuracy={resultado["cvs"].get("accuracy","?")}', 'success')
    return redirect(url_for('ml_stats'))


# ════════════════════════════════════════════════════════════════════
# EXPORTAR EXCEL — openpyxl (equivalente a ExcelJS en Python)
# ════════════════════════════════════════════════════════════════════
VERDE     = 'FF2E7D4F'
VERDE_CLR = 'FFE1F5EE'
GRIS_HD   = 'FF444441'
GRIS_FIL  = 'FFF1EFE8'
BLANCO    = 'FFFFFFFF'

def _borde(grosor='thin'):
    s = Side(style=grosor)
    return Border(left=s, right=s, top=s, bottom=s)


def _encabezado_institucional(ws, titulo, subtitulo='Municipalidad Provincial de Yau'):
    """Crea cabecera institucional verde en la hoja."""
    ws.merge_cells('A1:H1')
    c = ws['A1']
    c.value          = 'GESTIMUNI — SISTEMA DE GESTIÓN MUNICIPAL CON IA'
    c.font            = Font(name='Calibri', bold=True, size=14, color=BLANCO)
    c.fill            = PatternFill('solid', fgColor=VERDE)
    c.alignment       = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 28

    ws.merge_cells('A2:H2')
    c = ws['A2']
    c.value           = subtitulo
    c.font            = Font(name='Calibri', size=10, color=BLANCO, italic=True)
    c.fill            = PatternFill('solid', fgColor='FF1B5E35')
    c.alignment       = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A3:H3')
    c = ws['A3']
    c.value           = titulo
    c.font            = Font(name='Calibri', bold=True, size=12, color=GRIS_HD)
    c.fill            = PatternFill('solid', fgColor=VERDE_CLR)
    c.alignment       = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[3].height = 22

    ws.merge_cells('A4:D4')
    ws['A4'].value = f'Fecha de generación: {datetime.now().strftime("%d/%m/%Y %H:%M")}'
    ws['A4'].font  = Font(name='Calibri', size=9, italic=True, color='FF888780')
    ws.row_dimensions[4].height = 16


def _fila_encabezado_tabla(ws, fila, columnas):
    for col_idx, titulo in enumerate(columnas, 1):
        c = ws.cell(row=fila, column=col_idx, value=titulo)
        c.font      = Font(name='Calibri', bold=True, size=10, color=BLANCO)
        c.fill      = PatternFill('solid', fgColor=VERDE)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border    = _borde()
    ws.row_dimensions[fila].height = 20


def _fila_datos(ws, fila, datos, alternado=False):
    fill_color = GRIS_FIL if alternado else BLANCO
    for col_idx, valor in enumerate(datos, 1):
        c = ws.cell(row=fila, column=col_idx, value=valor)
        c.font      = Font(name='Calibri', size=9)
        c.fill      = PatternFill('solid', fgColor=fill_color)
        c.alignment = Alignment(vertical='center', wrap_text=True)
        c.border    = _borde('thin')
    ws.row_dimensions[fila].height = 16


@app.route('/exportar/tramites')
@login_required
@rol_requerido('admin', 'empleado')
def exportar_tramites_excel():
    estado   = request.args.get('estado', '')
    prioridad= request.args.get('prioridad', '')

    query = Tramite.query
    if estado:
        query = query.filter_by(estado=estado)
    if prioridad:
        query = query.filter_by(prioridad_ml=prioridad)
    if current_user.rol == 'empleado':
        query = query.filter_by(empleado_id=current_user.id)
    lista = query.order_by(Tramite.fecha_registro.desc()).all()

    wb = Workbook()
    ws = wb.active
    ws.title = 'Trámites'

    _encabezado_institucional(ws,
        f'REPORTE DE TRÁMITES'
        + (f' — Estado: {estado.upper()}' if estado else '')
        + (f' — Prioridad: {prioridad.upper()}' if prioridad else ''))

    cols = ['N°', 'Ciudadano', 'DNI', 'Tipo de Trámite',
            'Urgencia', 'Prioridad ML', 'Estado', 'Confianza ML', 'Fecha']
    anchos = [5, 22, 10, 22, 10, 12, 12, 12, 16]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    _fila_encabezado_tabla(ws, 6, cols)

    COLORES_PRIORIDAD = {'critico': 'FFFCEBEB', 'normal': 'FFFAEEDA', 'bajo': 'FFE1F5EE'}
    for idx, t in enumerate(lista):
        fila = 7 + idx
        dni_c = t.ciudadano.dni or '—'
        datos = [
            idx + 1,
            t.ciudadano.nombre,
            dni_c,
            t.tipo.replace('_', ' ').title(),
            t.urgencia.upper(),
            t.prioridad_ml.upper(),
            t.estado.upper().replace('_', ' '),
            f'{round(t.confianza_ml * 100, 1)}%',
            t.fecha_registro.strftime('%d/%m/%Y %H:%M')
        ]
        _fila_datos(ws, fila, datos, idx % 2 == 1)
        # Color según prioridad en columna F
        color_p = COLORES_PRIORIDAD.get(t.prioridad_ml, BLANCO)
        ws.cell(fila, 6).fill = PatternFill('solid', fgColor=color_p)

    # Fila resumen
    fila_res = 7 + len(lista) + 1
    ws.merge_cells(f'A{fila_res}:G{fila_res}')
    ws[f'A{fila_res}'].value = f'Total: {len(lista)} trámites'
    ws[f'A{fila_res}'].font  = Font(bold=True, size=10)
    ws[f'A{fila_res}'].fill  = PatternFill('solid', fgColor=VERDE_CLR)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    nombre_archivo = f'tramites_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
    return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=nombre_archivo)


@app.route('/exportar/curriculos')
@login_required
@rol_requerido('admin', 'empleado')
def exportar_curriculos_excel():
    lista = Postulante.query.order_by(Postulante.puntaje_ml.desc()).all()

    wb = Workbook()
    ws = wb.active
    ws.title = 'Postulantes'

    _encabezado_institucional(ws, 'REPORTE DE EVALUACIÓN DE CURRÍCULOS — ML')

    cols   = ['N°', 'Nombre', 'DNI', 'Puesto', 'Exp. (años)',
              'Educación', 'Habilidades', 'Idiomas', 'Resultado ML',
              'Puntaje ML', 'Probabilidad', 'Fecha']
    anchos = [4, 22, 10, 18, 10, 14, 11, 9, 12, 11, 12, 16]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    _fila_encabezado_tabla(ws, 6, cols)

    for idx, p in enumerate(lista):
        fila = 7 + idx
        datos = [
            idx + 1, p.nombre, p.dni or '—', p.puesto,
            p.experiencia, p.educacion.title(),
            p.habilidades, p.idiomas or '0',
            p.resultado_ml.upper(),
            f'{p.puntaje_ml}%',
            f'{round(p.probabilidad * 100, 1)}%',
            p.fecha_registro.strftime('%d/%m/%Y')
        ]
        _fila_datos(ws, fila, datos, idx % 2 == 1)
        color_r = 'FFE1F5EE' if p.resultado_ml == 'apto' else 'FFFCEBEB'
        ws.cell(fila, 9).fill = PatternFill('solid', fgColor=color_r)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    nombre_archivo = f'curriculos_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
    return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=nombre_archivo)

@app.route('/auditoria')
@login_required
@rol_requerido('admin')
def auditoria():
    registros = Auditoria.query.order_by(Auditoria.fecha.desc()).all()
    return render_template('auditoria.html', registros=registros)

# ════════════════════════════════════════════════════════════════════
# INICIALIZAR BD
# ════════════════════════════════════════════════════════════════════
def inicializar_bd():
    with app.app_context():
        db.create_all()
        if not Usuario.query.first():
            usuarios_default = [
                Usuario(nombre='Administrador', username='admin',
                        password=generate_password_hash('admin123'),
                        rol='admin', email='admin@gestimuni.pe',
                        dni='00000001'),
                Usuario(nombre='María Quispe', username='empleado1',
                        password=generate_password_hash('emp123'),
                        rol='empleado', email='empleado@gestimuni.pe',
                        dni='12345678'),
                Usuario(nombre='Juan Pérez', username='ciudadano1',
                        password=generate_password_hash('1234'),
                        rol='ciudadano', email='ciudadano@gmail.com',
                        dni='87654321'),
            ]
            db.session.add_all(usuarios_default)
            db.session.commit()
            print('[BD] Inicializada con usuarios por defecto.')


if __name__ == '__main__':
    inicializar_bd()
    app.run(debug=True)