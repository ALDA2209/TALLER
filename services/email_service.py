"""
services/email_service.py
Envía correos HTML profesionales al ciudadano cuando cambia el estado de su trámite.
"""
from flask_mail import Message
from flask import current_app


def _plantilla_base(cuerpo_html: str) -> str:
    return f"""
<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f4f4f4;font-family:Arial,Helvetica,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f4;padding:30px 0;">
  <tr><td align="center">
    <table width="560" cellpadding="0" cellspacing="0"
           style="background:#ffffff;border-radius:10px;overflow:hidden;
                  border:1px solid #dee2e6;max-width:560px;width:100%;">
      <!-- CABECERA -->
      <tr>
        <td style="background:#2E7D4F;padding:22px 30px;text-align:center;">
          <h1 style="color:#ffffff;margin:0;font-size:22px;letter-spacing:1px;">
            🏛 GestiMuni Huánuco
          </h1>
          <p style="color:#a8d5ba;margin:6px 0 0;font-size:13px;">
            Portal de Servicios Municipales Digitales
          </p>
        </td>
      </tr>
      <!-- CUERPO -->
      <tr><td style="padding:28px 30px;">{cuerpo_html}</td></tr>
      <!-- PIE -->
      <tr>
        <td style="background:#f8f9fa;padding:16px 30px;text-align:center;
                   border-top:1px solid #dee2e6;">
          <p style="color:#aaa;font-size:11px;margin:0;">
            © 2026 GestiMuni Huánuco — Municipalidad Provincial de Yau<br>
            Este correo es generado automáticamente, por favor no respondas.
          </p>
        </td>
      </tr>
    </table>
  </td></tr>
</table>
</body>
</html>"""


def _color_estado(estado: str):
    mapa = {
        'aprobado':  ('#1D9E75', '#E1F5EE', '✅'),
        'rechazado': ('#A32D2D', '#FCEBEB', '❌'),
        'en_proceso':('#185FA5', '#E6F1FB', '🔄'),
        'pendiente': ('#854F0B', '#FAEEDA', '⏳'),
    }
    return mapa.get(estado, ('#444444', '#f5f5f5', '📋'))


def enviar_alerta_nuevo_tramite(mail, ciudadano, tramite):
    """Correo de confirmación cuando el ciudadano registra un trámite."""
    color, bg, icono = _color_estado('pendiente')
    prioridad_color = {
        'critico': '#A32D2D', 'normal': '#854F0B', 'bajo': '#1D9E75'
    }.get(tramite.prioridad_ml, '#444')

    cuerpo = f"""
    <p style="color:#333;font-size:15px;margin:0 0 16px;">
      Hola <strong>{ciudadano.nombre}</strong>,
    </p>
    <p style="color:#555;font-size:14px;margin:0 0 20px;">
      Tu trámite ha sido registrado exitosamente en el sistema GestiMuni.
      A continuación el resumen:
    </p>
    <!-- Tarjeta de trámite -->
    <table width="100%" cellpadding="0" cellspacing="0"
           style="background:#f8f9fa;border-radius:8px;border:1px solid #dee2e6;
                  margin-bottom:20px;">
      <tr>
        <td style="padding:16px 20px;">
          <table width="100%" cellpadding="4" cellspacing="0">
            <tr>
              <td style="color:#888;font-size:12px;width:140px;">N° de trámite</td>
              <td style="color:#333;font-size:13px;font-weight:bold;">#{tramite.id}</td>
            </tr>
            <tr>
              <td style="color:#888;font-size:12px;">Tipo</td>
              <td style="color:#333;font-size:13px;">
                {tramite.tipo.replace('_',' ').title()}
              </td>
            </tr>
            <tr>
              <td style="color:#888;font-size:12px;">Fecha de registro</td>
              <td style="color:#333;font-size:13px;">
                {tramite.fecha_registro.strftime('%d/%m/%Y %H:%M')}
              </td>
            </tr>
            <tr>
              <td style="color:#888;font-size:12px;">Prioridad IA</td>
              <td>
                <span style="background:{prioridad_color};color:#fff;
                             font-size:11px;padding:3px 10px;border-radius:4px;
                             font-weight:bold;">
                  {tramite.prioridad_ml.upper()}
                </span>
                <span style="color:#888;font-size:11px;margin-left:6px;">
                  (confianza: {round(tramite.confianza_ml*100,1)}%)
                </span>
              </td>
            </tr>
            <tr>
              <td style="color:#888;font-size:12px;">Estado actual</td>
              <td>
                <span style="background:{bg};color:{color};font-size:11px;
                             padding:3px 10px;border-radius:4px;font-weight:bold;">
                  PENDIENTE
                </span>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
    <p style="color:#555;font-size:13px;margin:0 0 8px;">
      Te notificaremos por este correo cada vez que cambie el estado de tu solicitud.
    </p>
    <p style="color:#888;font-size:12px;margin:0;">
      También puedes ingresar al portal en cualquier momento para consultar el avance.
    </p>"""

    html = _plantilla_base(cuerpo)
    try:
        msg = Message(
            subject=f'GestiMuni — Trámite #{tramite.id} registrado correctamente',
            recipients=[ciudadano.email],
            html=html
        )
        mail.send(msg)
        return True
    except Exception as e:
        current_app.logger.error(f'[EMAIL] Error enviando confirmación: {e}')
        return False


def enviar_alerta_cambio_estado(mail, ciudadano, tramite, observaciones=''):
    """Correo cuando admin/empleado cambia el estado del trámite."""
    color, bg, icono = _color_estado(tramite.estado)

    obs_html = ''
    if observaciones:
        obs_html = f"""
        <div style="background:#fff8e1;border-left:3px solid #EF9F27;
                    padding:12px 16px;border-radius:0 6px 6px 0;margin-bottom:20px;">
          <p style="color:#633806;font-size:13px;margin:0;">
            <strong>Nota del operador:</strong> {observaciones}
          </p>
        </div>"""

    cuerpo = f"""
    <p style="color:#333;font-size:15px;margin:0 0 16px;">
      Hola <strong>{ciudadano.nombre}</strong>,
    </p>
    <p style="color:#555;font-size:14px;margin:0 0 20px;">
      El estado de tu trámite <strong>#{tramite.id}</strong>
      ({tramite.tipo.replace('_',' ').title()}) ha sido actualizado:
    </p>
    <!-- Estado grande -->
    <div style="text-align:center;padding:20px;background:{bg};
                border-radius:8px;margin-bottom:20px;">
      <div style="font-size:32px;margin-bottom:8px;">{icono}</div>
      <div style="font-size:22px;font-weight:bold;color:{color};
                  letter-spacing:2px;">
        {tramite.estado.upper().replace('_',' ')}
      </div>
      <div style="font-size:12px;color:{color};margin-top:4px;opacity:.8;">
        Actualizado el {tramite.fecha_actualizacion.strftime('%d/%m/%Y a las %H:%M') if tramite.fecha_actualizacion else 'ahora'}
      </div>
    </div>
    {obs_html}
    <table width="100%" cellpadding="4" cellspacing="0"
           style="font-size:13px;color:#555;margin-bottom:20px;">
      <tr>
        <td style="color:#888;">N° de trámite:</td>
        <td style="font-weight:bold;">#{tramite.id}</td>
      </tr>
      <tr>
        <td style="color:#888;">Tipo:</td>
        <td>{tramite.tipo.replace('_',' ').title()}</td>
      </tr>
      <tr>
        <td style="color:#888;">Prioridad IA:</td>
        <td>{tramite.prioridad_ml.upper()}</td>
      </tr>
    </table>
    <p style="color:#888;font-size:12px;margin:0;">
      Ingresa al portal para ver el detalle completo de tu trámite.
    </p>"""

    html = _plantilla_base(cuerpo)
    try:
        msg = Message(
            subject=f'GestiMuni — Tu trámite #{tramite.id} cambió a {tramite.estado.upper().replace("_"," ")}',
            recipients=[ciudadano.email],
            html=html
        )
        mail.send(msg)
        return True
    except Exception as e:
        current_app.logger.error(f'[EMAIL] Error enviando alerta estado: {e}')
        return False


def enviar_alerta_tramite_vencido(mail, ciudadano, tramite, dias_pendiente):
    """Correo automático cuando un trámite lleva demasiados días sin respuesta."""
    cuerpo = f"""
    <p style="color:#333;font-size:15px;margin:0 0 16px;">
      Hola <strong>{ciudadano.nombre}</strong>,
    </p>
    <div style="background:#FAEEDA;border-left:4px solid #EF9F27;
                padding:16px 20px;border-radius:0 8px 8px 0;margin-bottom:20px;">
      <p style="color:#633806;font-size:14px;margin:0;">
        ⚠ Tu trámite <strong>#{tramite.id}</strong> lleva
        <strong>{dias_pendiente} días</strong> sin actualización.
        Estamos trabajando para resolverlo a la brevedad.
      </p>
    </div>
    <table width="100%" cellpadding="4" cellspacing="0"
           style="font-size:13px;color:#555;background:#f8f9fa;
                  border-radius:8px;padding:12px;">
      <tr><td style="color:#888;padding:4px;">N° de trámite:</td>
          <td style="font-weight:bold;">#{tramite.id}</td></tr>
      <tr><td style="color:#888;padding:4px;">Tipo:</td>
          <td>{tramite.tipo.replace('_',' ').title()}</td></tr>
      <tr><td style="color:#888;padding:4px;">Días pendiente:</td>
          <td style="color:#854F0B;font-weight:bold;">{dias_pendiente} días</td></tr>
    </table>
    <p style="color:#888;font-size:12px;margin-top:16px;">
      Si tienes urgencia, visita la municipalidad con tu número de trámite.
    </p>"""

    html = _plantilla_base(cuerpo)
    try:
        msg = Message(
            subject=f'GestiMuni — Tu trámite #{tramite.id} está pendiente de atención',
            recipients=[ciudadano.email],
            html=html
        )
        mail.send(msg)
        return True
    except Exception as e:
        current_app.logger.error(f'[EMAIL] Error enviando alerta vencido: {e}')
        return False