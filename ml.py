import numpy as np
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder

# ── MODELO 1: Clasificador de trámites ──────────────────────────────
# Datos de entrenamiento simulados
tipos = ['licencia_construccion', 'certificado_residencia', 'permiso_negocio',
         'denuncia_peligro', 'partida_nacimiento', 'licencia_funcionamiento']

X_tramites = [
    [0, 2], [0, 1], [1, 2], [2, 0], [1, 1], [2, 2],
    [0, 2], [2, 1], [1, 0], [0, 1], [2, 2], [1, 2],
    [0, 0], [2, 0], [1, 1], [0, 2], [2, 1], [1, 0],
]
# 0=critico, 1=normal, 2=bajo
y_tramites = [0, 1, 0, 0, 1, 2, 0, 1, 2, 1, 0, 0, 2, 0, 1, 0, 1, 2]

modelo_tramites = RandomForestClassifier(n_estimators=10, random_state=42)
modelo_tramites.fit(X_tramites, y_tramites)

# ── MODELO 2: Evaluador de CVs ───────────────────────────────────────
# experiencia(años), educacion(0-3), habilidades(count)
X_cvs = [
    [5, 3, 8], [1, 1, 2], [3, 2, 5], [0, 0, 1], [7, 3, 10],
    [2, 1, 3], [4, 2, 6], [0, 1, 1], [6, 3, 9], [1, 0, 2],
    [3, 2, 4], [8, 3, 10], [0, 0, 0], [2, 2, 5], [5, 3, 7],
]
# 1=apto, 0=no_apto
y_cvs = [1, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1, 1, 0, 1, 1]

modelo_cvs = RandomForestClassifier(n_estimators=10, random_state=42)
modelo_cvs.fit(X_cvs, y_cvs)


def clasificar_tramite(tipo, urgencia):
    tipos_map = {
        'licencia_construccion': 0, 'denuncia_peligro': 0,
        'licencia_funcionamiento': 1, 'permiso_negocio': 1,
        'certificado_residencia': 2, 'partida_nacimiento': 2
    }
    urgencia_map = {'alta': 0, 'media': 1, 'baja': 2}
    prioridades = {0: 'critico', 1: 'normal', 2: 'bajo'}

    t = tipos_map.get(tipo, 1)
    u = urgencia_map.get(urgencia, 1)
    resultado = modelo_tramites.predict([[t, u]])[0]
    return prioridades[resultado]


def evaluar_cv(experiencia, educacion, habilidades):
    educacion_map = {
        'secundaria': 0, 'tecnico': 1,
        'universitario': 2, 'postgrado': 3
    }
    edu = educacion_map.get(educacion, 1)
    prob = modelo_cvs.predict_proba([[experiencia, edu, habilidades]])[0][1]
    resultado = 'apto' if prob >= 0.5 else 'no_apto'
    puntaje = round(prob * 100, 2)
    return resultado, puntaje