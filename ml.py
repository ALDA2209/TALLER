"""
ml.py — Motor de Machine Learning de GestiMuni
Incluye:
  - Dataset sintético realista (500+ muestras)
  - Pipeline con StandardScaler + RandomForestClassifier
  - GridSearchCV para hiperparámetros
  - Métricas completas: accuracy, F1, matriz de confusión
  - Serialización con joblib
  - Funciones de predicción con probabilidades
"""

import numpy as np
import os
import json
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.pipeline import Pipeline
from sklearn.model_selection import train_test_split, GridSearchCV, cross_val_score
from sklearn.metrics import (accuracy_score, f1_score, classification_report,
                              confusion_matrix, roc_auc_score)
import joblib
import warnings
warnings.filterwarnings('ignore')

# ────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN DE RUTAS
# ────────────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
MODEL_DIR   = os.path.join(BASE_DIR, 'models_ml')
os.makedirs(MODEL_DIR, exist_ok=True)

PATH_TRAMITES_MODEL  = os.path.join(MODEL_DIR, 'modelo_tramites.pkl')
PATH_CVS_MODEL       = os.path.join(MODEL_DIR, 'modelo_cvs.pkl')
PATH_TRAMITES_STATS  = os.path.join(MODEL_DIR, 'stats_tramites.json')
PATH_CVS_STATS       = os.path.join(MODEL_DIR, 'stats_cvs.json')


# ────────────────────────────────────────────────────────────────────
# DATASET TRÁMITES — 540 muestras sintéticas realistas
# ────────────────────────────────────────────────────────────────────
def _generar_dataset_tramites(n=540, seed=42):
    """
    Features:
      tipo_cod   : 0=licencia_construccion, 1=licencia_funcionamiento,
                   2=permiso_negocio, 3=denuncia_peligro,
                   4=certificado_residencia, 5=partida_nacimiento
      urgencia   : 0=alta, 1=media, 2=baja
      dias_ant   : días desde el ingreso (0-60)
      n_docs     : número de documentos adjuntos (0-5)

    Etiqueta: 0=critico, 1=normal, 2=bajo
    """
    rng = np.random.RandomState(seed)
    X, y = [], []

    reglas = {
        # (tipo, urgencia) → prioridad base, varianza
        (0, 0): (0, 0.05),  # construccion + alta   → critico
        (0, 1): (0, 0.15),
        (0, 2): (1, 0.15),
        (1, 0): (0, 0.10),  # funcionamiento + alta → critico
        (1, 1): (1, 0.10),
        (1, 2): (1, 0.20),
        (2, 0): (0, 0.10),
        (2, 1): (1, 0.10),
        (2, 2): (2, 0.10),
        (3, 0): (0, 0.02),  # denuncia peligro      → casi siempre critico
        (3, 1): (0, 0.05),
        (3, 2): (0, 0.10),
        (4, 0): (1, 0.10),
        (4, 1): (2, 0.10),
        (4, 2): (2, 0.05),
        (5, 0): (1, 0.15),
        (5, 1): (2, 0.10),
        (5, 2): (2, 0.05),
    }

    for _ in range(n):
        tipo    = rng.randint(0, 6)
        urgencia= rng.randint(0, 3)
        dias    = rng.randint(0, 61)
        n_docs  = rng.randint(0, 6)

        base, var = reglas.get((tipo, urgencia), (1, 0.15))
        # días sin atender aumentan la prioridad
        bonus = -1 if dias > 30 else 0
        raw = base + bonus + rng.normal(0, var)
        etq = int(np.clip(round(raw), 0, 2))

        X.append([tipo, urgencia, dias, n_docs])
        y.append(etq)

    return np.array(X), np.array(y)


# ────────────────────────────────────────────────────────────────────
# DATASET CVs — 480 muestras
# ────────────────────────────────────────────────────────────────────
def _generar_dataset_cvs(n=480, seed=7):
    """
    Features:
      experiencia : 0-15 años
      educacion   : 0=secundaria, 1=tecnico, 2=universitario, 3=postgrado
      habilidades : 0-10
      idiomas     : 0-3 (número de idiomas adicionales)

    Etiqueta: 1=apto, 0=no_apto
    """
    rng = np.random.RandomState(seed)
    X, y = [], []

    for _ in range(n):
        exp   = rng.randint(0, 16)
        edu   = rng.randint(0, 4)
        hab   = rng.randint(0, 11)
        idm   = rng.randint(0, 4)

        # Puntaje ponderado con ruido
        score = (exp * 0.35 + edu * 1.5 + hab * 0.4 + idm * 0.8
                 + rng.normal(0, 0.8))
        # umbral≈5.5 → apto
        etq = 1 if score >= 5.5 else 0

        X.append([exp, edu, hab, idm])
        y.append(etq)

    return np.array(X), np.array(y)


# ────────────────────────────────────────────────────────────────────
# ENTRENAMIENTO
# ────────────────────────────────────────────────────────────────────
def _entrenar_modelo_tramites():
    X, y = _generar_dataset_tramites()
    X_tr, X_te, y_tr, y_te = train_test_split(
        X, y, test_size=0.2, random_state=42, stratify=y)

    pipeline = Pipeline([
        ('scaler', StandardScaler()),
        ('clf', RandomForestClassifier(random_state=42))
    ])

    param_grid = {
        'clf__n_estimators': [100, 200],
        'clf__max_depth':    [None, 10, 20],
        'clf__min_samples_split': [2, 5],
    }
    gs = GridSearchCV(pipeline, param_grid, cv=5,
                      scoring='f1_macro', n_jobs=-1)
    gs.fit(X_tr, y_tr)

    best = gs.best_estimator_
    y_pred = best.predict(X_te)

    stats = {
        'accuracy':        round(float(accuracy_score(y_te, y_pred)), 4),
        'f1_macro':        round(float(f1_score(y_te, y_pred, average='macro')), 4),
        'f1_weighted':     round(float(f1_score(y_te, y_pred, average='weighted')), 4),
        'best_params':     gs.best_params_,
        'cv_scores':       [round(float(s), 4) for s in
                            cross_val_score(best, X, y, cv=5, scoring='f1_macro')],
        'confusion_matrix': confusion_matrix(y_te, y_pred).tolist(),
        'report':          classification_report(y_te, y_pred,
                               target_names=['critico','normal','bajo'],
                               output_dict=True),
        'feature_names':   ['tipo_tramite','urgencia','dias_espera','num_documentos'],
        'feature_importances': [
            round(float(v), 4)
            for v in best.named_steps['clf'].feature_importances_
        ],
        'n_train':  len(X_tr),
        'n_test':   len(X_te),
        'clases':   ['critico', 'normal', 'bajo'],
    }

    joblib.dump(best, PATH_TRAMITES_MODEL)
    with open(PATH_TRAMITES_STATS, 'w', encoding='utf-8') as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)

    print(f"[ML] Trámites — accuracy={stats['accuracy']}, F1={stats['f1_macro']}")
    return best, stats


def _entrenar_modelo_cvs():
    X, y = _generar_dataset_cvs()
    X_tr, X_te, y_tr, y_te = train_test_split(
        X, y, test_size=0.2, random_state=7, stratify=y)

    pipeline = Pipeline([
        ('scaler', StandardScaler()),
        ('clf', GradientBoostingClassifier(random_state=7))
    ])

    param_grid = {
        'clf__n_estimators':   [100, 200],
        'clf__learning_rate':  [0.05, 0.1],
        'clf__max_depth':      [3, 5],
    }
    gs = GridSearchCV(pipeline, param_grid, cv=5,
                      scoring='f1', n_jobs=-1)
    gs.fit(X_tr, y_tr)

    best = gs.best_estimator_
    y_pred     = best.predict(X_te)
    y_prob     = best.predict_proba(X_te)[:, 1]

    stats = {
        'accuracy':        round(float(accuracy_score(y_te, y_pred)), 4),
        'f1':              round(float(f1_score(y_te, y_pred)), 4),
        'roc_auc':         round(float(roc_auc_score(y_te, y_prob)), 4),
        'best_params':     gs.best_params_,
        'cv_scores':       [round(float(s), 4) for s in
                            cross_val_score(best, X, y, cv=5, scoring='f1')],
        'confusion_matrix': confusion_matrix(y_te, y_pred).tolist(),
        'report':          classification_report(y_te, y_pred,
                               target_names=['no_apto','apto'],
                               output_dict=True),
        'feature_names':   ['experiencia_anios','nivel_educacion',
                             'num_habilidades','num_idiomas'],
        'feature_importances': [
            round(float(v), 4)
            for v in best.named_steps['clf'].feature_importances_
        ],
        'n_train':  len(X_tr),
        'n_test':   len(X_te),
        'clases':   ['no_apto', 'apto'],
    }

    joblib.dump(best, PATH_CVS_MODEL)
    with open(PATH_CVS_STATS, 'w', encoding='utf-8') as f:
        json.dump(stats, f, ensure_ascii=False, indent=2)

    print(f"[ML] CVs — accuracy={stats['accuracy']}, F1={stats['f1']}, AUC={stats['roc_auc']}")
    return best, stats


# ────────────────────────────────────────────────────────────────────
# CARGA O ENTRENAMIENTO AL IMPORTAR
# ────────────────────────────────────────────────────────────────────
def _cargar_o_entrenar(path_model, train_fn):
    if os.path.exists(path_model):
        return joblib.load(path_model)
    model, _ = train_fn()
    return model


modelo_tramites = _cargar_o_entrenar(PATH_TRAMITES_MODEL, _entrenar_modelo_tramites)
modelo_cvs      = _cargar_o_entrenar(PATH_CVS_MODEL,      _entrenar_modelo_cvs)


# ────────────────────────────────────────────────────────────────────
# API PÚBLICA
# ────────────────────────────────────────────────────────────────────
TIPOS_MAP = {
    'licencia_construccion':   0,
    'licencia_funcionamiento': 1,
    'permiso_negocio':         2,
    'denuncia_peligro':        3,
    'certificado_residencia':  4,
    'partida_nacimiento':      5,
}
URGENCIA_MAP = {'alta': 0, 'media': 1, 'baja': 2}
PRIORIDADES  = {0: 'critico', 1: 'normal', 2: 'bajo'}

EDUCACION_MAP = {
    'secundaria': 0, 'tecnico': 1,
    'universitario': 2, 'postgrado': 3,
}


def clasificar_tramite(tipo, urgencia, dias_espera=0, num_docs=0):
    """Devuelve (prioridad_str, confianza_float 0-1)."""
    t = TIPOS_MAP.get(tipo, 2)
    u = URGENCIA_MAP.get(urgencia, 1)
    feat = np.array([[t, u, dias_espera, num_docs]])
    resultado  = modelo_tramites.predict(feat)[0]
    proba      = modelo_tramites.predict_proba(feat)[0]
    confianza  = round(float(max(proba)), 4)
    return PRIORIDADES[resultado], confianza


def evaluar_cv(experiencia, educacion, habilidades, idiomas=0):
    """Devuelve (resultado_str, puntaje_pct, probabilidad_float)."""
    edu   = EDUCACION_MAP.get(educacion, 1) if isinstance(educacion, str) else int(educacion)
    feat  = np.array([[experiencia, edu, habilidades, idiomas]])
    prob  = float(modelo_cvs.predict_proba(feat)[0][1])
    resultado = 'apto' if prob >= 0.5 else 'no_apto'
    puntaje   = round(prob * 100, 2)
    return resultado, puntaje, round(prob, 4)


def obtener_stats_tramites():
    """Lee el JSON de métricas del modelo de trámites."""
    if os.path.exists(PATH_TRAMITES_STATS):
        with open(PATH_TRAMITES_STATS, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def obtener_stats_cvs():
    """Lee el JSON de métricas del modelo de CVs."""
    if os.path.exists(PATH_CVS_STATS):
        with open(PATH_CVS_STATS, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def reentrenar():
    """Fuerza reentrenamiento y recarga los modelos globales."""
    global modelo_tramites, modelo_cvs
    modelo_tramites, _ = _entrenar_modelo_tramites()
    modelo_cvs,      _ = _entrenar_modelo_cvs()
    return {
        'tramites': obtener_stats_tramites(),
        'cvs':      obtener_stats_cvs(),
    }