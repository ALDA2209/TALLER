import pandas as pd
import joblib

from sklearn.compose import ColumnTransformer
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import accuracy_score, classification_report
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder


def entrenar_modelo_tramites():
    df = pd.read_csv("ml_models/dataset_tramites.csv")

    X = df[[
        "tipo",
        "urgencia",
        "dias_resolucion",
        "cantidad_documentos"
    ]]
    y = df["prioridad"]

    preprocesador = ColumnTransformer(
        transformers=[
            ("categoricas", OneHotEncoder(handle_unknown="ignore"), ["tipo", "urgencia"])
        ],
        remainder="passthrough"
    )

    modelo = Pipeline(steps=[
        ("preprocesador", preprocesador),
        ("clasificador", RandomForestClassifier(
            n_estimators=200,
            random_state=42,
            max_depth=6
        ))
    ])

    X_train, X_test, y_train, y_test = train_test_split(
        X,
        y,
        test_size=0.25,
        random_state=42,
        stratify=y
    )

    modelo.fit(X_train, y_train)
    pred = modelo.predict(X_test)

    accuracy = accuracy_score(y_test, pred)
    reporte = classification_report(y_test, pred, zero_division=0)

    joblib.dump(modelo, "ml_models/modelo_tramites.pkl")

    return accuracy, reporte


def entrenar_modelo_curriculos():
    df = pd.read_csv("ml_models/dataset_curriculos.csv")

    X = df[[
        "experiencia",
        "educacion",
        "habilidades",
        "certificaciones"
    ]]
    y = df["resultado"]

    preprocesador = ColumnTransformer(
        transformers=[
            ("categoricas", OneHotEncoder(handle_unknown="ignore"), ["educacion"])
        ],
        remainder="passthrough"
    )

    modelo = Pipeline(steps=[
        ("preprocesador", preprocesador),
        ("clasificador", RandomForestClassifier(
            n_estimators=200,
            random_state=42,
            max_depth=6
        ))
    ])

    X_train, X_test, y_train, y_test = train_test_split(
        X,
        y,
        test_size=0.25,
        random_state=42,
        stratify=y
    )

    modelo.fit(X_train, y_train)
    pred = modelo.predict(X_test)

    accuracy = accuracy_score(y_test, pred)
    reporte = classification_report(y_test, pred, zero_division=0)

    joblib.dump(modelo, "ml_models/modelo_curriculos.pkl")

    return accuracy, reporte


if __name__ == "__main__":
    acc_tramites, rep_tramites = entrenar_modelo_tramites()
    acc_cv, rep_cv = entrenar_modelo_curriculos()

    with open("ml_models/metricas_modelos.txt", "w", encoding="utf-8") as f:
        f.write("MÉTRICAS DEL MODELO DE TRÁMITES\n")
        f.write("================================\n")
        f.write(f"Accuracy: {acc_tramites:.2f}\n\n")
        f.write(rep_tramites)

        f.write("\n\nMÉTRICAS DEL MODELO DE CURRÍCULOS\n")
        f.write("=================================\n")
        f.write(f"Accuracy: {acc_cv:.2f}\n\n")
        f.write(rep_cv)

    print("Modelos entrenados correctamente.")
    print(f"Accuracy trámites: {acc_tramites:.2f}")
    print(f"Accuracy currículos: {acc_cv:.2f}")