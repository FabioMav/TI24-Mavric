# TI24 — Proyecto Final: Bagging: Random Forests
**Estudiante:** Fabio Mavric  
**Universidad del Valle — Cochabamba, Bolivia**  
**Materia:** Big Data / TI24  

---

## Tema: Bagging: Random Forests

Este proyecto implementa y analiza el algoritmo **Random Forest** (Bagging de Árboles de Decisión) aplicado al dataset de enfermedades cardiovasculares, comparándolo con un árbol de decisión simple.

---

## Dataset

**Cardiovascular Disease Dataset** — Kaggle  
- **Fuente:** https://www.kaggle.com/datasets/sulianova/cardiovascular-disease-dataset  
- **Filas:** 70,000 pacientes  
- **Columnas:** 12 variables + target  
- **Variables categóricas:** gender, cholesterol, gluc, smoke, alco, active  
- **Variables numéricas:** age, height, weight, ap_hi, ap_lo (+ derivadas: bmi, pulse_pressure)  
- **Target:** cardio (0 = sin enfermedad, 1 = con enfermedad)  

---

## Estructura del Repositorio

```
repositorio-TI24-Mavric/
├── README.md
├── requirements.txt
├── data/
│   ├── raw/
│   │   └── cardio_train.csv        ← COLOCAR AQUÍ EL DATASET
│   └── processed/
│       ├── X_train.csv
│       ├── X_test.csv
│       ├── y_train.csv
│       ├── y_test.csv
│       ├── rf_results.json
│       └── resultados_finales.csv
├── notebooks/
│   ├── 01_EDA.ipynb
│   ├── 02_preprocessing.ipynb
│   ├── 03_model_main.ipynb
│   └── 04_model_comparison.ipynb
└── docs/
    ├── informe_final.pdf
    └── presentacion.pdf
```

---

## Instrucciones de Ejecución

1. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   ```

2. Descargar el dataset de Kaggle y colocarlo en `data/raw/cardio_train.csv`

3. Ejecutar los notebooks **en orden**:
   - `01_EDA.ipynb` → Exploración de datos
   - `02_preprocessing.ipynb` → Limpieza y preprocesamiento
   - `03_model_main.ipynb` → Entrenamiento Random Forest
   - `04_model_comparison.ipynb` → Comparación RF vs Decision Tree

---

## Resultados Esperados

| Modelo | Accuracy | F1-Score | AUC-ROC |
|--------|----------|----------|---------|
| Decision Tree | ~71% | ~0.70 | ~0.71 |
| **Random Forest** | **~73%** | **~0.73** | **~0.80** |

---

## Referencias

1. Breiman, L. (2001). Random Forests. *Machine Learning, 45*(1), 5–32.
2. James, G., Witten, D., Hastie, T., & Tibshirani, R. (2021). *An Introduction to Statistical Learning* (2nd ed.). Springer.
3. Géron, A. (2022). *Hands-On Machine Learning with Scikit-Learn, Keras, and TensorFlow* (3rd ed.). O'Reilly Media.
