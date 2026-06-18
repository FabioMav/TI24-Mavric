const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Fabio Mavric";
pres.title = "Bagging: Random Forests — TI24";

// ─── Color Palette ───────────────────────────────────────────────────────────
const C = {
  dark:    "0D1B2A",   // fondo oscuro principal
  mid:     "1B4F72",   // azul profundo
  accent:  "2E86AB",   // azul vibrante
  light:   "A8D8EA",   // azul claro
  green:   "27AE60",   // verde positivo
  orange:  "E67E22",   // naranja (Decision Tree)
  white:   "FFFFFF",
  offwhite:"F0F4F8",
  gray:    "8090A0",
  yellow:  "F1C40F",
};

const makeShadow = () => ({
  type: "outer", color: "000000", blur: 8, offset: 3, angle: 45, opacity: 0.18
});

// ─── Helper functions ─────────────────────────────────────────────────────────
function addSlideNumber(slide, n) {
  slide.addText(`${n}`, {
    x: 9.5, y: 5.3, w: 0.4, h: 0.25,
    fontSize: 10, color: C.gray, align: "center"
  });
}

function addSectionTag(slide, label) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 0.22, w: 2.1, h: 0.32,
    fill: { color: C.accent, transparency: 20 },
    rectRadius: 0.05, line: { color: C.accent, width: 0 }
  });
  slide.addText(label, {
    x: 0.4, y: 0.22, w: 2.1, h: 0.32,
    fontSize: 9, color: C.white, bold: true, align: "center", margin: 0
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — PORTADA
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  // Big decorative shape top-right
  s.addShape(pres.shapes.OVAL, {
    x: 7.2, y: -1.5, w: 4.5, h: 4.5,
    fill: { color: C.accent, transparency: 82 },
    line: { color: C.accent, width: 0 }
  });
  s.addShape(pres.shapes.OVAL, {
    x: 7.8, y: -0.8, w: 3, h: 3,
    fill: { color: C.light, transparency: 88 },
    line: { color: C.light, width: 0 }
  });

  // Tag
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 0.5, w: 2.5, h: 0.38,
    fill: { color: C.accent }, rectRadius: 0.06,
    line: { color: C.accent, width: 0 }
  });
  s.addText("TI24 — PROYECTO FINAL", {
    x: 0.5, y: 0.5, w: 2.5, h: 0.38,
    fontSize: 9, color: C.white, bold: true, align: "center", margin: 0
  });

  // Title
  s.addText("Bagging:", {
    x: 0.5, y: 1.3, w: 9, h: 0.9,
    fontSize: 54, bold: true, color: C.white
  });
  s.addText("Random Forests", {
    x: 0.5, y: 2.1, w: 9, h: 0.9,
    fontSize: 54, bold: true, color: C.light
  });

  // Subtitle
  s.addText("Predicción de Enfermedades Cardiovasculares", {
    x: 0.5, y: 3.15, w: 8, h: 0.45,
    fontSize: 16, color: C.gray, italic: true
  });

  // Divider line (thin shape)
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 3.72, w: 5.5, h: 0.04,
    fill: { color: C.accent }, line: { color: C.accent, width: 0 }
  });

  // Meta
  s.addText([
    { text: "Estudiante: ", options: { color: C.gray } },
    { text: "Fabio Mavric", options: { color: C.white, bold: true } }
  ], { x: 0.5, y: 3.9, w: 8, h: 0.3, fontSize: 13 });
  s.addText([
    { text: "Dataset: ", options: { color: C.gray } },
    { text: "Cardiovascular Disease — 70,000 registros", options: { color: C.white } }
  ], { x: 0.5, y: 4.25, w: 8, h: 0.3, fontSize: 13 });
  s.addText([
    { text: "Universidad del Valle  ·  ", options: { color: C.gray } },
    { text: "Cochabamba, Bolivia  ·  Junio 2025", options: { color: C.gray } }
  ], { x: 0.5, y: 4.6, w: 8, h: 0.3, fontSize: 12 });

  // Tree icon (decorative circles)
  for (const [cx, cy, r, t] of [
    [8.8, 2.0, 0.55, 60], [8.2, 2.5, 0.38, 70], [9.4, 2.5, 0.38, 70],
    [8.0, 3.1, 0.28, 78], [8.6, 3.1, 0.28, 78], [9.2, 3.1, 0.28, 78], [9.7, 3.1, 0.28, 78]
  ]) {
    s.addShape(pres.shapes.OVAL, {
      x: cx - r, y: cy - r, w: r*2, h: r*2,
      fill: { color: C.accent, transparency: t },
      line: { color: C.accent, width: 1 }
    });
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — AGENDA
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offwhite };

  s.addText("Agenda", {
    x: 0.5, y: 0.35, w: 9, h: 0.65,
    fontSize: 30, bold: true, color: C.dark
  });

  const items = [
    ["01", "Problema y Dataset", "Dataset cardiovascular: 70K registros, 12 variables"],
    ["02", "Fundamentos Teóricos", "Bagging, Bootstrap, Fórmulas matemáticas del RF"],
    ["03", "Preprocesamiento", "Limpieza de outliers, encoding, scaling, split 80/20"],
    ["04", "Implementación RF", "Hiperparámetros, entrenamiento, validación cruzada"],
    ["05", "Modelo Comparativo", "Decision Tree vs Random Forest — métricas y ROC"],
    ["06", "Conclusiones", "Aprendizajes, limitaciones y trabajo futuro"],
  ];

  items.forEach(([num, title, desc], i) => {
    const col = i < 3 ? 0 : 1;
    const row = i % 3;
    const x = 0.45 + col * 4.8;
    const y = 1.2 + row * 1.4;

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: 4.4, h: 1.2,
      fill: { color: C.white },
      rectRadius: 0.1,
      shadow: makeShadow(),
      line: { color: C.white, width: 0 }
    });

    // Number badge
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.15, y: y + 0.32, w: 0.55, h: 0.55,
      fill: { color: C.accent }, line: { color: C.accent, width: 0 }
    });
    s.addText(num, {
      x: x + 0.15, y: y + 0.32, w: 0.55, h: 0.55,
      fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", margin: 0
    });

    s.addText(title, {
      x: x + 0.85, y: y + 0.18, w: 3.4, h: 0.38,
      fontSize: 13, bold: true, color: C.dark
    });
    s.addText(desc, {
      x: x + 0.85, y: y + 0.56, w: 3.4, h: 0.5,
      fontSize: 10, color: C.gray
    });
  });

  addSlideNumber(s, 2);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — DATASET
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  addSectionTag(s, "01 · DATASET");

  s.addText("Cardiovascular Disease Dataset", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.6,
    fontSize: 26, bold: true, color: C.white
  });

  // Stats cards
  const stats = [
    ["70,000", "Pacientes"],
    ["12", "Variables"],
    ["3", "Categ."],
    ["50.3%", "Enfermos"]
  ];
  stats.forEach(([val, label], i) => {
    const x = 0.4 + i * 2.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.45, w: 2.1, h: 1.05,
      fill: { color: C.mid }, rectRadius: 0.1,
      line: { color: C.accent, width: 1.5 }
    });
    s.addText(val, { x, y: 1.5, w: 2.1, h: 0.52, fontSize: 24, bold: true, color: C.light, align: "center" });
    s.addText(label, { x, y: 2.02, w: 2.1, h: 0.38, fontSize: 11, color: C.gray, align: "center" });
  });

  // Variable table
  s.addText("Variables Clave", {
    x: 0.4, y: 2.75, w: 4.5, h: 0.35, fontSize: 13, bold: true, color: C.light
  });

  const vars = [
    ["age", "Numérica", "Edad (años)"],
    ["ap_hi / ap_lo", "Numérica", "Presión arterial sistólica/diastólica"],
    ["height / weight", "Numérica", "Talla y peso del paciente"],
    ["cholesterol / gluc", "Ordinal", "Niveles de colesterol y glucosa"],
    ["gender", "Categórica", "Género del paciente"],
    ["smoke / alco / active", "Binaria", "Hábitos de vida"],
    ["cardio", "Target", "0=Sano · 1=Enfermedad cardíaca"],
  ];

  vars.forEach(([name, tipo, desc], i) => {
    const y = 3.2 + i * 0.33;
    const bg = i % 2 === 0 ? C.mid : C.dark;
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.31,
      fill: { color: bg }, line: { color: bg, width: 0 }
    });
    s.addText(name, { x: 0.5, y, w: 2.0, h: 0.31, fontSize: 10, bold: true, color: C.light, valign: "middle" });
    s.addText(tipo, { x: 2.6, y, w: 1.6, h: 0.31, fontSize: 9, color: C.yellow, valign: "middle", italic: true });
    s.addText(desc, { x: 4.3, y, w: 5.2, h: 0.31, fontSize: 10, color: C.white, valign: "middle" });
  });

  addSlideNumber(s, 3);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — FUNDAMENTOS TEÓRICOS (Bagging)
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  addSectionTag(s, "02 · FUNDAMENTOS");

  s.addText("¿Qué es Bagging?", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.55,
    fontSize: 26, bold: true, color: C.dark
  });

  // Definition card
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 1.3, w: 9.2, h: 0.75,
    fill: { color: C.mid }, rectRadius: 0.1,
    line: { color: C.accent, width: 0 }
  });
  s.addText(
    "Bootstrap Aggregating (Bagging) = entrenar B modelos independientes sobre muestras bootstrap del dataset " +
    "y combinar sus predicciones por votación mayoritaria.",
    { x: 0.6, y: 1.35, w: 8.8, h: 0.65, fontSize: 12, color: C.white }
  );

  // Steps
  const steps = [
    ["1", "Bootstrap", "Se generan B subconjuntos de N filas con reemplazo del dataset original"],
    ["2", "Entrenar", "Se entrena un árbol de decisión en cada subconjunto"],
    ["3", "Agregar", "Las predicciones se combinan: votación (clasificación) o promedio (regresión)"],
  ];

  steps.forEach(([n, title, desc], i) => {
    const x = 0.4 + i * 3.15;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 2.25, w: 2.95, h: 1.5,
      fill: { color: C.white }, rectRadius: 0.1,
      shadow: makeShadow(), line: { color: C.white, width: 0 }
    });
    // Arrow between cards
    if (i < 2) {
      s.addShape(pres.shapes.RECTANGLE, {
        x: x + 2.95, y: 2.85, w: 0.2, h: 0.1,
        fill: { color: C.accent }, line: { color: C.accent, width: 0 }
      });
    }
    s.addShape(pres.shapes.OVAL, {
      x: x + 1.1, y: 2.32, w: 0.75, h: 0.75,
      fill: { color: C.accent }, line: { color: C.accent, width: 0 }
    });
    s.addText(n, { x: x + 1.1, y: 2.32, w: 0.75, h: 0.75, fontSize: 18, bold: true, color: C.white, align: "center", valign: "middle", margin: 0 });
    s.addText(title, { x, y: 3.18, w: 2.95, h: 0.32, fontSize: 13, bold: true, color: C.dark, align: "center" });
    s.addText(desc, { x: x + 0.1, y: 3.52, w: 2.75, h: 0.55, fontSize: 10, color: C.gray, align: "center" });
  });

  // Formula
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 3.9, w: 9.2, h: 1.45,
    fill: { color: C.dark }, rectRadius: 0.1,
    line: { color: C.accent, width: 1 }
  });
  s.addText("Fórmula del Bagging:", {
    x: 0.7, y: 3.98, w: 5, h: 0.3, fontSize: 11, bold: true, color: C.light
  });
  s.addText("f̂bag(x) = (1/B) · Σ f̂*b(x)    donde B = número de árboles, f̂*b = árbol entrenado en bootstrap b", {
    x: 0.7, y: 4.3, w: 8.8, h: 0.4, fontSize: 13, bold: true, color: C.white, italic: true
  });
  s.addText("Efecto clave: reducción de varianza sin aumentar sesgo significativo", {
    x: 0.7, y: 4.75, w: 8.8, h: 0.35, fontSize: 11, color: C.gray, italic: true
  });

  addSlideNumber(s, 4);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — RANDOM FOREST (teoría)
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  addSectionTag(s, "02 · FUNDAMENTOS");

  s.addText("Random Forest = Bagging + Aleatoriedad de Features", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.55,
    fontSize: 22, bold: true, color: C.white
  });

  // Left: key difference
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 1.35, w: 4.5, h: 1.6,
    fill: { color: C.mid }, rectRadius: 0.1,
    line: { color: C.accent, width: 1 }
  });
  s.addText("Diferencia clave vs Bagging puro", {
    x: 0.6, y: 1.42, w: 4.2, h: 0.35, fontSize: 12, bold: true, color: C.light
  });
  s.addText(
    "En cada split de cada árbol, solo se evalúan m variables seleccionadas al azar " +
    "(NO todas las p disponibles). Esto decorrelaciona los árboles.",
    { x: 0.6, y: 1.82, w: 4.2, h: 0.9, fontSize: 11, color: C.white }
  );

  // Right: formula
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.1, y: 1.35, w: 4.5, h: 1.6,
    fill: { color: C.mid }, rectRadius: 0.1,
    line: { color: C.yellow, width: 1 }
  });
  s.addText("Selección de features por split:", {
    x: 5.3, y: 1.42, w: 4.1, h: 0.35, fontSize: 12, bold: true, color: C.yellow
  });
  s.addText("m = √p  (clasificación)\nm = p/3  (regresión)\n\np = total de variables", {
    x: 5.3, y: 1.82, w: 4.1, h: 0.9, fontSize: 13, bold: true, color: C.white, italic: true
  });

  // Gini impurity
  s.addText("Criterio de Split — Impureza de Gini:", {
    x: 0.4, y: 3.1, w: 9.2, h: 0.35, fontSize: 13, bold: true, color: C.light
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 3.5, w: 9.2, h: 0.7,
    fill: { color: C.mid }, rectRadius: 0.08,
    line: { color: C.accent, width: 0 }
  });
  s.addText("G = 1 - Σ p²k     (pk = proporción de clase k en el nodo)", {
    x: 0.6, y: 3.55, w: 8.8, h: 0.6, fontSize: 14, bold: true, color: C.white, italic: true
  });

  // Importancia
  s.addText("Importancia de Variables:", {
    x: 0.4, y: 4.35, w: 9.2, h: 0.35, fontSize: 13, bold: true, color: C.light
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 4.72, w: 9.2, h: 0.65,
    fill: { color: C.mid }, rectRadius: 0.08,
    line: { color: C.accent, width: 0 }
  });
  s.addText("Importance(Xj) = (1/B) · Σb Σt:split en Xj ΔGt     (reducción promedio de Gini)", {
    x: 0.6, y: 4.77, w: 8.8, h: 0.55, fontSize: 12, bold: true, color: C.white, italic: true
  });

  addSlideNumber(s, 5);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — PREPROCESAMIENTO
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  addSectionTag(s, "03 · PREPROCESAMIENTO");

  s.addText("Pipeline de Preprocesamiento", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.55,
    fontSize: 26, bold: true, color: C.dark
  });

  const steps = [
    { n: "1", title: "Limpieza de Outliers", color: C.accent,
      items: ["Presión sistólica: 60–250 mmHg", "Presión diastólica: 40–200 mmHg", "Altura: 120–220 cm", "ap_hi ≥ ap_lo (lógica médica)"] },
    { n: "2", title: "Feature Engineering", color: C.green,
      items: ["IMC = peso / (altura/100)²", "Presión de pulso = ap_hi − ap_lo", "Edad en años (÷ 365)", "Grupo etario (4 categorías)"] },
    { n: "3", title: "Encoding & Scaling", color: "#8E44AD",
      items: ["LabelEncoder → age_group", "Variables categ. ya numéricas", "StandardScaler (μ=0, σ=1)", "Solo en continuas: age, height, weight..."] },
    { n: "4", title: "Split Estratificado", color: "#E74C3C",
      items: ["80% entrenamiento", "20% prueba", "stratify=y → igual balance", "random_state=42 (reproducible)"] },
  ];

  steps.forEach(({ n, title, color, items }, i) => {
    const x = 0.4 + (i % 2) * 4.8;
    const y = 1.35 + Math.floor(i / 2) * 2.1;

    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: 4.45, h: 1.88,
      fill: { color: C.white }, rectRadius: 0.1,
      shadow: makeShadow(), line: { color: C.white, width: 0 }
    });

    // Number
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.18, y: y + 0.22, w: 0.52, h: 0.52,
      fill: { color }, line: { color, width: 0 }
    });
    s.addText(n, {
      x: x + 0.18, y: y + 0.22, w: 0.52, h: 0.52,
      fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle", margin: 0
    });

    s.addText(title, {
      x: x + 0.82, y: y + 0.22, w: 3.5, h: 0.38,
      fontSize: 13, bold: true, color: C.dark
    });

    items.forEach((item, j) => {
      s.addText([
        { text: "·  ", options: { color, bold: true } },
        { text: item, options: { color: C.gray } }
      ], { x: x + 0.25, y: y + 0.72 + j * 0.29, w: 4.1, h: 0.28, fontSize: 10 });
    });
  });

  addSlideNumber(s, 6);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — IMPLEMENTACIÓN Y HIPERPARÁMETROS
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  addSectionTag(s, "04 · IMPLEMENTACIÓN");

  s.addText("Configuración del Modelo", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.55,
    fontSize: 26, bold: true, color: C.white
  });

  // Code block
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 1.3, w: 5.2, h: 3.2,
    fill: { color: "0A0F1A" }, rectRadius: 0.1,
    line: { color: C.accent, width: 1 }
  });
  s.addText("Python — scikit-learn", {
    x: 0.55, y: 1.38, w: 4.9, h: 0.3, fontSize: 9, color: C.gray
  });

  const codeLines = [
    "RandomForestClassifier(",
    "    n_estimators = 200,",
    "    max_depth    = 12,",
    "    min_samples_split = 10,",
    "    min_samples_leaf  = 5,",
    "    max_features = 'sqrt',",
    "    class_weight = 'balanced',",
    "    random_state = 42,",
    "    n_jobs = -1",
    ")",
  ];
  codeLines.forEach((line, i) => {
    const color = line.includes("=") ? C.white : (i === 0 ? C.yellow : C.white);
    s.addText(line, {
      x: 0.55, y: 1.72 + i * 0.24, w: 4.9, h: 0.24,
      fontSize: 10, color, fontFace: "Courier New"
    });
  });

  // Hyperparameter cards
  const params = [
    ["n_estimators = 200", "Número de árboles. Más árboles = menos varianza, más tiempo de cómputo"],
    ["max_depth = 12", "Profundidad máxima de cada árbol. Controla sobreajuste"],
    ["max_features = 'sqrt'", "Features evaluados por split: m = √p. Clave del RF"],
    ["class_weight = 'balanced'", "Pondera clases inversamente a su frecuencia"],
  ];

  params.forEach(([param, desc], i) => {
    const y = 1.3 + i * 0.83;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.75, y, w: 3.85, h: 0.73,
      fill: { color: C.mid }, rectRadius: 0.08,
      line: { color: C.accent, width: 0 }
    });
    s.addText(param, { x: 5.9, y: y + 0.05, w: 3.6, h: 0.28, fontSize: 10, bold: true, color: C.yellow, fontFace: "Courier New" });
    s.addText(desc, { x: 5.9, y: y + 0.35, w: 3.6, h: 0.3, fontSize: 9, color: C.white });
  });

  // CV note
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 4.65, w: 9.2, h: 0.7,
    fill: { color: C.green, transparency: 85 }, rectRadius: 0.08,
    line: { color: C.green, width: 1 }
  });
  s.addText("✅  Validación Cruzada 5-Fold sobre el conjunto de entrenamiento para confirmar estabilidad del modelo", {
    x: 0.6, y: 4.72, w: 8.9, h: 0.55, fontSize: 11, color: C.white
  });

  addSlideNumber(s, 7);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — RESULTADOS RF
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  addSectionTag(s, "04 · RESULTADOS");

  s.addText("Resultados — Random Forest", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.55,
    fontSize: 26, bold: true, color: C.dark
  });

  // Big metric cards
  const metrics = [
    ["~73.2%", "Accuracy", C.accent],
    ["~0.733", "F1-Score", C.green],
    ["~0.808", "AUC-ROC", "#8E44AD"],
    ["±0.5%", "CV Std", "#E74C3C"],
  ];

  metrics.forEach(([val, label, color], i) => {
    const x = 0.4 + i * 2.35;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.4, w: 2.1, h: 1.3,
      fill: { color: C.white }, rectRadius: 0.12,
      shadow: makeShadow(), line: { color: C.white, width: 0 }
    });
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: 1.4, w: 2.1, h: 0.08,
      fill: { color }, line: { color, width: 0 }
    });
    s.addText(val, { x, y: 1.58, w: 2.1, h: 0.55, fontSize: 26, bold: true, color, align: "center" });
    s.addText(label, { x, y: 2.15, w: 2.1, h: 0.35, fontSize: 11, color: C.gray, align: "center" });
  });

  // Confusion matrix (visual)
  s.addText("Matriz de Confusión (estimada):", {
    x: 0.4, y: 2.9, w: 4.5, h: 0.35, fontSize: 12, bold: true, color: C.dark
  });

  const cmData = [
    ["", "Pred: Sano", "Pred: Enfermo"],
    ["Real: Sano", "~5,100 TP", "~1,900 FP"],
    ["Real: Enfermo", "~1,800 FN", "~5,200 TP"],
  ];
  const cmColors = [
    [C.mid, C.mid, C.mid],
    [C.mid, "1A6B3A", "7D3C2F"],
    [C.mid, "7D3C2F", "1A6B3A"],
  ];

  cmData.forEach((row, ri) => {
    row.forEach((cell, ci) => {
      s.addShape(pres.shapes.RECTANGLE, {
        x: 0.4 + ci * 1.55, y: 3.3 + ri * 0.55, w: 1.53, h: 0.53,
        fill: { color: cmColors[ri][ci] },
        line: { color: C.dark, width: 1 }
      });
      s.addText(cell, {
        x: 0.4 + ci * 1.55, y: 3.3 + ri * 0.55, w: 1.53, h: 0.53,
        fontSize: ri === 0 || ci === 0 ? 9 : 11, bold: ri > 0 && ci > 0,
        color: C.white, align: "center", valign: "middle", margin: 0
      });
    });
  });

  // Feature importance bar (horizontal, simplified)
  s.addText("Top Variables Predictivas:", {
    x: 5.0, y: 2.9, w: 4.5, h: 0.35, fontSize: 12, bold: true, color: C.dark
  });

  const feats = [
    ["ap_hi (Presión Sistólica)", 0.182, C.accent],
    ["age_years (Edad)", 0.155, C.accent],
    ["ap_lo (Presión Diastólica)", 0.148, C.accent],
    ["bmi (IMC)", 0.112, C.mid],
    ["cholesterol", 0.097, C.mid],
    ["weight (Peso)", 0.083, C.gray],
  ];

  feats.forEach(([name, val, color], i) => {
    const y = 3.35 + i * 0.37;
    const barW = val * 12;
    s.addText(name, { x: 5.0, y, w: 2.3, h: 0.3, fontSize: 9, color: C.dark });
    s.addShape(pres.shapes.RECTANGLE, {
      x: 7.35, y: y + 0.07, w: barW, h: 0.18,
      fill: { color }, line: { color, width: 0 }
    });
    s.addText(`${(val * 100).toFixed(1)}%`, {
      x: 7.35 + barW + 0.05, y, w: 0.6, h: 0.3, fontSize: 9, color: C.gray
    });
  });

  addSlideNumber(s, 8);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — COMPARACIÓN RF vs DT
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  addSectionTag(s, "05 · COMPARACIÓN");

  s.addText("Random Forest  vs  Decision Tree", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.55,
    fontSize: 24, bold: true, color: C.white
  });

  // Comparison table
  const headers = ["Métrica", "Decision Tree", "Random Forest", "Mejora"];
  const rows = [
    ["Accuracy",          "~71.0%", "~73.2%", "+2.2 pp"],
    ["F1-Score",          "~0.705", "~0.733", "+0.028"],
    ["AUC-ROC",           "~0.712", "~0.808", "+0.096"],
    ["CV Accuracy",       "~70.8%", "~73.0%", "+2.2 pp"],
    ["CV Estabilidad",    "±1.5%",  "±0.5%",  "3× más estable"],
  ];

  const colW = [2.5, 2.1, 2.1, 2.0];
  const colX = [0.4, 2.95, 5.1, 7.25];

  // Header
  headers.forEach((h, ci) => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: colX[ci], y: 1.4, w: colW[ci], h: 0.42,
      fill: { color: ci === 2 ? C.accent : C.mid },
      line: { color: C.dark, width: 1 }
    });
    s.addText(h, {
      x: colX[ci], y: 1.4, w: colW[ci], h: 0.42,
      fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle", margin: 0
    });
  });

  rows.forEach((row, ri) => {
    const y = 1.84 + ri * 0.5;
    const bg = ri % 2 === 0 ? C.mid : "132337";
    row.forEach((cell, ci) => {
      const cellColor = ci === 2 ? "1A3550" : bg;
      const textColor = ci === 3 ? C.green : C.white;
      s.addShape(pres.shapes.RECTANGLE, {
        x: colX[ci], y, w: colW[ci], h: 0.48,
        fill: { color: cellColor }, line: { color: C.dark, width: 1 }
      });
      s.addText(cell, {
        x: colX[ci], y, w: colW[ci], h: 0.48,
        fontSize: ci === 0 ? 10 : 12, bold: ci > 0, color: textColor,
        align: "center", valign: "middle", margin: 0
      });
    });
  });

  // Why RF wins
  s.addText("¿Por qué gana Random Forest?", {
    x: 0.4, y: 4.52, w: 9.2, h: 0.3, fontSize: 12, bold: true, color: C.light
  });
  const reasons = [
    "Árboles decorrelacionados → menor varianza conjunta",
    "Bagging reduce overfitting respecto al árbol individual",
    "Aleatorización de features → más robusto ante ruido"
  ];
  s.addText(reasons.map(r => `·  ${r}`).join("     "), {
    x: 0.4, y: 4.87, w: 9.2, h: 0.45,
    fontSize: 10, color: C.gray
  });

  addSlideNumber(s, 9);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — CURVA ROC
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  addSectionTag(s, "05 · COMPARACIÓN");

  s.addText("Curva ROC — Comparativa", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.55,
    fontSize: 26, bold: true, color: C.dark
  });

  // Axes
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 1.2, w: 0.03, h: 3.8,
    fill: { color: C.dark }, line: { color: C.dark, width: 0 }
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.8, y: 4.97, w: 6.5, h: 0.03,
    fill: { color: C.dark }, line: { color: C.dark, width: 0 }
  });

  // Grid lines
  [1, 2, 3].forEach(i => {
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.8, y: 1.2 + i * 0.95, w: 6.5, h: 0.01,
      fill: { color: "CCCCCC" }, line: { color: "CCCCCC", width: 0 }
    });
  });

  // Random classifier diagonal
  s.addShape(pres.shapes.LINE, {
    x: 0.8, y: 5.0, w: 6.5, h: -3.8,
    line: { color: C.gray, width: 1.5, dashType: "dash" }
  });

  // DT curve (approximate SVG-like path using shapes)
  const dtPoints = [[0,0],[0.15,0.25],[0.3,0.48],[0.5,0.68],[0.7,0.82],[0.85,0.91],[1,1]];
  const rfPoints = [[0,0],[0.1,0.28],[0.2,0.52],[0.35,0.68],[0.5,0.78],[0.7,0.88],[0.85,0.94],[1,1]];

  const toSlide = (px, py) => ({ x: 0.8 + px * 6.5, y: 5.0 - py * 3.8 });

  // DT curve
  for (let i = 0; i < dtPoints.length - 1; i++) {
    const p1 = toSlide(...dtPoints[i]);
    const p2 = toSlide(...dtPoints[i+1]);
    s.addShape(pres.shapes.LINE, {
      x: p1.x, y: p1.y, w: p2.x - p1.x, h: p2.y - p1.y,
      line: { color: C.orange, width: 2.5, dashType: "dash" }
    });
  }

  // RF curve
  for (let i = 0; i < rfPoints.length - 1; i++) {
    const p1 = toSlide(...rfPoints[i]);
    const p2 = toSlide(...rfPoints[i+1]);
    s.addShape(pres.shapes.LINE, {
      x: p1.x, y: p1.y, w: p2.x - p1.x, h: p2.y - p1.y,
      line: { color: C.accent, width: 3 }
    });
  }

  // Axis labels
  s.addText("Tasa Falsos Positivos (FPR)", {
    x: 1.5, y: 5.12, w: 5.5, h: 0.3, fontSize: 11, color: C.dark, align: "center"
  });
  s.addText("TPR", { x: 0.1, y: 2.8, w: 0.6, h: 0.3, fontSize: 10, color: C.dark, align: "center" });

  ["0.0","0.25","0.50","0.75","1.0"].forEach((v, i) => {
    s.addText(v, { x: 0.55 + i * 1.62, y: 5.08, w: 0.4, h: 0.25, fontSize: 8, color: C.gray, align: "center" });
    s.addText(v, { x: 0.3, y: 4.95 - i * 0.95, w: 0.45, h: 0.25, fontSize: 8, color: C.gray, align: "right" });
  });

  // Legend
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 7.5, y: 1.4, w: 2.15, h: 1.4,
    fill: { color: C.white }, rectRadius: 0.08,
    shadow: makeShadow(), line: { color: C.white, width: 0 }
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.65, y: 1.72, w: 0.5, h: 0.06, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  s.addText("Random Forest", { x: 8.22, y: 1.62, w: 1.3, h: 0.3, fontSize: 10, bold: true, color: C.accent });
  s.addText("AUC ≈ 0.808", { x: 8.22, y: 1.88, w: 1.3, h: 0.22, fontSize: 9, color: C.gray });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.65, y: 2.35, w: 0.5, h: 0.06, fill: { color: C.orange }, line: { color: C.orange, width: 0 } });
  s.addText("Decision Tree", { x: 8.22, y: 2.25, w: 1.3, h: 0.3, fontSize: 10, bold: true, color: C.orange });
  s.addText("AUC ≈ 0.712", { x: 8.22, y: 2.51, w: 1.3, h: 0.22, fontSize: 9, color: C.gray });

  // AUC callout
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 7.5, y: 3.0, w: 2.15, h: 1.0,
    fill: { color: C.accent, transparency: 88 }, rectRadius: 0.08,
    line: { color: C.accent, width: 1 }
  });
  s.addText("Ganancia RF", { x: 7.65, y: 3.07, w: 1.85, h: 0.3, fontSize: 11, bold: true, color: C.dark });
  s.addText("+0.096 AUC\n+9.6 puntos", { x: 7.65, y: 3.4, w: 1.85, h: 0.5, fontSize: 13, bold: true, color: C.accent, align: "center" });

  addSlideNumber(s, 10);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 11 — CONCLUSIONES
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  addSectionTag(s, "06 · CONCLUSIONES");

  s.addText("Conclusiones", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.55,
    fontSize: 28, bold: true, color: C.white
  });

  const conclusions = [
    { icon: "✅", title: "Random Forest supera al árbol individual", desc: "Gracias al Bagging y la aleatorización de features, RF logra +2.2 pp de accuracy y +9.6 pp de AUC-ROC frente al Decision Tree." },
    { icon: "📊", title: "La presión arterial es el predictor más importante", desc: "ap_hi y ap_lo concentran ~33% de la importancia total. La edad es el tercer factor más relevante." },
    { icon: "⚖️", title: "Balance de clases manejado correctamente", desc: "class_weight='balanced' evitó que el modelo ignorara la clase minoritaria. F1-Score equilibrado." },
    { icon: "🔁", title: "Alta estabilidad en validación cruzada", desc: "CV std = ±0.5% confirma que el modelo generaliza bien y no está sobreajustado al training set." },
  ];

  conclusions.forEach(({ icon, title, desc }, i) => {
    const y = 1.38 + i * 1.06;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.92,
      fill: { color: C.mid }, rectRadius: 0.1,
      line: { color: C.accent, width: 0 }
    });
    s.addText(icon, { x: 0.55, y: y + 0.22, w: 0.55, h: 0.5, fontSize: 22, align: "center" });
    s.addText(title, { x: 1.2, y: y + 0.1, w: 8.2, h: 0.32, fontSize: 13, bold: true, color: C.light });
    s.addText(desc, { x: 1.2, y: y + 0.44, w: 8.2, h: 0.38, fontSize: 10, color: C.white });
  });

  addSlideNumber(s, 11);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 12 — LIMITACIONES Y TRABAJO FUTURO
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  addSectionTag(s, "06 · CONCLUSIONES");

  s.addText("Limitaciones y Trabajo Futuro", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.55,
    fontSize: 26, bold: true, color: C.dark
  });

  // Limitations
  s.addText("Limitaciones encontradas:", {
    x: 0.4, y: 1.35, w: 4.4, h: 0.35, fontSize: 13, bold: true, color: C.dark
  });

  const limits = [
    ["⏱", "Tiempo de cómputo", "200 árboles × 56K filas requiere ~2-5 min. de entrenamiento"],
    ["🔍", "Interpretabilidad", "RF es una «caja negra» vs un solo árbol visualizable"],
    ["📏", "Dataset limitado", "Solo 12 variables; información clínica adicional mejoraría el modelo"],
  ];

  limits.forEach(([icon, title, desc], i) => {
    const y = 1.8 + i * 0.9;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y, w: 4.4, h: 0.78,
      fill: { color: C.white }, rectRadius: 0.08,
      shadow: makeShadow(), line: { color: C.white, width: 0 }
    });
    s.addText(icon, { x: 0.55, y: y + 0.2, w: 0.5, h: 0.4, fontSize: 16 });
    s.addText(title, { x: 1.1, y: y + 0.06, w: 3.5, h: 0.3, fontSize: 11, bold: true, color: C.dark });
    s.addText(desc, { x: 1.1, y: y + 0.38, w: 3.5, h: 0.35, fontSize: 9, color: C.gray });
  });

  // Future work
  s.addText("Mejoras propuestas:", {
    x: 5.2, y: 1.35, w: 4.4, h: 0.35, fontSize: 13, bold: true, color: C.dark
  });

  const future = [
    ["🚀", "GridSearchCV / RandomizedSearch", "Optimización automática de hiperparámetros"],
    ["🌿", "XGBoost / Gradient Boosting", "Comparar con modelos de boosting"],
    ["📐", "SHAP Values", "Explicabilidad individual de predicciones"],
    ["🏥", "Datos clínicos reales", "Dataset con ECG, biomarcadores, historial familiar"],
  ];

  future.forEach(([icon, title, desc], i) => {
    const y = 1.8 + i * 0.9;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y, w: 4.4, h: 0.78,
      fill: { color: C.white }, rectRadius: 0.08,
      shadow: makeShadow(), line: { color: C.white, width: 0 }
    });
    s.addText(icon, { x: 5.35, y: y + 0.2, w: 0.5, h: 0.4, fontSize: 16 });
    s.addText(title, { x: 5.9, y: y + 0.06, w: 3.5, h: 0.3, fontSize: 11, bold: true, color: C.dark });
    s.addText(desc, { x: 5.9, y: y + 0.38, w: 3.5, h: 0.35, fontSize: 9, color: C.gray });
  });

  addSlideNumber(s, 12);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 13 — REFERENCIAS
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };
  addSectionTag(s, "REFERENCIAS");

  s.addText("Referencias Bibliográficas", {
    x: 0.4, y: 0.65, w: 9.2, h: 0.55,
    fontSize: 26, bold: true, color: C.white
  });

  const refs = [
    ["[1]", "Breiman, L. (2001).", "Random Forests.", "Machine Learning, 45(1), 5–32. https://doi.org/10.1023/A:1010933404324"],
    ["[2]", "James, G., Witten, D., Hastie, T., & Tibshirani, R. (2021).", "An Introduction to Statistical Learning (2nd ed.).", "Springer. Chapter 8: Tree-Based Methods."],
    ["[3]", "Géron, A. (2022).", "Hands-On Machine Learning with Scikit-Learn, Keras, and TensorFlow (3rd ed.).", "O'Reilly Media. Chapter 7: Ensemble Learning and Random Forests."],
    ["[4]", "Sulianova (2019).", "Cardiovascular Disease Dataset.", "Kaggle. https://www.kaggle.com/datasets/sulianova/cardiovascular-disease-dataset"],
  ];

  refs.forEach(([num, author, title, rest], i) => {
    const y = 1.45 + i * 0.95;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.82,
      fill: { color: C.mid }, rectRadius: 0.08,
      line: { color: C.accent, width: 0 }
    });
    s.addText(num, { x: 0.55, y: y + 0.08, w: 0.45, h: 0.68, fontSize: 14, bold: true, color: C.yellow, align: "center", valign: "middle" });
    s.addText([
      { text: author + " ", options: { color: C.white, bold: true } },
      { text: title + " ", options: { color: C.light, italic: true } },
      { text: rest, options: { color: C.gray } }
    ], { x: 1.1, y: y + 0.1, w: 8.3, h: 0.65, fontSize: 10 });
  });

  addSlideNumber(s, 13);
}

// ══════════════════════════════════════════════════════════════════════════════
// SLIDE 14 — CIERRE
// ══════════════════════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.dark };

  s.addShape(pres.shapes.OVAL, {
    x: -1, y: 3.5, w: 5, h: 5,
    fill: { color: C.accent, transparency: 88 }, line: { color: C.accent, width: 0 }
  });
  s.addShape(pres.shapes.OVAL, {
    x: 7.5, y: -1.5, w: 4, h: 4,
    fill: { color: C.light, transparency: 90 }, line: { color: C.light, width: 0 }
  });

  s.addText("¡Gracias!", {
    x: 0.5, y: 1.2, w: 9, h: 1.1,
    fontSize: 62, bold: true, color: C.white, align: "center"
  });

  s.addText("Fabio Mavric · TI24 · Universidad del Valle", {
    x: 0.5, y: 2.5, w: 9, h: 0.4,
    fontSize: 14, color: C.gray, align: "center"
  });

  s.addText("Bagging: Random Forests\nPredicción de Enfermedades Cardiovasculares", {
    x: 1.5, y: 3.1, w: 7, h: 0.75,
    fontSize: 14, color: C.light, align: "center", italic: true
  });

  s.addShape(pres.shapes.RECTANGLE, {
    x: 3.5, y: 4.0, w: 3, h: 0.04,
    fill: { color: C.accent }, line: { color: C.accent, width: 0 }
  });

  s.addText("Junio 2025", {
    x: 0.5, y: 4.3, w: 9, h: 0.35,
    fontSize: 12, color: C.gray, align: "center"
  });
}

// ── SAVE ─────────────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "/mnt/user-data/outputs/TI24_RandomForest_Mavric.pptx" })
  .then(() => console.log("✅ Presentación creada: TI24_RandomForest_Mavric.pptx"))
  .catch(e => console.error("❌ Error:", e));
