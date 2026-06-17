document.getElementById("generate").addEventListener("click", async () => {
const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Revolut y Transformación de la Banca Tradicional";

// Color palette: Deep Navy + Violet accent (fintech/financial premium)
const C = {
  navy:    "0D1B2A",
  violet:  "5C4EFF",
  teal:    "00C2C7",
  white:   "FFFFFF",
  offwhite:"F4F6FB",
  lightgray:"E8ECF2",
  midgray: "8A94A6",
  darktext:"1A2235",
  gold:    "FFB800",
};

const fontTitle = "Cambria";
const fontBody  = "Calibri";

// ─── Helper: slide background ────────────────────────────────────────────────
function darkSlide(pres) {
  const s = pres.addSlide();
  s.background = { color: C.navy };
  return s;
}
function lightSlide(pres) {
  const s = pres.addSlide();
  s.background = { color: C.offwhite };
  return s;
}

// ─── Helper: section divider pill ────────────────────────────────────────────
function addPill(slide, text, x, y, w) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h: 0.28, fill: { color: C.violet }, rectRadius: 0.14, line: { color: C.violet }
  });
  slide.addText(text, {
    x, y, w, h: 0.28,
    fontSize: 9, fontFace: fontBody, color: C.white, bold: true,
    align: "center", valign: "middle", margin: 0
  });
}

// ─── Helper: stat box ────────────────────────────────────────────────────────
function addStatBox(slide, x, y, w, h, number, label) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: C.navy },
    shadow: { type: "outer", color: "000000", blur: 8, offset: 3, angle: 45, opacity: 0.18 },
    rectRadius: 0.12,
    line: { color: C.teal, width: 1.5 }
  });
  slide.addText(number, {
    x, y: y + 0.1, w, h: h * 0.55,
    fontSize: 28, fontFace: fontTitle, color: C.teal, bold: true,
    align: "center", valign: "bottom"
  });
  slide.addText(label, {
    x, y: y + h * 0.55, w, h: h * 0.4,
    fontSize: 10, fontFace: fontBody, color: C.lightgray,
    align: "center", valign: "top", margin: [0, 4, 0, 4]
  });
}

// ─── Helper: content card ────────────────────────────────────────────────────
function addCard(slide, x, y, w, h, title, body, titleColor) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: C.white },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 45, opacity: 0.10 },
    rectRadius: 0.1,
    line: { color: C.lightgray, width: 0.5 }
  });
  slide.addText(title, {
    x: x + 0.15, y: y + 0.12, w: w - 0.3, h: 0.3,
    fontSize: 11, fontFace: fontTitle, color: titleColor || C.violet, bold: true
  });
  slide.addText(body, {
    x: x + 0.15, y: y + 0.42, w: w - 0.3, h: h - 0.55,
    fontSize: 9.5, fontFace: fontBody, color: C.darktext,
    align: "left", valign: "top"
  });
}

// ─── Helper: dark card ───────────────────────────────────────────────────────
function addDarkCard(slide, x, y, w, h, title, body, accentColor) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h,
    fill: { color: "111E30" },
    rectRadius: 0.1,
    line: { color: accentColor || C.teal, width: 1 }
  });
  slide.addText(title, {
    x: x + 0.15, y: y + 0.12, w: w - 0.3, h: 0.28,
    fontSize: 11, fontFace: fontTitle, color: accentColor || C.teal, bold: true
  });
  slide.addText(body, {
    x: x + 0.15, y: y + 0.42, w: w - 0.3, h: h - 0.55,
    fontSize: 9.5, fontFace: fontBody, color: C.lightgray,
    align: "left", valign: "top"
  });
}

// ─── Slide 1: TITLE ──────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  // Teal accent bar top
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.teal }, line: { color: C.teal }
  });
  // Decorative circle (background)
  s.addShape(pres.shapes.OVAL, {
    x: 6.5, y: -1.2, w: 5.5, h: 5.5,
    fill: { color: C.violet, transparency: 82 }, line: { color: C.violet, transparency: 82 }
  });
  s.addShape(pres.shapes.OVAL, {
    x: 7.5, y: 2.5, w: 4, h: 4,
    fill: { color: C.teal, transparency: 88 }, line: { color: C.teal, transparency: 88 }
  });

  s.addText("REVOLUT", {
    x: 0.6, y: 0.9, w: 8.5, h: 0.9,
    fontSize: 54, fontFace: fontTitle, color: C.teal, bold: true, charSpacing: 8
  });
  s.addText("y la Transformación de la Banca Tradicional", {
    x: 0.6, y: 1.75, w: 7.8, h: 0.65,
    fontSize: 22, fontFace: fontTitle, color: C.white, italic: true
  });
  s.addText("Análisis Estratégico de la Ventaja Competitiva en el Sector Financiero Europeo", {
    x: 0.6, y: 2.5, w: 7.2, h: 0.5,
    fontSize: 13, fontFace: fontBody, color: C.midgray
  });
  // Divider
  s.addShape(pres.shapes.LINE, {
    x: 0.6, y: 3.12, w: 4.5, h: 0,
    line: { color: C.teal, width: 1.5 }
  });
  s.addText("Jessica Elizabeth Rojas Rodríguez", {
    x: 0.6, y: 3.25, w: 5.5, h: 0.3,
    fontSize: 12, fontFace: fontBody, color: C.white, bold: true
  });
  s.addText("MBA · Universidad de Valladolid  |  Tutor: Ignacio Paunero", {
    x: 0.6, y: 3.55, w: 6.5, h: 0.28,
    fontSize: 10, fontFace: fontBody, color: C.midgray
  });
  s.addText("Trabajo de Fin de Máster · 2025–2026", {
    x: 0.6, y: 3.85, w: 6.5, h: 0.28,
    fontSize: 10, fontFace: fontBody, color: C.midgray
  });
  // Bottom bar
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 5.32, w: 10, h: 0.3, fill: { color: "060D16" }, line: { color: "060D16" }
  });
  s.addText("Facultad de Ciencias Económicas y Empresariales · Universidad de Valladolid", {
    x: 0, y: 5.32, w: 10, h: 0.3,
    fontSize: 8.5, fontFace: fontBody, color: C.midgray, align: "center", valign: "middle"
  });
}

// ─── Slide 2: AGENDA ─────────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  s.addText("Estructura del Análisis", {
    x: 0.5, y: 0.3, w: 9, h: 0.55,
    fontSize: 28, fontFace: fontTitle, color: C.navy, bold: true
  });
  s.addShape(pres.shapes.LINE, { x: 0.5, y: 0.9, w: 9, h: 0, line: { color: C.lightgray, width: 0.75 } });

  const items = [
    ["01", "CONTEXTO & PREGUNTA DE INVESTIGACIÓN", "Justificación, antecedentes, paradoja Revolut"],
    ["02", "MARCO TEÓRICO", "Porter, Barney VRIN, Kim-Mauborgne, Bowman, BCG"],
    ["03", "REVOLUT & SECTOR FINANCIERO", "Historia, regulación, transformación competitiva"],
    ["04", "ANÁLISIS ESTRATÉGICO 7 PUNTOS", "PESTEL, Porter 5 Fuerzas, Costes, CX, Datos, Curva de Valor"],
    ["05", "VENTAJA COMPETITIVA SISTÉMICA", "VRIN + Reloj de Bowman + Posicionamiento híbrido"],
    ["06", "LIMITACIONES ESTRUCTURALES", "Regulación, dependencia de terceros, confianza institucional"],
    ["07", "CONCLUSIONES", "Disrupción lograda — consolidación pendiente"],
  ];

  items.forEach(([num, title, sub], i) => {
    const y = 1.1 + i * 0.6;
    // Number circle
    slide_shape(s, num, 0.5, y, C.violet);
    s.addText(title, {
      x: 1.25, y: y + 0.02, w: 4.8, h: 0.3,
      fontSize: 11, fontFace: fontTitle, color: C.navy, bold: true
    });
    s.addText(sub, {
      x: 1.25, y: y + 0.3, w: 5.5, h: 0.22,
      fontSize: 8.5, fontFace: fontBody, color: C.midgray
    });
  });

  function slide_shape(s, txt, x, y, col) {
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.65, h: 0.45, fill: { color: col }, line: { color: col } });
    s.addText(txt, { x, y, w: 0.65, h: 0.45, fontSize: 10, fontFace: fontBody, color: C.white, bold: true, align: "center", valign: "middle" });
  }

  // Right column: key numbers
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 7.0, y: 1.0, w: 2.7, h: 4.3,
    fill: { color: C.navy }, rectRadius: 0.15, line: { color: C.navy }
  });
  s.addText("El caso Revolut\nen cifras", {
    x: 7.0, y: 1.1, w: 2.7, h: 0.55,
    fontSize: 12, fontFace: fontTitle, color: C.teal, bold: true, align: "center"
  });
  const kpis = [["52M", "usuarios globales"],["45.000M$","valoración 2025"],["19.8%","nuevas altas España"],["< 0.3%","banco principal"]];
  kpis.forEach(([val, lbl], i) => {
    s.addText(val, { x: 7.0, y: 1.75 + i * 0.9, w: 2.7, h: 0.42, fontSize: 20, fontFace: fontTitle, color: C.teal, bold: true, align: "center" });
    s.addText(lbl, { x: 7.0, y: 2.12 + i * 0.9, w: 2.7, h: 0.25, fontSize: 9, fontFace: fontBody, color: C.lightgray, align: "center" });
  });
}

// ─── Slide 3: PREGUNTA CENTRAL / PARADOJA ────────────────────────────────────
{
  const s = darkSlide(pres);
  addPill(s, "PREGUNTA DE INVESTIGACIÓN", 0.5, 0.28, 4);
  s.addText("¿Es sostenible la ventaja competitiva de Revolut?", {
    x: 0.5, y: 0.65, w: 9, h: 1.1,
    fontSize: 30, fontFace: fontTitle, color: C.white, bold: true
  });

  // Paradox box
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 1.85, w: 9, h: 1.2,
    fill: { color: C.violet, transparency: 80 },
    line: { color: C.violet, width: 1.5 }, rectRadius: 0.15
  });
  s.addText("LA PARADOJA REVOLUT", {
    x: 0.7, y: 1.95, w: 4, h: 0.3,
    fontSize: 10, fontFace: fontBody, color: C.teal, bold: true
  });
  s.addText("Revolut concentra el 19,8% de las NUEVAS ALTAS en España (más que BBVA, Santander o CaixaBank),\npero solo aparece como banco PRINCIPAL en el < 0,3% de los usuarios.", {
    x: 0.7, y: 2.25, w: 8.5, h: 0.7,
    fontSize: 12, fontFace: fontBody, color: C.white, italic: true
  });

  // 4 drivers
  const drivers = [
    ["⚡", "Eficiencia Operativa", "Arquitectura cloud-native\nsin infraestructura física"],
    ["📊", "Acumulación de Datos", "Activo estratégico acumulativo\ny autorreforzante"],
    ["🆓", "Modelo Freemium", "Captación a bajo coste\nescalada por suscripción"],
    ["🚀", "Velocidad de Innovación", "Ciclos de semanas vs.\nmeses de la banca"],
  ];
  drivers.forEach(([icon, title, body], i) => {
    const x = 0.5 + i * 2.38;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 3.25, w: 2.2, h: 2.1,
      fill: { color: "111E30" }, rectRadius: 0.12,
      line: { color: C.teal, width: 1 }
    });
    s.addText(icon, { x, y: 3.3, w: 2.2, h: 0.5, fontSize: 20, align: "center" });
    s.addText(title, { x: x + 0.1, y: 3.8, w: 2.0, h: 0.35, fontSize: 10, fontFace: fontTitle, color: C.teal, bold: true, align: "center" });
    s.addText(body, { x: x + 0.1, y: 4.15, w: 2.0, h: 0.9, fontSize: 9, fontFace: fontBody, color: C.lightgray, align: "center" });
  });
}

// ─── Slide 4: MARCO TEÓRICO ──────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  addPill(s, "MARCO TEÓRICO", 0.5, 0.2, 2.5);
  s.addText("Herramientas de Análisis Estratégico", {
    x: 0.5, y: 0.55, w: 9, h: 0.55,
    fontSize: 26, fontFace: fontTitle, color: C.navy, bold: true
  });

  const frameworks = [
    { x: 0.35, y: 1.2, title: "Porter (2015)", body: "Ventaja competitiva,\ncadena de valor\ny 5 fuerzas", color: C.violet },
    { x: 2.85, y: 1.2, title: "Barney VRIN (2004)", body: "Sostenibilidad de\nrecursos: Valioso,\nRaro, Inimitable,\nNo sustituible", color: C.teal },
    { x: 5.35, y: 1.2, title: "Kim & Mauborgne (2005)", body: "Océano azul\ny Curva de Valor:\nnuevos espacios\ncompetitivos", color: C.gold },
    { x: 7.85, y: 1.2, title: "Bowman (1995)", body: "Reloj Estratégico:\nprecio vs.\nvalor añadido\npercibido", color: "#E85C50" },
    { x: 0.35, y: 3.15, title: "Abell (1980)", body: "Modelo de\ndefinición del\nnegocio: cliente,\nnecesidad, tecnología", color: "#50B5A0" },
    { x: 2.85, y: 3.15, title: "BCG / UEN", body: "Cartera de productos:\nEstrellas, Vacas,\nInterrogantes, Perros", color: C.violet },
    { x: 5.35, y: 3.15, title: "PESTEL + Grupos", body: "Macroentorno y\ngrupos estratégicos\nen el sector\nfintech europeo", color: C.teal },
    { x: 7.85, y: 3.15, title: "DAFO", body: "Síntesis del\nanálisis interno\ny externo de\nRevolut", color: C.gold },
  ];

  frameworks.forEach(({ x, y, title, body, color }) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: 2.3, h: 2.15,
      fill: { color: C.white },
      shadow: { type: "outer", color: "000000", blur: 5, offset: 2, angle: 45, opacity: 0.08 },
      rectRadius: 0.12, line: { color: color, width: 2 }
    });
    s.addText(title, {
      x: x + 0.1, y: y + 0.12, w: 2.1, h: 0.35,
      fontSize: 10, fontFace: fontTitle, color: color, bold: true
    });
    s.addText(body, {
      x: x + 0.1, y: y + 0.48, w: 2.1, h: 1.55,
      fontSize: 9.5, fontFace: fontBody, color: C.darktext
    });
  });
}

// ─── Slide 5: HISTORIA REVOLUT ───────────────────────────────────────────────
{
  const s = darkSlide(pres);
  addPill(s, "REVOLUT: HISTORIA Y EVOLUCIÓN", 0.5, 0.2, 4);
  s.addText("De tarjeta prepago a plataforma financiera integral", {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 22, fontFace: fontTitle, color: C.white, bold: true
  });

  const timeline = [
    { year: "2015", color: C.teal, event: "FUNDACIÓN", detail: "Tarjeta prepago sin comisión en divisas · Storonsky & Yatsenko · 100K usuarios (referidos)" },
    { year: "2018", color: C.violet, event: "LICENCIA BANCARIA", detail: "Licencia del Banco de Lituania · Pasaporte europeo (27 estados EEE) · Challenger bank" },
    { year: "2020", color: C.gold, event: "PLATAFORMA INTEGRAL", detail: "Inversión en bolsa, cripto, seguros · Valoración: 4.500M$ · Expansión a nuevos mercados" },
    { year: "2023", color: "#E85C50", event: "PRIMER BENEFICIO", detail: "Ingresos: 2.200M$ · EBITDA: 545M$ · 35M usuarios · Modelo freemium validado" },
    { year: "2024–25", color: C.teal, event: "LIDERAZGO EN ESPAÑA", detail: "19,8% nuevas altas · Licencia UK obtenida · 52M usuarios · Valoración: 45.000M$" },
  ];

  // Timeline line
  s.addShape(pres.shapes.LINE, {
    x: 1.35, y: 1.35, w: 0, h: 3.6,
    line: { color: C.teal, width: 1.5 }
  });

  timeline.forEach(({ year, color, event, detail }, i) => {
    const y = 1.2 + i * 0.78;
    // Dot
    s.addShape(pres.shapes.OVAL, { x: 1.12, y: y + 0.02, w: 0.45, h: 0.38, fill: { color }, line: { color } });
    // Year
    s.addText(year, { x: 0.0, y: y + 0.04, w: 1.1, h: 0.3, fontSize: 11, fontFace: fontTitle, color, bold: true, align: "right" });
    // Event + detail
    s.addText(event + "  ", { x: 1.65, y: y, w: 2.4, h: 0.3, fontSize: 10, fontFace: fontTitle, color, bold: true });
    s.addText(detail, { x: 1.65, y: y + 0.28, w: 7.9, h: 0.4, fontSize: 9, fontFace: fontBody, color: C.lightgray });
  });
}

// ─── Slide 6: REGULACIÓN ─────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  addPill(s, "MARCO REGULATORIO", 0.5, 0.2, 3);
  s.addText("La Regulación como Variable Estratégica", {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 24, fontFace: fontTitle, color: C.navy, bold: true
  });

  // 3 dimensions
  const dims = [
    { x: 0.4, color: C.violet, num: "1", title: "Filtro de Entrada", items: ["Basilea III: requisitos de capital mínimo", "Licencias BCE, EBA, nacionales", "LCR 186%, NSFR 134% (banca española)", "Barrera insalvable sin escala suficiente"] },
    { x: 3.6, color: C.teal, num: "2", title: "Reordenador de Rivalidad", items: ["Asimetría: banca vs. fintech (EMD2/PSD2)", "Pasaporte europeo: licencia única EEE", "Sandboxes regulatorios (FCA 2016, BdE 2021)", "Ventaja de Revolut: Lituania → 27 países"] },
    { x: 6.8, color: "#E85C50", num: "3", title: "Coste Permanente y Creciente", items: ["AML/KYC: Reg. UE 2024/1620 obligatorio", "DORA + MiCAR: ampliación del perímetro", "GDPR: protección de datos recurrente", "Multa 3,5M€ Banco Lituania (2025)"] },
  ];
  dims.forEach(({ x, color, num, title, items }) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.2, w: 3.1, h: 4.15,
      fill: { color: C.white },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 45, opacity: 0.09 },
      rectRadius: 0.12, line: { color: color, width: 2 }
    });
    s.addShape(pres.shapes.OVAL, { x: x + 0.12, y: 1.28, w: 0.52, h: 0.43, fill: { color }, line: { color } });
    s.addText(num, { x: x + 0.12, y: 1.28, w: 0.52, h: 0.43, fontSize: 13, fontFace: fontTitle, color: C.white, bold: true, align: "center", valign: "middle" });
    s.addText(title, { x: x + 0.72, y: 1.3, w: 2.35, h: 0.4, fontSize: 11, fontFace: fontTitle, color, bold: true });
    items.forEach((item, j) => {
      s.addText([{ text: "› ", options: { color, bold: true } }, { text: item, options: { color: C.darktext } }], {
        x: x + 0.18, y: 1.85 + j * 0.55, w: 2.85, h: 0.48,
        fontSize: 9.5, fontFace: fontBody
      });
    });
  });

  // Bottom insight
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 5.1, w: 9.2, h: 0.35,
    fill: { color: C.navy }, rectRadius: 0.1, line: { color: C.navy }
  });
  s.addText("Revolut convirtió la regulación en ventaja: licencia lituana (2018) → pasaporte europeo → operación en 27 países con una sola licencia", {
    x: 0.55, y: 5.1, w: 9.0, h: 0.35,
    fontSize: 9.5, fontFace: fontBody, color: C.teal, align: "center", valign: "middle", italic: true
  });
}

// ─── Slide 7: ANÁLISIS PESTEL ────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  addPill(s, "ANÁLISIS ESTRATÉGICO · PUNTO 4.1", 0.5, 0.18, 4);
  s.addText("Análisis PESTEL: Macroentorno de Revolut", {
    x: 0.5, y: 0.52, w: 9, h: 0.5,
    fontSize: 22, fontFace: fontTitle, color: C.white, bold: true
  });

  const pestel = [
    { l: "P", col: "#5C4EFF", title: "Político / Legal", body: "Mercado único digital UE · MiCAR regula cripto · PSD2 habilita open banking · Pasaporte europeo" },
    { l: "E", col: C.teal, title: "Económico", body: "Tipos bajos erosionaron ventaja depósitos banca · Subida tipos 2022 presiona costes · Crecimiento e-commerce y pagos cross-border" },
    { l: "S", col: C.gold, title: "Social", body: "Generaciones Millennial/Z: conveniencia, transparencia, personalización · Desconfianza en banca post-2008 · Smartphone como canal financiero" },
    { l: "T", col: "#50B5A0", title: "Tecnológico", body: "Cloud a bajo coste · APIs abiertas · IA y machine learning · Biometría reduce fricción onboarding · Riesgo: phishing, malware, ransomware" },
    { l: "E", col: "#FF8C42", title: "Ecológico", body: "Presión ESG creciente · Huella digital menor vs. redes de oficinas físicas · Regulación sostenibilidad financiera UE" },
    { l: "L", col: "#E85C50", title: "Legal (Detalle)", body: "Basilea III, DORA, MiCAR, GDPR, PSD2 · Coste AML/KYC creciente · Sandboxes regulatorios como catalizadores de innovación" },
  ];
  pestel.forEach(({ l, col, title, body }, i) => {
    const col_i = i % 3;
    const row_i = Math.floor(i / 3);
    const x = 0.4 + col_i * 3.2;
    const y = 1.15 + row_i * 1.9;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: 3.0, h: 1.7,
      fill: { color: "111E30" }, rectRadius: 0.1,
      line: { color: col, width: 1.5 }
    });
    // Letter badge
    s.addShape(pres.shapes.OVAL, { x: x + 0.1, y: y + 0.12, w: 0.42, h: 0.38, fill: { color: col }, line: { color: col } });
    s.addText(l, { x: x + 0.1, y: y + 0.12, w: 0.42, h: 0.38, fontSize: 12, fontFace: fontTitle, color: C.white, bold: true, align: "center", valign: "middle" });
    s.addText(title, { x: x + 0.6, y: y + 0.14, w: 2.35, h: 0.28, fontSize: 10, fontFace: fontTitle, color: col, bold: true });
    s.addText(body, { x: x + 0.1, y: y + 0.48, w: 2.85, h: 1.1, fontSize: 9, fontFace: fontBody, color: C.lightgray });
  });
}

// ─── Slide 8: CINCO FUERZAS PORTER ───────────────────────────────────────────
{
  const s = lightSlide(pres);
  addPill(s, "ANÁLISIS ESTRATÉGICO · PUNTO 4.1", 0.5, 0.2, 4);
  s.addText("Las 5 Fuerzas de Porter en el Sector Fintech Europeo", {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 22, fontFace: fontTitle, color: C.navy, bold: true
  });

  // Center box (Rivalidad)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 3.5, y: 2.2, w: 3.0, h: 1.6,
    fill: { color: C.navy }, rectRadius: 0.12, line: { color: C.teal, width: 2 }
  });
  s.addText("RIVALIDAD", { x: 3.5, y: 2.28, w: 3.0, h: 0.35, fontSize: 11, fontFace: fontTitle, color: C.teal, bold: true, align: "center" });
  s.addText("ALTA · Intensificada", { x: 3.5, y: 2.6, w: 3.0, h: 0.28, fontSize: 9, fontFace: fontBody, color: C.gold, bold: true, align: "center" });
  s.addText("N26 · Monzo · Wise · Bunq\nBBVA IA · CaixaBank GalaxIA", { x: 3.5, y: 2.88, w: 3.0, h: 0.7, fontSize: 8.5, fontFace: fontBody, color: C.lightgray, align: "center" });

  // Arrows (connector lines)
  const arrows = [[4.85, 1.5, 4.85, 2.2], [4.85, 3.8, 4.85, 4.5], [0.8, 3.0, 3.5, 3.0], [7.0, 3.0, 6.5, 3.0]];
  arrows.forEach(([x1, y1, x2, y2]) => {
    s.addShape(pres.shapes.LINE, { x: x1, y: y1, w: x2 - x1, h: y2 - y1, line: { color: C.midgray, width: 1, dashType: "dash" } });
  });

  // 4 outer forces
  const forces = [
    { x: 3.2, y: 0.92, w: 3.6, title: "Nuevos Entrantes", intensity: "BAJA en banca · MEDIA en fintech", detail: "Regulación bloquea entrada masiva · Sandboxes facilitan startups · Revolut demostró que la barrera es gestionable" },
    { x: 3.2, y: 4.05, w: 3.6, title: "Sustitutivos (BigTech)", intensity: "ALTA · Mayor amenaza futura", detail: "Google, Apple, Amazon: ecosistemas con 100M+ usuarios · Sin licencias bancarias plenas en Europa (aún)" },
    { x: 0.35, y: 2.1, w: 2.6, title: "Proveedores", intensity: "BAJO · Multi-Cloud", detail: "AWS/GCP/Azure intercambiables · Sin dependencia de legacy TI · Agilidad de migración" },
    { x: 7.05, y: 2.1, w: 2.6, title: "Clientes", intensity: "ALTA · Bajo coste de cambio", detail: "Comparación inmediata · Alta en minutos · Revolut: valoración top en App Store" },
  ];
  forces.forEach(({ x, y, w, title, intensity, detail }) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h: 1.75,
      fill: { color: C.white },
      shadow: { type: "outer", color: "000000", blur: 5, offset: 2, angle: 45, opacity: 0.09 },
      rectRadius: 0.1, line: { color: C.lightgray, width: 0.75 }
    });
    s.addText(title, { x: x + 0.1, y: y + 0.1, w: w - 0.2, h: 0.3, fontSize: 10, fontFace: fontTitle, color: C.navy, bold: true });
    s.addText(intensity, { x: x + 0.1, y: y + 0.38, w: w - 0.2, h: 0.25, fontSize: 9, fontFace: fontBody, color: C.violet, bold: true });
    s.addText(detail, { x: x + 0.1, y: y + 0.65, w: w - 0.2, h: 0.95, fontSize: 8.5, fontFace: fontBody, color: C.darktext });
  });
}

// ─── Slide 9: VENTAJA EN COSTES ──────────────────────────────────────────────
{
  const s = darkSlide(pres);
  addPill(s, "ANÁLISIS ESTRATÉGICO · PUNTO 4.2", 0.5, 0.18, 4);
  s.addText("La Ventaja en Costes como Pilar Estratégico", {
    x: 0.5, y: 0.52, w: 9, h: 0.5,
    fontSize: 22, fontFace: fontTitle, color: C.white, bold: true
  });

  // Comparison table
  const headers = ["", "BANCA TRADICIONAL", "REVOLUT"];
  const rows = [
    ["Infraestructura", "Red de sucursales, cajeros, sistemas legacy", "Arquitectura cloud-native, APIs modulares, IA"],
    ["Actividades primarias", "Atención presencial, formularios, semanas de espera", "App móvil, onboarding < 10 min, IA y ML"],
    ["Actividades de apoyo", "Edificios, cajeros, sistemas propietarios", "Código, datos, desarrolladores"],
    ["Escala", "Coste fijo elevado proporcional a clientes", "Escala global sin incrementar costes fijos"],
    ["Captación de clientes", "CAC alto: publicidad, oficinas, personal", "CAC bajo: referidos y modelo freemium"],
  ];

  const colW = [2.2, 3.5, 3.5];
  const colX = [0.4, 2.65, 6.2];
  // Header row
  headers.forEach((h, ci) => {
    if (ci === 0) return;
    s.addShape(pres.shapes.RECTANGLE, {
      x: colX[ci], y: 1.18, w: colW[ci], h: 0.38,
      fill: { color: ci === 2 ? C.teal : C.midgray }, line: { color: ci === 2 ? C.teal : C.midgray }
    });
    s.addText(h, { x: colX[ci], y: 1.18, w: colW[ci], h: 0.38, fontSize: 10, fontFace: fontTitle, color: C.navy, bold: true, align: "center", valign: "middle" });
  });
  rows.forEach((row, ri) => {
    const y = 1.58 + ri * 0.65;
    const bg = ri % 2 === 0 ? "111E30" : "0D1B2A";
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 9.3, h: 0.62, fill: { color: bg }, line: { color: bg } });
    row.forEach((cell, ci) => {
      s.addText(cell, {
        x: colX[ci] + 0.05, y: y + 0.04, w: colW[ci] - 0.1, h: 0.54,
        fontSize: 9, fontFace: ci === 0 ? fontTitle : fontBody,
        color: ci === 2 ? C.teal : ci === 0 ? C.gold : C.lightgray,
        bold: ci === 0, valign: "middle"
      });
    });
  });

  // Ingresos fact
  s.addText("📈  Ingresos Revolut +45% → £923M · EBITDA 2023: 545M$ · Primer beneficio neto certificado en 2023", {
    x: 0.4, y: 4.98, w: 9.2, h: 0.35,
    fontSize: 9.5, fontFace: fontBody, color: C.teal, italic: true
  });
}

// ─── Slide 10: MODELO FREEMIUM ───────────────────────────────────────────────
{
  const s = lightSlide(pres);
  addPill(s, "ANÁLISIS ESTRATÉGICO · PUNTO 4.2", 0.5, 0.2, 4);
  s.addText("Modelo Freemium: Captación y Monetización", {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 24, fontFace: fontTitle, color: C.navy, bold: true
  });

  // Pyramid / subscription levels
  const levels = [
    { label: "ULTRA / METAL", pct: 0.2, color: C.gold, x: 3.5, w: 3.0, y: 1.25, desc: "Salas VIP, cashback, metal card, máx. beneficios" },
    { label: "PREMIUM / PLUS", pct: 0.3, color: C.violet, x: 2.8, w: 4.4, y: 2.05, desc: "Seguros viaje, cripto ampliado, soporte prioritario" },
    { label: "STANDARD (Freemium)", pct: 0.5, color: C.teal, x: 1.8, w: 6.4, y: 2.85, desc: "Cuenta gratuita · Captación sin fricción · Motor de referidos" },
  ];
  levels.forEach(({ label, pct, color, x, w, y, desc }) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h: 0.68,
      fill: { color }, rectRadius: 0.08, line: { color }
    });
    s.addText(label, {
      x: x + 0.1, y: y + 0.05, w: w - 0.2, h: 0.3,
      fontSize: 10, fontFace: fontTitle, color: C.white, bold: true, align: "center"
    });
    s.addText(desc, {
      x: x + 0.1, y: y + 0.35, w: w - 0.2, h: 0.28,
      fontSize: 8.5, fontFace: fontBody, color: C.white, align: "center"
    });
  });

  // 4 levers
  const levers = [
    ["Ahorro de tiempo", "Onboarding < 10 min\nNo filas, no papeles"],
    ["Ahorro de costes", "Sin comisiones en divisas\nTransferencias gratis"],
    ["Mejor experiencia UX", "App intuitiva, notificaciones\nen tiempo real, control gasto"],
    ["Eliminación de barreras", "Acceso al sistema financiero\nsin burocracia tradicional"],
  ];
  levers.forEach(([title, body], i) => {
    addCard(s, 0.4 + i * 2.38, 3.72, 2.2, 1.55, title, body, C.violet);
  });
  s.addText("Valoración: 45.000M$ (2025)  ·  De 4.500M$ (2020) a 45.000M$ (2024)  ·  Fintech privada más valiosa de Europa", {
    x: 0.4, y: 5.25, w: 9.2, h: 0.25,
    fontSize: 9, fontFace: fontBody, color: C.midgray, align: "center", italic: true
  });
}

// ─── Slide 11: 4.3 CURVA DE VALOR ────────────────────────────────────────────
{
  const s = darkSlide(pres);
  addPill(s, "ANÁLISIS ESTRATÉGICO · PUNTO 4.3", 0.5, 0.18, 3.8);
  s.addText("Curva de Valor: Revolut vs. Banca Tradicional", {
    x: 0.5, y: 0.52, w: 9, h: 0.5,
    fontSize: 22, fontFace: fontTitle, color: C.white, bold: true
  });

  // Bar chart
  const factors = [
    { label: "Comisiones divisas", bank: 1, rev: 9 },
    { label: "Onboarding digital", bank: 2, rev: 9 },
    { label: "Transparencia", bank: 3, rev: 9 },
    { label: "Velocidad transf.", bank: 2, rev: 9 },
    { label: "Cobertura geográf.", bank: 7, rev: 8 },
    { label: "Amplitud producto", bank: 8, rev: 7 },
    { label: "Garantía depósitos", bank: 9, rev: 5 },
    { label: "Confianza inst.", bank: 9, rev: 3 },
    { label: "Banco principal", bank: 9, rev: 1 },
  ];

  const chartX = 0.4, chartY = 1.2, chartW = 9.2, chartH = 3.5;
  const nFact = factors.length;
  const barW = chartW / nFact;
  const maxH = chartH * 0.85;

  [2, 4, 6, 8, 10].forEach(v => {
    const yg = chartY + maxH - (v / 10) * maxH;
    s.addShape(pres.shapes.LINE, { x: chartX, y: yg, w: chartW, h: 0, line: { color: "1E3050", width: 0.5 } });
    s.addText(String(v), { x: chartX - 0.35, y: yg - 0.12, w: 0.3, h: 0.24, fontSize: 7.5, fontFace: fontBody, color: C.midgray, align: "right" });
  });

  factors.forEach(({ label, bank, rev }, i) => {
    const x = chartX + i * barW;
    const bH = (bank / 10) * maxH;
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + barW * 0.05, y: chartY + maxH - bH, w: barW * 0.35, h: bH,
      fill: { color: C.midgray, transparency: 30 }, line: { color: C.midgray }
    });
    const rH = (rev / 10) * maxH;
    s.addShape(pres.shapes.RECTANGLE, {
      x: x + barW * 0.48, y: chartY + maxH - rH, w: barW * 0.35, h: rH,
      fill: { color: C.teal, transparency: 20 }, line: { color: C.teal }
    });
    s.addText(label, {
      x: x, y: chartY + maxH + 0.08, w: barW, h: 0.55,
      fontSize: 7, fontFace: fontBody, color: C.lightgray, align: "center"
    });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: 4.85, w: 0.35, h: 0.2, fill: { color: C.midgray }, line: { color: C.midgray } });
  s.addText("Banca Tradicional", { x: 6.9, y: 4.83, w: 1.5, h: 0.24, fontSize: 9, fontFace: fontBody, color: C.lightgray });
  s.addShape(pres.shapes.RECTANGLE, { x: 8.3, y: 4.85, w: 0.35, h: 0.2, fill: { color: C.teal }, line: { color: C.teal } });
  s.addText("Revolut", { x: 8.7, y: 4.83, w: 1.0, h: 0.24, fontSize: 9, fontFace: fontBody, color: C.teal });

  s.addText("Revolut lidera en experiencia digital; la brecha persiste en confianza institucional y primacía como banco principal", {
    x: 0.4, y: 5.22, w: 9.2, h: 0.28,
    fontSize: 9, fontFace: fontBody, color: C.midgray, italic: true, align: "center"
  });
}

// ─── Slide NEW: 4.6 TRANSFORMACIÓN SECTORIAL Y REACCIÓN DE LA COMPETENCIA ────
{
  const s = lightSlide(pres);
  addPill(s, "ANÁLISIS ESTRATÉGICO · PUNTO 4.6", 0.5, 0.2, 3.8);
  s.addText("Transformación Sectorial y Reacción de la Competencia", {
    x: 0.5, y: 0.55, w: 9, h: 0.55,
    fontSize: 22, fontFace: fontTitle, color: C.navy, bold: true
  });

  // Left column: ruptura estructural (4.6.1)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 1.25, w: 4.5, h: 4.1,
    fill: { color: C.white },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 45, opacity: 0.09 },
    rectRadius: 0.12, line: { color: C.navy, width: 1.5 }
  });
  s.addText("4.6.1  Del Modelo Presencial al Digital", {
    x: 0.55, y: 1.35, w: 4.2, h: 0.32,
    fontSize: 11, fontFace: fontTitle, color: C.navy, bold: true
  });

  // Pre-Revolut chain
  s.addText("ANTES (Banca Tradicional)", { x: 0.55, y: 1.75, w: 4.2, h: 0.25, fontSize: 9, fontFace: fontTitle, color: C.midgray, bold: true });
  const preChain = ["Acudir a sucursal → documentación física", "Alta: varios días de espera", "Transferencias internacionales: 1-3 días hábiles", "Intermediarios corresponsales + comisiones", "Atención presencial → fricción percibida como coste"];
  preChain.forEach((item, i) => {
    s.addText([{ text: "✗  ", options: { color: "#E85C50", bold: true } }, { text: item }], {
      x: 0.6, y: 2.02 + i * 0.37, w: 4.2, h: 0.33,
      fontSize: 8.5, fontFace: fontBody, color: C.darktext
    });
  });

  // Post-Revolut chain
  s.addText("DESPUÉS (Revolut)", { x: 0.55, y: 3.9, w: 4.2, h: 0.25, fontSize: 9, fontFace: fontTitle, color: C.teal, bold: true });
  const postChain = ["Onboarding móvil < 10 min · biometría + IA", "Transferencias internacionales en segundos · tipo interbancario"];
  postChain.forEach((item, i) => {
    s.addText([{ text: "✓  ", options: { color: C.teal, bold: true } }, { text: item }], {
      x: 0.6, y: 4.18 + i * 0.42, w: 4.2, h: 0.38,
      fontSize: 8.5, fontFace: fontBody, color: C.darktext
    });
  });

  // Right column: respuesta competidores (4.6.2 y 4.6.3)
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.1, y: 1.25, w: 4.5, h: 4.1,
    fill: { color: C.white },
    shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 45, opacity: 0.09 },
    rectRadius: 0.12, line: { color: C.violet, width: 1.5 }
  });
  s.addText("4.6.2–4.6.3  Respuesta de la Banca e Implicaciones", {
    x: 5.25, y: 1.35, w: 4.2, h: 0.32,
    fontSize: 11, fontFace: fontTitle, color: C.violet, bold: true
  });

  const responses = [
    { actor: "Openbank (Santander)", desc: "Relanzado 2017 · infraestructura cloud · 3M+ clientes España · expandido a Alemania, Portugal, Argentina" },
    { actor: "BBVA", desc: "Plataforma IA para asesoramiento financiero personalizado en tiempo real" },
    { actor: "CaixaBank GalaxIA", desc: "Iniciativa IA para banca minorista · reducción de brecha en experiencia de usuario" },
  ];
  responses.forEach(({ actor, desc }, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 5.2, y: 1.8 + i * 0.82, w: 4.25, h: 0.72,
      fill: { color: C.offwhite }, rectRadius: 0.08,
      line: { color: C.lightgray, width: 0.75 }
    });
    s.addText(actor, { x: 5.3, y: 1.87 + i * 0.82, w: 4.05, h: 0.25, fontSize: 9.5, fontFace: fontTitle, color: C.violet, bold: true });
    s.addText(desc, { x: 5.3, y: 2.1 + i * 0.82, w: 4.05, h: 0.38, fontSize: 8.5, fontFace: fontBody, color: C.darktext });
  });

  // Implication box
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.2, y: 4.3, w: 4.25, h: 1.0,
    fill: { color: C.navy }, rectRadius: 0.1,
    line: { color: C.teal, width: 1 }
  });
  s.addText("Implicación clave (4.6.3)", { x: 5.32, y: 4.38, w: 4.0, h: 0.28, fontSize: 9.5, fontFace: fontTitle, color: C.teal, bold: true });
  s.addText("La sostenibilidad de Revolut depende de que su ritmo de innovación siga superando el de sus imitadores. Las ventajas puramente digitales tenderán a reducirse con el tiempo.", {
    x: 5.32, y: 4.67, w: 4.0, h: 0.55,
    fontSize: 8.5, fontFace: fontBody, color: C.lightgray
  });
}

// ─── Slide NEW: 4.7 EVOLUCIÓN DE LA CURVA DE VALOR: SÍNTESIS ─────────────────
{
  const s = darkSlide(pres);
  addPill(s, "ANÁLISIS ESTRATÉGICO · PUNTO 4.7", 0.5, 0.18, 3.8);
  s.addText("Evolución de la Curva de Valor: Síntesis y Resultados", {
    x: 0.5, y: 0.52, w: 9, h: 0.5,
    fontSize: 22, fontFace: fontTitle, color: C.white, bold: true
  });

  // 3 evolutionary phases as timeline cards
  const phases = [
    {
      year: "2015", color: C.teal, title: "Curva estrecha — un factor",
      body: "Solo un factor diferenciado: tipo de cambio sin comisión al máximo. En todos los demás (amplitud, garantías, presencia, historial), curva plana o inexistente. Captura de segmento desatendido: viajeros y expatriados.",
      tag: "OCÉANO AZUL PURO"
    },
    {
      year: "2018–22", color: C.violet, title: "Curva en expansión — posición híbrida",
      body: "Incorporación de inversión, cripto, seguros, crédito. La curva sube en amplitud de producto. El Reloj de Bowman (1995) sitúa a Revolut en zona híbrida: precio competitivo + valor añadido creciente. Licencia lituana → pasaporte europeo.",
      tag: "BOWMAN: ZONA HÍBRIDA"
    },
    {
      year: "2023–25", color: C.gold, title: "Síntesis — la frontera no resuelta",
      body: "La curva de valor de Revolut supera a la banca en experiencia digital, velocidad y coste. Pero la brecha en confianza institucional y banco principal sigue siendo estructural. La paradoja: 19,8% nuevas altas vs. < 0,3% banco principal.",
      tag: "CONSOLIDACIÓN PENDIENTE"
    },
  ];

  phases.forEach(({ year, color, title, body, tag }, i) => {
    const y = 1.2 + i * 1.45;
    // Year badge
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y, w: 1.2, h: 1.25,
      fill: { color }, rectRadius: 0.1, line: { color }
    });
    s.addText(year, { x: 0.4, y: y + 0.28, w: 1.2, h: 0.45, fontSize: 16, fontFace: fontTitle, color: C.white, bold: true, align: "center" });
    s.addText(tag, { x: 0.4, y: y + 0.72, w: 1.2, h: 0.45, fontSize: 7, fontFace: fontBody, color: C.white, align: "center" });

    // Content card
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 1.75, y, w: 7.9, h: 1.25,
      fill: { color: "111E30" }, rectRadius: 0.1,
      line: { color, width: 1 }
    });
    s.addText(title, { x: 1.9, y: y + 0.1, w: 7.65, h: 0.3, fontSize: 11, fontFace: fontTitle, color, bold: true });
    s.addText(body, { x: 1.9, y: y + 0.42, w: 7.65, h: 0.72, fontSize: 9, fontFace: fontBody, color: C.lightgray });
  });

  // DAFO note bottom
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 5.1, w: 9.2, h: 0.35,
    fill: { color: C.violet, transparency: 75 }, rectRadius: 0.1,
    line: { color: C.violet, width: 1 }
  });
  s.addText("Este punto actúa como síntesis del DAFO (Humphrey, 1960): la curva de valor refleja visualmente qué ventajas son sostenibles y cuáles están bajo presión competitiva creciente.", {
    x: 0.55, y: 5.1, w: 9.0, h: 0.35,
    fontSize: 9, fontFace: fontBody, color: C.white, align: "center", valign: "middle", italic: true
  });
}

// ─── Slide 12: DATOS COMO ACTIVO ESTRATÉGICO ─────────────────────────────────
{
  const s = lightSlide(pres);
  addPill(s, "ANÁLISIS ESTRATÉGICO · PUNTO 4.5", 0.5, 0.2, 4);
  s.addText("Datos de Clientes: Activo Estratégico Acumulativo", {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 22, fontFace: fontTitle, color: C.navy, bold: true
  });

  // VRIN evaluation
  const vrin = [
    { letter: "V", label: "Valioso", desc: "Mejora rentabilidad por usuario · Reduce coste del riesgo crediticio · Alimenta modelos de scoring y antifraude", color: C.violet },
    { letter: "R", label: "Raro", desc: "Requiere volumen y antigüedad de relación acumulada · 10 años de datos de comportamiento financiero real", color: C.teal },
    { letter: "I", label: "Inimitable", desc: "Dimensión histórica no recreable ni comprable · Combinación única de transacciones + contexto financiero", color: C.gold },
    { letter: "N", label: "No sustituible", desc: "Ningún dato de terceros ofrece la misma granularidad en conducta financiera real del usuario", color: "#E85C50" },
  ];
  vrin.forEach(({ letter, label, desc, color }, i) => {
    const x = 0.4 + i * 2.38;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.2, w: 2.2, h: 2.8,
      fill: { color: C.white },
      shadow: { type: "outer", color: "000000", blur: 6, offset: 2, angle: 45, opacity: 0.1 },
      rectRadius: 0.12, line: { color: color, width: 2 }
    });
    s.addShape(pres.shapes.OVAL, { x: x + 0.75, y: 1.3, w: 0.7, h: 0.6, fill: { color }, line: { color } });
    s.addText(letter, { x: x + 0.75, y: 1.3, w: 0.7, h: 0.6, fontSize: 18, fontFace: fontTitle, color: C.white, bold: true, align: "center", valign: "middle" });
    s.addText(label, { x: x + 0.1, y: 2.0, w: 2.0, h: 0.3, fontSize: 11, fontFace: fontTitle, color, bold: true, align: "center" });
    s.addText(desc, { x: x + 0.1, y: 2.35, w: 2.0, h: 1.5, fontSize: 9, fontFace: fontBody, color: C.darktext });
  });

  // Data flywheel
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.4, y: 4.15, w: 9.2, h: 1.15,
    fill: { color: C.navy }, rectRadius: 0.12, line: { color: C.teal, width: 1 }
  });
  s.addText("EL FLYWHEEL DE DATOS →", { x: 0.6, y: 4.25, w: 2.5, h: 0.3, fontSize: 9.5, fontFace: fontTitle, color: C.teal, bold: true });
  s.addText("Más usuarios activos  →  Más datos  →  Mejor personalización  →  Mayor retención  →  Más usuarios activos", {
    x: 0.6, y: 4.55, w: 8.8, h: 0.5,
    fontSize: 11, fontFace: fontBody, color: C.white, align: "center"
  });
  s.addText("Condición: cumplimiento RGPD · Riesgo: ciberseguridad (phishing, malware) · ISO 27001 + PCI-DSS certificados", {
    x: 0.6, y: 5.02, w: 8.8, h: 0.22,
    fontSize: 8.5, fontFace: fontBody, color: C.midgray, align: "center", italic: true
  });
}

// ─── Slide 13: DIVERSIFICACIÓN E INTERNACIONALIZACIÓN ────────────────────────
{
  const s = darkSlide(pres);
  addPill(s, "ANÁLISIS ESTRATÉGICO · PUNTO 4.4", 0.5, 0.18, 4);
  s.addText("Expansión del Portafolio e Internacionalización", {
    x: 0.5, y: 0.52, w: 9, h: 0.5,
    fontSize: 22, fontFace: fontTitle, color: C.white, bold: true
  });

  // 5 phases
  const phases = [
    { num: "①", year: "2015", title: "Tarjeta prepago", body: "Tipo de cambio interbancario\nSin comisiones divisas", color: C.teal },
    { num: "②", year: "2016-18", title: "Cuenta operativa", body: "IBAN, transferencias\nDomiciliación nómina", color: "#50B5A0" },
    { num: "③", year: "2019-20", title: "Inversión", body: "Acciones fraccionadas\nETFs desde la app", color: C.violet },
    { num: "④", year: "2020-21", title: "Criptomonedas", body: "Compra, venta, custodia\nMiCAR → marco regulado", color: C.gold },
    { num: "⑤", year: "2022-25", title: "Banco completo", body: "Seguros, crédito, ahorro\nLicencia bancaria UK (2024)", color: "#E85C50" },
  ];
  phases.forEach(({ num, year, title, body, color }, i) => {
    const x = 0.35 + i * 1.88;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: 1.2, w: 1.75, h: 2.2,
      fill: { color: "111E30" }, rectRadius: 0.1,
      line: { color, width: 1.5 }
    });
    s.addText(num, { x, y: 1.25, w: 1.75, h: 0.45, fontSize: 18, fontFace: fontTitle, color, bold: true, align: "center" });
    s.addText(year, { x, y: 1.65, w: 1.75, h: 0.25, fontSize: 9, fontFace: fontBody, color: C.midgray, align: "center" });
    s.addText(title, { x: x + 0.08, y: 1.9, w: 1.6, h: 0.32, fontSize: 10, fontFace: fontTitle, color, bold: true, align: "center" });
    s.addText(body, { x: x + 0.08, y: 2.22, w: 1.6, h: 0.95, fontSize: 8.5, fontFace: fontBody, color: C.lightgray, align: "center" });
    // Arrow between phases
    if (i < 4) {
      s.addShape(pres.shapes.LINE, { x: x + 1.78, y: 2.1, w: 0.1, h: 0, line: { color: C.midgray, width: 1 } });
    }
  });

  // Internacionalización 3 levers
  const levers = [
    { title: "Infraestructura Tech", desc: "Cloud escala sin coste fijo · APIs abiertas para integración local · Sin presencia física requerida" },
    { title: "Demanda Global", desc: "Movilidad laboral · E-commerce transfronterizo · Remesas · Servicios multidivisa más eficientes" },
    { title: "Gestión Regulatoria", desc: "Estrategia por jurisdicción · UK licencia 2024 · EE.UU. en proceso · 48 países, 52M usuarios" },
  ];
  levers.forEach(({ title, desc }, i) => {
    addDarkCard(s, 0.4 + i * 3.2, 3.6, 3.0, 1.65, title, desc, C.teal);
  });
}

// ─── Slide 14: DAFO ──────────────────────────────────────────────────────────
{
  const s = lightSlide(pres);
  addPill(s, "SÍNTESIS ESTRATÉGICA", 0.5, 0.2, 2.5);
  s.addText("Análisis DAFO de Revolut", {
    x: 0.5, y: 0.55, w: 9, h: 0.5,
    fontSize: 26, fontFace: fontTitle, color: C.navy, bold: true
  });

  const quad = [
    { x: 0.4, y: 1.2, color: "#2E7D52", label: "FORTALEZAS", items: ["Arquitectura cloud-native: eficiencia operativa", "Ecosistema integrado: pagos, cripto, inversión", "Datos y personalización basada en IA", "Modelo freemium: captación a bajo CAC", "Velocidad de innovación: semanas vs. meses", "Valoración 45.000M$: líder fintech EU"] },
    { x: 5.15, y: 1.2, color: C.violet, label: "DEBILIDADES", items: ["Confianza institucional limitada (neobanco)", "Banco principal < 0,3% de usuarios", "Complejidad regulatoria creciente (50 países)", "Dependencia tecnológica de terceros (Visa, AWS)", "Modelo monetización no probado fuera Europa", "Sanción AML 3,5M€ Lituania (2025)"] },
    { x: 0.4, y: 3.35, color: C.teal, label: "OPORTUNIDADES", items: ["Expansión USA, India, mercados emergentes", "IA y automatización financiera avanzada", "Open Banking PSD2: acceso a datos banca", "MiCAR: marco legal para cripto", "Segmento 18-34 años: 50% base clientes", "Licencia UK 2024: palanca de confianza"] },
    { x: 5.15, y: 3.35, color: "#E85C50", label: "AMENAZAS", items: ["BigTech sin licencia bancaria (Google, Apple)", "DORA + MiCAR: coste compliance permanente", "Imitación acelerada: BBVA IA, CaixaBank GalaxIA", "Saturación y presión de márgenes en Europa", "Riesgos ciberseguridad y continuidad servicio", "6 entidades controlan 75% depósitos España"] },
  ];
  quad.forEach(({ x, y, color, label, items }) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: 4.65, h: 1.95,
      fill: { color: C.white },
      shadow: { type: "outer", color: "000000", blur: 5, offset: 2, angle: 45, opacity: 0.09 },
      rectRadius: 0.1, line: { color, width: 1.5 }
    });
    s.addText(label, { x: x + 0.12, y: y + 0.1, w: 4.4, h: 0.28, fontSize: 10, fontFace: fontTitle, color, bold: true });
    items.forEach((item, j) => {
      s.addText([{ text: "• ", options: { color, bold: true } }, { text: item }], {
        x: x + 0.12, y: y + 0.38 + j * 0.245, w: 4.4, h: 0.23,
        fontSize: 8.5, fontFace: fontBody, color: C.darktext
      });
    });
  });
}

// ─── Slide 15: CONCLUSIONES ──────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  addPill(s, "CONCLUSIONES", 0.5, 0.18, 2.5);
  s.addText("Disrupción lograda · Consolidación pendiente", {
    x: 0.5, y: 0.52, w: 9, h: 0.55,
    fontSize: 22, fontFace: fontTitle, color: C.teal, bold: true
  });

  const conclusions = [
    { num: "5.1", title: "Customer focus como disrupción", body: "Revolut rediseñó la cadena de valor bancaria desde el cliente: eliminó la fricción que la banca generaba estructuralmente (comisiones, burocracia, tiempos de espera)", color: C.teal },
    { num: "5.2", title: "Ventaja competitiva sistémica (VRIN)", body: "La integración de eficiencia digital + datos acumulados + modelo freemium + velocidad de innovación cumple los 4 criterios VRIN de forma interdependiente", color: C.violet },
    { num: "5.3", title: "Posicionamiento híbrido", body: "19,8% nuevas altas (España) pero < 0,3% banco principal. Sin primacía, el posicionamiento híbrido no es sostenible a largo plazo (Bowman: zona híbrida estable)", color: C.gold },
    { num: "5.4–5.5", title: "Limitaciones estructurales", body: "Coste regulatorio permanente y creciente · Dependencia de Visa, AWS y APIs de terceros · Modelo monetización no probado en mercados no europeos", color: "#E85C50" },
    { num: "5.8", title: "Confianza: última frontera estratégica", body: "Licencia bancaria UK (2024) es el primer paso. La confianza institucional no puede acelerarse con tecnología: se construye con historial en momentos de crisis", color: "#50B5A0" },
  ];
  conclusions.forEach(({ num, title, body, color }, i) => {
    const y = 1.18 + i * 0.84;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.4, y, w: 9.2, h: 0.78,
      fill: { color: "111E30" }, rectRadius: 0.1,
      line: { color, width: 1 }
    });
    s.addText(num, { x: 0.5, y: y + 0.05, w: 0.6, h: 0.6, fontSize: 10, fontFace: fontTitle, color, bold: true, align: "center", valign: "middle" });
    s.addShape(pres.shapes.LINE, { x: 1.12, y: y + 0.12, w: 0, h: 0.52, line: { color, width: 1 } });
    s.addText(title, { x: 1.25, y: y + 0.08, w: 2.5, h: 0.28, fontSize: 10, fontFace: fontTitle, color, bold: true });
    s.addText(body, { x: 1.25, y: y + 0.35, w: 8.1, h: 0.38, fontSize: 9, fontFace: fontBody, color: C.lightgray });
  });
}

// ─── Slide 16: CIERRE ────────────────────────────────────────────────────────
{
  const s = darkSlide(pres);
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.teal }, line: { color: C.teal } });
  s.addShape(pres.shapes.OVAL, { x: 5.5, y: -0.8, w: 5.5, h: 5.5, fill: { color: C.violet, transparency: 88 }, line: { color: C.violet, transparency: 88 } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 2.5, w: 5, h: 5, fill: { color: C.teal, transparency: 90 }, line: { color: C.teal, transparency: 90 } });

  s.addText("\"El futuro de Revolut se jugará en el terreno\nde la confianza institucional.\"", {
    x: 0.8, y: 0.8, w: 8.4, h: 1.8,
    fontSize: 24, fontFace: fontTitle, color: C.white, italic: true, align: "center"
  });
  s.addShape(pres.shapes.LINE, { x: 3, y: 2.75, w: 4, h: 0, line: { color: C.teal, width: 1.5 } });

  s.addText("REVOLUT", { x: 0.8, y: 2.95, w: 8.4, h: 0.7, fontSize: 36, fontFace: fontTitle, color: C.teal, bold: true, align: "center", charSpacing: 8 });
  s.addText("Disrupción lograda · Consolidación pendiente", { x: 0.8, y: 3.65, w: 8.4, h: 0.4, fontSize: 15, fontFace: fontTitle, color: C.midgray, italic: true, align: "center" });

  s.addShape(pres.shapes.LINE, { x: 3, y: 4.2, w: 4, h: 0, line: { color: C.midgray, width: 0.5 } });

  s.addText("Jessica Elizabeth Rojas Rodríguez\nMBA · Universidad de Valladolid · 2025–2026", {
    x: 0.8, y: 4.35, w: 8.4, h: 0.7,
    fontSize: 11, fontFace: fontBody, color: C.midgray, align: "center"
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.33, w: 10, h: 0.29, fill: { color: "060D16" }, line: { color: "060D16" } });
  s.addText("Facultad de Ciencias Económicas y Empresariales · Universidad de Valladolid", {
    x: 0, y: 5.33, w: 10, h: 0.29,
    fontSize: 8.5, fontFace: fontBody, color: C.midgray, align: "center", valign: "middle"
  });
}

pres.writeFile({ fileName: "/mnt/user-data/outputs/Revolut_TFM_Presentacion.pptx" })
  .then(() => console.log("✅ Presentación creada"))
  .catch(e => console.error("❌", e));
  }
});

