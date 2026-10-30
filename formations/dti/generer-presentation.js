// Génère la présentation de Division Terrestre Internationale (pptxgenjs + icônes react-icons).
// Usage : NODE_PATH=<node_modules avec pptxgenjs, react-icons, react, react-dom, sharp> node generer-presentation.js <apply_theme.js>
const pptxgen = require("pptxgenjs");
const React = require("react");
const { renderToStaticMarkup } = require("react-dom/server");
const sharp = require("sharp");
const JSZip = require("jszip");
const fa = require("react-icons/fa6");
const { applyTheme } = require(process.argv[2]);
const OUT = __dirname + "/presentation-division-terrestre-internationale.pptx";

// Bleu nuit dominant, bleus d'appui, rouge en simple touche
const HEX = { navy: "1E2340", blue: "3B5BA9", sky: "A9BCEB", tint: "EEF2FA", line: "E1E6F0", red: "C60007", green: "3F8A6A", grey: "6B7584", ink: "1B2430", white: "FFFFFF", bg: "F7F9FD" };
const THEME = { name: "DTI", headFontFace: "Arial", bodyFontFace: "Calibri",
  colors: { dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.tint, accent1: HEX.blue, accent2: HEX.sky,
    accent3: HEX.green, accent4: HEX.red, accent5: HEX.grey, accent6: HEX.line, hlink: HEX.blue, folHlink: HEX.grey } };

// --- Images générées : icônes et fonds (route stylisée)
const ICONS = {};
const png64 = (buf) => "image/png;base64," + buf.toString("base64");
async function loadIcons(list) {
  for (const [name, hex] of list) {
    if (ICONS[name + hex]) continue;
    const svg = renderToStaticMarkup(React.createElement(fa[name], { color: "#" + hex, size: 256 }));
    ICONS[name + hex] = png64(await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer());
  }
}
const ic = (name, hex) => ICONS[name + hex];
const svgPng = async (svg) => png64(await sharp(Buffer.from(svg)).png().toBuffer());
const BG_DARK = `<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080">
<defs><linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#1E2340"/><stop offset="1" stop-color="#2B3A72"/></linearGradient>
<radialGradient id="r" cx="0.78" cy="0.4" r="0.5"><stop offset="0" stop-color="#3B5BA9" stop-opacity="0.55"/><stop offset="1" stop-color="#3B5BA9" stop-opacity="0"/></radialGradient></defs>
<rect width="1920" height="1080" fill="url(#g)"/><rect width="1920" height="1080" fill="url(#r)"/>
<path d="M0 820 C 520 640, 980 1000, 1920 700" stroke="#FFFFFF" stroke-opacity="0.06" stroke-width="2" fill="none"/>
<path d="M0 900 C 600 760, 1100 1060, 1920 820" stroke="#FFFFFF" stroke-opacity="0.05" stroke-width="2" fill="none"/>
<path d="M1010 1080 L1455 640 L1545 640 L1990 1080 Z" fill="#FFFFFF" fill-opacity="0.06"/>
<path d="M1500 650 L1500 1080" stroke="#FFFFFF" stroke-opacity="0.35" stroke-width="7" stroke-dasharray="46 34"/>
<path d="M1455 640 L1010 1080 M1545 640 L1990 1080" stroke="#A9BCEB" stroke-opacity="0.35" stroke-width="3"/></svg>`;
const BG_LIGHT = `<svg xmlns="http://www.w3.org/2000/svg" width="1920" height="1080">
<rect width="1920" height="1080" fill="#F7F9FD"/>
<path d="M0 1010 C 500 900, 900 1060, 1920 880" stroke="#E3E9F5" stroke-width="2" fill="none"/>
<path d="M1150 1080 C 1450 930, 1650 860, 1920 830" stroke="#EAEFF8" stroke-width="64" fill="none"/>
<path d="M1150 1080 C 1450 930, 1650 860, 1920 830" stroke="#FFFFFF" stroke-width="4" stroke-dasharray="26 22" fill="none"/></svg>`;

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "Division Terrestre Internationale";
  const C = pres.SchemeColor;
  const bgDark = await svgPng(BG_DARK), bgLight = await svgPng(BG_LIGHT);

  pres.defineSlideMaster({ title: "SOMBRE", background: { data: bgDark },
    objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.0, w: 7.2, h: 1.9, fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left" }, text: "" } }] });
  pres.defineSlideMaster({ title: "IMPACT", background: { data: bgDark },
    slideNumber: { x: 12.25, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent2, align: "right" },
    objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 0.75, w: 11.7, h: 0.9, fontSize: 32, bold: true, color: C.background1, valign: "middle", align: "left" }, text: "" } }] });
  pres.defineSlideMaster({ title: "CONTENU", background: { data: bgLight },
    slideNumber: { x: 12.25, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
    objects: [
      { image: { data: ic("FaTruckFast", HEX.grey), x: 0.6, y: 7.05, w: 0.22, h: 0.22 } },
      { text: { text: "Division Terrestre Internationale", options: { x: 0.9, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent5 } } },
      { placeholder: { options: { name: "title", type: "title", x: 1.6, y: 0.62, w: 11.1, h: 0.7, fontSize: 28, bold: true, color: C.text2, valign: "middle", align: "left" }, text: "" } },
    ] });

  let n = 0;
  const A = (step) => ({ objectName: `anim${step}-${++n}` }); // repère d'animation (étape d'apparition)
  const shadow = () => ({ type: "outer", color: HEX.navy, opacity: 0.08, blur: 10, offset: 3, angle: 90 });
  function content(title, sec, icon) {
    const s = pres.addSlide({ masterName: "CONTENU", sectionTitle: sec });
    s.addText(title, { placeholder: "title" });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 0.45, w: 0.8, h: 0.8, rectRadius: 0.15, fill: { color: C.text2 }, line: { color: C.text2 } });
    s.addImage({ data: ic(icon, HEX.white), x: 0.8, y: 0.65, w: 0.4, h: 0.4 });
    s.addShape(pres.shapes.OVAL, { x: 1.62, y: 0.42, w: 0.1, h: 0.1, fill: { color: C.accent4 }, line: { color: C.accent4 } });
    s.addText(sec.toUpperCase(), { isTextBox: true, x: 1.8, y: 0.3, w: 8, h: 0.34, fontSize: 10, bold: true, color: C.accent5, charSpacing: 2, margin: 0, valign: "middle" });
    return s;
  }
  function tx(s, t, o) { s.addText(t, Object.assign({ isTextBox: true, fontSize: 15, color: C.text1, valign: "top" }, o)); }
  function panel(s, x, y, w, h, fill, step, withShadow) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { color: fill === C.background1 ? C.accent6 : fill, width: 0.75 } },
      withShadow ? { shadow: shadow() } : {}, step ? A(step) : {}));
  }
  function iconTile(s, icon, x, y, d, fill, iconHex, step) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w: d, h: d, rectRadius: d * 0.22, fill: { color: fill }, line: { color: fill } }, step ? A(step) : {}));
    s.addImage(Object.assign({ data: ic(icon, iconHex), x: x + d * 0.26, y: y + d * 0.26, w: d * 0.48, h: d * 0.48 }, step ? A(step) : {}));
  }
  // Carte blanche avec icône sur fond bleu pâle
  function card(s, x, y, w, h, icon, head, body, step, iconHex) {
    panel(s, x, y, w, h, C.background1, step, true);
    iconTile(s, icon, x + 0.22, y + 0.22, 0.6, C.background2, iconHex || HEX.blue, step);
    tx(s, [{ text: head, options: { bold: true, fontSize: 16, color: C.text2, breakLine: true, paraSpaceAfter: 4 } }, { text: body, options: { fontSize: 14, color: C.text1 } }],
      Object.assign({ x: x + 0.98, y: y + 0.14, w: w - 1.15, h: h - 0.28 }, A(step)));
  }
  // Diapositive « module »
  function module(title, sec, icon, menu, intro, feats, benefit, notes, panelTitle) {
    const s = content(title, sec, icon);
    tx(s, [{ text: "Menu  ", options: { bold: true, color: C.accent1 } }, { text: menu, options: { color: C.accent5 } }], { x: 0.6, y: 1.5, w: 8.4, h: 0.32, fontSize: 12, valign: "middle", margin: 0 });
    tx(s, intro, { x: 0.6, y: 1.88, w: 8.4, h: 0.5, fontSize: 17, color: C.text2, margin: 0 });
    feats.forEach((f, i) => card(s, 0.6 + (i % 2) * 4.3, 2.6 + Math.floor(i / 2) * 2.05, 4.1, 1.85, f[0], f[1], f[2], i + 1));
    const st = feats.length + 1;
    panel(s, 9.35, 1.5, 3.4, 5.2, C.text2, st, true);
    tx(s, panelTitle || "À quoi ça sert", Object.assign({ x: 9.6, y: 1.7, w: 3.0, h: 0.5, fontSize: 18, bold: true, color: C.accent2, margin: 0 }, A(st)));
    benefit.forEach((b, i) => { const y = 2.45 + i * 1.3;
      s.addImage(Object.assign({ data: ic("FaCircleCheck", HEX.sky), x: 9.6, y: y + 0.04, w: 0.32, h: 0.32 }, A(st + 1 + i)));
      tx(s, b, Object.assign({ x: 10.05, y, w: 2.55, h: 1.15, fontSize: 14, color: C.background1, margin: 0 }, A(st + 1 + i))); });
    s.addNotes(notes);
    return s;
  }
  // Grand camion sur la route (diapositives sombres)
  function hero(s, icon) {
    s.addShape(pres.shapes.OVAL, { x: 9.2, y: 0.9, w: 3.9, h: 3.9, fill: { color: C.accent1, transparency: 55 }, line: { color: C.accent2, transparency: 60, width: 1 } });
    s.addImage(Object.assign({ data: ic(icon, HEX.white), x: 10.05, y: 1.75, w: 2.2, h: 2.2 }, A(1)));
  }

  // 1 — Titre
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "SOMBRE", sectionTitle: "Introduction" });
  hero(s, "FaTruckFast");
  s.addShape(pres.shapes.OVAL, { x: 0.8, y: 1.55, w: 0.12, h: 0.12, fill: { color: C.accent4 }, line: { color: C.accent4 } });
  tx(s, "PRISE EN MAIN DU LOGICIEL", { x: 1.0, y: 1.45, w: 6, h: 0.32, fontSize: 12, bold: true, color: C.accent2, charSpacing: 3, margin: 0, valign: "middle" });
  s.addText("Division Terrestre Internationale", { placeholder: "title" });
  tx(s, "Gérer les camions, les chauffeurs, les trajets et l'argent de l'entreprise", Object.assign({ x: 0.8, y: 4.05, w: 7.2, h: 0.9, fontSize: 20, color: C.accent2 }, A(2)));
  // Présentateur et date
  iconTile(s, "FaUserTie", 0.8, 5.4, 0.7, C.accent1, HEX.white, 3);
  tx(s, [{ text: "PRÉSENTÉ PAR", options: { fontSize: 10, bold: true, color: C.accent2, charSpacing: 2, breakLine: true } },
    { text: "Armel Meledje", options: { fontSize: 19, bold: true, color: C.background1, breakLine: true } },
    { text: "Développeur Senior – Ingénieur de conception", options: { fontSize: 13, color: C.accent2 } }], Object.assign({ x: 1.7, y: 5.3, w: 4.6, h: 1.1, margin: 0, valign: "middle" }, A(3)));
  s.addShape(pres.shapes.LINE, Object.assign({ x: 6.15, y: 5.4, w: 0, h: 0.9, line: { color: HEX.sky, width: 1, transparency: 50 } }, A(4)));
  iconTile(s, "FaCalendarDays", 6.4, 5.4, 0.7, C.accent1, HEX.white, 4);
  tx(s, [{ text: "DATE", options: { fontSize: 10, bold: true, color: C.accent2, charSpacing: 2, breakLine: true } },
    { text: "Vendredi 30 octobre 2026", options: { fontSize: 17, bold: true, color: C.background1 } }], Object.assign({ x: 7.3, y: 5.3, w: 3.0, h: 1.1, margin: 0, valign: "middle" }, A(4)));
  s.addNotes("Bonjour à tous. Vous utilisez déjà le logiciel Division Terrestre Internationale. Aujourd'hui, on va voir ensemble comment bien l'utiliser, tous les jours. D'abord, on fait le tour des rubriques. Ensuite, on parle des bons réflexes. Puis je vous montre un exemple en direct. À la fin, vous pourrez poser toutes vos questions.");

  // 2 — Pourquoi
  s = content("Pourquoi bien l'utiliser au quotidien ?", "Introduction", "FaLightbulb");
  tx(s, "Quand on oublie de noter", Object.assign({ x: 0.6, y: 1.5, w: 5.6, h: 0.45, fontSize: 19, bold: true, color: C.accent5 }, A(1)));
  [["FaShuffle", "Informations introuvables", "Une information restée dans un cahier, les autres ne peuvent pas la voir."],
   ["FaHourglassHalf", "Aucune alerte", "Une date de fin pas notée, c'est une assurance ou un permis qui expire sans prévenir."],
   ["FaCircleQuestion", "Chiffres faux", "Un plein ou une dépense oubliés, et les chiffres du logiciel sont faux."]]
    .forEach((d, i) => card(s, 0.6, 2.05 + i * 1.6, 5.6, 1.4, d[0], d[1], d[2], 2 + i, HEX.grey));
  s.addShape(pres.shapes.RIGHT_ARROW, Object.assign({ x: 6.42, y: 3.85, w: 0.55, h: 0.55, fill: { color: C.accent1 }, line: { color: C.accent1 } }, A(5)));
  tx(s, "Quand chacun l'utilise bien", Object.assign({ x: 7.2, y: 1.5, w: 5.6, h: 0.45, fontSize: 19, bold: true, color: C.accent1 }, A(6)));
  [["FaLayerGroup", "Tout est retrouvé en quelques secondes", "Véhicules, chauffeurs, trajets, clients, carburant, entretien et finances."],
   ["FaBell", "Les alertes marchent", "Papiers périmés signalés, entretiens rappelés au bon moment."],
   ["FaChartLine", "Des chiffres justes", "On sait ce que rapporte et ce que coûte chaque camion."]]
    .forEach((d, i) => card(s, 7.2, 2.05 + i * 1.6, 5.55, 1.4, d[0], d[1], d[2], 7 + i));
  s.addNotes("Le logiciel est déjà là. Mais il nous aide seulement si on y note bien les choses. Si on oublie de noter une date ou un plein, le logiciel ne peut pas nous prévenir, et les chiffres sont faux. Si chacun note bien, on retrouve tout vite, les alertes marchent et les chiffres sont justes. Question à la salle : aujourd'hui, est-ce que tout est bien noté ?");

  // 3 — Modules
  s = content("18 rubriques, 6 grandes familles", "Introduction", "FaSitemap");
  s.addShape(pres.shapes.OVAL, Object.assign({ x: 0.45, y: 2.25, w: 3.0, h: 3.0, fill: { color: C.background2 }, line: { color: C.accent6 } }, A(1)));
  s.addShape(pres.shapes.OVAL, Object.assign({ x: 0.75, y: 2.55, w: 2.4, h: 2.4, fill: { color: C.text2 }, line: { color: C.text2 } }, A(1)));
  tx(s, [{ text: "18", options: { fontSize: 60, bold: true, breakLine: true } }, { text: "rubriques", options: { fontSize: 17, color: C.accent2 } }], Object.assign({ x: 0.75, y: 2.55, w: 2.4, h: 2.4, align: "center", valign: "middle", color: C.background1 }, A(1)));
  [["FaGaugeHigh", "Suivre l'activité", "Dashboard · Availability · Reports"], ["FaTruck", "Les camions", "Vehicle's · Vehicle Vendor's · Tracking · Geofence"], ["FaUsers", "Les personnes", "Driver's · Employee"],
   ["FaRoute", "Clients et trajets", "Trips · Customer · Coupon Management"], ["FaWrench", "Entretien et carburant", "Maintenance · Parts Stock · Fuel · Reminder"], ["FaCoins", "Argent et réglages", "Accounts · Setting's"]]
    .forEach((d, i) => card(s, 3.8 + (i % 2) * 4.5, 1.6 + Math.floor(i / 2) * 1.75, 4.3, 1.55, d[0], d[1], d[2], 2 + i));
  s.addNotes("Dans le menu à gauche de l'écran, il y a 18 rubriques. Pour s'y retrouver, je les ai rangées en 6 familles. Les noms dans le menu sont en anglais : sur la page suivante, je vous donne leur traduction.");



  // Petit lexique : les mots anglais du menu
  s = content("Petit lexique : les mots anglais du menu", "Introduction", "FaLanguage");
  [["Dashboard", "Tableau de bord"], ["Availability", "Disponibilité des camions"], ["Vehicle's", "Véhicules"], ["Driver's", "Chauffeurs"],
   ["Vehicle Vendor's", "Loueurs de véhicules"], ["Trips", "Trajets"], ["Customer", "Clients"], ["Coupon", "Code de réduction"],
   ["Maintenance", "Entretien"], ["Parts Stock", "Stock de pièces"], ["Fuel", "Carburant"], ["Reminder", "Rappel"],
   ["Accounts", "Comptes (argent)"], ["Tracking", "Suivi sur la carte"], ["Geofence", "Zone sur la carte"], ["Reports", "Rapports"],
   ["Employee", "Employés"], ["Setting's", "Réglages"], ["Add / Save", "Ajouter / Enregistrer"], ["Edit / Delete", "Modifier / Supprimer"]]
    .forEach((d, i) => { const x = 0.6 + (i % 4) * 3.05, y = 1.55 + Math.floor(i / 4) * 1.03, st = Math.floor(i / 4) + 1;
      panel(s, x, y, 2.9, 0.88, C.background1, st, true);
      tx(s, [{ text: d[0], options: { bold: true, fontSize: 14, color: C.accent1, breakLine: true } }, { text: d[1], options: { fontSize: 14, color: C.text2 } }],
        Object.assign({ x: x + 0.2, y, w: 2.6, h: 0.88, valign: "middle", margin: 0 }, A(st))); });
  s.addNotes("Le logiciel affiche ses menus en anglais. Voici la traduction des mots que vous verrez le plus souvent. Pas besoin de tout retenir : gardez cette page sous la main au début.");
  // 4 — Le parcours d'un trajet (récit)
  s = pres.addSlide({ masterName: "IMPACT", sectionTitle: "Introduction" });
  s.addText("Un trajet, du premier appel au dernier franc", { placeholder: "title" });
  tx(s, "Chaque étape du parcours d'un camion a sa place dans le logiciel.", Object.assign({ x: 0.8, y: 1.65, w: 11, h: 0.5, fontSize: 18, color: C.accent2 }, A(1)));
  s.addShape(pres.shapes.LINE, Object.assign({ x: 0.9, y: 3.85, w: 11.6, h: 0, line: { color: "3A4A82", width: 30, cap: "round" } }, A(1)));
  s.addShape(pres.shapes.LINE, Object.assign({ x: 1.0, y: 3.85, w: 11.4, h: 0, line: { color: HEX.sky, width: 2, dashType: "dash" } }, A(1)));
  [["FaPhone", "Réserver", "Client, véhicule disponible, chauffeur"], ["FaTruckFast", "Rouler", "Voir où est le camion sur la carte"], ["FaGasPump", "Faire le plein", "Chaque plein noté pour son camion"],
   ["FaWrench", "Entretenir", "Ne pas oublier l'entretien, suivre les pièces"], ["FaCoins", "Encaisser", "Noter l'argent qui entre et qui sort"], ["FaChartLine", "Faire le point", "Voir ce qui rapporte et ce qui coûte"]]
    .forEach((d, i) => { const cx = 1.45 + i * 2.08, st = i + 2;
      s.addShape(pres.shapes.OVAL, Object.assign({ x: cx - 0.62, y: 3.23, w: 1.24, h: 1.24, fill: { color: i === 1 ? C.accent1 : C.background1 }, line: { color: "3A4A82", width: 4 } }, A(st)));
      s.addImage(Object.assign({ data: ic(d[0], i === 1 ? HEX.white : HEX.navy), x: cx - 0.29, y: 3.56, w: 0.58, h: 0.58 }, A(st)));
      tx(s, [{ text: d[1], options: { bold: true, fontSize: 18, color: C.background1, breakLine: true, paraSpaceAfter: 4 } }, { text: d[2], options: { fontSize: 13, color: C.accent2 } }], Object.assign({ x: cx - 1.0, y: 4.75, w: 2.0, h: 1.3, align: "center" }, A(st))); });
  s.addNotes("Prenons l'histoire d'un trajet. Un client appelle. On réserve un camion libre et un chauffeur. Le camion roule, et on le voit sur la carte. Il fait le plein. Il passe à l'entretien. Le client paie. À la fin, on regarde ce que ce trajet a rapporté. Chaque étape a sa place dans le logiciel : on va les voir une par une.");

  // Modules
  pres.addSection({ title: "Suivre l'activité" });
  module("Le tableau de bord : la journée en un coup d'œil", "Suivre l'activité", "FaGaugeHigh", "Dashboard · Availability",
    "Dès qu'on se connecte, on voit l'essentiel de la journée.",
    [["FaHashtag", "Les chiffres du jour", "Nombre de camions, chauffeurs, clients, trajets ; argent entré et sorti."], ["FaTruckFront", "Où sont les camions", "Camions libres ou en route, et leur position."],
     ["FaBell", "Rappels et classements", "Entretiens à faire ; ce qui rapporte et ce qui coûte le plus."], ["FaCalendarDays", "Calendrier des camions", "Quels camions sont pris, par mois, semaine ou jour."]],
    ["La direction voit l'activité sans attendre", "Les agents voient tout de suite les camions libres", "On voit tout de suite ce qui est urgent"],
    "Le tableau de bord, c'est la première page quand on se connecte. On y voit les chiffres du jour, les camions libres ou en route, et ce qui est urgent. Avant de réserver un trajet, regardez toujours le calendrier (Availability) pour voir quels camions sont libres.");

  pres.addSection({ title: "Les camions" });
  module("Gérer les véhicules", "Les camions", "FaTruck", "Vehicle's › Vehicle List · Add Vehicle · Vehicle Group · Route",
    "Chaque camion a sa fiche, comme une carte d'identité.",
    [["FaIdCard", "Fiche du camion", "Immatriculation, modèle, kilométrage, photo et papiers."], ["FaCalendarXmark", "Dates de fin", "Immatriculation et assurance : en rouge quand la date est passée."],
     ["FaLayerGroup", "Types et routes", "Camion, bus, voiture… et la liste des routes habituelles."], ["FaTag", "Prix et propriétaire", "Prix habituel, camion à nous ou loué, boîtier GPS."]],
    ["On ne roule plus avec une assurance périmée", "Ajouter beaucoup de camions d'un coup avec un fichier tableur", "Sortir la liste en Excel ou en PDF"],
    "Chaque camion a sa fiche, comme une carte d'identité. Regardez les dates de fin de l'assurance et de l'immatriculation : quand la date est passée, elle s'affiche en rouge. Sur toutes les listes, les petits boutons Excel et PDF permettent de sortir la liste. On peut aussi ajouter beaucoup de camions d'un coup avec un fichier tableur.");

  module("Chauffeurs et loueurs de camions", "Les camions", "FaUserTie", "Driver's · Vehicle Vendor's",
    "Savoir qui conduit, avec quel permis, et à qui on loue des camions.",
    [["FaAddressCard", "Fiche chauffeur", "Contact, âge, adresse, expérience, date d'entrée, photo et document."], ["FaIdBadge", "Permis", "Numéro et date de fin, en rouge quand elle est passée."],
     ["FaHandshake", "Loueurs de camions", "Société, personne à appeler, téléphone, adresse."], ["FaFileContract", "Contrats", "Date et copie du contrat."]],
    ["Un permis périmé se voit tout de suite", "Tous les contrats des loueurs au même endroit", "Ajouter beaucoup de chauffeurs d'un coup avec un fichier tableur"],
    "Pour chaque chauffeur, on note son permis et sa date de fin. Quand un permis est périmé, il s'affiche en rouge : il ne faut pas lui confier de camion. Les loueurs, ce sont les entreprises qui nous louent des camions ; on garde leurs contrats ici.");

  module("Voir où sont les camions", "Les camions", "FaMapLocationDot", "Tracking › History Tracking · Live Location · Geofence",
    "Voir les camions sur une carte, et être prévenu quand ils entrent ou sortent d'un lieu.",
    [["FaLocationCrosshairs", "En ce moment", "Où se trouve chaque camion, sur la carte."], ["FaClockRotateLeft", "Les trajets passés", "Le chemin pris par un camion, sur les dates choisies."],
     ["FaDrawPolygon", "Zones sur la carte", "On dessine un lieu sur la carte : dépôt, port, client…"], ["FaBullhorn", "Alertes de zone", "On est prévenu quand un camion entre ou sort de ce lieu."]],
    ["Vérifier la route vraiment prise", "Réagir vite si un camion sort de sa route", "Dire au client où en est sa livraison"],
    "Ici, on voit les camions sur une carte. Attention : il faut qu'un petit boîtier GPS soit installé dans le camion. On peut voir où est le camion maintenant, ou le chemin qu'il a pris avant. On peut aussi dessiner un lieu sur la carte, comme le dépôt, et être prévenu quand un camion y entre ou en sort.");

  pres.addSection({ title: "Clients et trajets" });
  s = module("Trajets, clients et réductions", "Clients et trajets", "FaRoute", "Trips · Customer · Coupon Management",
    "De la réservation à la facture, chaque trajet est suivi.",
    [["FaClipboardList", "Réservation", "Client, dates, camion, chauffeur, départ, arrivée, kilomètres, arrêts."], ["FaSignal", "Suivi du trajet", "Où en est le trajet (réservé, commencé…), distance, numéro."],
     ["FaUsers", "Clients", "Coordonnées, ce qu'il reste à payer, client actif ou non."], ["FaTicket", "Codes de réduction", "Montant de la réduction, dates, nombre d'utilisations."]],
    ["Chaque trajet lié à un client, un camion, un chauffeur", "On voit qui n'a pas encore payé", "Le client reçoit sa réservation par e-mail ou SMS"],
    "C'est ici qu'on crée les trajets. On choisit le client, le camion, le chauffeur, le départ et l'arrivée. On peut faire payer de 9 façons : un prix fixe, à la tonne, au kilo, au kilomètre, au trajet, au jour, à l'heure, au litre ou au sac. Le client peut recevoir sa réservation par e-mail ou par SMS. On voit aussi ce que chaque client doit encore payer.");
  tx(s, [{ text: "9 façons de faire payer : ", options: { bold: true, color: C.accent1 } }, { text: "fixe · tonne · kg · km · trajet · jour · heure · litre · sac", options: { color: C.text2 } }], Object.assign({ x: 0.6, y: 6.62, w: 8.6, h: 0.3, fontSize: 13, margin: 0 }, A(9)));

  pres.addSection({ title: "Entretien et carburant" });
  module("Entretien, pièces et rappels", "Entretien et carburant", "FaWrench", "Maintenance · Parts Stock · Reminder",
    "Prévoir l'entretien avant la panne, et connaître son coût.",
    [["FaScrewdriverWrench", "Entretiens", "Camion, dates, mécanicien, garage, coût et pièces utilisées."], ["FaUserGear", "Mécaniciens et garages", "La liste des personnes et garages qui réparent."],
     ["FaBoxesStacked", "Stock de pièces", "Pièces en réserve, prix, achats payés ou non."], ["FaBell", "Rappels", "Ce qu'il faut faire et quand ; on coche une fois fait."]],
    ["Moins de pannes grâce aux rappels", "On sait combien coûte l'entretien de chaque camion", "On sait quelles pièces on a et ce qu'on a acheté"],
    "Ici, on note chaque entretien : quel camion, quel mécanicien, combien ça a coûté et quelles pièces on a utilisées. Les rappels servent à ne pas oublier une vidange ou une révision. Ils apparaissent aussi sur le tableau de bord.");

  module("Carburant", "Entretien et carburant", "FaGasPump", "Fuel › Fuel Management · Add Fuel · Fuel Vendor",
    "Chaque plein est noté, avec le camion et le chauffeur.",
    [["FaFillDrip", "Pleins", "Date, camion, chauffeur, litres, prix, kilométrage, remarque."], ["FaArrowRightArrowLeft", "D'où vient le carburant", "Notre cuve ou une station."],
     ["FaOilCan", "Cuve", "Remplir la cuve et voir combien de litres il reste."], ["FaStore", "Stations", "La liste des stations et vendeurs de carburant."]],
    ["Voir si un camion consomme trop", "Savoir combien il reste dans la cuve", "Un bilan du carburant sur les dates choisies"],
    "À chaque plein, on note les litres, le prix et le kilométrage du camion. Grâce au kilométrage, on peut voir si un camion consomme trop. Le carburant peut venir de notre cuve ou d'une station : pour la cuve, le logiciel compte combien de litres il reste.");

  pres.addSection({ title: "Argent et rapports" });
  module("L'argent : entrées et sorties", "Argent et rapports", "FaCoins", "Accounts › Accounts · Add Transactions · Transactions · Ac Transfers / Deposit · Category",
    "Chaque entrée et chaque sortie d'argent est notée, avec le camion concerné.",
    [["FaBuildingColumns", "Comptes", "Les caisses et comptes de l'entreprise, avec ce qu'il y a dedans."], ["FaReceipt", "Entrées et sorties", "Recette ou dépense : compte, type, camion, montant."],
     ["FaMoneyBillTransfer", "Virements", "Passer de l'argent d'un compte à un autre."], ["FaTags", "Types", "Recettes de trajet, carburant, salaires…"]],
    ["On voit ce que rapporte chaque camion", "On retrouve tout par date, avec les totaux", "Des types adaptés à notre entreprise"],
    "Ici, on note tout l'argent qui entre et qui sort : les recettes des trajets, le carburant, les salaires… Chaque opération est liée à un camion. Comme ça, on voit ce que chaque camion rapporte. On peut chercher par dates, et le logiciel fait les totaux tout seul.");

  // 12 — Rapports
  s = content("11 rapports prêts à l'emploi", "Argent et rapports", "FaChartPie");
  tx(s, "On choisit les dates, puis on clique sur « Generate Report ».", { x: 0.6, y: 1.55, w: 12, h: 0.45, fontSize: 17, color: C.text2 });
  [["FaClipboardList", "Réservations"], ["FaScaleBalanced", "Revenus et dépenses"], ["FaGasPump", "Carburant"], ["FaUserTie", "Chauffeurs"], ["FaTicket", "Coupons"], ["FaBell", "Rappels"],
   ["FaWrench", "Entretien"], ["FaArrowTrendUp", "Top revenus"], ["FaArrowTrendDown", "Top dépenses"], ["FaTruck", "Top véhicules"], ["FaRoute", "Top routes"]]
    .forEach((r, i) => { const x = 0.6 + (i % 4) * 3.1, y = 2.2 + Math.floor(i / 4) * 1.5, top = i >= 7;
      panel(s, x, y, 2.9, 1.25, top ? C.text2 : C.background1, i + 1, true);
      iconTile(s, r[0], x + 0.22, y + 0.3, 0.65, top ? C.accent1 : C.background2, top ? HEX.white : HEX.blue, i + 1);
      tx(s, r[1], Object.assign({ x: x + 1.02, y, w: 1.8, h: 1.25, fontSize: 15, bold: true, valign: "middle", color: top ? C.background1 : C.text2 }, A(i + 1))); });
  tx(s, "Les « Top » montrent ce qui rapporte ou coûte le plus.", { x: 9.9, y: 5.2, w: 2.85, h: 1.25, fontSize: 13, italic: true, color: C.accent5, valign: "middle" });
  s.addNotes("Les rapports sont des bilans tout prêts. On choisit les dates, on clique, et le logiciel fait le travail. Les rapports « Top » montrent ce qui rapporte le plus et ce qui coûte le plus. Plus besoin de faire les calculs à la main.");

  pres.addSection({ title: "Comptes et réglages" });
  module("Les comptes des employés", "Comptes et réglages", "FaUserShield", "Employee › Employee Management · Add Employee",
    "Chaque utilisateur a son compte et ne voit que ce qu'il doit voir.",
    [["FaUserPlus", "Un compte par personne", "Nom, téléphone, e-mail, identifiant et mot de passe."], ["FaListCheck", "Ce que chacun peut faire", "Des cases à cocher : voir, ajouter, modifier, supprimer."],
     ["FaCubes", "Pour chaque rubrique", "Camions, chauffeurs, argent, clients, carburant, entretien…"], ["FaToggleOn", "Compte ouvert ou bloqué", "On bloque le compte de quelqu'un qui part."]],
    ["Chacun voit seulement son travail", "Les chiffres de l'argent restent protégés", "On sait qui a noté chaque opération"],
    "Chaque personne a son propre compte, avec son mot de passe. Pour chaque compte, on coche ce que la personne peut faire : voir, ajouter, modifier ou supprimer. Comme ça, chacun voit seulement ce dont il a besoin. Quand quelqu'un quitte l'entreprise, on bloque son compte.");

  module("Les réglages", "Comptes et réglages", "FaGear", "Setting's",
    "Les réglages sont faits par l'administrateur, une fois pour toutes.",
    [["FaBuilding", "Notre entreprise", "Nom, adresse, contacts, logo, couleurs, heure."], ["FaEnvelope", "Messages", "Envoi d'e-mails et de SMS aux clients."],
     ["FaFileInvoice", "Réservations et factures", "Numéros des réservations, réglage des factures."], ["FaDatabase", "Copies de sécurité", "Une copie de toutes les informations en un clic."]],
    ["Documents aux couleurs de l'entreprise", "Les clients sont prévenus tout seuls", "Rien n'est perdu grâce aux copies"],
    "Les réglages, c'est l'administrateur qui s'en occupe. On y met le nom et le logo de l'entreprise, les messages envoyés aux clients, et on fait des copies de sécurité de toutes les informations.");



  // Sensibilisation : sécurité, confidentialité et qualité des données
  pres.addSection({ title: "Les bons réflexes" });
  s = pres.addSlide({ masterName: "IMPACT", sectionTitle: "Les bons réflexes" });
  s.addText("Bien utiliser le logiciel : les bons réflexes", { placeholder: "title" });
  tx(s, "Le logiciel protège nos informations, mais la sécurité dépend surtout de chacun de nous.", Object.assign({ x: 0.8, y: 1.65, w: 11.5, h: 0.5, fontSize: 18, color: C.accent2 }, A(1)));
  [["FaShieldHalved", "Sécurité", "Protéger son compte et l'accès au logiciel."],
   ["FaUserSecret", "Confidentialité", "Garder pour nous les informations des clients, des chauffeurs et de l'entreprise."],
   ["FaClipboardCheck", "Bien noter", "Noter juste, complet et tout de suite, pour des chiffres justes."]]
    .forEach((d, i) => { const x = 0.8 + i * 3.95, st = i + 2;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y: 2.55, w: 3.65, h: 3.6, rectRadius: 0.15, fill: { color: "28356B" }, line: { color: "4A5C9A", width: 1 } }, A(st)));
      s.addShape(pres.shapes.OVAL, Object.assign({ x: x + 1.27, y: 2.85, w: 1.1, h: 1.1, fill: { color: C.background1 }, line: { color: "3A4A82", width: 4 } }, A(st)));
      s.addImage(Object.assign({ data: ic(d[0], HEX.navy), x: x + 1.55, y: 3.13, w: 0.54, h: 0.54 }, A(st)));
      tx(s, [{ text: d[1], options: { bold: true, fontSize: 21, color: C.background1, breakLine: true, paraSpaceAfter: 6 } }, { text: d[2], options: { fontSize: 15, color: C.accent2 } }], Object.assign({ x: x + 0.25, y: 4.15, w: 3.15, h: 1.8, align: "center" }, A(st))); });
  s.addNotes("Le logiciel protège nos informations, mais il ne peut pas tout faire seul. La plupart des problèmes viennent de petits gestes : un mot de passe prêté, une liste envoyée sur WhatsApp, un plein pas noté. On va voir les bons réflexes, très simples.");

  // Diapositive « à faire / à éviter »
  function dosDonts(title, icon, intro, dos, donts, hint, notes) {
    const s = content(title, "Les bons réflexes", icon);
    tx(s, intro, { x: 0.6, y: 1.5, w: 12.1, h: 0.5, fontSize: 17, color: C.text2, margin: 0 });
    [["À faire", "FaCircleCheck", HEX.green, dos], ["À éviter", "FaCircleXmark", HEX.grey, donts]].forEach((col, c) => {
      const x = 0.6 + c * 6.15, st = c * 6 + 1;
      panel(s, x, 2.15, 5.95, 4.0, C.background1, st, true);
      tx(s, col[0], Object.assign({ x: x + 0.3, y: 2.3, w: 5.4, h: 0.5, fontSize: 19, bold: true, color: c ? C.accent5 : C.accent3, margin: 0 }, A(st)));
      col[3].forEach((t, j) => { const y = 2.95 + j * 0.78;
        s.addImage(Object.assign({ data: ic(col[1], col[2]), x: x + 0.3, y: y + 0.04, w: 0.32, h: 0.32 }, A(st + 1 + j)));
        tx(s, t, Object.assign({ x: x + 0.8, y, w: 4.95, h: 0.72, fontSize: 15, margin: 0 }, A(st + 1 + j))); }); });
    tx(s, [{ text: "Dans le logiciel  ", options: { bold: true, color: C.accent1 } }, { text: hint, options: { color: C.text2 } }], Object.assign({ x: 0.6, y: 6.35, w: 12.1, h: 0.4, fontSize: 13, margin: 0 }, A(13)));
    s.addNotes(notes);
  }

  dosDonts("Sécurité : protéger son compte", "FaShieldHalved", "Votre compte, c'est votre signature : tout ce qui est fait avec est enregistré à votre nom.",
    ["Choisir un mot de passe long et personnel", "Le changer dès qu'il a pu être vu", "Verrouiller son ordinateur quand on s'absente", "Se déconnecter sur un ordinateur partagé"],
    ["Prêter son mot de passe, même à un collègue pressé", "Noter son mot de passe sur un post-it", "Donner son mot de passe à quelqu'un qui le demande par message", "Se connecter sur un ordinateur ou un Wi-Fi qu'on ne connaît pas"],
    "le menu « Change Password » sert à changer son mot de passe.",
    "Votre compte, c'est comme votre signature. Tout ce qui est fait avec votre compte est noté à votre nom. Si vous prêtez votre mot de passe, on ne sait plus qui a fait quoi. Et attention : personne ne vous demandera jamais votre mot de passe par téléphone, par e-mail ou par WhatsApp. Question à la salle : qui a déjà prêté son mot de passe ?");

  dosDonts("Confidentialité : protéger les informations", "FaUserSecret", "Numéros des clients, permis des chauffeurs, salaires, chiffres de l'entreprise : ce sont des informations privées.",
    ["Consulter uniquement ce dont on a besoin", "Garder les fichiers sortis du logiciel dans les dossiers de l'entreprise", "Supprimer ces fichiers quand on n'en a plus besoin", "Demander l'avis de son responsable en cas de doute"],
    ["Envoyer une liste ou une photo de l'écran sur WhatsApp", "Communiquer les coordonnées d'un client ou d'un chauffeur", "Laisser un document imprimé traîner", "Parler des chiffres de l'entreprise à l'extérieur"],
    "chacun voit seulement ce que son compte permet ; on sort une liste (Excel, PDF) seulement si le travail le demande.",
    "Les numéros des clients, les permis des chauffeurs, les salaires : ce sont des informations privées. Dans le logiciel, elles sont protégées. Mais dès qu'on les sort, en fichier ou en photo de l'écran, elles ne sont plus protégées. Exemple : une liste de clients envoyée par erreur dans un groupe WhatsApp.");

  dosDonts("Bien noter : des informations justes", "FaClipboardCheck", "Les chiffres du logiciel sont justes seulement si ce qu'on note est juste.",
    ["Noter tout de suite : un plein, un trajet, une dépense", "Vérifier camion, chauffeur, montant et date avant d'enregistrer", "Joindre les documents (permis, assurance, contrat)", "Mettre à jour les dates de fin des papiers"],
    ["Tout noter en fin de semaine ou de mois", "Créer deux fois le même client ou le même type", "Laisser des essais dans le logiciel", "Corriger en silence l'erreur d'un collègue"],
    "chercher avant de créer ; les cases avec * sont obligatoires.",
    "Le logiciel donne des chiffres justes seulement si on note juste. Notez les choses tout de suite, pas en fin de mois. Vérifiez avant d'enregistrer. Avant de créer un client, cherchez s'il existe déjà, pour ne pas le créer deux fois. Exemple : dans notre logiciel, le type « Trip Income » existe deux fois, et les recettes sont coupées en deux.");

  // En cas de problème
  s = content("En cas de problème : les bons réflexes", "Les bons réflexes", "FaLifeRing");
  tx(s, "Prévenir vite vaut mieux que réparer tard. Prévenir n'est jamais une faute.", { x: 0.6, y: 1.5, w: 12.1, h: 0.5, fontSize: 17, color: C.text2, margin: 0 });
  [["FaKey", "Mon mot de passe a pu être vu", "Je le change tout de suite (Change Password) et je préviens mon responsable."],
   ["FaEnvelopeOpenText", "Un message me demande mon mot de passe", "Je ne réponds pas, je ne clique pas, je préviens."],
   ["FaPenToSquare", "Je vois une erreur", "Je la corrige si mon compte le permet, sinon je préviens mon responsable."],
   ["FaUserMinus", "Un collègue quitte l'entreprise", "Le responsable fait bloquer son compte le jour même."]]
    .forEach((d, i) => card(s, 0.6 + (i % 2) * 6.15, 2.25 + Math.floor(i / 2) * 2.25, 5.95, 2.0, d[0], d[1], d[2], i + 1));
  s.addNotes("Si vous avez un problème, prévenez vite. Ce n'est jamais une faute de prévenir. Si votre mot de passe a pu être vu, changez-le tout de suite. Si un message vous demande votre mot de passe, ne répondez pas. Si vous voyez une erreur, corrigez-la ou prévenez votre responsable. Dites bien à qui s'adresser dans l'entreprise.");

  // Quiz
  s = pres.addSlide({ masterName: "IMPACT", sectionTitle: "Les bons réflexes" });
  s.addText("Vrai ou faux ?", { placeholder: "title" });
  tx(s, "Votez à main levée.", Object.assign({ x: 0.8, y: 1.65, w: 11.5, h: 0.5, fontSize: 18, color: C.accent2 }, A(1)));
  ["Je peux prêter mon mot de passe à un collègue pressé.", "Je peux envoyer la liste des clients sur WhatsApp.",
   "Je note le plein dès qu'il est fait.", "Le compte d'un collègue qui part doit être bloqué."]
    .forEach((q, i) => { const x = 0.8 + (i % 2) * 5.95, y = 2.5 + Math.floor(i / 2) * 1.9, st = i + 2;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w: 5.7, h: 1.6, rectRadius: 0.15, fill: { color: "28356B" }, line: { color: "4A5C9A", width: 1 } }, A(st)));
      s.addShape(pres.shapes.OVAL, Object.assign({ x: x + 0.3, y: y + 0.45, w: 0.7, h: 0.7, fill: { color: C.background1 }, line: { color: C.background1 } }, A(st)));
      tx(s, String(i + 1), Object.assign({ x: x + 0.3, y: y + 0.45, w: 0.7, h: 0.7, fontSize: 20, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0 }, A(st)));
      tx(s, q, Object.assign({ x: x + 1.25, y, w: 4.25, h: 1.6, fontSize: 17, color: C.background1, valign: "middle", margin: 0 }, A(st))); });
  s.addNotes("Lisez chaque phrase, faites voter à main levée, puis donnez la réponse. RÉPONSES : 1) FAUX : tout ce qui est fait avec votre compte est noté à votre nom. 2) FAUX : ce sont des informations privées, elles ne doivent pas sortir. 3) VRAI : on note tout de suite pour avoir des chiffres justes. 4) VRAI : le jour même de son départ.");

  // Chiffres clés
  pres.addSection({ title: "Pour finir" });
  s = pres.addSlide({ masterName: "IMPACT", sectionTitle: "Pour finir" });
  s.addText("En résumé, en quelques chiffres", { placeholder: "title" });
  [["18", "rubriques", "pour tout le travail de l'entreprise"], ["11", "rapports", "prêts à l'emploi"], ["9", "façons", "de faire payer un trajet"], ["1", "seule application", "pour la direction et les agents"]]
    .forEach((d, i) => { const x = 0.8 + i * 3.0, st = i + 1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y: 2.1, w: 2.75, h: 3.6, rectRadius: 0.15, fill: { color: "28356B" }, line: { color: "4A5C9A", width: 1 } }, A(st)));
      tx(s, d[0], Object.assign({ x, y: 2.35, w: 2.75, h: 1.6, fontSize: 80, bold: true, color: i === 3 ? C.accent2 : C.background1, align: "center", valign: "middle", fontFace: "Arial" }, A(st)));
      tx(s, [{ text: d[1], options: { bold: true, fontSize: 19, color: C.background1, breakLine: true, paraSpaceAfter: 4 } }, { text: d[2], options: { fontSize: 14, color: C.accent2 } }], Object.assign({ x: x + 0.15, y: 4.05, w: 2.45, h: 1.4, align: "center" }, A(st))); });
  tx(s, "Et aussi : les alertes sur les dates de fin, les camions sur la carte, les listes en Excel ou PDF, les copies de sécurité.", Object.assign({ x: 0.8, y: 6.1, w: 11.7, h: 0.5, fontSize: 15, color: C.accent2 }, A(5)));
  s.addNotes("Pour résumer en chiffres : 18 rubriques, 11 rapports, 9 façons de faire payer un trajet, et une seule application pour tout le monde.");

  // Ce que ça change
  s = content("Qui fait quoi au quotidien ?", "Pour finir", "FaPeopleArrows");
  [["FaUserTie", "La direction", C.text2, ["Consulte le tableau de bord chaque jour", "Regarde les rapports : camions, routes, dépenses", "Décide qui peut faire quoi dans le logiciel", "Vérifie que tout est bien noté"]],
   ["FaUsers", "Les agents", C.accent1, ["Vérifient le calendrier avant de réserver", "Notent trajets, pleins et dépenses tout de suite", "Mettent à jour les papiers et les dates de fin", "Préviennent en cas d'erreur ou de problème"]]]
    .forEach((d, i) => { const x = 0.6 + i * 6.2, st = i * 5 + 1;
      panel(s, x, 1.6, 5.9, 5.1, C.background1, st, true);
      iconTile(s, d[0], x + 0.3, 1.85, 0.85, d[2], HEX.white, st);
      tx(s, d[1], Object.assign({ x: x + 1.35, y: 1.85, w: 4.3, h: 0.85, fontSize: 22, bold: true, color: C.text2, valign: "middle" }, A(st)));
      d[3].forEach((t, j) => { const y = 3.0 + j * 0.88;
        s.addImage(Object.assign({ data: ic("FaCircleCheck", HEX.blue), x: x + 0.35, y: y + 0.05, w: 0.32, h: 0.32 }, A(st + 1 + j)));
        tx(s, t, Object.assign({ x: x + 0.85, y, w: 4.85, h: 0.8, fontSize: 17, margin: 0 }, A(st + 1 + j))); }); });
  s.addNotes("Chacun a son rôle. La direction regarde le tableau de bord et les rapports, et décide qui peut faire quoi. Les agents notent tout de suite les trajets, les pleins et les dépenses, et préviennent en cas de problème. Question : quelle est la première chose que vous ferez demain dans le logiciel ?");

  // 16 — Démo : les étapes posées sur une route
  s = content("Démonstration : un circuit de travail", "Pour finir", "FaPlay");
  tx(s, "On le fait ensemble, en direct, du client jusqu'au rapport.", { x: 0.6, y: 1.55, w: 12, h: 0.45, fontSize: 17, color: C.text2 });
  s.addShape(pres.shapes.LINE, Object.assign({ x: 0.6, y: 3.05, w: 12.1, h: 0, line: { color: "D5DCEB", width: 26, cap: "round" } }, A(1)));
  s.addShape(pres.shapes.LINE, Object.assign({ x: 0.75, y: 3.05, w: 11.8, h: 0, line: { color: HEX.white, width: 2, dashType: "dash" } }, A(1)));
  s.addImage(Object.assign({ data: ic("FaTruckFast", HEX.navy), x: 0.55, y: 2.15, w: 0.6, h: 0.6 }, A(1)));
  const steps = [["FaUserPlus", "Client", "Customer › Add Customer"], ["FaRoute", "Trajet", "Trips › Add Trips"], ["FaGasPump", "Carburant", "Fuel › Add Fuel"], ["FaReceipt", "Argent reçu", "Accounts › Add Transactions"], ["FaChartLine", "Tableau de bord et rapport", "Dashboard · Reports"]];
  steps.forEach((d, i) => { const cx = 1.95 + i * 2.35, st = i + 2;
    s.addShape(pres.shapes.OVAL, Object.assign({ x: cx - 0.6, y: 2.45, w: 1.2, h: 1.2, fill: { color: C.text2 }, line: { color: C.background1, width: 3 }, shadow: shadow() }, A(st)));
    s.addImage(Object.assign({ data: ic(d[0], HEX.white), x: cx - 0.27, y: 2.78, w: 0.54, h: 0.54 }, A(st)));
    s.addShape(pres.shapes.OVAL, Object.assign({ x: cx + 0.28, y: 2.35, w: 0.42, h: 0.42, fill: { color: C.accent4 }, line: { color: C.background1, width: 2 } }, A(st)));
    tx(s, String(i + 1), Object.assign({ x: cx + 0.28, y: 2.35, w: 0.42, h: 0.42, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 }, A(st)));
    tx(s, [{ text: d[1], options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } }, { text: d[2], options: { fontSize: 12, color: C.accent5 } }], Object.assign({ x: cx - 1.1, y: 3.85, w: 2.2, h: 1.1, align: "center" }, A(st))); });
  panel(s, 0.6, 5.3, 12.1, 1.3, C.text2, 7, true);
  iconTile(s, "FaBullseye", 0.85, 5.55, 0.8, C.accent1, HEX.white, 7);
  tx(s, [{ text: "Objectif  ", options: { bold: true, color: C.accent2 } }, { text: "voir qu'on note une seule fois, et que tout se met à jour : le calendrier, l'argent, le tableau de bord et les rapports." }],
    Object.assign({ x: 1.9, y: 5.3, w: 10.6, h: 1.3, fontSize: 15, color: C.background1, valign: "middle" }, A(7)));
  s.addNotes("On fait un exemple ensemble, en direct. 1) On crée un client pour l'essai. 2) On crée un trajet pour ce client, avec un camion et un chauffeur. 3) On note un plein pour ce camion. 4) On note l'argent reçu pour le trajet. 5) On regarde le tableau de bord et un rapport : tout s'est mis à jour tout seul. À la fin, supprimer les données de l'essai.");

  // 17 — Questions
  s = pres.addSlide({ masterName: "SOMBRE", sectionTitle: "Pour finir" });
  hero(s, "FaComments");
  s.addText("Questions et échanges", { placeholder: "title" });
  tx(s, "Merci de votre attention", Object.assign({ x: 0.8, y: 4.05, w: 7.2, h: 0.6, fontSize: 20, color: C.accent2 }, A(2)));
  s.addNotes("Merci pour votre attention. Avez-vous des questions ? Dites-moi aussi ce qui vous semble difficile, pour qu'on puisse vous accompagner.");

  // 18 — Réserve : coûts (masquée, à afficher seulement sur demande de la direction)
  pres.addSection({ title: "Réserve" });
  s = content("Coûts de fonctionnement", "Réserve", "FaFileInvoiceDollar");
  s.hidden = true;
  [["FaServer", "Hébergement du serveur", "par an", 1.6], ["FaHeadset", "Maintenance", "par mois, formation des agents comprise", 4.2]].forEach((d, i) => {
    panel(s, 0.6, d[3], 5.6, 2.3, C.background1, i + 1, true);
    iconTile(s, d[0], 0.85, d[3] + 0.25, 0.6, C.background2, HEX.blue, i + 1);
    tx(s, d[1], Object.assign({ x: 1.6, y: d[3] + 0.25, w: 4.4, h: 0.6, fontSize: 16, bold: true, color: C.text2, valign: "middle" }, A(i + 1)));
    tx(s, [{ text: "100 000 FCFA", options: { bold: true, fontSize: 32, color: C.text2, breakLine: true } }, { text: d[2], options: { fontSize: 16, color: C.accent5 } }], Object.assign({ x: 0.85, y: d[3] + 1.0, w: 5.2, h: 1.2 }, A(i + 1))); });
  tx(s, "Calendrier de paiement de la maintenance", Object.assign({ x: 6.7, y: 1.6, w: 6, h: 0.45, bold: true, fontSize: 17, color: C.text2 }, A(3)));
  panel(s, 6.7, 2.2, 6.05, 1.9, C.text2, 3, true);
  tx(s, [{ text: "Premier paiement : 300 000 FCFA", options: { bold: true, fontSize: 20, color: C.background1, breakLine: true } }, { text: "1er mois de maintenance + 2 mois payés d'avance (mois 2 et 3)", options: { fontSize: 15, color: C.accent2 } }], Object.assign({ x: 6.95, y: 2.3, w: 5.6, h: 1.7, valign: "middle" }, A(3)));
  s.addShape(pres.shapes.DOWN_ARROW, Object.assign({ x: 9.5, y: 4.2, w: 0.45, h: 0.45, fill: { color: C.accent1 }, line: { color: C.accent1 } }, A(4)));
  panel(s, 6.7, 4.75, 6.05, 1.75, C.background1, 4, true);
  tx(s, [{ text: "À partir du 4ᵉ mois : 100 000 FCFA", options: { bold: true, fontSize: 20, color: C.text2, breakLine: true } }, { text: "réglés chaque mois", options: { fontSize: 15, color: C.accent5 } }], Object.assign({ x: 6.95, y: 4.85, w: 5.6, h: 1.55, valign: "middle" }, A(4)));
  s.addNotes("DIAPOSITIVE MASQUÉE : elle ne s'affiche pas pendant le diaporama. Pour l'afficher si la direction pose la question, taper 27 puis Entrée pendant le diaporama. Ne présenter que ces montants, sans autres frais.");

  await pres.writeFile({ fileName: OUT });
}

// --- Transitions et animations (post-traitement OOXML) : fondu entre diapositives,
// apparition automatique en fondu des éléments marqués « animN-… », étape par étape.
function timingXml(groups) {
  let id = 3, pars = "", blds = "";
  const steps = Object.keys(groups).map(Number).sort((a, b) => a - b);
  steps.forEach((step, k) => {
    groups[step].forEach(({ spid, text }) => {
      const a = id++, b = id++, c = id++;
      pars += `<p:par><p:cTn id="${a}" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect"><p:stCondLst><p:cond delay="${k * 250}"/></p:stCondLst><p:childTnLst>`
        + `<p:set><p:cBhvr><p:cTn id="${b}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`
        + `<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${c}" dur="500"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>`;
      if (text) blds += `<p:bldP spid="${spid}" grpId="0" animBg="1"/>`;
    });
  });
  const grp = id++;
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>`
    + `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>`
    + `<p:par><p:cTn id="${grp}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>`
    + `<p:par><p:cTn id="${id++}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${pars}</p:childTnLst></p:cTn></p:par>`
    + `</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>`
    + `<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>`
    + `<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>`
    + `</p:childTnLst></p:cTn></p:par></p:tnLst>${blds ? `<p:bldLst>${blds}</p:bldLst>` : ""}</p:timing>`;
}
async function animate() {
  const fs = require("fs");
  const zip = await JSZip.loadAsync(fs.readFileSync(OUT));
  for (const path of Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p))) {
    let xml = await zip.file(path).async("string");
    const groups = {};
    // chaque forme : <p:sp>, <p:pic> ou <p:cxnSp> dont le cNvPr porte un nom « animN-… »
    const re = /<p:(sp|pic|cxnSp)>([\s\S]*?)<\/p:\1>/g; let m;
    while ((m = re.exec(xml))) {
      const nv = /<p:cNvPr id="(\d+)" name="anim(\d+)-/.exec(m[2]);
      if (!nv) continue;
      (groups[nv[2]] = groups[nv[2]] || []).push({ spid: nv[1], text: m[1] === "sp" && /<a:t>/.test(m[2]) });
    }
    const trans = `<p:transition spd="med"><p:fade/></p:transition>`;
    const insert = trans + (Object.keys(groups).length ? timingXml(groups) : "");
    xml = xml.includes("<p:extLst>") && xml.lastIndexOf("<p:extLst>") > xml.indexOf("</p:clrMapOvr>")
      ? xml.replace(/<p:extLst>(?![\s\S]*<p:extLst>)/, insert + "<p:extLst>")
      : xml.replace("</p:sld>", insert + "</p:sld>");
    zip.file(path, xml);
  }
  fs.writeFileSync(OUT, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

(async () => {
  await loadIcons([
    ...["FaTruckFast", "FaLightbulb", "FaLayerGroup", "FaBell", "FaChartLine", "FaSitemap", "FaGaugeHigh", "FaTruck", "FaUsers", "FaRoute", "FaWrench", "FaCoins",
      "FaUserTie", "FaMapLocationDot", "FaChartPie", "FaUserShield", "FaGear", "FaStar", "FaPlay", "FaComments", "FaFileInvoiceDollar", "FaUserPlus", "FaGasPump",
      "FaReceipt", "FaArrowTrendUp", "FaArrowTrendDown", "FaBullseye", "FaPhone", "FaWrench", "FaCalendarDays", "FaShieldHalved", "FaUserSecret", "FaClipboardCheck", "FaLifeRing", "FaPeopleArrows", "FaLanguage"].map(n => [n, HEX.white]),
    ...["FaEnvelopeOpenText", "FaPenToSquare", "FaUserMinus"].map(n => [n, HEX.blue]), ["FaCircleCheck", HEX.green], ["FaCircleXmark", HEX.grey],
    ...["FaKey", "FaLock", "FaRoad", "FaUserLock", "FaAddressBook", "FaFingerprint", "FaFileExport", "FaAsterisk", "FaListUl", "FaCalculator", "FaFileCsv"].map(n => [n, HEX.blue]),
    ...["FaLayerGroup", "FaBell", "FaChartLine", "FaGaugeHigh", "FaTruck", "FaUsers", "FaRoute", "FaWrench", "FaCoins", "FaHashtag", "FaTruckFront", "FaCalendarDays",
      "FaIdCard", "FaCalendarXmark", "FaTag", "FaAddressCard", "FaIdBadge", "FaHandshake", "FaFileContract", "FaLocationCrosshairs", "FaClockRotateLeft", "FaDrawPolygon",
      "FaBullhorn", "FaClipboardList", "FaSignal", "FaTicket", "FaScrewdriverWrench", "FaUserGear", "FaBoxesStacked", "FaFillDrip", "FaArrowRightArrowLeft", "FaOilCan",
      "FaStore", "FaBuildingColumns", "FaReceipt", "FaMoneyBillTransfer", "FaTags", "FaScaleBalanced", "FaGasPump", "FaUserTie", "FaListCheck", "FaCubes", "FaToggleOn",
      "FaBuilding", "FaEnvelope", "FaFileInvoice", "FaDatabase", "FaServer", "FaHeadset", "FaUserPlus", "FaCircleCheck"].map(n => [n, HEX.blue]),
    ...["FaShuffle", "FaHourglassHalf", "FaCircleQuestion", "FaTruckFast"].map(n => [n, HEX.grey]),
    ["FaCircleCheck", HEX.sky], ...["FaTruckFast", "FaPhone", "FaGasPump", "FaWrench", "FaCoins", "FaChartLine", "FaShieldHalved", "FaUserSecret", "FaClipboardCheck"].map(n => [n, HEX.navy]),
  ]);
  await build();
  await applyTheme(OUT, THEME);
  await animate();
  console.log("ok");
})();
