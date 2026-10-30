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
  function module(title, sec, icon, menu, intro, feats, benefit, notes) {
    const s = content(title, sec, icon);
    tx(s, [{ text: "Menu  ", options: { bold: true, color: C.accent1 } }, { text: menu, options: { color: C.accent5 } }], { x: 0.6, y: 1.5, w: 8.4, h: 0.32, fontSize: 12, valign: "middle", margin: 0 });
    tx(s, intro, { x: 0.6, y: 1.88, w: 8.4, h: 0.5, fontSize: 17, color: C.text2, margin: 0 });
    feats.forEach((f, i) => card(s, 0.6 + (i % 2) * 4.3, 2.6 + Math.floor(i / 2) * 2.05, 4.1, 1.85, f[0], f[1], f[2], i + 1));
    const st = feats.length + 1;
    panel(s, 9.35, 1.5, 3.4, 5.2, C.text2, st, true);
    tx(s, "Ce que ça apporte", Object.assign({ x: 9.6, y: 1.7, w: 3.0, h: 0.5, fontSize: 18, bold: true, color: C.accent2, margin: 0 }, A(st)));
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
  tx(s, "PRÉSENTATION DU LOGICIEL", { x: 1.0, y: 1.45, w: 6, h: 0.32, fontSize: 12, bold: true, color: C.accent2, charSpacing: 3, margin: 0, valign: "middle" });
  s.addText("Division Terrestre Internationale", { placeholder: "title" });
  tx(s, "Gestion de la flotte, des chauffeurs, des trajets et des finances", Object.assign({ x: 0.8, y: 4.05, w: 7.2, h: 0.9, fontSize: 20, color: C.accent2 }, A(2)));
  tx(s, [{ text: "Présentation : [à compléter]", options: { breakLine: true } }, { text: "Date : [à compléter]" }], Object.assign({ x: 0.8, y: 5.6, w: 6, h: 0.9, fontSize: 14, color: C.background1 }, A(3)));
  s.addNotes("Accueillir les participants (agents et direction). Annoncer le déroulé : présentation du logiciel (ce qu'il sait faire, ses modules, comment et pourquoi l'utiliser), puis démonstration en direct d'un petit circuit de travail, puis questions.");

  // 2 — Pourquoi
  s = content("Pourquoi ce logiciel ?", "Introduction", "FaLightbulb");
  tx(s, "Aujourd'hui", Object.assign({ x: 0.6, y: 1.5, w: 5.6, h: 0.45, fontSize: 19, bold: true, color: C.accent5 }, A(1)));
  [["FaShuffle", "Informations dispersées", "Cahiers, tableurs, appels : difficile de savoir où est un camion et qui le conduit."],
   ["FaHourglassHalf", "Échéances oubliées", "Assurances, immatriculations et permis expirés découverts trop tard."],
   ["FaCircleQuestion", "Coûts mal suivis", "Carburant, entretien et pièces difficiles à rattacher à chaque véhicule."]]
    .forEach((d, i) => card(s, 0.6, 2.05 + i * 1.6, 5.6, 1.4, d[0], d[1], d[2], 2 + i, HEX.grey));
  s.addShape(pres.shapes.RIGHT_ARROW, Object.assign({ x: 6.42, y: 3.85, w: 0.55, h: 0.55, fill: { color: C.accent1 }, line: { color: C.accent1 } }, A(5)));
  tx(s, "Avec Division Terrestre Internationale", Object.assign({ x: 7.2, y: 1.5, w: 5.6, h: 0.45, fontSize: 19, bold: true, color: C.accent1 }, A(6)));
  [["FaLayerGroup", "Tout au même endroit", "Véhicules, chauffeurs, trajets, clients, carburant, entretien et finances."],
   ["FaBell", "Alertes et rappels", "Documents expirés signalés, rappels d'entretien, tableau de bord du jour."],
   ["FaChartLine", "Décisions chiffrées", "Revenus, dépenses et rapports par véhicule, chauffeur ou itinéraire."]]
    .forEach((d, i) => card(s, 7.2, 2.05 + i * 1.6, 5.55, 1.4, d[0], d[1], d[2], 7 + i));
  s.addNotes("Partir du quotidien des participants : demander comment ils savent aujourd'hui quel camion est disponible, quand expire une assurance, ou combien a coûté un véhicule le mois dernier. Le logiciel répond à ces trois besoins : centraliser, alerter, chiffrer.");

  // 3 — Modules
  s = content("18 modules, 6 grandes familles", "Introduction", "FaSitemap");
  s.addShape(pres.shapes.OVAL, Object.assign({ x: 0.45, y: 2.25, w: 3.0, h: 3.0, fill: { color: C.background2 }, line: { color: C.accent6 } }, A(1)));
  s.addShape(pres.shapes.OVAL, Object.assign({ x: 0.75, y: 2.55, w: 2.4, h: 2.4, fill: { color: C.text2 }, line: { color: C.text2 } }, A(1)));
  tx(s, [{ text: "18", options: { fontSize: 60, bold: true, breakLine: true } }, { text: "modules", options: { fontSize: 17, color: C.accent2 } }], Object.assign({ x: 0.75, y: 2.55, w: 2.4, h: 2.4, align: "center", valign: "middle", color: C.background1 }, A(1)));
  [["FaGaugeHigh", "Pilotage", "Dashboard · Availability · Reports"], ["FaTruck", "Flotte", "Vehicle's · Vehicle Vendor's · Tracking · Geofence"], ["FaUsers", "Personnes", "Driver's · Employee"],
   ["FaRoute", "Activité commerciale", "Trips · Customer · Coupon Management"], ["FaWrench", "Exploitation", "Maintenance · Parts Stock · Fuel · Reminder"], ["FaCoins", "Finances et système", "Accounts · Setting's"]]
    .forEach((d, i) => card(s, 3.8 + (i % 2) * 4.5, 1.6 + Math.floor(i / 2) * 1.75, 4.3, 1.55, d[0], d[1], d[2], 2 + i));
  s.addNotes("Présenter la carte du logiciel : 18 modules dans le menu latéral, regroupés ici en 6 familles pour faciliter la lecture. Les libellés restent en anglais dans l'application : on les garde tels quels pour que chacun s'y retrouve à l'écran. Chaque famille est détaillée dans les diapositives suivantes.");


  // 4 — Le parcours d'un trajet (récit)
  s = pres.addSlide({ masterName: "IMPACT", sectionTitle: "Introduction" });
  s.addText("Un trajet, du premier appel au dernier franc", { placeholder: "title" });
  tx(s, "Chaque étape du parcours d'un camion a sa place dans le logiciel.", Object.assign({ x: 0.8, y: 1.65, w: 11, h: 0.5, fontSize: 18, color: C.accent2 }, A(1)));
  s.addShape(pres.shapes.LINE, Object.assign({ x: 0.9, y: 3.85, w: 11.6, h: 0, line: { color: "3A4A82", width: 30, cap: "round" } }, A(1)));
  s.addShape(pres.shapes.LINE, Object.assign({ x: 1.0, y: 3.85, w: 11.4, h: 0, line: { color: HEX.sky, width: 2, dashType: "dash" } }, A(1)));
  [["FaPhone", "Réserver", "Client, véhicule disponible, chauffeur"], ["FaTruckFast", "Rouler", "Suivi GPS, zones, statut du trajet"], ["FaGasPump", "Alimenter", "Pleins rattachés au véhicule"],
   ["FaWrench", "Entretenir", "Rappels, mécaniciens, pièces"], ["FaCoins", "Encaisser", "Transactions et comptes"], ["FaChartLine", "Analyser", "Tableau de bord et rapports"]]
    .forEach((d, i) => { const cx = 1.45 + i * 2.08, st = i + 2;
      s.addShape(pres.shapes.OVAL, Object.assign({ x: cx - 0.62, y: 3.23, w: 1.24, h: 1.24, fill: { color: i === 1 ? C.accent1 : C.background1 }, line: { color: "3A4A82", width: 4 } }, A(st)));
      s.addImage(Object.assign({ data: ic(d[0], i === 1 ? HEX.white : HEX.navy), x: cx - 0.29, y: 3.56, w: 0.58, h: 0.58 }, A(st)));
      tx(s, [{ text: d[1], options: { bold: true, fontSize: 18, color: C.background1, breakLine: true, paraSpaceAfter: 4 } }, { text: d[2], options: { fontSize: 13, color: C.accent2 } }], Object.assign({ x: cx - 1.0, y: 4.75, w: 2.0, h: 1.3, align: "center" }, A(st))); });
  s.addNotes("Raconter le parcours d'un trajet comme une histoire : un client appelle, on réserve un camion disponible avec un chauffeur, le camion roule et on le suit, il fait le plein, il passe à l'entretien, on encaisse, puis on analyse la rentabilité. Annoncer que chaque étape correspond à un module présenté ensuite.");

  // Modules
  pres.addSection({ title: "Pilotage" });
  module("Le tableau de bord : la journée en un coup d'œil", "Pilotage", "FaGaugeHigh", "Dashboard · Availability",
    "Dès la connexion, l'essentiel de l'activité est affiché.",
    [["FaHashtag", "Compteurs et transactions", "Véhicules, chauffeurs, clients, trajets du jour ; revenus et dépenses."], ["FaTruckFront", "Statut de la flotte", "Véhicules disponibles ou en trajet, et leur position actuelle."],
     ["FaBell", "Rappels et tops", "Entretiens à prévoir, principaux revenus et dépenses."], ["FaCalendarDays", "Calendrier de disponibilité", "Occupation des véhicules par mois, semaine ou jour."]],
    ["La direction voit l'activité sans demander de rapport", "Les agents savent tout de suite quels véhicules sont libres", "Les urgences sont visibles dès l'ouverture"],
    "Montrer le tableau de bord en direct pendant la démonstration. Insister sur le calendrier Availability : c'est le premier réflexe avant de planifier un trajet.");

  pres.addSection({ title: "Flotte" });
  module("Gérer les véhicules", "Flotte", "FaTruck", "Vehicle's › Vehicle List · Add Vehicle · Vehicle Group · Route",
    "Une fiche complète pour chaque véhicule, de l'immatriculation au GPS.",
    [["FaIdCard", "Fiche véhicule", "Immatriculation, modèle, châssis, moteur, compteur, photo et documents."], ["FaCalendarXmark", "Échéances", "Expiration de l'immatriculation et de l'assurance, signalée à l'écran."],
     ["FaLayerGroup", "Types et itinéraires", "Camion, bus, voiture… et liste des itinéraires."], ["FaTag", "Tarif et propriété", "Facturation et coût par défaut, véhicule en propre ou loué, liaison GPS."]],
    ["Plus d'assurance expirée qui passe inaperçue", "Toute la flotte saisie d'un coup par import CSV", "Export Excel, CSV ou PDF"],
    "Points à montrer : la colonne des dates d'expiration en rouge, l'import en masse (Bulk Import), et les boutons d'export présents sur toutes les listes du logiciel.");

  module("Chauffeurs et fournisseurs de véhicules", "Flotte", "FaUserTie", "Driver's · Vehicle Vendor's",
    "Savoir qui conduit, avec quel permis, et avec quels partenaires on travaille.",
    [["FaAddressCard", "Fiche chauffeur", "Contact, âge, adresse, expérience, date d'entrée, photo et document."], ["FaIdBadge", "Permis", "Numéro et date d'expiration, signalée quand elle est dépassée."],
     ["FaHandshake", "Fournisseurs de véhicules", "Société, contact, téléphone, adresse et statut."], ["FaFileContract", "Contrats", "Date de contrat et document du contrat joint."]],
    ["Un permis expiré est repéré tout de suite", "Les contrats des loueurs sont centralisés", "Import des chauffeurs en masse par CSV"],
    "Faire le lien avec la sécurité et la conformité : un permis expiré est affiché en rouge dans la liste. Les fournisseurs de véhicules (Vehicle Vendor's) concernent les véhicules loués à des partenaires.");

  module("Suivi GPS et zones géographiques", "Flotte", "FaMapLocationDot", "Tracking › History Tracking · Live Location · Geofence",
    "Savoir où sont les véhicules, et être prévenu quand ils entrent ou sortent d'une zone.",
    [["FaLocationCrosshairs", "Position en direct", "Localisation actuelle des véhicules sur la carte."], ["FaClockRotateLeft", "Historique des parcours", "Trajet d'un véhicule sur une période choisie."],
     ["FaDrawPolygon", "Zones (geofences)", "Zones dessinées sur la carte : cercle, rectangle, polygone…"], ["FaBullhorn", "Événements de zone", "Entrées et sorties enregistrées, avec notification."]],
    ["Contrôler les itinéraires réellement suivis", "Réagir vite en cas d'écart", "Rassurer les clients sur leurs livraisons"],
    "Préciser que le suivi GPS nécessite un boîtier GPS relié à chaque véhicule (synchronisation Traccar, visible sur le tableau de bord). Montrer la carte de l'historique si un véhicule est équipé.");

  pres.addSection({ title: "Activité commerciale" });
  s = module("Trajets, clients et coupons", "Activité commerciale", "FaRoute", "Trips · Customer · Coupon Management",
    "De la réservation à la facturation, chaque trajet est suivi.",
    [["FaClipboardList", "Réservation", "Client, dates, véhicule, chauffeur, départ, arrivée, kilométrage, ramassages."], ["FaSignal", "Suivi du trajet", "Statut (réservé, commencé…), distance et numéro de réservation."],
     ["FaUsers", "Clients", "Coordonnées, solde impayé et statut actif ou inactif."], ["FaTicket", "Coupons", "Codes de réduction : montant, dates, limite d'utilisation."]],
    ["Chaque trajet lié à un client, un véhicule, un chauffeur", "Les impayés clients sont visibles", "Confirmation par e-mail ou SMS"],
    "Insister sur la souplesse de facturation (encart en bas). La réservation peut être envoyée au client par e-mail ou SMS. Ce module servira de point de départ à la démonstration.");
  tx(s, [{ text: "9 modes de facturation : ", options: { bold: true, color: C.accent1 } }, { text: "fixe · tonne · kg · km · trajet · jour · heure · litre · sac", options: { color: C.text2 } }], Object.assign({ x: 0.6, y: 6.62, w: 8.6, h: 0.3, fontSize: 13, margin: 0 }, A(9)));

  pres.addSection({ title: "Exploitation" });
  module("Entretien, pièces et rappels", "Exploitation", "FaWrench", "Maintenance · Parts Stock · Reminder",
    "Anticiper les entretiens et maîtriser leur coût.",
    [["FaScrewdriverWrench", "Entretiens", "Véhicule, dates, mécanicien, fournisseur, coût et pièces utilisées."], ["FaUserGear", "Mécaniciens et fournisseurs", "Annuaire des intervenants, avec catégorie."],
     ["FaBoxesStacked", "Stock de pièces", "Inventaire, prix et historique des achats avec statut de paiement."], ["FaBell", "Rappels", "Échéances par véhicule et service, marquées comme terminées."]],
    ["Moins de pannes grâce aux rappels", "Coût réel de l'entretien par véhicule", "Stock de pièces et achats tracés"],
    "Montrer qu'un entretien consomme des pièces du stock et peut notifier des membres de l'équipe. Les rappels apparaissent aussi sur le tableau de bord.");

  module("Carburant", "Exploitation", "FaGasPump", "Fuel › Fuel Management · Add Fuel · Fuel Vendor",
    "Chaque plein est enregistré et rattaché à un véhicule et un chauffeur.",
    [["FaFillDrip", "Pleins", "Date, véhicule, chauffeur, quantité, montant, compteur et commentaire."], ["FaArrowRightArrowLeft", "Provenance", "Cuve interne ou fournisseur extérieur."],
     ["FaOilCan", "Cuve", "Réapprovisionnement et stock disponible en litres."], ["FaStore", "Fournisseurs", "Liste des stations ou fournisseurs de carburant."]],
    ["Repérer une consommation anormale", "Contrôler le stock de la cuve", "Rapport carburant par période"],
    "Le compteur kilométrique saisi à chaque plein permet de comparer consommation et distance. Le rapport Fuel du module Reports reprend ces données.");

  pres.addSection({ title: "Finances et pilotage" });
  module("Finances", "Finances et pilotage", "FaCoins", "Accounts › Accounts · Add Transactions · Transactions · Ac Transfers / Deposit · Category",
    "Revenus et dépenses enregistrés par compte, catégorie et véhicule.",
    [["FaBuildingColumns", "Comptes", "Comptes financiers avec leur solde."], ["FaReceipt", "Transactions", "Revenu ou dépense : compte, catégorie, véhicule, montant, référence."],
     ["FaMoneyBillTransfer", "Virements et dépôts", "Mouvements d'un compte à un autre."], ["FaTags", "Catégories", "Revenus de trajet, carburant, salaires, transfrontalier…"]],
    ["Rentabilité visible par véhicule", "Historique filtrable avec totaux", "Catégories adaptées à l'entreprise"],
    "Montrer la liste des transactions filtrée par période, avec le total des débits et des crédits. Les catégories sont personnalisables : à définir avec la direction financière.");

  // 12 — Rapports
  s = content("11 rapports prêts à l'emploi", "Finances et pilotage", "FaChartPie");
  tx(s, "Choisir une période, un filtre, puis « Generate Report ».", { x: 0.6, y: 1.55, w: 12, h: 0.45, fontSize: 17, color: C.text2 });
  [["FaClipboardList", "Réservations"], ["FaScaleBalanced", "Revenus et dépenses"], ["FaGasPump", "Carburant"], ["FaUserTie", "Chauffeurs"], ["FaTicket", "Coupons"], ["FaBell", "Rappels"],
   ["FaWrench", "Entretien"], ["FaArrowTrendUp", "Top revenus"], ["FaArrowTrendDown", "Top dépenses"], ["FaTruck", "Top véhicules"], ["FaRoute", "Top itinéraires"]]
    .forEach((r, i) => { const x = 0.6 + (i % 4) * 3.1, y = 2.2 + Math.floor(i / 4) * 1.5, top = i >= 7;
      panel(s, x, y, 2.9, 1.25, top ? C.text2 : C.background1, i + 1, true);
      iconTile(s, r[0], x + 0.22, y + 0.3, 0.65, top ? C.accent1 : C.background2, top ? HEX.white : HEX.blue, i + 1);
      tx(s, r[1], Object.assign({ x: x + 1.02, y, w: 1.8, h: 1.25, fontSize: 15, bold: true, valign: "middle", color: top ? C.background1 : C.text2 }, A(i + 1))); });
  tx(s, "Les « Top » classent les postes qui rapportent ou coûtent le plus.", { x: 9.9, y: 5.2, w: 2.85, h: 1.25, fontSize: 13, italic: true, color: C.accent5, valign: "middle" });
  s.addNotes("Pour la direction : les rapports Top véhicules et Top itinéraires montrent où l'entreprise gagne et perd de l'argent. Pour les agents : les rapports remplacent les synthèses faites à la main.");

  pres.addSection({ title: "Administration" });
  module("Employés et droits d'accès", "Administration", "FaUserShield", "Employee › Employee Management · Add Employee",
    "Chaque utilisateur a son compte et ne voit que ce qu'il doit voir.",
    [["FaUserPlus", "Comptes utilisateurs", "Nom, téléphone, e-mail, identifiant, mot de passe, dernière connexion."], ["FaListCheck", "Permissions par module", "Cases à cocher : voir, ajouter, modifier, supprimer…"],
     ["FaCubes", "Modules concernés", "Véhicules, chauffeurs, transactions, clients, carburant, entretien, GPS…"], ["FaToggleOn", "Statut", "Compte actif ou désactivé."]],
    ["Chacun travaille dans son périmètre", "Les données financières restent protégées", "On sait qui a créé chaque transaction"],
    "Les droits se règlent finement, module par module et action par action. Les profils (exploitation, finance, direction…) seront définis avec l'organisation avant la mise en service.");

  module("Paramètres et sécurité", "Administration", "FaGear", "Setting's",
    "Le logiciel s'adapte à l'entreprise et protège ses données.",
    [["FaBuilding", "Identité de l'entreprise", "Nom, adresse, contacts, logo, couleurs, fuseau horaire."], ["FaEnvelope", "Communication", "Envoi d'e-mails et de SMS, avec modèles de messages."],
     ["FaFileInvoice", "Réservations et factures", "Préfixe des numéros de réservation, réglages des factures."], ["FaDatabase", "Sauvegardes", "Sauvegarde de la base de données en un clic."]],
    ["Documents aux couleurs de l'entreprise", "Clients informés automatiquement", "Données protégées par des sauvegardes"],
    "Les paramètres sont réservés à l'administrateur. Mentionner aussi la configuration de la carte (MAP API) et du GPS (Traccar Config).");


  // Chiffres clés
  pres.addSection({ title: "Conclusion" });
  s = pres.addSlide({ masterName: "IMPACT", sectionTitle: "Conclusion" });
  s.addText("En résumé, en quelques chiffres", { placeholder: "title" });
  [["18", "modules", "pour couvrir toute l'activité"], ["11", "rapports", "prêts à l'emploi"], ["9", "modes", "de facturation des trajets"], ["1", "seule application", "pour la direction et les agents"]]
    .forEach((d, i) => { const x = 0.8 + i * 3.0, st = i + 1;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y: 2.1, w: 2.75, h: 3.6, rectRadius: 0.15, fill: { color: "28356B" }, line: { color: "4A5C9A", width: 1 } }, A(st)));
      tx(s, d[0], Object.assign({ x, y: 2.35, w: 2.75, h: 1.6, fontSize: 80, bold: true, color: i === 3 ? C.accent2 : C.background1, align: "center", valign: "middle", fontFace: "Arial" }, A(st)));
      tx(s, [{ text: d[1], options: { bold: true, fontSize: 19, color: C.background1, breakLine: true, paraSpaceAfter: 4 } }, { text: d[2], options: { fontSize: 14, color: C.accent2 } }], Object.assign({ x: x + 0.15, y: 4.05, w: 2.45, h: 1.4, align: "center" }, A(st))); });
  tx(s, "Plus des alertes sur les échéances, le suivi GPS, les exports Excel, CSV, PDF et les sauvegardes.", Object.assign({ x: 0.8, y: 6.1, w: 11.7, h: 0.5, fontSize: 15, color: C.accent2 }, A(5)));
  s.addNotes("Diapositive de respiration avant la synthèse : laisser les chiffres apparaître un à un et les commenter brièvement.");

  // Ce que ça change
  s = content("Ce que ça change pour chacun", "Conclusion", "FaStar");
  [["FaUserTie", "Pour la direction", C.text2, ["Vision en temps réel de l'activité", "Rentabilité par véhicule, chauffeur et itinéraire", "Coûts de carburant et d'entretien maîtrisés", "Décisions appuyées sur des rapports"]],
   ["FaUsers", "Pour les agents", C.accent1, ["Une seule application au lieu de plusieurs outils", "Informations retrouvées en quelques secondes", "Rappels automatiques des échéances", "Moins de saisie grâce aux imports et exports"]]]
    .forEach((d, i) => { const x = 0.6 + i * 6.2, st = i * 5 + 1;
      panel(s, x, 1.6, 5.9, 5.1, C.background1, st, true);
      iconTile(s, d[0], x + 0.3, 1.85, 0.85, d[2], HEX.white, st);
      tx(s, d[1], Object.assign({ x: x + 1.35, y: 1.85, w: 4.3, h: 0.85, fontSize: 22, bold: true, color: C.text2, valign: "middle" }, A(st)));
      d[3].forEach((t, j) => { const y = 3.0 + j * 0.88;
        s.addImage(Object.assign({ data: ic("FaCircleCheck", HEX.blue), x: x + 0.35, y: y + 0.05, w: 0.32, h: 0.32 }, A(st + 1 + j)));
        tx(s, t, Object.assign({ x: x + 0.85, y, w: 4.85, h: 0.8, fontSize: 17, margin: 0 }, A(st + 1 + j))); }); });
  s.addNotes("Synthèse orientée bénéfices, avant la démonstration. Laisser la direction et les agents réagir : quel bénéfice compte le plus pour eux ?");

  // 16 — Démo : les étapes posées sur une route
  s = content("Démonstration : un circuit de travail", "Conclusion", "FaPlay");
  tx(s, "En direct dans l'application, de la commande au rapport.", { x: 0.6, y: 1.55, w: 12, h: 0.45, fontSize: 17, color: C.text2 });
  s.addShape(pres.shapes.LINE, Object.assign({ x: 0.6, y: 3.05, w: 12.1, h: 0, line: { color: "D5DCEB", width: 26, cap: "round" } }, A(1)));
  s.addShape(pres.shapes.LINE, Object.assign({ x: 0.75, y: 3.05, w: 11.8, h: 0, line: { color: HEX.white, width: 2, dashType: "dash" } }, A(1)));
  s.addImage(Object.assign({ data: ic("FaTruckFast", HEX.navy), x: 0.55, y: 2.15, w: 0.6, h: 0.6 }, A(1)));
  const steps = [["FaUserPlus", "Client", "Customer › Add Customer"], ["FaRoute", "Trajet", "Trips › Add Trips"], ["FaGasPump", "Carburant", "Fuel › Add Fuel"], ["FaReceipt", "Transaction", "Accounts › Add Transactions"], ["FaChartLine", "Tableau de bord et rapport", "Dashboard · Reports"]];
  steps.forEach((d, i) => { const cx = 1.95 + i * 2.35, st = i + 2;
    s.addShape(pres.shapes.OVAL, Object.assign({ x: cx - 0.6, y: 2.45, w: 1.2, h: 1.2, fill: { color: C.text2 }, line: { color: C.background1, width: 3 }, shadow: shadow() }, A(st)));
    s.addImage(Object.assign({ data: ic(d[0], HEX.white), x: cx - 0.27, y: 2.78, w: 0.54, h: 0.54 }, A(st)));
    s.addShape(pres.shapes.OVAL, Object.assign({ x: cx + 0.28, y: 2.35, w: 0.42, h: 0.42, fill: { color: C.accent4 }, line: { color: C.background1, width: 2 } }, A(st)));
    tx(s, String(i + 1), Object.assign({ x: cx + 0.28, y: 2.35, w: 0.42, h: 0.42, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 }, A(st)));
    tx(s, [{ text: d[1], options: { bold: true, fontSize: 16, color: C.text2, breakLine: true } }, { text: d[2], options: { fontSize: 12, color: C.accent5 } }], Object.assign({ x: cx - 1.1, y: 3.85, w: 2.2, h: 1.1, align: "center" }, A(st))); });
  panel(s, 0.6, 5.3, 12.1, 1.3, C.text2, 7, true);
  iconTile(s, "FaBullseye", 0.85, 5.55, 0.8, C.accent1, HEX.white, 7);
  tx(s, [{ text: "Objectif  ", options: { bold: true, color: C.accent2 } }, { text: "montrer qu'une seule saisie alimente tout le reste : le trajet apparaît dans le calendrier, le carburant et la transaction dans les finances, puis dans le tableau de bord et les rapports." }],
    Object.assign({ x: 1.9, y: 5.3, w: 10.6, h: 1.3, fontSize: 15, color: C.background1, valign: "middle" }, A(7)));
  s.addNotes("DÉROULÉ : 1) créer un client fictif ; 2) créer un trajet pour ce client avec un véhicule et un chauffeur, choisir un mode de facturation ; 3) enregistrer un plein pour ce véhicule ; 4) enregistrer la transaction de revenu du trajet (catégorie Trip Income) ; 5) revenir au tableau de bord et générer un rapport sur la période. Utiliser uniquement des données de test, et les supprimer après la démonstration.");

  // 17 — Questions
  s = pres.addSlide({ masterName: "SOMBRE", sectionTitle: "Conclusion" });
  hero(s, "FaComments");
  s.addText("Questions et échanges", { placeholder: "title" });
  tx(s, "Merci de votre attention", Object.assign({ x: 0.8, y: 4.05, w: 7.2, h: 0.6, fontSize: 20, color: C.accent2 }, A(2)));
  s.addNotes("Recueillir les questions et les besoins : modules prioritaires, profils de droits à créer, données à importer (véhicules, chauffeurs), formation des agents.");

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
  s.addNotes("DIAPOSITIVE MASQUÉE : elle ne s'affiche pas pendant le diaporama. Pour l'afficher si la direction pose la question, taper 20 puis Entrée pendant le diaporama. Ne présenter que ces montants, sans autres frais.");

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
      "FaReceipt", "FaArrowTrendUp", "FaArrowTrendDown", "FaBullseye", "FaPhone", "FaWrench"].map(n => [n, HEX.white]),
    ...["FaLayerGroup", "FaBell", "FaChartLine", "FaGaugeHigh", "FaTruck", "FaUsers", "FaRoute", "FaWrench", "FaCoins", "FaHashtag", "FaTruckFront", "FaCalendarDays",
      "FaIdCard", "FaCalendarXmark", "FaTag", "FaAddressCard", "FaIdBadge", "FaHandshake", "FaFileContract", "FaLocationCrosshairs", "FaClockRotateLeft", "FaDrawPolygon",
      "FaBullhorn", "FaClipboardList", "FaSignal", "FaTicket", "FaScrewdriverWrench", "FaUserGear", "FaBoxesStacked", "FaFillDrip", "FaArrowRightArrowLeft", "FaOilCan",
      "FaStore", "FaBuildingColumns", "FaReceipt", "FaMoneyBillTransfer", "FaTags", "FaScaleBalanced", "FaGasPump", "FaUserTie", "FaListCheck", "FaCubes", "FaToggleOn",
      "FaBuilding", "FaEnvelope", "FaFileInvoice", "FaDatabase", "FaServer", "FaHeadset", "FaUserPlus", "FaCircleCheck"].map(n => [n, HEX.blue]),
    ...["FaShuffle", "FaHourglassHalf", "FaCircleQuestion", "FaTruckFast"].map(n => [n, HEX.grey]),
    ["FaCircleCheck", HEX.sky], ...["FaTruckFast", "FaPhone", "FaGasPump", "FaWrench", "FaCoins", "FaChartLine"].map(n => [n, HEX.navy]),
  ]);
  await build();
  await applyTheme(OUT, THEME);
  await animate();
  console.log("ok");
})();
