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

const HEX = { red: "C60007", navy: "1E2340", ink: "1B2430", teal: "2E6F8E", green: "4E8A5B", amber: "E8A317", grey: "6B7584", light: "EEF1F5", line: "C9D3DD", white: "FFFFFF" };
const THEME = { name: "DTI", headFontFace: "Arial", bodyFontFace: "Calibri",
  colors: { dk1: HEX.ink, lt1: HEX.white, dk2: HEX.navy, lt2: HEX.light, accent1: HEX.red, accent2: HEX.teal,
    accent3: HEX.green, accent4: HEX.amber, accent5: HEX.grey, accent6: HEX.line, hlink: HEX.teal, folHlink: HEX.grey } };

// --- Icônes : SVG react-icons rastérisé en PNG
const ICONS = {};
async function loadIcons(list) {
  for (const [name, hex] of list) {
    const key = name + hex;
    if (ICONS[key]) continue;
    const svg = renderToStaticMarkup(React.createElement(fa[name], { color: "#" + hex, size: 256 }));
    const png = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
    ICONS[key] = "image/png;base64," + png.toString("base64");
  }
}
const ic = (name, hex) => ICONS[name + hex];

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  pres.title = "Division Terrestre Internationale";
  const C = pres.SchemeColor;

  pres.defineSlideMaster({ title: "SOMBRE", background: { color: HEX.navy },
    objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.0, w: 7.6, h: 1.9, fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left" }, text: "" } }] });
  pres.defineSlideMaster({ title: "CONTENU", background: { color: HEX.white },
    slideNumber: { x: 12.25, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
    objects: [
      { text: { text: "Division Terrestre Internationale", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent5 } } },
      { placeholder: { options: { name: "title", type: "title", x: 1.55, y: 0.4, w: 11.2, h: 0.8, fontSize: 28, bold: true, color: C.text2, valign: "middle", align: "left" }, text: "" } },
    ] });

  let n = 0;
  const A = (step) => ({ objectName: `anim${step}-${++n}` }); // repère d'animation (étape d'apparition)
  function content(title, sec, icon) {
    const s = pres.addSlide({ masterName: "CONTENU", sectionTitle: sec });
    s.addText(title, { placeholder: "title" });
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: 0.45, w: 0.7, h: 0.7, fill: { color: C.accent1 }, line: { color: C.accent1 } });
    s.addImage({ data: ic(icon, HEX.white), x: 0.77, y: 0.62, w: 0.36, h: 0.36 });
    return s;
  }
  function tx(s, t, o) { s.addText(t, Object.assign({ isTextBox: true, fontSize: 15, color: C.text1, valign: "top" }, o)); }
  function rbox(s, x, y, w, h, fill, step) {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, Object.assign({ x, y, w, h, rectRadius: 0.1, fill: { color: fill }, line: { color: fill } }, step ? A(step) : {}));
  }
  function iconDot(s, icon, x, y, d, fill, iconHex, step) {
    s.addShape(pres.shapes.OVAL, Object.assign({ x, y, w: d, h: d, fill: { color: fill }, line: { color: fill } }, step ? A(step) : {}));
    s.addImage(Object.assign({ data: ic(icon, iconHex), x: x + d * 0.25, y: y + d * 0.25, w: d * 0.5, h: d * 0.5 }, step ? A(step) : {}));
  }
  // Carte avec icône
  function card(s, x, y, w, h, icon, head, body, step) {
    rbox(s, x, y, w, h, C.background2, step);
    iconDot(s, icon, x + 0.2, y + 0.2, 0.55, C.accent1, HEX.white, step);
    tx(s, [{ text: head, options: { bold: true, fontSize: 16, color: C.text2, breakLine: true, paraSpaceAfter: 4 } }, { text: body, options: { fontSize: 14 } }],
      Object.assign({ x: x + 0.9, y: y + 0.12, w: w - 1.05, h: h - 0.24 }, A(step)));
  }
  // Diapositive « module »
  function module(title, sec, icon, menu, intro, feats, benefit, notes) {
    const s = content(title, sec, icon);
    rbox(s, 0.6, 1.4, 8.4, 0.42, C.background2);
    tx(s, [{ text: "Menu  ", options: { bold: true, color: C.accent1 } }, { text: menu, options: { color: C.text2 } }], { x: 0.75, y: 1.4, w: 8.2, h: 0.42, fontSize: 12, valign: "middle", margin: 0 });
    tx(s, intro, { x: 0.6, y: 1.95, w: 8.4, h: 0.5, fontSize: 17, italic: true, color: C.accent2, margin: 0 });
    feats.forEach((f, i) => card(s, 0.6 + (i % 2) * 4.3, 2.65 + Math.floor(i / 2) * 2.05, 4.1, 1.85, f[0], f[1], f[2], i + 1));
    const st = feats.length + 1;
    rbox(s, 9.35, 1.4, 3.4, 5.3, C.text2, st);
    tx(s, "Ce que ça apporte", Object.assign({ x: 9.6, y: 1.6, w: 3.0, h: 0.5, fontSize: 18, bold: true, color: C.accent4, margin: 0 }, A(st)));
    benefit.forEach((b, i) => { const y = 2.35 + i * 1.3;
      s.addImage(Object.assign({ data: ic("FaCircleCheck", HEX.amber), x: 9.6, y: y + 0.05, w: 0.32, h: 0.32 }, A(st + 1 + i)));
      tx(s, b, Object.assign({ x: 10.05, y, w: 2.55, h: 1.15, fontSize: 14, color: C.background1, margin: 0 }, A(st + 1 + i))); });
    s.addNotes(notes);
    return s;
  }

  // 1 — Titre
  pres.addSection({ title: "Introduction" });
  let s = pres.addSlide({ masterName: "SOMBRE", sectionTitle: "Introduction" });
  s.addShape(pres.shapes.OVAL, { x: 8.6, y: -0.9, w: 6.6, h: 6.6, fill: { color: C.accent1 }, line: { color: C.accent1 } });
  s.addShape(pres.shapes.OVAL, { x: 10.4, y: 4.9, w: 3.6, h: 3.6, fill: { color: C.accent2, transparency: 30 }, line: { color: C.accent2, transparency: 30 } });
  s.addImage(Object.assign({ data: ic("FaTruckFast", HEX.white), x: 9.9, y: 1.0, w: 3.6, h: 3.6 }, A(1)));
  s.addText("Division Terrestre Internationale", { placeholder: "title" });
  tx(s, "Le logiciel de gestion de la flotte, des chauffeurs, des trajets et des finances", Object.assign({ x: 0.8, y: 4.05, w: 7.4, h: 0.9, fontSize: 20, color: C.accent4 }, A(2)));
  tx(s, [{ text: "Présentation : [à compléter]", options: { breakLine: true } }, { text: "Date : [à compléter]" }], Object.assign({ x: 0.8, y: 5.6, w: 6, h: 0.9, fontSize: 14, color: C.background1 }, A(3)));
  s.addNotes("Accueillir les participants (agents et direction). Annoncer le déroulé : présentation du logiciel (ce qu'il sait faire, ses modules, comment et pourquoi l'utiliser), puis démonstration en direct d'un petit circuit de travail, puis questions.");

  // 2 — Pourquoi
  s = content("Pourquoi ce logiciel ?", "Introduction", "FaLightbulb");
  tx(s, "Aujourd'hui", Object.assign({ x: 0.6, y: 1.45, w: 5.6, h: 0.45, fontSize: 20, bold: true, color: C.accent1 }, A(1)));
  [["FaShuffle", "Informations dispersées", "Cahiers, tableurs, appels : difficile de savoir où est un camion et qui le conduit."],
   ["FaTriangleExclamation", "Échéances oubliées", "Assurances, immatriculations et permis expirés découverts trop tard."],
   ["FaCircleQuestion", "Coûts mal suivis", "Carburant, entretien et pièces difficiles à rattacher à chaque véhicule."]]
    .forEach((d, i) => card(s, 0.6, 2.05 + i * 1.6, 5.6, 1.4, d[0], d[1], d[2], 2 + i));
  s.addShape(pres.shapes.RIGHT_ARROW, Object.assign({ x: 6.4, y: 3.85, w: 0.6, h: 0.6, fill: { color: C.accent5 }, line: { color: C.accent5 } }, A(5)));
  tx(s, "Avec Division Terrestre Internationale", Object.assign({ x: 7.2, y: 1.45, w: 5.6, h: 0.45, fontSize: 20, bold: true, color: C.accent3 }, A(6)));
  [["FaLayerGroup", "Tout au même endroit", "Véhicules, chauffeurs, trajets, clients, carburant, entretien et finances."],
   ["FaBell", "Alertes et rappels", "Documents expirés signalés, rappels d'entretien, tableau de bord du jour."],
   ["FaChartLine", "Décisions chiffrées", "Revenus, dépenses et rapports par véhicule, chauffeur ou itinéraire."]]
    .forEach((d, i) => card(s, 7.2, 2.05 + i * 1.6, 5.55, 1.4, d[0], d[1], d[2], 7 + i));
  s.addNotes("Partir du quotidien des participants : demander comment ils savent aujourd'hui quel camion est disponible, quand expire une assurance, ou combien a coûté un véhicule le mois dernier. Le logiciel répond à ces trois besoins : centraliser, alerter, chiffrer.");

  // 3 — Modules
  s = content("18 modules, 6 grandes familles", "Introduction", "FaSitemap");
  s.addShape(pres.shapes.OVAL, Object.assign({ x: 0.6, y: 2.45, w: 2.7, h: 2.7, fill: { color: C.accent1 }, line: { color: C.accent1 } }, A(1)));
  tx(s, [{ text: "18", options: { fontSize: 64, bold: true, breakLine: true } }, { text: "modules", options: { fontSize: 18 } }], Object.assign({ x: 0.6, y: 2.45, w: 2.7, h: 2.7, align: "center", valign: "middle", color: C.background1 }, A(1)));
  [["FaGaugeHigh", "Pilotage", "Dashboard · Availability · Reports"], ["FaTruck", "Flotte", "Vehicle's · Vehicle Vendor's · Tracking · Geofence"], ["FaUsers", "Personnes", "Driver's · Employee"],
   ["FaRoute", "Activité commerciale", "Trips · Customer · Coupon Management"], ["FaWrench", "Exploitation", "Maintenance · Parts Stock · Fuel · Reminder"], ["FaCoins", "Finances et système", "Accounts · Setting's"]]
    .forEach((d, i) => card(s, 3.75 + (i % 2) * 4.55, 1.5 + Math.floor(i / 2) * 1.8, 4.35, 1.6, d[0], d[1], d[2], 2 + i));
  s.addNotes("Présenter la carte du logiciel : 18 modules dans le menu latéral, regroupés ici en 6 familles pour faciliter la lecture. Les libellés restent en anglais dans l'application : on les garde tels quels pour que chacun s'y retrouve à l'écran. Chaque famille est détaillée dans les diapositives suivantes.");

  // 4 à 14 — Modules
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
    [["FaIdCard", "Fiche véhicule", "Immatriculation, modèle, châssis, moteur, compteur, photo et documents."], ["FaCalendarXmark", "Échéances", "Expiration de l'immatriculation et de l'assurance, signalée en rouge."],
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
  tx(s, [{ text: "9 modes de facturation : ", options: { bold: true, color: C.accent1 } }, { text: "fixe · tonne · kg · km · trajet · jour · heure · litre · sac" }], Object.assign({ x: 0.6, y: 6.62, w: 8.6, h: 0.3, fontSize: 13, margin: 0 }, A(9)));

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
  tx(s, "Choisir une période, un filtre, puis « Generate Report ».", { x: 0.6, y: 1.4, w: 12, h: 0.45, fontSize: 17, italic: true, color: C.accent2 });
  [["FaClipboardList", "Réservations"], ["FaScaleBalanced", "Revenus et dépenses"], ["FaGasPump", "Carburant"], ["FaUserTie", "Chauffeurs"], ["FaTicket", "Coupons"], ["FaBell", "Rappels"],
   ["FaWrench", "Entretien"], ["FaArrowTrendUp", "Top revenus"], ["FaArrowTrendDown", "Top dépenses"], ["FaTruck", "Top véhicules"], ["FaRoute", "Top itinéraires"]]
    .forEach((r, i) => { const x = 0.6 + (i % 4) * 3.1, y = 2.05 + Math.floor(i / 4) * 1.5, top = i >= 7;
      rbox(s, x, y, 2.9, 1.25, top ? C.text2 : C.background2, i + 1);
      iconDot(s, r[0], x + 0.2, y + 0.3, 0.65, top ? C.accent4 : C.accent1, top ? HEX.navy : HEX.white, i + 1);
      tx(s, r[1], Object.assign({ x: x + 1.0, y, w: 1.8, h: 1.25, fontSize: 15, bold: true, valign: "middle", color: top ? C.background1 : C.text2 }, A(i + 1))); });
  tx(s, "Les « Top » classent les postes qui rapportent ou coûtent le plus.", { x: 9.9, y: 5.05, w: 2.85, h: 1.25, fontSize: 13, italic: true, color: C.accent5, valign: "middle" });
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

  // 15 — Ce que ça change
  pres.addSection({ title: "Conclusion" });
  s = content("Ce que ça change pour chacun", "Conclusion", "FaStar");
  [["FaUserTie", "Pour la direction", C.accent1, ["Vision en temps réel de l'activité", "Rentabilité par véhicule, chauffeur et itinéraire", "Coûts de carburant et d'entretien maîtrisés", "Décisions appuyées sur des rapports"]],
   ["FaUsers", "Pour les agents", C.accent2, ["Une seule application au lieu de plusieurs outils", "Informations retrouvées en quelques secondes", "Rappels automatiques des échéances", "Moins de saisie grâce aux imports et exports"]]]
    .forEach((d, i) => { const x = 0.6 + i * 6.2, st = i * 5 + 1;
      rbox(s, x, 1.5, 5.9, 5.2, C.background2, st);
      iconDot(s, d[0], x + 0.3, 1.75, 0.85, d[2], HEX.white, st);
      tx(s, d[1], Object.assign({ x: x + 1.35, y: 1.75, w: 4.3, h: 0.85, fontSize: 22, bold: true, color: C.text2, valign: "middle" }, A(st)));
      d[3].forEach((t, j) => { const y = 2.95 + j * 0.88;
        s.addImage(Object.assign({ data: ic("FaCircleCheck", i ? HEX.teal : HEX.red), x: x + 0.35, y: y + 0.06, w: 0.34, h: 0.34 }, A(st + 1 + j)));
        tx(s, t, Object.assign({ x: x + 0.85, y, w: 4.85, h: 0.8, fontSize: 17, margin: 0 }, A(st + 1 + j))); }); });
  s.addNotes("Synthèse orientée bénéfices, avant la démonstration. Laisser la direction et les agents réagir : quel bénéfice compte le plus pour eux ?");

  // 16 — Démo
  s = content("Démonstration : un circuit de travail", "Conclusion", "FaPlay");
  tx(s, "En direct dans l'application", { x: 0.6, y: 1.4, w: 12, h: 0.45, fontSize: 17, italic: true, color: C.accent2 });
  const steps = [["FaUserPlus", "Client", "Customer › Add Customer"], ["FaRoute", "Trajet", "Trips › Add Trips"], ["FaGasPump", "Carburant", "Fuel › Add Fuel"], ["FaReceipt", "Transaction", "Accounts › Add Transactions"], ["FaChartLine", "Tableau de bord et rapport", "Dashboard · Reports"]];
  steps.forEach((d, i) => { const x = 0.6 + i * 2.5, st = i * 2 + 1;
    iconDot(s, d[0], x + 0.5, 2.15, 1.3, C.accent1, HEX.white, st);
    s.addShape(pres.shapes.OVAL, Object.assign({ x: x + 1.5, y: 2.05, w: 0.45, h: 0.45, fill: { color: C.accent4 }, line: { color: C.background1, width: 2 } }, A(st)));
    tx(s, String(i + 1), Object.assign({ x: x + 1.5, y: 2.05, w: 0.45, h: 0.45, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0 }, A(st)));
    tx(s, [{ text: d[1], options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } }, { text: d[2], options: { fontSize: 13, color: C.accent5 } }], Object.assign({ x, y: 3.6, w: 2.3, h: 1.2, align: "center" }, A(st)));
    if (i < steps.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, Object.assign({ x: x + 2.0, y: 2.6, w: 0.42, h: 0.4, fill: { color: C.accent5 }, line: { color: C.accent5 } }, A(st + 1))); });
  rbox(s, 0.6, 5.2, 12.1, 1.4, C.text2, 10);
  iconDot(s, "FaBullseye", 0.85, 5.5, 0.8, C.accent4, HEX.navy, 10);
  tx(s, [{ text: "Objectif  ", options: { bold: true, color: C.accent4 } }, { text: "montrer qu'une seule saisie alimente tout le reste : le trajet apparaît dans le calendrier, le carburant et la transaction dans les finances, puis dans le tableau de bord et les rapports." }],
    Object.assign({ x: 1.9, y: 5.2, w: 10.6, h: 1.4, fontSize: 15, color: C.background1, valign: "middle" }, A(10)));
  s.addNotes("DÉROULÉ : 1) créer un client fictif ; 2) créer un trajet pour ce client avec un véhicule et un chauffeur, choisir un mode de facturation ; 3) enregistrer un plein pour ce véhicule ; 4) enregistrer la transaction de revenu du trajet (catégorie Trip Income) ; 5) revenir au tableau de bord et générer un rapport sur la période. Utiliser uniquement des données de test, et les supprimer après la démonstration.");

  // 17 — Questions
  s = pres.addSlide({ masterName: "SOMBRE", sectionTitle: "Conclusion" });
  s.addShape(pres.shapes.OVAL, { x: 8.6, y: -0.9, w: 6.6, h: 6.6, fill: { color: C.accent1 }, line: { color: C.accent1 } });
  s.addShape(pres.shapes.OVAL, { x: 10.4, y: 4.9, w: 3.6, h: 3.6, fill: { color: C.accent2, transparency: 30 }, line: { color: C.accent2, transparency: 30 } });
  s.addImage(Object.assign({ data: ic("FaComments", HEX.white), x: 10.0, y: 1.1, w: 3.4, h: 3.4 }, A(1)));
  s.addText("Questions et échanges", { placeholder: "title" });
  tx(s, "Merci de votre attention", Object.assign({ x: 0.8, y: 4.05, w: 7.4, h: 0.6, fontSize: 20, color: C.accent4 }, A(2)));
  s.addNotes("Recueillir les questions et les besoins : modules prioritaires, profils de droits à créer, données à importer (véhicules, chauffeurs), formation des agents.");

  // 18 — Réserve : coûts (masquée, à afficher seulement sur demande de la direction)
  pres.addSection({ title: "Réserve" });
  s = content("Coûts de fonctionnement", "Réserve", "FaFileInvoiceDollar");
  s.hidden = true;
  [["FaServer", "Hébergement du serveur", "par an", 1.5], ["FaHeadset", "Maintenance", "par mois, formation des agents comprise", 4.1]].forEach((d, i) => {
    rbox(s, 0.6, d[3], 5.6, 2.3, C.background2, i + 1);
    iconDot(s, d[0], 0.85, d[3] + 0.25, 0.6, C.accent1, HEX.white, i + 1);
    tx(s, d[1], Object.assign({ x: 1.6, y: d[3] + 0.25, w: 4.4, h: 0.6, fontSize: 16, bold: true, color: C.text2, valign: "middle" }, A(i + 1)));
    tx(s, [{ text: "100 000 FCFA", options: { bold: true, fontSize: 32, color: C.accent1, breakLine: true } }, { text: d[2], options: { fontSize: 16 } }], Object.assign({ x: 0.85, y: d[3] + 1.0, w: 5.2, h: 1.2 }, A(i + 1))); });
  tx(s, "Calendrier de paiement de la maintenance", Object.assign({ x: 6.7, y: 1.5, w: 6, h: 0.45, bold: true, fontSize: 17, color: C.text2 }, A(3)));
  rbox(s, 6.7, 2.15, 6.05, 1.9, C.text2, 3);
  tx(s, [{ text: "Premier paiement : 300 000 FCFA", options: { bold: true, fontSize: 20, color: C.accent4, breakLine: true } }, { text: "1er mois de maintenance + 2 mois payés d'avance (mois 2 et 3)", options: { fontSize: 15 } }], Object.assign({ x: 6.9, y: 2.25, w: 5.7, h: 1.7, color: C.background1, valign: "middle" }, A(3)));
  s.addShape(pres.shapes.DOWN_ARROW, Object.assign({ x: 9.5, y: 4.15, w: 0.45, h: 0.45, fill: { color: C.accent5 }, line: { color: C.accent5 } }, A(4)));
  rbox(s, 6.7, 4.7, 6.05, 1.7, C.background2, 4);
  tx(s, [{ text: "À partir du 4ᵉ mois : 100 000 FCFA", options: { bold: true, fontSize: 20, color: C.text2, breakLine: true } }, { text: "réglés chaque mois", options: { fontSize: 15 } }], Object.assign({ x: 6.9, y: 4.8, w: 5.7, h: 1.5, valign: "middle" }, A(4)));
  s.addNotes("DIAPOSITIVE MASQUÉE : elle ne s'affiche pas pendant le diaporama. Pour l'afficher si la direction pose la question, taper 18 puis Entrée pendant le diaporama. Ne présenter que ces montants, sans autres frais.");

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
    // chaque forme : <p:sp> ou <p:pic> dont le cNvPr porte un nom « animN-… »
    const re = /<p:(sp|pic)>([\s\S]*?)<\/p:\1>/g; let m;
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
    ...["FaTruckFast", "FaLightbulb", "FaShuffle", "FaTriangleExclamation", "FaCircleQuestion", "FaLayerGroup", "FaBell", "FaChartLine", "FaSitemap", "FaGaugeHigh", "FaTruck",
      "FaUsers", "FaRoute", "FaWrench", "FaCoins", "FaHashtag", "FaTruckFront", "FaCalendarDays", "FaIdCard", "FaCalendarXmark", "FaTag", "FaUserTie", "FaAddressCard", "FaIdBadge",
      "FaHandshake", "FaFileContract", "FaMapLocationDot", "FaLocationCrosshairs", "FaClockRotateLeft", "FaDrawPolygon", "FaBullhorn", "FaClipboardList", "FaSignal", "FaTicket",
      "FaScrewdriverWrench", "FaUserGear", "FaBoxesStacked", "FaGasPump", "FaFillDrip", "FaArrowRightArrowLeft", "FaOilCan", "FaStore", "FaBuildingColumns", "FaReceipt",
      "FaMoneyBillTransfer", "FaTags", "FaChartPie", "FaScaleBalanced", "FaUserShield", "FaUserPlus", "FaListCheck", "FaCubes", "FaToggleOn", "FaGear", "FaBuilding", "FaEnvelope",
      "FaFileInvoice", "FaDatabase", "FaStar", "FaPlay", "FaComments", "FaFileInvoiceDollar", "FaServer", "FaHeadset"].map(n => [n, HEX.white]),
    ["FaCircleCheck", HEX.amber], ["FaCircleCheck", HEX.red], ["FaCircleCheck", HEX.teal], ["FaBullseye", HEX.navy],
    ...["FaArrowTrendUp", "FaArrowTrendDown", "FaTruck", "FaRoute"].map(n => [n, HEX.navy]),
  ]);
  await build();
  await applyTheme(OUT, THEME);
  await animate();
  console.log("ok");
})();
