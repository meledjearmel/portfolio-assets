const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/2d22e13e-03a7-4f53-876d-c8368813d997_77b68b09-cce3-4263-bfaa-f3c76c1f97e3/pptx/scripts/apply_theme.js");
const OUT = "/home/user/portfolio-assets/formations/dti/formation-division-terrestre-internationale.pptx";
const THEME = { name: "DTI Formation", headFontFace: "Arial", bodyFontFace: "Calibri",
  colors: { dk1: "1B2430", lt1: "FFFFFF", dk2: "24364A", lt2: "EEF1F4", accent1: "E8A317", accent2: "2E6F8E",
    accent3: "4E8A5B", accent4: "B5482F", accent5: "7A8796", accent6: "C9D3DD", hlink: "2E6F8E", folHlink: "7A8796" } };
const HEX = { amber: "E8A317", navy: "24364A", ink: "1B2430", teal: "2E6F8E", green: "4E8A5B", red: "B5482F", grey: "7A8796", light: "EEF1F4", line: "C9D3DD", white: "FFFFFF" };
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Formation à Division Terrestre Internationale";
const C = pres.SchemeColor;

pres.defineSlideMaster({ title: "TITRE", background: { color: HEX.navy },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.0, w: 11.7, h: 2.0, fontSize: 38, bold: true, color: C.background1, valign: "bottom" }, text: "" } }] });
pres.defineSlideMaster({ title: "CONTENU", background: { color: HEX.white }, margin: [0.5, 0.6, 0.6, 0.6],
  slideNumber: { x: 12.3, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
  objects: [
    { text: { text: "Division Terrestre Internationale — Formation initiale", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent5 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontSize: 30, bold: true, color: C.text2, valign: "middle" }, text: "" } },
  ] });

let n = 0;
function content(title, sec) { const s = pres.addSlide({ masterName: "CONTENU", sectionTitle: sec }); s.addText(title, { placeholder: "title" }); return s; }
function tx(s, t, o) { s.addText(t, Object.assign({ isTextBox: true, fontSize: 15, color: C.text1, valign: "top" }, o)); }
function bullets(s, items, o) {
  s.addText(items.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < items.length - 1, paraSpaceAfter: 6 } })),
    Object.assign({ isTextBox: true, fontSize: 15, color: C.text1, valign: "top" }, o));
}
function badge(s, num, x, y, d, fill) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill || C.accent1 }, line: { color: fill || C.accent1 } });
  s.addText(String(num), { isTextBox: true, x, y, w: d, h: d, fontSize: d > 0.5 ? 18 : 13, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0 });
}
// Emplacement de capture + légende numérotée
function capture(s, ecran, markers, box) {
  const b = Object.assign({ x: 0.6, y: 1.45, w: 7.4, h: 5.3 }, box);
  s.addShape(pres.shapes.RECTANGLE, { objectName: "Capture à ajouter", x: b.x, y: b.y, w: b.w, h: b.h, fill: { color: C.background2 }, line: { color: C.accent5, width: 1.25, dashType: "dash" } });
  tx(s, [{ text: "Capture à ajouter", options: { bold: true, fontSize: 20, color: C.accent4, breakLine: true } },
    { text: "Écran : " + ecran, options: { fontSize: 14, color: C.text2, breakLine: true } },
    { text: "Masquer noms, numéros et données personnelles (ou remplacer par des exemples fictifs).", options: { fontSize: 11, italic: true, color: C.accent5 } }],
    { x: b.x + 0.4, y: b.y + b.h / 2 - 0.8, w: b.w - 0.8, h: 1.6, align: "center", valign: "middle" });
  if (markers && markers.length) {
    // repères à placer sur la capture
    markers.forEach((m, i) => badge(s, i + 1, b.x + 0.25 + i * 0.55, b.y + 0.25, 0.42));
    const lx = b.x + b.w + 0.35, lw = 12.73 - lx;
    tx(s, "Repères à placer sur la capture", { x: lx, y: b.y, w: lw, h: 0.4, fontSize: 13, bold: true, color: C.accent5, margin: 0 });
    const step = Math.min(0.85, (b.h - 0.5) / markers.length);
    markers.forEach((m, i) => {
      const y = b.y + 0.5 + i * step;
      badge(s, i + 1, lx, y, 0.42);
      tx(s, m, { x: lx + 0.55, y: y - 0.04, w: lw - 0.55, h: step, fontSize: 14, margin: 0 });
    });
  }
}
function card(s, x, y, w, h, head, body, fill) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill || C.background2 }, line: { color: fill || C.background2 } });
  tx(s, [{ text: head, options: { bold: true, fontSize: 16, color: C.text2, breakLine: true, paraSpaceAfter: 6 } }, { text: body, options: { fontSize: 14 } }], { x: x + 0.2, y: y + 0.15, w: w - 0.4, h: h - 0.3 });
}
const NOM = "« Track » peut encore apparaître dans certains menus : c'est le nom technique de Division Terrestre Internationale.";

// 1 — Titre
pres.addSection({ title: "Introduction" });
let s = pres.addSlide({ masterName: "TITRE", sectionTitle: "Introduction" });
s.addText("Formation à Division Terrestre Internationale : gestion de flotte, du personnel et des finances", { placeholder: "title" });
tx(s, "Formation initiale — environ 2 heures, avec démonstrations et exercices en groupe", { x: 0.8, y: 4.2, w: 11.7, h: 0.5, fontSize: 18, color: C.accent1 });
tx(s, [{ text: "Organisation : [à compléter]", options: { breakLine: true } }, { text: "Formateur : [à compléter]", options: { breakLine: true } }, { text: "Date : [à compléter]" }],
  { x: 0.8, y: 5.1, w: 6, h: 1.3, fontSize: 16, color: C.background1, paraSpaceAfter: 4 });
s.addNotes("Accueillir les participants, se présenter et faire un rapide tour de table (nom, service, utilisation prévue de l'application). Compléter les champs Organisation, Formateur et Date avant la séance. Annoncer la durée (environ 2 h) et l'alternance entre explications, démonstrations et exercices en groupe.");

// 2 — Programme
s = content("Objectifs et déroulé de la séance", "Introduction");
tx(s, "À la fin de la formation, vous saurez :", { x: 0.6, y: 1.4, w: 6, h: 0.4, bold: true, color: C.text2 });
bullets(s, ["Comprendre les fonctions de base de l'application", "Retrouver les informations sur les camions, les chauffeurs et le personnel", "Comprendre les informations financières présentées", "Savoir qui est responsable de chaque opération", "Définir en groupe les permissions nécessaires selon les rôles", "Connaître les coûts d'hébergement et de maintenance"], { x: 0.6, y: 1.9, w: 6.0, h: 4.8 });
const prog = [["15 min", "Introduction et interface"], ["45 min", "Camions, chauffeurs, personnel, trajets"], ["20 min", "Carburant, finances, entretien, suivi"], ["30 min", "Ateliers : rôles, permissions, exercice"], ["10 min", "Coûts, bonnes pratiques, quiz, questions"]];
prog.forEach((p, i) => { const y = 1.5 + i * 1.02;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.2, y, w: 5.5, h: 0.85, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 } });
  tx(s, p[0], { x: 7.35, y, w: 1.3, h: 0.85, bold: true, color: C.accent2, valign: "middle" });
  tx(s, p[1], { x: 8.65, y, w: 3.9, h: 0.85, valign: "middle" }); });
s.addNotes("Présenter les objectifs et le déroulé. Les durées sont indicatives : les ajuster selon le groupe. Préciser que certains écrans seront vus plus en détail lors de séances complémentaires (feuille de route en fin de présentation).");

// 3 — Pourquoi
s = content("Pourquoi utiliser l'application ?", "Introduction");
tx(s, "Un seul outil pour suivre l'activité et retrouver la bonne information au bon moment.", { x: 0.6, y: 1.35, w: 12, h: 0.5, fontSize: 17, italic: true, color: C.accent2 });
[["Camions", "Retrouver un véhicule, son statut, son chauffeur et ses trajets."], ["Chauffeurs", "Consulter les profils, les permis et le suivi des chauffeurs."], ["Personnel", "Accéder aux rubriques des employés, de la présence et de la paie."], ["Finances", "Consulter les transactions et les coûts selon les écrans disponibles."]]
  .forEach((d, i) => { const x = 0.6 + i * 3.1; badge(s, i + 1, x + 0.2, 2.25, 0.7); card(s, x, 3.1, 2.85, 2.6, d[0], d[1]); });
tx(s, NOM, { x: 0.6, y: 6.1, w: 12, h: 0.5, fontSize: 12, color: C.accent5 });
s.addNotes("Demander aux participants comment ces informations sont suivies aujourd'hui (papier, tableurs, appels). Présenter les quatre domaines. Prévenir que le nom technique « Track » peut apparaître dans certains menus : il s'agit bien de Division Terrestre Internationale. On garde les libellés anglais des menus pour se repérer à l'écran.");

// 4 — Interface
pres.addSection({ title: "Interface" });
s = content("Découvrir l'interface", "Interface");
capture(s, "page d'accueil après connexion, menu latéral déplié", ["Menu latéral : accès à toutes les rubriques (camions, chauffeurs, trajets, carburant…)", "Tableau de bord : vue d'ensemble de l'activité", "Zone de navigation principale (en-tête, recherche, compte), selon la capture", "Mention « Track » éventuelle : nom technique de l'application"]);
s.addNotes("DÉMONSTRATION : se connecter et parcourir le menu latéral sans ouvrir de formulaire. Lire à voix haute les libellés des menus tels qu'ils apparaissent. Une fois la capture ajoutée, placer les repères 1 à 4 avec des flèches ou des encadrés. Ne décrire que les zones réellement visibles sur la capture.");

// 5 — Tableau de bord
s = content("Lire le tableau de bord", "Interface");
capture(s, "tableau de bord complet", ["Flotte : indicateurs sur les véhicules", "Chauffeurs : indicateurs sur les chauffeurs", "Trajets : activité en cours et à venir", "Carburant : consommation et dépenses", "Transactions : mouvements financiers", "Alertes : points à surveiller"]);
tx(s, "Les chiffres affichés sont des exemples d'écran, pas des résultats réels.", { x: 0.6, y: 6.8, w: 7.4, h: 0.2, fontSize: 11, italic: true, color: C.accent4, margin: 0 });
s.addNotes("Expliquer chaque indicateur à partir de la capture, en rappelant que les chiffres sont des exemples d'écran. EXERCICE RAPIDE : demander à un participant de dire où il regarderait d'abord pour savoir si une alerte demande une action. Ne pas interpréter de calcul qui ne serait pas visible à l'écran.");

// 6 — Camions retrouver
pres.addSection({ title: "Camions" });
s = content("Gestion des camions : retrouver un véhicule", "Camions");
capture(s, "Vehicle Management — liste des véhicules", ["Menu « Vehicle Management »", "Recherche par nom, modèle ou immatriculation", "Filtres de statut", "Vue en cartes ou en liste (si visible)"]);
s.addNotes("DÉMONSTRATION : ouvrir « Vehicle Management », rechercher un véhicule par immatriculation, puis par modèle. Appliquer un filtre de statut. Basculer entre la vue en cartes et la vue en liste si l'écran le permet. Utiliser uniquement des immatriculations fictives ou masquées dans les supports.");

// 7 — Fiche camion
s = content("Gestion des camions : lire une fiche", "Camions");
capture(s, "fiche détaillée d'un véhicule", ["Statut du véhicule", "Chauffeur associé", "Carburant courant", "Type de véhicule"], { w: 6.6 });
s.addNotes("Lire une fiche avec le groupe. Insister sur la différence entre consulter une fiche (aucun risque) et la modifier (action réservée selon les permissions qui seront validées). Ne pas cliquer sur un bouton de modification pendant la démonstration. Si un bouton de modification est visible sur la capture, le signaler sans détailler la procédure.");
tx(s, [{ text: "Consulter", options: { bold: true, color: C.accent3 } }, { text: " une fiche ≠ " }, { text: "modifier", options: { bold: true, color: C.accent4 } }, { text: " une fiche (selon les droits validés)." }], { x: 0.6, y: 6.8, w: 6.6, h: 0.2, fontSize: 12, margin: 0 });

// 8 — Groupes et itinéraires
s = content("Groupes et itinéraires", "Camions");
card(s, 0.6, 1.5, 5.9, 2.2, "Vehicle Group", "Rubrique de regroupement des véhicules. Procédure détaillée présentée seulement avec une capture dédiée.");
card(s, 6.8, 1.5, 5.9, 2.2, "Route", "Rubrique des itinéraires. Procédure détaillée présentée seulement avec une capture dédiée.");
capture(s, "Vehicle Group (liste) et Route (liste)", [], { x: 0.6, y: 4.0, w: 12.1, h: 2.7 });
s.addNotes("Montrer seulement où se trouvent les menus « Vehicle Group » et « Route » dans le menu latéral. Ne pas décrire de procédure de création ou de modification tant que les captures correspondantes ne sont pas fournies. Ce sujet est approfondi dans une séance complémentaire.");

// 9 — Chauffeurs retrouver
pres.addSection({ title: "Chauffeurs et personnel" });
s = content("Gestion des chauffeurs : retrouver un profil", "Chauffeurs et personnel");
capture(s, "Driver Info — liste des chauffeurs", ["Menu « Driver Info »", "Filtres de statut", "Recherche par nom, permis ou contact"]);
s.addNotes("DÉMONSTRATION : ouvrir « Driver Info », filtrer par statut, puis rechercher un chauffeur par nom, numéro de permis ou contact. Sur les supports, remplacer les noms, numéros de permis et contacts par des exemples fictifs (ex. « Chauffeur A », « Permis 000-EXEMPLE »).");

// 10 — Infos chauffeurs
s = content("Informations et suivi des chauffeurs", "Chauffeurs et personnel");
capture(s, "fiche d'un chauffeur", ["Statut du chauffeur", "Permis de conduire", "Date d'expiration du permis", "Date d'inscription"], { w: 6.6 });
tx(s, [{ text: "Driver Performance", options: { bold: true } }, { text: " : rubrique à approfondir avec une capture dédiée." }], { x: 0.6, y: 6.8, w: 6.6, h: 0.2, fontSize: 12, margin: 0 });
s.addNotes("Lire une fiche chauffeur. Insister sur la date d'expiration du permis : c'est une information à surveiller. Citer « Driver Performance » comme rubrique existante, sans décrire ses indicateurs : ils seront présentés lors d'une séance complémentaire avec une capture.");

// 11 — Personnel
s = content("Gestion du personnel", "Chauffeurs et personnel");
[["Employés", "Rubrique des dossiers des employés."], ["Présence", "Rubrique de suivi de la présence."], ["Paie", "Rubrique liée à la paie."]]
  .forEach((d, i) => card(s, 0.6 + i * 4.1, 1.5, 3.8, 1.6, d[0], d[1]));
capture(s, "rubriques Employés, Présence et Paie (relever les libellés exacts des menus)", [], { x: 0.6, y: 3.4, w: 12.1, h: 3.3 });
s.addNotes("Situer les rubriques du personnel dans le menu, avec leurs libellés exacts. Ne pas décrire les champs ni les procédures (ajout d'un employé, saisie de la présence, calcul de la paie) tant que les captures ne sont pas fournies. Noter les questions des participants RH pour la séance complémentaire dédiée. Les données du personnel sont sensibles : à masquer systématiquement.");

// 12 — Trajets statuts
pres.addSection({ title: "Trajets" });
s = content("Gestion des trajets : les statuts", "Trajets");
[["Upcoming", "À venir", C.accent2], ["Ongoing", "En cours", C.accent1], ["Completed", "Terminé", C.accent3], ["Cancelled", "Annulé", C.accent4]]
  .forEach((d, i) => { const x = 0.6 + i * 3.1;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.6, w: 2.8, h: 1.5, rectRadius: 0.1, fill: { color: d[2] }, line: { color: d[2] } });
    tx(s, [{ text: d[0], options: { bold: true, fontSize: 22, breakLine: true } }, { text: d[1], options: { fontSize: 15 } }], { x, y: 1.6, w: 2.8, h: 1.5, align: "center", valign: "middle", color: C.background1 });
    if (i < 3) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 2.83, y: 2.2, w: 0.24, h: 0.3, fill: { color: C.accent5 }, line: { color: C.accent5 } }); });
capture(s, "liste des trajets avec les onglets ou filtres de statut", [], { x: 0.6, y: 3.5, w: 12.1, h: 3.2 });
s.addNotes("Présenter les quatre statuts tels qu'ils apparaissent dans l'application. Un trajet annulé (« Cancelled ») ne passe pas forcément par les autres statuts. Demander : « Où chercheriez-vous un trajet prévu demain ? » (réponse : Upcoming).");

// 13 — Consulter un trajet
s = content("Consulter un trajet", "Trajets");
capture(s, "détail d'un trajet", ["Véhicule et identifiant du trajet", "Itinéraire", "Client", "Chauffeur et statut du chauffeur", "Dates de début et de fin", "Détails du trajet"]);
s.addNotes("Ouvrir un trajet et parcourir les informations dans l'ordre des repères. Remplacer le nom du client et du chauffeur par des exemples fictifs sur la capture. Rappeler que l'on consulte : aucune modification pendant la formation.");

// 14 — Relier
s = content("Relier camion, chauffeur et trajet", "Trajets");
[["Camion", "Vehicle Management"], ["Chauffeur", "Driver Info"], ["Trajet", "liste des trajets"]].forEach((d, i) => { const x = 0.9 + i * 4.2;
  s.addShape(pres.shapes.OVAL, { x: x + 0.6, y: 1.6, w: 2.2, h: 2.2, fill: { color: C.text2 }, line: { color: C.text2 } });
  tx(s, [{ text: d[0], options: { bold: true, fontSize: 20, breakLine: true } }, { text: d[1], options: { fontSize: 12 } }], { x: x + 0.6, y: 1.6, w: 2.2, h: 2.2, align: "center", valign: "middle", color: C.background1 });
  if (i < 2) s.addShape(pres.shapes.LEFT_RIGHT_ARROW, { x: x + 3.0, y: 2.45, w: 1.6, h: 0.5, fill: { color: C.accent1 }, line: { color: C.accent1 } }); });
tx(s, "Démonstration en consultation uniquement — aucune création, aucune modification", { x: 0.6, y: 4.2, w: 12.1, h: 0.4, bold: true, color: C.accent4, align: "center" });
bullets(s, ["Partir d'un trajet : noter le véhicule et le chauffeur affichés", "Ouvrir la fiche du véhicule : vérifier le chauffeur associé et le statut", "Ouvrir le profil du chauffeur : vérifier le statut et la validité du permis"], { x: 2.0, y: 4.8, w: 9.5, h: 1.8 });
s.addNotes("DÉMONSTRATION GUIDÉE : 1) ouvrir un trajet « Ongoing » ; 2) relever le véhicule et le chauffeur ; 3) retrouver le véhicule dans « Vehicle Management » et comparer le chauffeur associé ; 4) ouvrir le profil dans « Driver Info » et vérifier la date d'expiration du permis. Question au groupe : que faire si les informations ne concordent pas ? (réponse attendue : ne pas corriger soi-même sans droit, signaler au responsable).");

// 15 — Carburant
pres.addSection({ title: "Carburant, finances et suivi" });
s = content("Suivi du carburant", "Carburant, finances et suivi");
capture(s, "Fuel Info — liste des pleins", ["Menu « Fuel Info » et recherche", "Bouton « Add Fuel »", "Véhicule et date", "Quantité et coût", "Chauffeur", "Reçu (lorsqu'il est présent)"]);
tx(s, "Les quantités et montants visibles sont des exemples d'écran.", { x: 0.6, y: 6.8, w: 7.4, h: 0.2, fontSize: 11, italic: true, color: C.accent4, margin: 0 });
s.addNotes("Montrer la recherche dans « Fuel Info » et lire une ligne : véhicule, date, quantité, coût, chauffeur, reçu. Signaler l'existence du bouton « Add Fuel » sans faire de saisie : la procédure d'ajout et le droit de saisie seront discutés dans l'atelier sur les permissions.");

// 16 — Finances
s = content("Finances : comprendre les informations disponibles", "Carburant, finances et suivi");
capture(s, "Capture financière à ajouter : transactions, coûts, revenus ou dépenses", [], { w: 7.4 });
card(s, 8.3, 1.45, 4.4, 5.3, "Ce que l'on présentera", "Uniquement les informations visibles à l'écran : transactions, coûts, revenus ou dépenses.\n\nAucune catégorie, aucun calcul et aucune procédure comptable ne sont décrits sans capture.\n\nLes transactions apparaissent déjà comme indicateur du tableau de bord.");
s.addNotes("Faire le lien avec l'indicateur « transactions » du tableau de bord. Ne pas décrire de catégories, de calculs ni de procédures comptables tant que l'écran financier n'est pas fourni. Noter les questions des responsables finance pour la séance complémentaire dédiée.");

// 17 — Entretien
s = content("Entretien, rappels et pièces", "Carburant, finances et suivi");
[["Maintenance", "Suivi de l'entretien des véhicules."], ["Reminder", "Rappels des échéances à ne pas manquer."], ["Parts Stock", "Suivi du stock de pièces."]]
  .forEach((d, i) => { const x = 0.6 + i * 4.1; badge(s, i + 1, x, 1.6, 0.6, C.accent2); card(s, x, 2.4, 3.8, 2.0, d[0], d[1]); });
tx(s, "Rôle général uniquement : les procédures détaillées seront vues lors d'une séance dédiée, avec captures.", { x: 0.6, y: 4.9, w: 12.1, h: 0.5, color: C.accent5, italic: true });
tx(s, "Captures à ajouter : Maintenance, Reminder, Parts Stock", { x: 0.6, y: 5.6, w: 12.1, h: 0.5, bold: true, color: C.accent4 });
s.addNotes("Situer les trois menus et expliquer leur rôle général. Ne pas détailler la création d'un entretien, d'un rappel ou d'un mouvement de stock. Demander qui, dans l'organisation, s'occupe aujourd'hui de l'entretien et des pièces : utile pour l'atelier sur les rôles.");

// 18 — Suivi
s = content("Suivi, zones et alertes", "Carburant, finances et suivi");
[["Tracking", "Suivre la position des véhicules."], ["Geofence", "Définir des zones géographiques de référence."], ["Alerts", "Être averti des événements à surveiller."]]
  .forEach((d, i) => { const x = 0.6 + i * 4.1; badge(s, i + 1, x, 1.6, 0.6, C.accent3); card(s, x, 2.4, 3.8, 2.0, d[0], d[1]); });
tx(s, "Utilité générale seulement : le fonctionnement précis des écrans sera présenté avec des captures.", { x: 0.6, y: 4.9, w: 12.1, h: 0.5, color: C.accent5, italic: true });
tx(s, "Captures à ajouter : Tracking, Geofence, Alerts", { x: 0.6, y: 5.6, w: 12.1, h: 0.5, bold: true, color: C.accent4 });
s.addNotes("Présenter l'utilité générale des trois rubriques. Faire le lien avec l'indicateur « alertes » du tableau de bord. Ne pas décrire de réglages (création de zone, seuils d'alerte) sans capture.");

// 19 — Qui utilise quoi
pres.addSection({ title: "Ateliers" });
s = content("Atelier : qui utilise quoi ?", "Ateliers");
tx(s, "En groupes de 3 à 4 — 10 minutes", { x: 0.6, y: 1.35, w: 12, h: 0.4, bold: true, color: C.accent2 });
const acts = ["Consulter les camions", "Mettre à jour les dossiers", "Suivre les chauffeurs", "Consulter les dépenses", "Produire des rapports"];
acts.forEach((a, i) => { const y = 1.95 + i * 0.95;
  badge(s, i + 1, 0.6, y + 0.1, 0.5);
  tx(s, a, { x: 1.3, y, w: 4.5, h: 0.7, fontSize: 17, valign: "middle" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.0, y, w: 6.7, h: 0.7, rectRadius: 0.06, fill: { color: C.background2 }, line: { color: C.accent6, dashType: "dash" } });
  tx(s, "Qui le fait aujourd'hui ? Qui devrait le faire ?", { x: 6.2, y, w: 6.4, h: 0.7, fontSize: 13, color: C.accent5, valign: "middle", italic: true }); });
s.addNotes("Former des groupes en mélangeant les services. Chaque groupe indique, pour chaque activité, quel métier la réalise aujourd'hui et lequel devrait la réaliser dans l'application. Mise en commun de 5 minutes : noter les réponses au tableau. Elles servent de base directe à l'atelier sur les permissions.");

// 20 — Permissions
s = content("Atelier : matrice des permissions à valider", "Ateliers");
const H = { bold: true, color: C.background1, fill: { color: C.text2 }, align: "center", fontSize: 13 };
const cell = t => ({ text: t, options: { align: "center", fontSize: 13, color: C.text1 } });
const rows = [["Rôle indicatif", "Consulter", "Ajouter", "Modifier", "Valider", "Administrer"].map((t, i) => ({ text: t, options: Object.assign({}, H, i === 0 ? { align: "left" } : {}) }))];
["Administrateur", "Responsable de flotte", "Agent d'exploitation", "RH", "Finance", "Chauffeur"].forEach(r => rows.push([{ text: r, options: { bold: true, fontSize: 13, color: C.text2 } }, cell("?"), cell("?"), cell("?"), cell("?"), cell("?")]));
s.addTable(rows, { x: 0.6, y: 1.45, w: 8.2, colW: [2.45, 1.15, 1.15, 1.15, 1.15, 1.15], rowH: 0.55, border: { type: "solid", pt: 0.75, color: HEX.line }, fill: { color: HEX.white } });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.2, y: 1.45, w: 3.5, h: 3.85, rectRadius: 0.08, fill: { color: C.accent1 }, line: { color: C.accent1 } });
tx(s, [{ text: "Base de discussion", options: { bold: true, fontSize: 17, breakLine: true, paraSpaceAfter: 8 } }, { text: "Ces droits ne sont pas configurés dans l'application. La matrice est à valider par l'organisation, puis à vérifier dans l'application avant toute diffusion.", options: { fontSize: 14 } }], { x: 9.4, y: 1.6, w: 3.1, h: 3.6 });
tx(s, "Remplir chaque case par domaine si besoin : camions, chauffeurs, personnel, finances.", { x: 0.6, y: 5.55, w: 12.1, h: 0.5, fontSize: 13, color: C.accent5, italic: true });
s.addNotes("Construire la matrice avec les participants, rôle par rôle, à partir des réponses de l'atelier précédent. Pour chaque case : « Ce rôle a-t-il besoin de ce droit pour faire son travail ? ». Principe à rappeler : donner le minimum nécessaire. Préciser à plusieurs reprises qu'il s'agit d'une proposition : les droits réels doivent être validés par l'organisation et vérifiés dans l'application avant diffusion de la matrice. Photographier ou recopier la matrice remplie en fin d'atelier.");

// 21 — Exercice
s = content("Exercice pratique en équipe", "Ateliers");
[["A", "Retrouver un camion à partir de son immatriculation et noter son statut."], ["B", "Vérifier le chauffeur associé à un camion et la validité de son permis."], ["C", "Consulter un trajet en cours : véhicule, itinéraire, dates, chauffeur."], ["D", "Repérer une information de carburant : date, quantité, coût, reçu."], ["E", "Dire qui devrait traiter une opération financière ou RH donnée."]]
  .forEach((d, i) => { const col = i % 3, row = Math.floor(i / 3), x = 0.6 + col * 4.1, y = 1.5 + row * 2.6;
    card(s, x, y, 3.8, 2.3, "Scénario " + d[0], d[1]); });
s.addNotes("Attribuer un scénario par groupe (15 minutes). Utiliser un environnement de démonstration ou des données fictives. Consigne : consulter uniquement, ne rien créer ni modifier. Chaque groupe présente en une minute son chemin dans les menus. Scénario E : s'appuyer sur la matrice de l'atelier précédent ; la bonne réponse est le rôle retenu par le groupe, à confirmer par l'organisation.");

// 22 — Coûts
pres.addSection({ title: "Coûts et bonnes pratiques" });
s = content("Coûts de fonctionnement", "Coûts et bonnes pratiques");
card(s, 0.6, 1.5, 3.9, 2.4, "Hébergement du serveur", "");
tx(s, [{ text: "100 000 FCFA", options: { bold: true, fontSize: 30, color: C.accent2, breakLine: true } }, { text: "par an", options: { fontSize: 15 } }], { x: 0.8, y: 2.3, w: 3.5, h: 1.4 });
card(s, 0.6, 4.2, 3.9, 2.4, "Maintenance", "");
tx(s, [{ text: "100 000 FCFA", options: { bold: true, fontSize: 30, color: C.accent2, breakLine: true } }, { text: "par mois, formation des agents comprise", options: { fontSize: 14 } }], { x: 0.8, y: 5.0, w: 3.5, h: 1.4 });
tx(s, "Calendrier de paiement de la maintenance", { x: 5.0, y: 1.5, w: 7.7, h: 0.4, bold: true, color: C.text2 });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.0, y: 2.1, w: 4.6, h: 2.3, rectRadius: 0.08, fill: { color: C.accent1 }, line: { color: C.accent1 } });
tx(s, [{ text: "Premier paiement", options: { bold: true, fontSize: 15, breakLine: true } }, { text: "300 000 FCFA", options: { bold: true, fontSize: 28, breakLine: true } }, { text: "Mois 1 + mois 2 et 3 payés d'avance", options: { fontSize: 14 } }], { x: 5.2, y: 2.2, w: 4.2, h: 2.1, valign: "middle" });
s.addShape(pres.shapes.RIGHT_ARROW, { x: 9.7, y: 3.0, w: 0.5, h: 0.5, fill: { color: C.accent5 }, line: { color: C.accent5 } });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.3, y: 2.1, w: 2.4, h: 2.3, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 } });
tx(s, [{ text: "À partir du 4ᵉ mois", options: { bold: true, fontSize: 14, breakLine: true } }, { text: "100 000 FCFA", options: { bold: true, fontSize: 20, color: C.accent2, breakLine: true } }, { text: "chaque mois", options: { fontSize: 14 } }], { x: 10.4, y: 2.2, w: 2.2, h: 2.1, valign: "middle" });
tx(s, "Montants communiqués par l'organisation. Aucun autre frais n'est présenté ici.", { x: 5.0, y: 4.8, w: 7.7, h: 0.5, fontSize: 13, italic: true, color: C.accent5 });
s.addNotes("Présenter uniquement les montants communiqués. Hébergement du serveur : 100 000 FCFA par an. Maintenance : 100 000 FCFA par mois, formation des agents comprise. Le premier paiement est de 300 000 FCFA : il couvre le premier mois de maintenance et deux mois payés d'avance (mois 2 et 3). À partir du quatrième mois, le règlement est mensuel : 100 000 FCFA par mois. Ne pas calculer ni annoncer d'autres frais. Renvoyer toute question contractuelle vers l'organisation.");

// 23 — Bonnes pratiques
s = content("Bonnes pratiques et responsabilités", "Coûts et bonnes pratiques");
[["Vérifier", "le bon camion ou le bon dossier avant toute action (immatriculation, nom, identifiant)."], ["Contrôler", "les informations avant toute modification."], ["Respecter", "les permissions validées par l'organisation."], ["Signaler", "les erreurs constatées et les besoins de formation."]]
  .forEach((d, i) => { const col = i % 2, row = Math.floor(i / 2), x = 0.6 + col * 6.15, y = 1.6 + row * 2.5;
    badge(s, i + 1, x, y + 0.15, 0.7, C.accent3);
    tx(s, [{ text: d[0], options: { bold: true, fontSize: 20, color: C.text2, breakLine: true } }, { text: d[1], options: { fontSize: 15 } }], { x: x + 0.95, y, w: 4.9, h: 2.0 }); });
s.addNotes("Insister sur la responsabilité de chacun : une donnée erronée (mauvais camion, mauvais chauffeur) fausse les suivis de tous. Préciser à qui signaler une erreur ou un besoin de formation (contact à confirmer par l'organisation).");

// 24 — Quiz
pres.addSection({ title: "Conclusion" });
s = content("Vérification des acquis", "Conclusion");
["Dans quel menu retrouvez-vous un camion à partir de son immatriculation ?", "Où vérifiez-vous la date d'expiration du permis d'un chauffeur ?", "Quel statut affiche un trajet prévu mais pas encore commencé ?", "Quelles informations figurent sur une ligne de carburant ?", "Où trouvez-vous les rubriques liées au personnel ?", "Les permissions présentées sont-elles déjà configurées ?"]
  .forEach((q, i) => { const col = i % 2, row = Math.floor(i / 2), x = 0.6 + col * 6.15, y = 1.5 + row * 1.75;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 5.95, h: 1.5, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 } });
    badge(s, i + 1, x + 0.2, y + 0.45, 0.6);
    tx(s, q, { x: x + 1.0, y, w: 4.8, h: 1.5, valign: "middle" }); });
s.addNotes("RÉPONSES : 1) « Vehicle Management », recherche par nom, modèle ou immatriculation. 2) Dans « Driver Info », sur le profil du chauffeur. 3) « Upcoming ». 4) Véhicule, date, quantité, coût, chauffeur et reçu lorsqu'il est présent. 5) Dans les rubriques Employés, Présence et Paie du menu (libellés exacts à confirmer avec la capture). 6) Non : la matrice est une base de discussion, à valider par l'organisation puis à vérifier dans l'application avant diffusion.");

// 25 — Récap
s = pres.addSlide({ masterName: "TITRE", sectionTitle: "Conclusion" });
s.addText("Récapitulatif et questions", { placeholder: "title" });
tx(s, [{ text: "Se repérer dans l'interface et lire le tableau de bord", options: { bullet: true, breakLine: true } }, { text: "Retrouver camions, chauffeurs, trajets et pleins de carburant", options: { bullet: true, breakLine: true } }, { text: "Situer le personnel, les finances, l'entretien et le suivi", options: { bullet: true, breakLine: true } }, { text: "Rôles et permissions : une matrice à valider", options: { bullet: true, breakLine: true } }, { text: "Coûts : hébergement 100 000 FCFA/an, maintenance 100 000 FCFA/mois", options: { bullet: true } }],
  { x: 0.8, y: 4.3, w: 7.2, h: 2.6, fontSize: 16, color: C.background1, paraSpaceAfter: 4 });
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.5, y: 4.3, w: 4.0, h: 2.6, rectRadius: 0.08, fill: { color: C.accent1 }, line: { color: C.accent1 } });
tx(s, [{ text: "Questions à suivre", options: { bold: true, fontSize: 16, breakLine: true } }, { text: "1. ", options: { breakLine: true } }, { text: "2. ", options: { breakLine: true } }, { text: "3. " }], { x: 8.7, y: 4.4, w: 3.6, h: 2.4, fontSize: 15, paraSpaceAfter: 6 });
s.addNotes("Résumer les acquis. Recueillir les questions restées sans réponse et les noter dans l'encadré « Questions à suivre » avec la personne chargée d'y répondre. Rappeler la suite : validation de la matrice des permissions et séances complémentaires.");

// 26 — Feuille de route
pres.addSection({ title: "Après la formation" });
s = content("Feuille de route : séances complémentaires", "Après la formation");
tx(s, "Approfondissements futurs, distincts de la formation initiale — chacun avec des captures d'écran supplémentaires", { x: 0.6, y: 1.3, w: 12.1, h: 0.4, fontSize: 14, italic: true, color: C.accent5 });
["Gestion avancée des camions, groupes et itinéraires", "Performance des chauffeurs et suivi détaillé des trajets", "Gestion du personnel, présence et paie", "Finances, transactions et rapports", "Entretien, rappels, carburant et stock de pièces", "Suivi GPS, geofences et alertes", "Permissions et accompagnement des nouveaux agents"]
  .forEach((t, i) => { const col = i < 4 ? 0 : 1, row = i < 4 ? i : i - 4, x = 0.6 + col * 6.15, y = 1.95 + row * 1.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 5.95, h: 1.0, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 } });
    tx(s, "Séance " + (i + 1), { x: x + 0.2, y, w: 1.3, h: 1.0, bold: true, color: C.accent2, valign: "middle" });
    tx(s, t, { x: x + 1.5, y, w: 4.3, h: 1.0, valign: "middle" }); });
s.addNotes("Présenter ces sujets comme des approfondissements à planifier, pas comme du contenu de la formation initiale. Chaque séance nécessite les captures des écrans concernés. Dates et ordre à définir avec l'organisation selon les priorités des services.");

(async () => { await pres.writeFile({ fileName: OUT }); await applyTheme(OUT, THEME); console.log("ok"); })();
