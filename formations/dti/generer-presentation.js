// Génère la présentation de Division Terrestre Internationale (pptxgenjs).
// Usage : NODE_PATH=<node_modules contenant pptxgenjs> node generer-presentation.js <apply_theme.js>
const pptxgen = require("pptxgenjs");
const { applyTheme } = require(process.argv[2]);
const OUT = __dirname + "/presentation-division-terrestre-internationale.pptx";
const THEME = { name: "DTI", headFontFace: "Arial", bodyFontFace: "Calibri",
  colors: { dk1: "1B2430", lt1: "FFFFFF", dk2: "1E2340", lt2: "EEF1F5", accent1: "C60007", accent2: "2E6F8E",
    accent3: "4E8A5B", accent4: "E8A317", accent5: "6B7584", accent6: "C9D3DD", hlink: "2E6F8E", folHlink: "6B7584" } };
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Division Terrestre Internationale";
const C = pres.SchemeColor;

pres.defineSlideMaster({ title: "SOMBRE", background: { color: "1E2340" },
  objects: [{ placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.6, w: 11.7, h: 2.0, fontSize: 40, bold: true, color: C.background1, valign: "bottom" }, text: "" } }] });
pres.defineSlideMaster({ title: "CONTENU", background: { color: "FFFFFF" },
  slideNumber: { x: 12.3, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent5, align: "right" },
  objects: [
    { text: { text: "Division Terrestre Internationale", options: { x: 0.6, y: 7.0, w: 8, h: 0.3, fontSize: 10, color: C.accent5 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.9, fontSize: 30, bold: true, color: C.text2, valign: "middle" }, text: "" } },
  ] });

function content(title, sec) { const s = pres.addSlide({ masterName: "CONTENU", sectionTitle: sec }); s.addText(title, { placeholder: "title" }); return s; }
function tx(s, t, o) { s.addText(t, Object.assign({ isTextBox: true, fontSize: 15, color: C.text1, valign: "top" }, o)); }
function badge(s, label, x, y, d, fill, color) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill || C.accent1 }, line: { color: fill || C.accent1 } });
  tx(s, String(label), { x, y, w: d, h: d, fontSize: d > 0.6 ? 18 : 13, bold: true, color: color || C.background1, align: "center", valign: "middle", margin: 0 });
}
function card(s, x, y, w, h, head, body, fill) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill || C.background2 }, line: { color: fill || C.background2 } });
  tx(s, [{ text: head, options: { bold: true, fontSize: 16, color: C.text2, breakLine: true, paraSpaceAfter: 4 } }, { text: body, options: { fontSize: 14 } }], { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: h - 0.24 });
}
// Diapositive « module » : sous-titre, grille de fonctions, encadré bénéfice
function module(title, sec, menu, intro, feats, benefit, notes) {
  const s = content(title, sec);
  tx(s, [{ text: "Menu : ", options: { bold: true } }, { text: menu }], { x: 0.6, y: 1.25, w: 8.4, h: 0.35, fontSize: 13, color: C.accent5, margin: 0 });
  tx(s, intro, { x: 0.6, y: 1.65, w: 8.4, h: 0.6, fontSize: 16, italic: true, color: C.accent2, margin: 0 });
  const cols = 2, cw = 4.1, ch = feats.length > 4 ? 1.35 : 1.9;
  feats.forEach((f, i) => { const x = 0.6 + (i % cols) * (cw + 0.2), y = 2.45 + Math.floor(i / cols) * (ch + 0.2); card(s, x, y, cw, ch, f[0], f[1]); });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 9.35, y: 1.3, w: 3.4, h: 5.4, rectRadius: 0.08, fill: { color: C.text2 }, line: { color: C.text2 } });
  tx(s, [{ text: "Ce que ça apporte", options: { bold: true, fontSize: 17, color: C.accent4, breakLine: true, paraSpaceAfter: 10 } },
    ...benefit.map((b, i) => ({ text: b, options: { bullet: true, breakLine: i < benefit.length - 1, paraSpaceAfter: 8 } }))],
    { x: 9.55, y: 1.5, w: 3.0, h: 5.0, fontSize: 14, color: C.background1 });
  s.addNotes(notes);
  return s;
}

// 1 — Titre
pres.addSection({ title: "Introduction" });
let s = pres.addSlide({ masterName: "SOMBRE", sectionTitle: "Introduction" });
s.addText("Division Terrestre Internationale", { placeholder: "title" });
tx(s, "Le logiciel de gestion de la flotte, des chauffeurs, des trajets et des finances", { x: 0.8, y: 3.8, w: 11.7, h: 0.6, fontSize: 20, color: C.accent4 });
tx(s, [{ text: "Présentation : [à compléter]", options: { breakLine: true } }, { text: "Date : [à compléter]" }], { x: 0.8, y: 5.2, w: 6, h: 0.9, fontSize: 15, color: C.background1 });
s.addNotes("Accueillir les participants (agents et direction). Annoncer le déroulé : présentation du logiciel (ce qu'il sait faire, ses modules, comment et pourquoi l'utiliser), puis démonstration en direct d'un petit circuit de travail, puis questions.");

// 2 — Pourquoi
s = content("Pourquoi ce logiciel ?", "Introduction");
tx(s, "Aujourd'hui", { x: 0.6, y: 1.4, w: 5.6, h: 0.45, fontSize: 20, bold: true, color: C.accent1 });
[["Informations dispersées", "Cahiers, tableurs, appels : difficile de savoir où est un camion et qui le conduit."],
 ["Échéances oubliées", "Assurances, immatriculations et permis expirés découverts trop tard."],
 ["Coûts mal suivis", "Carburant, entretien et pièces difficiles à rattacher à chaque véhicule."]]
  .forEach((d, i) => card(s, 0.6, 2.0 + i * 1.55, 5.6, 1.35, d[0], d[1]));
s.addShape(pres.shapes.RIGHT_ARROW, { x: 6.4, y: 3.7, w: 0.6, h: 0.6, fill: { color: C.accent5 }, line: { color: C.accent5 } });
tx(s, "Avec Division Terrestre Internationale", { x: 7.2, y: 1.4, w: 5.6, h: 0.45, fontSize: 20, bold: true, color: C.accent3 });
[["Tout au même endroit", "Véhicules, chauffeurs, trajets, clients, carburant, entretien et finances."],
 ["Alertes et rappels", "Documents expirés signalés, rappels d'entretien, tableau de bord du jour."],
 ["Décisions chiffrées", "Revenus, dépenses et rapports par véhicule, chauffeur ou itinéraire."]]
  .forEach((d, i) => card(s, 7.2, 2.0 + i * 1.55, 5.55, 1.35, d[0], d[1]));
s.addNotes("Partir du quotidien des participants : demander comment ils savent aujourd'hui quel camion est disponible, quand expire une assurance, ou combien a coûté un véhicule le mois dernier. Le logiciel répond à ces trois besoins : centraliser, alerter, chiffrer.");

// 3 — Modules
s = content("18 modules, 6 grandes familles", "Introduction");
s.addShape(pres.shapes.OVAL, { x: 0.6, y: 2.4, w: 2.6, h: 2.6, fill: { color: C.accent1 }, line: { color: C.accent1 } });
tx(s, [{ text: "18", options: { fontSize: 60, bold: true, breakLine: true } }, { text: "modules", options: { fontSize: 18 } }], { x: 0.6, y: 2.4, w: 2.6, h: 2.6, align: "center", valign: "middle", color: C.background1 });
[["Pilotage", "Dashboard · Availability · Reports"], ["Flotte", "Vehicle's · Vehicle Vendor's · Tracking · Geofence"], ["Personnes", "Driver's · Employee"],
 ["Activité commerciale", "Trips · Customer · Coupon Management"], ["Exploitation", "Maintenance · Parts Stock · Fuel · Reminder"], ["Finances et système", "Accounts · Setting's"]]
  .forEach((d, i) => { const x = 3.7 + (i % 2) * 4.55, y = 1.45 + Math.floor(i / 2) * 1.8; card(s, x, y, 4.35, 1.6, d[0], d[1]); });
s.addNotes("Présenter la carte du logiciel : 18 modules dans le menu latéral, regroupés ici en 6 familles pour faciliter la lecture. Les libellés restent en anglais dans l'application : on les garde tels quels pour que chacun s'y retrouve à l'écran. Chaque famille est détaillée dans les diapositives suivantes.");

// 4 — Tableau de bord
pres.addSection({ title: "Pilotage" });
module("Le tableau de bord : la journée en un coup d'œil", "Pilotage", "Dashboard · Availability",
  "Dès la connexion, l'essentiel de l'activité est affiché.",
  [["Compteurs", "Véhicules, chauffeurs, clients et trajets du jour."], ["Transactions", "Courbe des revenus et des dépenses, et totaux du jour."],
   ["Statut de la flotte", "Véhicules disponibles ou en trajet, position actuelle."], ["Rappels", "Entretiens à prévoir, avec ajout rapide d'un rappel."],
   ["Top revenus et dépenses", "Graphiques des principaux postes."], ["Calendrier de disponibilité", "Occupation des véhicules par mois, semaine ou jour."]],
  ["La direction voit l'activité sans demander de rapport", "Les agents savent tout de suite quels véhicules sont libres", "Les urgences (rappels) sont visibles dès l'ouverture"],
  "Montrer le tableau de bord en direct pendant la démonstration. Insister sur le calendrier Availability : c'est le premier réflexe avant de planifier un trajet.");

// 5 — Véhicules
pres.addSection({ title: "Flotte" });
module("Gérer les véhicules", "Flotte", "Vehicle's › Vehicle List · Add Vehicle · Vehicle Group · Route",
  "Une fiche complète pour chaque véhicule, de l'immatriculation au GPS.",
  [["Fiche véhicule", "Immatriculation, nom, modèle, châssis, moteur, constructeur, compteur, photo et documents."], ["Échéances", "Dates d'expiration de l'immatriculation et de l'assurance, signalées quand elles sont dépassées."],
   ["Types et itinéraires", "Types de véhicules (camion, bus, voiture…) et liste des itinéraires."], ["Tarif et propriété", "Type de facturation et coût par défaut, véhicule en propre ou loué, liaison GPS."]],
  ["Plus d'assurance expirée qui passe inaperçue", "Saisie rapide de toute la flotte par import CSV", "Export de la liste en Excel, CSV ou PDF"],
  "Points à montrer : la colonne des dates d'expiration en rouge, l'import en masse (Bulk Import), et les boutons d'export présents sur toutes les listes du logiciel.");

// 6 — Chauffeurs et fournisseurs
module("Chauffeurs et fournisseurs de véhicules", "Flotte", "Driver's · Vehicle Vendor's",
  "Savoir qui conduit, avec quel permis, et avec quels partenaires on travaille.",
  [["Fiche chauffeur", "Contact, âge, adresse, expérience, date d'entrée, photo et document."], ["Permis", "Numéro et date d'expiration, signalée quand elle est dépassée."],
   ["Fournisseurs de véhicules", "Société, contact, téléphone, adresse et statut."], ["Contrats", "Date de contrat et document du contrat joint."]],
  ["Un chauffeur au permis expiré est repéré tout de suite", "Les contrats des loueurs sont centralisés", "Import des chauffeurs en masse par CSV"],
  "Faire le lien avec la sécurité et la conformité : un permis expiré est affiché en rouge dans la liste. Les fournisseurs de véhicules (Vehicle Vendor's) concernent les véhicules loués à des partenaires.");

// 7 — GPS et zones
module("Suivi GPS et zones géographiques", "Flotte", "Tracking › History Tracking · Live Location · Geofence",
  "Savoir où sont les véhicules, et être prévenu quand ils entrent ou sortent d'une zone.",
  [["Position en direct", "Localisation actuelle des véhicules sur la carte."], ["Historique des parcours", "Trajet d'un véhicule sur une période choisie."],
   ["Zones (geofences)", "Zones dessinées sur la carte : cercle, rectangle, polygone…"], ["Événements de zone", "Entrées et sorties enregistrées, avec notification des personnes choisies."]],
  ["Contrôler les itinéraires réellement suivis", "Réagir vite en cas d'écart", "Rassurer les clients sur leurs livraisons"],
  "Préciser que le suivi GPS nécessite un boîtier GPS relié à chaque véhicule (synchronisation Traccar, visible sur le tableau de bord). Montrer la carte de l'historique si un véhicule est équipé.");

// 8 — Trajets et clients
pres.addSection({ title: "Activité commerciale" });
s = module("Trajets, clients et coupons", "Activité commerciale", "Trips · Customer · Coupon Management",
  "De la réservation à la facturation, chaque trajet est suivi.",
  [["Réservation", "Client, dates, véhicule, chauffeur, départ, arrivée, kilométrage, points de ramassage."], ["Suivi du trajet", "Statut (réservé, commencé…), distance et numéro de réservation."],
   ["Clients", "Coordonnées, solde impayé et statut actif ou inactif."], ["Coupons", "Codes de réduction : montant, dates, limite d'utilisation."]],
  ["Chaque trajet est rattaché à un client, un véhicule et un chauffeur", "Les impayés clients sont visibles", "Confirmation de réservation par e-mail ou SMS"],
  "Insister sur la souplesse de facturation (diapositive suivante en encart). La réservation peut être envoyée au client par e-mail ou SMS. Ce module servira de point de départ à la démonstration.");
tx(s, [{ text: "9 modes de facturation : ", options: { bold: true, color: C.accent1 } }, { text: "fixe · par tonne · kg · km · trajet · jour · heure · litre · sac" }], { x: 0.6, y: 6.55, w: 8.6, h: 0.35, fontSize: 13, margin: 0 });

// 9 — Entretien
pres.addSection({ title: "Exploitation" });
module("Entretien, pièces et rappels", "Exploitation", "Maintenance · Parts Stock · Reminder",
  "Anticiper les entretiens et maîtriser leur coût.",
  [["Entretiens", "Véhicule, dates, mécanicien, fournisseur, compteur, coût, détail et pièces utilisées."], ["Mécaniciens et fournisseurs", "Annuaire des intervenants, avec catégorie."],
   ["Stock de pièces", "Inventaire, prix, stock de départ et historique des achats avec statut de paiement."], ["Rappels", "Échéances par véhicule et service, marquées comme terminées une fois faites."]],
  ["Moins de pannes grâce aux rappels", "Coût réel de l'entretien par véhicule", "Stock de pièces suivi, achats tracés"],
  "Montrer qu'un entretien consomme des pièces du stock et peut notifier des membres de l'équipe. Les rappels apparaissent aussi sur le tableau de bord.");

// 10 — Carburant
module("Carburant", "Exploitation", "Fuel › Fuel Management · Add Fuel · Fuel Vendor",
  "Chaque plein est enregistré et rattaché à un véhicule et un chauffeur.",
  [["Pleins", "Date, véhicule, chauffeur, quantité, montant, compteur et commentaire."], ["Provenance", "Cuve interne (avec son stock en litres) ou fournisseur."],
   ["Cuve", "Réapprovisionnement de la cuve et suivi du stock disponible."], ["Fournisseurs", "Liste des stations ou fournisseurs de carburant."]],
  ["Repérer une consommation anormale", "Contrôler le stock de la cuve", "Rapport carburant par période"],
  "Le compteur kilométrique saisi à chaque plein permet de comparer consommation et distance. Le rapport Fuel du module Reports reprend ces données.");

// 11 — Finances
pres.addSection({ title: "Finances et pilotage" });
module("Finances", "Finances et pilotage", "Accounts › Accounts · Add Transactions · Transactions · Ac Transfers / Deposit · Transactions Category",
  "Revenus et dépenses enregistrés par compte, catégorie et véhicule.",
  [["Comptes", "Comptes financiers avec leur solde."], ["Transactions", "Revenu ou dépense : compte, catégorie, véhicule, montant, référence, note."],
   ["Virements et dépôts", "Mouvements d'un compte à un autre."], ["Catégories", "Ex. revenus de trajet, carburant, salaires, téléphone, transfrontalier…"]],
  ["Rentabilité visible par véhicule", "Historique filtrable par période, avec totaux", "Catégories adaptées à l'entreprise"],
  "Montrer la liste des transactions filtrée par période, avec le total des débits et des crédits. Les catégories sont personnalisables : à définir avec la direction financière.");

// 12 — Rapports
s = content("11 rapports prêts à l'emploi", "Finances et pilotage");
tx(s, "Choisir une période, un filtre, puis « Generate Report ».", { x: 0.6, y: 1.3, w: 12, h: 0.45, fontSize: 16, italic: true, color: C.accent2 });
["Réservations", "Revenus et dépenses", "Carburant", "Chauffeurs", "Coupons", "Rappels", "Entretien", "Top revenus", "Top dépenses", "Top véhicules", "Top itinéraires"]
  .forEach((r, i) => { const col = i % 4, row = Math.floor(i / 4), x = 0.6 + col * 3.1, y = 2.0 + row * 1.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.9, h: 1.25, rectRadius: 0.08, fill: { color: i >= 7 ? C.text2 : C.background2 }, line: { color: i >= 7 ? C.text2 : C.background2 } });
    badge(s, i + 1, x + 0.2, y + 0.33, 0.6, i >= 7 ? C.accent4 : C.accent1, i >= 7 ? C.text1 : C.background1);
    tx(s, r, { x: x + 0.95, y, w: 1.85, h: 1.25, fontSize: 15, bold: true, valign: "middle", color: i >= 7 ? C.background1 : C.text2 }); });
tx(s, "Les « Top » classent les postes qui rapportent ou coûtent le plus.", { x: 0.6, y: 6.5, w: 12, h: 0.4, fontSize: 13, color: C.accent5 });
s.addNotes("Pour la direction : les rapports Top véhicules et Top itinéraires montrent où l'entreprise gagne et perd de l'argent. Pour les agents : les rapports remplacent les synthèses faites à la main.");

// 13 — Employés et droits
pres.addSection({ title: "Administration" });
module("Employés et droits d'accès", "Administration", "Employee › Employee Management · Add Employee",
  "Chaque utilisateur a son compte et ne voit que ce qu'il doit voir.",
  [["Comptes utilisateurs", "Nom, téléphone, e-mail, identifiant, mot de passe, dernière connexion."], ["Permissions par module", "Cases à cocher : voir la liste, ajouter, modifier, supprimer…"],
   ["Exemples", "Véhicules, chauffeurs, transactions, clients, carburant, rappels, entretien, pièces, suivi GPS…"], ["Statut", "Compte actif ou désactivé."]],
  ["Chacun travaille dans son périmètre", "Les données sensibles (finances) restent protégées", "On sait qui a créé chaque transaction"],
  "Les droits se règlent finement, module par module et action par action. Les profils (exploitation, finance, direction…) seront définis avec l'organisation avant la mise en service.");

// 14 — Paramètres
module("Paramètres et sécurité", "Administration", "Setting's",
  "Le logiciel s'adapte à l'entreprise et protège ses données.",
  [["Identité de l'entreprise", "Nom, adresse, téléphone, e-mail, logo, couleurs, fuseau horaire."], ["Communication", "Configuration de l'envoi d'e-mails et de SMS, avec modèles de messages."],
   ["Réservations et factures", "Préfixe des numéros de réservation, réglages des factures."], ["Sauvegardes", "Sauvegarde de la base de données en un clic (DB Backup)."]],
  ["Documents aux couleurs de l'entreprise", "Clients informés automatiquement", "Données protégées par des sauvegardes"],
  "Les paramètres sont réservés à l'administrateur. Mentionner aussi la configuration de la carte (MAP API) et du GPS (Traccar Config).");

// 15 — Ce que ça change
pres.addSection({ title: "Conclusion" });
s = content("Ce que ça change pour chacun", "Conclusion");
[["Pour la direction", C.accent1, ["Vision en temps réel de l'activité", "Rentabilité par véhicule, chauffeur et itinéraire", "Coûts de carburant et d'entretien maîtrisés", "Décisions appuyées sur des rapports"]],
 ["Pour les agents", C.accent2, ["Une seule application au lieu de plusieurs outils", "Informations retrouvées en quelques secondes", "Rappels automatiques des échéances", "Moins de saisie grâce aux imports et exports"]]]
  .forEach((d, i) => { const x = 0.6 + i * 6.2;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.5, w: 5.9, h: 5.1, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 } });
    badge(s, i === 0 ? "D" : "A", x + 0.3, y = 1.75, 0.8, d[1]);
    tx(s, d[0], { x: x + 1.3, y: 1.75, w: 4.4, h: 0.8, fontSize: 22, bold: true, color: C.text2, valign: "middle" });
    tx(s, d[2].map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < 3, paraSpaceAfter: 10 } })), { x: x + 0.3, y: 2.9, w: 5.3, h: 3.5, fontSize: 17 }); });
s.addNotes("Synthèse orientée bénéfices, avant la démonstration. Laisser la direction et les agents réagir : quel bénéfice compte le plus pour eux ?");

// 16 — Démo
s = content("Démonstration : un circuit de travail", "Conclusion");
tx(s, "En direct dans l'application", { x: 0.6, y: 1.3, w: 12, h: 0.45, fontSize: 16, italic: true, color: C.accent2 });
const steps = [["Client", "Customer › Add Customer"], ["Trajet", "Trips › Add Trips"], ["Carburant", "Fuel › Add Fuel"], ["Transaction", "Accounts › Add Transactions"], ["Tableau de bord et rapport", "Dashboard · Reports"]];
steps.forEach((d, i) => { const x = 0.6 + i * 2.5;
  s.addShape(pres.shapes.OVAL, { x: x + 0.55, y: 2.2, w: 1.2, h: 1.2, fill: { color: C.accent1 }, line: { color: C.accent1 } });
  tx(s, String(i + 1), { x: x + 0.55, y: 2.2, w: 1.2, h: 1.2, fontSize: 28, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0 });
  if (i < steps.length - 1) s.addShape(pres.shapes.RIGHT_ARROW, { x: x + 1.95, y: 2.6, w: 0.45, h: 0.4, fill: { color: C.accent5 }, line: { color: C.accent5 } });
  tx(s, [{ text: d[0], options: { bold: true, fontSize: 17, color: C.text2, breakLine: true } }, { text: d[1], options: { fontSize: 13, color: C.accent5 } }], { x, y: 3.6, w: 2.3, h: 1.2, align: "center" }); });
card(s, 0.6, 5.2, 12.1, 1.4, "Objectif", "Montrer qu'une seule saisie alimente tout le reste : le trajet apparaît dans le calendrier, le carburant et la transaction dans les finances, puis dans le tableau de bord et les rapports.");
s.addNotes("DÉROULÉ : 1) créer un client fictif ; 2) créer un trajet pour ce client avec un véhicule et un chauffeur, choisir un mode de facturation ; 3) enregistrer un plein pour ce véhicule ; 4) enregistrer la transaction de revenu du trajet (catégorie Trip Income) ; 5) revenir au tableau de bord et générer un rapport sur la période. Utiliser uniquement des données de test, et les supprimer après la démonstration.");

// 17 — Questions
s = pres.addSlide({ masterName: "SOMBRE", sectionTitle: "Conclusion" });
s.addText("Questions et échanges", { placeholder: "title" });
tx(s, "Merci de votre attention", { x: 0.8, y: 3.8, w: 11.7, h: 0.6, fontSize: 20, color: C.accent4 });
s.addNotes("Recueillir les questions et les besoins : modules prioritaires, profils de droits à créer, données à importer (véhicules, chauffeurs), formation des agents.");

(async () => { await pres.writeFile({ fileName: OUT }); await applyTheme(OUT, THEME); console.log("ok"); })();
