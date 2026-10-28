
const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const initSqlJs = require('sql.js');
const XLSX = require('xlsx');

// Le dossier des données reste le même, quel que soit le nom affiché de l'application.
app.setPath('userData', path.join(app.getPath('appData'), 'karate-club-manager'));

// ── AUTH BEGIN ──────────────────────────────────────────────
// AUTHENTIFICATION, RÔLES ET JOURNAL DES ACTIONS
// Toutes les vérifications se font ICI (processus principal) : l'interface ne peut pas les contourner.
const crypto = require('crypto');
let session = null;
let essaisRecup = 0;

const ROLES = ['president', 'tresorier', 'secretaire', 'coach'];
const hacher = (mdp, sel) => crypto.scryptSync(String(mdp), sel, 64).toString('hex');
const nouveauSel = () => crypto.randomBytes(16).toString('hex');
const egaux = (a, b) => { const x = Buffer.from(a, 'hex'), y = Buffer.from(b, 'hex'); return x.length === y.length && crypto.timingSafeEqual(x, y); };
const utilisateurPublic = u => ({ id: u.id, nom: u.nom, identifiant: u.identifiant, role: u.role, actif: u.actif, doit_changer: u.doit_changer });
function cleRecuperation() {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', b = crypto.randomBytes(16);
  let s = ''; for (let i = 0; i < 16; i++) s += alphabet[b[i] % alphabet.length];
  return s.match(/.{4}/g).join('-');
}
function enregistrerCle(cle) {
  const sel = nouveauSel();
  db.run("INSERT OR REPLACE INTO securite(cle, valeur) VALUES('recup_sel', ?)", [sel]);
  db.run("INSERT OR REPLACE INTO securite(cle, valeur) VALUES('recup_hash', ?)", [hacher(cle.replace(/-/g, ''), sel)]);
}
function ecrireJournal(utilisateur, role, action, detail) {
  db.run('INSERT INTO journal(horodatage, utilisateur, role, action, detail) VALUES(?,?,?,?,?)',
    [new Date().toISOString(), utilisateur, role, action, String(detail || '').slice(0, 400)]);
}
function verifierMotDePasse(mdp) {
  if (typeof mdp !== 'string' || mdp.length < 8) throw new Error('Mot de passe trop court (8 caractères minimum).');
}

// ── Droits par rôle (le président peut tout) ──
const G_ATH = ['ath:save', 'ath:groupes', 'ath:groupe', 'doc:add', 'xl:import', 'dos:set'];
const G_GRP = ['grp:save', 'grp:del'];
const G_SEA = ['pres:save', 'pc:save', 'ann:set'];
const G_FIN = ['pay:cancel', 'ins:set', 'ins:mode', 'ins:delai', 'ins:assur', 'saison:ouvrir'];
const LECTURE = ['ath:list', 'doc:list', 'doc:open', 'row:list', 'pres:get', 'pres:stats', 'param:get', 'xl:model'];
const DROITS = {
  tresorier: new Set([...LECTURE, ...G_FIN, 'list:export', 'xl:export', 'lien:ouvrir', 'row:save', 'row:del']),
  secretaire: new Set([...LECTURE, ...G_ATH, ...G_GRP, ...G_SEA, 'ins:set', 'ins:assur', 'list:export', 'xl:export', 'lien:ouvrir', 'row:save', 'row:del']),
  coach: new Set([...LECTURE, ...G_SEA])
};
const TABLES_ECRITURE = {
  tresorier: ['paiements', 'relances', 'saisons'],
  secretaire: ['grades', 'competitions', 'resultats', 'coachs', 'groupes', 'creneaux', 'annulations', 'relances'],
  coach: []
};
const TABLES_LECTURE_INTERDITE = { coach: ['paiements', 'relances'] };

function autoriser(canal, args) {
  if (!session) return { ok: false, erreur: 'Non connecté.' };
  const r = session.role;
  if (r === 'president') return { ok: true };
  const refus = { ok: false, erreur: 'Accès refusé : droits insuffisants.' };
  if (!DROITS[r] || !DROITS[r].has(canal)) return refus;
  if ((canal === 'row:save' || canal === 'row:del') && !(TABLES_ECRITURE[r] || []).includes(args[0])) return refus;
  if (canal === 'row:list' && (TABLES_LECTURE_INTERDITE[r] || []).includes(args[0])) return { ok: false, vide: [] };
  if (canal === 'doc:list' && r === 'coach') return { ok: false, vide: [] };
  if (canal === 'doc:open' && r === 'coach') return refus;
  if (canal === 'ins:set' && r === 'secretaire' && args[3] != null) return refus; // pas de remise
  return { ok: true };
}

const LIBELLES = {
  'ath:save': 'Athlète enregistré', 'ath:del': 'Athlète supprimé', 'ath:groupes': 'Affectation groupée', 'ath:groupe': 'Changement de groupe',
  'doc:add': 'Document ajouté', 'xl:import': 'Import Excel', 'xl:export': 'Export Excel des athlètes', 'dos:set': 'Dossier d’inscription',
  'grp:save': 'Groupe enregistré', 'grp:del': 'Groupe supprimé', 'pres:save': 'Appel enregistré', 'pc:save': 'Présence des coachs',
  'ann:set': 'Séance annulée', 'pay:cancel': 'Paiement annulé', 'ins:set': 'Inscription modifiée', 'ins:mode': 'Formule de paiement',
  'ins:delai': 'Accord de paiement', 'ins:assur': 'Assurance transmise', 'saison:ouvrir': 'Nouvelle saison', 'row:save': 'Enregistrement',
  'row:del': 'Suppression', 'param:set': 'Paramètre modifié', 'sauv:export': 'Sauvegarde', 'sauv:restore': 'Restauration',
  'list:export': 'Export de liste', 'lien:ouvrir': 'Message WhatsApp', 'usr:save': 'Utilisateur modifié', 'usr:reset': 'Mot de passe réinitialisé'
};
function journaliser(canal, args, res) {
  if (!session || !LIBELLES[canal]) return;
  let detail;
  if (canal === 'usr:save') detail = `${args[0] && (args[0].identifiant || '#' + args[0].id)} (${args[0] && args[0].role})`;
  else if (canal === 'usr:reset') detail = `utilisateur n°${args[0]}`;
  else if (canal === 'row:save' || canal === 'row:del') detail = `${args[0]} ${JSON.stringify(args[1] ?? '').slice(0, 280)}`;
  else if (canal === 'list:export') detail = `${args[0]} (${(args[2] || []).length} lignes)`;
  else if (canal === 'lien:ouvrir') detail = 'lien WhatsApp ouvert';
  else detail = JSON.stringify(args).slice(0, 300);
  try { ecrireJournal(session.nom, session.role, LIBELLES[canal], detail); persist(); } catch (_) { /* le journal ne doit jamais bloquer l'action */ }
}

// Tous les canaux enregistrés ensuite passent par ce contrôle central.
const handleBrut = ipcMain.handle.bind(ipcMain);
ipcMain.handle = (canal, fn) => handleBrut(canal, async (e, ...args) => {
  const v = autoriser(canal, args);
  if (!v.ok) { if ('vide' in v) return v.vide; throw new Error(v.erreur); }
  const res = await fn(e, ...args);
  journaliser(canal, args, res);
  return res;
});

// ── Canaux d'authentification (accessibles sans session) ──
handleBrut('auth:etat', () => {
  const n = q('SELECT COUNT(*) AS n FROM utilisateurs')[0].n;
  const u = session ? q('SELECT * FROM utilisateurs WHERE id=?', [session.id])[0] : null;
  return { aucun: n === 0, connecte: !!u, utilisateur: u ? utilisateurPublic(u) : null };
});

handleBrut('auth:setup', (e, nom, identifiant, mdp) => {
  if (q('SELECT COUNT(*) AS n FROM utilisateurs')[0].n > 0) throw new Error('La configuration initiale est déjà faite.');
  nom = String(nom || '').trim(); identifiant = String(identifiant || '').trim().toLowerCase();
  if (!nom || !/^[a-z0-9._-]{3,30}$/.test(identifiant)) throw new Error('Nom obligatoire et identifiant de 3 à 30 caractères (lettres minuscules, chiffres, . _ -).');
  verifierMotDePasse(mdp);
  const sel = nouveauSel();
  db.run("INSERT INTO utilisateurs(nom, identifiant, role, sel, hash, actif, doit_changer, cree_le) VALUES(?,?, 'president', ?, ?, 1, 0, ?)",
    [nom, identifiant, sel, hacher(mdp, sel), new Date().toISOString()]);
  const cle = cleRecuperation(); enregistrerCle(cle);
  const u = q('SELECT * FROM utilisateurs WHERE identifiant=?', [identifiant])[0];
  session = { id: u.id, nom: u.nom, role: u.role };
  ecrireJournal(u.nom, u.role, 'Création du compte président', identifiant);
  persist();
  return { cle, utilisateur: utilisateurPublic(u) };
});

handleBrut('auth:login', (e, identifiant, mdp) => {
  const maintenant = Date.now();
  const u = q('SELECT * FROM utilisateurs WHERE identifiant=?', [String(identifiant || '').trim().toLowerCase()])[0];
  if (u && u.verrou_jusqua > maintenant) return { ok: false, erreur: `Compte verrouillé. Réessayez dans ${Math.ceil((u.verrou_jusqua - maintenant) / 60000)} min.` };
  if (!u || !u.actif || !egaux(hacher(mdp, u.sel), u.hash)) {
    if (u) {
      const n = u.echecs + 1;
      db.run('UPDATE utilisateurs SET echecs=?, verrou_jusqua=? WHERE id=?', [n >= 5 ? 0 : n, n >= 5 ? maintenant + 5 * 60000 : 0, u.id]);
    }
    ecrireJournal(u ? u.nom : String(identifiant || '').slice(0, 40), u ? u.role : '', 'Échec de connexion', '');
    persist();
    return { ok: false, erreur: 'Identifiant ou mot de passe incorrect.' };
  }
  db.run('UPDATE utilisateurs SET echecs=0, verrou_jusqua=0 WHERE id=?', [u.id]);
  session = { id: u.id, nom: u.nom, role: u.role };
  ecrireJournal(u.nom, u.role, 'Connexion', '');
  persist();
  return { ok: true, utilisateur: utilisateurPublic(u) };
});

handleBrut('auth:logout', () => {
  if (session) { ecrireJournal(session.nom, session.role, 'Déconnexion', ''); persist(); }
  session = null;
  return true;
});

handleBrut('auth:changer', (e, ancien, nouveau) => {
  if (!session) throw new Error('Non connecté.');
  const u = q('SELECT * FROM utilisateurs WHERE id=?', [session.id])[0];
  if (!egaux(hacher(ancien, u.sel), u.hash)) throw new Error('Le mot de passe actuel est incorrect.');
  verifierMotDePasse(nouveau);
  if (nouveau === ancien) throw new Error('Le nouveau mot de passe doit être différent.');
  const sel = nouveauSel();
  db.run('UPDATE utilisateurs SET sel=?, hash=?, doit_changer=0 WHERE id=?', [sel, hacher(nouveau, sel), u.id]);
  ecrireJournal(u.nom, u.role, 'Mot de passe modifié', '');
  persist();
  return true;
});

handleBrut('auth:recuperer', (e, cle, nouveau) => {
  if (essaisRecup >= 5) return { ok: false, erreur: 'Trop de tentatives. Fermez puis relancez l’application.' };
  const sel = (q("SELECT valeur FROM securite WHERE cle='recup_sel'")[0] || {}).valeur;
  const hash = (q("SELECT valeur FROM securite WHERE cle='recup_hash'")[0] || {}).valeur;
  const saisie = String(cle || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  if (!sel || !hash || !egaux(hacher(saisie, sel), hash)) { essaisRecup++; return { ok: false, erreur: 'Clé de récupération incorrecte.' }; }
  verifierMotDePasse(nouveau);
  const p = q("SELECT * FROM utilisateurs WHERE role='president' AND actif=1 ORDER BY id LIMIT 1")[0];
  if (!p) return { ok: false, erreur: 'Aucun compte président actif.' };
  const s2 = nouveauSel();
  db.run('UPDATE utilisateurs SET sel=?, hash=?, doit_changer=1, echecs=0, verrou_jusqua=0 WHERE id=?', [s2, hacher(nouveau, s2), p.id]);
  const nouvelle = cleRecuperation(); enregistrerCle(nouvelle);
  ecrireJournal(p.nom, p.role, 'Récupération du compte président par clé', '');
  essaisRecup = 0;
  persist();
  return { ok: true, identifiant: p.identifiant, cle: nouvelle };
});

// ── Gestion des utilisateurs et journal (président uniquement : aucun autre rôle n'a ces canaux) ──
ipcMain.handle('usr:list', () => q('SELECT id, nom, identifiant, role, actif, doit_changer FROM utilisateurs ORDER BY id'));

ipcMain.handle('usr:save', (e, u) => {
  if (!u || !ROLES.includes(u.role)) throw new Error('Rôle invalide.');
  const nom = String(u.nom || '').trim();
  if (!nom) throw new Error('Le nom est obligatoire.');
  if (u.id) {
    const cur = q('SELECT * FROM utilisateurs WHERE id=?', [u.id])[0];
    if (!cur) throw new Error('Utilisateur introuvable.');
    if (session && cur.id === session.id && !u.actif) throw new Error('Vous ne pouvez pas désactiver votre propre compte.');
    db.run('UPDATE utilisateurs SET nom=?, role=?, actif=? WHERE id=?', [nom, u.role, u.actif ? 1 : 0, u.id]);
    if (q("SELECT COUNT(*) AS n FROM utilisateurs WHERE role='president' AND actif=1")[0].n < 1) {
      db.run('UPDATE utilisateurs SET nom=?, role=?, actif=? WHERE id=?', [cur.nom, cur.role, cur.actif, cur.id]);
      throw new Error('Il doit rester au moins un président actif.');
    }
  } else {
    const identifiant = String(u.identifiant || '').trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,30}$/.test(identifiant)) throw new Error('Identifiant : 3 à 30 caractères (lettres minuscules, chiffres, . _ -).');
    if (q('SELECT 1 FROM utilisateurs WHERE identifiant=?', [identifiant]).length) throw new Error('Cet identifiant existe déjà.');
    verifierMotDePasse(u.motdepasse);
    const sel = nouveauSel();
    db.run('INSERT INTO utilisateurs(nom, identifiant, role, sel, hash, actif, doit_changer, cree_le) VALUES(?,?,?,?,?,1,1,?)',
      [nom, identifiant, u.role, sel, hacher(u.motdepasse, sel), new Date().toISOString()]);
  }
  persist();
  return true;
});

ipcMain.handle('usr:reset', (e, id, mdp) => {
  verifierMotDePasse(mdp);
  if (!q('SELECT 1 FROM utilisateurs WHERE id=?', [id]).length) throw new Error('Utilisateur introuvable.');
  const sel = nouveauSel();
  db.run('UPDATE utilisateurs SET sel=?, hash=?, doit_changer=1, echecs=0, verrou_jusqua=0 WHERE id=?', [sel, hacher(mdp, sel), id]);
  persist();
  return true;
});

ipcMain.handle('jrn:list', (e, f = {}) => {
  const cond = [], p = [];
  if (f.du) { cond.push('horodatage >= ?'); p.push(new Date(f.du + 'T00:00:00').toISOString()); }
  if (f.au) { cond.push('horodatage <= ?'); p.push(new Date(f.au + 'T23:59:59').toISOString()); }
  if (f.utilisateur) { cond.push('utilisateur LIKE ?'); p.push('%' + f.utilisateur + '%'); }
  if (f.texte) { cond.push('(action LIKE ? OR detail LIKE ?)'); p.push('%' + f.texte + '%', '%' + f.texte + '%'); }
  return q(`SELECT * FROM journal ${cond.length ? 'WHERE ' + cond.join(' AND ') : ''} ORDER BY id DESC LIMIT 1000`, p);
});
// ── AUTH END ────────────────────────────────────────────────


let db;
let dbFile;
let mainWindow;

const dir = () =>
  process.env.PORTABLE_EXECUTABLE_DIR || app.getPath('userData');

const PARAMETRES_DEFAUT = {
  nom_club: 'Club Sportif des Arts Martiaux – Béjaïa',
  adresse: '',
  telephone1: '0770 18 84 95',
  telephone2: '0550 87 59 70',
  president: 'TALBI Mohammed',
  federation: '',
  ligue: '',
  reglement: '',
  logo: 'logo.png',
  whatsapp_numero: '',
  whatsapp_groupe: ''
};

process.on('uncaughtException', e => {
  try {
    dialog.showErrorBox(
      'Erreur ASAMB',
      String(e.stack || e)
    );
  } catch (_) {
    console.error(e);
  }
});

// ─────────────────────────────────────────────
// BASE DE DONNÉES
// Ne jamais modifier les migrations historiques.
// Toute évolution doit être ajoutée à la fin.
// ─────────────────────────────────────────────

const COLS0 = [
  'nom', 'prenom', 'nom_ar', 'prenom_ar',
  'date_naissance', 'lieu_naissance', 'lieu_naissance_ar',
  'sexe', 'pere_ar', 'mere_ar', 'telephone', 'adresse',
  'ceinture', 'groupe', 'licence_no', 'licence_expiration',
  'certificat_expiration', 'paye_jusqua'
];

const COLS = COLS0
  .filter(c => !c.endsWith('_ar') && c !== 'paye_jusqua')
  .concat([
    'date_integration',
    'groupe_sanguin',
    'taille',
    'poids',
    'profession',
    'club_precedent',
    'profession_pere'
  ]);

const MIGRATIONS = [
  // Migration 0
  `CREATE TABLE athletes(
    id INTEGER PRIMARY KEY,
    ${COLS0.map(c => c + ' TEXT').join(', ')}
  );
  CREATE TABLE documents(
    id INTEGER PRIMARY KEY,
    athlete_id INTEGER,
    type TEXT,
    fichier TEXT,
    date_ajout TEXT DEFAULT CURRENT_DATE
  );`,

  // Migration 1
  `CREATE TABLE coachs(
    id INTEGER PRIMARY KEY,
    nom TEXT,
    telephone TEXT,
    diplome TEXT
  );
  CREATE TABLE groupes(
    id INTEGER PRIMARY KEY,
    nom TEXT,
    categorie TEXT,
    coach_id INTEGER,
    horaire TEXT
  );
  CREATE TABLE presences(
    id INTEGER PRIMARY KEY,
    athlete_id INTEGER,
    date TEXT,
    present INTEGER,
    UNIQUE(athlete_id, date)
  );
  CREATE TABLE paiements(
    id INTEGER PRIMARY KEY,
    athlete_id INTEGER,
    montant TEXT,
    date TEXT,
    jusqua TEXT
  );
  CREATE TABLE grades(
    id INTEGER PRIMARY KEY,
    athlete_id INTEGER,
    ceinture TEXT,
    date TEXT
  );
  CREATE TABLE competitions(
    id INTEGER PRIMARY KEY,
    nom TEXT,
    date TEXT,
    lieu TEXT
  );
  CREATE TABLE resultats(
    id INTEGER PRIMARY KEY,
    competition_id INTEGER,
    athlete_id INTEGER,
    resultat TEXT
  );`,

  // Migration 2
  `CREATE TABLE saisons(
    id INTEGER PRIMARY KEY,
    nom TEXT,
    debut TEXT,
    fin TEXT,
    frais INTEGER DEFAULT 0,
    tarif INTEGER DEFAULT 0,
    actif INTEGER DEFAULT 0
  );
  CREATE TABLE inscriptions(
    id INTEGER PRIMARY KEY,
    saison_id INTEGER,
    athlete_id INTEGER,
    remise INTEGER DEFAULT 0,
    statut TEXT DEFAULT 'actif',
    UNIQUE(saison_id, athlete_id)
  );
  ALTER TABLE paiements ADD COLUMN saison_id INTEGER;
  ALTER TABLE paiements ADD COLUMN mode TEXT;
  ALTER TABLE paiements ADD COLUMN note TEXT;`,

  // Migration 3
  `CREATE TABLE creneaux(
    id INTEGER PRIMARY KEY,
    groupe_id INTEGER,
    jour TEXT,
    debut TEXT,
    fin TEXT
  );
  ALTER TABLE saisons ADD COLUMN frais_grade INTEGER DEFAULT 1000;
  ALTER TABLE grades ADD COLUMN frais INTEGER;
  ALTER TABLE competitions ADD COLUMN frais INTEGER DEFAULT 0;
  ALTER TABLE resultats ADD COLUMN frais INTEGER;
  ALTER TABLE paiements ADD COLUMN type TEXT;
  ALTER TABLE paiements ADD COLUMN ref_id INTEGER;`,

  // Migration 4
  `ALTER TABLE inscriptions ADD COLUMN mode TEXT DEFAULT 'annuel';
  ALTER TABLE inscriptions ADD COLUMN date_debut TEXT;
  ALTER TABLE saisons ADD COLUMN nb_mois INTEGER DEFAULT 10;
  ALTER TABLE paiements ADD COLUMN periode TEXT;
  ALTER TABLE paiements ADD COLUMN annule INTEGER DEFAULT 0;
  ALTER TABLE paiements ADD COLUMN motif TEXT;
  ALTER TABLE grades ADD COLUMN issue TEXT DEFAULT 'Inscrit';`,

  // Migration 5
  `CREATE TABLE presences_coachs(
    id INTEGER PRIMARY KEY,
    coach_id INTEGER,
    groupe_id INTEGER,
    date TEXT,
    present INTEGER,
    UNIQUE(coach_id, groupe_id, date)
  );
  CREATE TABLE annulations(
    id INTEGER PRIMARY KEY,
    groupe_id INTEGER,
    date TEXT,
    motif TEXT,
    UNIQUE(groupe_id, date)
  );`,

  // Migration 6
  `ALTER TABLE athletes ADD COLUMN date_integration TEXT;
  ALTER TABLE inscriptions ADD COLUMN delai TEXT;`,

  // Migration 7
  `ALTER TABLE athletes ADD COLUMN groupe_sanguin TEXT;
  ALTER TABLE athletes ADD COLUMN taille TEXT;
  ALTER TABLE athletes ADD COLUMN poids TEXT;
  ALTER TABLE athletes ADD COLUMN profession TEXT;
  ALTER TABLE athletes ADD COLUMN club_precedent TEXT;
  ALTER TABLE athletes ADD COLUMN profession_pere TEXT;
  CREATE TABLE parametres(
    cle TEXT PRIMARY KEY,
    valeur TEXT
  );`,

  // Migration 8 : garantit la présence des paramètres du club
  `CREATE TABLE IF NOT EXISTS parametres(
    cle TEXT PRIMARY KEY,
    valeur TEXT
  );`,

  // Migration 9 : suivi de l'assurance (transmission à l'assureur)
  `ALTER TABLE inscriptions ADD COLUMN assur_transmise TEXT;
  ALTER TABLE inscriptions ADD COLUMN assur_ref TEXT;`,

  // Migration 10 : circuit de validation des dossiers d'inscription
  `ALTER TABLE inscriptions ADD COLUMN dossier_statut TEXT;
  ALTER TABLE inscriptions ADD COLUMN dossier_pieces TEXT;
  ALTER TABLE inscriptions ADD COLUMN dossier_motif TEXT;
  ALTER TABLE inscriptions ADD COLUMN dossier_date TEXT;`,

  // Migration 11 : journal des relances envoyées aux familles
  `CREATE TABLE IF NOT EXISTS relances(
    id INTEGER PRIMARY KEY,
    athlete_id INTEGER,
    date TEXT,
    canal TEXT,
    message TEXT
  );`,

  // Migration 12 : lien du groupe WhatsApp de chaque groupe d'entraînement
  `ALTER TABLE groupes ADD COLUMN lien_whatsapp TEXT;`,

  // Migration 13 : utilisateurs, rôles, journal des actions, clé de récupération
  `CREATE TABLE IF NOT EXISTS utilisateurs(
    id INTEGER PRIMARY KEY,
    nom TEXT,
    identifiant TEXT UNIQUE,
    role TEXT,
    sel TEXT,
    hash TEXT,
    actif INTEGER DEFAULT 1,
    doit_changer INTEGER DEFAULT 0,
    echecs INTEGER DEFAULT 0,
    verrou_jusqua INTEGER DEFAULT 0,
    cree_le TEXT
  );
  CREATE TABLE IF NOT EXISTS journal(
    id INTEGER PRIMARY KEY,
    horodatage TEXT,
    utilisateur TEXT,
    role TEXT,
    action TEXT,
    detail TEXT
  );
  CREATE TABLE IF NOT EXISTS securite(
    cle TEXT PRIMARY KEY,
    valeur TEXT
  );`
];

const q = (sql, params = []) => {
  const stmt = db.prepare(sql);
  const rows = [];

  try {
    stmt.bind(params);

    while (stmt.step()) {
      rows.push(stmt.getAsObject());
    }
  } finally {
    stmt.free();
  }

  return rows;
};

const persist = () => {
  fs.writeFileSync(dbFile, Buffer.from(db.export()));
};

function createBackup() {
  if (!fs.existsSync(dbFile)) return;

  const backupDir = path.join(dir(), 'sauvegardes');
  fs.mkdirSync(backupDir, { recursive: true });

  const stamp = new Date()
    .toISOString()
    .replace(/[:.]/g, '-');

  const backupFile = path.join(
    backupDir,
    `club_${stamp}.db`
  );

  fs.copyFileSync(dbFile, backupFile);

  // Conserve les 15 sauvegardes les plus récentes.
  const backups = fs.readdirSync(backupDir)
    .filter(name => /^club_.*\.db$/.test(name))
    .map(name => ({
      name,
      time: fs.statSync(path.join(backupDir, name)).mtimeMs
    }))
    .sort((a, b) => b.time - a.time);

  backups.slice(15).forEach(item => {
    fs.unlinkSync(path.join(backupDir, item.name));
  });
}

async function init() {
  fs.mkdirSync(path.join(dir(), 'documents'), {
    recursive: true
  });

  fs.mkdirSync(path.join(dir(), 'sauvegardes'), {
    recursive: true
  });

  dbFile = path.join(dir(), 'club.db');

  // En développement, le fichier WASM se trouve dans node_modules.
  // Dans l'application packagée, il peut être placé dans app.asar.unpacked.
  const SQL = await initSqlJs({
    locateFile: file => {
      const unpackedPath = path.join(
        process.resourcesPath || __dirname,
        'app.asar.unpacked',
        'node_modules',
        'sql.js',
        'dist',
        file
      );

      const developmentPath = path.join(
        __dirname,
        'node_modules',
        'sql.js',
        'dist',
        file
      );

      const resourcesPath = path.join(
        process.resourcesPath || __dirname,
        'node_modules',
        'sql.js',
        'dist',
        file
      );

      if (fs.existsSync(unpackedPath)) return unpackedPath;
      if (fs.existsSync(developmentPath)) return developmentPath;
      if (fs.existsSync(resourcesPath)) return resourcesPath;

      return developmentPath;
    }
  });

  // Sauvegarde avant d'ouvrir ou de faire évoluer la base existante.
  createBackup();

  db = fs.existsSync(dbFile)
    ? new SQL.Database(fs.readFileSync(dbFile))
    : new SQL.Database();

  const versionRow = q('PRAGMA user_version')[0];
  let version = Number(versionRow?.user_version || 0);

  if (version > MIGRATIONS.length) {
    throw new Error(
      `La base de données est en version ${version}, ` +
      `mais cette application ne connaît que la version ${MIGRATIONS.length}.`
    );
  }

  while (version < MIGRATIONS.length) {
    db.exec('BEGIN');

    try {
      db.exec(MIGRATIONS[version]);
      version += 1;
      db.run(`PRAGMA user_version = ${version}`);
      db.exec('COMMIT');
    } catch (error) {
      try {
        db.exec('ROLLBACK');
      } catch (_) {}

      throw error;
    }
  }

  // Crée une première saison si nécessaire.
  if (!q('SELECT id FROM saisons LIMIT 1')[0]) {
    const now = new Date();
    const year = now.getFullYear();
    const startYear = now.getMonth() >= 8 ? year : year - 1;

    db.run(
      `INSERT INTO saisons(nom, debut, fin, actif)
       VALUES(?,?,?,1)`,
      [
        `${startYear}-${startYear + 1}`,
        `${startYear}-09-01`,
        `${startYear + 1}-08-31`
      ]
    );

    const seasonId = q(
      'SELECT last_insert_rowid() AS id'
    )[0].id;

    db.run(
      `INSERT INTO inscriptions(saison_id, athlete_id)
       SELECT ?, id FROM athletes`,
      [seasonId]
    );

    db.run(
      'UPDATE paiements SET saison_id=? WHERE saison_id IS NULL',
      [seasonId]
    );
  }

  // Insère les paramètres par défaut sans écraser les réglages existants.
  for (const [key, value] of Object.entries(PARAMETRES_DEFAUT)) {
    db.run(
      `INSERT OR IGNORE INTO parametres(cle, valeur)
       VALUES(?,?)`,
      [key, value]
    );
  }

  persist();
}

// ─────────────────────────────────────────────
// ATHLÈTES
// ─────────────────────────────────────────────

function saveAth(a, quiet = false, noIns = false) {
  if (!a || typeof a !== 'object') {
    throw new Error('Données athlète invalides.');
  }

  const fields = COLS.filter(c =>
    a.id ? Object.prototype.hasOwnProperty.call(a, c) :
      a[c] != null && a[c] !== ''
  );

  let id = a.id;

  if (id) {
    if (fields.length) {
      db.run(
        `UPDATE athletes SET ${fields.map(c => c + '=?').join(', ')}
         WHERE id=?`,
        [...fields.map(c => a[c] ?? null), id]
      );
    }
  } else {
    if (!fields.length) {
      throw new Error('Aucune donnée à enregistrer pour cet athlète.');
    }

    db.run(
      `INSERT INTO athletes(${fields.join(',')})
       VALUES(${fields.map(() => '?').join(',')})`,
      fields.map(c => a[c])
    );

    id = q('SELECT last_insert_rowid() AS id')[0].id;

    const season = q(
      'SELECT id FROM saisons WHERE actif=1 LIMIT 1'
    )[0];

    if (season && !noIns) {
      db.run(
        `INSERT OR IGNORE INTO inscriptions
         (saison_id, athlete_id, date_debut)
         VALUES(?,?,date('now'))`,
        [season.id, id]
      );
    }
  }

  if (!quiet) persist();

  return id;
}

ipcMain.handle('ath:list', () =>
  q('SELECT * FROM athletes ORDER BY nom, prenom')
);

ipcMain.handle('ath:save', (e, athlete) =>
  saveAth(athlete)
);

ipcMain.handle('ath:del', (e, id) => {
  db.run('DELETE FROM athletes WHERE id=?', [id]);
  db.run('DELETE FROM documents WHERE athlete_id=?', [id]);
  persist();
});

ipcMain.handle('ath:groupes', (e, ids, groupName) => {
  if (!Array.isArray(ids)) return;

  ids.forEach(id => {
    db.run(
      'UPDATE athletes SET groupe=? WHERE id=?',
      [groupName, id]
    );
  });

  persist();
});

ipcMain.handle('ath:groupe', (e, id, groupName) => {
  db.run(
    'UPDATE athletes SET groupe=? WHERE id=?',
    [groupName, id]
  );

  persist();
});

// ─────────────────────────────────────────────
// DOCUMENTS DES ATHLÈTES
// ─────────────────────────────────────────────

ipcMain.handle('doc:list', (e, id) =>
  q(
    'SELECT * FROM documents WHERE athlete_id=? ORDER BY id DESC',
    [id]
  )
);

ipcMain.handle('doc:open', async (e, relativeFile) => {
  if (typeof relativeFile !== 'string' || !relativeFile.trim()) {
    throw new Error('Chemin du document invalide.');
  }

  const documentsDir = path.resolve(dir(), 'documents');
  const target = path.resolve(documentsDir, relativeFile);

  if (
    target !== documentsDir &&
    !target.startsWith(documentsDir + path.sep)
  ) {
    throw new Error('Accès au document refusé.');
  }

  if (!fs.existsSync(target) || !fs.statSync(target).isFile()) {
    throw new Error('Le document demandé est introuvable.');
  }

  return shell.openPath(target);
});

ipcMain.handle('doc:add', async (e, athleteId, type) => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    filters: [{
      name: 'Documents',
      extensions: ['pdf', 'jpg', 'jpeg', 'png']
    }]
  });

  if (result.canceled || !result.filePaths.length) return;

  const athleteDir = path.join(
    dir(),
    'documents',
    String(athleteId)
  );

  fs.mkdirSync(athleteDir, { recursive: true });

  for (const source of result.filePaths) {
    const filename =
      `${Date.now()}_${path.basename(source)}`;

    const relativeFile = `${athleteId}/${filename}`;
    const destination = path.join(dir(), 'documents', relativeFile);

    fs.copyFileSync(source, destination);

    db.run(
      `INSERT INTO documents(athlete_id, type, fichier)
       VALUES(?,?,?)`,
      [athleteId, type, relativeFile]
    );
  }

  persist();
});

// ─────────────────────────────────────────────
// IMPORT / EXPORT EXCEL
// ─────────────────────────────────────────────

const xl = async (worksheet, defaultName) => {
  const result = await dialog.showSaveDialog({
    defaultPath: defaultName,
    filters: [{
      name: 'Fichier Excel',
      extensions: ['xlsx']
    }]
  });

  if (result.canceled || !result.filePath) return null;

  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    'Athletes'
  );

  XLSX.writeFile(workbook, result.filePath);

  return result.filePath;
};

ipcMain.handle('xl:model', () =>
  xl(
    XLSX.utils.aoa_to_sheet([COLS]),
    'modele_athletes.xlsx'
  )
);

ipcMain.handle('xl:export', () =>
  xl(
    XLSX.utils.json_to_sheet(
      q(`SELECT ${COLS.join(',')} FROM athletes`),
      { header: COLS }
    ),
    'athletes.xlsx'
  )
);

ipcMain.handle('xl:import', async (e, inscrire) => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [{
      name: 'Excel / CSV',
      extensions: ['xlsx', 'xls', 'csv']
    }]
  });

  if (result.canceled || !result.filePaths[0]) return null;

  const workbook = XLSX.readFile(result.filePaths[0], {
    cellDates: true
  });

  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

  if (!firstSheet) {
    throw new Error('Le fichier ne contient aucune feuille exploitable.');
  }

  const iso = value => {
    if (value instanceof Date) {
      const year = value.getFullYear();
      const month = String(value.getMonth() + 1).padStart(2, '0');
      const day = String(value.getDate()).padStart(2, '0');

      return `${year}-${month}-${day}`;
    }

    const match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
      .exec(String(value ?? '').trim());

    if (match) {
      return `${match[3]}-${match[2].padStart(2, '0')}-${match[1].padStart(2, '0')}`;
    }

    return String(value ?? '').trim();
  };

  const key = athlete =>
    [athlete.nom, athlete.prenom, athlete.date_naissance]
      .map(value => String(value ?? '').trim().toLowerCase())
      .join('|');

  const existing = new Set(
    q('SELECT nom, prenom, date_naissance FROM athletes').map(key)
  );

  const rows = XLSX.utils.sheet_to_json(firstSheet, {
    raw: true,
    defval: null
  }).filter(row => row.nom);

  let imported = 0;
  let skipped = 0;

  db.run('BEGIN');

  try {
    rows.forEach(row => {
      COLS.forEach(column => {
        if (row[column] != null) {
          row[column] =
            column.includes('date') || column.endsWith('expiration')
              ? iso(row[column])
              : String(row[column]).trim();
        }
      });

      const athleteKey = key(row);

      if (existing.has(athleteKey)) {
        skipped++;
        return;
      }

      existing.add(athleteKey);

      saveAth(row, true, !inscrire);
      imported++;
    });

    db.run('COMMIT');
  } catch (error) {
    db.run('ROLLBACK');
    throw error;
  }

  persist();

  return { n: imported, skip: skipped };
});

// ─────────────────────────────────────────────
// MODULES GÉNÉRIQUES
// ─────────────────────────────────────────────

const T = [
  'relances',
  'presences_coachs',
  'annulations',
  'presences',
  'creneaux',
  'saisons',
  'inscriptions',
  'coachs',
  'groupes',
  'paiements',
  'grades',
  'competitions',
  'resultats'
];

const cols = table =>
  q(`PRAGMA table_info(${table})`)
    .map(column => column.name)
    .filter(name => name !== 'id');

ipcMain.handle('row:list', (e, table) => {
  if (!T.includes(table)) return [];

  return q(`SELECT * FROM ${table} ORDER BY id DESC`);
});

ipcMain.handle('row:save', (e, table, data) => {
  if (!T.includes(table)) {
    throw new Error('Table non autorisée.');
  }

  if (!data || typeof data !== 'object') {
    throw new Error('Données à enregistrer invalides.');
  }

  const fields = cols(table).filter(column =>
    Object.prototype.hasOwnProperty.call(data, column)
  );

  if (data.id) {
    if (fields.length) {
      db.run(
        `UPDATE ${table} SET ${fields.map(c => c + '=?').join(', ')}
         WHERE id=?`,
        [...fields.map(c => data[c] ?? null), data.id]
      );
    }
  } else {
    if (!fields.length) {
      throw new Error('Aucune donnée à enregistrer.');
    }

    db.run(
      `INSERT INTO ${table}(${fields.join(',')})
       VALUES(${fields.map(() => '?').join(',')})`,
      fields.map(c => data[c] ?? null)
    );
  }

  if (table === 'paiements' && data.jusqua) {
    db.run(
      'UPDATE athletes SET paye_jusqua=? WHERE id=?',
      [data.jusqua, data.athlete_id]
    );
  }

  if (
    table === 'grades' &&
    data.ceinture &&
    data.issue === 'Réussi'
  ) {
    db.run(
      'UPDATE athletes SET ceinture=? WHERE id=?',
      [data.ceinture, data.athlete_id]
    );
  }

  persist();

  return data.id || q(
    `SELECT MAX(id) AS id FROM ${table}`
  )[0].id;
});

ipcMain.handle('row:del', (e, table, id) => {
  // Les paiements sont conservés dans l'historique.
  if (T.includes(table) && table !== 'paiements') {
    db.run(`DELETE FROM ${table} WHERE id=?`, [id]);
    persist();
  }
});

// ─────────────────────────────────────────────
// SAISONS ET INSCRIPTIONS
// ─────────────────────────────────────────────

ipcMain.handle('saison:ouvrir', (e, season) => {
  const previous = q(
    'SELECT id FROM saisons WHERE actif=1 LIMIT 1'
  )[0];

  db.run('UPDATE saisons SET actif=0');

  db.run(
    `INSERT INTO saisons
     (nom, debut, fin, frais, tarif, frais_grade, nb_mois, actif)
     VALUES(?,?,?,?,?,?,?,1)`,
    [
      season.nom,
      season.debut,
      season.fin,
      season.frais,
      season.tarif,
      season.frais_grade ?? 1000,
      season.nb_mois ?? 10
    ]
  );

  const id = q(
    'SELECT last_insert_rowid() AS id'
  )[0].id;

  if (season.renew && previous) {
    db.run(
      `INSERT INTO inscriptions
       (saison_id, athlete_id, remise, statut, mode)
       SELECT ?, athlete_id, 0, 'actif', mode
       FROM inscriptions
       WHERE saison_id=? AND statut='actif'`,
      [id, previous.id]
    );
  }

  persist();
  return id;
});

ipcMain.handle('ins:set', (e, seasonId, athleteId, status, discount) => {
  db.run(
    `INSERT INTO inscriptions
     (saison_id, athlete_id, statut, remise, date_debut)
     VALUES(?,?,?,?,date('now'))
     ON CONFLICT(saison_id, athlete_id)
     DO UPDATE SET
       statut=COALESCE(?, statut),
       remise=COALESCE(?, remise)`,
    [
      seasonId,
      athleteId,
      status ?? 'actif',
      discount ?? 0,
      status ?? null,
      discount ?? null
    ]
  );

  persist();
});

ipcMain.handle('ins:mode', (e, seasonId, athleteId, mode) => {
  db.run(
    'UPDATE inscriptions SET mode=? WHERE saison_id=? AND athlete_id=?',
    [mode, seasonId, athleteId]
  );

  persist();
});

ipcMain.handle('ins:delai', (e, seasonId, athleteId, date) => {
  db.run(
    'UPDATE inscriptions SET delai=? WHERE saison_id=? AND athlete_id=?',
    [date, seasonId, athleteId]
  );

  persist();
});

// ─────────────────────────────────────────────
// GROUPES ET CRÉNEAUX
// ─────────────────────────────────────────────

ipcMain.handle('ins:assur', (e, seasonId, athleteIds, date, ref) => {
  if (!Array.isArray(athleteIds)) return;
  athleteIds.forEach(id => db.run(
    'UPDATE inscriptions SET assur_transmise=?, assur_ref=? WHERE saison_id=? AND athlete_id=?',
    [date ?? null, ref ?? null, seasonId, id]
  ));
  persist();
});

// Export d'une liste (en-têtes + lignes) vers un fichier Excel
ipcMain.handle('list:export', async (e, nom, entetes, lignes) => {
  if (!Array.isArray(entetes) || !Array.isArray(lignes)) return false;
  const r = await dialog.showSaveDialog({ defaultPath: String(nom || 'liste.xlsx'), filters: [{ name: 'Excel', extensions: ['xlsx'] }] });
  if (!r.filePath) return false;
  const ws = XLSX.utils.aoa_to_sheet([entetes, ...lignes]);
  ws['!cols'] = entetes.map(() => ({ wch: 24 }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Liste');
  XLSX.writeFile(wb, r.filePath);
  return true;
});

ipcMain.handle('dos:set', (e, seasonId, athleteId, statut, pieces, motif, date) => {
  if (statut !== null && statut !== 'valide' && statut !== 'rejete') return;
  db.run(
    'UPDATE inscriptions SET dossier_statut=?, dossier_pieces=?, dossier_motif=?, dossier_date=? WHERE saison_id=? AND athlete_id=?',
    [statut ?? null, pieces ?? null, motif ?? null, date ?? null, seasonId, athleteId]
  );
  persist();
});

// ── Sauvegarde / restauration (dossier complet : base + documents scannés) ──
ipcMain.handle('sauv:export', async () => {
  const r = await dialog.showOpenDialog({ title: 'Choisir le dossier de destination (clé USB…)', properties: ['openDirectory', 'createDirectory'] });
  if (r.canceled || !r.filePaths[0]) return null;
  persist();
  const dest = path.join(r.filePaths[0], `ASAMB-sauvegarde-${new Date().toISOString().slice(0, 10)}`);
  fs.mkdirSync(dest, { recursive: true });
  fs.copyFileSync(dbFile, path.join(dest, 'club.db'));
  const docs = path.join(dir(), 'documents');
  if (fs.existsSync(docs)) fs.cpSync(docs, path.join(dest, 'documents'), { recursive: true });
  fs.writeFileSync(path.join(dest, 'LISEZMOI.txt'), 'Sauvegarde ASAMB. Pour restaurer : menu Paramètres > Restaurer une sauvegarde, puis choisir CE dossier.');
  return dest;
});

ipcMain.handle('sauv:restore', async () => {
  const r = await dialog.showOpenDialog({ title: 'Choisir le dossier de sauvegarde ASAMB', properties: ['openDirectory'] });
  if (r.canceled || !r.filePaths[0]) return null;
  const src = r.filePaths[0], srcDb = path.join(src, 'club.db');
  if (!fs.existsSync(srcDb)) return { erreur: 'Ce dossier ne contient pas de fichier club.db.' };
  const fd = fs.openSync(srcDb, 'r'), head = Buffer.alloc(15);
  fs.readSync(fd, head, 0, 15, 0); fs.closeSync(fd);
  if (head.toString('latin1') !== 'SQLite format 3') return { erreur: 'Ce fichier club.db n’est pas une base valide.' };
  persist();
  const secu = path.join(dir(), 'sauvegardes', `avant-restauration-${Date.now()}`);
  fs.mkdirSync(secu, { recursive: true });
  fs.copyFileSync(dbFile, path.join(secu, 'club.db'));
  fs.copyFileSync(srcDb, dbFile);
  const docs = path.join(src, 'documents');
  if (fs.existsSync(docs)) fs.cpSync(docs, path.join(dir(), 'documents'), { recursive: true, force: true });
  app.relaunch();
  app.exit(0); // pas de 'before-quit' : la base restaurée ne doit pas être réécrasée par la mémoire
  return { ok: true };
});

// Ouverture d'un lien WhatsApp (seules adresses autorisées : wa.me/<numéro>?text=... et chat.whatsapp.com/<invitation>)
ipcMain.handle('lien:ouvrir', (e, url) => {
  if (!/^https:\/\/(wa\.me\/\d{8,15}(\?text=[^\s]*)?|chat\.whatsapp\.com\/[A-Za-z0-9]{10,40})$/.test(String(url || ''))) return false;
  shell.openExternal(String(url));
  return true;
});

ipcMain.handle('grp:save', (e, group, slots = []) => {
  let id = group.id;
  const lien = /^https:\/\/chat\.whatsapp\.com\/[A-Za-z0-9]{10,40}$/.test(String(group.lien_whatsapp || '')) ? group.lien_whatsapp : null;

  if (id) {
    const oldGroup = q(
      'SELECT nom FROM groupes WHERE id=?',
      [id]
    )[0];

    if (oldGroup && oldGroup.nom !== group.nom) {
      db.run(
        'UPDATE athletes SET groupe=? WHERE groupe=?',
        [group.nom, oldGroup.nom]
      );
    }

    db.run(
      `UPDATE groupes
       SET nom=?, categorie=?, coach_id=?, lien_whatsapp=?
       WHERE id=?`,
      [
        group.nom,
        group.categorie ?? null,
        group.coach_id ?? null,
        lien,
        id
      ]
    );
  } else {
    db.run(
      'INSERT INTO groupes(nom, categorie, coach_id, lien_whatsapp) VALUES(?,?,?,?)',
      [
        group.nom,
        group.categorie ?? null,
        group.coach_id ?? null,
        lien
      ]
    );

    id = q('SELECT last_insert_rowid() AS id')[0].id;
  }

  db.run('DELETE FROM creneaux WHERE groupe_id=?', [id]);

  slots.forEach(slot => {
    db.run(
      'INSERT INTO creneaux(groupe_id, jour, debut, fin) VALUES(?,?,?,?)',
      [id, slot.jour, slot.debut, slot.fin]
    );
  });

  persist();
  return id;
});

ipcMain.handle('grp:del', (e, id) => {
  db.run('DELETE FROM creneaux WHERE groupe_id=?', [id]);
  db.run('DELETE FROM groupes WHERE id=?', [id]);
  persist();
});

// ─────────────────────────────────────────────
// PAIEMENTS
// ─────────────────────────────────────────────

// Un paiement n'est jamais supprimé : il est annulé avec un motif.
ipcMain.handle('pay:cancel', (e, id, reason) => {
  db.run(
    'UPDATE paiements SET annule=1, motif=? WHERE id=?',
    [reason, id]
  );

  persist();
});

// ─────────────────────────────────────────────
// PRÉSENCES DES COACHS ET ANNULATIONS
// ─────────────────────────────────────────────

ipcMain.handle('pc:save', (e, date, groupId, rows) => {
  if (!Array.isArray(rows)) return;

  rows.forEach(([coachId, present]) => {
    db.run(
      `INSERT OR REPLACE INTO presences_coachs
       (coach_id, groupe_id, date, present)
       VALUES(?,?,?,?)`,
      [coachId, groupId, date, present]
    );
  });

  persist();
});

ipcMain.handle('ann:set', (e, groupIds, date, reason) => {
  if (!Array.isArray(groupIds)) return;

  groupIds.forEach(groupId => {
    db.run(
      `INSERT OR REPLACE INTO annulations(groupe_id, date, motif)
       VALUES(?,?,?)`,
      [groupId, date, reason]
    );
  });

  persist();
});

// ─────────────────────────────────────────────
// PRÉSENCES DES ATHLÈTES
// ─────────────────────────────────────────────

ipcMain.handle('pres:get', (e, date) =>
  q(
    'SELECT athlete_id, present FROM presences WHERE date=?',
    [date]
  )
);

ipcMain.handle('pres:stats', () =>
  q(
    `SELECT athlete_id,
            SUM(present) AS p,
            COUNT(*) AS n
     FROM presences
     GROUP BY athlete_id`
  )
);

ipcMain.handle('pres:save', (e, date, rows) => {
  if (!Array.isArray(rows)) return;

  rows.forEach(([athleteId, present]) => {
    db.run(
      `INSERT OR REPLACE INTO presences
       (athlete_id, date, present)
       VALUES(?,?,?)`,
      [athleteId, date, present]
    );
  });

  persist();
});

// ─────────────────────────────────────────────
// PARAMÈTRES DU CLUB
// ─────────────────────────────────────────────

// Renvoie les paramètres enregistrés, avec les valeurs par défaut
// pour les clés qui ne seraient pas encore présentes.
ipcMain.handle('param:get', () => {
  const saved = q(
    'SELECT cle, valeur FROM parametres'
  );

  const result = { ...PARAMETRES_DEFAUT };

  saved.forEach(row => {
    result[row.cle] = row.valeur ?? '';
  });

  return result;
});

ipcMain.handle('param:set', (e, key, value) => {
  if (!Object.prototype.hasOwnProperty.call(PARAMETRES_DEFAUT, key)) {
    throw new Error(`Paramètre non autorisé : ${key}`);
  }

  const safeValue = value == null ? '' : String(value);

  db.run(
    `INSERT INTO parametres(cle, valeur)
     VALUES(?,?)
     ON CONFLICT(cle)
     DO UPDATE SET valeur=excluded.valeur`,
    [key, safeValue]
  );

  persist();

  return true;
});

// ─────────────────────────────────────────────
// DÉMARRAGE DE L'APPLICATION
// ─────────────────────────────────────────────

// Une seule instance à la fois : deux fenêtres s'écraseraient mutuellement la base de données.
const verrou = app.requestSingleInstanceLock();
if (!verrou) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });
}

app.whenReady().then(async () => {
  if (!verrou) return;
  try {
    await init();
  } catch (error) {
    dialog.showErrorBox(
      'Base de données',
      String(error.stack || error)
    );

    app.quit();
    return;
  }

  mainWindow = new BrowserWindow({
    width: 1280,
    height: 820,
    minWidth: 1000,
    minHeight: 650,
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, 'index.html'));
});

app.on('before-quit', () => {
  try {
    if (db && dbFile) {
      fs.writeFileSync(dbFile, Buffer.from(db.export()));
    }
  } catch (error) {
    console.error('Erreur lors de la sauvegarde finale :', error);
  }
});
