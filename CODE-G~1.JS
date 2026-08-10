// ══════════════════════════════════════════════════════════════
//  JTCF — Positionnement · doGet corrigé
//  Remplace intégralement la fonction doGet existante.
//
//  CORRECTIF : la détection de structure reposait sur la colonne N
//  (row[13]). Dès qu'elle était vide, une ligne récente était lue
//  avec l'ancien mapping — d'où l'heure affichée en formation et la
//  formation affichée en nom.
//  On détecte désormais sur la colonne B : une heure dans le nouveau
//  format, un intitulé de formation dans l'ancien.
// ══════════════════════════════════════════════════════════════

// La colonne B contient-elle une heure ? (Date, fraction de jour, ou "HH:MM")
function estColonneHeure(v) {
  if (v instanceof Date) return true;
  if (typeof v === 'number' && v >= 0 && v < 1) return true;
  if (typeof v === 'string') return /^\d{1,2}\s*[:hH]\s*\d{2}/.test(v.trim());
  return false;
}

function formaterHeure(v) {
  if (v instanceof Date) {
    return ('0' + v.getHours()).slice(-2) + ':' + ('0' + v.getMinutes()).slice(-2);
  }
  if (typeof v === 'number' && v >= 0 && v < 1) {
    var tot = Math.round(v * 24 * 60);
    return ('0' + Math.floor(tot / 60)).slice(-2) + ':' + ('0' + (tot % 60)).slice(-2);
  }
  return String(v || '').substring(0, 5);
}

function formaterDate(v) {
  if (v instanceof Date) return v.toLocaleDateString('fr-FR');
  return v || '';
}

function doGet(e) {
  const params = e ? e.parameter : {};
  const code   = params.code || '';

  if (code !== 'JTCF-ADMIN-2026') {
    return ContentService.createTextOutput(JSON.stringify({status:"ERREUR", message:"Code incorrect"}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data  = sheet.getDataRange().getValues();

  const resultats = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;

    // ── DÉTECTION DE STRUCTURE ────────────────────────────────
    // Nouveau format (14 colonnes) : B = heure
    // Ancien format  ( 9 colonnes) : B = intitulé de formation
    const isNouveau = estColonneHeure(row[1]);

    let r;

    if (isNouveau) {
      r = {
        id: i,
        date:        formaterDate(row[0]),
        heure:       formaterHeure(row[1]),
        formation:   row[2] || '',
        formationId: row[3] || '',
        nom:         row[4] || '',
        prenom:      row[5] || '',
        ddn:         row[6] instanceof Date ? row[6].toLocaleDateString('fr-FR') : (row[6] || ''),
        tel:         row[7] || '',
        email:       row[8] || '',
        niveau:      row[9] || '',
        scoreQCM:    row[10] || '',
        scorePct:    row[11] || '',
        reponses:    row[13] || ''
      };
    } else {
      const texte = String(row[8] || '');
      r = {
        id: i,
        date:        formaterDate(row[0]),
        heure:       '',
        formation:   row[1] || '',
        formationId: '',
        nom:         row[2] || '',
        prenom:      row[3] || '',
        ddn:         row[4] instanceof Date ? row[4].toLocaleDateString('fr-FR') : (row[4] || ''),
        tel:         row[5] || '',
        email:       row[6] || '',
        niveau:      row[7] || '',
        scoreQCM:    '',
        scorePct:    '',
        reponses:    texte
      };
      r.scoreQCM = calculerScoreDepuisTexte(texte, r.formation);
      if (r.scoreQCM) {
        const p = r.scoreQCM.split('/');
        r.scorePct = Math.round(parseInt(p[0]) / parseInt(p[1]) * 100) + '%';
      }
    }

    // ── FILET DE SÉCURITÉ ─────────────────────────────────────
    // L'en-tête du texte des réponses fait foi : il est écrit par le
    // formulaire et ne dépend d'aucune colonne.
    const txt = String(r.reponses || '');
    const mF = txt.match(/^\s*FORMATION\s*:\s*(.+)$/mi);
    const mC = txt.match(/^\s*CANDIDAT\s*:\s*(.+)$/mi);
    if (mF) r.formation = mF[1].trim();
    if (mC) {
      const parts = mC[1].trim().split(/\s+/);
      r.prenom = parts.shift() || '';
      r.nom    = parts.join(' ');
      if (!r.nom) { r.nom = r.prenom; r.prenom = ''; }
    }

    if (r.formation) resultats.push(r);
  }

  const filtreFormation = params.formation || '';
  const liste = filtreFormation
    ? resultats.filter(function(r){ return String(r.formation).includes(filtreFormation); })
    : resultats;

  return ContentService.createTextOutput(JSON.stringify({status:"OK", resultats: liste.reverse()}))
    .setMimeType(ContentService.MimeType.JSON);
}
