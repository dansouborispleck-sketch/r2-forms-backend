require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const pdf = require('pdf-parse');
const fetch = require('node-fetch');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 4000;

app.use(cors({ origin: '*' }));
app.use(express.json({ limit: '10mb' }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });
app.use(express.static(path.join(__dirname, 'public')));

app.get('/health', (req, res) => res.json({ status: 'ok', version: '3.0.0' }));

// ============ IMPORT ============
app.post('/api/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier reçu' });
    const { originalname, buffer } = req.file;
    const ext = originalname.split('.').pop().toLowerCase();
    let text = '';
    console.log(`[IMPORT] ${originalname}`);

    if (['txt','csv'].includes(ext)) text = buffer.toString('utf-8');
    else if (ext === 'pdf') {
      const d = await pdf(buffer);
      text = d.text;
      if (!text || text.trim().length < 20) return res.status(422).json({ error: 'PDF_SCANNED', message: 'PDF scanné illisible. Collez le texte directement.' });
    }
    else if (ext === 'docx') { const r = await mammoth.extractRawText({ buffer }); text = r.value; }
    else if (ext === 'doc') {
      text = buffer.toString('latin1').replace(/[^\x20-\x7E\n\r\u00C0-\u024F]/g,' ').replace(/\s+/g,' ').trim();
      if (text.length < 30) return res.status(422).json({ error: 'DOC_OLD', message: 'Format .doc non supporté. Enregistrez en .docx.' });
    }
    else if (['xlsx','xls'].includes(ext)) {
      const wb = XLSX.read(buffer, { type: 'buffer' });
      wb.SheetNames.forEach(n => { text += `\n=== ${n} ===\n` + XLSX.utils.sheet_to_csv(wb.Sheets[n]); });
    }
    else return res.status(400).json({ error: 'FORMAT_UNSUPPORTED', message: `Format .${ext} non supporté.` });

    text = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n').replace(/\n{3,}/g,'\n\n').trim();
    if (text.length < 20) return res.status(422).json({ error: 'EMPTY', message: 'Fichier vide. Collez le texte directement.' });

    console.log(`[IMPORT] ✓ ${text.length} chars`);
    res.json({ success: true, text, metadata: { filename: originalname, chars: text.length } });
  } catch(err) {
    console.error('[IMPORT ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ ANALYSE ============
app.post('/api/analyse', async (req, res) => {
  try {
    const { text, tool } = req.body;
    if (!text || text.trim().length < 10) return res.status(400).json({ error: 'Texte trop court' });
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'Clé API manquante' });

    const inputText = text.slice(0, 8000);
    const truncated = text.length > 8000;
    console.log(`[ANALYSE] ${text.length} chars → ${tool}${truncated?' (tronqué)':''}`);

    const system = `Expert collecte de donnees terrain. Analyse le questionnaire et extrais TOUTES les questions en JSON compact.

FORMAT JSON COMPACT OBLIGATOIRE (sans indentation):
{"title":"titre","questions":[{"id":"q1","num":1,"label":"libelle complet de la question","question_class":"CLASS","type":"TYPE","required":true,"hint":"","choices":["choix1","choix2"],"group":"groupe ou null","formats":[FORMAT_ARRAY],"suggested_format_idx":0,"suggestions":[]}],"groups":[]}

QUESTION_CLASS et TYPE et FORMATS:

quantitative (age,poids,revenu,score,quantite,mesure...): type=integer ou decimal
formats=[{"id":"A","name":"Nombre entier","type":"integer","note":"Ex: 25, 150 sans virgule"},{"id":"B","name":"Nombre avec virgule","type":"decimal","note":"Ex: 65,5 kg ou 37,8 degres"},{"id":"C","name":"Valeur sur une echelle","type":"range","note":"Ex: satisfaction de 1 a 10"}]

qualitative_choice (sexe,niveau,statut,oui/non,categories avec options): type=select_one
IMPORTANT: choices[] doit contenir les vraies options extraites du document
formats=[{"id":"A","name":"Une seule reponse au choix","type":"select_one","note":"Le repondant coche une seule case"},{"id":"B","name":"Plusieurs reponses possibles","type":"select_multiple","note":"Le repondant peut cocher plusieurs cases"},{"id":"C","name":"Reponse ecrite libre","type":"text","note":"Le repondant ecrit lui-meme"}]

qualitative_open (nom,commentaire,description,adresse,opinion): type=text
formats=[{"id":"A","name":"Reponse courte","type":"text","note":"Quelques mots ex: nom, profession"},{"id":"B","name":"Reponse longue","type":"text","note":"Plusieurs phrases ex: commentaire"}]

date_time: type=date
formats=[{"id":"A","name":"Date jour/mois/annee","type":"date","note":"Ex: 15/03/2024"},{"id":"B","name":"Heure","type":"time","note":"Ex: 14h30"},{"id":"C","name":"Date et heure ensemble","type":"datetime","note":"Ex: 15/03/2024 14h30"}]

geopoint: formats=[{"id":"A","name":"Localisation GPS un point precis","type":"geopoint","note":"Capture GPS automatique"}]
geotrace: formats=[{"id":"A","name":"Tracer un chemin GPS","type":"geotrace","note":"Trajet sur carte"}]
geoshape: formats=[{"id":"A","name":"Delimiter une zone GPS","type":"geoshape","note":"Perimetre sur carte"}]
media_photo: formats=[{"id":"A","name":"Prendre une photo","type":"image","note":"Photo avec appareil"}]
media_audio: formats=[{"id":"A","name":"Enregistrer un son","type":"audio","note":"Enregistrement audio"}]
media_video: formats=[{"id":"A","name":"Enregistrer une video","type":"video","note":"Enregistrement video"}]
media_file: formats=[{"id":"A","name":"Joindre un fichier","type":"file","note":"PDF, Excel ou autre"}]
barcode: formats=[{"id":"A","name":"Scanner un code-barres","type":"barcode","note":"QR code ou code-barres"}]
acknowledge: formats=[{"id":"A","name":"Case a cocher pour confirmer","type":"acknowledge","note":"Ex: J accepte les conditions"}]
ranking: formats=[{"id":"A","name":"Classer par ordre de preference","type":"rank","note":"Du plus au moins important"}]
scale: formats=[{"id":"A","name":"Valeur sur une echelle","type":"range","note":"Ex: douleur de 0 a 10"}]
calculate: formats=[{"id":"A","name":"Valeur calculee automatiquement","type":"calculate","note":"Calcul a partir d autres reponses"}]
note: formats=[{"id":"A","name":"Message information","type":"note","note":"Texte affiche sans saisie"}]

SUGGESTIONS (seulement si vraiment detecte dans la logique du questionnaire):
Format: {"type":"skip_logic","label":"libelle court","description":"explication claire sans jargon","value":"formule XLSForm relevant","confidence":"high|medium|low"}
{"type":"calculate","label":"libelle court","description":"explication","value":"formule XLSForm calculation","confidence":"high|medium|low"}
{"type":"constraint","label":"libelle court","description":"explication","value":"formule XLSForm constraint","confidence":"high|medium|low"}

REGLES IMPORTANTES:
- required=true par defaut pour toutes les questions
- Extrais TOUTES les questions sans exception
- Pour qualitative_choice: extrais les vraies options dans choices[]
- suggested_format_idx = index du format le plus adapte (0, 1 ou 2)
- JSON compact sans indentation obligatoire
- Outil cible: ${tool}`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 4096,
        system,
        messages: [{ role: 'user', content: `Extrais toutes les questions de ce questionnaire:\n\n${inputText}` }]
      })
    });

    if (!response.ok) {
      const e = await response.json();
      console.error('[CLAUDE ERROR]', e);
      return res.status(502).json({ error: 'CLAUDE_ERROR', message: 'Erreur API Claude.' });
    }

    const data = await response.json();
    const rawText = data.content?.[0]?.text || '{}';

    let form;
    try {
      let jsonStr = rawText.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim();
      const s = jsonStr.indexOf('{');
      if (s > 0) jsonStr = jsonStr.slice(s);
      try { form = JSON.parse(jsonStr); }
      catch(e) {
        const lc = jsonStr.lastIndexOf('},');
        if (lc > 100) jsonStr = jsonStr.slice(0, lc+1);
        jsonStr = jsonStr.replace(/,\s*$/,'');
        const ob = (jsonStr.match(/\[/g)||[]).length-(jsonStr.match(/\]/g)||[]).length;
        const cb = (jsonStr.match(/\{/g)||[]).length-(jsonStr.match(/\}/g)||[]).length;
        for(let i=0;i<Math.max(0,ob);i++) jsonStr+=']';
        for(let i=0;i<Math.max(0,cb);i++) jsonStr+='}';
        form = JSON.parse(jsonStr);
        console.log('[PARSE] JSON repare');
      }
    } catch(pe) {
      console.error('[PARSE ERROR]', pe.message);
      return res.status(502).json({ error: 'PARSE_ERROR', message: 'Reponse invalide.' });
    }

    if (!form.questions || !Array.isArray(form.questions)) {
      return res.status(422).json({ error: 'NO_QUESTIONS', message: 'Aucune question detectee.' });
    }

    console.log(`[ANALYSE] ✓ ${form.questions.length} questions`);
    res.json({
      success: true, form, truncated,
      warning: truncated ? `Questionnaire long (${text.length} caracteres). Seuls les ${inputText.length} premiers caracteres ont ete analyses. Verifiez que toutes vos questions sont presentes.` : null
    });

  } catch(err) {
    console.error('[ANALYSE ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ CORRECTION ============
app.post('/api/correct', async (req, res) => {
  try {
    const { form, instructions } = req.body;
    if (!form || !instructions) return res.status(400).json({ error: 'Donnees manquantes' });
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'Cle API manquante' });

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5', max_tokens: 4096,
        system: 'Expert collecte de donnees. Applique exactement les corrections au formulaire JSON. Retourne JSON corrige UNIQUEMENT sans markdown.',
        messages: [{ role: 'user', content: `Formulaire:\n${JSON.stringify(form)}\n\nCorrections:\n${instructions}` }]
      })
    });

    if (!response.ok) return res.status(502).json({ error: 'CLAUDE_ERROR' });
    const data = await response.json();
    let raw = data.content?.[0]?.text||'{}';
    raw = raw.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim();
    const m = raw.match(/\{[\s\S]*/);
    const corrected = JSON.parse(m ? m[0] : '{}');
    res.json({ success: true, form: corrected });
  } catch(err) {
    res.status(500).json({ error: 'SERVER_ERROR' });
  }
});

// ============ DEPLOIEMENT KOBO ============
app.post('/api/deploy/kobo', async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { username, password, server = 'https://kf.kobotoolbox.org' } = credentials;
    if (!form || !username || !password) return res.status(400).json({ error: 'Donnees manquantes' });

    console.log(`[DEPLOY] KoboToolbox → ${server}`);

    const tokenRes = await fetch(`${server}/token/?format=json`, {
      headers: { 'Authorization': 'Basic ' + Buffer.from(`${username}:${password}`).toString('base64') }
    });
    if (!tokenRes.ok) return res.status(401).json({ error: 'AUTH_ERROR', message: 'Identifiants incorrects.' });
    const { token } = await tokenRes.json();
    const auth = { 'Authorization': 'Token ' + token, 'Content-Type': 'application/json' };

    const assetRes = await fetch(`${server}/api/v2/assets/?format=json`, {
      method: 'POST', headers: auth,
      body: JSON.stringify({ name: form.title || 'Formulaire R2', asset_type: 'survey' })
    });
    if (!assetRes.ok) return res.status(502).json({ error: 'ASSET_ERROR', message: 'Erreur creation formulaire.' });
    const { uid } = await assetRes.json();

    const koboContent = buildKoboContent(form);
    const patchRes = await fetch(`${server}/api/v2/assets/${uid}/?format=json`, {
      method: 'PATCH', headers: auth,
      body: JSON.stringify({ name: form.title || 'Formulaire R2', content: koboContent })
    });
    if (!patchRes.ok) {
      const e = await patchRes.text();
      console.error('[PATCH ERROR]', e);
      return res.status(502).json({ error: 'PATCH_ERROR', message: 'Erreur import questionnaire.' });
    }

    await fetch(`${server}/api/v2/assets/${uid}/deployment/?format=json`, {
      method: 'POST', headers: auth, body: JSON.stringify({ active: true })
    });

    console.log(`[DEPLOY] ✓ ${uid} — ${form.questions?.length} questions`);
    res.json({ success: true, uid, url: `${server}/#/forms/${uid}/summary`, questions: form.questions?.length || 0 });
  } catch(err) {
    console.error('[DEPLOY ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ BUILDER KOBOTOOLBOX ============
function buildKoboContent(form) {
  const survey = [];
  const choices = [];
  const seen = new Set();

  // Grouper les questions
  const groups = {};
  (form.questions || []).forEach(q => {
    const g = q.group || 'general';
    if (!groups[g]) groups[g] = [];
    groups[g].push(q);
  });

  Object.entries(groups).forEach(([gname, qs]) => {
    const gId = gname.replace(/\s+/g,'_').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9_]/g,'').slice(0,32);

    if (gname !== 'general') {
      survey.push({ type: 'begin_group', name: gId, label: gname });
    }

    qs.forEach(q => {
      // Determine le type final selon le choix de l'utilisateur
      const fmtIdx = q.validatedFormatIdx !== undefined ? q.validatedFormatIdx : (q.suggested_format_idx || 0);
      const selectedFmt = q.formats?.[fmtIdx];
      const t = q.selectedType || selectedFmt?.type || q.type || 'text';
      const name = (q.id || ('q'+q.num)).replace(/[^a-zA-Z0-9_]/g,'_');

      const row = {
        type: t,
        name,
        label: q.label || '',
        required: q.required !== false ? 'yes' : 'no',
        hint: q.hint || ''
      };

      // Logique de saut
      if (q.relevant && q.relevant.trim()) row.relevant = q.relevant.trim();

      // Calcul
      if (t === 'calculate' && q.calculation) {
        row.calculation = q.calculation;
        delete row.required;
      }

      // Contrainte (min/max + chiffres)
      const constraints = [];
      if (q.numMin !== '' && q.numMin !== undefined && q.numMin !== null) constraints.push(`. >= ${q.numMin}`);
      if (q.numMax !== '' && q.numMax !== undefined && q.numMax !== null) constraints.push(`. <= ${q.numMax}`);
      if (q.numDigitsBefore) constraints.push(`string-length(substring-before(string(.), '.')) <= ${q.numDigitsBefore}`);
      if (q.numDigitsAfter) constraints.push(`string-length(substring-after(string(.), '.')) <= ${q.numDigitsAfter}`);
      if (q.constraint && q.constraint.trim()) constraints.push(q.constraint.trim());
      if (constraints.length > 0) {
        row.constraint = constraints.join(' and ');
        row.constraint_message = 'Valeur hors des limites acceptees';
      }

      // Range
      if (t === 'range') {
        row.parameters = `start=${q.numMin||1} end=${q.numMax||10}`;
        delete row.required;
      }

      // Note / calculate — pas de required
      if (['note','calculate'].includes(t)) delete row.required;

      // Choix
      if (['select_one','select_multiple','rank'].includes(t)) {
        const listName = 'list_' + name;
        row.type = t + ' ' + listName;
        if (!seen.has(listName)) {
          seen.add(listName);
          (q.choices || []).forEach((c, i) => {
            const label = typeof c === 'string' ? c : (c.label || String(c));
            const val = label.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9_]/g,'_').replace(/__+/g,'_').replace(/^_|_$/g,'').slice(0,30) || ('c'+(i+1));
            choices.push({ list_name: listName, name: val, label });
          });
        }
      }

      survey.push(row);
    });

    if (gname !== 'general') {
      survey.push({ type: 'end_group', name: gId });
    }
  });

  return {
    survey, choices,
    settings: [{
      form_title: form.title || 'Formulaire',
      form_id: (form.title||'formulaire').replace(/\s+/g,'_').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9_]/g,'').slice(0,32),
      version: '1'
    }]
  };
}

app.listen(PORT, () => {
  console.log(`\n╔══════════════════════════════════╗\n║   R2 Forms — Backend v3.0        ║\n║   Port : ${PORT}                     ║\n╚══════════════════════════════════╝\n`);
});
