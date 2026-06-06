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

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }
});

app.use(express.static(path.join(__dirname, 'public')));

app.get('/health', (req, res) => {
  res.json({ status: 'ok', version: '2.0.0', timestamp: new Date().toISOString() });
});

// ============ IMPORT FICHIER ============
app.post('/api/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier reçu' });

    const { originalname, buffer } = req.file;
    const ext = originalname.split('.').pop().toLowerCase();
    let extractedText = '';

    console.log(`[IMPORT] ${originalname} (${(buffer.length/1024).toFixed(0)} KB)`);

    if (['txt', 'csv'].includes(ext)) {
      extractedText = buffer.toString('utf-8');
    }
    else if (ext === 'pdf') {
      const pdfData = await pdf(buffer);
      extractedText = pdfData.text;
      if (!extractedText || extractedText.trim().length < 20) {
        return res.status(422).json({ error: 'PDF_SCANNED', message: 'PDF scanné illisible. Collez le texte directement.' });
      }
    }
    else if (ext === 'docx') {
      const result = await mammoth.extractRawText({ buffer });
      extractedText = result.value;
    }
    else if (ext === 'doc') {
      extractedText = buffer.toString('latin1').replace(/[^\x20-\x7E\n\r\u00C0-\u024F]/g, ' ').replace(/\s+/g, ' ').trim();
      if (extractedText.length < 30) return res.status(422).json({ error: 'DOC_OLD', message: 'Format .doc non supporté. Enregistrez en .docx.' });
    }
    else if (['xlsx', 'xls'].includes(ext)) {
      const workbook = XLSX.read(buffer, { type: 'buffer' });
      let allText = '';
      workbook.SheetNames.forEach(name => {
        allText += `\n=== ${name} ===\n` + XLSX.utils.sheet_to_csv(workbook.Sheets[name]);
      });
      extractedText = allText.trim();
    }
    else {
      return res.status(400).json({ error: 'FORMAT_UNSUPPORTED', message: `Format .${ext} non supporté.` });
    }

    extractedText = extractedText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n{3,}/g, '\n\n').trim();

    if (extractedText.length < 20) {
      return res.status(422).json({ error: 'EMPTY', message: 'Fichier vide ou illisible. Collez le texte directement.' });
    }

    console.log(`[IMPORT] ✓ ${extractedText.length} caractères`);
    res.json({ success: true, text: extractedText, metadata: { filename: originalname, chars: extractedText.length } });

  } catch (err) {
    console.error('[IMPORT ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: 'Erreur import: ' + err.message });
  }
});

// ============ ANALYSE CLAUDE ============
app.post('/api/analyse', async (req, res) => {
  try {
    const { text, tool } = req.body;
    if (!text || text.trim().length < 10) return res.status(400).json({ error: 'Texte trop court' });

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'Clé API manquante' });

    // Limiter le texte pour économiser la mémoire
    const inputText = text.slice(0, 8000);
    console.log(`[ANALYSE] ${text.length} chars (envoi: ${inputText.length}) → ${tool}`);

    const system = `Expert collecte de donnees. Extrais toutes les questions du questionnaire en JSON compact.

FORMAT (JSON compact, sans indentation):
{"title":"titre","questions":[{"id":"q1","num":1,"label":"libelle complet","question_class":"quantitative|qualitative_choice|qualitative_open|date_time|geopoint|geotrace|geoshape|media_photo|media_audio|media_video|media_file|barcode|acknowledge|ranking|scale|calculate|note","type":"integer|decimal|text|select_one|select_multiple|date|time|datetime|geopoint|geotrace|geoshape|image|audio|video|file|barcode|acknowledge|rank|range|calculate|note","required":true,"hint":"","choices":[],"group":"groupe ou null","formats":[],"suggested_format_idx":0,"suggestions":[]}],"groups":[]}

FORMATS PAR CLASSE:
quantitative:[{"id":"A","name":"Nombre entier","type":"integer","note":"Ex: 25"},{"id":"B","name":"Nombre avec virgule","type":"decimal","note":"Ex: 65,5"},{"id":"C","name":"Valeur sur une echelle","type":"range","note":"Ex: 1 a 10"}]
qualitative_choice:[{"id":"A","name":"Une seule reponse au choix","type":"select_one","note":"Une case"},{"id":"B","name":"Plusieurs reponses possibles","type":"select_multiple","note":"Plusieurs cases"},{"id":"C","name":"Reponse ecrite libre","type":"text","note":"Libre"}]
qualitative_open:[{"id":"A","name":"Reponse courte","type":"text","note":"Quelques mots"},{"id":"B","name":"Reponse longue","type":"text","note":"Plusieurs phrases"}]
date_time:[{"id":"A","name":"Date","type":"date","note":"jj/mm/aaaa"},{"id":"B","name":"Heure","type":"time","note":"hh:mm"},{"id":"C","name":"Date et heure","type":"datetime","note":"jj/mm/aaaa hh:mm"}]
geopoint:[{"id":"A","name":"Localisation GPS","type":"geopoint","note":"Point GPS"}]
geotrace:[{"id":"A","name":"Tracer un chemin GPS","type":"geotrace","note":"Trajet"}]
geoshape:[{"id":"A","name":"Delimiter une zone GPS","type":"geoshape","note":"Zone"}]
media_photo:[{"id":"A","name":"Prendre une photo","type":"image","note":"Photo"}]
media_audio:[{"id":"A","name":"Enregistrer un son","type":"audio","note":"Audio"}]
media_video:[{"id":"A","name":"Enregistrer une video","type":"video","note":"Video"}]
media_file:[{"id":"A","name":"Joindre un fichier","type":"file","note":"Fichier"}]
barcode:[{"id":"A","name":"Scanner un code-barres","type":"barcode","note":"QR/barcode"}]
acknowledge:[{"id":"A","name":"Case a cocher pour confirmer","type":"acknowledge","note":"Confirmation"}]
ranking:[{"id":"A","name":"Classer par ordre de preference","type":"rank","note":"Ordre"}]
scale:[{"id":"A","name":"Valeur sur une echelle","type":"range","note":"Curseur"}]
calculate:[{"id":"A","name":"Valeur calculee automatiquement","type":"calculate","note":"Calcul"}]
note:[{"id":"A","name":"Message information","type":"note","note":"Sans saisie"}]

SUGGESTIONS (seulement si detecte dans le doc):
{"type":"skip_logic|calculate|constraint","label":"court","description":"clair","value":"formule XLSForm","confidence":"high|medium|low"}

REGLES: required=true par defaut. Extrais TOUTES questions. choices[] pour qualitative_choice. JSON compact. Outil: ${tool}`;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 4096,
        system,
        messages: [{ role: 'user', content: `Extrais toutes les questions:\n\n${inputText}` }]
      })
    });

    if (!response.ok) {
      const errData = await response.json();
      console.error('[CLAUDE ERROR]', errData);
      return res.status(502).json({ error: 'CLAUDE_ERROR', message: 'Erreur API Claude.' });
    }

    const data = await response.json();
    const rawText = data.content?.[0]?.text || '{}';

    let form;
    try {
      let jsonStr = rawText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      const jsonStart = jsonStr.indexOf('{');
      if (jsonStart > 0) jsonStr = jsonStr.slice(jsonStart);

      try {
        form = JSON.parse(jsonStr);
      } catch(e) {
        // Réparer JSON tronqué
        const lastComplete = jsonStr.lastIndexOf('},');
        if (lastComplete > 100) jsonStr = jsonStr.slice(0, lastComplete + 1);
        jsonStr = jsonStr.replace(/,\s*$/, '');
        const opens = (jsonStr.match(/\[/g)||[]).length - (jsonStr.match(/\]/g)||[]).length;
        const openb = (jsonStr.match(/\{/g)||[]).length - (jsonStr.match(/\}/g)||[]).length;
        for(let i = 0; i < Math.max(0,opens); i++) jsonStr += ']';
        for(let i = 0; i < Math.max(0,openb); i++) jsonStr += '}';
        form = JSON.parse(jsonStr);
        console.log('[PARSE] JSON réparé');
      }
    } catch (parseErr) {
      console.error('[PARSE ERROR]', parseErr.message);
      return res.status(502).json({ error: 'PARSE_ERROR', message: 'Réponse invalide.' });
    }

    if (!form.questions || !Array.isArray(form.questions)) {
      return res.status(422).json({ error: 'NO_QUESTIONS', message: 'Aucune question détectée.' });
    }

    const truncated = text.length > 8000;
    console.log(`[ANALYSE] ✓ ${form.questions.length} questions${truncated ? ' (questionnaire tronqué à 8000 chars)' : ''}`);
    
    res.json({ 
      success: true, 
      form,
      truncated,
      warning: truncated ? `Votre questionnaire est long (${text.length} caractères). Seules les ${inputText.length} premiers caractères ont été analysés. Les questions restantes peuvent être ajoutées manuellement.` : null
    });

  } catch (err) {
    console.error('[ANALYSE ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: 'Erreur analyse.' });
  }
});

// ============ CORRECTION CLAUDE ============
app.post('/api/correct', async (req, res) => {
  try {
    const { form, instructions } = req.body;
    if (!form || !instructions) return res.status(400).json({ error: 'Données manquantes' });

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'Clé API manquante' });

    console.log(`[CORRECT] ${instructions.slice(0, 80)}`);

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 4096,
        system: 'Expert collecte de donnees. Applique les corrections au formulaire JSON et retourne JSON corrigé UNIQUEMENT, sans markdown.',
        messages: [{ role: 'user', content: `Formulaire:\n${JSON.stringify(form)}\n\nCorrections:\n${instructions}` }]
      })
    });

    if (!response.ok) return res.status(502).json({ error: 'CLAUDE_ERROR' });

    const data = await response.json();
    let rawText = data.content?.[0]?.text || '{}';
    rawText = rawText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const match = rawText.match(/\{[\s\S]*/);
    const corrected = JSON.parse(match ? match[0] : '{}');

    console.log(`[CORRECT] ✓ ${corrected.questions?.length} questions`);
    res.json({ success: true, form: corrected });

  } catch (err) {
    console.error('[CORRECT ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR' });
  }
});

// ============ DÉPLOIEMENT KOBOTOOLBOX ============
app.post('/api/deploy/kobo', async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { username, password, server = 'https://kf.kobotoolbox.org' } = credentials;
    if (!form || !username || !password) return res.status(400).json({ error: 'Données manquantes' });

    console.log(`[DEPLOY] KoboToolbox → ${server}`);

    const tokenRes = await fetch(`${server}/token/?format=json`, {
      method: 'GET',
      headers: { 'Authorization': 'Basic ' + Buffer.from(`${username}:${password}`).toString('base64') }
    });
    if (!tokenRes.ok) return res.status(401).json({ error: 'AUTH_ERROR', message: 'Identifiants incorrects.' });

    const { token } = await tokenRes.json();
    const auth = { 'Authorization': 'Token ' + token, 'Content-Type': 'application/json' };

    const assetRes = await fetch(`${server}/api/v2/assets/?format=json`, {
      method: 'POST', headers: auth,
      body: JSON.stringify({ name: form.title || 'Formulaire R2', asset_type: 'survey' })
    });
    if (!assetRes.ok) return res.status(502).json({ error: 'ASSET_ERROR', message: 'Erreur création formulaire.' });

    const { uid: assetUid } = await assetRes.json();
    const koboContent = buildKoboContent(form);

    const patchRes = await fetch(`${server}/api/v2/assets/${assetUid}/?format=json`, {
      method: 'PATCH', headers: auth,
      body: JSON.stringify({ name: form.title || 'Formulaire R2', content: koboContent })
    });
    if (!patchRes.ok) {
      const errBody = await patchRes.text();
      console.error('[PATCH ERROR]', errBody);
      return res.status(502).json({ error: 'PATCH_ERROR', message: 'Erreur import questionnaire.' });
    }

    await fetch(`${server}/api/v2/assets/${assetUid}/deployment/?format=json`, {
      method: 'POST', headers: auth,
      body: JSON.stringify({ active: true })
    });

    console.log(`[DEPLOY] ✓ ${assetUid}`);
    res.json({ success: true, uid: assetUid, url: `${server}/#/forms/${assetUid}/summary`, questions: form.questions?.length || 0 });

  } catch (err) {
    console.error('[DEPLOY ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ BUILDER KOBOTOOLBOX ============
function buildKoboContent(form) {
  const survey = [];
  const choices = [];
  const seen = new Set();

  const groups = {};
  (form.questions || []).forEach(q => {
    const g = q.group || 'general';
    if (!groups[g]) groups[g] = [];
    groups[g].push(q);
  });

  Object.entries(groups).forEach(([gname, qs]) => {
    if (gname !== 'general') {
      survey.push({ type: 'begin_group', name: gname.replace(/\s+/g,'_').toLowerCase().replace(/[^a-z0-9_]/g,''), label: gname });
    }

    qs.forEach(q => {
      const t = q.selectedType || q.type || 'text';
      const name = (q.id || ('q'+q.num)).replace(/[^a-zA-Z0-9_]/g,'_');
      const row = { type: t, name, label: q.label || '', required: q.required !== false ? 'yes' : 'no', hint: q.hint || '' };

      if (q.relevant && q.relevant.trim()) row.relevant = q.relevant.trim();
      if (t === 'calculate' && q.calculation) { row.calculation = q.calculation; delete row.required; }
      if (q.constraint && q.constraint.trim()) { row.constraint = q.constraint.trim(); row.constraint_message = 'Valeur hors limites'; }
      if (t === 'range' && (q.numMin || q.numMax)) row.parameters = `start=${q.numMin||1} end=${q.numMax||10}`;
      if (['note','calculate'].includes(t)) delete row.required;

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
      survey.push({ type: 'end_group', name: gname.replace(/\s+/g,'_').toLowerCase().replace(/[^a-z0-9_]/g,'') });
    }
  });

  return {
    survey,
    choices,
    settings: [{ form_title: form.title || 'Formulaire', form_id: (form.title||'formulaire').replace(/\s+/g,'_').toLowerCase().replace(/[^a-z0-9_]/g,'').slice(0,32), version: '1' }]
  };
}

app.listen(PORT, () => {
  console.log(`
╔══════════════════════════════════╗
║   R2 Forms — Backend Server      ║
║   Port : ${PORT}                     ║
║   Import : PDF, Word, Excel...   ║
║   Analyse : Claude Sonnet        ║
║   Deploy : KoboToolbox API v2    ║
╚══════════════════════════════════╝
  `);
});
