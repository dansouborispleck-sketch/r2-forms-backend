// ============================================
// R2 Forms Backend v4.1.1
// Corrigé : validation des choices, déploiement multi-outils
// ============================================

require('dotenv').config();

const express = require('express');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const helmet = require('helmet');
const multer = require('multer');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const pdf = require('pdf-parse');
const fetch = require('node-fetch');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 4000;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;

// ============ SÉCURITÉ ============

app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com"],
      fontSrc: ["'self'", "https://fonts.gstatic.com"],
      scriptSrc: ["'self'", "'unsafe-inline'"],
      imgSrc: ["'self'", "data:", "https:"],
      connectSrc: ["'self'", "https://api.anthropic.com"]
    }
  }
}));

app.use(cors({
  origin: 'https://dansouborispleck-sketch.github.io'
}));

const limiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 20,
  message: { error: 'Trop de requêtes. Attendez 15 minutes.' }
});

const analyseLimiter = rateLimit({
  windowMs: 60 * 60 * 1000,
  max: 10,
  message: { error: 'Limite d\'analyses atteinte (10/heure).' }
});

app.use('/api/', limiter);
app.use('/api/analyse', analyseLimiter);
app.use('/api/correct', analyseLimiter);
app.use('/api/verify', analyseLimiter);

app.use(express.json({ limit: '2mb' }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = ['txt','csv','pdf','docx','doc','xlsx','xls','odt'];
    const ext = file.originalname.split('.').pop().toLowerCase();
    if (ok.includes(ext)) cb(null, true);
    else cb(new Error('Format non autorisé'), false);
  }
});

// ============ PROMPT CLAUDE ============

const SYSTEM_PROMPT = `Expert collecte données terrain. Extrais structure questionnaire en JSON compact.

REGLES:
1. TOUTES questions: numérotées, sans numero, sous-questions
2. GROUPES: titres complets de sections
3. "Autres" → question libre auto après (required=false, label="Si autre, précisez :")
4. SAUTS: detecter conditions implicites
5. choice_values: codes numeriques [0,1,2...]

FORMAT JSON:
{"title":"...","coherence_report":[],"questions":[{"id":"q1","num":1,"label":"...","question_class":"qualitative_choice","type":"select_one","required":true,"hint":"","choices":["Masculin","Féminin"],"choice_values":["0","1"],"group":"TITRE","formats":[{"id":"A","name":"Choix unique","type":"select_one","note":"Une seule réponse"},{"id":"B","name":"Choix multiple","type":"select_multiple","note":"Plusieurs réponses"},{"id":"C","name":"Texte libre","type":"text","note":"Réponse libre"}],"suggested_format_idx":0,"suggestions":[]}],"groups":[]}

CLASSES: quantitative,qualitative_choice,qualitative_open,date_time,geopoint,geotrace,geoshape,media_photo,media_audio,media_video,media_file,barcode,acknowledge,ranking,scale,calculate,note

IMPORTANT: Pour les questions à choix (qualitative_choice), choices DOIT contenir les vraies options: ["Masculin","Féminin"] et choice_values: ["0","1"].

SUGGESTIONS: {"type":"skip_logic|calculate|constraint","label":"court","description":"clair","value":"formule XLSForm","confidence":"high|medium|low"}

required=true par defaut. JSON compact.`;

// ============ FONCTIONS AUXILIAIRES ============

function parseJSON(text) {
  let clean = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
  const start = clean.indexOf('{');
  if (start > 0) clean = clean.slice(start);
  
  try {
    return JSON.parse(clean);
  } catch (e) {
    const openB = (clean.match(/\[/g) || []).length - (clean.match(/\]/g) || []).length;
    const openC = (clean.match(/\{/g) || []).length - (clean.match(/\}/g) || []).length;
    clean = clean.replace(/,\s*$/, '');
    for (let i = 0; i < Math.max(0, openB); i++) clean += ']';
    for (let i = 0; i < Math.max(0, openC); i++) clean += '}';
    return JSON.parse(clean);
  }
}

// Validation et normalisation du formulaire
function normalizeForm(form) {
  if (!form || !form.questions) return form;
  
  form.questions.forEach((q, idx) => {
    // Assurer que num existe
    if (!q.num) q.num = idx + 1;
    if (!q.id) q.id = 'q' + q.num;
    
    // Assurer que choices est un tableau de strings
    if (q.choices && Array.isArray(q.choices)) {
      q.choices = q.choices.map(c => {
        if (typeof c === 'string') return c;
        if (c && typeof c === 'object' && c.label) return c.label;
        return String(c);
      }).filter(c => c && c.trim());
    }
    
    // Assurer que choice_values existe
    if (!q.choice_values || !Array.isArray(q.choice_values)) {
      if (q.choices && q.choices.length > 0) {
        q.choice_values = q.choices.map((_, i) => String(i));
      }
    }
    
    // Assurer que formats existe pour les questions à choix
    if (q.question_class === 'qualitative_choice' || (q.choices && q.choices.length > 0)) {
      if (!q.formats || q.formats.length === 0) {
        q.formats = [
          { id: 'A', name: 'Choix unique', type: 'select_one', note: 'Une seule réponse possible' },
          { id: 'B', name: 'Choix multiple', type: 'select_multiple', note: 'Plusieurs réponses possibles' },
          { id: 'C', name: 'Texte libre', type: 'text', note: 'Réponse libre' }
        ];
      }
    }
    
    // Assurer question_class
    if (!q.question_class) {
      if (q.choices && q.choices.length > 0) q.question_class = 'qualitative_choice';
      else if (q.type === 'integer' || q.type === 'decimal') q.question_class = 'quantitative';
      else if (q.type === 'date' || q.type === 'time') q.question_class = 'date_time';
      else if (q.type === 'geopoint') q.question_class = 'geopoint';
      else if (q.type === 'image') q.question_class = 'media_photo';
      else q.question_class = 'qualitative_open';
    }
    
    // Assurer suggested_format_idx
    if (q.suggested_format_idx === undefined || q.suggested_format_idx === null) {
      q.suggested_format_idx = 0;
    }
  });
  
  return form;
}

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
    const gId = gname.replace(/\s+/g, '_').toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9_]/g, '').slice(0, 32) || 'group';

    if (gname !== 'general') survey.push({ type: 'begin_group', name: gId, label: gname });

    qs.forEach(q => {
      const fmtIdx = q.validatedFormatIdx !== undefined ? q.validatedFormatIdx : (q.suggested_format_idx || 0);
      const selectedFmt = q.formats?.[fmtIdx] || q.formats?.[0];
      let t = q.selectedType || selectedFmt?.type || q.type || 'text';
      const name = (q.id || ('q' + q.num)).replace(/[^a-zA-Z0-9_]/g, '_');

      let row = {
        type: t,
        name: name,
        label: q.label || '',
        required: q.required !== false ? 'yes' : 'no',
        hint: q.hint || ''
      };

      let relevant = q.relevant || '';
      if (!relevant && q.suggestions) {
        q.suggestions.forEach(s => {
          if (s.type === 'skip_logic' && s.value && q.confirmedSuggestions?.[s._idx]) {
            relevant = s.value;
          }
        });
      }
      if (relevant) row.relevant = relevant;

      if (t === 'calculate') {
        let calc = q.calculation || '';
        if (!calc && q.suggestions) {
          q.suggestions.forEach(s => {
            if (s.type === 'calculate' && s.value && q.confirmedSuggestions?.[s._idx]) {
              calc = s.value;
            }
          });
        }
        if (calc) row.calculation = calc;
        delete row.required;
      }

      const constraints = [];
      if (q.numMin !== '' && q.numMin != null) constraints.push('. >= ' + q.numMin);
      if (q.numMax !== '' && q.numMax != null) constraints.push('. <= ' + q.numMax);
      if (q.constraint?.trim()) constraints.push(q.constraint.trim());
      if (constraints.length > 0) {
        row.constraint = constraints.join(' and ');
        row.constraint_message = 'Valeur hors limites';
      }

      if (t === 'range') {
        row.parameters = 'start=' + (q.numMin || 1) + ' end=' + (q.numMax || 10);
        delete row.required;
      }

      if (t === 'note' || t === 'calculate') delete row.required;

      if (t === 'select_one' || t === 'select_multiple' || t === 'rank') {
        const listName = 'list_' + name;
        row.type = t + ' ' + listName;
        if (!seen.has(listName)) {
          seen.add(listName);
          const vals = q.choice_values || [];
          const chs = q.choices || [];
          chs.forEach((c, i) => {
            const label = typeof c === 'string' ? c : (c.label || String(c));
            const val = (vals[i] !== undefined && vals[i] !== '') ? String(vals[i]) : String(i);
            choices.push({ list_name: listName, name: val, label: label });
          });
        }
      }

      survey.push(row);
    });

    if (gname !== 'general') survey.push({ type: 'end_group', name: gId });
  });

  const formId = (form.title || 'formulaire').replace(/\s+/g, '_').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9_]/g, '').slice(0, 32);

  return {
    survey: survey,
    choices: choices,
    settings: [{ form_title: form.title || 'Formulaire', form_id: formId, version: '1' }]
  };
}

// ============ ROUTES ============

app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    version: '4.1.1',
    anthropic: ANTHROPIC_API_KEY ? 'ok' : 'missing'
  });
});

// ============ IMPORT ============
app.post('/api/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier recu' });

    const { originalname, buffer } = req.file;
    const ext = originalname.split('.').pop().toLowerCase();
    let text = '';

    console.log('[IMPORT] ' + originalname);

    if (['txt', 'csv'].includes(ext)) {
      text = buffer.toString('utf-8');
    } else if (ext === 'pdf') {
      const d = await pdf(buffer);
      text = d.text;
      if (!text || text.trim().length < 20) {
        return res.status(422).json({ error: 'PDF_SCANNED', message: 'PDF scanné illisible. Collez le texte.' });
      }
    } else if (ext === 'docx') {
      const r = await mammoth.extractRawText({ buffer });
      text = r.value;
    } else if (ext === 'doc') {
      text = buffer.toString('latin1').replace(/[^\x20-\x7E\n\r\u00C0-\u024F]/g, ' ').replace(/\s+/g, ' ').trim();
      if (text.length < 30) {
        return res.status(422).json({ error: 'DOC_OLD', message: 'Format .doc non supporté. Enregistrez en .docx.' });
      }
    } else if (['xlsx', 'xls'].includes(ext)) {
      const wb = XLSX.read(buffer, { type: 'buffer' });
      wb.SheetNames.forEach(n => {
        text += '\n=== ' + n + ' ===\n' + XLSX.utils.sheet_to_csv(wb.Sheets[n]);
      });
    } else if (ext === 'odt') {
      const raw = buffer.toString('utf-8');
      const matches = raw.match(/<<text:p[^>]*>([^<<]{2,})<<\/text:p>/g) || [];
      text = matches.map(m => m.replace(/<<[^>]+>/g, '').trim()).filter(t => t.length > 1).join('\n');
    } else {
      return res.status(400).json({ error: 'FORMAT_UNSUPPORTED', message: 'Format .' + ext + ' non supporté.' });
    }

    text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n{3,}/g, '\n\n').trim();

    if (text.length < 20) {
      return res.status(422).json({ error: 'EMPTY', message: 'Fichier vide. Collez le texte.' });
    }

    console.log('[IMPORT] OK - ' + text.length + ' caractères');
    res.json({ success: true, text: text, metadata: { filename: originalname, chars: text.length } });

  } catch (err) {
    console.error('[IMPORT ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ ANALYSE ============
app.post('/api/analyse', analyseLimiter, async (req, res) => {
  try {
    const { text, tool } = req.body;
    if (!text || text.trim().length < 10) {
      return res.status(400).json({ error: 'Texte trop court' });
    }
    if (!ANTHROPIC_API_KEY) {
      return res.status(500).json({ error: 'Cle API Anthropic manquante' });
    }

    const MAX = 6000;
    let inputText = text;
    if (text.length > MAX) {
      inputText = text.slice(0, Math.floor(MAX * 0.7)) + '\n\n[...]\n\n' + text.slice(-Math.floor(MAX * 0.3));
    }

    console.log('[ANALYSE] ' + text.length + ' chars → ' + tool);

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 24000,
        system: SYSTEM_PROMPT,
        messages: [{ role: 'user', content: 'Extrais TOUTES les questions:\n\n' + inputText }]
      })
    });

    if (!response.ok) {
      const e = await response.json().catch(() => ({}));
      console.error('[CLAUDE ERROR]', e);
      return res.status(502).json({ error: 'CLAUDE_ERROR', message: 'Erreur API Claude.' });
    }

    const data = await response.json();
    const rawText = data.content?.[0]?.text || '{}';

    if (data.usage) {
      const cost = (data.usage.input_tokens * 3 / 1000000) + (data.usage.output_tokens * 15 / 1000000);
      console.log('[COUT] $' + cost.toFixed(4));
    }

    let form = parseJSON(rawText);
    
    // NORMALISATION : corriger les données manquantes
    form = normalizeForm(form);

    if (!form.questions || !Array.isArray(form.questions)) {
      return res.status(422).json({ error: 'NO_QUESTIONS', message: 'Aucune question detectee.' });
    }

    // Log de debug pour vérifier les choices
    form.questions.forEach(q => {
      if (q.choices && q.choices.length > 0) {
        console.log('[CHOICES] Q' + q.num + ': ' + q.choices.join(', '));
      }
    });

    console.log('[ANALYSE] OK - ' + form.questions.length + ' questions');
    res.json({
      success: true,
      form: form,
      truncated: text.length > MAX
    });

  } catch (err) {
    console.error('[ANALYSE ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ CORRECTION ============
app.post('/api/correct', analyseLimiter, async (req, res) => {
  try {
    const { form, instructions } = req.body;
    if (!form || !instructions) return res.status(400).json({ error: 'Donnees manquantes' });
    if (!ANTHROPIC_API_KEY) return res.status(500).json({ error: 'Cle API manquante' });

    console.log('[CORRECT] ' + instructions.slice(0, 100));

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 16000,
        system: 'Expert collecte de donnees. Applique exactement les corrections au formulaire JSON. Retourne JSON corrige UNIQUEMENT sans markdown. Conserve les choices et choice_values.',
        messages: [{
          role: 'user',
          content: 'Formulaire:\n' + JSON.stringify(form).slice(0, 8000) + '\n\nCorrections:\n' + instructions
        }]
      })
    });

    if (!response.ok) return res.status(502).json({ error: 'CLAUDE_ERROR' });

    const data = await response.json();
    let raw = data.content?.[0]?.text || '{}';
    raw = raw.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const m = raw.match(/\{[\s\S]*/);
    let corrected = parseJSON(m ? m[0] : raw);
    
    // Normaliser après correction
    corrected = normalizeForm(corrected);

    console.log('[CORRECT] OK - ' + corrected.questions?.length + ' questions');
    res.json({ success: true, form: corrected });

  } catch (err) {
    console.error('[CORRECT ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ VÉRIFICATION ============
app.post('/api/verify', analyseLimiter, async (req, res) => {
  try {
    const { form } = req.body;
    if (!form) return res.status(400).json({ error: 'Formulaire manquant' });

    const hasComplex = form.questions?.some(q =>
      q.suggestions?.some(s => s.type === 'skip_logic') ||
      q.relevant ||
      q.question_class === 'calculate'
    );

    if (!hasComplex) {
      return res.json({ valid: true, issues: [], skipped: true });
    }

    if (!ANTHROPIC_API_KEY) return res.status(500).json({ error: 'Cle API manquante' });

    const koboContent = buildKoboContent(form);
    const xlsformStr = JSON.stringify(koboContent, null, 2).slice(0, 5000);

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 4096,
        system: 'Expert XLSForm. Verifie le contenu et reponds en JSON: {"valid":true/false,"issues":["probleme"]}',
        messages: [{ role: 'user', content: 'Verifie:\n\n' + xlsformStr }]
      })
    });

    if (!response.ok) return res.json({ valid: true, issues: [] });

    const data = await response.json();
    const raw = (data.content?.[0]?.text || '{"valid":true,"issues":[]}')
      .replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();

    let result;
    try {
      const si = raw.indexOf('{');
      result = JSON.parse(si >= 0 ? raw.slice(si) : raw);
    } catch (e) {
      result = { valid: true, issues: [] };
    }

    res.json({ valid: result.valid, issues: result.issues || [] });

  } catch (err) {
    res.json({ valid: true, issues: [], error: err.message });
  }
});

// ============ DÉPLOIEMENT KOBO ============
app.post('/api/deploy/kobo', async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { username, password, server = 'https://kf.kobotoolbox.org' } = credentials;

    if (!form || !username || !password) {
      return res.status(400).json({ error: 'Donnees manquantes' });
    }

    console.log('[DEPLOY] KoboToolbox → ' + server);

    const tokenRes = await fetch(server + '/token/?format=json', {
      headers: {
        'Authorization': 'Basic ' + Buffer.from(username + ':' + password).toString('base64')
      }
    });

    if (!tokenRes.ok) {
      return res.status(401).json({ error: 'AUTH_ERROR', message: 'Identifiants incorrects.' });
    }

    const { token } = await tokenRes.json();
    const auth = {
      'Authorization': 'Token ' + token,
      'Content-Type': 'application/json'
    };

    const assetRes = await fetch(server + '/api/v2/assets/?format=json', {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({
        name: form.title || 'Formulaire R2',
        asset_type: 'survey'
      })
    });

    if (!assetRes.ok) {
      return res.status(502).json({ error: 'ASSET_ERROR', message: 'Erreur creation formulaire.' });
    }

    const { uid } = await assetRes.json();

    const koboContent = buildKoboContent(form);
    const patchRes = await fetch(server + '/api/v2/assets/' + uid + '/?format=json', {
      method: 'PATCH',
      headers: auth,
      body: JSON.stringify({
        name: form.title || 'Formulaire R2',
        content: koboContent
      })
    });

    if (!patchRes.ok) {
      const err = await patchRes.text();
      console.error('[PATCH ERROR]', err.slice(0, 200));
      return res.status(502).json({ error: 'PATCH_ERROR', message: 'Erreur import questionnaire.' });
    }

    await fetch(server + '/api/v2/assets/' + uid + '/deployment/?format=json', {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({ active: true })
    });

    console.log('[DEPLOY] Kobo OK - ' + uid);
    res.json({
      success: true,
      uid: uid,
      url: server + '/#/forms/' + uid + '/summary',
      questions: form.questions?.length || 0
    });

  } catch (err) {
    console.error('[DEPLOY KOBO ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ DÉPLOIEMENT ODK ============
app.post('/api/deploy/odk', async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { server, email, password } = credentials;

    if (!form || !server || !email || !password) {
      return res.status(400).json({ error: 'Donnees manquantes', message: 'Server, email et password requis' });
    }

    console.log('[DEPLOY] ODK Central → ' + server);

    const sessionRes = await fetch(server + '/v1/sessions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password })
    });

    if (!sessionRes.ok) {
      return res.status(401).json({ error: 'AUTH_ERROR', message: 'Identifiants ODK incorrects.' });
    }

    const { token } = await sessionRes.json();
    const cookie = '__Host-session=' + token;

    const projectsRes = await fetch(server + '/v1/projects', {
      headers: { 'Cookie': cookie, 'X-Extended-Metadata': 'true' }
    });

    let projectId;
    if (projectsRes.ok) {
      const projects = await projectsRes.json();
      const r2Project = projects.find(p => p.name === 'R2 Forms');
      if (r2Project) {
        projectId = r2Project.id;
      } else {
        const createRes = await fetch(server + '/v1/projects', {
          method: 'POST',
          headers: { 'Cookie': cookie, 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: 'R2 Forms' })
        });
        const newProject = await createRes.json();
        projectId = newProject.id;
      }
    }

    const wb = XLSX.utils.book_new();
    const koboContent = buildKoboContent(form);

    const wsSurvey = XLSX.utils.json_to_sheet(koboContent.survey);
    XLSX.utils.book_append_sheet(wb, wsSurvey, 'survey');

    if (koboContent.choices.length > 0) {
      const wsChoices = XLSX.utils.json_to_sheet(koboContent.choices);
      XLSX.utils.book_append_sheet(wb, wsChoices, 'choices');
    }

    const wsSettings = XLSX.utils.json_to_sheet(koboContent.settings);
    XLSX.utils.book_append_sheet(wb, wsSettings, 'settings');

    const xlsxBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    const FormData = require('form-data');
    const formData = new FormData();
    formData.append('xlsForm', xlsxBuffer, {
      filename: 'form.xlsx',
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });

    const uploadRes = await fetch(server + '/v1/projects/' + projectId + '/forms?publish=true', {
      method: 'POST',
      headers: {
        'Cookie': cookie,
        ...formData.getHeaders()
      },
      body: formData
    });

    if (!uploadRes.ok) {
      const err = await uploadRes.text();
      console.error('[ODK UPLOAD ERROR]', err.slice(0, 200));
      return res.status(502).json({ error: 'UPLOAD_ERROR', message: 'Erreur upload ODK.' });
    }

    const uploadData = await uploadRes.json();

    console.log('[DEPLOY] ODK OK - ' + uploadData.xmlFormId);
    res.json({
      success: true,
      projectId: projectId,
      formId: uploadData.xmlFormId,
      url: server + '/#/projects/' + projectId + '/forms/' + uploadData.xmlFormId,
      questions: form.questions?.length || 0
    });

  } catch (err) {
    console.error('[DEPLOY ODK ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ DÉPLOIEMENT JOTFORM ============
app.post('/api/deploy/jotform', async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { apiKey } = credentials;

    if (!form || !apiKey) {
      return res.status(400).json({ error: 'Donnees manquantes' });
    }

    console.log('[DEPLOY] JotForm');

    const userRes = await fetch('https://api.jotform.com/user?apiKey=' + apiKey);
    if (!userRes.ok) {
      return res.status(401).json({ error: 'AUTH_ERROR', message: 'Cle API JotForm incorrecte.' });
    }

    const userData = await userRes.json();
    if (userData.responseCode !== 200) {
      return res.status(401).json({ error: 'AUTH_ERROR', message: 'Cle API JotForm invalide.' });
    }

    const createRes = await fetch('https://api.jotform.com/form?apiKey=' + apiKey, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: 'questions%5B0%5D%5Btype%5D=control_head&questions%5B0%5D%5Btext%5D=' +
        encodeURIComponent(form.title || 'Formulaire R2') +
        '&properties%5Btitle%5D=' + encodeURIComponent(form.title || 'Formulaire R2')
    });

    if (!createRes.ok) {
      return res.status(502).json({ error: 'CREATE_ERROR', message: 'Erreur creation JotForm.' });
    }

    const createData = await createRes.json();
    const formId = createData.content?.id;

    if (!formId) {
      return res.status(502).json({ error: 'CREATE_ERROR', message: 'ID formulaire non recu.' });
    }

    const questions = form.questions || [];
    for (let i = 0; i < questions.length; i++) {
      const q = questions[i];
      const t = q.selectedType || q.type || 'text';
      let qType = 'control_textbox';

      if (t === 'select_one') qType = 'control_radio';
      else if (t === 'select_multiple') qType = 'control_checkbox';
      else if (t === 'integer' || t === 'decimal') qType = 'control_number';
      else if (t === 'date') qType = 'control_datetime';
      else if (t === 'image') qType = 'control_fileupload';

      const qData = new URLSearchParams();
      qData.append('questions[' + (i + 1) + '][type]', qType);
      qData.append('questions[' + (i + 1) + '][text]', q.label || '');
      qData.append('questions[' + (i + 1) + '][required]', q.required ? 'Yes' : 'No');
      qData.append('questions[' + (i + 1) + '][order]', String(i + 1));

      if ((qType === 'control_radio' || qType === 'control_checkbox') && q.choices?.length > 0) {
        q.choices.forEach(c => {
          const label = typeof c === 'string' ? c : (c.label || String(c));
          qData.append('questions[' + (i + 1) + '][options]', label);
        });
      }

      await fetch('https://api.jotform.com/form/' + formId + '/questions?apiKey=' + apiKey, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: qData.toString()
      });
    }

    console.log('[DEPLOY] JotForm OK - ' + formId);
    res.json({
      success: true,
      formId: formId,
      url: 'https://www.jotform.com/' + formId,
      questions: questions.length
    });

  } catch (err) {
    console.error('[DEPLOY JOTFORM ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ DÉPLOIEMENT GOOGLE FORMS ============
app.post('/api/deploy/google', async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { accessToken } = credentials;

    if (!form || !accessToken) {
      return res.status(400).json({ error: 'Token Google manquant' });
    }

    console.log('[DEPLOY] Google Forms - ' + form.questions?.length + ' questions');

    const createRes = await fetch('https://forms.googleapis.com/v1/forms', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + accessToken,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        info: {
          title: form.title || 'Formulaire R2',
          documentTitle: form.title || 'Formulaire R2'
        }
      })
    });

    if (!createRes.ok) {
      const err = await createRes.json();
      return res.status(502).json({
        error: 'CREATE_ERROR',
        message: err.error?.message || 'Erreur creation Google Form.'
      });
    }

    const { formId } = await createRes.json();

    const requests = [];
    let itemIndex = 0;

    const groups = {};
    form.questions.forEach(q => {
      const g = q.group || 'Général';
      if (!groups[g]) groups[g] = [];
      groups[g].push(q);
    });

    Object.entries(groups).forEach(([gname, qs]) => {
      if (itemIndex > 0) {
        requests.push({
          createItem: {
            item: { title: gname, pageBreakItem: {} },
            location: { index: itemIndex }
          }
        });
        itemIndex++;
      }

      qs.forEach(q => {
        const t = q.selectedType || q.type || 'text';
        const choices = (q.choices || []).map(c => ({
          value: typeof c === 'string' ? c : (c.label || String(c))
        }));

        let item = { title: q.label || '' };
        let question = {};

        if (t === 'select_one') {
          question.choiceQuestion = { type: 'RADIO', options: choices, shuffle: false };
        } else if (t === 'select_multiple') {
          question.choiceQuestion = { type: 'CHECKBOX', options: choices, shuffle: false };
        } else if (t === 'integer' || t === 'decimal') {
          question.textQuestion = { paragraph: false };
        } else if (t === 'date') {
          question.dateQuestion = { includeTime: false, includeYear: true };
        } else if (t === 'time') {
          question.timeQuestion = {};
        } else if (t === 'image') {
          question.fileUploadQuestion = { folderId: '' };
        } else {
          question.textQuestion = { paragraph: (q.label || '').length > 50 };
        }

        item.questionItem = { question, required: q.required !== false };

        requests.push({
          createItem: {
            item,
            location: { index: itemIndex }
          }
        });
        itemIndex++;
      });
    });

    if (requests.length > 0) {
      const batchRes = await fetch('https://forms.googleapis.com/v1/forms/' + formId + ':batchUpdate', {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ requests })
      });

      if (!batchRes.ok) {
        const err = await batchRes.json();
        console.warn('[GOOGLE BATCH WARN]', JSON.stringify(err).slice(0, 300));
      }
    }

    console.log('[DEPLOY] Google Forms OK - ' + formId);
    res.json({
      success: true,
      formId: formId,
      url: 'https://docs.google.com/forms/d/' + formId + '/edit',
      viewUrl: 'https://docs.google.com/forms/d/e/' + formId + '/viewform',
      questions: form.questions?.length || 0
    });

  } catch (err) {
    console.error('[DEPLOY GOOGLE ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ EXPORT XLSFORM ============
app.post('/api/export/xlsform', async (req, res) => {
  try {
    const { form } = req.body;
    if (!form) return res.status(400).json({ error: 'Formulaire manquant' });

    console.log('[EXPORT] XLSForm');

    const wb = XLSX.utils.book_new();
    const koboContent = buildKoboContent(form);

    const wsSurvey = XLSX.utils.json_to_sheet(koboContent.survey);
    XLSX.utils.book_append_sheet(wb, wsSurvey, 'survey');

    if (koboContent.choices.length > 0) {
      const wsChoices = XLSX.utils.json_to_sheet(koboContent.choices);
      XLSX.utils.book_append_sheet(wb, wsChoices, 'choices');
    }

    const wsSettings = XLSX.utils.json_to_sheet(koboContent.settings);
    XLSX.utils.book_append_sheet(wb, wsSettings, 'settings');

    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    const filename = (form.title || 'formulaire').replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '') + '.xlsx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + filename + '"');
    res.send(buffer);

    console.log('[EXPORT] XLSForm OK - ' + filename);

  } catch (err) {
    console.error('[EXPORT XLSFORM ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ TALLY.SO ============
app.post('/api/deploy/tally', async (req, res) => {
  try {
    const { form } = req.body;
    if (!form) return res.status(400).json({ error: 'Formulaire manquant' });

    console.log('[DEPLOY] Tally.so - template');

    const template = {
      title: form.title || 'Formulaire R2',
      fields: (form.questions || []).map((q, i) => {
        const t = q.selectedType || q.type;
        const base = {
          id: 'field_' + i,
          label: q.label,
          required: q.required !== false,
          description: q.hint || ''
        };

        if (t === 'select_one') return { ...base, type: 'RADIO', options: q.choices || [] };
        if (t === 'select_multiple') return { ...base, type: 'CHECKBOX', options: q.choices || [] };
        if (t === 'integer' || t === 'decimal') return { ...base, type: 'NUMBER' };
        if (t === 'date') return { ...base, type: 'DATE' };
        if (t === 'image') return { ...base, type: 'FILE_UPLOAD', accept: 'image/*' };
        return { ...base, type: 'TEXT' };
      })
    };

    res.json({
      success: true,
      method: 'manual_import',
      template: template,
      instructions: {
        fr: '1. Créez un formulaire vierge sur Tally.so\n2. Cliquez sur "Import"\n3. Collez le JSON fourni',
        en: '1. Create a blank form on Tally.so\n2. Click "Import"\n3. Paste the provided JSON'
      },
      questions: form.questions?.length || 0
    });

  } catch (err) {
    console.error('[DEPLOY TALLY ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ ERREURS ============

app.use((err, req, res, next) => {
  console.error('[ERREUR]', err.message);
  res.status(500).json({ error: 'SERVER_ERROR', message: 'Erreur serveur' });
});

app.use((req, res) => {
  res.status(404).json({ error: 'NOT_FOUND', message: 'Route non trouvée' });
});

// ============ DÉMARRAGE ============

app.listen(PORT, () => {
  console.log('╔════════════════════════════════════════╗');
  console.log('║  R2 Forms Backend v4.1.1               ║');
  console.log('║  Port: ' + PORT + '                           ║');
  console.log('║  Anthropic: ' + (ANTHROPIC_API_KEY ? 'OK' : 'MANQUANT') + '                  ║');
  console.log('╚════════════════════════════════════════╝');
});
