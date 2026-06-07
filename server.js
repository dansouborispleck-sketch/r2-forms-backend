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

app.get('/health', (req, res) => res.json({ status: 'ok', version: '4.0.0' }));

// ============ IMPORT ============
app.post('/api/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier recu' });
    const { originalname, buffer } = req.file;
    const ext = originalname.split('.').pop().toLowerCase();
    let text = '';
    console.log('[IMPORT] ' + originalname);
    if (['txt','csv'].includes(ext)) {
      text = buffer.toString('utf-8');
    } else if (ext === 'pdf') {
      const d = await pdf(buffer);
      text = d.text;
      if (!text || text.trim().length < 20)
        return res.status(422).json({ error: 'PDF_SCANNED', message: 'PDF scanne illisible. Collez le texte directement.' });
    } else if (ext === 'docx') {
      const r = await mammoth.extractRawText({ buffer });
      text = r.value;
    } else if (ext === 'doc') {
      text = buffer.toString('latin1').replace(/[^\x20-\x7E\n\r\u00C0-\u024F]/g,' ').replace(/\s+/g,' ').trim();
      if (text.length < 30)
        return res.status(422).json({ error: 'DOC_OLD', message: 'Format .doc non supporte. Enregistrez en .docx.' });
    } else if (['xlsx','xls'].includes(ext)) {
      const wb = XLSX.read(buffer, { type: 'buffer' });
      wb.SheetNames.forEach(function(n) { text += '\n=== ' + n + ' ===\n' + XLSX.utils.sheet_to_csv(wb.Sheets[n]); });
    } else {
      return res.status(400).json({ error: 'FORMAT_UNSUPPORTED', message: 'Format .' + ext + ' non supporte.' });
    }
    text = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n').replace(/\n{3,}/g,'\n\n').trim();
    if (text.length < 20)
      return res.status(422).json({ error: 'EMPTY', message: 'Fichier vide. Collez le texte directement.' });
    console.log('[IMPORT] ok ' + text.length + ' chars');
    res.json({ success: true, text: text, metadata: { filename: originalname, chars: text.length } });
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
    if (!apiKey) return res.status(500).json({ error: 'Cle API manquante' });
    const inputText = text.slice(0, 8000);
    const truncated = text.length > 8000;
    console.log('[ANALYSE] ' + text.length + ' chars -> ' + tool);

    const system = 'Expert en collecte de donnees terrain. Analyse le questionnaire et extrais TOUTES les questions.\n\n' +
'REGLES ABSOLUES:\n' +
'1. Extrais TOUTES les questions: numerotees, sans numero (date/lieu en entete), sous-questions implicites.\n' +
'2. GROUPES: Utilise le TITRE COMPLET de la section (ex: "Facteurs socio-demographiques" et NON "A").\n' +
'3. CHAMP AUTRES: Quand une modalite contient "Autres" ou "Autre":\n' +
'   - Garde la modalite dans choices[]\n' +
'   - Cree AUTOMATIQUEMENT une question de saisie libre juste apres (required=false)\n' +
'   - Label: "Si autre, precisez :"\n' +
'4. SAUTS IMPLICITES: Detecte les conditions meme non ecrites.\n' +
'5. VALEURS XLSForm: Pour skip_logic, utilise codes numeriques (0,1,2...) pas les libelles.\n' +
'6. coherence_report: liste des observations (questions manquantes, sauts detectes, incoherences).\n\n' +
'FORMAT JSON COMPACT (sans indentation):\n' +
'{"title":"titre","coherence_report":["obs1"],"questions":[{"id":"q1","num":1,"label":"libelle","question_class":"CLASS","type":"TYPE","required":true,"hint":"","choices":[],"choice_values":[],"group":"TITRE COMPLET","formats":[],"suggested_format_idx":0,"suggestions":[]}],"groups":[]}\n\n' +
'FORMATS PAR CLASSE:\n' +
'quantitative:[{"id":"A","name":"Nombre entier","type":"integer","note":"Ex: 25"},{"id":"B","name":"Nombre avec virgule","type":"decimal","note":"Ex: 65,5"},{"id":"C","name":"Valeur sur une echelle","type":"range","note":"Ex: 1 a 10"}]\n' +
'qualitative_choice:[{"id":"A","name":"Une seule reponse au choix","type":"select_one","note":"Une case"},{"id":"B","name":"Plusieurs reponses possibles","type":"select_multiple","note":"Plusieurs cases"},{"id":"C","name":"Reponse ecrite libre","type":"text","note":"Libre"}]\n' +
'Pour mention "Plusieurs reponses" ou "Cochez toutes": suggested_format_idx=1\n' +
'qualitative_open:[{"id":"A","name":"Reponse courte","type":"text","note":"Quelques mots"},{"id":"B","name":"Reponse longue","type":"text","note":"Plusieurs phrases"}]\n' +
'date_time:[{"id":"A","name":"Date","type":"date","note":"jj/mm/aaaa"},{"id":"B","name":"Heure","type":"time","note":"hh:mm"},{"id":"C","name":"Date et heure","type":"datetime","note":"jj/mm/aaaa hh:mm"}]\n' +
'geopoint:[{"id":"A","name":"Localisation GPS","type":"geopoint","note":"GPS auto"}]\n' +
'geotrace:[{"id":"A","name":"Tracer un chemin GPS","type":"geotrace","note":"Trajet"}]\n' +
'geoshape:[{"id":"A","name":"Delimiter une zone GPS","type":"geoshape","note":"Zone"}]\n' +
'media_photo:[{"id":"A","name":"Prendre une photo","type":"image","note":"Photo"}]\n' +
'media_audio:[{"id":"A","name":"Enregistrer un son","type":"audio","note":"Audio"}]\n' +
'media_video:[{"id":"A","name":"Enregistrer une video","type":"video","note":"Video"}]\n' +
'media_file:[{"id":"A","name":"Joindre un fichier","type":"file","note":"Fichier"}]\n' +
'barcode:[{"id":"A","name":"Scanner un code-barres","type":"barcode","note":"QR/barcode"}]\n' +
'acknowledge:[{"id":"A","name":"Case a cocher pour confirmer","type":"acknowledge","note":"Confirmation"}]\n' +
'ranking:[{"id":"A","name":"Classer par ordre de preference","type":"rank","note":"Ordre"}]\n' +
'scale:[{"id":"A","name":"Valeur sur une echelle","type":"range","note":"Curseur"}]\n' +
'calculate:[{"id":"A","name":"Valeur calculee automatiquement","type":"calculate","note":"Calcul"}]\n' +
'note:[{"id":"A","name":"Message information","type":"note","note":"Sans saisie"}]\n\n' +
'SUGGESTIONS: {"type":"skip_logic|calculate|constraint","label":"court","description":"clair","value":"formule XLSForm","confidence":"high|medium|low"}\n' +
'REGLES: required=true par defaut. JSON compact. choice_values=codes numeriques ["0","1",...]. Outil: ' + tool;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 32000,
        system: system,
        messages: [{ role: 'user', content: 'Extrais TOUTES les questions:\n\n' + inputText }]
      })
    });

    if (!response.ok) {
      const e = await response.json();
      console.error('[CLAUDE ERROR]', e);
      return res.status(502).json({ error: 'CLAUDE_ERROR', message: 'Erreur API Claude.' });
    }

    const data = await response.json();
    const rawText = data.content && data.content[0] ? data.content[0].text : '{}';

    let form;
    try {
      let jsonStr = rawText.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim();
      const s = jsonStr.indexOf('{');
      if (s > 0) jsonStr = jsonStr.slice(s);
      try {
        form = JSON.parse(jsonStr);
      } catch(e) {
        const lc = jsonStr.lastIndexOf('},');
        if (lc > 100) jsonStr = jsonStr.slice(0, lc+1);
        jsonStr = jsonStr.replace(/,\s*$/,'');
        const ob = (jsonStr.match(/\[/g)||[]).length-(jsonStr.match(/\]/g)||[]).length;
        const cb = (jsonStr.match(/\{/g)||[]).length-(jsonStr.match(/\}/g)||[]).length;
        for(var i=0;i<Math.max(0,ob);i++) jsonStr+=']';
        for(var j=0;j<Math.max(0,cb);j++) jsonStr+='}';
        form = JSON.parse(jsonStr);
        console.log('[PARSE] JSON repare');
      }
    } catch(pe) {
      console.error('[PARSE ERROR]', pe.message);
      return res.status(502).json({ error: 'PARSE_ERROR', message: 'Reponse invalide.' });
    }

    if (!form.questions || !Array.isArray(form.questions))
      return res.status(422).json({ error: 'NO_QUESTIONS', message: 'Aucune question detectee.' });

    console.log('[ANALYSE] ok ' + form.questions.length + ' questions');
    res.json({
      success: true, form: form, truncated: truncated,
      warning: truncated ? 'Questionnaire tronque a ' + inputText.length + ' chars sur ' + text.length + ' total.' : null
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
        messages: [{ role: 'user', content: 'Formulaire:\n' + JSON.stringify(form) + '\n\nCorrections:\n' + instructions }]
      })
    });
    if (!response.ok) return res.status(502).json({ error: 'CLAUDE_ERROR' });
    const data = await response.json();
    let raw = data.content && data.content[0] ? data.content[0].text : '{}';
    raw = raw.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim();
    const m = raw.match(/\{[\s\S]*/);
    const corrected = JSON.parse(m ? m[0] : '{}');
    res.json({ success: true, form: corrected });
  } catch(err) {
    res.status(500).json({ error: 'SERVER_ERROR' });
  }
});

// ============ VERIFICATION XLSFORM ============
app.post('/api/verify', async (req, res) => {
  try {
    const { form } = req.body;
    if (!form) return res.status(400).json({ error: 'Formulaire manquant' });
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'Cle API manquante' });
    const koboContent = buildKoboContent(form);
    const xlsformStr = JSON.stringify(koboContent, null, 2);
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5', max_tokens: 2048,
        system: 'Expert XLSForm. Verifie le contenu et reponds en JSON: {"valid":true/false,"issues":["probleme"],"fixed_questions":[{"id":"q1","relevant":"formule corrigee"}]}',
        messages: [{ role: 'user', content: 'Verifie:\n\n' + xlsformStr.slice(0, 6000) }]
      })
    });
    if (!response.ok) return res.json({ valid: true, issues: [] });
    const data = await response.json();
    const raw = (data.content && data.content[0] ? data.content[0].text : '{"valid":true,"issues":[]}').replace(/```json\n?/g,'').replace(/```\n?/g,'').trim();
    const si = raw.indexOf('{');
    let result;
    try { result = JSON.parse(si >= 0 ? raw.slice(si) : raw); } catch(e) { result = { valid: true, issues: [] }; }
    if (result.fixed_questions && result.fixed_questions.length > 0) {
      result.fixed_questions.forEach(function(fix) {
        const q = form.questions.find(function(q) { return q.id === fix.id; });
        if (q && fix.relevant) q.relevant = fix.relevant;
      });
    }
    res.json({ valid: result.valid, issues: result.issues || [], form: form });
  } catch(err) {
    res.json({ valid: true, issues: [] });
  }
});

// ============ DEPLOIEMENT KOBO ============
app.post('/api/deploy/kobo', async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { username, password, server = 'https://kf.kobotoolbox.org' } = credentials;
    if (!form || !username || !password) return res.status(400).json({ error: 'Donnees manquantes' });
    console.log('[DEPLOY] KoboToolbox -> ' + server);
    const tokenRes = await fetch(server + '/token/?format=json', {
      headers: { 'Authorization': 'Basic ' + Buffer.from(username + ':' + password).toString('base64') }
    });
    if (!tokenRes.ok) return res.status(401).json({ error: 'AUTH_ERROR', message: 'Identifiants incorrects.' });
    const token = (await tokenRes.json()).token;
    const auth = { 'Authorization': 'Token ' + token, 'Content-Type': 'application/json' };
    const assetRes = await fetch(server + '/api/v2/assets/?format=json', {
      method: 'POST', headers: auth,
      body: JSON.stringify({ name: form.title || 'Formulaire R2', asset_type: 'survey' })
    });
    if (!assetRes.ok) return res.status(502).json({ error: 'ASSET_ERROR', message: 'Erreur creation formulaire.' });
    const uid = (await assetRes.json()).uid;
    const koboContent = buildKoboContent(form);
    const patchRes = await fetch(server + '/api/v2/assets/' + uid + '/?format=json', {
      method: 'PATCH', headers: auth,
      body: JSON.stringify({ name: form.title || 'Formulaire R2', content: koboContent })
    });
    if (!patchRes.ok) {
      console.error('[PATCH ERROR]', await patchRes.text());
      return res.status(502).json({ error: 'PATCH_ERROR', message: 'Erreur import questionnaire.' });
    }
    await fetch(server + '/api/v2/assets/' + uid + '/deployment/?format=json', {
      method: 'POST', headers: auth, body: JSON.stringify({ active: true })
    });
    console.log('[DEPLOY] Kobo ok ' + uid);
    res.json({ success: true, uid: uid, url: server + '/#/forms/' + uid + '/summary', questions: (form.questions||[]).length });
  } catch(err) {
    console.error('[DEPLOY KOBO ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ DEPLOIEMENT JOTFORM ============
app.post('/api/deploy/jotform', async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { apiKey } = credentials;
    if (!form || !apiKey) return res.status(400).json({ error: 'Donnees manquantes' });
    console.log('[DEPLOY] JotForm');
    const userRes = await fetch('https://api.jotform.com/user?apiKey=' + apiKey);
    if (!userRes.ok) return res.status(401).json({ error: 'AUTH_ERROR', message: 'Cle API JotForm incorrecte.' });
    const userData = await userRes.json();
    if (userData.responseCode !== 200) return res.status(401).json({ error: 'AUTH_ERROR', message: 'Cle API invalide.' });
    const createRes = await fetch('https://api.jotform.com/form?apiKey=' + apiKey, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: 'questions%5B0%5D%5Btype%5D=control_head&questions%5B0%5D%5Btext%5D=' + encodeURIComponent(form.title || 'Formulaire R2') + '&properties%5Btitle%5D=' + encodeURIComponent(form.title || 'Formulaire R2')
    });
    if (!createRes.ok) return res.status(502).json({ error: 'CREATE_ERROR', message: 'Erreur creation formulaire JotForm.' });
    const createData = await createRes.json();
    const formId = createData.content && createData.content.id;
    if (!formId) return res.status(502).json({ error: 'CREATE_ERROR', message: 'ID formulaire non recu.' });
    const questions = form.questions || [];
    for (var qi = 0; qi < questions.length; qi++) {
      var q = questions[qi];
      var t = q.selectedType || q.type || 'text';
      var qType = 'control_textbox';
      if (t === 'select_one') qType = 'control_radio';
      else if (t === 'select_multiple') qType = 'control_checkbox';
      else if (t === 'integer' || t === 'decimal') qType = 'control_number';
      else if (t === 'date') qType = 'control_datetime';
      else if (t === 'image') qType = 'control_fileupload';
      var qData = new URLSearchParams();
      qData.append('questions[' + (qi+1) + '][type]', qType);
      qData.append('questions[' + (qi+1) + '][text]', q.label || '');
      qData.append('questions[' + (qi+1) + '][required]', q.required ? 'Yes' : 'No');
      qData.append('questions[' + (qi+1) + '][order]', String(qi+1));
      if ((qType === 'control_radio' || qType === 'control_checkbox') && q.choices && q.choices.length > 0) {
        q.choices.forEach(function(c) {
          var label = typeof c === 'string' ? c : (c.label || String(c));
          qData.append('questions[' + (qi+1) + '][options]', label);
        });
      }
      await fetch('https://api.jotform.com/form/' + formId + '/questions?apiKey=' + apiKey, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: qData.toString()
      });
    }
    console.log('[DEPLOY] JotForm ok ' + formId);
    res.json({ success: true, formId: formId, url: 'https://www.jotform.com/' + formId, questions: questions.length });
  } catch(err) {
    console.error('[DEPLOY JOTFORM ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ DEPLOIEMENT GOOGLE FORMS ============
app.post('/api/deploy/google', async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { accessToken } = credentials;
    if (!form || !accessToken) return res.status(400).json({ error: 'Token Google manquant' });
    const questions = form.questions || [];
    console.log('[DEPLOY] Google Forms - ' + questions.length + ' questions');

    // 1. Creer le formulaire vide
    const createRes = await fetch('https://forms.googleapis.com/v1/forms', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
      body: JSON.stringify({ info: { title: form.title || 'Formulaire R2', documentTitle: form.title || 'Formulaire R2' } })
    });
    if (!createRes.ok) return res.status(502).json({ error: 'CREATE_ERROR', message: 'Erreur creation Google Form.' });
    const formId = (await createRes.json()).formId;
    console.log('[GOOGLE] FormId: ' + formId);

    // Helper: construire un questionItem selon le type
    function makeQItem(q) {
      var t = q.selectedType || q.type || 'text';
      var required = q.required !== false;
      var choices = (q.choices||[]).map(function(c){ return { value: typeof c==='string'?c:(c.label||String(c)) }; });
      if (t==='select_one' && choices.length>0) return { question:{ required:required, choiceQuestion:{ type:'RADIO', options:choices, shuffle:false } } };
      if (t==='select_multiple' && choices.length>0) return { question:{ required:required, choiceQuestion:{ type:'CHECKBOX', options:choices, shuffle:false } } };
      if (t==='date') return { question:{ required:required, dateQuestion:{ includeTime:false, includeYear:true } } };
      if (t==='time'||t==='datetime') return { question:{ required:required, timeQuestion:{ duration:false } } };
      if (t==='scale'||t==='range') return { question:{ required:required, scaleQuestion:{ low:parseInt(q.numMin)||1, high:parseInt(q.numMax)||10, lowLabel:'Min', highLabel:'Max' } } };
      return { question:{ required:required, textQuestion:{ paragraph:(q.label||'').length>50 } } };
    }

    function makeHint(q) {
      var t = q.selectedType || q.type || 'text';
      var hint = q.hint || '';
      if (t==='integer'||t==='decimal') {
        var parts=[];
        if (q.numMin!==''&&q.numMin!=null) parts.push('min: '+q.numMin);
        if (q.numMax!==''&&q.numMax!=null) parts.push('max: '+q.numMax);
        if (parts.length>0) return (hint?hint+' — ':'')+parts.join(', ');
      }
      return hint;
    }

    // 2. Analyser tous les sauts confirmes pour determiner les sections necessaires
    // Un saut = une question A dont certaines reponses doivent afficher des questions B,C,D
    // et d'autres reponses doivent les sauter
    // Dans Google Forms: B,C,D doivent etre dans une section separee

    // Identifier toutes les questions cibles de sauts
    // q.relevant contient la formule XLSForm: "${q1} = '1'" -> la question q1 est le declencheur
    var skipTargets = {}; // questionId -> true (cette question est ciblee par un saut)
    questions.forEach(function(q) {
      if (q.relevant && q.relevant.trim()) {
        // Extraire l'ID de la question source du relevant
        var match = q.relevant.match(/\$\{([^}]+)\}/);
        if (match) skipTargets[q.id] = match[1]; // q.id -> id_source
      }
      // Detecter aussi les questions "Si autre, precisez"
      if (q.label) {
        var l = q.label.toLowerCase();
        if (l.includes('si autre')||l.includes('precisez')||l.includes('preciser')) {
          skipTargets[q.id] = '_autre';
        }
      }
    });

    // 3. Construire la structure des sections Google Forms
    // Principe: chaque groupe de questions conditionnelles = une section separee
    // Les questions non conditionnelles restent dans leur section de groupe

    // Construire des "blocs" de questions:
    // - Bloc normal: questions sans saut entrant
    // - Bloc conditionnel: questions avec saut entrant (meme source, meme condition)

    var buildRequests = [];
    var sectionMeta = []; // { type:'section'|'question', label, qIdx, q, isConditional, sourceId, condValue }
    var itemIdx = 0;

    // Grouper les questions par leur groupe thematique
    var groupMap = {};
    var groupOrder = [];
    questions.forEach(function(q) {
      var g = q.group || 'general';
      if (!groupMap[g]) { groupMap[g] = []; groupOrder.push(g); }
      groupMap[g].push(q);
    });

    // Pour chaque groupe, identifier les sous-groupes conditionnels
    groupOrder.forEach(function(gname) {
      var qs = groupMap[gname];

      // Section de groupe principale
      if (gname !== 'general') {
        buildRequests.push({ createItem:{ item:{ title:gname, pageBreakItem:{} }, location:{ index:itemIdx } } });
        sectionMeta.push({ type:'section', label:gname, itemIdx:itemIdx });
        itemIdx++;
      }

      // Identifier les blocs: normal vs conditionnel
      // Un bloc conditionnel commence quand une question a un relevant
      var i = 0;
      while (i < qs.length) {
        var q = qs[i];
        var hasRelevant = q.relevant && q.relevant.trim();
        var isAutreQ = q.label && (q.label.toLowerCase().includes('si autre')||q.label.toLowerCase().includes('precisez'));

        if (hasRelevant || isAutreQ) {
          // Cette question et les suivantes avec le meme relevant forment un bloc conditionnel
          // Creer une section vide pour ce bloc
          var sectionLabel = '_cond_' + itemIdx;
          buildRequests.push({ createItem:{ item:{ title:'', pageBreakItem:{} }, location:{ index:itemIdx } } });
          sectionMeta.push({ type:'section', label:sectionLabel, itemIdx:itemIdx, isConditional:true, relevant: q.relevant||'_autre' });
          itemIdx++;

          // Ajouter toutes les questions consecutives avec le meme relevant (ou "Si autre")
          while (i < qs.length) {
            var qc = qs[i];
            var qcRelevant = qc.relevant && qc.relevant.trim();
            var qcIsAutre = qc.label && (qc.label.toLowerCase().includes('si autre')||qc.label.toLowerCase().includes('precisez'));
            // Continuer si meme source de saut
            if (i > 0 && !qcRelevant && !qcIsAutre) break;
            if (i > 0 && qcRelevant && qcRelevant !== q.relevant) break;

            buildRequests.push({ createItem:{ item:{ title:qc.label||('Q'+(itemIdx+1)), description:makeHint(qc), questionItem:makeQItem(qc) }, location:{ index:itemIdx } } });
            sectionMeta.push({ type:'question', q:qc, itemIdx:itemIdx, isConditional:true });
            itemIdx++;
            i++;
          }
        } else {
          // Question normale
          buildRequests.push({ createItem:{ item:{ title:q.label||('Q'+(itemIdx+1)), description:makeHint(q), questionItem:makeQItem(q) }, location:{ index:itemIdx } } });
          sectionMeta.push({ type:'question', q:q, itemIdx:itemIdx, isConditional:false });
          itemIdx++;
          i++;
        }
      }
    });

    // Section finale pour absorber les sauts "passer les conditionnels"
    buildRequests.push({ createItem:{ item:{ title:'', pageBreakItem:{} }, location:{ index:itemIdx } } });
    sectionMeta.push({ type:'section', label:'_final', itemIdx:itemIdx, isFinal:true });

    // 4. Envoyer le batch de creation
    console.log('[GOOGLE] Creation de ' + buildRequests.length + ' items (questions + sections)');
    var batch1Res = await fetch('https://forms.googleapis.com/v1/forms/' + formId + ':batchUpdate', {
      method:'POST',
      headers:{ 'Authorization':'Bearer '+accessToken, 'Content-Type':'application/json' },
      body: JSON.stringify({ requests:buildRequests, includeFormInResponse:true })
    });
    if (!batch1Res.ok) {
      var e = await batch1Res.json();
      console.error('[GOOGLE BATCH1]', JSON.stringify(e).slice(0,300));
      return res.status(502).json({ error:'BATCH_ERROR', message:'Erreur ajout questions.' });
    }
    var batch1Data = await batch1Res.json();
    var updatedForm = batch1Data.form;

    // 5. Appliquer la navigation goToSectionId
    if (updatedForm && updatedForm.items) {
      // Mapper index Google -> itemId
      var orderedItems = updatedForm.items.slice().sort(function(a,b){ return (a.index||0)-(b.index||0); });
      orderedItems.forEach(function(item, i) { if (sectionMeta[i]) sectionMeta[i].googleItemId = item.itemId; });

      var navRequests = [];

      // Pour chaque question normale a choix qui precede un bloc conditionnel
      sectionMeta.forEach(function(sm, idx) {
        if (sm.type !== 'question' || sm.isConditional) return;
        var q = sm.q;
        var t = q.selectedType || q.type || 'text';
        if (t !== 'select_one' && t !== 'select_multiple') return;

        // Chercher les sections conditionnelles qui suivent cette question
        // et qui correspondent a ses modalites
        var nextSectionIdx = null;
        var conditionalSections = [];

        for (var i2 = idx+1; i2 < sectionMeta.length; i2++) {
          var sm2 = sectionMeta[i2];
          if (sm2.type === 'section' && !sm2.isConditional && !sm2.isFinal) break; // autre groupe
          if (sm2.type === 'section' && sm2.isConditional) {
            conditionalSections.push({ sectionId: sm2.googleItemId, relevant: sm2.relevant });
          }
          if (sm2.type === 'section' && (sm2.isFinal || (!sm2.isConditional && sm2.label !== '_final'))) {
            nextSectionIdx = sm2.googleItemId;
            break;
          }
        }

        if (conditionalSections.length === 0) return;

        // Determiner pour chaque choix vers quelle section aller
        var choices = q.choices || [];
        var choiceVals = q.choice_values || [];

        // Section finale par defaut (apres tous les conditionnels)
        var finalSectionId = null;
        for (var i3 = idx+1; i3 < sectionMeta.length; i3++) {
          var sm3 = sectionMeta[i3];
          if (sm3.type === 'section' && (sm3.isFinal || (!sm3.isConditional && !sm3.isFinal && i3 > idx+2))) {
            finalSectionId = sm3.googleItemId;
            break;
          }
        }
        if (!finalSectionId && sectionMeta[sectionMeta.length-1]) {
          finalSectionId = sectionMeta[sectionMeta.length-1].googleItemId;
        }

        var opts = choices.map(function(c, ci) {
          var label = typeof c === 'string' ? c : (c.label||String(c));
          var val = (choiceVals[ci] !== undefined) ? String(choiceVals[ci]) : String(ci);
          var opt = { value: label };

          // Chercher si un des blocs conditionnels correspond a ce choix
          var targetSection = null;
          conditionalSections.forEach(function(cs) {
            if (!cs.relevant) return;
            // Verifier si le relevant pointe vers ce choix
            // Ex: "${q5} = '1'" -> val = '1'
            if (cs.relevant === '_autre') {
              // Pour "Autres, precisez"
              if (label.toLowerCase().includes('autre')||label.toLowerCase().includes('other')) {
                targetSection = cs.sectionId;
              }
            } else {
              // Verifier si la valeur du choix correspond au relevant
              var valInRelevant = cs.relevant.match(/=\s*'([^']+)'/);
              if (valInRelevant && (valInRelevant[1] === val || valInRelevant[1] === label.toLowerCase())) {
                targetSection = cs.sectionId;
              }
            }
          });

          if (targetSection) opt.goToSectionId = targetSection;
          else if (finalSectionId) opt.goToSectionId = finalSectionId;
          else opt.goToAction = 'NEXT_SECTION';
          return opt;
        });

        if (!sm.googleItemId) return;

        navRequests.push({ updateItem:{
          item:{ itemId:sm.googleItemId, title:q.label||'', questionItem:{ question:{ required:q.required!==false, choiceQuestion:{ type:t==='select_multiple'?'CHECKBOX':'RADIO', options:opts, shuffle:false } } } },
          location:{ index:sm.itemIdx },
          updateMask:'questionItem'
        }});
      });

      if (navRequests.length > 0) {
        console.log('[GOOGLE] ' + navRequests.length + ' navigations a configurer');
        var batch2Res = await fetch('https://forms.googleapis.com/v1/forms/' + formId + ':batchUpdate', {
          method:'POST',
          headers:{ 'Authorization':'Bearer '+accessToken, 'Content-Type':'application/json' },
          body: JSON.stringify({ requests:navRequests })
        });
        if (!batch2Res.ok) {
          var e2 = await batch2Res.json();
          console.warn('[GOOGLE BATCH2]', JSON.stringify(e2).slice(0,300));
        } else {
          console.log('[GOOGLE] Navigation configuree avec succes');
        }
      }
    }

    console.log('[DEPLOY] Google Forms ok ' + formId);
    res.json({ success:true, formId:formId, url:'https://docs.google.com/forms/d/'+formId+'/edit', questions:questions.length });

  } catch(err) {
    console.error('[DEPLOY GOOGLE ERROR]', err.message);
    res.status(500).json({ error:'SERVER_ERROR', message:err.message });
  }
});

// ============ BUILDER KOBOTOOLBOX ============
function buildKoboContent(form) {
  var survey = [];
  var choices = [];
  var seen = new Set();
  var groups = {};
  (form.questions || []).forEach(function(q) {
    var g = q.group || 'general';
    if (!groups[g]) groups[g] = [];
    groups[g].push(q);
  });
  Object.keys(groups).forEach(function(gname) {
    var qs = groups[gname];
    var gId = gname.replace(/\s+/g,'_').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9_]/g,'').slice(0,32) || 'group';
    if (gname !== 'general') survey.push({ type: 'begin_group', name: gId, label: gname });
    qs.forEach(function(q) {
      var fmtIdx = q.validatedFormatIdx !== undefined ? q.validatedFormatIdx : (q.suggested_format_idx || 0);
      var selectedFmt = q.formats && q.formats[fmtIdx] ? q.formats[fmtIdx] : null;
      var t = q.selectedType || (selectedFmt ? selectedFmt.type : null) || q.type || 'text';
      var name = (q.id || ('q'+q.num)).replace(/[^a-zA-Z0-9_]/g,'_');
      var row = { type: t, name: name, label: q.label || '', required: q.required !== false ? 'yes' : 'no', hint: q.hint || '' };
      // Skip logic
      var relevant = q.relevant || '';
      if (!relevant && q.suggestions) {
        q.suggestions.forEach(function(s, si) {
          if (s.type === 'skip_logic' && s.value && q.confirmedSuggestions && q.confirmedSuggestions[si]) relevant = s.value;
        });
      }
      if (relevant) row.relevant = relevant;
      // Calculate
      if (t === 'calculate') {
        var calc = q.calculation || '';
        if (!calc && q.suggestions) {
          q.suggestions.forEach(function(s, si) {
            if (s.type === 'calculate' && s.value && q.confirmedSuggestions && q.confirmedSuggestions[si]) calc = s.value;
          });
        }
        if (calc) row.calculation = calc;
        delete row.required;
      }
      // Constraints
      var constraints = [];
      if (q.numMin !== '' && q.numMin != null) constraints.push('. >= ' + q.numMin);
      if (q.numMax !== '' && q.numMax != null) constraints.push('. <= ' + q.numMax);
      if (q.numDigitsBefore) constraints.push('string-length(substring-before(string(.), \'.\')) <= ' + q.numDigitsBefore);
      if (q.numDigitsAfter) constraints.push('string-length(substring-after(string(.), \'.\')) <= ' + q.numDigitsAfter);
      if (q.constraint && q.constraint.trim()) constraints.push(q.constraint.trim());
      if (q.suggestions) {
        q.suggestions.forEach(function(s, si) {
          if (s.type === 'constraint' && s.value && q.confirmedSuggestions && q.confirmedSuggestions[si]) constraints.push(s.value);
        });
      }
      if (constraints.length > 0) { row.constraint = constraints.join(' and '); row.constraint_message = 'Valeur hors limites'; }
      if (t === 'range') { row.parameters = 'start=' + (q.numMin||1) + ' end=' + (q.numMax||10); delete row.required; }
      if (t === 'note' || t === 'calculate') delete row.required;
      // Choices
      if (t === 'select_one' || t === 'select_multiple' || t === 'rank') {
        var listName = 'list_' + name;
        row.type = t + ' ' + listName;
        if (!seen.has(listName)) {
          seen.add(listName);
          var choiceVals = q.choice_values || [];
          (q.choices || []).forEach(function(c, i) {
            var label = typeof c === 'string' ? c : (c.label || String(c));
            var val = (choiceVals[i] !== undefined && choiceVals[i] !== '') ? String(choiceVals[i]) : String(i);
            choices.push({ list_name: listName, name: val, label: label });
          });
        }
      }
      survey.push(row);
    });
    if (gname !== 'general') survey.push({ type: 'end_group', name: gId });
  });
  var formId = (form.title||'formulaire').replace(/\s+/g,'_').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9_]/g,'').slice(0,32);
  return { survey: survey, choices: choices, settings: [{ form_title: form.title || 'Formulaire', form_id: formId, version: '1' }] };
}

app.listen(PORT, function() {
  console.log('\n╔══════════════════════════════════╗\n║   R2 Forms Backend v4.0          ║\n║   Port: ' + PORT + '                    ║\n╚══════════════════════════════════╝\n');
});
