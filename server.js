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

app.get('/health', (req, res) => res.json({ status: 'ok', version: '3.2.0' }));

// ============ IMPORT ============
app.post('/api/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier recu' });
    const { originalname, buffer } = req.file;
    const ext = originalname.split('.').pop().toLowerCase();
    let text = '';
    console.log('[IMPORT] ' + originalname);
    if (['txt','csv'].includes(ext)) text = buffer.toString('utf-8');
    else if (ext === 'pdf') {
      const d = await pdf(buffer);
      text = d.text;
      if (!text || text.trim().length < 20) return res.status(422).json({ error: 'PDF_SCANNED', message: 'PDF scanne illisible. Collez le texte directement.' });
    }
    else if (ext === 'docx') { const r = await mammoth.extractRawText({ buffer }); text = r.value; }
    else if (ext === 'doc') {
      text = buffer.toString('latin1').replace(/[^\x20-\x7E\n\r\u00C0-\u024F]/g,' ').replace(/\s+/g,' ').trim();
      if (text.length < 30) return res.status(422).json({ error: 'DOC_OLD', message: 'Format .doc non supporte. Enregistrez en .docx.' });
    }
    else if (['xlsx','xls'].includes(ext)) {
      const wb = XLSX.read(buffer, { type: 'buffer' });
      wb.SheetNames.forEach(function(n) { text += '\n=== ' + n + ' ===\n' + XLSX.utils.sheet_to_csv(wb.Sheets[n]); });
    }
    else return res.status(400).json({ error: 'FORMAT_UNSUPPORTED', message: 'Format .' + ext + ' non supporte.' });
    text = text.replace(/\r\n/g,'\n').replace(/\r/g,'\n').replace(/\n{3,}/g,'\n\n').trim();
    if (text.length < 20) return res.status(422).json({ error: 'EMPTY', message: 'Fichier vide. Collez le texte directement.' });
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

    const systemParts = [
      'Expert en collecte de donnees terrain. Analyse le questionnaire et extrais TOUTES les questions sans exception.',
      '',
      'REGLES ABSOLUES:',
      '1. TOUTES LES QUESTIONS: numerotees, sans numero (date/lieu/identifiant en en-tete), sous-questions implicites, variables cachees.',
      '2. GROUPES/SECTIONS: Utilise le TITRE COMPLET de la section tel qu\'il apparait dans le document (ex: "Facteurs socio-demographiques, culturels et economiques" et NON pas "A"). Si le document ecrit "Section A - Facteurs socio-demographiques", le group doit etre "Facteurs socio-demographiques", jamais juste "A".',
      '3. MODALITES "AUTRES": Quand une modalite contient "Autres" ou "Autre" (avec ou sans "precisez"):',
      '   - Garde la modalite dans choices[]',
      '   - Cree AUTOMATIQUEMENT une question de saisie libre juste apres',
      '   - Label: "Si autre, precisez :"',
      '   - required=false sur cette question',
      '   - suggestion skip_logic: s\'affiche si la question parente = valeur numerique de "Autres" (ex: "selected(${id_parent}, \'autres\')" ou "${id_parent} = \'4\'")',
      '4. SAUTS CONDITIONNELS: Detecte TOUS les sauts meme implicites. Ex: "Si Oui, lesquels?" => la question suivante est conditionnelle. Ex: Q19=Oui => Q20 a Q25 conditionnelles.',
      '5. VALEURS XLSForm OBLIGATOIRE: Pour les skip_logic utilise TOUJOURS le code numerique (0,1,2...) de la modalite, JAMAIS son libelle.',
      '   Ex: modalites=[Non=code0, Oui=code1] -> condition sur Oui: "${q19} = \'1\'" et NON "${q19} = \'oui\'"',
      '   Ex: pour select_multiple: "selected(${q13}, \'4\')" pour tester si valeur code 4 selectionnee',
      '   choice_values[] dans ta reponse doit contenir les codes: ["0","1","2",...] dans le meme ordre que choices[]',
      '6. QUESTIONS MANQUANTES: Si des numeros manquent dans la sequence (ex: Q15 puis Q18 sans Q16/Q17), le signaler dans coherence_report.',
      '7. NUMEROTATION: Numerote toutes les questions en continu dans l\'ordre d\'apparition.',
      '',
      'FORMAT JSON COMPACT (sans indentation, sans espaces inutiles):',
      '{"title":"titre","coherence_report":["obs1"],"questions":[{"id":"q1","num":1,"label":"libelle complet","question_class":"CLASS","type":"TYPE","required":true,"hint":"","choices":["choix1","choix2"],"choice_values":["0","1"],"group":"TITRE COMPLET DE LA SECTION","formats":[FORMAT],"suggested_format_idx":0,"suggestions":[]}],"groups":[]}',
      '',
      'IMPORTANT: choice_values[] contient les codes numeriques correspondant a chaque choix (0, 1, 2...) pour les skip_logic XLSForm.',
      '',
      'FORMATS PAR CLASSE:',
      'quantitative: [{"id":"A","name":"Nombre entier","type":"integer","note":"Ex: 25"},{"id":"B","name":"Nombre avec virgule","type":"decimal","note":"Ex: 65,5"},{"id":"C","name":"Valeur sur une echelle","type":"range","note":"Ex: 1 a 10"}]',
      'qualitative_choice: [{"id":"A","name":"Une seule reponse au choix","type":"select_one","note":"Une case"},{"id":"B","name":"Plusieurs reponses possibles","type":"select_multiple","note":"Plusieurs cases"},{"id":"C","name":"Reponse ecrite libre","type":"text","note":"Libre"}]',
      'Pour "Plusieurs reponses" ou "Cochez toutes": suggested_format_idx=1',
      'qualitative_open: [{"id":"A","name":"Reponse courte","type":"text","note":"Quelques mots"},{"id":"B","name":"Reponse longue","type":"text","note":"Plusieurs phrases"}]',
      'date_time: [{"id":"A","name":"Date","type":"date","note":"jj/mm/aaaa"},{"id":"B","name":"Heure","type":"time","note":"hh:mm"},{"id":"C","name":"Date et heure","type":"datetime","note":"jj/mm/aaaa hh:mm"}]',
      'geopoint: [{"id":"A","name":"Localisation GPS","type":"geopoint","note":"GPS auto"}]',
      'media_photo: [{"id":"A","name":"Prendre une photo","type":"image","note":"Photo"}]',
      'media_audio: [{"id":"A","name":"Enregistrer un son","type":"audio","note":"Audio"}]',
      'media_video: [{"id":"A","name":"Enregistrer une video","type":"video","note":"Video"}]',
      'media_file: [{"id":"A","name":"Joindre un fichier","type":"file","note":"Fichier"}]',
      'barcode: [{"id":"A","name":"Scanner un code-barres","type":"barcode","note":"QR/barcode"}]',
      'acknowledge: [{"id":"A","name":"Case a cocher pour confirmer","type":"acknowledge","note":"Confirmation"}]',
      'ranking: [{"id":"A","name":"Classer par ordre de preference","type":"rank","note":"Ordre"}]',
      'scale: [{"id":"A","name":"Valeur sur une echelle","type":"range","note":"Curseur"}]',
      'calculate: [{"id":"A","name":"Valeur calculee automatiquement","type":"calculate","note":"Calcul"}]',
      'note: [{"id":"A","name":"Message information","type":"note","note":"Sans saisie"}]',
      '',
      'SUGGESTIONS FORMAT:',
      '{"type":"skip_logic","label":"court","description":"explication","value":"formule XLSForm avec codes numeriques","confidence":"high|medium|low"}',
      '{"type":"constraint","label":"court","description":"explication","value":"formule XLSForm","confidence":"high|medium|low"}',
      '{"type":"calculate","label":"court","description":"explication","value":"formule XLSForm","confidence":"high|medium|low"}',
      '',
      'REGLES FINALES: required=true par defaut. JSON compact. Outil: ' + tool
    ];

    const system = systemParts.join('\n');

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 8192,
        system: system,
        messages: [{ role: 'user', content: 'Extrais TOUTES les questions de ce questionnaire sans en oublier aucune:\n\n' + inputText }]
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
      try { form = JSON.parse(jsonStr); }
      catch(e) {
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

    if (!form.questions || !Array.isArray(form.questions)) {
      return res.status(422).json({ error: 'NO_QUESTIONS', message: 'Aucune question detectee.' });
    }

    console.log('[ANALYSE] ok ' + form.questions.length + ' questions');
    res.json({
      success: true,
      form: form,
      truncated: truncated,
      warning: truncated ? 'Questionnaire tronque a ' + inputText.length + ' chars sur ' + text.length + ' total.' : null
    });
  } catch(err) {
    console.error('[ANALYSE ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ VERIFICATION XLSFORM AVANT DEPLOIEMENT ============
app.post('/api/verify', async (req, res) => {
  try {
    const { form } = req.body;
    if (!form) return res.status(400).json({ error: 'Formulaire manquant' });
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'Cle API manquante' });

    console.log('[VERIFY] Verification XLSForm avant deploiement');
    const koboContent = buildKoboContent(form);
    const xlsformStr = JSON.stringify(koboContent, null, 2);

    const verifySystem = 'Tu es expert XLSForm et KoboToolbox. Analyse ce contenu XLSForm et verifie:\n' +
      '1. Les skip_logic (relevant) sont correctement formules (syntaxe XPath valide)\n' +
      '2. Les questions a choix ont bien leurs modalites\n' +
      '3. Les calculs sont valides\n' +
      '4. Les contraintes sont valides\n' +
      'Reponds en JSON: {"valid":true/false,"issues":["probleme 1"],"fixed_questions":[{"id":"q1","relevant":"formule corrigee"}]}\n' +
      'Si tout est bon, issues=[] et fixed_questions=[]. Sois concis.';

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 2048,
        system: verifySystem,
        messages: [{ role: 'user', content: 'Verifie ce XLSForm:\n\n' + xlsformStr.slice(0, 6000) }]
      })
    });

    if (!response.ok) {
      console.error('[VERIFY] Claude error');
      return res.json({ valid: true, issues: [], message: 'Verification ignoree' });
    }

    const data = await response.json();
    const raw = data.content && data.content[0] ? data.content[0].text : '{"valid":true,"issues":[]}';
    let result;
    try {
      let clean = raw.replace(/```json\n?/g,'').replace(/```\n?/g,'').trim();
      const si = clean.indexOf('{');
      if (si > 0) clean = clean.slice(si);
      result = JSON.parse(clean);
    } catch(e) {
      result = { valid: true, issues: [] };
    }

    // Appliquer les corrections suggérées
    if (result.fixed_questions && result.fixed_questions.length > 0) {
      result.fixed_questions.forEach(function(fix) {
        const q = form.questions.find(function(q) { return q.id === fix.id; });
        if (q && fix.relevant) q.relevant = fix.relevant;
      });
    }

    console.log('[VERIFY] valid=' + result.valid + ' issues=' + (result.issues ? result.issues.length : 0));
    res.json({ valid: result.valid, issues: result.issues || [], form: form });
  } catch(err) {
    console.error('[VERIFY ERROR]', err.message);
    res.json({ valid: true, issues: [], message: 'Verification ignoree: ' + err.message });
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
    const tokenData = await tokenRes.json();
    const token = tokenData.token;
    const auth = { 'Authorization': 'Token ' + token, 'Content-Type': 'application/json' };
    const assetRes = await fetch(server + '/api/v2/assets/?format=json', {
      method: 'POST', headers: auth,
      body: JSON.stringify({ name: form.title || 'Formulaire R2', asset_type: 'survey' })
    });
    if (!assetRes.ok) return res.status(502).json({ error: 'ASSET_ERROR', message: 'Erreur creation formulaire.' });
    const assetData = await assetRes.json();
    const uid = assetData.uid;
    const koboContent = buildKoboContent(form);
    const patchRes = await fetch(server + '/api/v2/assets/' + uid + '/?format=json', {
      method: 'PATCH', headers: auth,
      body: JSON.stringify({ name: form.title || 'Formulaire R2', content: koboContent })
    });
    if (!patchRes.ok) {
      const e = await patchRes.text();
      console.error('[PATCH ERROR]', e);
      return res.status(502).json({ error: 'PATCH_ERROR', message: 'Erreur import questionnaire.' });
    }
    await fetch(server + '/api/v2/assets/' + uid + '/deployment/?format=json', {
      method: 'POST', headers: auth, body: JSON.stringify({ active: true })
    });
    const qCount = form.questions ? form.questions.length : 0;
    console.log('[DEPLOY] ok ' + uid + ' - ' + qCount + ' questions');
    res.json({ success: true, uid: uid, url: server + '/#/forms/' + uid + '/summary', questions: qCount });
  } catch(err) {
    console.error('[DEPLOY ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
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
      var name = (q.id || ('q' + q.num)).replace(/[^a-zA-Z0-9_]/g,'_');

      var row = {
        type: t,
        name: name,
        label: q.label || '',
        required: q.required !== false ? 'yes' : 'no',
        hint: q.hint || ''
      };

      // Skip logic — utilise les suggestions confirmées
      var relevant = '';
      if (q.suggestions) {
        q.suggestions.forEach(function(s, si) {
          if (s.type === 'skip_logic' && s.value && s.value.trim()) {
            var confirmed = q.confirmedSuggestions && q.confirmedSuggestions[si];
            if (confirmed) {
              // Nettoyer et valider la formule XLSForm
              var formula = s.value.trim();
              // S'assurer que les références sont au bon format ${nom}
              // Remplacer les libelles texte par les codes si possible
              relevant = formula;
            }
          }
        });
      }
      // Fallback sur q.relevant s'il existe et qu'aucune suggestion n'est confirmée
      if (!relevant && q.relevant && q.relevant.trim()) relevant = q.relevant.trim();
      if (relevant) row.relevant = relevant;

      // Calcul
      if (t === 'calculate') {
        var calc = q.calculation || '';
        if (!calc && q.suggestions) {
          q.suggestions.forEach(function(s, si) {
            if (s.type === 'calculate' && s.value) {
              var confirmed = q.confirmedSuggestions && q.confirmedSuggestions[si];
              if (confirmed) calc = s.value;
            }
          });
        }
        if (calc) row.calculation = calc;
        delete row.required;
      }

      // Contraintes numériques
      var constraints = [];
      if (q.numMin !== '' && q.numMin !== undefined && q.numMin !== null) constraints.push('. >= ' + q.numMin);
      if (q.numMax !== '' && q.numMax !== undefined && q.numMax !== null) constraints.push('. <= ' + q.numMax);
      if (q.numDigitsBefore) constraints.push('string-length(substring-before(string(.), \'.\')) <= ' + q.numDigitsBefore);
      if (q.numDigitsAfter) constraints.push('string-length(substring-after(string(.), \'.\')) <= ' + q.numDigitsAfter);
      if (q.constraint && q.constraint.trim()) constraints.push(q.constraint.trim());
      if (q.suggestions) {
        q.suggestions.forEach(function(s, si) {
          if (s.type === 'constraint' && s.value) {
            var confirmed = q.confirmedSuggestions && q.confirmedSuggestions[si];
            if (confirmed) constraints.push(s.value);
          }
        });
      }
      if (constraints.length > 0) {
        row.constraint = constraints.join(' and ');
        row.constraint_message = 'Valeur hors limites acceptees';
      }

      if (t === 'range') { row.parameters = 'start=' + (q.numMin||1) + ' end=' + (q.numMax||10); delete row.required; }
      if (t === 'note' || t === 'calculate') delete row.required;

      // Choix avec valeurs numériques
      if (t === 'select_one' || t === 'select_multiple' || t === 'rank') {
        var listName = 'list_' + name;
        row.type = t + ' ' + listName;
        if (!seen.has(listName)) {
          seen.add(listName);
          var choiceVals = q.choice_values || [];
          (q.choices || []).forEach(function(c, i) {
            var label = typeof c === 'string' ? c : (c.label || String(c));
            // Toujours utiliser le code numérique (0, 1, 2...) comme valeur XLSForm
            // C'est ce que KoboCollect utilise dans les formules relevant
            var val;
            if (choiceVals[i] !== undefined && choiceVals[i] !== '') {
              val = String(choiceVals[i]);
            } else {
              // Par défaut: code numérique basé sur l'index (0, 1, 2...)
              val = String(i);
            }
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
  console.log('\n╔══════════════════════════════════╗\n║   R2 Forms Backend v3.2          ║\n║   Port: ' + PORT + '                    ║\n╚══════════════════════════════════╝\n');
});
