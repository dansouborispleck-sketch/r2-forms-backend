require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const pdf = require('pdf-parse');
const fetch = require('node-fetch');
const path = require('path');
const cloudinary = require('cloudinary').v2;
const FormData = require('form-data');

// Config Cloudinary
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME || 'tk6wqscg',
  api_key: process.env.CLOUDINARY_API_KEY || '682714838266378',
  api_secret: process.env.CLOUDINARY_API_SECRET || 'GF3nuz5v7kvuwrcdJUJHqzEzv3g'
});

const app = express();
const PORT = process.env.PORT || 4000;

app.use(cors({ origin: '*' }));
app.use(express.json({ limit: '50mb' }));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
app.use(express.static(path.join(__dirname, 'public')));

app.get('/health', (req, res) => res.json({ status: 'ok', version: '5.0.0' }));

// ============ IMPORT ============
app.post('/api/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier recu' });
    const { originalname, buffer } = req.file;
    const ext = originalname.split('.').pop().toLowerCase();
    let text = '';
    console.log('[IMPORT] ' + originalname + ' (' + buffer.length + ' bytes)');

    if (['txt','csv'].includes(ext)) {
      // Essayer UTF-8 d'abord, puis latin1
      try {
        text = buffer.toString('utf-8');
      } catch(e) {
        text = buffer.toString('latin1');
      }
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

    // Normaliser le texte — préserver les accents
    text = text
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();

    if (text.length < 20)
      return res.status(422).json({ error: 'EMPTY', message: 'Fichier vide. Collez le texte directement.' });

    console.log('[IMPORT] ok ' + text.length + ' chars, accents: ' + (text.includes('é') || text.includes('è') || text.includes('à')));
    res.json({ success: true, text: text, metadata: { filename: originalname, chars: text.length } });
  } catch(err) {
    console.error('[IMPORT ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ ANALYSE (CHUNKING COMPLET) ============
app.post('/api/analyse', async (req, res) => {
  try {
    const { text, tool } = req.body;
    if (!text || text.trim().length < 10) return res.status(400).json({ error: 'Texte trop court' });
    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'Cle API manquante' });

    // Chunking — morceaux de 6000 chars, découpe aux sauts de ligne
    const CHUNK_SIZE = 6000;
    const chunks = [];
    let pos = 0;
    while (pos < text.length) {
      let end = Math.min(pos + CHUNK_SIZE, text.length);
      if (end < text.length) {
        const lastNL = text.lastIndexOf('\n', end);
        if (lastNL > pos + CHUNK_SIZE * 0.6) end = lastNL;
      }
      chunks.push(text.slice(pos, end));
      pos = end;
    }

    console.log('[ANALYSE] ' + text.length + ' chars -> ' + chunks.length + ' chunks -> ' + tool);

    const system = 'Expert en collecte de donnees terrain. Analyse le questionnaire et extrais TOUTES les questions et le texte introductif.\n\n' +
'REGLES ABSOLUES:\n' +
'1. TEXTE AVANT LA PREMIERE QUESTION: Tout texte apparaissant AVANT la premiere question numerotee (avant "1)" ou "Q1" ou "1.") doit IMPERATIVEMENT etre inclus comme premiere question de type "note". REGLES STRICTES: (a) Le label = le texte INTEGRAL copie mot pour mot, JAMAIS un titre invente comme "Introduction" ou "Presentation"; (b) Ne jamais ignorer ce texte meme s\'il semble long; (c) Si le questionnaire n\'a pas de texte introductif, ne pas creer de note fictive.\n' +
'2. Extrais TOUTES les questions: numerotees, sans numero, sous-questions implicites, questions en entete.\n' +
'3. GROUPES: Utilise le TITRE COMPLET de la section.\n' +
'4. ACCENTS: Preserve absolument tous les accents (é, è, à, ù, â, ê, î, ô, û, ç, etc.).\n' +
'5. CHAMP AUTRES: Quand une modalite contient "Autres" ou "Autre", cree automatiquement une question "Si autre, precisez :" juste apres (required=false).\n' +
'6. SAUTS: Detecte les conditions. Pour skip_logic, utilise les vrais IDs des questions (format q1, q2...), PAS des IDs inventes.\n' +
'7. PAS de suggestions de type calculate — le statisticien s\'en charge.\n' +
'8. coherence_report: observations utiles.\n\n' +
'FORMAT JSON COMPACT:\n' +
'{"title":"titre","coherence_report":["obs"],"questions":[{"id":"q1","num":1,"label":"libelle","question_class":"CLASS","type":"TYPE","required":true,"hint":"","choices":[],"choice_values":[],"group":"TITRE COMPLET","formats":[],"suggested_format_idx":0,"suggestions":[]}],"groups":[]}\n\n' +
'CLASSES ET FORMATS OBLIGATOIRES:\n' +
'- quantitative -> formats:[{id:"A",name:"Nombre entier",type:"integer",note:"Ex: 25"},{id:"B",name:"Nombre decimal",type:"decimal",note:"Ex: 65,5"},{id:"C",name:"Echelle",type:"range",note:"Ex: 1 a 10"}], suggested_format_idx:0\n' +
'- qualitative_choice -> formats:[{id:"A",name:"Choix unique",type:"select_one",note:"Une seule reponse"},{id:"B",name:"Choix multiple",type:"select_multiple",note:"Plusieurs reponses"},{id:"C",name:"Texte libre",type:"text",note:"Reponse libre"}], suggested_format_idx:0 (ou 1 si "plusieurs reponses" mentionnes)\n' +
'- qualitative_open -> formats:[{id:"A",name:"Reponse courte",type:"text",note:"Quelques mots"},{id:"B",name:"Reponse longue",type:"text",note:"Plusieurs phrases"}], suggested_format_idx:0\n' +
'- date_time -> formats:[{id:"A",name:"Date",type:"date",note:"jj/mm/aaaa"},{id:"B",name:"Heure",type:"time",note:"hh:mm"},{id:"C",name:"Date et heure",type:"datetime",note:"jj/mm/aaaa hh:mm"}], suggested_format_idx:0\n' +
'- geopoint -> formats:[{id:"A",name:"GPS",type:"geopoint",note:"Capture GPS auto"}], suggested_format_idx:0\n' +
'- media_photo -> formats:[{id:"A",name:"Photo",type:"image",note:"Photo"}], suggested_format_idx:0\n' +
'- media_audio -> formats:[{id:"A",name:"Audio",type:"audio",note:"Enregistrement"}], suggested_format_idx:0\n' +
'- media_video -> formats:[{id:"A",name:"Video",type:"video",note:"Video"}], suggested_format_idx:0\n' +
'- media_file -> formats:[{id:"A",name:"Fichier",type:"file",note:"Piece jointe"}], suggested_format_idx:0\n' +
'- barcode -> formats:[{id:"A",name:"Code-barres",type:"barcode",note:"Scan"}], suggested_format_idx:0\n' +
'- acknowledge -> formats:[{id:"A",name:"Confirmation",type:"acknowledge",note:"Case a cocher"}], suggested_format_idx:0\n' +
'- ranking -> formats:[{id:"A",name:"Classement",type:"rank",note:"Ordre de preference"}], suggested_format_idx:0\n' +
'- scale -> formats:[{id:"A",name:"Echelle",type:"range",note:"Curseur"}], suggested_format_idx:0\n' +
'- note -> formats:[{id:"A",name:"Note",type:"note",note:"Texte sans saisie"}], suggested_format_idx:0\n' +
'IMPORTANT: Chaque question DOIT avoir son tableau formats[] rempli selon sa classe. Ne jamais laisser formats:[]\n\n' +
'SUGGESTIONS (skip_logic, constraint et audio uniquement):\n' +
'AUDIO: Si la question demande une description libre, une narration, un recit, une explication detaillee, une opinion longue, ou tout contenu qui serait mieux capture par la voix (ex: "Decrivez...", "Racontez...", "Expliquez...", "Donnez votre opinion sur...", "Comment decririez-vous...", "Quelles sont vos impressions..."), ajoute une suggestion de type audio:\n' +
'{"type":"audio","label":"Enregistrement audio","description":"Cette question peut etre mieux repondue par enregistrement vocal","confidence":"high"}\n' +
'skip_logic et constraint:\n' +
'{"type":"skip_logic|constraint","label":"court","description":"clair","value":"formule XLSForm avec vrais IDs","confidence":"high|medium|low"}\n\n' +
'Outil cible: ' + tool;

    async function analyseChunk(chunkText, chunkNum, totalChunks, previousIds) {
      const contextNote = totalChunks > 1
        ? '\n\nPartie ' + chunkNum + '/' + totalChunks + '.' +
          (chunkNum === 1 ? ' Extrait depuis le debut.' : ' Continue — ne repete pas les questions deja extraites.') +
          (chunkNum === totalChunks ? ' Derniere partie.' : '') +
          (previousIds.length > 0 ? '\nDernieres questions extraites: ' + previousIds.slice(-8).join(', ') : '')
        : '';

      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
        body: JSON.stringify({
          model: 'claude-sonnet-4-5',
          max_tokens: 32000,
          system: system + contextNote,
          messages: [{ role: 'user', content: 'Extrais TOUTES les questions de cette partie du questionnaire:\n\n' + chunkText }]
        })
      });

      if (!response.ok) {
        const e = await response.json();
        throw new Error('Claude: ' + (e.error && e.error.message || JSON.stringify(e).slice(0, 200)));
      }

      const data = await response.json();
      const rawText = data.content && data.content[0] ? data.content[0].text : '{}';
      let jsonStr = rawText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      const s = jsonStr.indexOf('{');
      if (s > 0) jsonStr = jsonStr.slice(s);

      try {
        return JSON.parse(jsonStr);
      } catch(e) {
        // Tenter de reparer le JSON tronque
        const lc = jsonStr.lastIndexOf('},');
        if (lc > 100) jsonStr = jsonStr.slice(0, lc + 1);
        jsonStr = jsonStr.replace(/,\s*$/, '');
        const ob = (jsonStr.match(/\[/g) || []).length - (jsonStr.match(/\]/g) || []).length;
        const cb = (jsonStr.match(/\{/g) || []).length - (jsonStr.match(/\}/g) || []).length;
        for (var i = 0; i < Math.max(0, ob); i++) jsonStr += ']';
        for (var j = 0; j < Math.max(0, cb); j++) jsonStr += '}';
        return JSON.parse(jsonStr);
      }
    }

    let allQuestions = [];
    let formTitle = '';
    let allGroups = new Set();
    let coherenceReport = [];
    let previousIds = [];
    const failedChunks = [];

    for (var ci = 0; ci < chunks.length; ci++) {
      var chunkNum = ci + 1;
      console.log('[ANALYSE] Chunk ' + chunkNum + '/' + chunks.length + ' (' + chunks[ci].length + ' chars)');
      var success = false;

      for (var attempt = 1; attempt <= 3; attempt++) {
        try {
          if (attempt > 1) await new Promise(function(r) { setTimeout(r, 1500 * attempt); });
          var result = await analyseChunk(chunks[ci], chunkNum, chunks.length, previousIds);

          if (chunkNum === 1 && result.title) formTitle = result.title;
          if (result.coherence_report) result.coherence_report.forEach(function(r) { coherenceReport.push(r); });

          if (result.questions && Array.isArray(result.questions) && result.questions.length > 0) {
            // Deduplication par label
            var existingLabels = new Set(allQuestions.map(function(q) { return (q.label || '').trim().toLowerCase(); }));
            var newQs = result.questions.filter(function(q) {
              var lbl = (q.label || '').trim().toLowerCase();
              return lbl.length > 0 && !existingLabels.has(lbl);
            });
            // Renumeroter
            newQs.forEach(function(q, idx) {
              q.num = allQuestions.length + idx + 1;
              q.id = 'q' + q.num;
            });
            allQuestions = allQuestions.concat(newQs);
            previousIds = allQuestions.slice(-10).map(function(q) { return q.id; });
          }

          if (result.groups) result.groups.forEach(function(g) { allGroups.add(g); });
          console.log('[ANALYSE] Chunk ' + chunkNum + ' ok -> ' + (result.questions ? result.questions.length : 0) + ' nouvelles questions. Total: ' + allQuestions.length);
          success = true;
          break;
        } catch(chunkErr) {
          console.error('[ANALYSE] Chunk ' + chunkNum + ' attempt ' + attempt + ' failed:', chunkErr.message);
        }
      }

      if (!success) {
        failedChunks.push(chunkNum);
        console.error('[ANALYSE] Chunk ' + chunkNum + ' abandonne apres 3 tentatives');
      }

      // Pause entre chunks pour eviter rate limiting
      if (ci < chunks.length - 1) await new Promise(function(r) { setTimeout(r, 600); });
    }

    if (allQuestions.length === 0)
      return res.status(422).json({ error: 'NO_QUESTIONS', message: 'Aucune question detectee.' });

    var form = {
      title: formTitle || 'Formulaire',
      coherence_report: coherenceReport,
      questions: allQuestions,
      groups: Array.from(allGroups)
    };

    console.log('[ANALYSE] TOTAL: ' + allQuestions.length + ' questions, ' + chunks.length + ' chunks, ' + failedChunks.length + ' echecs');

    res.json({
      success: true,
      form: form,
      chunks_total: chunks.length,
      chunks_failed: failedChunks.length,
      warning: failedChunks.length > 0
        ? failedChunks.length + ' partie(s) non traitee(s) (chunks ' + failedChunks.join(', ') + '). Verifiez que toutes les questions sont presentes.'
        : null
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
        model: 'claude-sonnet-4-5', max_tokens: 8192,
        system: 'Expert collecte de donnees. Applique exactement les corrections au formulaire JSON. Preserve tous les accents. Retourne JSON corrige UNIQUEMENT sans markdown.',
        messages: [{ role: 'user', content: 'Formulaire:\n' + JSON.stringify(form) + '\n\nCorrections:\n' + instructions }]
      })
    });
    if (!response.ok) return res.status(502).json({ error: 'CLAUDE_ERROR' });
    const data = await response.json();
    let raw = data.content && data.content[0] ? data.content[0].text : '{}';
    raw = raw.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const m = raw.match(/\{[\s\S]*/);
    const corrected = JSON.parse(m ? m[0] : '{}');
    res.json({ success: true, form: corrected });
  } catch(err) {
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ VALIDATION ET AUTO-CORRECTION XLSFORM ============
app.post('/api/verify', async (req, res) => {
  try {
    const { form } = req.body;
    if (!form) return res.status(400).json({ error: 'Formulaire manquant' });

    // Auto-correction du formulaire avant de construire le XLSForm
    var questions = form.questions || [];
    var validQIds = new Set(questions.map(function(q) {
      return (q.id || ('q' + q.num)).replace(/[^a-zA-Z0-9_]/g, '_');
    }));

    var fixCount = 0;

    questions.forEach(function(q) {
      // 1. Corriger les sauts avec IDs invalides
      if (q.relevant) {
        var refs = q.relevant.match(/\$\{([^}]+)\}/g) || [];
        var needsFix = false;
        refs.forEach(function(ref) {
          var refId = ref.replace('${', '').replace('}', '');
          if (!validQIds.has(refId)) needsFix = true;
        });
        if (needsFix) {
          // Tenter de corriger par similarite de numero
          var fixed = q.relevant;
          refs.forEach(function(ref) {
            var refId = ref.replace('${', '').replace('}', '');
            if (!validQIds.has(refId)) {
              // Chercher par numero dans l'ID
              var numMatch = refId.match(/\d+/);
              if (numMatch) {
                var num = parseInt(numMatch[0]);
                var target = questions.find(function(qq) { return qq.num === num; });
                if (target) {
                  var targetId = (target.id || ('q' + target.num)).replace(/[^a-zA-Z0-9_]/g, '_');
                  fixed = fixed.replace('${' + refId + '}', '${' + targetId + '}');
                  fixCount++;
                  console.log('[VERIFY] Saut corrige par num: ${' + refId + '} -> ${' + targetId + '}');
                } else {
                  // Pas de correspondance par numero -> chercher par similarite de label
                  var bestMatch = null;
                  var bestScore = 0;
                  questions.forEach(function(qq) {
                    var qqId = (qq.id || ('q' + qq.num)).replace(/[^a-zA-Z0-9_]/g, '_');
                    // Similarite entre refId et qqId
                    var shorter = refId.length < qqId.length ? refId : qqId;
                    var longer = refId.length < qqId.length ? qqId : refId;
                    var score = 0;
                    for (var ci = 0; ci < shorter.length; ci++) {
                      if (shorter[ci] === longer[ci]) score++;
                    }
                    score = score / longer.length;
                    if (score > bestScore && score > 0.4) { bestScore = score; bestMatch = qqId; }
                  });
                  if (bestMatch) {
                    fixed = fixed.replace('${' + refId + '}', '${' + bestMatch + '}');
                    fixCount++;
                    console.log('[VERIFY] Saut corrige par similarite: ${' + refId + '} -> ${' + bestMatch + '} (score:' + bestScore.toFixed(2) + ')');
                  } else {
                    // Dernier recours: utiliser la question precedente dans le formulaire
                    var qIdx = questions.findIndex(function(qq) { return qq.relevant === q.relevant; });
                    if (qIdx > 0) {
                      var prevId = (questions[qIdx-1].id || ('q' + questions[qIdx-1].num)).replace(/[^a-zA-Z0-9_]/g, '_');
                      fixed = fixed.replace('${' + refId + '}', '${' + prevId + '}');
                      fixCount++;
                      console.log('[VERIFY] Saut corrige par question precedente: ${' + refId + '} -> ${' + prevId + '}');
                    } else {
                      // Utiliser q1 comme fallback absolu — jamais supprimer
                      var fallbackId = (questions[0].id || 'q1').replace(/[^a-zA-Z0-9_]/g, '_');
                      fixed = fixed.replace('${' + refId + '}', '${' + fallbackId + '}');
                      fixCount++;
                      console.log('[VERIFY] Saut corrige par fallback: ${' + refId + '} -> ${' + fallbackId + '}');
                    }
                  }
                }
              } else {
                // Pas de numero dans refId -> chercher par similarite de label
                var bestMatch2 = null;
                var bestScore2 = 0;
                questions.forEach(function(qq) {
                  var qqId = (qq.id || ('q' + qq.num)).replace(/[^a-zA-Z0-9_]/g, '_');
                  var qqLabel = (qq.label || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'_');
                  // Comparer refId avec qqId et qqLabel
                  var score1 = 0, score2 = 0;
                  var shorter1 = refId.length < qqId.length ? refId : qqId;
                  var longer1 = refId.length < qqId.length ? qqId : refId;
                  for (var ci = 0; ci < shorter1.length; ci++) { if (shorter1[ci] === longer1[ci]) score1++; }
                  score1 = score1 / Math.max(longer1.length, 1);
                  // Chercher refId dans le label normalise
                  if (qqLabel.includes(refId) || refId.includes(qqLabel.slice(0, 6))) score2 = 0.7;
                  var score = Math.max(score1, score2);
                  if (score > bestScore2 && score > 0.35) { bestScore2 = score; bestMatch2 = qqId; }
                });
                if (bestMatch2) {
                  fixed = fixed.replace('${' + refId + '}', '${' + bestMatch2 + '}');
                  fixCount++;
                  console.log('[VERIFY] Saut corrige par label: ${' + refId + '} -> ${' + bestMatch2 + '} (score:' + bestScore2.toFixed(2) + ')');
                } else {
                  // Fallback absolu — prendre la question juste avant celle qui a ce saut
                  var qIdx2 = questions.findIndex(function(qq) { return qq.relevant === q.relevant; });
                  var prevIdx = qIdx2 > 0 ? qIdx2 - 1 : 0;
                  var prevId2 = (questions[prevIdx].id || ('q' + questions[prevIdx].num)).replace(/[^a-zA-Z0-9_]/g, '_');
                  fixed = fixed.replace('${' + refId + '}', '${' + prevId2 + '}');
                  fixCount++;
                  console.log('[VERIFY] Saut corrige par fallback absolu: ${' + refId + '} -> ${' + prevId2 + '}');
                }
              }
            }
          });
          q.relevant = fixed;
        }
      }

      // 2. Corriger les questions select sans choix -> convertir en texte libre
      var t = q.selectedType || q.type || 'text';
      if ((t === 'select_one' || t === 'select_multiple') && (!q.choices || q.choices.length === 0)) {
        console.log('[VERIFY] Question sans choix convertie en texte: ' + q.label);
        q.selectedType = 'text';
        q.type = 'text';
        fixCount++;
      }

      // 3. Corriger les suggestions skip_logic avec IDs invalides
      if (q.suggestions) {
        q.suggestions.forEach(function(s) {
          if (s.type === 'skip_logic' && s.value) {
            var sRefs = s.value.match(/\$\{([^}]+)\}/g) || [];
            var hasInvalid = sRefs.some(function(ref) {
              return !validQIds.has(ref.replace('${','').replace('}',''));
            });
            if (hasInvalid) {
              s.value = ''; // Vider la suggestion invalide
              fixCount++;
            }
          }
        });
      }
    });

    console.log('[VERIFY] Auto-correction: ' + fixCount + ' corrections appliquees');
    res.json({ valid: true, issues: [], form: form, fixed: fixCount });

  } catch(err) {
    console.error('[VERIFY ERROR]', err.message);
    // Ne jamais bloquer — retourner le formulaire tel quel
    res.json({ valid: true, issues: [], form: req.body.form });
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
      body: JSON.stringify({ name: form.title || 'Formulaire DIGUE', asset_type: 'survey' })
    });
    if (!assetRes.ok) return res.status(502).json({ error: 'ASSET_ERROR', message: 'Erreur creation formulaire.' });
    const uid = (await assetRes.json()).uid;

    const koboContent = buildKoboContent(form);

    // Collecter les images pour les retourner au frontend (téléchargement ZIP automatique)
    const imageAttachments = [];
    (form.questions || []).forEach(function(q) {
      var choiceImages = q.choiceImages || {};
      Object.keys(choiceImages).forEach(function(ci) {
        var img = choiceImages[ci];
        if (img && img.url) {
          var choices = q.choices || [];
          var label = typeof choices[ci] === 'string' ? choices[ci] : (choices[ci] && choices[ci].label) || ('choice_' + ci);
          var key = label.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'_').replace(/_+/g,'_').slice(0,50);
          imageAttachments.push({ url: img.url, filename: key + '.png', label: label });
        }
      });
    });
    if (imageAttachments.length > 0) {
      console.log('[DEPLOY] ' + imageAttachments.length + ' images a telecharger par utilisateur');
    }

    const patchRes = await fetch(server + '/api/v2/assets/' + uid + '/?format=json', {
      method: 'PATCH', headers: auth,
      body: JSON.stringify({ name: form.title || 'Formulaire DIGUE', content: koboContent })
    });
    if (!patchRes.ok) {
      const errText = await patchRes.text();
      console.error('[PATCH ERROR]', errText.slice(0, 300));
      return res.status(502).json({ error: 'PATCH_ERROR', message: 'Erreur import questionnaire. Verifiez le XLSForm.' });
    }

    // Déployer le formulaire
    await fetch(server + '/api/v2/assets/' + uid + '/deployment/?format=json', {
      method: 'POST', headers: auth, body: JSON.stringify({ active: true })
    });



    console.log('[DEPLOY] Kobo ok ' + uid);
    res.json({
      success: true,
      uid: uid,
      url: server + '/#/forms/' + uid + '/summary',
      questions: (form.questions || []).length,
      images: imageAttachments  // URLs des images pour téléchargement ZIP côté frontend
    });
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
    const userData = await userRes.json();
    if (!userRes.ok || userData.responseCode !== 200)
      return res.status(401).json({ error: 'AUTH_ERROR', message: 'Cle API JotForm incorrecte ou insuffisante (Full Access requis).' });

    const questions = form.questions || [];
    const jotformQuestions = {};
    jotformQuestions['0'] = { type: 'control_head', text: form.title || 'Formulaire DIGUE', order: '1', name: 'header' };

    questions.forEach(function(q, qi) {
      var t = q.selectedType || q.type || 'text';
      var qType = 'control_textbox';
      if (t === 'select_one') qType = 'control_radio';
      else if (t === 'select_multiple') qType = 'control_checkbox';
      else if (t === 'integer' || t === 'decimal') qType = 'control_number';
      else if (t === 'date') qType = 'control_datetime';
      else if (t === 'image') qType = 'control_fileupload';
      else if (t === 'note') qType = 'control_text';

      var qObj = { type: qType, text: q.label || ('Question ' + (qi+1)), order: String(qi+2), name: 'q' + (qi+1), required: q.required !== false ? 'Yes' : 'No' };
      if ((qType === 'control_radio' || qType === 'control_checkbox') && q.choices && q.choices.length > 0) {
        qObj.options = q.choices.map(function(c) { return typeof c === 'string' ? c : (c.label || String(c)); }).join('|');
      }
      jotformQuestions[String(qi+1)] = qObj;
    });

    const createBody = new URLSearchParams();
    createBody.append('properties[title]', form.title || 'Formulaire DIGUE');
    Object.keys(jotformQuestions).forEach(function(key) {
      var q = jotformQuestions[key];
      Object.keys(q).forEach(function(field) {
        createBody.append('questions[' + key + '][' + field + ']', q[field]);
      });
    });

    const createRes = await fetch('https://api.jotform.com/form?apiKey=' + apiKey, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: createBody.toString()
    });
    const createData = await createRes.json();
    if (!createRes.ok || createData.responseCode !== 200)
      return res.status(502).json({ error: 'CREATE_ERROR', message: 'Erreur creation JotForm: ' + (createData.message || createData.responseCode) });

    const formId = createData.content && createData.content.id;
    if (!formId) return res.status(502).json({ error: 'CREATE_ERROR', message: 'ID formulaire JotForm non recu.' });

    console.log('[DEPLOY] JotForm ok ' + formId);
    res.json({ success: true, formId: formId, url: 'https://www.jotform.com/build/' + formId, questions: questions.length });
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

    const createRes = await fetch('https://forms.googleapis.com/v1/forms', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
      body: JSON.stringify({ info: { title: form.title || 'Formulaire DIGUE', documentTitle: form.title || 'Formulaire DIGUE' } })
    });
    if (!createRes.ok) return res.status(502).json({ error: 'CREATE_ERROR', message: 'Erreur creation Google Form.' });
    const formId = (await createRes.json()).formId;

    const groupMap = {}, groupOrder = [];
    questions.forEach(function(q) {
      var g = q.group || 'General';
      if (!groupMap[g]) { groupMap[g] = []; groupOrder.push(g); }
      groupMap[g].push(q);
    });

    const buildRequests = [], sectionItems = [];
    var itemIdx = 0;

    function makeQItem(q) {
      var t = q.selectedType || q.type || 'text';
      var required = q.required !== false;
      var choices = (q.choices || []).map(function(c) { return { value: typeof c === 'string' ? c : (c.label || String(c)) }; });
      if (t === 'select_one' && choices.length > 0) return { question: { required: required, choiceQuestion: { type: 'RADIO', options: choices, shuffle: false } } };
      if (t === 'select_multiple' && choices.length > 0) return { question: { required: required, choiceQuestion: { type: 'CHECKBOX', options: choices, shuffle: false } } };
      if (t === 'date') return { question: { required: required, dateQuestion: { includeTime: false, includeYear: true } } };
      if (t === 'time' || t === 'datetime') return { question: { required: required, timeQuestion: { duration: false } } };
      if (t === 'scale' || t === 'range') return { question: { required: required, scaleQuestion: { low: parseInt(q.numMin) || 1, high: parseInt(q.numMax) || 10, lowLabel: 'Min', highLabel: 'Max' } } };
      return { question: { required: required, textQuestion: { paragraph: (q.label || '').length > 50 } } };
    }

    function isAutre(q) {
      if (!q.label) return false;
      var l = q.label.toLowerCase();
      return l.includes('si autre') || l.includes('precisez') || l.includes('preciser');
    }

    groupOrder.forEach(function(gname) {
      if (gname !== 'general' && gname !== 'General') {
        buildRequests.push({ createItem: { item: { title: gname, pageBreakItem: {} }, location: { index: itemIdx } } });
        sectionItems.push({ type: 'section', name: gname, idx: itemIdx });
        itemIdx++;
      }
      groupMap[gname].forEach(function(q) {
        if (isAutre(q)) {
          buildRequests.push({ createItem: { item: { title: '', pageBreakItem: {} }, location: { index: itemIdx } } });
          sectionItems.push({ type: 'section', name: '_a' + itemIdx, idx: itemIdx, isAutreSection: true });
          itemIdx++;
        }
        var hint = q.hint || '';
        var t = q.selectedType || q.type || 'text';
        if ((t === 'integer' || t === 'decimal') && (q.numMin || q.numMax)) {
          var parts = [];
          if (q.numMin) parts.push('min: ' + q.numMin);
          if (q.numMax) parts.push('max: ' + q.numMax);
          if (parts.length) hint = (hint ? hint + ' — ' : '') + parts.join(', ');
        }
        buildRequests.push({ createItem: { item: { title: q.label || ('Q' + (itemIdx+1)), description: hint, questionItem: makeQItem(q) }, location: { index: itemIdx } } });
        sectionItems.push({ type: 'question', q: q, idx: itemIdx, isAutre: isAutre(q) });
        itemIdx++;
      });
    });
    buildRequests.push({ createItem: { item: { title: '', pageBreakItem: {} }, location: { index: itemIdx } } });
    sectionItems.push({ type: 'section', name: '_final', idx: itemIdx, isFinal: true });

    const batch1Res = await fetch('https://forms.googleapis.com/v1/forms/' + formId + ':batchUpdate', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
      body: JSON.stringify({ requests: buildRequests, includeFormInResponse: true })
    });
    if (!batch1Res.ok) {
      const e = await batch1Res.json();
      return res.status(502).json({ error: 'BATCH_ERROR', message: 'Erreur ajout questions Google Forms.' });
    }
    const updatedForm = (await batch1Res.json()).form;

    if (updatedForm && updatedForm.items) {
      const orderedItems = updatedForm.items.slice().sort(function(a, b) { return (a.index || 0) - (b.index || 0); });
      orderedItems.forEach(function(item, i) { if (sectionItems[i]) sectionItems[i].googleItemId = item.itemId; });

      const navRequests = [];
      sectionItems.forEach(function(si, idx) {
        if (si.type !== 'question') return;
        var nextSi = sectionItems[idx + 1], nextNextSi = sectionItems[idx + 2];
        if (!nextSi || !nextSi.isAutreSection || !nextNextSi || !nextNextSi.isAutre) return;
        var q = si.q;
        var autreIdx = -1;
        (q.choices || []).forEach(function(c, ci) {
          var l = (typeof c === 'string' ? c : c.label || '').toLowerCase();
          if (l.includes('autre') || l.includes('other')) autreIdx = ci;
        });
        if (autreIdx < 0 || !nextSi.googleItemId || !si.googleItemId) return;
        var skipId = null;
        for (var i2 = idx + 3; i2 < sectionItems.length; i2++) {
          if (sectionItems[i2].type === 'section') { skipId = sectionItems[i2].googleItemId; break; }
        }
        var t = q.selectedType || q.type || 'text';
        var opts = (q.choices || []).map(function(c, ci) {
          var label = typeof c === 'string' ? c : (c.label || String(c));
          var opt = { value: label };
          if (ci === autreIdx) opt.goToSectionId = nextSi.googleItemId;
          else if (skipId) opt.goToSectionId = skipId;
          else opt.goToAction = 'NEXT_SECTION';
          return opt;
        });
        navRequests.push({ updateItem: { item: { itemId: si.googleItemId, title: q.label || '', questionItem: { question: { required: q.required !== false, choiceQuestion: { type: t === 'select_multiple' ? 'CHECKBOX' : 'RADIO', options: opts, shuffle: false } } } }, location: { index: si.idx }, updateMask: 'questionItem' } });
      });

      if (navRequests.length > 0) {
        await fetch('https://forms.googleapis.com/v1/forms/' + formId + ':batchUpdate', {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + accessToken, 'Content-Type': 'application/json' },
          body: JSON.stringify({ requests: navRequests })
        });
      }
    }

    console.log('[DEPLOY] Google Forms ok ' + formId);
    res.json({ success: true, formId: formId, url: 'https://docs.google.com/forms/d/' + formId + '/edit', questions: questions.length });
  } catch(err) {
    console.error('[DEPLOY GOOGLE ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ EXPORT EXCEL ============
app.post('/api/deploy/excel', async (req, res) => {
  try {
    const { form } = req.body;
    if (!form) return res.status(400).json({ error: 'Formulaire manquant' });
    const questions = form.questions || [];
    console.log('[EXCEL] ' + questions.length + ' questions');

    const stopWords = new Set(['quel','quelle','quels','quelles','est','sont','avez','vous','votre','vos','etes','avons','ont','avoir','quoi','comment','combien','pourquoi','si','oui','non','les','des','une','un','la','le','au','aux','du','de','que','qui','quand','par','pour','avec','sans','sur','sous','dans','entre','vers','chez','en','et','mais','donc','car','ni','or','pas','ne','plus','moins','tres','bien','tout','tous','toute','toutes','autre','autres','meme','plusieurs','quelques']);

    function makeVarName(label, usedNames) {
      var s = label.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
      var words = s.replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(function(w) { return w.length > 2 && !stopWords.has(w); });
      var base = words.slice(0, 3).join('_').slice(0, 25).replace(/_+$/, '') || 'variable';
      var name = base;
      if (usedNames[base]) { usedNames[base]++; name = base + '_' + usedNames[base]; }
      else usedNames[base] = 1;
      return name;
    }

    var groupMap = {}, groupOrder = [], usedNames = {}, colMeta = [];
    questions.forEach(function(q) {
      var g = q.group || 'General';
      if (!groupMap[g]) { groupMap[g] = []; groupOrder.push(g); }
      groupMap[g].push(q);
      colMeta.push({ q: q, group: g, varname: makeVarName(q.label || q.id || 'var', usedNames) });
    });

    var wb = new ExcelJS.Workbook();
    wb.creator = 'DIGUE';
    var ws1 = wb.addWorksheet('Saisie');
    var ws2 = wb.addWorksheet('Dictionnaire');

    var FILL_SEC = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    var FILL_VAR = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E75B6' } };
    var FILL_ODD = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF3FB' } };
    var FILL_EVN = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
    var FNT_WH = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
    var FNT_VAR = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 9 };
    var FNT_DAT = { name: 'Arial', size: 9 };
    var BRD = { top: { style: 'thin', color: { argb: 'FFBDD7EE' } }, bottom: { style: 'thin', color: { argb: 'FFBDD7EE' } }, left: { style: 'thin', color: { argb: 'FFBDD7EE' } }, right: { style: 'thin', color: { argb: 'FFBDD7EE' } } };

    var r1 = ws1.getRow(1); r1.height = 22;
    var col = 1;
    groupOrder.forEach(function(g) {
      var n = groupMap[g].length;
      var cell = ws1.getCell(1, col);
      cell.value = g; cell.font = FNT_WH; cell.fill = FILL_SEC; cell.border = BRD;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      if (n > 1) ws1.mergeCells(1, col, 1, col + n - 1);
      col += n;
    });

    var r2 = ws1.getRow(2); r2.height = 20;
    colMeta.forEach(function(meta, i) {
      var cell = ws1.getCell(2, i + 1);
      cell.value = meta.varname; cell.font = FNT_VAR; cell.fill = FILL_VAR; cell.border = BRD;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });

    ws1.views = [{ state: 'frozen', ySplit: 2 }];
    var N = 100;
    for (var row = 3; row <= 2 + N; row++) {
      var fill = (row % 2 === 1) ? FILL_ODD : FILL_EVN;
      ws1.getRow(row).height = 18;
      colMeta.forEach(function(meta, i) {
        var cell = ws1.getCell(row, i + 1);
        cell.fill = fill; cell.border = BRD; cell.font = FNT_DAT;
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
      });
    }

    colMeta.forEach(function(meta, i) {
      var q = meta.q;
      var t = q.selectedType || q.type || 'text';
      var colNum = i + 1;
      for (var row = 3; row <= 2 + N; row++) {
        var cell = ws1.getCell(row, colNum);
        if (t === 'select_one' || t === 'select_multiple') {
          var choices = (q.choices || []).map(function(c) { return typeof c === 'string' ? c : (c.label || String(c)); });
          if (choices.length > 0) {
            var formula = '"' + choices.slice(0, 20).join(',') + '"';
            if (formula.length <= 255) {
              cell.dataValidation = { type: 'list', allowBlank: true, formulae: [formula], showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Valeur invalide', error: 'Choisissez une valeur dans la liste' };
            }
          }
        } else if (t === 'integer') {
          var nm = q.numMin, nx = q.numMax;
          var dv = { type: 'whole', allowBlank: true, showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Valeur invalide', error: 'Entrez un nombre entier' };
          if (nm && nx) { dv.operator = 'between'; dv.formulae = [String(nm), String(nx)]; }
          else if (nm) { dv.operator = 'greaterThanOrEqual'; dv.formulae = [String(nm)]; }
          else if (nx) { dv.operator = 'lessThanOrEqual'; dv.formulae = [String(nx)]; }
          else { dv.operator = 'between'; dv.formulae = ['-999999', '999999']; }
          cell.dataValidation = dv; cell.numFmt = '0';
        } else if (t === 'decimal') {
          var nm2 = q.numMin, nx2 = q.numMax;
          var dv2 = { type: 'decimal', allowBlank: true, showErrorMessage: true, errorStyle: 'stop', errorTitle: 'Valeur invalide', error: 'Entrez un nombre decimal' };
          if (nm2 && nx2) { dv2.operator = 'between'; dv2.formulae = [String(nm2), String(nx2)]; }
          else { dv2.operator = 'between'; dv2.formulae = ['-999999', '999999']; }
          cell.dataValidation = dv2;
          var after = q.numDigitsAfter;
          cell.numFmt = '0.' + '0'.repeat(after && String(after).match(/^\d+$/) ? parseInt(after) : 2);
        } else if (t === 'date') {
          cell.dataValidation = { type: 'date', allowBlank: true, showInputMessage: true, promptTitle: 'Date', prompt: 'Format: JJ/MM/AAAA' };
          cell.numFmt = 'DD/MM/YYYY';
        } else if (t === 'time') {
          cell.numFmt = 'HH:MM';
        } else if (t === 'datetime') {
          cell.numFmt = 'DD/MM/YYYY HH:MM';
        }
      }
      ws1.getColumn(colNum).width = Math.max(15, Math.min(28, meta.varname.length + 4));
    });

    var typeLabels = { select_one: 'Choix unique', select_multiple: 'Choix multiple', integer: 'Nombre entier', decimal: 'Nombre decimal', date: 'Date', time: 'Heure', datetime: 'Date et heure', text: 'Texte libre', calculate: 'Calcul auto', geopoint: 'GPS', image: 'Photo', audio: 'Audio', video: 'Video', file: 'Fichier', barcode: 'Code-barres', acknowledge: 'Confirmation', rank: 'Classement', range: 'Echelle', note: 'Note/Texte' };
    var hdrs = ['Variable', 'Libelle complet de la question', 'Type', 'Section', 'Obligatoire', 'Modalites'];
    hdrs.forEach(function(h, j) {
      var cell = ws2.getCell(1, j + 1);
      cell.value = h; cell.font = FNT_WH; cell.fill = FILL_SEC; cell.border = BRD;
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    ws2.getRow(1).height = 20;
    ws2.views = [{ state: 'frozen', ySplit: 1 }];
    ws2.columns = [{ width: 22 }, { width: 55 }, { width: 18 }, { width: 25 }, { width: 12 }, { width: 45 }];

    colMeta.forEach(function(meta, r) {
      var q = meta.q;
      var t = q.selectedType || q.type || 'text';
      var choices = (q.choices || []).map(function(c) { return typeof c === 'string' ? c : (c.label || String(c)); }).join(' / ');
      var fill = (r % 2 === 0) ? FILL_ODD : FILL_EVN;
      var rowData = [meta.varname, q.label || '', typeLabels[t] || t, meta.group, q.required !== false ? 'Oui' : 'Non', choices];
      rowData.forEach(function(val, j) {
        var cell = ws2.getCell(r + 2, j + 1);
        cell.value = val; cell.font = FNT_DAT; cell.fill = fill; cell.border = BRD;
        cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: (j === 1 || j === 5) };
      });
    });

    var filename = (form.title || 'formulaire').replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 40) + '_masque.xlsx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + filename + '"');
    await wb.xlsx.write(res);
    console.log('[EXCEL] ok - ' + filename);
  } catch(err) {
    console.error('[EXCEL ERROR]', err.message);
    if (!res.headersSent) res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ BUILDER KOBOTOOLBOX ============
function buildKoboContent(form) {
  var survey = [];
  var choices = [];
  var seen = new Set();
  var groups = {};
  var groupOrder = [];

  (form.questions || []).forEach(function(q) {
    var g = q.group || 'general';
    if (!groups[g]) { groups[g] = []; groupOrder.push(g); }
    groups[g].push(q);
  });

  // Corriger les sauts invalides avant de construire
  var validQIds = new Set((form.questions || []).map(function(q) {
    return (q.id || ('q' + q.num)).replace(/[^a-zA-Z0-9_]/g, '_');
  }));

  groupOrder.forEach(function(gname) {
    var qs = groups[gname];
    var gId = gname.replace(/\s+/g, '_').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9_]/g, '').slice(0, 32) || 'group';

    if (gname !== 'general') survey.push({ type: 'begin_group', name: gId, label: gname });

    qs.forEach(function(q) {
      var fmtIdx = q.validatedFormatIdx !== undefined ? q.validatedFormatIdx : (q.suggested_format_idx || 0);
      var selectedFmt = q.formats && q.formats[fmtIdx] ? q.formats[fmtIdx] : null;
      var t = q.selectedType || (selectedFmt ? selectedFmt.type : null) || q.type || 'text';
      var name = (q.id || ('q' + q.num)).replace(/[^a-zA-Z0-9_]/g, '_');
      var row = { type: t, name: name, label: q.label || '', required: q.required !== false ? 'yes' : 'no', hint: q.hint || '' };

      // Skip logic — verifier que les IDs references existent et pas de cycle
      var relevant = q.relevant || '';
      if (!relevant && q.suggestions) {
        q.suggestions.forEach(function(s, si) {
          if (s.type === 'skip_logic' && s.value && q.confirmedSuggestions && q.confirmedSuggestions[si]) relevant = s.value;
        });
      }
      if (relevant) {
        var refs = relevant.match(/\$\{([^}]+)\}/g) || [];
        // 1. Verifier que tous les IDs existent
        var allValid = refs.every(function(ref) {
          return validQIds.has(ref.replace('${', '').replace('}', ''));
        });
        if (!allValid) {
          console.log('[KOBO] Saut invalide ignore pour ' + name + ': ' + relevant);
          relevant = '';
        }
        // 2. Verifier pas de cycle: la question ne peut pas referencer une question qui la referencera
        if (relevant) {
          var refIds = refs.map(function(r) { return r.replace('${','').replace('}',''); });
          var hasCycle = refIds.some(function(refId) {
            // Chercher si la question referencee a elle-meme un relevant qui pointe vers la question courante
            var refQ = (form.questions || []).find(function(qq) {
              return (qq.id || ('q'+qq.num)).replace(/[^a-zA-Z0-9_]/g,'_') === refId;
            });
            if (!refQ || !refQ.relevant) return false;
            var backRefs = (refQ.relevant.match(/\$\{([^}]+)\}/g) || []).map(function(r){ return r.replace('${','').replace('}',''); });
            return backRefs.indexOf(name) >= 0;
          });
          if (hasCycle) {
            console.log('[KOBO] Cycle detecte et supprime pour ' + name + ': ' + relevant);
            relevant = '';
          }
        }
        if (relevant) row.relevant = relevant;
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

      if (t === 'range') { row.parameters = 'start=' + (q.numMin || 1) + ' end=' + (q.numMax || 10); delete row.required; }
      if (t === 'note' || t === 'calculate') delete row.required;
      if (t === 'audio' && q.required === false) delete row.required;

      // Choices
      if (t === 'select_one' || t === 'select_multiple' || t === 'rank') {
        var listName = 'list_' + name;
        row.type = t + ' ' + listName;
        if (!seen.has(listName)) {
          seen.add(listName);
          var choiceVals = q.choice_values || [];
          var choiceImages = q.choiceImages || {};
          (q.choices || []).forEach(function(c, i) {
            var label = typeof c === 'string' ? c : (c.label || String(c));
            var val = (choiceVals[i] !== undefined && choiceVals[i] !== '') ? String(choiceVals[i]) : String(i);
            var choiceRow = { list_name: listName, name: val, label: label };
            // Ajouter l'image si disponible
            if (choiceImages[i] && choiceImages[i].url) {
              // KoboToolbox attend le nom du fichier (pas l'URL) dans media::image
              var imgLabel = typeof c === 'string' ? c : (c.label || String(c));
              var imgKey = imgLabel.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'_').replace(/_+/g,'_').slice(0,50);
              choiceRow['media::image'] = imgKey + '.png';
            }
            choices.push(choiceRow);
          });
        }
      }

      // Audio — type audio directement supporté par KoboToolbox
      if (t === 'audio' || t === 'video' || t === 'image' || t === 'file') {
        // Ces types sont natifs dans XLSForm — rien de spécial à faire
        // S'assurer que required est absent pour audio optionnel
        if (q.required === false) delete row.required;
      }

      survey.push(row);
    });

    if (gname !== 'general') survey.push({ type: 'end_group', name: gId });
  });

  var formId = (form.title || 'formulaire').replace(/\s+/g, '_').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9_]/g, '').slice(0, 32);
  return { survey: survey, choices: choices, settings: [{ form_title: form.title || 'Formulaire', form_id: formId, version: '1' }] };
}

// ============ GENERATION D'IMAGES POUR MODALITES ============
// Cache en mémoire des images déjà générées (label normalisé -> url Cloudinary)
var imageCache = {};

function normalizeLabel(label) {
  return label.toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]/g, '_')
    .replace(/_+/g, '_')
    .slice(0, 50);
}

app.post('/api/generate-image', async (req, res) => {
  try {
    const { label, context } = req.body;
    if (!label) return res.status(400).json({ error: 'Label manquant' });

    const key = normalizeLabel(label);
    console.log('[IMAGE] Demande pour: ' + label + ' (key: ' + key + ')');

    // 1. Chercher dans le cache mémoire
    if (imageCache[key]) {
      console.log('[IMAGE] Cache hit: ' + key);
      return res.json({ success: true, url: imageCache[key], fromCache: true });
    }

    // 2. Chercher dans Cloudinary (base persistante)
    try {
      const cloudResult = await cloudinary.api.resource('digue/choices/' + key);
      if (cloudResult && cloudResult.secure_url) {
        imageCache[key] = cloudResult.secure_url;
        console.log('[IMAGE] Cloudinary hit: ' + key);
        return res.json({ success: true, url: cloudResult.secure_url, fromCache: true });
      }
    } catch(cloudErr) {
      // Image pas encore en Cloudinary - on va la générer
    }

    // 3. Générer avec DALL-E via OpenAI
    const stabilityKey = process.env.STABILITY_API_KEY;
    if (!stabilityKey) {
      // Pas de clé Stability — fallback SVG
      try {
        const initials = label.trim().slice(0, 2).toUpperCase();
        const colors = ['E8132A','1E40AF','10B981','F59E0B','7C3AED','0891B2','DC2626','059669'];
        const color = colors[label.charCodeAt(0) % colors.length];
        const svgContent = '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="200" viewBox="0 0 200 200"><rect width="200" height="200" rx="20" fill="#' + color + '"/><text x="100" y="120" font-family="Arial,sans-serif" font-size="72" font-weight="bold" fill="white" text-anchor="middle">' + initials + '</text></svg>';
        const svgBuffer = Buffer.from(svgContent);
        const uploadResult = await new Promise(function(resolve, reject) {
          const stream = cloudinary.uploader.upload_stream(
            { public_id: 'digue/choices/' + key, resource_type: 'image', format: 'png',
              access_mode: 'public', type: 'upload',
              transformation: [{ width: 200, height: 200 }] },
            function(error, result) { if (error) reject(error); else resolve(result); }
          );
          stream.end(svgBuffer);
        });
        imageCache[key] = uploadResult.secure_url;
        return res.json({ success: true, url: uploadResult.secure_url, fromCache: false, fallback: true });
      } catch(svgErr) {
        const fallbackUrl = 'https://ui-avatars.com/api/?name=' + encodeURIComponent(label) + '&size=200&background=E8132A&color=fff&bold=true';
        imageCache[key] = fallbackUrl;
        return res.json({ success: true, url: fallbackUrl, fromCache: false, fallback: true });
      }
    }

    // Générer avec Stability AI (Stable Diffusion)
    const prompt = 'Simple, clear, flat illustration representing "' + label + '" for a data collection survey. ' +
      (context ? 'Survey context: ' + context + '. ' : '') +
      'Minimalist icon style, white background, single clear symbol, bright colors, no text, no letters.';

    console.log('[IMAGE] Génération Stability AI pour: ' + label + ' (key: ' + key + ')');

    const formData = new FormData();
    formData.append('prompt', prompt);
    formData.append('output_format', 'png');
    formData.append('aspect_ratio', '1:1');
    formData.append('style_preset', 'digital-art');

    const imgRes = await fetch('https://api.stability.ai/v2beta/stable-image/generate/core', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + stabilityKey,
        'Accept': 'image/*',
        ...formData.getHeaders()
      },
      body: formData
    });

    if (!imgRes.ok) {
      const errText = await imgRes.text();
      console.error('[IMAGE] Stability AI error:', errText.slice(0, 200));
      const fallbackUrl = 'https://ui-avatars.com/api/?name=' + encodeURIComponent(label) + '&size=200&background=1F4E79&color=fff&bold=true';
      imageCache[key] = fallbackUrl;
      return res.json({ success: true, url: fallbackUrl, fromCache: false, fallback: true });
    }

    const imageBuffer = Buffer.from(await imgRes.arrayBuffer());
    const imageUrl = null; // On upload directement le buffer

    // 4. Uploader le buffer vers Cloudinary pour persistance
    const uploadResult = await new Promise(function(resolve, reject) {
      const stream = cloudinary.uploader.upload_stream(
        { public_id: 'digue/choices/' + key, overwrite: true, resource_type: 'image',
          access_mode: 'public', type: 'upload',
          transformation: [{ width: 400, height: 400, crop: 'fill', quality: 'auto' }] },
        function(error, result) { if (error) reject(error); else resolve(result); }
      );
      stream.end(imageBuffer);
    });

    const finalUrl = uploadResult.secure_url;
    imageCache[key] = finalUrl;
    console.log('[IMAGE] Générée et stockée: ' + key + ' -> ' + finalUrl);
    res.json({ success: true, url: finalUrl, fromCache: false });

  } catch(err) {
    console.error('[IMAGE ERROR]', err.message);
    const fallbackUrl = 'https://ui-avatars.com/api/?name=' + encodeURIComponent(req.body.label || 'img') + '&size=200&background=10B981&color=fff&bold=true';
    res.json({ success: true, url: fallbackUrl, fromCache: false, fallback: true });
  }
});

// Upload image utilisateur vers Cloudinary
app.post('/api/upload-image', upload.single('image'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Aucune image reçue' });
    const key = req.body.key || ('upload_' + Date.now());

    const result = await new Promise(function(resolve, reject) {
      const stream = cloudinary.uploader.upload_stream(
        { public_id: 'digue/choices/' + key, resource_type: 'image', transformation: [{ width: 400, height: 400, crop: 'fill', quality: 'auto' }] },
        function(error, result) { if (error) reject(error); else resolve(result); }
      );
      stream.end(req.file.buffer);
    });

    console.log('[UPLOAD] Image uploadée: ' + result.secure_url);
    res.json({ success: true, url: result.secure_url });
  } catch(err) {
    console.error('[UPLOAD ERROR]', err.message);
    res.status(500).json({ error: 'Erreur upload: ' + err.message });
  }
});

// ============ PAIEMENT FEDAPAY ============
app.post('/api/payment/initiate', async (req, res) => {
  try {
    const { amount, description, customer } = req.body;
    if (!amount || !description) return res.status(400).json({ error: 'Donnees manquantes' });

    const FEDAPAY_SECRET = process.env.FEDAPAY_SECRET_KEY || 'sk_sandbox_MbYxK8aQRbGAkAVjN-p2nTrh';
    const FEDAPAY_BASE = FEDAPAY_SECRET.includes('sandbox') ? 'https://sandbox-api.fedapay.com' : 'https://api.fedapay.com';

    // Creer une transaction FedaPay
    const txRes = await fetch(FEDAPAY_BASE + '/v1/transactions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + FEDAPAY_SECRET
      },
      body: JSON.stringify({
        description: description,
        amount: amount,
        currency: { iso: 'XOF' },
        callback_url: 'https://dansouborispleck-sketch.github.io/r2-forms/',
        customer: customer || { email: 'client@r2forms.bj' }
      })
    });

    if (!txRes.ok) {
      const e = await txRes.json();
      console.error('[FEDAPAY] Error:', JSON.stringify(e));
      return res.status(502).json({ error: 'PAYMENT_ERROR', message: 'Erreur initialisation paiement.' });
    }

    const txData = await txRes.json();
    console.log('[FEDAPAY] Response:', JSON.stringify(txData).slice(0, 300));

    // FedaPay retourne la transaction dans différents formats selon la version
    const transactionId = (txData.id) ||
                          (txData['v1/transaction'] && txData['v1/transaction'].id) ||
                          (txData.transaction && txData.transaction.id);

    if (!transactionId) {
      console.error('[FEDAPAY] ID non trouvé dans:', JSON.stringify(txData).slice(0, 200));
      return res.status(502).json({ error: 'PAYMENT_ERROR', message: 'ID transaction non recu.' });
    }

    // Generer le token de paiement
    const tokenRes = await fetch(FEDAPAY_BASE + '/v1/transactions/' + transactionId + '/token', {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + FEDAPAY_SECRET, 'Content-Type': 'application/json' }
    });

    if (!tokenRes.ok) {
      const tokenErr = await tokenRes.json();
      console.error('[FEDAPAY TOKEN ERROR]', JSON.stringify(tokenErr));
      return res.status(502).json({ error: 'TOKEN_ERROR', message: 'Erreur generation token.' });
    }

    const tokenData = await tokenRes.json();
    console.log('[FEDAPAY] Token response:', JSON.stringify(tokenData).slice(0, 300));
    // FedaPay retourne le token et l'URL directe
    const token = tokenData.token ||
                  (tokenData['v1/transaction'] && tokenData['v1/transaction'].token) ||
                  tokenData.url;
    const checkoutUrl = tokenData.url ||
                        (tokenData['v1/transaction'] && tokenData['v1/transaction'].url) ||
                        'https://sandbox-checkout.fedapay.com/' + token;

    console.log('[FEDAPAY] Transaction creee: ' + transactionId + ' - ' + amount + ' XOF - URL: ' + checkoutUrl);
    res.json({ success: true, transactionId: transactionId, token: token, checkoutUrl: checkoutUrl, amount: amount });

  } catch(err) {
    console.error('[PAYMENT ERROR]', err.message);
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

app.post('/api/payment/verify', async (req, res) => {
  try {
    const { transactionId } = req.body;
    if (!transactionId) return res.status(400).json({ error: 'Transaction ID manquant' });

    const FEDAPAY_SECRET = process.env.FEDAPAY_SECRET_KEY || 'sk_sandbox_MbYxK8aQRbGAkAVjN-p2nTrh';
    const FEDAPAY_BASE = FEDAPAY_SECRET.includes('sandbox') ? 'https://sandbox-api.fedapay.com' : 'https://api.fedapay.com';

    const verifyRes = await fetch(FEDAPAY_BASE + '/v1/transactions/' + transactionId, {
      headers: { 'Authorization': 'Bearer ' + FEDAPAY_SECRET }
    });

    if (!verifyRes.ok) return res.status(502).json({ error: 'VERIFY_ERROR' });

    const data = await verifyRes.json();
    console.log('[FEDAPAY VERIFY]', JSON.stringify(data).slice(0, 200));
    const tx = data['v1/transaction'] || data.transaction || data;
    const status = tx && tx.status;

    console.log('[FEDAPAY] Verification ' + transactionId + ': ' + status);
    res.json({ success: true, status: status, approved: status === 'approved' });

  } catch(err) {
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ TELECHARGEMENT ZIP DES IMAGES ============
app.post('/api/download-images-zip', async (req, res) => {
  try {
    const { images, title } = req.body;
    if (!images || images.length === 0) return res.status(400).json({ error: 'Aucune image' });

    const archiver = require('archiver');
    const archive = archiver('zip', { zlib: { level: 9 } });

    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="' + (title || 'images').replace(/[^a-zA-Z0-9_-]/g,'_') + '_images.zip"');
    archive.pipe(res);

    for (var i = 0; i < images.length; i++) {
      var img = images[i];
      try {
        var imgRes = await fetch(img.url);
        if (imgRes.ok) {
          var imgBuffer = Buffer.from(await imgRes.arrayBuffer());
          archive.append(imgBuffer, { name: img.filename });
          console.log('[ZIP] Ajouté: ' + img.filename);
        }
      } catch(e) {
        console.error('[ZIP] Erreur image:', img.filename, e.message);
      }
    }

    await archive.finalize();
    console.log('[ZIP] Archive générée avec ' + images.length + ' images');
  } catch(err) {
    console.error('[ZIP ERROR]', err.message);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// Route pour vider le cache images (admin)
app.post('/api/clear-image-cache', async (req, res) => {
  try {
    imageCache = {};
    // Supprimer toutes les images dans Cloudinary digue/choices/
    const result = await cloudinary.api.delete_resources_by_prefix('digue/choices/');
    console.log('[CACHE] Images supprimées:', result);
    res.json({ success: true, message: 'Cache vidé et images Cloudinary supprimées' });
  } catch(err) {
    res.json({ success: true, message: 'Cache mémoire vidé (erreur Cloudinary: ' + err.message + ')' });
  }
});

app.listen(PORT, function() {
  console.log('\n╔══════════════════════════════════╗\n║   DIGUE Backend v5.0          ║\n║   Port: ' + PORT + '                    ║\n╚══════════════════════════════════╝\n');
});
