/**
 * R2 Forms — Backend Server
 * Gère : import de fichiers, analyse Claude, déploiement KoboToolbox
 */

require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const pdf = require('pdf-parse');
const fetch = require('node-fetch');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 4000;

// ============ MIDDLEWARE ============
app.use(cors({ origin: '*' }));
app.use(express.json({ limit: '20mb' }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 }
});

app.use(express.static(path.join(__dirname, 'public')));

// ============ ROUTE: SANTÉ ============
app.get('/health', (req, res) => {
  res.json({ status: 'ok', version: '1.0.0', timestamp: new Date().toISOString() });
});

// ============ ROUTE: IMPORT FICHIER ============
app.post('/api/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'Aucun fichier reçu' });
    }

    const { originalname, buffer, mimetype } = req.file;
    const ext = originalname.split('.').pop().toLowerCase();
    let extractedText = '';
    let metadata = { filename: originalname, size: buffer.length, type: ext };

    console.log(`[IMPORT] ${originalname} (${(buffer.length/1024).toFixed(0)} KB)`);

    if (['txt', 'csv'].includes(ext)) {
      extractedText = buffer.toString('utf-8');
    }
    else if (ext === 'pdf') {
      try {
        const pdfData = await pdf(buffer);
        extractedText = pdfData.text;
        metadata.pages = pdfData.numpages;
        if (!extractedText || extractedText.trim().length < 20) {
          return res.status(422).json({
            error: 'PDF_SCANNED',
            message: 'Ce PDF semble être une image scannée. Veuillez coller le texte directement.',
            metadata
          });
        }
      } catch (pdfErr) {
        return res.status(422).json({ error: 'PDF_ERROR', message: 'Impossible de lire ce PDF.', metadata });
      }
    }
    else if (ext === 'docx') {
      try {
        const result = await mammoth.extractRawText({ buffer });
        extractedText = result.value;
      } catch (wordErr) {
        return res.status(422).json({ error: 'DOCX_ERROR', message: 'Impossible de lire ce fichier Word.' });
      }
    }
    else if (ext === 'doc') {
      extractedText = buffer.toString('latin1')
        .replace(/[^\x20-\x7E\n\r\u00C0-\u024F]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      if (extractedText.length < 30) {
        return res.status(422).json({
          error: 'DOC_OLD_FORMAT',
          message: 'Format .doc ancien non supporté. Enregistrez en .docx ou collez le texte.'
        });
      }
    }
    else if (['xlsx', 'xls'].includes(ext)) {
      try {
        const workbook = XLSX.read(buffer, { type: 'buffer' });
        let allText = '';
        workbook.SheetNames.forEach(sheetName => {
          const sheet = workbook.Sheets[sheetName];
          const csvData = XLSX.utils.sheet_to_csv(sheet);
          allText += `\n\n=== Feuille: ${sheetName} ===\n${csvData}`;
        });
        extractedText = allText.trim();
      } catch (xlsxErr) {
        return res.status(422).json({ error: 'XLSX_ERROR', message: 'Impossible de lire ce fichier Excel.' });
      }
    }
    else if (ext === 'odt') {
      try {
        const text = buffer.toString('utf-8');
        const textMatches = text.match(/<text:p[^>]*>([^<]{2,})<\/text:p>/g) || [];
        extractedText = textMatches
          .map(m => m.replace(/<[^>]+>/g, '').trim())
          .filter(t => t.length > 1)
          .join('\n');
        if (!extractedText) {
          extractedText = text.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 8000);
        }
      } catch (odtErr) {
        return res.status(422).json({ error: 'ODT_ERROR', message: 'Impossible de lire ce fichier ODT.' });
      }
    }
    else {
      return res.status(400).json({ error: 'FORMAT_UNSUPPORTED', message: `Format .${ext} non supporté.` });
    }

    extractedText = extractedText
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();

    if (extractedText.length < 20) {
      return res.status(422).json({
        error: 'EMPTY_CONTENT',
        message: 'Le fichier semble vide ou illisible. Collez le texte directement.'
      });
    }

    console.log(`[IMPORT] ✓ ${extractedText.length} caractères extraits`);

    res.json({
      success: true,
      text: extractedText,
      preview: extractedText.slice(0, 300) + (extractedText.length > 300 ? '...' : ''),
      metadata: { ...metadata, chars: extractedText.length }
    });

  } catch (err) {
    console.error('[IMPORT ERROR]', err);
    res.status(500).json({ error: 'SERVER_ERROR', message: 'Erreur serveur lors de l\'import.' });
  }
});

// ============ ROUTE: ANALYSE CLAUDE ============
// Chunking : découpe le texte en morceaux et fusionne les résultats

const CHUNK_SIZE = 5000;    // caractères par chunk
const CHUNK_OVERLAP = 200;  // overlap pour ne pas couper une question en deux

async function analyseChunk(apiKey, chunkText, tool, chunkNum, totalChunks, previousIds) {
  const isFirst = chunkNum === 1;
  const isLast = chunkNum === totalChunks;
  
  const contextNote = totalChunks > 1 
    ? `\nCeci est la partie ${chunkNum} sur ${totalChunks} du questionnaire.${isFirst ? ' Commence depuis le début.' : ' Continue l\'extraction — ne répète pas les questions déjà extraites.'}${isLast ? ' C\'est la dernière partie.' : ''}`
    : '';

  const previousNote = previousIds.length > 0
    ? `\nQuestions déjà extraites (IDs à ne pas répéter): ${previousIds.join(', ')}`
    : '';

  const system = `Tu es un expert en collecte de données terrain pour les ONG, universités et entreprises en Afrique de l'Ouest.
Tu analyses des questionnaires et extrais leur structure COMPLÈTE avec toute la logique interne.

Réponds UNIQUEMENT en JSON valide compact (sans indentation, sans espaces inutiles), sans markdown ni texte autour.

=== CLASSIFICATION DES QUESTIONS ===
Pour chaque question, détermine sa classe selon sa nature réelle :
- quantitative : variable numérique mesurable (âge, poids, taille, revenu, score, durée, quantité, nombre de...). JAMAIS de choix pour ces questions.
- qualitative_choice : réponse choisie parmi des options définies (sexe, niveau d'instruction, statut, oui/non, catégories...)
- qualitative_open : réponse libre sans options (nom, commentaire, description, adresse, opinion...)
- date_time : question de date, heure ou date+heure
- geopoint : localisation GPS précise (un point)
- geotrace : trajet ou chemin GPS
- geoshape : zone ou périmètre GPS
- media_photo : question demandant une photo
- media_audio : question demandant un enregistrement audio
- media_video : question demandant une vidéo
- media_file : question demandant un fichier joint
- barcode : scan de code-barres ou QR code
- acknowledge : question d'accord ou de confirmation
- ranking : classement par ordre de préférence
- scale : échelle ou curseur numérique
- calculate : variable calculée automatiquement
- note : information sans réponse

=== FORMAT DE SORTIE (JSON COMPACT OBLIGATOIRE) ===
{"title":"titre","summary":"résumé","questions":[{"id":"q1","num":1,"label":"libellé","question_class":"quantitative","type":"integer","required":true,"hint":"","choices":[],"group":"groupe ou null","formats":[],"suggested_format_idx":0,"suggestions":[]}],"groups":[]}

=== FORMATS SELON LA CLASSE ===
quantitative: formats=[{"id":"A","name":"Nombre entier","type":"integer","note":"Ex: 25, 150 — sans virgule"},{"id":"B","name":"Nombre avec virgule","type":"decimal","note":"Ex: 65,5 kg, 37,8°C"},{"id":"C","name":"Valeur sur une echelle","type":"range","note":"Ex: satisfaction de 1 a 10"}]
qualitative_choice: formats=[{"id":"A","name":"Une seule reponse au choix","type":"select_one","note":"Le repondant ne coche qu'une case"},{"id":"B","name":"Plusieurs reponses possibles","type":"select_multiple","note":"Le repondant peut cocher plusieurs cases"},{"id":"C","name":"Reponse ecrite libre","type":"text","note":"Le repondant ecrit lui-meme sa reponse"}]
qualitative_open: formats=[{"id":"A","name":"Reponse courte","type":"text","note":"Quelques mots"},{"id":"B","name":"Reponse longue","type":"text","note":"Plusieurs phrases"}]
date_time: formats=[{"id":"A","name":"Date (jour/mois/annee)","type":"date","note":"Ex: 15/03/2024"},{"id":"B","name":"Heure","type":"time","note":"Ex: 14h30"},{"id":"C","name":"Date et heure","type":"datetime","note":"Ex: 15/03/2024 a 14h30"}]
geopoint: formats=[{"id":"A","name":"Localisation GPS (un point precis)","type":"geopoint","note":"Capture automatiquement la position GPS"}]
geotrace: formats=[{"id":"A","name":"Tracer un chemin GPS","type":"geotrace","note":"Enregistre un trajet sur la carte"}]
geoshape: formats=[{"id":"A","name":"Delimiter une zone GPS","type":"geoshape","note":"Dessine un perimetre sur la carte"}]
media_photo: formats=[{"id":"A","name":"Prendre une photo","type":"image","note":"Photo avec l'appareil photo"}]
media_audio: formats=[{"id":"A","name":"Enregistrer un son","type":"audio","note":"Enregistrement audio"}]
media_video: formats=[{"id":"A","name":"Enregistrer une video","type":"video","note":"Enregistrement video"}]
media_file: formats=[{"id":"A","name":"Joindre un fichier","type":"file","note":"Document PDF, Excel ou autre"}]
barcode: formats=[{"id":"A","name":"Scanner un code-barres","type":"barcode","note":"Scan QR code ou code-barres"}]
acknowledge: formats=[{"id":"A","name":"Case a cocher pour confirmer","type":"acknowledge","note":"Une seule case a cocher"}]
ranking: formats=[{"id":"A","name":"Classer par ordre de preference","type":"rank","note":"Du plus important au moins important"}]
scale: formats=[{"id":"A","name":"Valeur sur une echelle","type":"range","note":"Ex: satisfaction de 1 a 5"}]
calculate: formats=[{"id":"A","name":"Valeur calculee automatiquement","type":"calculate","note":"Calculee a partir des autres reponses"}]
note: formats=[{"id":"A","name":"Message d'information","type":"note","note":"Affiche un texte sans demander de reponse"}]

=== SUGGESTIONS ===
Pour chaque question, ajoute les suggestions pertinentes detectees dans la logique du questionnaire:
Format: {"type":"skip_logic|calculate|constraint","label":"libelle court","description":"explication claire","value":"formule XLSForm","confidence":"high|medium|low"}
- skip_logic: condition d'affichage. value ex: "\${id_q} = 'oui'" ou "selected(\${id_q}, 'valeur')"
- calculate: calcul auto. value ex: "\${q_poids} div (\${q_taille} * \${q_taille}) * 10000"
- constraint: contrainte logique. value ex: ". >= 0 and . <= 150"

=== REGLES ===
- Toutes questions required:true par defaut sauf si explicitement facultatif
- Respecte l'ordre original des questions
- Regroupe par thematique dans "group"
- Extrais TOUTES les questions et sous-questions sans exception
- Pour qualitative_choice: extrais les vraies modalites dans choices[]
- Numerote les questions en continu (num) meme sur plusieurs parties
- Outil cible: ${tool}${contextNote}${previousNote}`;

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-5',
      max_tokens: 8000,
      system,
      messages: [{ role: 'user', content: `Extrais toutes les questions de cette partie du questionnaire:\n\n${chunkText}` }]
    })
  });

  if (!response.ok) {
    const errData = await response.json();
    throw new Error(`Claude API error: ${JSON.stringify(errData)}`);
  }

  const data = await response.json();
  const rawText = data.content?.[0]?.text || '{}';

  // Parse JSON
  let cleanText = rawText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
  const jsonMatch = cleanText.match(/\{[\s\S]*/);
  let jsonStr = jsonMatch ? jsonMatch[0] : cleanText;

  try {
    return JSON.parse(jsonStr);
  } catch(e) {
    // Réparer JSON tronqué
    const lastComplete = jsonStr.lastIndexOf('},');
    if (lastComplete > 0) jsonStr = jsonStr.slice(0, lastComplete + 1);
    jsonStr = jsonStr.replace(/,\s*$/, '');
    const opens = (jsonStr.match(/\[/g)||[]).length - (jsonStr.match(/\]/g)||[]).length;
    const openb = (jsonStr.match(/\{/g)||[]).length - (jsonStr.match(/\}/g)||[]).length;
    for(let i = 0; i < Math.max(0,opens); i++) jsonStr += ']';
    for(let i = 0; i < Math.max(0,openb); i++) jsonStr += '}';
    return JSON.parse(jsonStr);
  }
}

function splitIntoChunks(text, chunkSize, overlap) {
  if (text.length <= chunkSize) return [text];
  
  const chunks = [];
  let pos = 0;
  
  while (pos < text.length) {
    let end = Math.min(pos + chunkSize, text.length);
    
    // Ne pas couper au milieu d'une ligne — reculer jusqu'au dernier saut de ligne
    if (end < text.length) {
      const lastNewline = text.lastIndexOf('\n', end);
      if (lastNewline > pos + chunkSize * 0.7) {
        end = lastNewline;
      }
    }
    
    chunks.push(text.slice(pos, end));
    pos = end - overlap;
    if (pos >= text.length) break;
  }
  
  return chunks;
}

app.post('/api/analyse', async (req, res) => {
  try {
    const { text, tool } = req.body;
    if (!text || text.trim().length < 10) {
      return res.status(400).json({ error: 'Texte manquant ou trop court' });
    }

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) {
      return res.status(500).json({ error: 'Clé API non configurée' });
    }

    console.log(`[ANALYSE] ${text.length} chars → ${tool}`);

    // Découper en chunks si nécessaire
    const chunks = splitIntoChunks(text, CHUNK_SIZE, CHUNK_OVERLAP);
    console.log(`[ANALYSE] ${chunks.length} chunk(s) à traiter`);

    let allQuestions = [];
    let formTitle = '';
    let formSummary = '';
    let allGroups = new Set();
    let previousIds = [];

    // Traiter chaque chunk séquentiellement avec retry
    const MAX_RETRIES = 3;
    const failedChunks = [];

    for (let i = 0; i < chunks.length; i++) {
      const chunkNum = i + 1;
      console.log(`[ANALYSE] Chunk ${chunkNum}/${chunks.length} (${chunks[i].length} chars)`);
      
      let success = false;
      let lastError = null;

      // Retry jusqu'à MAX_RETRIES fois
      for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
        try {
          if (attempt > 1) {
            console.log(`[ANALYSE] Chunk ${chunkNum} — Tentative ${attempt}/${MAX_RETRIES}`);
            await new Promise(r => setTimeout(r, 1000 * attempt)); // backoff
          }

          const result = await analyseChunk(apiKey, chunks[i], tool, chunkNum, chunks.length, previousIds);
          
          if (chunkNum === 1) {
            formTitle = result.title || 'Formulaire';
            formSummary = result.summary || '';
          }
          
          if (result.questions && Array.isArray(result.questions)) {
            const existingLabels = new Set(allQuestions.map(q => q.label?.trim().toLowerCase()));
            const newQuestions = result.questions.filter(q => {
              const label = q.label?.trim().toLowerCase();
              return label && !existingLabels.has(label);
            });
            
            newQuestions.forEach((q, idx) => {
              q.num = allQuestions.length + idx + 1;
              q.id = `q${q.num}`;
              existingLabels.add(q.label?.trim().toLowerCase());
            });
            
            allQuestions = allQuestions.concat(newQuestions);
            previousIds = allQuestions.map(q => q.id);
          }
          
          if (result.groups && Array.isArray(result.groups)) {
            result.groups.forEach(g => allGroups.add(g));
          }
          
          console.log(`[ANALYSE] Chunk ${chunkNum}: ${result.questions?.length || 0} questions → Total: ${allQuestions.length}`);
          success = true;
          break; // Succès — on sort du retry

        } catch(chunkErr) {
          lastError = chunkErr.message;
          console.error(`[ANALYSE] Chunk ${chunkNum} tentative ${attempt} échouée:`, chunkErr.message);
        }
      }

      if (!success) {
        // Chunk définitivement échoué après MAX_RETRIES tentatives
        failedChunks.push(chunkNum);
        console.error(`[ANALYSE] ❌ Chunk ${chunkNum} abandonné après ${MAX_RETRIES} tentatives`);
      }

      // Pause entre chunks pour éviter le rate limiting
      if (i < chunks.length - 1) {
        await new Promise(r => setTimeout(r, 500));
      }
    }

    if (allQuestions.length === 0) {
      return res.status(422).json({ 
        error: 'NO_QUESTIONS', 
        message: 'Aucune question détectée dans le document.' 
      });
    }

    // Avertir si des chunks ont échoué
    const hasFailures = failedChunks.length > 0;
    if (hasFailures) {
      console.warn(`[ANALYSE] ⚠️ ${failedChunks.length} chunk(s) échoué(s): ${failedChunks.join(', ')}`);
    }

    const form = {
      title: formTitle,
      summary: formSummary,
      questions: allQuestions,
      groups: [...allGroups]
    };

    console.log(`[ANALYSE] ✓ ${form.questions.length} questions extraites au total`);
    
    // Retourner avec avertissement si des sections ont échoué
    const responseData = { 
      success: true, 
      form,
      chunks_total: chunks.length,
      chunks_failed: failedChunks.length,
      warning: hasFailures 
        ? `Attention : ${failedChunks.length} section(s) du questionnaire n'ont pas pu être traitées (sections ${failedChunks.join(', ')}). Le masque peut être incomplet. Nous vous recommandons de vérifier toutes les questions avant de déployer.`
        : null
    };
    
    res.json(responseData);

  } catch (err) {
    console.error('[ANALYSE ERROR]', err);
    res.status(500).json({ error: 'SERVER_ERROR', message: 'Erreur serveur lors de l\'analyse.' });
  }
});

// ============ ROUTE: CORRECTION CLAUDE ============
app.post('/api/correct', async (req, res) => {
  try {
    const { form, instructions } = req.body;
    if (!form || !instructions) {
      return res.status(400).json({ error: 'Données manquantes' });
    }

    const apiKey = process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(500).json({ error: 'Clé API non configurée' });

    console.log(`[CORRECT] Instructions: ${instructions.slice(0, 100)}`);

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-sonnet-4-5',
        max_tokens: 16000,
        system: `Tu es un expert en collecte de données. L'utilisateur te donne un formulaire JSON et des instructions de correction en langage naturel.
Applique EXACTEMENT les corrections demandées et retourne le formulaire JSON corrigé UNIQUEMENT, sans texte autour, sans markdown, sans balises \`\`\`json.`,
        messages: [{
          role: 'user',
          content: `Formulaire actuel:\n${JSON.stringify(form, null, 2)}\n\nInstructions de correction:\n${instructions}`
        }]
      })
    });

    if (!response.ok) return res.status(502).json({ error: 'CLAUDE_ERROR' });

    const data = await response.json();
    const rawText = data.content?.[0]?.text || '{}';
    let cleanText = rawText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
    const match = cleanText.match(/\{[\s\S]*\}/);
    const corrected = JSON.parse(match ? match[0] : '{}');

    console.log(`[CORRECT] ✓ ${corrected.questions?.length} questions après correction`);
    res.json({ success: true, form: corrected });

  } catch (err) {
    console.error('[CORRECT ERROR]', err);
    res.status(500).json({ error: 'SERVER_ERROR' });
  }
});

// ============ ROUTE: DÉPLOIEMENT KOBOTOOLBOX ============
app.post('/api/deploy/kobo', async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { username, password, server = 'https://kf.kobotoolbox.org' } = credentials;

    if (!form || !username || !password) {
      return res.status(400).json({ error: 'Données manquantes' });
    }

    console.log(`[DEPLOY] KoboToolbox → ${server}`);

    const tokenRes = await fetch(`${server}/token/?format=json`, {
      method: 'GET',
      headers: { 'Authorization': 'Basic ' + Buffer.from(`${username}:${password}`).toString('base64') }
    });

    if (!tokenRes.ok) {
      return res.status(401).json({ error: 'AUTH_ERROR', message: 'Identifiants incorrects. Vérifiez votre email et mot de passe KoboToolbox.' });
    }

    const tokenData = await tokenRes.json();
    const token = tokenData.token;
    const auth = { 'Authorization': 'Token ' + token, 'Content-Type': 'application/json' };

    const assetRes = await fetch(`${server}/api/v2/assets/?format=json`, {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({ name: form.title || 'Formulaire R2', asset_type: 'survey' })
    });

    if (!assetRes.ok) {
      return res.status(502).json({ error: 'ASSET_ERROR', message: 'Erreur lors de la création du formulaire.' });
    }

    const asset = await assetRes.json();
    const assetUid = asset.uid;
    const koboContent = buildKoboContent(form);

    const patchRes = await fetch(`${server}/api/v2/assets/${assetUid}/?format=json`, {
      method: 'PATCH',
      headers: auth,
      body: JSON.stringify({
        name: form.title || 'Formulaire R2',
        content: koboContent
      })
    });

    if (!patchRes.ok) {
      const errBody = await patchRes.text();
      console.error('[PATCH ERROR]', errBody);
      return res.status(502).json({ error: 'PATCH_ERROR', message: 'Erreur lors de l\'import du questionnaire.' });
    }

    const deployRes = await fetch(`${server}/api/v2/assets/${assetUid}/deployment/?format=json`, {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({ active: true })
    });

    const deployOk = deployRes.ok || deployRes.status === 201;
    console.log(`[DEPLOY] ✓ Asset ${assetUid} déployé`);

    res.json({
      success: true,
      uid: assetUid,
      url: `${server}/#/forms/${assetUid}/summary`,
      deployed: deployOk,
      questions: form.questions?.length || 0
    });

  } catch (err) {
    console.error('[DEPLOY ERROR]', err);
    res.status(500).json({ error: 'SERVER_ERROR', message: 'Erreur serveur lors du déploiement.' });
  }
});

// ============ BUILDER KOBOTOOLBOX ============
function buildKoboContent(form) {
  const survey = [];
  const choices = [];
  const choiceListsSeen = new Set();

  const groups = {};
  form.questions.forEach(q => {
    const g = q.group || 'general';
    if (!groups[g]) groups[g] = [];
    groups[g].push(q);
  });

  Object.entries(groups).forEach(([groupName, qs]) => {
    if (groupName !== 'general') {
      survey.push({ type: 'begin_group', name: groupName.replace(/\s+/g, '_').toLowerCase(), label: groupName });
    }

    qs.forEach(q => {
      const t = q.selectedType || q.type || 'text';
      const name = (q.id || ('q' + q.num)).replace(/[^a-zA-Z0-9_]/g, '_');

      const row = {
        type: t,
        name: name,
        label: q.label || '',
        required: q.required !== false ? 'yes' : 'no',
        hint: q.hint || ''
      };

      // Logique de saut (relevant)
      if (q.relevant && q.relevant.trim()) {
        row.relevant = q.relevant.trim();
      }

      // Calcul automatique
      if (t === 'calculate' && q.calculation) {
        row.calculation = q.calculation;
        row.type = 'calculate';
        delete row.required;
      }

      // Contrainte de valeur (min/max + chiffres avant/après virgule)
      if (q.constraint && q.constraint.trim()) {
        row.constraint = q.constraint.trim();
        row.constraint_message = lang === 'fr' ? 'Valeur hors des limites acceptées' : 'Value out of accepted range';
      }

      // Range — paramètres min/max
      if (t === 'range' && (q.numMin || q.numMax)) {
        const rangeMin = q.numMin || '1';
        const rangeMax = q.numMax || '10';
        row.parameters = `start=${rangeMin} end=${rangeMax}`;
      }

      // Questions à choix (select_one, select_multiple, rank)
      if (t === 'select_one' || t === 'select_multiple' || t === 'rank') {
        const listName = 'list_' + name;
        row.type = t + ' ' + listName;
        if (!choiceListsSeen.has(listName)) {
          choiceListsSeen.add(listName);
          (q.choices || []).forEach((c, i) => {
            const label = typeof c === 'string' ? c : (c.label || String(c));
            const val = label.toLowerCase()
              .normalize('NFD').replace(/[̀-ͯ]/g, '')
              .replace(/[^a-z0-9_]/g, '_')
              .replace(/__+/g,'_')
              .replace(/^_|_$/g,'')
              .slice(0,30) || ('c' + (i+1));
            choices.push({ list_name: listName, name: val, label: label });
          });
        }
      }

      // Types sans required (note, calculate)
      if (t === 'note' || t === 'calculate') {
        delete row.required;
      }

      survey.push(row);
    });

    if (groupName !== 'general') {
      survey.push({ type: 'end_group', name: groupName.replace(/\s+/g, '_').toLowerCase() });
    }
  });

  return {
    survey,
    choices,
    settings: [{
      form_title: form.title || 'Formulaire',
      form_id: (form.title || 'formulaire').replace(/\s+/g, '_').toLowerCase().replace(/[^a-z0-9_]/g, ''),
      version: '1'
    }]
  };
}

// ============ DÉMARRAGE ============
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
