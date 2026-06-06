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

    const system = `Tu es un expert en collecte de données terrain pour les ONG, universités et entreprises en Afrique de l'Ouest.
Tu analyses des questionnaires et extrais leur structure COMPLÈTE avec toute la logique interne.

Réponds UNIQUEMENT en JSON valide, sans markdown ni texte autour.

=== CLASSIFICATION DES QUESTIONS ===
Pour chaque question, détermine sa classe selon sa nature réelle :

- quantitative : variable numérique mesurable (âge, poids, taille, revenu, score, durée, distance, température, quantité, nombre de...). JAMAIS de choix pour ces questions.
- qualitative_choice : réponse choisie parmi des options définies (sexe, niveau d'instruction, statut, oui/non, catégories...)
- qualitative_open : réponse libre sans options (nom, commentaire, description, adresse, opinion...)
- date_time : question de date, heure ou date+heure
- geopoint : localisation GPS précise (un point)
- geotrace : trajet ou chemin GPS (plusieurs points formant une ligne)
- geoshape : zone ou périmètre GPS (surface fermée)
- media_photo : question demandant une photo
- media_audio : question demandant un enregistrement audio
- media_video : question demandant une vidéo
- media_file : question demandant un fichier joint
- barcode : scan de code-barres ou QR code
- acknowledge : question d'accord ou de confirmation (case à cocher simple)
- ranking : question de classement par ordre de préférence
- scale : échelle ou curseur numérique (satisfaction de 1 à 5, score de 0 à 10...)
- calculate : variable calculée automatiquement à partir d'autres questions
- note : information ou instruction sans réponse

=== FORMAT DE SORTIE ===
{
  "title": "titre du formulaire",
  "summary": "résumé en 1 phrase",
  "questions": [
    {
      "id": "q1",
      "num": 1,
      "label": "libellé complet de la question",
      "question_class": "quantitative|qualitative_choice|qualitative_open|date_time|geopoint|geotrace|geoshape|media_photo|media_audio|media_video|media_file|barcode|acknowledge|ranking|scale|calculate|note",
      "type": "integer|decimal|text|select_one|select_multiple|date|time|datetime|geopoint|geotrace|geoshape|image|audio|video|file|barcode|acknowledge|rank|range|calculate|note",
      "required": true,
      "hint": "",
      "choices": [],
      "group": "nom du groupe ou null",
      "formats": [],
      "suggested_format_idx": 0,
      "suggestions": []
    }
  ],
  "groups": ["Groupe 1"]
}

=== FORMATS SELON LA CLASSE ===

quantitative:
formats = [
  {"id":"A","name":"Nombre entier","type":"integer","note":"Ex : 25, 150, 3 — sans virgule"},
  {"id":"B","name":"Nombre avec virgule","type":"decimal","note":"Ex : 65,5 kg, 37,8°C, 12,45%"},
  {"id":"C","name":"Valeur sur une échelle","type":"range","note":"Ex : satisfaction de 1 à 10, score de 0 à 5"}
]
suggested_format_idx = 0 si entier attendu, 1 si décimal, 2 si échelle

qualitative_choice:
formats = [
  {"id":"A","name":"Une seule réponse au choix","type":"select_one","note":"Ex : Masculin ou Féminin — le répondant ne coche qu'une case"},
  {"id":"B","name":"Plusieurs réponses possibles","type":"select_multiple","note":"Ex : Maladies déclarées — le répondant peut cocher plusieurs cases"},
  {"id":"C","name":"Réponse écrite libre","type":"text","note":"Le répondant écrit lui-même sa réponse sans liste prédéfinie"}
]

qualitative_open:
formats = [
  {"id":"A","name":"Réponse courte","type":"text","note":"Ex : nom, prénom, profession — quelques mots"},
  {"id":"B","name":"Réponse longue","type":"text","note":"Ex : commentaire, description, observation — plusieurs phrases"}
]

date_time:
formats = [
  {"id":"A","name":"Date (jour/mois/année)","type":"date","note":"Ex : 15/03/2024"},
  {"id":"B","name":"Heure","type":"time","note":"Ex : 14h30"},
  {"id":"C","name":"Date et heure ensemble","type":"datetime","note":"Ex : 15/03/2024 à 14h30"}
]

geopoint: formats = [{"id":"A","name":"Localisation GPS (un point précis)","type":"geopoint","note":"Capture automatiquement la position GPS du téléphone"}]
geotrace: formats = [{"id":"A","name":"Tracer un chemin GPS","type":"geotrace","note":"Enregistre un trajet ou une route sur la carte"}]
geoshape: formats = [{"id":"A","name":"Délimiter une zone GPS","type":"geoshape","note":"Dessine un périmètre ou une superficie sur la carte"}]
media_photo: formats = [{"id":"A","name":"Prendre une photo","type":"image","note":"L'enquêteur prend une photo avec l'appareil photo"}]
media_audio: formats = [{"id":"A","name":"Enregistrer un son","type":"audio","note":"Enregistre un message audio ou un entretien"}]
media_video: formats = [{"id":"A","name":"Enregistrer une vidéo","type":"video","note":"Enregistre une vidéo avec la caméra"}]
media_file: formats = [{"id":"A","name":"Joindre un fichier","type":"file","note":"Le répondant joint un document PDF, Excel ou autre"}]
barcode: formats = [{"id":"A","name":"Scanner un code-barres","type":"barcode","note":"Scanne un code QR ou code-barres avec la caméra"}]
acknowledge: formats = [{"id":"A","name":"Case à cocher pour confirmer","type":"acknowledge","note":"Ex : J'ai lu et accepté les conditions — une seule case à cocher"}]
ranking: formats = [{"id":"A","name":"Classer par ordre de préférence","type":"rank","note":"Ex : Classer ces services du plus important au moins important"}]
scale: formats = [{"id":"A","name":"Valeur sur une échelle","type":"range","note":"Ex : Satisfaction de 1 à 5, douleur de 0 à 10"}]
calculate: formats = [{"id":"A","name":"Valeur calculée automatiquement","type":"calculate","note":"La valeur est calculée à partir des autres réponses, invisible pour le répondant"}]
note: formats = [{"id":"A","name":"Message d'information","type":"note","note":"Affiche un texte ou une instruction sans demander de réponse"}]

=== SUGGESTIONS ===
Analyse TOUTE la logique du questionnaire. Pour chaque question, génère les suggestions pertinentes :
Format : {"type":"skip_logic|calculate|constraint","label":"libellé court en langage simple","description":"explication claire sans jargon","value":"formule XLSForm technique","confidence":"high|medium|low"}

- skip_logic : saut conditionnel. La condition s'applique sur la question qui doit s'afficher conditionnellement. value = formule XLSForm ex: "\${id_q} = 'oui'" ou "selected(\${id_q}, 'valeur')"
- calculate : calcul auto. Toutes formes : somme, différence, produit, quotient, pourcentage, moyenne, min, max. value = formule XLSForm
- constraint : contrainte logique. value = formule XLSForm ex: ". >= 0 and . <= 150"

Ne génère des suggestions que si vraiment pertinent. confidence=high si explicitement dans le doc, medium si fortement suggéré.

=== AUTRES RÈGLES ===
- Toutes questions required:true par défaut sauf si explicitement facultatif
- Respecte l'ordre original. Regroupe par thématique dans "group"
- Extrais TOUTES les questions et sous-questions
- Pour qualitative_choice : extrais les vraies modalités dans choices[]
- Outil cible: \${tool}
`;



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
        system,
        messages: [{ role: 'user', content: `Analyse ce questionnaire et extrais sa structure:\n\n${text.slice(0, 8000)}` }]
      })
    });

    if (!response.ok) {
      const errData = await response.json();
      console.error('[CLAUDE ERROR]', errData);
      return res.status(502).json({ error: 'CLAUDE_ERROR', message: 'Erreur API analyse.' });
    }

    const data = await response.json();
    const rawText = data.content?.[0]?.text || '{}';

    let form;
    try {
      // Nettoyer les balises markdown si présentes
      let cleanText = rawText.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      const jsonMatch = cleanText.match(/\{[\s\S]*\}/);
      let jsonStr = jsonMatch ? jsonMatch[0] : cleanText;

      // Tenter de parser directement
      try {
        form = JSON.parse(jsonStr);
      } catch(e) {
        // Réparer JSON tronqué en fermant les structures ouvertes
        console.log('[PARSE] Tentative réparation JSON tronqué...');
        const opens = (jsonStr.match(/\[/g)||[]).length - (jsonStr.match(/\]/g)||[]).length;
        const openb = (jsonStr.match(/\{/g)||[]).length - (jsonStr.match(/\}/g)||[]).length;
        // Supprimer la dernière virgule si présente avant fermeture
        jsonStr = jsonStr.replace(/,\s*$/, '');
        for(let i = 0; i < opens; i++) jsonStr += ']';
        for(let i = 0; i < openb; i++) jsonStr += '}';
        form = JSON.parse(jsonStr);
        console.log('[PARSE] ✓ JSON réparé avec succès');
      }
    } catch (parseErr) {
      console.error('[PARSE ERROR]', parseErr, rawText.slice(0, 200));
      return res.status(502).json({ error: 'PARSE_ERROR', message: 'Réponse invalide de l\'analyse.' });
    }

    if (!form.questions || !Array.isArray(form.questions)) {
      return res.status(422).json({ error: 'NO_QUESTIONS', message: 'Aucune question détectée dans le document.' });
    }

    console.log(`[ANALYSE] ✓ ${form.questions.length} questions extraites`);
    res.json({ success: true, form });

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
