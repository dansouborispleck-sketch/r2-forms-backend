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
Tu analyses des questionnaires et guides d'entretien et tu extrais leur structure pour les convertir en masques de saisie numériques.

Réponds UNIQUEMENT en JSON valide, sans markdown ni texte autour. Pas de balises \`\`\`json.
Format exact:
{
  "title": "titre du formulaire",
  "questions": [
    {
      "id": "q1",
      "num": 1,
      "label": "libellé complet de la question",
      "type": "text",
      "required": false,
      "hint": "",
      "choices": ["choix1", "choix2"],
      "group": "nom du groupe thématique ou null",
      "formats": [
        {"id": "A", "name": "Choix unique", "type": "select_one", "note": "Une seule réponse possible"},
        {"id": "B", "name": "Choix multiple", "type": "select_multiple", "note": "Plusieurs réponses possibles"},
        {"id": "C", "name": "Texte libre", "type": "text", "note": "Le répondant écrit sa propre réponse"}
      ]
    }
  ],
  "groups": ["Groupe 1", "Groupe 2"]
}

Règles importantes:
- Extrais TOUTES les questions du document, même les sous-questions
- Pour les questions à choix, mets les vraies options dans "choices" et propose select_one + select_multiple + text dans formats
- Pour les questions ouvertes, propose text + integer/decimal/date selon le contexte dans formats (2 options max)
- Pour GPS/localisation: type geopoint
- Pour photos: type image
- Regroupe les questions par thématique logique dans "group"
- Respecte l'ordre original des questions
- Outil cible: ${tool}`;

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
      const row = {
        type: t,
        name: q.id || ('q' + q.num),
        label: q.label || '',
        required: q.required ? 'yes' : 'no',
        hint: q.hint || ''
      };

      if (t === 'select_one' || t === 'select_multiple') {
        const listName = 'list_' + (q.id || q.num);
        row.type = t + ' ' + listName;
        if (!choiceListsSeen.has(listName)) {
          choiceListsSeen.add(listName);
          (q.choices || []).forEach((c, i) => {
            const label = typeof c === 'string' ? c : (c.label || String(c));
            choices.push({ list_name: listName, name: 'c' + (i + 1), label: label });
          });
        }
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
