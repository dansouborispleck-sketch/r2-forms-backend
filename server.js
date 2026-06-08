// ============================================
// R2 Forms Backend v4.1
// Sécurisé · Optimisé · Multi-outils
// ============================================

require('dotenv').config();

const express = require('express');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const helmet = require('helmet');
const { body, validationResult } = require('express-validator');
const multer = require('multer');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const pdf = require('pdf-parse');
const fetch = require('node-fetch');
const path = require('path');
const crypto = require('crypto');
const { createClient } = require('redis');

// ============ CONFIGURATION ============

const PORT = process.env.PORT || 4000;
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const OPENAI_API_KEY = process.env.OPENAI_API_KEY; // Fallback
const FLUTTERWAVE_SECRET_KEY = process.env.FLUTTERWAVE_SECRET_KEY; // Pour paiement futur

// Origines autorisées (votre GitHub Pages + dev local)
const ALLOWED_ORIGINS = [
  'https://dansouborispleck-sketch.github.io',
  'https://r2-forms.vercel.app',
  'http://localhost:3000',
  'http://127.0.0.1:5500'
];

// Redis client (cache)
let redis = null;
if (process.env.REDIS_URL) {
  redis = createClient({ url: process.env.REDIS_URL });
  redis.on('error', (err) => console.error('[REDIS ERROR]', err.message));
  redis.connect().catch(() => console.log('[REDIS] Non connecté, cache désactivé'));
}

// ============ LOGGING ============

const logger = {
  info: (obj) => console.log(JSON.stringify({ level: 'info', timestamp: new Date().toISOString(), ...obj })),
  warn: (obj) => console.log(JSON.stringify({ level: 'warn', timestamp: new Date().toISOString(), ...obj })),
  error: (obj) => console.error(JSON.stringify({ level: 'error', timestamp: new Date().toISOString(), ...obj }))
};

// ============ HELPERS ============

// Hash pour cache
function hashContent(text, tool) {
  return crypto.createHash('sha256').update(text.slice(0, 5000) + tool).digest('hex');
}

// Sauvegarder dans le cache
async function saveCache(key, data, ttl = 7 * 24 * 60 * 60) {
  if (!redis) return;
  try {
    await redis.setEx(key, ttl, JSON.stringify(data));
    logger.info({ event: 'CACHE_SET', key: key.slice(0, 16) });
  } catch (e) {
    logger.warn({ event: 'CACHE_WRITE_ERROR', error: e.message });
  }
}

// Vérifier le cache
async function checkCache(key) {
  if (!redis) return null;
  try {
    const cached = await redis.get(key);
    if (cached) {
      logger.info({ event: 'CACHE_HIT', key: key.slice(0, 16) });
      return JSON.parse(cached);
    }
  } catch (e) {
    logger.warn({ event: 'CACHE_READ_ERROR', error: e.message });
  }
  return null;
}

// Parser la réponse Claude (JSON robuste)
function parseClaudeResponse(rawText) {
  let jsonStr = rawText
    .replace(/```json\n?/g, '')
    .replace(/```\n?/g, '')
    .trim();
  
  const startIdx = jsonStr.indexOf('{');
  if (startIdx > 0) jsonStr = jsonStr.slice(startIdx);
  
  try {
    return JSON.parse(jsonStr);
  } catch (e) {
    // Réparation JSON tronqué
    const openBrackets = (jsonStr.match(/\[/g) || []).length - (jsonStr.match(/\]/g) || []).length;
    const openBraces = (jsonStr.match(/\{/g) || []).length - (jsonStr.match(/\}/g) || []).length;
    
    jsonStr = jsonStr.replace(/,\s*$/, '');
    for (let i = 0; i < Math.max(0, openBrackets); i++) jsonStr += ']';
    for (let i = 0; i < Math.max(0, openBraces); i++) jsonStr += '}';
    
    try {
      return JSON.parse(jsonStr);
    } catch (e2) {
      throw new Error('JSON parsing failed after repair attempt');
    }
  }
}

// Fallback OpenAI (10x moins cher)
async function analyseWithOpenAI(text, tool) {
  if (!OPENAI_API_KEY) throw new Error('Fallback OpenAI non configuré');
  
  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: 'gpt-4o-mini',
      messages: [
        { role: 'system', content: OPTIMIZED_SYSTEM_PROMPT },
        { role: 'user', content: 'Extrais TOUTES les questions:\n\n' + text.slice(0, 6000) }
      ],
      max_tokens: 16000,
      temperature: 0.1
    })
  });

  if (!response.ok) throw new Error('Fallback API error: ' + response.status);
  const data = await response.json();
  return parseClaudeResponse(data.choices[0].message.content);
}

// Prompt système optimisé (court = moins cher)
const OPTIMIZED_SYSTEM_PROMPT = `Expert collecte données terrain. Extrais structure questionnaire en JSON compact.

RÈGLES:
1. TOUTES questions: numérotées, sans numéro, sous-questions
2. GROUPES: titres complets de sections
3. "Autres" → question libre auto après (required=false, label="Si autre, précisez :")
4. SAUTS: détecter conditions implicites
5. choice_values: codes numériques [0,1,2...]
6. coherence_report: liste observations

FORMAT JSON:
{"title":"...","coherence_report":[],"questions":[{"id":"q1","num":1,"label":"...","question_class":"CLASS","type":"TYPE","required":true,"hint":"","choices":[],"choice_values":[],"group":"TITRE","formats":[],"suggested_format_idx":0,"suggestions":[]}],"groups":[]}

CLASSES: quantitative,qualitative_choice,qualitative_open,date_time,geopoint,geotrace,geoshape,media_photo,media_audio,media_video,media_file,barcode,acknowledge,ranking,scale,calculate,note

SUGGESTIONS: {"type":"skip_logic|calculate|constraint","label":"court","description":"clair","value":"formule XLSForm","confidence":"high|medium|low"}

required=true par défaut. JSON compact.`;

// ============ APP SETUP ============

const app = express();

// Helmet (sécurité headers)
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      styleSrc: ["'self'", "'unsafe-inline'", "https://fonts.googleapis.com"],
      fontSrc: ["'self'", "https://fonts.gstatic.com"],
      scriptSrc: ["'self'", "'unsafe-inline'"],
      imgSrc: ["'self'", "data:", "https:"],
      connectSrc: ["'self'", "https://api.anthropic.com", "https://api.openai.com"]
    }
  }
}));

// CORS restreint
app.use(cors({
  origin: function (origin, callback) {
    if (!origin) {
      if (process.env.NODE_ENV === 'production') {
        return callback(new Error('CORS: origin manquant'), false);
      }
      return callback(null, true);
    }
    if (ALLOWED_ORIGINS.includes(origin)) return callback(null, true);
    logger.warn({ event: 'CORS_REJECTED', origin });
    return callback(new Error('CORS: origin non autorisé'), false);
  },
  methods: ['POST', 'GET', 'OPTIONS'],
  allowedHeaders: ['Content-Type'],
  credentials: false
}));

// Rate limiting général
const generalLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 20,
  standardHeaders: true,
  legacyHeaders: false,
  message: {
    error: 'RATE_LIMIT',
    message: 'Trop de requêtes. Limite: 20 requêtes / 15 minutes.',
    retryAfter: 900
  },
  skip: (req) => req.path === '/health'
});

// Rate limiting strict pour les routes coûteuses
const expensiveLimiter = rateLimit({
  windowMs: 60 * 60 * 1000,
  max: 10,
  message: {
    error: 'RATE_LIMIT_EXPENSIVE',
    message: 'Limite d\'analyses atteinte (10/heure).',
    retryAfter: 3600
  }
});

app.use('/api/', generalLimiter);
app.use('/api/analyse', expensiveLimiter);
app.use('/api/correct', expensiveLimiter);
app.use('/api/verify', expensiveLimiter);

// Body parser
app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: true, limit: '2mb' }));

// Multer upload
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024, files: 1 },
  fileFilter: (req, file, cb) => {
    const allowedExts = ['txt', 'csv', 'pdf', 'docx', 'doc', 'xlsx', 'xls', 'odt'];
    const allowedMimes = [
      'text/plain', 'text/csv',
      'application/pdf',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.oasis.opendocument.text'
    ];
    const ext = file.originalname.split('.').pop().toLowerCase();
    
    if (allowedExts.includes(ext) && allowedMimes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error('Format non autorisé: ' + ext), false);
    }
  }
});

// Middleware validation
const handleValidation = (req, res, next) => {
  const errors = validationResult(req);
  if (!errors.isEmpty()) {
    return res.status(400).json({
      error: 'VALIDATION_ERROR',
      details: errors.array().map(e => ({ field: e.path, message: e.msg }))
    });
  }
  next();
};

// Middleware logging
app.use((req, res, next) => {
  req.startTime = Date.now();
  res.on('finish', () => {
    logger.info({
      method: req.method,
      path: req.path,
      ip: req.ip?.slice(0, 15),
      status: res.statusCode,
      duration: (Date.now() - req.startTime) + 'ms'
    });
  });
  next();
});

// ============ ROUTES ============

// Health check
app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    version: '4.1.0',
    timestamp: new Date().toISOString(),
    cache: redis ? 'connected' : 'disabled',
    anthropic: ANTHROPIC_API_KEY ? 'configured' : 'missing',
    openai: OPENAI_API_KEY ? 'configured' : 'missing'
  });
});

// ============ IMPORT ============
app.post('/api/import', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Aucun fichier recu' });

    const { originalname, buffer } = req.file;
    const ext = originalname.split('.').pop().toLowerCase();
    let text = '';

    logger.info({ event: 'IMPORT_START', file: originalname, size: buffer.length });

    if (['txt', 'csv'].includes(ext)) {
      text = buffer.toString('utf-8');
    } else if (ext === 'pdf') {
      const d = await pdf(buffer);
      text = d.text;
      if (!text || text.trim().length < 20) {
        return res.status(422).json({
          error: 'PDF_SCANNED',
          message: 'PDF scanné illisible. Collez le texte directement.'
        });
      }
    } else if (ext === 'docx') {
      const r = await mammoth.extractRawText({ buffer });
      text = r.value;
    } else if (ext === 'doc') {
      text = buffer.toString('latin1')
        .replace(/[^\x20-\x7E\n\r\u00C0-\u024F]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
      if (text.length < 30) {
        return res.status(422).json({
          error: 'DOC_OLD',
          message: 'Format .doc non supporté. Enregistrez en .docx.'
        });
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

    logger.info({ event: 'IMPORT_OK', chars: text.length });
    res.json({
      success: true,
      text: text,
      metadata: { filename: originalname, chars: text.length }
    });

  } catch (err) {
    logger.error({ event: 'IMPORT_ERROR', error: err.message });
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ ANALYSE ============
const validateAnalyse = [
  body('text').isString().trim().isLength({ min: 20, max: 50000 }).withMessage('Texte: 20-50000 caractères'),
  body('tool').isIn(['kobo', 'odk', 'jotform', 'xlsform', 'google']).withMessage('Outil non supporté')
];

app.post('/api/analyse', validateAnalyse, handleValidation, async (req, res, next) => {
  // Vérifier le cache
  const cacheKey = 'analyse:' + hashContent(req.body.text, req.body.tool);
  const cached = await checkCache(cacheKey);
  if (cached) {
    return res.json({
      success: true,
      form: cached,
      cached: true,
      saved: '~0.15 USD'
    });
  }
  req.cacheKey = cacheKey;
  next();
}, async (req, res) => {
  try {
    const { text, tool } = req.body;
    if (!ANTHROPIC_API_KEY) return res.status(500).json({ error: 'Cle API Anthropic manquante' });

    // Troncature intelligente
    const MAX_INPUT = 6000;
    let inputText = text;
    if (text.length > MAX_INPUT) {
      const keepStart = Math.floor(MAX_INPUT * 0.7);
      const keepEnd = Math.floor(MAX_INPUT * 0.3);
      inputText = text.slice(0, keepStart) + '\n\n[...]\n\n' + text.slice(-keepEnd);
    }

    logger.info({ event: 'ANALYSE_START', tool, textLength: text.length, inputLength: inputText.length });

    // Appel Claude
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
        system: OPTIMIZED_SYSTEM_PROMPT,
        messages: [{ role: 'user', content: 'Extrais TOUTES les questions:\n\n' + inputText }]
      })
    });

    if (!response.ok) {
      // Fallback OpenAI
      if (OPENAI_API_KEY) {
        logger.warn({ event: 'CLAUDE_FALLBACK', status: response.status });
        try {
          const fallbackForm = await analyseWithOpenAI(text, tool);
          if (req.cacheKey) await saveCache(req.cacheKey, fallbackForm);
          return res.json({
            success: true,
            form: fallbackForm,
            fallback: 'openai',
            warning: 'Analyse par modèle alternatif (coût réduit)'
          });
        } catch (fallbackErr) {
          logger.error({ event: 'FALLBACK_FAILED', error: fallbackErr.message });
        }
      }
      
      const e = await response.json().catch(() => ({}));
      logger.error({ event: 'CLAUDE_ERROR', status: response.status, error: e });
      return res.status(502).json({ error: 'CLAUDE_ERROR', message: 'Erreur API Claude.' });
    }

    const data = await response.json();
    const rawText = data.content?.[0]?.text || '{}';

    // Log coûts
    const usage = data.usage;
    if (usage) {
      const cost = (usage.input_tokens * 3 / 1_000_000) + (usage.output_tokens * 15 / 1_000_000);
      logger.info({
        event: 'CLAUDE_COST',
        inputTokens: usage.input_tokens,
        outputTokens: usage.output_tokens,
        costUSD: cost.toFixed(4)
      });
    }

    const form = parseClaudeResponse(rawText);

    if (!form.questions || !Array.isArray(form.questions)) {
      return res.status(422).json({ error: 'NO_QUESTIONS', message: 'Aucune question detectee.' });
    }

    // Sauvegarder cache
    if (req.cacheKey) await saveCache(req.cacheKey, form);

    logger.info({ event: 'ANALYSE_OK', questions: form.questions.length });
    res.json({
      success: true,
      form: form,
      truncated: text.length > MAX_INPUT,
      stats: {
        questions: form.questions.length,
        groups: form.groups?.length || 0,
        suggestions: form.questions.reduce((a, q) => a + (q.suggestions?.length || 0), 0)
      }
    });

  } catch (err) {
    logger.error({ event: 'ANALYSE_ERROR', error: err.message });
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ CORRECTION ============
const validateCorrect = [
  body('form').isObject().withMessage('Formulaire requis'),
  body('instructions').isString().trim().isLength({ min: 5, max: 2000 }).withMessage('Instructions: 5-2000 caractères')
];

app.post('/api/correct', validateCorrect, handleValidation, expensiveLimiter, async (req, res) => {
  try {
    const { form, instructions } = req.body;
    if (!ANTHROPIC_API_KEY) return res.status(500).json({ error: 'Cle API manquante' });

    logger.info({ event: 'CORRECT_START', instructionsLength: instructions.length });

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
        system: 'Expert collecte de donnees. Applique exactement les corrections au formulaire JSON. Retourne JSON corrige UNIQUEMENT sans markdown.',
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
    const corrected = parseClaudeResponse(m ? m[0] : raw);

    logger.info({ event: 'CORRECT_OK', questions: corrected.questions?.length });
    res.json({ success: true, form: corrected });

  } catch (err) {
    logger.error({ event: 'CORRECT_ERROR', error: err.message });
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ VÉRIFICATION (optionnelle) ============
app.post('/api/verify', expensiveLimiter, async (req, res) => {
  try {
    const { form, force = false } = req.body;
    if (!form) return res.status(400).json({ error: 'Formulaire manquant' });

    // Skip si pas de logique complexe et pas forcé
    const hasComplexLogic = form.questions?.some(q =>
      q.suggestions?.some(s => s.type === 'skip_logic') ||
      q.relevant ||
      q.question_class === 'calculate'
    );

    if (!force && !hasComplexLogic) {
      return res.json({
        valid: true,
        issues: [],
        skipped: true,
        message: 'Verification automatique: pas de logique complexe'
      });
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
        system: 'Expert XLSForm. Verifie le contenu et reponds en JSON: {"valid":true/false,"issues":["probleme"],"fixed_questions":[{"id":"q1","relevant":"formule"}]}',
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

    // Appliquer corrections automatiques
    if (result.fixed_questions) {
      result.fixed_questions.forEach(fix => {
        const q = form.questions.find(q => q.id === fix.id);
        if (q && fix.relevant) q.relevant = fix.relevant;
      });
    }

    res.json({
      valid: result.valid,
      issues: result.issues || [],
      form: form,
      verified: true
    });

  } catch (err) {
    logger.error({ event: 'VERIFY_ERROR', error: err.message });
    res.json({ valid: true, issues: [], error: err.message });
  }
});

// ============ BUILDER KOBOTOOLBOX (partagé) ============
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

    if (gname !== 'general') {
      survey.push({ type: 'begin_group', name: gId, label: gname });
    }

    qs.forEach(q => {
      const fmtIdx = q.validatedFormatIdx !== undefined
        ? q.validatedFormatIdx
        : (q.suggested_format_idx || 0);
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

      // Skip logic
      let relevant = q.relevant || '';
      if (!relevant && q.suggestions) {
        q.suggestions.forEach((s, si) => {
          if (s.type === 'skip_logic' && s.value && q.confirmedSuggestions?.[si]) {
            relevant = s.value;
          }
        });
      }
      if (relevant) row.relevant = relevant;

      // Calculate
      if (t === 'calculate') {
        let calc = q.calculation || '';
        if (!calc && q.suggestions) {
          q.suggestions.forEach((s, si) => {
            if (s.type === 'calculate' && s.value && q.confirmedSuggestions?.[si]) {
              calc = s.value;
            }
          });
        }
        if (calc) row.calculation = calc;
        delete row.required;
      }

      // Constraints
      const constraints = [];
      if (q.numMin !== '' && q.numMin != null) constraints.push('. >= ' + q.numMin);
      if (q.numMax !== '' && q.numMax != null) constraints.push('. <= ' + q.numMax);
      if (q.numDigitsBefore) constraints.push('string-length(substring-before(string(.), \'.\')) <= ' + q.numDigitsBefore);
      if (q.numDigitsAfter) constraints.push('string-length(substring-after(string(.), \'.\')) <= ' + q.numDigitsAfter);
      if (q.constraint?.trim()) constraints.push(q.constraint.trim());
      if (q.suggestions) {
        q.suggestions.forEach((s, si) => {
          if (s.type === 'constraint' && s.value && q.confirmedSuggestions?.[si]) {
            constraints.push(s.value);
          }
        });
      }
      if (constraints.length > 0) {
        row.constraint = constraints.join(' and ');
        row.constraint_message = 'Valeur hors limites';
      }

      // Range parameters
      if (t === 'range') {
        row.parameters = 'start=' + (q.numMin || 1) + ' end=' + (q.numMax || 10);
        delete row.required;
      }

      // Note / calculate: pas de required
      if (t === 'note' || t === 'calculate') delete row.required;

      // Choices
      if (t === 'select_one' || t === 'select_multiple' || t === 'rank') {
        const listName = 'list_' + name;
        row.type = t + ' ' + listName;
        if (!seen.has(listName)) {
          seen.add(listName);
          const choiceVals = q.choice_values || [];
          (q.choices || []).forEach((c, i) => {
            const label = typeof c === 'string' ? c : (c.label || String(c));
            const val = (choiceVals[i] !== undefined && choiceVals[i] !== '') ? String(choiceVals[i]) : String(i);
            choices.push({ list_name: listName, name: val, label: label });
          });
        }
      }

      survey.push(row);
    });

    if (gname !== 'general') {
      survey.push({ type: 'end_group', name: gId });
    }
  });

  const formId = (form.title || 'formulaire').replace(/\s+/g, '_').toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9_]/g, '').slice(0, 32);

  return {
    survey: survey,
    choices: choices,
    settings: [{ form_title: form.title || 'Formulaire', form_id: formId, version: '1' }]
  };
}

// ============ DÉPLOIEMENT KOBO (outil sélectionné) ============
const validateDeploy = [
  body('form').isObject().withMessage('Formulaire requis'),
  body('credentials').isObject().withMessage('Credentials requis')
];

app.post('/api/deploy/kobo', validateDeploy, handleValidation, async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { username, password, server = 'https://kf.kobotoolbox.org' } = credentials;

    if (!form || !username || !password) {
      return res.status(400).json({ error: 'Donnees manquantes' });
    }

    logger.info({ event: 'DEPLOY_KOBO_START', server });

    // 1. Authentification
    const tokenRes = await fetch(server + '/token/?format=json', {
      headers: {
        'Authorization': 'Basic ' + Buffer.from(username + ':' + password).toString('base64')
      }
    });

    if (!tokenRes.ok) {
      return res.status(401).json({
        error: 'AUTH_ERROR',
        message: 'Identifiants KoboToolbox incorrects.'
      });
    }

    const { token } = await tokenRes.json();
    const auth = {
      'Authorization': 'Token ' + token,
      'Content-Type': 'application/json'
    };

    // 2. Créer l'asset
    const assetRes = await fetch(server + '/api/v2/assets/?format=json', {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({
        name: form.title || 'Formulaire R2',
        asset_type: 'survey'
      })
    });

    if (!assetRes.ok) {
      return res.status(502).json({
        error: 'ASSET_ERROR',
        message: 'Erreur creation formulaire.'
      });
    }

    const { uid } = await assetRes.json();

    // 3. Mettre à jour le contenu
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
      const errBody = await patchRes.text();
      logger.error({ event: 'KOBOPATCH_ERROR', error: errBody.slice(0, 200) });
      return res.status(502).json({
        error: 'PATCH_ERROR',
        message: 'Erreur import questionnaire.'
      });
    }

    // 4. Déployer
    await fetch(server + '/api/v2/assets/' + uid + '/deployment/?format=json', {
      method: 'POST',
      headers: auth,
      body: JSON.stringify({ active: true })
    });

    logger.info({ event: 'DEPLOY_KOBO_OK', uid });

    res.json({
      success: true,
      uid: uid,
      url: server + '/#/forms/' + uid + '/summary',
      questions: form.questions?.length || 0
    });

  } catch (err) {
    logger.error({ event: 'DEPLOY_KOBO_ERROR', error: err.message });
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ DÉPLOIEMENT ODK CENTRAL (outil sélectionné) ============
app.post('/api/deploy/odk', validateDeploy, handleValidation, async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { server, email, password } = credentials;

    if (!form || !server || !email || !password) {
      return res.status(400).json({
        error: 'Donnees manquantes',
        message: 'Server, email et password requis'
      });
    }

    logger.info({ event: 'DEPLOY_ODK_START', server });

    // 1. Authentification (session cookie)
    const sessionRes = await fetch(server + '/v1/sessions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ email, password })
    });

    if (!sessionRes.ok) {
      return res.status(401).json({
        error: 'AUTH_ERROR',
        message: 'Identifiants ODK Central incorrects.'
      });
    }

    const { token } = await sessionRes.json();
    const cookie = `__Host-session=${token}`;

    // 2. Lister les projets
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
        // Créer le projet
        const createProjectRes = await fetch(server + '/v1/projects', {
          method: 'POST',
          headers: { 'Cookie': cookie, 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: 'R2 Forms' })
        });
        const newProject = await createProjectRes.json();
        projectId = newProject.id;
      }
    }

    // 3. Construire le XLSForm en vrai .xlsx
    const wb = XLSX.utils.book_new();

    // Sheet survey
    const koboContent = buildKoboContent(form);
    const surveyRows = koboContent.survey.map(row => ({
      type: row.type,
      name: row.name,
      label: row.label,
      required: row.required,
      hint: row.hint || '',
      relevant: row.relevant || '',
      constraint: row.constraint || '',
      calculation: row.calculation || '',
      parameters: row.parameters || ''
    }));
    const wsSurvey = XLSX.utils.json_to_sheet(surveyRows);
    XLSX.utils.book_append_sheet(wb, wsSurvey, 'survey');

    // Sheet choices
    if (koboContent.choices.length > 0) {
      const wsChoices = XLSX.utils.json_to_sheet(koboContent.choices);
      XLSX.utils.book_append_sheet(wb, wsChoices, 'choices');
    }

    // Sheet settings
    const wsSettings = XLSX.utils.json_to_sheet(koboContent.settings);
    XLSX.utils.book_append_sheet(wb, wsSettings, 'settings');

    const xlsxBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    // 4. Upload du formulaire
    const formData = new (require('form-data'))();
    formData.append('xlsForm', new Blob([xlsxBuffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }), 'form.xlsx');

    const uploadRes = await fetch(server + `/v1/projects/${projectId}/forms?publish=true`, {
      method: 'POST',
      headers: {
        'Cookie': cookie,
        ...formData.getHeaders()
      },
      body: formData
    });

    if (!uploadRes.ok) {
      const errText = await uploadRes.text();
      logger.error({ event: 'ODK_UPLOAD_ERROR', status: uploadRes.status, error: errText.slice(0, 200) });
      return res.status(502).json({
        error: 'UPLOAD_ERROR',
        message: 'Erreur upload formulaire ODK.'
      });
    }

    const uploadData = await uploadRes.json();

    logger.info({ event: 'DEPLOY_ODK_OK', projectId, formId: uploadData.xmlFormId });

    res.json({
      success: true,
      projectId: projectId,
      formId: uploadData.xmlFormId,
      url: `${server}/#/projects/${projectId}/forms/${uploadData.xmlFormId}`,
      questions: form.questions?.length || 0
    });

  } catch (err) {
    logger.error({ event: 'DEPLOY_ODK_ERROR', error: err.message });
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ DÉPLOIEMENT JOTFORM (outil sélectionné) ============
app.post('/api/deploy/jotform', validateDeploy, handleValidation, async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { apiKey } = credentials;

    if (!form || !apiKey) {
      return res.status(400).json({ error: 'Donnees manquantes' });
    }

    logger.info({ event: 'DEPLOY_JOTFORM_START' });

    // 1. Vérifier la clé API
    const userRes = await fetch('https://api.jotform.com/user?apiKey=' + apiKey);
    if (!userRes.ok) {
      return res.status(401).json({
        error: 'AUTH_ERROR',
        message: 'Cle API JotForm incorrecte.'
      });
    }

    const userData = await userRes.json();
    if (userData.responseCode !== 200) {
      return res.status(401).json({
        error: 'AUTH_ERROR',
        message: 'Cle API JotForm invalide.'
      });
    }

    // 2. Créer le formulaire
    const createRes = await fetch('https://api.jotform.com/form?apiKey=' + apiKey, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: 'questions%5B0%5D%5Btype%5D=control_head&questions%5B0%5D%5Btext%5D=' +
        encodeURIComponent(form.title || 'Formulaire R2') +
        '&properties%5Btitle%5D=' + encodeURIComponent(form.title || 'Formulaire R2')
    });

    if (!createRes.ok) {
      return res.status(502).json({
        error: 'CREATE_ERROR',
        message: 'Erreur creation formulaire JotForm.'
      });
    }

    const createData = await createRes.json();
    const formId = createData.content?.id;

    if (!formId) {
      return res.status(502).json({
        error: 'CREATE_ERROR',
        message: 'ID formulaire non recu.'
      });
    }

    // 3. Ajouter les questions
    const questions = form.questions || [];
    for (let qi = 0; qi < questions.length; qi++) {
      const q = questions[qi];
      const t = q.selectedType || q.type || 'text';
      let qType = 'control_textbox';

      if (t === 'select_one') qType = 'control_radio';
      else if (t === 'select_multiple') qType = 'control_checkbox';
      else if (t === 'integer' || t === 'decimal') qType = 'control_number';
      else if (t === 'date') qType = 'control_datetime';
      else if (t === 'image') qType = 'control_fileupload';

      const qData = new URLSearchParams();
      qData.append('questions[' + (qi + 1) + '][type]', qType);
      qData.append('questions[' + (qi + 1) + '][text]', q.label || '');
      qData.append('questions[' + (qi + 1) + '][required]', q.required ? 'Yes' : 'No');
      qData.append('questions[' + (qi + 1) + '][order]', String(qi + 1));

      if ((qType === 'control_radio' || qType === 'control_checkbox') && q.choices?.length > 0) {
        q.choices.forEach(c => {
          const label = typeof c === 'string' ? c : (c.label || String(c));
          qData.append('questions[' + (qi + 1) + '][options]', label);
        });
      }

      await fetch('https://api.jotform.com/form/' + formId + '/questions?apiKey=' + apiKey, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: qData.toString()
      });
    }

    logger.info({ event: 'DEPLOY_JOTFORM_OK', formId });

    res.json({
      success: true,
      formId: formId,
      url: 'https://www.jotform.com/' + formId,
      questions: questions.length
    });

  } catch (err) {
    logger.error({ event: 'DEPLOY_JOTFORM_ERROR', error: err.message });
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ DÉPLOIEMENT GOOGLE FORMS (outil sélectionné) ============
app.post('/api/deploy/google', validateDeploy, handleValidation, async (req, res) => {
  try {
    const { form, credentials } = req.body;
    const { accessToken } = credentials;

    if (!form || !accessToken) {
      return res.status(400).json({ error: 'Token Google manquant' });
    }

    logger.info({ event: 'DEPLOY_GOOGLE_START', questions: form.questions?.length });

    // 1. Créer le formulaire
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

    // 2. Construire les items
    const requests = [];
    let itemIndex = 0;

    const groups = {};
    form.questions.forEach(q => {
      const g = q.group || 'Général';
      if (!groups[g]) groups[g] = [];
      groups[g].push(q);
    });

    Object.entries(groups).forEach(([gname, qs]) => {
      // Page break pour chaque groupe (sauf le premier)
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

    // 3. Batch update
    if (requests.length > 0) {
      const batchRes = await fetch(`https://forms.googleapis.com/v1/forms/${formId}:batchUpdate`, {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + accessToken,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ requests })
      });

      if (!batchRes.ok) {
        const err = await batchRes.json();
        logger.warn({ event: 'GOOGLE_BATCH_WARN', error: JSON.stringify(err).slice(0, 300) });
      }
    }

    logger.info({ event: 'DEPLOY_GOOGLE_OK', formId });

    res.json({
      success: true,
      formId: formId,
      url: `https://docs.google.com/forms/d/${formId}/edit`,
      viewUrl: `https://docs.google.com/forms/d/e/${formId}/viewform`,
      questions: form.questions?.length || 0
    });

  } catch (err) {
    logger.error({ event: 'DEPLOY_GOOGLE_ERROR', error: err.message });
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ EXPORT XLSFORM (outil sélectionné) ============
app.post('/api/export/xlsform', async (req, res) => {
  try {
    const { form } = req.body;
    if (!form) return res.status(400).json({ error: 'Formulaire manquant' });

    logger.info({ event: 'EXPORT_XLSFORM_START' });

    const wb = XLSX.utils.book_new();
    const koboContent = buildKoboContent(form);

    // Sheet survey
    const surveyRows = koboContent.survey.map(row => ({
      type: row.type,
      name: row.name,
      label: row.label,
      required: row.required,
      hint: row.hint || '',
      relevant: row.relevant || '',
      constraint: row.constraint || '',
      calculation: row.calculation || '',
      parameters: row.parameters || ''
    }));
    const wsSurvey = XLSX.utils.json_to_sheet(surveyRows);
    XLSX.utils.book_append_sheet(wb, wsSurvey, 'survey');

    // Sheet choices
    if (koboContent.choices.length > 0) {
      const wsChoices = XLSX.utils.json_to_sheet(koboContent.choices);
      XLSX.utils.book_append_sheet(wb, wsChoices, 'choices');
    }

    // Sheet settings
    const wsSettings = XLSX.utils.json_to_sheet(koboContent.settings);
    XLSX.utils.book_append_sheet(wb, wsSettings, 'settings');

    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    const filename = (form.title || 'formulaire').replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '') + '.xlsx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(buffer);

    logger.info({ event: 'EXPORT_XLSFORM_OK', filename });

  } catch (err) {
    logger.error({ event: 'EXPORT_XLSFORM_ERROR', error: err.message });
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ TALLY.SO (outil sélectionné — pas d'API, template JSON) ============
app.post('/api/deploy/tally', async (req, res) => {
  try {
    const { form } = req.body;
    if (!form) return res.status(400).json({ error: 'Formulaire manquant' });

    logger.info({ event: 'DEPLOY_TALLY_START' });

    // Tally.so n'a pas d'API publique de création
    // Générer un template JSON pour import manuel
    const tallyTemplate = {
      title: form.title || 'Formulaire R2',
      description: 'Créé avec R2 Forms',
      fields: (form.questions || []).map((q, i) => {
        const t = q.selectedType || q.type;
        const base = {
          id: `field_${i}`,
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
      tallyTemplate: tallyTemplate,
      instructions: {
        fr: '1. Créez un formulaire vierge sur Tally.so\n2. Cliquez sur "Import"\n3. Collez le JSON fourni',
        en: '1. Create a blank form on Tally.so\n2. Click "Import"\n3. Paste the provided JSON'
      },
      questions: form.questions?.length || 0
    });

  } catch (err) {
    logger.error({ event: 'DEPLOY_TALLY_ERROR', error: err.message });
    res.status(500).json({ error: 'SERVER_ERROR', message: err.message });
  }
});

// ============ ERROR HANDLER ============

app.use((err, req, res, next) => {
  logger.error({
    event: 'UNHANDLED_ERROR',
    path: req.path,
    error: err.message,
    stack: err.stack?.slice(0, 200)
  });
  
  res.status(500).json({
    error: 'SERVER_ERROR',
    message: process.env.NODE_ENV === 'production' ? 'Erreur serveur' : err.message
  });
});

// 404
app.use((req, res) => {
  res.status(404).json({ error: 'NOT_FOUND', message: 'Route non trouvée' });
});

// ============ START ============

app.listen(PORT, () => {
  console.log(`
╔══════════════════════════════════════════════╗
║  R2 Forms Backend v4.1                       ║
║  Port: ${PORT}                                  ║
║  Cache: ${redis ? 'Redis OK' : 'Disabled'}                    ║
║  Anthropic: ${ANTHROPIC_API_KEY ? 'OK' : 'MISSING'}                    ║
║  OpenAI Fallback: ${OPENAI_API_KEY ? 'OK' : 'Disabled'}            ║
╚══════════════════════════════════════════════╝
  `);
});
