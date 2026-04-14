require('dotenv').config();
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const {
    getACSToken,
    getSpeechToken,
    getAvatarIceConfig,
    transcribeAudio,
    getAgentAnswer,
    processQuestion,
    config
} = require('./app');

const app = express();
const uploadsDir = path.join(__dirname, 'uploads');
const publicDir = path.join(__dirname, 'public');
const maxAudioUploadBytes = Number(process.env.MAX_AUDIO_UPLOAD_BYTES || 15 * 1024 * 1024);
const allowedOrigins = (process.env.CORS_ORIGIN || '')
    .split(',')
    .map(origin => origin.trim())
    .filter(Boolean);

// Ensure required directories exist.
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}
if (!fs.existsSync(publicDir)) {
    fs.mkdirSync(publicDir, { recursive: true });
}

if (allowedOrigins.length > 0) {
    app.use(cors({
        origin(origin, callback) {
            if (!origin || allowedOrigins.includes(origin)) {
                return callback(null, true);
            }

            return callback(new Error('Origin not allowed by CORS'));
        }
    }));
}

app.use(express.json({ limit: '1mb' }));
app.use(express.static(publicDir));

const upload = multer({
    dest: uploadsDir,
    limits: { fileSize: maxAudioUploadBytes }
});

function requireApiKey(req, res, next) {
    const configuredKey = process.env.APP_API_KEY;

    if (!configuredKey) {
        return next();
    }

    const providedKey = req.get('x-app-api-key');
    if (providedKey !== configuredKey) {
        return res.status(401).json({ error: 'Unauthorized' });
    }

    return next();
}

function cleanupUploadedFile(filePath) {
    if (filePath && fs.existsSync(filePath)) {
        fs.unlinkSync(filePath);
    }
}

// Health check
app.get('/health', (req, res) => {
    res.json({
        status: 'Running',
        wakeWord: config.avatar.wakeWord,
        avatar: config.avatar.character,
        avatarStyle: config.avatar.style,
        voice: config.avatar.voice,
        apiKeyRequired: Boolean(process.env.APP_API_KEY),
        timestamp: new Date().toISOString()
    });
});

// Get ACS token
app.get('/acs-token', requireApiKey, async (req, res) => {
    try {
        const result = await getACSToken();
        res.json(result);
    } catch (error) {
        console.error('ACS token error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Get speech token
app.get('/speech-token', requireApiKey, async (req, res) => {
    try {
        const result = await getSpeechToken();
        res.json(result);
    } catch (error) {
        console.error('Speech token error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.get('/avatar-ice', requireApiKey, async (req, res) => {
    try {
        const result = await getAvatarIceConfig();
        res.json(result);
    } catch (error) {
        console.error('Avatar ICE error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Process audio from Teams
app.post('/process-audio', requireApiKey, upload.single('audio'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No audio file' });
        }

        const audioBuffer = fs.readFileSync(req.file.path);
        const result = await processQuestion(audioBuffer);
        return res.json({ success: true, result });
    } catch (error) {
        console.error('Process audio error:', error);
        return res.status(500).json({ error: error.message });
    } finally {
        cleanupUploadedFile(req.file?.path);
    }
});

// Direct text question (for testing)
app.post('/ask', requireApiKey, async (req, res) => {
    try {
        const question = req.body?.question?.trim();
        if (!question) {
            return res.status(400).json({ error: 'No question provided' });
        }

        const answer = await getAgentAnswer(question);
        return res.json({ question, answer });
    } catch (error) {
        console.error('Ask error:', error);
        return res.status(500).json({ error: error.message });
    }
});

// Transcribe only (for testing)
app.post('/transcribe', requireApiKey, upload.single('audio'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: 'No audio file' });
        }

        const audioBuffer = fs.readFileSync(req.file.path);
        const text = await transcribeAudio(audioBuffer);
        return res.json({ text });
    } catch (error) {
        console.error('Transcribe error:', error);
        return res.status(500).json({ error: error.message });
    } finally {
        cleanupUploadedFile(req.file?.path);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`\nServer running: http://localhost:${PORT}`);
    console.log(`Wake word: "${config.avatar.wakeWord}"`);
    console.log(`Avatar: ${config.avatar.character} (${config.avatar.style})`);
    console.log(`Voice: ${config.avatar.voice}\n`);
});
