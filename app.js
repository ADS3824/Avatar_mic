require('dotenv').config();
const { CommunicationIdentityClient } = require('@azure/communication-identity');
const { AIProjectClient } = require('@azure/ai-projects');
const { ClientSecretCredential } = require('@azure/identity');
const { OpenAI } = require('openai');
const fs = require('fs');
const path = require('path');
const sdk = require('microsoft-cognitiveservices-speech-sdk');

// Config
const config = {
    openai: {
        key: process.env.AZURE_OPENAI_KEY,
        endpoint: process.env.AZURE_OPENAI_ENDPOINT,
        deployment: process.env.AZURE_OPENAI_DEPLOYMENT,
        whisper: process.env.WHISPER_DEPLOYMENT
    },
    speech: {
        key: process.env.AZURE_SPEECH_KEY,
        region: process.env.AZURE_SPEECH_REGION,
        endpoint: process.env.AZURE_SPEECH_ENDPOINT
    },
    acs: {
        connectionString: process.env.ACS_CONNECTION_STRING
    },
    agent: {
        id: process.env.AZURE_AGENT_ID,
        endpoint: process.env.AZURE_AI_PROJECT_ENDPOINT
    },
    avatar: {
        character: process.env.AVATAR_CHARACTER || 'jeff',
        style: process.env.AVATAR_STYLE || 'business',
        voice: process.env.VOICE_NAME || 'en-IN-ArjunNeural',
        wakeWord: process.env.WAKE_WORD || 'Aravindan Sir'
    },
    azure: {
        tenantId: process.env.AZURE_TENANT_ID,
        clientId: process.env.AZURE_CLIENT_ID,
        clientSecret: process.env.AZURE_CLIENT_SECRET
    }
};

// OpenAI Client (Foundry)
const openaiBaseUrl = config.openai.endpoint
    ? config.openai.endpoint.replace(/\/+$/, '')   // Foundry endpoint, no extra /openai
    : undefined;

const openaiClient = new OpenAI({
    apiKey: config.openai.key,
    baseURL: openaiBaseUrl,
    defaultHeaders: { 'api-key': config.openai.key },
    defaultQuery: { 'api-version': '2024-02-01' }
});

// Azure AI Agent Client
let agentClient;
try {
    const credential = new ClientSecretCredential(
        config.azure.tenantId,
        config.azure.clientId,
        config.azure.clientSecret
    );
    agentClient = new AIProjectClient(config.agent.endpoint, credential);
    console.log('✅ Agent client initialized');
} catch (err) {
    console.error('❌ Agent client error:', err.message);
}

// ─── STEP 1: GET ACS TOKEN ───
async function getACSToken() {
    const identityClient = new CommunicationIdentityClient(config.acs.connectionString);
    const user = await identityClient.createUser();
    const tokenResponse = await identityClient.getToken(user, ['voip']);
    console.log('✅ ACS Token obtained');
    return { user, token: tokenResponse.token, userId: user.communicationUserId };
}

// ─── STEP 2: GET SPEECH TOKEN ───
async function getSpeechToken() {
    const response = await fetch(
        `https://${config.speech.region}.api.cognitive.microsoft.com/sts/v1.0/issueToken`,
        {
            method: 'POST',
            headers: {
                'Ocp-Apim-Subscription-Key': config.speech.key,
                'Content-Type': 'application/x-www-form-urlencoded'
            }
        }
    );

    if (!response.ok) {
        const details = await response.text();
        throw new Error(`Speech token request failed (${response.status}): ${details}`);
    }

    const token = await response.text();
    return {
        token,
        region: config.speech.region,
        voice: config.avatar.voice,
        avatarCharacter: config.avatar.character,
        avatarStyle: config.avatar.style
    };
}

function buildAvatarRelayUrls(urls) {
    const relayMode = (process.env.AVATAR_RELAY_MODE || 'tcp443').trim().toLowerCase();
    const originalUrls = (Array.isArray(urls) ? urls : [urls].filter(Boolean))
        .filter(url => typeof url === 'string' && /^turns?:/i.test(url));
    const customRelayUrls = (process.env.AVATAR_TURN_URLS || '')
        .split(',')
        .map(url => url.trim())
        .filter(url => /^turns?:/i.test(url));
    const combined = [...customRelayUrls, ...originalUrls];
    const deduped = [];
    const seen = new Set();

    const pushUrl = url => {
        if (!url || seen.has(url)) {
            return;
        }
        seen.add(url);
        deduped.push(url);
    };

    combined.forEach(pushUrl);

    for (const url of [...deduped]) {
        const match = /^turns?:([^?]+)(\?.*)?$/i.exec(url);
        if (!match) {
            continue;
        }

        const hostPort = match[1];
        const query = match[2] || '';
        const host = hostPort.split(':')[0];
        if (!host) {
            continue;
        }

        pushUrl(`turns:${host}:443?transport=tcp`);
        if (relayMode === 'tcp') {
            pushUrl(`turn:${host}:3478?transport=tcp`);
        }

        if (relayMode === 'all' && !/transport=/i.test(query)) {
            pushUrl(`turn:${host}:3478?transport=udp`);
        }
    }

    const priority = url => {
        if (/^turns:.*transport=tcp/i.test(url)) return 0;
        if (/^turn:.*transport=tcp/i.test(url)) return 1;
        if (/^turns:/i.test(url)) return 2;
        if (/^turn:/i.test(url)) return 3;
        return 4;
    };

    const sorted = deduped.sort((left, right) => priority(left) - priority(right));
    const secureTcpRelayOnly = sorted.filter(url => /^turns:.*:443\?transport=tcp/i.test(url));

    return secureTcpRelayOnly.length > 0 ? secureTcpRelayOnly : sorted;
}

async function getAvatarIceConfig() {
    const response = await fetch(
        `https://${config.speech.region}.tts.speech.microsoft.com/cognitiveservices/avatar/relay/token/v1`,
        {
            method: 'GET',
            headers: {
                'Ocp-Apim-Subscription-Key': config.speech.key
            }
        }
    );

    if (!response.ok) {
        const details = await response.text();
        throw new Error(`Avatar ICE request failed (${response.status}): ${details}`);
    }

    const payload = await response.json();
    const urls = payload.urls || payload.Urls || payload.iceServers?.[0]?.urls || [];
    const username = payload.username || payload.Username || payload.iceServers?.[0]?.username;
    const credential = payload.credential || payload.Credential || payload.password || payload.Password || payload.iceServers?.[0]?.credential;

    if ((Array.isArray(urls) ? urls : [urls].filter(Boolean)).length === 0 || !username || !credential) {
        throw new Error('Avatar ICE response did not include a usable TURN configuration');
    }

    return {
        urls: buildAvatarRelayUrls(urls),
        username,
        credential
    };
}

// ─── STEP 3: TRANSCRIBE AUDIO ───
async function transcribeAudio(audioBuffer) {
    const uploadsDir = path.join(__dirname, 'uploads');
    const tempPath = path.join(uploadsDir, `temp_${Date.now()}.webm`);

    if (!fs.existsSync(uploadsDir)) {
        fs.mkdirSync(uploadsDir, { recursive: true });
    }

    fs.writeFileSync(tempPath, audioBuffer);

    try {
        const transcription = await openaiClient.audio.transcriptions.create({
            file: fs.createReadStream(tempPath),
            model: config.openai.whisper,
        });
        console.log('✅ Transcribed:', transcription.text);
        return transcription.text;
    } finally {
        if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath);
    }
}

// ─── STEP 4: CHECK WAKE WORD ───
function containsWakeWord(text) {
    const detected = text.toLowerCase().includes(config.avatar.wakeWord.toLowerCase());
    if (detected) console.log('🎯 Wake word detected!');
    return detected;
}

// ─── STEP 5: GET AGENT ANSWER ───
function extractMessageText(message) {
    if (!message || !Array.isArray(message.content)) {
        return '';
    }

    return message.content
        .map(item => {
            if (item?.type === 'text' && item.text?.value) {
                return item.text.value;
            }

            if (item?.type === 'output_text' && item.text) {
                return item.text;
            }

            return '';
        })
        .filter(Boolean)
        .join('\n')
        .trim();
}

async function getAgentAnswer(question) {
    console.log('🤔 Question:', question);

    if (!agentClient) {
        return 'I am sorry, the AI service is not available right now.';
    }

    try {
        const thread = await agentClient.agents.threads.create();
        
        await agentClient.agents.messages.create(
            thread.id,
            'user',
            question
        );

        const run = await agentClient.agents.runs.createAndPoll(
            thread.id,
            config.agent.id
        );

        if (run.status === 'completed') {
            const messages = agentClient.agents.messages.list(thread.id);
            let answer = '';

            for await (const message of messages) {
                if (message.role === 'assistant') {
                    answer = extractMessageText(message);
                    if (answer) break;
                }
            }

            if (!answer) {
                throw new Error('Assistant completed without returning any text');
            }
            console.log('✅ Answer:', answer);
            return answer;
        } else {
            console.error('❌ Run status:', run.status);
            return 'I am sorry, I could not process your question right now.';
        }
    } catch (err) {
        console.error('❌ Agent error:', err.message);
        return 'I encountered an error. Please try again.';
    }
}

// ─── STEP 6: AVATAR SPEAKS ───
async function speakWithAvatar(text, peerConnection) {
    console.log('🗣️ Speaking:', text);

    return new Promise((resolve, reject) => {
        const speechConfig = sdk.SpeechConfig.fromSubscription(
            config.speech.key,
            config.speech.region
        );
        speechConfig.speechSynthesisVoiceName = config.avatar.voice;

        const avatarConfig = new sdk.AvatarConfig(
            config.avatar.character,
            config.avatar.style,
            new sdk.AvatarVideoFormat()
        );

        const synthesizer = new sdk.AvatarSynthesizer(speechConfig, avatarConfig);

        if (peerConnection) {
            synthesizer.startAvatarAsync(peerConnection).then(() => {
                synthesizer.speakTextAsync(text, result => {
                    if (result.reason === sdk.ResultReason.SynthesizingAudioCompleted) {
                        console.log('✅ Avatar spoke successfully');
                        resolve(result);
                    } else {
                        reject(new Error(result.errorDetails));
                    }
                    synthesizer.close();
                }, err => { reject(err); synthesizer.close(); });
            });
        } else {
            synthesizer.speakTextAsync(text, result => {
                if (result.reason === sdk.ResultReason.SynthesizingAudioCompleted) {
                    console.log('✅ Avatar spoke successfully');
                    resolve(result);
                } else {
                    reject(new Error(result.errorDetails));
                }
                synthesizer.close();
            }, err => { reject(err); synthesizer.close(); });
        }
    });
}

// ─── MAIN FLOW ───
async function processQuestion(audioBuffer) {
    try {
        const text = await transcribeAudio(audioBuffer);
        if (!containsWakeWord(text)) {
            console.log('⏭️ No wake word, skipping...');
            return null;
        }
        const question = text.toLowerCase()
            .replace(config.avatar.wakeWord.toLowerCase(), '')
            .trim();
        if (!question) {
            const answer = `Please say ${config.avatar.wakeWord} followed by your question.`;
            await speakWithAvatar(answer);
            return {
                question: '',
                answer
            };
        }
        const answer = await getAgentAnswer(question);
        await speakWithAvatar(answer);
        return { question, answer };
    } catch (error) {
        console.error('❌ Flow error:', error);
        return null;
    }
}

module.exports = {
    getACSToken,
    getSpeechToken,
    getAvatarIceConfig,
    transcribeAudio,
    containsWakeWord,
    getAgentAnswer,
    speakWithAvatar,
    processQuestion,
    config
};

console.log('✅ App module loaded');
console.log(`🎯 Wake word: "${config.avatar.wakeWord}"`);
console.log(`👨 Avatar: ${config.avatar.character} (${config.avatar.style})`);
console.log(`🗣️ Voice: ${config.avatar.voice}`);
