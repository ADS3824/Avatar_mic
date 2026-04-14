import {
    CallClient,
    LocalAudioStream,
    LocalVideoStream
} from '@azure/communication-calling';
import { AzureCommunicationTokenCredential } from '@azure/communication-common';

// ─── State ───────────────────────────────────────────────────────────────────
let speechConfig, avatarSynthesizer, avatarPeerConnection;
let callAgent, call;
let isListening = false;
let isBusy = false;
let suppressListeningUntil = 0;
let audioProcessingQueue = Promise.resolve();
let currentMode = 'mic'; // 'mic' | 'teams'

// Teams audio
let meetingAudioContext, meetingAudioDestination;
const remoteAudioNodes = new Map();
let teamsMediaRecorder;
let avatarRemoteAudioStream, avatarRemoteVideoStream;
let avatarLocalAudioStream, avatarLocalVideoStream;

// Mic audio
let micMediaRecorder;
let micStream;

let runtimeConfig = { wakeWord: 'Aravindan Sir', apiKeyRequired: false };
let avatarRelayDiagnostics = {
    lastIceState: 'new', lastConnectionState: 'new',
    lastCandidateError: null, lastRelayUrls: []
};

const appApiKey = new URLSearchParams(window.location.search).get('apiKey')
    || window.localStorage.getItem('appApiKey') || '';
if (appApiKey) window.localStorage.setItem('appApiKey', appApiKey);

// ─── Helpers ─────────────────────────────────────────────────────────────────
function getSpeechSdk() {
    if (!window.SpeechSDK) throw new Error('Azure Speech SDK did not load. Refresh and try again.');
    return window.SpeechSDK;
}
function getWakeWordLC() { return runtimeConfig.wakeWord.toLowerCase(); }
function buildHeaders(base = {}) {
    return appApiKey ? { ...base, 'x-app-api-key': appApiKey } : base;
}
function setStatus(msg, type = 'info') {
    const el = document.getElementById('status');
    el.textContent = '';
    const dot = document.createElement('span');
    dot.className = `indicator ${type === 'ok' ? 'ind-on' : type === 'wait' ? 'ind-wait' : 'ind-off'}`;
    el.appendChild(dot);
    el.appendChild(document.createTextNode(msg));
}
function addConversation(question, answer) {
    const conv = document.getElementById('conversation');
    const empty = document.getElementById('conversationEmpty');
    if (empty) empty.remove();
    const q = document.createElement('p'); q.className = 'msg-q'; q.textContent = `Q: ${question}`;
    const a = document.createElement('p'); a.className = 'msg-a'; a.textContent = `A: ${answer}`;
    conv.appendChild(q); conv.appendChild(a);
    conv.scrollTop = conv.scrollHeight;
}
async function readJson(response) {
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.error || `Request failed (${response.status})`);
    return payload;
}
function getPreferredMimeType() {
    const candidates = ['audio/webm;codecs=opus', 'audio/webm', 'audio/ogg;codecs=opus'];
    return candidates.find(t => window.MediaRecorder && MediaRecorder.isTypeSupported(t));
}
function cloneStreamByKind(stream, kind) {
    const tracks = stream.getTracks().filter(t => !kind || t.kind === kind).map(t => t.clone());
    return new MediaStream(tracks);
}

// ─── Mode Switch ─────────────────────────────────────────────────────────────
function switchMode(mode) {
    if (isListening) stopAllListening();
    currentMode = mode;
    document.getElementById('tabMic').classList.toggle('active', mode === 'mic');
    document.getElementById('tabTeams').classList.toggle('active', mode === 'teams');
    document.getElementById('panelMic').classList.toggle('active', mode === 'mic');
    document.getElementById('panelTeams').classList.toggle('active', mode === 'teams');
    document.getElementById('modeStatus').textContent = mode === 'mic' ? '🎙️ Microphone' : '👥 Teams';
}

// ─── Avatar Relay ─────────────────────────────────────────────────────────────
function normalizeIceUrls(urls) {
    const list = Array.isArray(urls) ? urls : [urls].filter(Boolean);
    const turnOnly = list.filter(u => typeof u === 'string' && /^turns?:/i.test(u));
    const candidates = turnOnly.length > 0 ? turnOnly : list;
    const seen = new Set();
    const sorted = candidates.filter(u => {
        if (typeof u !== 'string' || seen.has(u)) return false;
        seen.add(u); return true;
    }).sort((a, b) => {
        const p = v => /^turns:.*transport=tcp/i.test(v) ? 0 : /^turn:.*transport=tcp/i.test(v) ? 1
            : /^turns:/i.test(v) ? 2 : /^turn:/i.test(v) ? 3 : 4;
        return p(a) - p(b);
    });
    const secure = sorted.filter(u => /^turns:.*:443\?transport=tcp/i.test(u));
    return secure.length > 0 ? secure : sorted;
}
function describeRelayIssue(prefix) {
    const e = avatarRelayDiagnostics.lastCandidateError;
    const relays = avatarRelayDiagnostics.lastRelayUrls.length > 0
        ? ` Relay: ${avatarRelayDiagnostics.lastRelayUrls.join(', ')}.` : '';
    if (e && e.errorCode) return `${prefix} ICE ${e.errorCode} ${e.errorText || ''} ${e.url || ''}.${relays}`;
    return `${prefix} WebRTC TURN relay failed. Check network/VPN/firewall.${relays}`;
}
function waitForRelay(pc, ms = 20000) {
    return new Promise((resolve, reject) => {
        const isOk = () => ['connected', 'completed'].includes(pc.iceConnectionState);
        if (isOk()) { resolve(); return; }
        let settled = false, timer;
        const cleanup = () => { clearTimeout(timer); pc.removeEventListener('iceconnectionstatechange', h); pc.removeEventListener('connectionstatechange', h); };
        const fail = m => { if (settled) return; settled = true; cleanup(); reject(new Error(describeRelayIssue(m))); };
        const succeed = () => { if (settled) return; settled = true; cleanup(); resolve(); };
        const h = () => { if (isOk()) { succeed(); return; } if (pc.iceConnectionState === 'failed' || pc.connectionState === 'failed') fail('Avatar relay failed.'); };
        timer = setTimeout(() => fail('Timed out waiting for avatar relay.'), ms);
        pc.addEventListener('iceconnectionstatechange', h);
        pc.addEventListener('connectionstatechange', h);
    });
}
function wireAvatarTracks(pc) {
    pc.addTransceiver('video', { direction: 'sendrecv' });
    pc.addTransceiver('audio', { direction: 'sendrecv' });
    pc.onconnectionstatechange = () => { avatarRelayDiagnostics.lastConnectionState = pc.connectionState; };
    pc.onicecandidateerror = ev => { avatarRelayDiagnostics.lastCandidateError = { url: ev.url || '', errorCode: ev.errorCode, errorText: ev.errorText || '' }; };
    pc.ontrack = ev => {
        if (ev.track.kind === 'video') {
            const vid = document.getElementById('avatarVideo');
            vid.srcObject = ev.streams[0]; vid.style.display = 'block';
            document.getElementById('placeholder').style.display = 'none';
            avatarRemoteVideoStream = cloneStreamByKind(ev.streams[0] || new MediaStream([ev.track]), 'video');
        }
        if (ev.track.kind === 'audio') avatarRemoteAudioStream = cloneStreamByKind(ev.streams[0] || new MediaStream([ev.track]), 'audio');
        syncAvatarMediaToTeams().catch(console.error);
    };
    let iceWarn;
    pc.oniceconnectionstatechange = () => {
        const s = pc.iceConnectionState;
        avatarRelayDiagnostics.lastIceState = s;
        if (s === 'connected' || s === 'completed') { clearTimeout(iceWarn); iceWarn = null; setStatus('Avatar relay connected.', 'ok'); }
        else if (s === 'failed') { clearTimeout(iceWarn); setStatus(describeRelayIssue('Avatar relay failed.'), 'info'); }
        else if (s === 'disconnected' && !iceWarn) {
            iceWarn = setTimeout(() => { iceWarn = null; if (pc.iceConnectionState === 'disconnected') setStatus('Avatar relay issue. Reconnect if video stops.', 'info'); }, 4000);
        }
    };
}

// ─── Avatar Connect / Disconnect ─────────────────────────────────────────────
async function connectAvatar() {
    setStatus('Connecting avatar relay...', 'wait');
    document.getElementById('connectAvatarBtn').disabled = true;
    try {
        const SpeechSDK = getSpeechSdk();
        const [speechToken, iceConfig] = await Promise.all([
            readJson(await fetch('/speech-token', { headers: buildHeaders() })),
            readJson(await fetch('/avatar-ice', { headers: buildHeaders() }))
        ]);
        const { token, region, voice, avatarCharacter, avatarStyle } = speechToken;
        const relayUrls = normalizeIceUrls(iceConfig.urls);
        if (!relayUrls.length) throw new Error('No TURN relay URLs returned.');
        avatarRelayDiagnostics = { lastIceState: 'new', lastConnectionState: 'new', lastCandidateError: null, lastRelayUrls: relayUrls.slice() };

        speechConfig = SpeechSDK.SpeechConfig.fromAuthorizationToken(token, region);
        speechConfig.speechSynthesisVoiceName = voice;
        const avatarConfig = new SpeechSDK.AvatarConfig(avatarCharacter, avatarStyle, new SpeechSDK.AvatarVideoFormat());
        avatarPeerConnection = new RTCPeerConnection({ iceServers: [{ urls: relayUrls, username: iceConfig.username, credential: iceConfig.credential }], iceTransportPolicy: 'relay' });
        wireAvatarTracks(avatarPeerConnection);
        avatarSynthesizer = new SpeechSDK.AvatarSynthesizer(speechConfig, avatarConfig);
        await avatarSynthesizer.startAvatarAsync(avatarPeerConnection);
        await waitForRelay(avatarPeerConnection);

        document.getElementById('avatarStatus').textContent = 'Connected';
        document.getElementById('avatarStatus').style.color = '#2ecc71';
        document.getElementById('disconnectAvatarBtn').disabled = false;
        document.getElementById('joinBtn').disabled = false;
        document.getElementById('micListenBtn').disabled = false;
        document.getElementById('speakTextBtn').disabled = false;
        setStatus('Avatar connected. Choose input mode and start listening.', 'ok');
        await avatarSynthesizer.speakTextAsync(
            'Hello. I am Aravindan Lite. Say ' + runtimeConfig.wakeWord + ' followed by your question and I will answer.'
        );
    } catch (err) {
        setStatus('Avatar connection failed: ' + err.message, 'info');
        if (avatarSynthesizer) { avatarSynthesizer.close(); avatarSynthesizer = null; }
        if (avatarPeerConnection) { avatarPeerConnection.close(); avatarPeerConnection = null; }
        document.getElementById('avatarStatus').textContent = 'Disconnected';
        document.getElementById('avatarStatus').style.color = '#e74c3c';
        document.getElementById('disconnectAvatarBtn').disabled = true;
        document.getElementById('joinBtn').disabled = true;
        document.getElementById('micListenBtn').disabled = true;
        document.getElementById('speakTextBtn').disabled = true;
        document.getElementById('connectAvatarBtn').disabled = false;
        console.error(err);
    }
}

async function disconnectAvatar() {
    stopAllListening();
    if (call) await leaveTeamsMeeting();
    else await stopSendingAvatarMediaToTeams();
    if (avatarSynthesizer) { avatarSynthesizer.close(); avatarSynthesizer = null; }
    if (avatarPeerConnection) { avatarPeerConnection.close(); avatarPeerConnection = null; }
    avatarRemoteAudioStream = null; avatarRemoteVideoStream = null;
    document.getElementById('avatarVideo').style.display = 'none';
    document.getElementById('placeholder').style.display = 'block';
    document.getElementById('avatarStatus').textContent = 'Disconnected';
    document.getElementById('avatarStatus').style.color = '#e74c3c';
    document.getElementById('connectAvatarBtn').disabled = false;
    document.getElementById('disconnectAvatarBtn').disabled = true;
    document.getElementById('joinBtn').disabled = true;
    document.getElementById('micListenBtn').disabled = true;
    document.getElementById('speakTextBtn').disabled = true;
    setStatus('Avatar disconnected.');
}

// ─── Microphone Listening ─────────────────────────────────────────────────────
async function toggleMicListening() {
    if (!isListening || currentMode !== 'mic') await startMicListening();
    else stopMicListening();
}

async function startMicListening() {
    if (!avatarSynthesizer) { setStatus('Connect avatar first.', 'info'); return; }
    try {
        micStream = await navigator.mediaDevices.getUserMedia({ audio: true, video: false });
        const mimeType = getPreferredMimeType();
        micMediaRecorder = mimeType
            ? new MediaRecorder(micStream, { mimeType })
            : new MediaRecorder(micStream);
        micMediaRecorder.ondataavailable = ev => {
            if (!ev.data || ev.data.size === 0) return;
            audioProcessingQueue = audioProcessingQueue
                .then(() => processAudioChunk(ev.data))
                .catch(e => console.error('Audio processing error:', e));
        };
        micMediaRecorder.start(5000);
        isListening = true;
        document.getElementById('micListenBtn').textContent = '🔴 Stop Listening';
        document.getElementById('micListenBtn').classList.add('recording');
        document.getElementById('listenStatus').textContent = 'On (Mic)';
        document.getElementById('listenStatus').style.color = '#2ecc71';
        setStatus('Listening via microphone for "' + runtimeConfig.wakeWord + '"...', 'ok');
    } catch (err) {
        setStatus('Microphone access failed: ' + err.message, 'info');
        console.error(err);
    }
}

function stopMicListening() {
    if (micMediaRecorder && micMediaRecorder.state !== 'inactive') micMediaRecorder.stop();
    micMediaRecorder = null;
    if (micStream) { micStream.getTracks().forEach(t => t.stop()); micStream = null; }
    isListening = false;
    document.getElementById('micListenBtn').textContent = '🎙️ Start Listening';
    document.getElementById('micListenBtn').classList.remove('recording');
    document.getElementById('listenStatus').textContent = 'Off';
    document.getElementById('listenStatus').style.color = '#e74c3c';
    setStatus('Microphone listening stopped.');
}

// ─── Teams Listening ──────────────────────────────────────────────────────────
async function toggleTeamsListening() {
    if (!isListening || currentMode !== 'teams') await startTeamsListening();
    else stopTeamsListening();
}

async function startTeamsListening() {
    if (!call || call.state !== 'Connected') { setStatus('Join the Teams meeting first.', 'info'); return; }
    try {
        await ensureMeetingAudioContext();
        const mimeType = getPreferredMimeType();
        teamsMediaRecorder = mimeType
            ? new MediaRecorder(meetingAudioDestination.stream, { mimeType })
            : new MediaRecorder(meetingAudioDestination.stream);
        teamsMediaRecorder.ondataavailable = ev => {
            if (!ev.data || ev.data.size === 0) return;
            audioProcessingQueue = audioProcessingQueue
                .then(() => processAudioChunk(ev.data))
                .catch(e => console.error('Audio processing error:', e));
        };
        teamsMediaRecorder.start(5000);
        isListening = true;
        document.getElementById('listenBtn').textContent = '🔴 Stop Listening';
        document.getElementById('listenBtn').classList.add('recording');
        document.getElementById('listenStatus').textContent = 'On (Teams)';
        document.getElementById('listenStatus').style.color = '#2ecc71';
        setStatus('Listening to Teams meeting for "' + runtimeConfig.wakeWord + '"...', 'ok');
    } catch (err) {
        setStatus('Teams listening failed: ' + err.message, 'info');
        console.error(err);
    }
}

function stopTeamsListening() {
    if (teamsMediaRecorder && teamsMediaRecorder.state !== 'inactive') teamsMediaRecorder.stop();
    teamsMediaRecorder = null;
    isListening = false;
    document.getElementById('listenBtn').textContent = '👂 Start Listening';
    document.getElementById('listenBtn').classList.remove('recording');
    document.getElementById('listenStatus').textContent = 'Off';
    document.getElementById('listenStatus').style.color = '#e74c3c';
    setStatus('Teams listening stopped.');
}

function stopAllListening() {
    stopMicListening();
    stopTeamsListening();
}

// ─── Audio Processing (shared by mic & teams) ─────────────────────────────────
async function processAudioChunk(blob) {
    if (isBusy || Date.now() < suppressListeningUntil) return;
    const formData = new FormData();
    const ext = blob.type.includes('ogg') ? 'ogg' : 'webm';
    formData.append('audio', blob, `chunk.${ext}`);
    const result = await readJson(await fetch('/transcribe', {
        method: 'POST', headers: buildHeaders(), body: formData
    }));
    const text = result.text;
    if (!text) return;
    const normalized = text.toLowerCase();
    const wakeWord = getWakeWordLC();
    if (!normalized.includes(wakeWord)) return;
    const question = normalized.replace(wakeWord, '').trim();
    if (!question) { setStatus('Wake word heard. Waiting for question...', 'wait'); return; }
    setStatus('Question detected: ' + question, 'wait');
    await askAndSpeak(question);
}

// ─── Ask & Speak ─────────────────────────────────────────────────────────────
async function askAndSpeak(question, speakAloud = true) {
    if (isBusy) { setStatus('Still processing previous answer...', 'wait'); return; }
    isBusy = true;
    try {
        setStatus('Getting answer from knowledge base...', 'wait');
        const result = await readJson(await fetch('/ask', {
            method: 'POST',
            headers: buildHeaders({ 'Content-Type': 'application/json' }),
            body: JSON.stringify({ question })
        }));
        addConversation(question, result.answer);
        if (speakAloud && avatarSynthesizer) {
            await syncAvatarMediaToTeams();
            setStatus('Avatar is speaking...', 'wait');
            await avatarSynthesizer.speakTextAsync(result.answer);
            suppressListeningUntil = Date.now() + 4000;
        }
        setStatus('Ready.', 'ok');
    } catch (err) {
        setStatus('Error: ' + err.message, 'info');
        console.error(err);
    } finally {
        isBusy = false;
    }
}

// Ask Directly — text only (no speech)
async function testQuestionTextOnly() {
    const question = document.getElementById('testQuestion').value.trim();
    if (!question) return;
    document.getElementById('testQuestion').value = '';
    await askAndSpeak(question, false);
}

// Ask Directly — answer + avatar speaks
async function testQuestionAndSpeak() {
    const question = document.getElementById('testQuestion').value.trim();
    if (!question) return;
    if (!avatarSynthesizer) { alert('Connect the avatar first.'); return; }
    document.getElementById('testQuestion').value = '';
    await askAndSpeak(question, true);
}

// ─── Speak Pasted Text ────────────────────────────────────────────────────────
async function speakPastedText() {
    const text = document.getElementById('speakTextArea').value.trim();
    if (!text) { setStatus('Please paste some text to speak.', 'info'); return; }
    if (!avatarSynthesizer) { alert('Connect the avatar first.'); return; }
    if (isBusy) { setStatus('Avatar is busy. Please wait.', 'wait'); return; }
    isBusy = true;
    document.getElementById('speakTextBtn').disabled = true;
    try {
        await syncAvatarMediaToTeams();
        setStatus('Avatar is speaking pasted text...', 'wait');
        await avatarSynthesizer.speakTextAsync(text);
        suppressListeningUntil = Date.now() + 2000;
        setStatus('Done speaking.', 'ok');
    } catch (err) {
        setStatus('Speak error: ' + err.message, 'info');
        console.error(err);
    } finally {
        isBusy = false;
        document.getElementById('speakTextBtn').disabled = false;
    }
}

// ─── Teams Meeting ────────────────────────────────────────────────────────────
async function ensureMeetingAudioContext() {
    if (!meetingAudioContext) {
        const AC = window.AudioContext || window.webkitAudioContext;
        if (!AC) throw new Error('Web Audio API not supported.');
        meetingAudioContext = new AC();
        meetingAudioDestination = meetingAudioContext.createMediaStreamDestination();
    }
    if (meetingAudioContext.state === 'suspended') await meetingAudioContext.resume();
}
function clearRemoteAudioNodes() {
    remoteAudioNodes.forEach(n => n.sourceNode.disconnect());
    remoteAudioNodes.clear();
}
async function attachRemoteAudioStream(s) {
    if (remoteAudioNodes.has(s)) return;
    await ensureMeetingAudioContext();
    const ms = await s.getMediaStream();
    const src = meetingAudioContext.createMediaStreamSource(ms);
    src.connect(meetingAudioDestination);
    remoteAudioNodes.set(s, { sourceNode: src });
}
function detachRemoteAudioStream(s) {
    const n = remoteAudioNodes.get(s);
    if (!n) return;
    n.sourceNode.disconnect();
    remoteAudioNodes.delete(s);
}
async function subscribeToMeetingAudio(activeCall) {
    await ensureMeetingAudioContext();
    activeCall.remoteAudioStreams.forEach(s => attachRemoteAudioStream(s).catch(console.error));
    activeCall.on('remoteAudioStreamsUpdated', ev => {
        ev.added.forEach(s => attachRemoteAudioStream(s).catch(console.error));
        ev.removed.forEach(detachRemoteAudioStream);
    });
}
async function stopSendingAvatarMediaToTeams() {
    if (!call) { avatarLocalAudioStream = null; avatarLocalVideoStream = null; return; }
    if (avatarLocalVideoStream) { await call.stopVideo(avatarLocalVideoStream).catch(console.error); avatarLocalVideoStream = null; }
    if (avatarLocalAudioStream) { await call.stopAudio().catch(console.error); avatarLocalAudioStream = null; }
}
async function syncAvatarMediaToTeams() {
    if (!call || call.state !== 'Connected') return;
    if (avatarRemoteAudioStream) {
        const ams = cloneStreamByKind(avatarRemoteAudioStream, 'audio');
        if (ams.getAudioTracks().length > 0) {
            if (!avatarLocalAudioStream) { avatarLocalAudioStream = new LocalAudioStream(ams); await call.startAudio(avatarLocalAudioStream); }
            else if (typeof avatarLocalAudioStream.setMediaStream === 'function') await avatarLocalAudioStream.setMediaStream(ams);
            if (call.isMuted) await call.unmute().catch(console.error);
        }
    }
    if (avatarRemoteVideoStream) {
        const vms = cloneStreamByKind(avatarRemoteVideoStream, 'video');
        if (vms.getVideoTracks().length > 0) {
            if (!avatarLocalVideoStream) { avatarLocalVideoStream = new LocalVideoStream(vms); await call.startVideo(avatarLocalVideoStream); }
            else if (typeof avatarLocalVideoStream.setMediaStream === 'function') await avatarLocalVideoStream.setMediaStream(vms);
        }
    }
}

async function joinTeamsMeeting() {
    const link = document.getElementById('meetingLink').value.trim();
    if (!link) { alert('Please paste a Teams meeting link.'); return; }
    if (!avatarSynthesizer) { alert('Connect the avatar first.'); return; }
    setStatus('Joining Teams meeting...', 'wait');
    document.getElementById('joinBtn').disabled = true;
    try {
        await ensureMeetingAudioContext();
        const { token } = await readJson(await fetch('/acs-token', { headers: buildHeaders() }));
        const callClient = new CallClient();
        const cred = new AzureCommunicationTokenCredential(token);
        callAgent = await callClient.createCallAgent(cred, { displayName: 'Aravindan Lite' });
        call = callAgent.join({ meetingLink: link }, { audioOptions: { muted: true }, videoOptions: { localVideoStreams: [] } });
        call.on('stateChanged', async () => {
            if (call.state === 'Connected') {
                document.getElementById('teamsStatus').textContent = 'Joined';
                document.getElementById('teamsStatus').style.color = '#2ecc71';
                document.getElementById('leaveBtn').disabled = false;
                document.getElementById('listenBtn').disabled = false;
                await subscribeToMeetingAudio(call);
                await syncAvatarMediaToTeams();
                setStatus('Joined Teams. Start listening.', 'ok');
            } else if (call.state === 'InLobby') {
                setStatus('Waiting in Teams lobby...', 'wait');
            } else if (call.state === 'Disconnected') {
                clearRemoteAudioNodes();
                document.getElementById('teamsStatus').textContent = 'Not joined';
                document.getElementById('teamsStatus').style.color = '#e74c3c';
                document.getElementById('leaveBtn').disabled = true;
                document.getElementById('listenBtn').disabled = true;
                document.getElementById('joinBtn').disabled = !avatarSynthesizer;
                setStatus('Disconnected from Teams.');
            }
        });
    } catch (err) {
        setStatus('Failed to join: ' + err.message, 'info');
        document.getElementById('joinBtn').disabled = false;
        console.error(err);
    }
}

async function leaveTeamsMeeting() {
    stopTeamsListening();
    await stopSendingAvatarMediaToTeams();
    clearRemoteAudioNodes();
    if (call) { await call.hangUp().catch(console.error); call = null; }
    if (callAgent && typeof callAgent.dispose === 'function') callAgent.dispose();
    callAgent = null;
    document.getElementById('teamsStatus').textContent = 'Not joined';
    document.getElementById('teamsStatus').style.color = '#e74c3c';
    document.getElementById('leaveBtn').disabled = true;
    document.getElementById('listenBtn').disabled = true;
    document.getElementById('joinBtn').disabled = !avatarSynthesizer;
    setStatus('Left Teams meeting.');
}

// ─── Init ─────────────────────────────────────────────────────────────────────
async function loadRuntimeConfig() {
    try {
        const health = await readJson(await fetch('/health', { headers: buildHeaders() }));
        runtimeConfig = { wakeWord: health.wakeWord || runtimeConfig.wakeWord, apiKeyRequired: Boolean(health.apiKeyRequired) };
        document.getElementById('wakeWordHintMic').textContent = runtimeConfig.wakeWord;
        document.getElementById('wakeWordHintTeams').textContent = runtimeConfig.wakeWord;
        if (runtimeConfig.apiKeyRequired && !appApiKey) {
            document.getElementById('connectAvatarBtn').disabled = true;
            setStatus('API key required. Open with ?apiKey=YOUR_KEY', 'info');
            return;
        }
    } catch (err) { console.error('Health check failed:', err); }
    setStatus('Ready — connect the avatar to begin.');
    document.getElementById('modeStatus').textContent = '🎙️ Microphone';
}

// ─── Expose to HTML ───────────────────────────────────────────────────────────
window.connectAvatar = connectAvatar;
window.disconnectAvatar = disconnectAvatar;
window.joinTeamsMeeting = joinTeamsMeeting;
window.leaveTeamsMeeting = leaveTeamsMeeting;
window.toggleMicListening = toggleMicListening;
window.toggleTeamsListening = toggleTeamsListening;
window.switchMode = switchMode;
window.testQuestionTextOnly = testQuestionTextOnly;
window.testQuestionAndSpeak = testQuestionAndSpeak;
window.speakPastedText = speakPastedText;

document.getElementById('testQuestion').addEventListener('keypress', ev => {
    if (ev.key === 'Enter') testQuestionAndSpeak();
});

loadRuntimeConfig();
