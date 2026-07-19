// ─────────────────────────────────────────────────────────────
//  NEXUS // TRACE — Project Report Generator
//  Generates a structured Word (.docx) report for submission
// ─────────────────────────────────────────────────────────────

const {
  Document, Packer, Paragraph, Table, TableRow, TableCell,
  TextRun, HeadingLevel, AlignmentType, BorderStyle,
  WidthType, ShadingType, TableLayoutType, VerticalAlign,
  Header, Footer, PageNumberElement, NumberFormat, Tab,
  UnderlineType, PageBreak, LevelFormat,
} = require("docx");
const fs = require("fs");
const path = require("path");

// ── Color Palette ─────────────────────────────────────────────
const COLOR = {
  primary:    "1A3C5E",   // Dark navy blue
  accent:     "2E86C1",   // Professional blue
  light:      "D6EAF8",   // Light blue fill
  heading:    "1A3C5E",
  subheading: "2E86C1",
  black:      "000000",
  white:      "FFFFFF",
  gray:       "5D6D7E",
  lightGray:  "F2F3F4",
  border:     "AEB6BF",
  green:      "1E8449",
  amber:      "D68910",
};

// ── Helpers ───────────────────────────────────────────────────
function bold(text, size = 22, color = COLOR.black) {
  return new TextRun({ text, bold: true, size, color, font: "Calibri" });
}

function normal(text, size = 22, color = COLOR.black) {
  return new TextRun({ text, size, color, font: "Calibri" });
}

function italic(text, size = 20, color = COLOR.gray) {
  return new TextRun({ text, italics: true, size, color, font: "Calibri" });
}

function h1(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 36, color: COLOR.white, font: "Calibri" })],
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    shading: { type: ShadingType.SOLID, color: COLOR.primary, fill: COLOR.primary },
    spacing: { before: 200, after: 200 },
  });
}

function h2(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 28, color: COLOR.white, font: "Calibri" })],
    heading: HeadingLevel.HEADING_2,
    shading: { type: ShadingType.SOLID, color: COLOR.accent, fill: COLOR.accent },
    spacing: { before: 280, after: 120 },
    indent: { left: 0 },
  });
}

function h3(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 24, color: COLOR.primary, font: "Calibri",
      underline: { type: UnderlineType.SINGLE, color: COLOR.accent } })],
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 80 },
  });
}

function para(text, size = 22, align = AlignmentType.JUSTIFIED) {
  return new Paragraph({
    children: [normal(text, size)],
    alignment: align,
    spacing: { before: 60, after: 80, line: 336 },
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    children: [normal(text, 20)],
    bullet: { level },
    spacing: { before: 40, after: 40, line: 300 },
    indent: { left: 360 + level * 360 },
  });
}

function keyValue(key, value) {
  return new Paragraph({
    children: [
      bold(key + ": ", 20, COLOR.primary),
      normal(value, 20),
    ],
    spacing: { before: 40, after: 40, line: 300 },
  });
}

function spacer(lines = 1) {
  return Array.from({ length: lines }, () =>
    new Paragraph({ children: [new TextRun("")], spacing: { before: 80, after: 80 } })
  );
}

function divider() {
  return new Paragraph({
    children: [new TextRun("")],
    border: { bottom: { color: COLOR.accent, space: 1, size: 6, style: BorderStyle.SINGLE } },
    spacing: { before: 120, after: 120 },
  });
}

function infoBox(label, text, bgColor = COLOR.light) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            children: [
              new Paragraph({ children: [bold(label, 20, COLOR.primary)], spacing: { before: 40, after: 20 } }),
              new Paragraph({ children: [normal(text, 20)], spacing: { before: 20, after: 40 } }),
            ],
            shading: { type: ShadingType.SOLID, color: bgColor, fill: bgColor },
            margins: { top: 100, bottom: 100, left: 200, right: 200 },
            borders: {
              top:    { style: BorderStyle.SINGLE, size: 4, color: COLOR.accent },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: COLOR.accent },
              left:   { style: BorderStyle.SINGLE, size: 12, color: COLOR.accent },
              right:  { style: BorderStyle.NONE },
            },
          }),
        ],
      }),
    ],
    margins: { top: 80, bottom: 80 },
  });
}

function twoColTable(rows, headers = []) {
  const headerRow = headers.length
    ? [new TableRow({
        tableHeader: true,
        children: headers.map(h => new TableCell({
          children: [new Paragraph({ children: [bold(h, 20, COLOR.white)], alignment: AlignmentType.CENTER })],
          shading: { type: ShadingType.SOLID, color: COLOR.primary, fill: COLOR.primary },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          verticalAlign: VerticalAlign.CENTER,
        })),
      })]
    : [];

  const dataRows = rows.map((row, i) =>
    new TableRow({
      children: row.map(cell => new TableCell({
        children: [new Paragraph({ children: [normal(cell, 20)], alignment: AlignmentType.LEFT })],
        shading: i % 2 === 0
          ? { type: ShadingType.SOLID, color: COLOR.lightGray, fill: COLOR.lightGray }
          : { type: ShadingType.CLEAR, color: COLOR.white, fill: COLOR.white },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        borders: {
          top:    { style: BorderStyle.SINGLE, size: 2, color: COLOR.border },
          bottom: { style: BorderStyle.SINGLE, size: 2, color: COLOR.border },
          left:   { style: BorderStyle.SINGLE, size: 2, color: COLOR.border },
          right:  { style: BorderStyle.SINGLE, size: 2, color: COLOR.border },
        },
      })),
    })
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [...headerRow, ...dataRows],
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ── Cover Page ────────────────────────────────────────────────
function buildCoverPage() {
  return [
    ...spacer(2),
    new Paragraph({
      children: [new TextRun({ text: "⬡", size: 120, color: COLOR.accent, font: "Segoe UI Symbol" })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 200, after: 100 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "NEXUS TRACE", bold: true, size: 64, color: COLOR.primary, font: "Calibri" })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 60 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Information Gathering & Device Fingerprinting System", italics: true, size: 28, color: COLOR.accent, font: "Calibri" })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: 300 },
    }),
    divider(),
    ...spacer(1),
    new Paragraph({
      children: [new TextRun({ text: "PROJECT REPORT", bold: true, size: 36, color: COLOR.primary, font: "Calibri" })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 100, after: 60 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Submitted in partial fulfillment of the course requirements", size: 22, color: COLOR.gray, font: "Calibri" })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 40, after: 200 },
    }),
    ...spacer(2),
    twoColTable([
      ["Submitted By",   "Praveen Narni"],
      ["Guided By",      "Sai Shanthan"],
      ["Project Title",  "NEXUS TRACE — Information Gathering Application"],
      ["Domain",         "Web Development / Cybersecurity Tools"],
      ["Technology",     "Node.js, Express.js, SQL.js (SQLite), HTML5, CSS3, Vanilla JS"],
      ["Date",           new Date().toLocaleDateString("en-IN", { dateStyle: "long" })],
    ], ["Field", "Details"]),
    ...spacer(3),
    pageBreak(),
  ];
}

// ── Table of Contents ─────────────────────────────────────────
function buildToC() {
  return [
    h1("TABLE OF CONTENTS"),
    ...spacer(1),
    ...[
      ["1.", "Introduction",                             "3"],
      ["2.", "Problem Statement",                        "3"],
      ["3.", "Objectives",                               "4"],
      ["4.", "Technology Stack",                         "4"],
      ["5.", "System Architecture",                      "5"],
      ["6.", "Database Design",                          "6"],
      ["7.", "Module-wise Description",                  "7"],
      ["8.", "API Documentation",                        "8"],
      ["9.", "Application Workflow",                     "9"],
      ["10.", "Screenshots / UI Description",            "10"],
      ["11.", "Security Considerations",                 "11"],
      ["12.", "Limitations & Future Enhancements",       "11"],
      ["13.", "Conclusion",                              "12"],
    ].map(([num, title, page]) =>
      new Paragraph({
        children: [
          bold(num + " ", 22, COLOR.primary),
          normal(title, 22),
          new TextRun({ text: `  ${"·".repeat(60 - title.length - num.length)}  ${page}`, size: 22, color: COLOR.gray, font: "Courier New" }),
        ],
        spacing: { before: 80, after: 80 },
      })
    ),
    pageBreak(),
  ];
}

// ── Section 1: Introduction ───────────────────────────────────
function buildIntro() {
  return [
    h1("1. INTRODUCTION"),
    para(
      "NEXUS TRACE is a web-based information gathering and device fingerprinting application developed as an educational cybersecurity demonstration tool. " +
      "The project simulates how malicious actors can collect sensitive information about a user — such as their device details, GPS location, camera images, and audio recordings — " +
      "by disguising a data collection page as a legitimate web form (in this case, a fake job application portal)."
    ),
    para(
      "This application was built to raise awareness about the kinds of data that modern web browsers expose to websites without the user realizing it. " +
      "It also serves as a practical demonstration of several browser APIs including the Geolocation API, MediaDevices API, and MediaRecorder API, combined with a " +
      "Node.js backend and a SQLite database for persistent data storage."
    ),
    ...spacer(1),
    infoBox(
      "📌 Important Note",
      "This project is built strictly for educational purposes — to understand how information gathering works on the web and to learn about browser security. " +
      "Deploying such a tool against unsuspecting users without consent is illegal and unethical."
    ),
    ...spacer(1),
    pageBreak(),
  ];
}

// ── Section 2: Problem Statement ─────────────────────────────
function buildProblemStatement() {
  return [
    h1("2. PROBLEM STATEMENT"),
    para(
      "In today's digital world, most internet users are unaware of how much personal data their browser shares with every website they visit. " +
      "A simple web page can, with the right code and user permission, access:"
    ),
    bullet("The operating system, browser, and device type of the visitor"),
    bullet("The exact GPS coordinates of the visitor"),
    bullet("Live camera images from both front and rear cameras"),
    bullet("Audio recordings from the microphone"),
    bullet("The visitor's IP address, city, ISP, and approximate location via IP geolocation"),
    ...spacer(1),
    para(
      "The challenge this project addresses is: How can all of this information be collected, processed, and stored efficiently through a single web page, " +
      "and how can it be presented in a clear, organized dashboard for review?"
    ),
    pageBreak(),
  ];
}

// ── Section 3: Objectives ─────────────────────────────────────
function buildObjectives() {
  return [
    h1("3. OBJECTIVES"),
    para("The primary objectives of the NEXUS TRACE project are:"),
    bullet("To build a full-stack web application using Node.js and Express.js"),
    bullet("To demonstrate real-world use of browser APIs: Geolocation, Camera (getUserMedia), and Audio (MediaRecorder)"),
    bullet("To implement a persistent database using SQLite (via sql.js) to store captured visitor records"),
    bullet("To design a visually engaging and functional frontend using HTML5, CSS3, and Vanilla JavaScript"),
    bullet("To create a secure admin panel that allows authorized users to view all collected records"),
    bullet("To build a \"social engineering\" styled bait page to understand how phishing/data collection attacks are structured"),
    bullet("To learn how server-side IP geolocation APIs work and how to integrate them"),
    pageBreak(),
  ];
}

// ── Section 4: Technology Stack ───────────────────────────────
function buildTechStack() {
  return [
    h1("4. TECHNOLOGY STACK"),
    para("The project is divided into a backend (server-side) and a frontend (client-side). Below is a detailed breakdown of all the technologies used:"),
    ...spacer(1),
    h3("4.1 Backend Technologies"),
    twoColTable([
      ["Node.js",        "JavaScript runtime environment used to run the server-side code."],
      ["Express.js",     "Lightweight web framework for Node.js used to create API routes and serve static files."],
      ["sql.js",         "A JavaScript library that compiles SQLite to WebAssembly (WASM). Used as the project's database engine. The database runs in memory and is saved to disk as a binary .db file."],
      ["node-fetch",     "A lightweight library that lets Node.js make HTTP requests to external APIs (used to call ip-api.com and ipify.org)."],
      ["fs (built-in)",  "Node.js built-in File System module, used to read and write the SQLite database file from disk."],
    ], ["Technology", "Purpose"]),
    ...spacer(1),
    h3("4.2 Frontend Technologies"),
    twoColTable([
      ["HTML5",              "Structure of all web pages — the bait form, terminal dashboard, admin panel, and QR code page."],
      ["CSS3 (Vanilla CSS)", "All styling is done using pure CSS. The design follows a dark cyberpunk theme with glow effects, animations, and a CRT scanline overlay."],
      ["Vanilla JavaScript", "All client-side logic: device fingerprinting, camera capture, audio recording, geolocation, and API calls to the backend."],
      ["Google Fonts",       "JetBrains Mono and Fira Code fonts are loaded from Google Fonts CDN for the terminal aesthetic."],
    ], ["Technology", "Purpose"]),
    ...spacer(1),
    h3("4.3 Browser APIs Used"),
    twoColTable([
      ["navigator.mediaDevices.getUserMedia()", "Requests access to the device camera to capture still images."],
      ["MediaRecorder API",                     "Records audio from the device microphone for up to 10 seconds."],
      ["navigator.geolocation.getCurrentPosition()", "Retrieves the device's GPS latitude, longitude, and accuracy."],
      ["Canvas API",                            "Used to draw a video frame onto a hidden canvas element and export it as a JPEG image."],
      ["Fetch API",                             "Used to POST the collected data payload to the backend server."],
    ], ["Browser API", "What it Does"]),
    ...spacer(1),
    h3("4.4 External APIs"),
    twoColTable([
      ["ip-api.com",    "A free IP geolocation API. Given an IP address, it returns the city, region, country, ISP, organisation, and approximate coordinates."],
      ["api.ipify.org", "A simple API that returns the server's own public IP address. Used when the app runs on localhost (which has a private/loopback IP)."],
    ], ["API", "Purpose"]),
    pageBreak(),
  ];
}

// ── Section 5: System Architecture ───────────────────────────
function buildArchitecture() {
  return [
    h1("5. SYSTEM ARCHITECTURE"),
    para(
      "The application follows a classic Client–Server architecture. The frontend (browser) communicates with the backend (Node.js server) " +
      "through HTTP API requests. The backend stores data in a local SQLite database."
    ),
    ...spacer(1),
    h3("5.1 Architecture Diagram (Text Representation)"),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({ children: [new TableCell({
          children: [
            new Paragraph({ children: [bold("[ VISITOR'S BROWSER ]", 20, COLOR.white)], alignment: AlignmentType.CENTER }),
            new Paragraph({ children: [normal("HTML + CSS + Vanilla JS", 18, COLOR.light)], alignment: AlignmentType.CENTER }),
            new Paragraph({ children: [normal("Camera API | Mic API | GPS API | Fetch API", 18, COLOR.light)], alignment: AlignmentType.CENTER }),
          ],
          shading: { type: ShadingType.SOLID, color: COLOR.primary, fill: COLOR.primary },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          borders: { top: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent }, right: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent } },
        })] }),
        new TableRow({ children: [new TableCell({
          children: [new Paragraph({ children: [bold("↕  HTTP / JSON  (POST /api/capture, GET /api/admin/records)", 18, COLOR.primary)], alignment: AlignmentType.CENTER })],
          shading: { type: ShadingType.SOLID, color: COLOR.lightGray, fill: COLOR.lightGray },
          margins: { top: 80, bottom: 80, left: 200, right: 200 },
          borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent }, right: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent } },
        })] }),
        new TableRow({ children: [new TableCell({
          children: [
            new Paragraph({ children: [bold("[ NODE.JS SERVER — server.js ]", 20, COLOR.white)], alignment: AlignmentType.CENTER }),
            new Paragraph({ children: [normal("Express.js Framework | Route Handlers | IP Geolocation Logic", 18, COLOR.light)], alignment: AlignmentType.CENTER }),
          ],
          shading: { type: ShadingType.SOLID, color: COLOR.accent, fill: COLOR.accent },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent }, right: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent } },
        })] }),
        new TableRow({ children: [new TableCell({
          children: [new Paragraph({ children: [bold("↕  sql.js (SQLite WASM)  ↔  /data/records.db", 18, COLOR.primary)], alignment: AlignmentType.CENTER })],
          shading: { type: ShadingType.SOLID, color: COLOR.lightGray, fill: COLOR.lightGray },
          margins: { top: 80, bottom: 80, left: 200, right: 200 },
          borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent }, right: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent } },
        })] }),
        new TableRow({ children: [new TableCell({
          children: [
            new Paragraph({ children: [bold("[ SQLite DATABASE — records.db ]", 20, COLOR.white)], alignment: AlignmentType.CENTER }),
            new Paragraph({ children: [normal("Table: captures | Stores all visitor records persistently", 18, COLOR.light)], alignment: AlignmentType.CENTER }),
          ],
          shading: { type: ShadingType.SOLID, color: COLOR.primary, fill: COLOR.primary },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent }, left: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent }, right: { style: BorderStyle.SINGLE, size: 6, color: COLOR.accent } },
        })] }),
      ],
    }),
    ...spacer(1),
    h3("5.2 File Structure"),
    twoColTable([
      ["server.js",          "Main backend file — Express server, database init, all API route handlers"],
      ["package.json",       "Project metadata and list of npm dependencies"],
      ["data/records.db",    "SQLite binary database file persisted on disk (~2MB+)"],
      ["public/index.html",  "Main page — the fake job application bait form + terminal dashboard"],
      ["public/admin.html",  "Admin panel — view all captured records including images and audio"],
      ["public/qr.html",     "QR code generation page for easy mobile access to the capture URL"],
      ["public/script.js",   "All client-side JavaScript logic (625 lines) — captures, UI, animations"],
      ["public/style.css",   "All CSS styling (24KB) — cyberpunk dark theme, animations, layouts"],
    ], ["File", "Purpose"]),
    pageBreak(),
  ];
}

// ── Section 6: Database Design ────────────────────────────────
function buildDatabase() {
  return [
    h1("6. DATABASE DESIGN"),
    para(
      "The application uses SQLite as its database engine. SQLite is a lightweight, file-based database — meaning all data is stored in a single .db file " +
      "on the server's disk (data/records.db). The sql.js library is used to interact with SQLite from Node.js through WebAssembly (WASM), " +
      "which means the database runs in memory and is saved to the file after every write operation."
    ),
    ...spacer(1),
    h3("6.1 Table: captures"),
    para("There is a single table in the database called captures. It stores one row per visitor."),
    ...spacer(1),
    twoColTable([
      ["id",              "INTEGER (PK, Auto Increment)", "Unique ID for each capture record"],
      ["timestamp_open",  "TEXT",                        "ISO 8601 timestamp of when the page was opened"],
      ["os",              "TEXT",                        "Detected operating system (e.g., WINDOWS_NT, ANDROID)"],
      ["browser",         "TEXT",                        "Detected browser name (e.g., CHROME, FIREFOX)"],
      ["browser_version", "TEXT",                        "Version number of the detected browser"],
      ["device_type",     "TEXT",                        "WORKSTATION, MOBILE_UNIT, or TABLET_UNIT"],
      ["screen_res",      "TEXT",                        "Screen resolution (e.g., 1920×1080)"],
      ["language",        "TEXT",                        "Browser/OS language setting (e.g., EN-IN)"],
      ["platform",        "TEXT",                        "Raw platform string from navigator.platform"],
      ["user_agent",      "TEXT",                        "Full browser user-agent string"],
      ["ip_address",      "TEXT",                        "Resolved public IP address of the visitor"],
      ["ip_city",         "TEXT",                        "City resolved from IP (via ip-api.com)"],
      ["ip_region",       "TEXT",                        "State/region resolved from IP"],
      ["ip_country",      "TEXT",                        "Country resolved from IP"],
      ["ip_lat",          "REAL",                        "Approximate latitude from IP geolocation"],
      ["ip_lon",          "REAL",                        "Approximate longitude from IP geolocation"],
      ["ip_isp",          "TEXT",                        "Internet Service Provider name"],
      ["ip_org",          "TEXT",                        "Organisation associated with the IP"],
      ["ip_timezone",     "TEXT",                        "Timezone from IP (e.g., Asia/Kolkata)"],
      ["geo_lat",         "REAL",                        "Precise GPS latitude from browser Geolocation API"],
      ["geo_lon",         "REAL",                        "Precise GPS longitude from browser Geolocation API"],
      ["geo_accuracy",    "REAL",                        "GPS accuracy in metres (from browser)"],
      ["front_image",     "TEXT",                        "Base64-encoded JPEG from the front (selfie) camera"],
      ["back_image",      "TEXT",                        "Base64-encoded JPEG from the rear camera"],
      ["audio_recording", "TEXT",                        "Base64-encoded WebM/OGG audio (10 seconds)"],
      ["created_at",      "TEXT",                        "Auto-set server timestamp (datetime('now'))"],
    ], ["Column Name", "Data Type", "Description"]),
    pageBreak(),
  ];
}

// ── Section 7: Module-wise Description ───────────────────────
function buildModules() {
  return [
    h1("7. MODULE-WISE DESCRIPTION"),
    para("The project is organized into the following logical modules:"),
    ...spacer(1),
    h3("7.1 Bait / Social Engineering Module (public/index.html)"),
    para(
      "This is the first page the visitor sees. It is designed to look like a legitimate job application portal for a fictional company — " +
      "\"Nexus Quantum Solutions\". The job posting is for a \"System Engineer\" role with an attractive salary, designed to make the visitor fill in their details. " +
      "The key trick is the \"Compliance Location Lock\" field — clicking on it triggers all the browser permission requests (location, camera, microphone) in the background."
    ),
    ...spacer(1),
    h3("7.2 Device Fingerprinting Module (public/script.js — scanDevice())"),
    para(
      "This module reads several properties available from the browser's JavaScript environment without requiring any user permission. " +
      "It detects the operating system, browser name and version, device type, screen resolution, browser language, and platform string by parsing the User-Agent string."
    ),
    ...spacer(1),
    h3("7.3 Camera Capture Module (public/script.js — captureCamera())"),
    para(
      "This module uses the browser's getUserMedia API to request camera access. When granted, it opens a hidden video stream, waits 800ms for the camera to warm up, " +
      "draws one frame onto a hidden HTML Canvas element, and exports it as a JPEG image encoded in Base64. " +
      "It attempts both the front-facing camera (facingMode: user) and the rear camera (facingMode: environment) in sequence."
    ),
    ...spacer(1),
    h3("7.4 Audio Recording Module (public/script.js — recordAudio())"),
    para(
      "This module uses the browser's MediaRecorder API to record 10 seconds of audio from the device microphone. " +
      "The recording is stored as audio chunks and then assembled into a Blob, which is converted to a Base64 data URL for transmission. " +
      "The audio is hidden from the visitor; it is only visible to the admin."
    ),
    ...spacer(1),
    h3("7.5 Geolocation Module (public/script.js — getGeoLocation())"),
    para(
      "This module calls the browser's built-in Geolocation API with high accuracy enabled. " +
      "It retrieves the visitor's precise GPS latitude, longitude, and accuracy (in metres). If the user denies permission, the module gracefully fails and the " +
      "IP-based approximate location is used instead."
    ),
    ...spacer(1),
    h3("7.6 Backend Capture & Storage Module (server.js — POST /api/capture)"),
    para(
      "This is the server-side module that receives the payload from the browser. It resolves the visitor's IP address, calls the external ip-api.com API " +
      "to get city, region, ISP, and coordinates, and then inserts all the data into the SQLite database via sql.js."
    ),
    ...spacer(1),
    h3("7.7 Admin Panel Module (public/admin.html)"),
    para(
      "A password-protected dashboard that allows the project administrator to view all captured records. " +
      "It lists all records in a table (without images/audio for speed), and allows clicking a record to view the full details including camera images and audio playback."
    ),
    ...spacer(1),
    h3("7.8 Terminal Dashboard / Reveal Module (public/script.js — runHackSequence())"),
    para(
      "After form submission and data transmission, the bait form is hidden and a retro cyberpunk terminal dashboard is shown. " +
      "An animated 'hack sequence' plays out — logging each phase (Device Scan, Biometric Capture, Voice Intercept, GPS Lock, Payload Transmission, Network Trace) " +
      "in real time with color-coded log entries, simulating a hacking tool interface."
    ),
    pageBreak(),
  ];
}

// ── Section 8: API Documentation ─────────────────────────────
function buildAPI() {
  return [
    h1("8. API DOCUMENTATION"),
    para("The backend exposes the following RESTful API endpoints:"),
    ...spacer(1),
    h3("8.1 POST /api/capture"),
    twoColTable([
      ["Method",      "POST"],
      ["URL",         "/api/capture"],
      ["Auth",        "None (public endpoint)"],
      ["Content-Type","application/json"],
      ["Request Body","JSON object with device metadata, GPS coords, base64 images, and audio"],
      ["Response",    "JSON: { success: true, recordId: number, data: { ip_address, ip_city, ... } }"],
      ["Purpose",     "Receives and stores a full visitor capture record in the database"],
    ], ["Property", "Value"]),
    ...spacer(1),
    h3("8.2 GET /api/admin/records"),
    twoColTable([
      ["Method",   "GET"],
      ["URL",      "/api/admin/records"],
      ["Auth",     "Header: X-Admin-Pass: nexus2026  OR  Query: ?pass=nexus2026"],
      ["Response", "JSON: { success: true, count: number, records: [ ... ] }"],
      ["Purpose",  "Returns all capture records (metadata only — no images or audio for performance)"],
    ], ["Property", "Value"]),
    ...spacer(1),
    h3("8.3 GET /api/admin/record/:id"),
    twoColTable([
      ["Method",   "GET"],
      ["URL",      "/api/admin/record/:id  (e.g., /api/admin/record/5)"],
      ["Auth",     "Header: X-Admin-Pass: nexus2026  OR  Query: ?pass=nexus2026"],
      ["Response", "JSON: { success: true, record: { all columns including images and audio } }"],
      ["Purpose",  "Returns a single full record by ID, including Base64 camera images and audio"],
    ], ["Property", "Value"]),
    pageBreak(),
  ];
}

// ── Section 9: Application Workflow ──────────────────────────
function buildWorkflow() {
  return [
    h1("9. APPLICATION WORKFLOW"),
    para("The following step-by-step flow describes how the application works from the moment a visitor opens the page:"),
    ...spacer(1),
    twoColTable([
      ["Step 1",  "Visitor opens the URL in a browser. They see a professional-looking job application form for 'Nexus Quantum Solutions — System Engineer'."],
      ["Step 2",  "Visitor fills in their name, email, phone number, and date of birth."],
      ["Step 3",  "Visitor clicks the 'Compliance Location Lock' field. This triggers three browser permission requests in the background: Location, Camera, and Microphone."],
      ["Step 4",  "If permission is granted: GPS coordinates are captured, front camera photo is taken, rear camera photo is taken, and 10 seconds of audio is recorded — all silently in the background."],
      ["Step 5",  "Visitor clicks 'Submit Secure Application'. The form is disabled and a loading spinner appears."],
      ["Step 6",  "The client-side JavaScript waits for all captures to finish (Promise.all), then scans the device fingerprint (OS, browser, screen resolution, etc.)."],
      ["Step 7",  "All captured data is compiled into a single JSON payload and sent to the server via POST /api/capture."],
      ["Step 8",  "The server resolves the visitor's IP address, calls ip-api.com to get geolocation data, and inserts the complete record into the SQLite database."],
      ["Step 9",  "The server responds with a success message, the record ID, and the IP geolocation data."],
      ["Step 10", "The bait form disappears. A cyberpunk terminal dashboard fades in, and an animated 'hack sequence' log plays through all 6 phases, revealing all captured data to the visitor."],
    ], ["Step", "Description"]),
    pageBreak(),
  ];
}

// ── Section 10: UI Description ────────────────────────────────
function buildUI() {
  return [
    h1("10. USER INTERFACE DESCRIPTION"),
    para("The application has three distinct UI pages, each with a specific purpose:"),
    ...spacer(1),
    h3("10.1 Bait Page (index.html — Phase 1)"),
    para(
      "The bait page is deliberately designed to look trustworthy and professional. It features a clean card layout with a company logo, job title, salary range, " +
      "and a standard-looking form. The color scheme is neutral and corporate. The 'Compliance Location Lock' field is styled to look like a standard location input."
    ),
    ...spacer(1),
    h3("10.2 Terminal Dashboard (index.html — Phase 2, revealed after form submit)"),
    para("The terminal dashboard has the following UI components:"),
    bullet("Top Ticker Bar: A scrolling banner with system status messages"),
    bullet("Top Bar: Shows the NEXUS//TRACE branding, a LIVE indicator, system protocol status, and a live clock"),
    bullet("Status Bar: Shows the visitor trace ID, encryption type, session ID, and current phase"),
    bullet("Operation Log Panel (left): A real-time log of all capture phases, color-coded (green = success, amber = warning, red = error)"),
    bullet("Data Reveal Panels (right): Five sections that appear one by one — Device Fingerprint, Biometric Capture (camera images), Voice Intercept (waveform), GPS Coordinates, and Network/IP Trace"),
    bullet("Bottom Bar: Shows payload delivery status, total record count, and extraction status"),
    bullet("Bottom Ticker: A reverse-scrolling banner"),
    ...spacer(1),
    h3("10.3 Admin Panel (admin.html)"),
    para(
      "A password-gated table that shows all capture records. Each row can be expanded to view the full record including camera photos (displayed as images) " +
      "and the audio recording (with an HTML5 audio player for playback)."
    ),
    ...spacer(1),
    h3("10.4 QR Code Page (qr.html)"),
    para(
      "A utility page that generates a QR code for the application's URL. This is useful for sharing the link to mobile devices for testing camera and microphone captures."
    ),
    pageBreak(),
  ];
}

// ── Section 11: Security Considerations ──────────────────────
function buildSecurity() {
  return [
    h1("11. SECURITY CONSIDERATIONS"),
    para("While building this project, the following security aspects were observed and noted:"),
    ...spacer(1),
    twoColTable([
      ["Browser Permission Model",   "Modern browsers require explicit user permission for camera, microphone, and location access. The application cannot bypass these — it relies on social engineering to convince the user to grant them."],
      ["HTTPS Requirement",          "Camera and microphone access (getUserMedia) are blocked by browsers on plain HTTP connections. For real-world deployment on mobile devices, HTTPS is mandatory (e.g., via ngrok tunnel or SSL certificate)."],
      ["Hardcoded Admin Password",   "The admin password 'nexus2026' is hardcoded in plaintext in server.js. In a production system, this should be stored as an environment variable and hashed."],
      ["No Rate Limiting",           "The /api/capture endpoint has no rate limiting, making it vulnerable to spam and denial-of-service attacks."],
      ["No Input Validation",        "Data received from the client is inserted into the database without sanitization. In a production system, all inputs should be validated and sanitized."],
      ["Large Data in SQLite",       "Storing Base64-encoded images and audio directly in SQLite TEXT columns causes the database file to grow rapidly. A proper solution would store media files on disk or in object storage."],
    ], ["Security Aspect", "Explanation"]),
    pageBreak(),
  ];
}

// ── Section 12: Limitations & Future Enhancements ────────────
function buildLimitations() {
  return [
    h1("12. LIMITATIONS & FUTURE ENHANCEMENTS"),
    ...spacer(1),
    h3("12.1 Current Limitations"),
    bullet("The database (sql.js) loads entirely into memory — not suitable for large datasets"),
    bullet("Admin authentication is a single hardcoded plaintext password"),
    bullet("No rate limiting or spam protection on the capture endpoint"),
    bullet("Base64 media in SQLite causes rapid database size growth"),
    bullet("No email/notification system when a new capture is received"),
    bullet("The application requires HTTPS for camera/mic on mobile (ngrok workaround needed)"),
    ...spacer(1),
    h3("12.2 Future Enhancements"),
    bullet("Replace sql.js with a proper database like PostgreSQL or MongoDB for scalability"),
    bullet("Store media files (images, audio) on disk or cloud storage (S3) and save only file paths in the DB"),
    bullet("Add JWT-based authentication for the admin panel"),
    bullet("Implement rate limiting using the express-rate-limit package"),
    bullet("Add real-time notifications (WebSockets or Server-Sent Events) for new captures"),
    bullet("Build a map view in the admin panel using Leaflet.js to visualize GPS locations"),
    bullet("Add export functionality (CSV/Excel export of all records)"),
    bullet("Implement email alerts when new captures arrive"),
    pageBreak(),
  ];
}

// ── Section 13: Conclusion ────────────────────────────────────
function buildConclusion() {
  return [
    h1("13. CONCLUSION"),
    para(
      "The NEXUS TRACE project successfully demonstrates how modern web technologies can be combined to build a full-stack information gathering application. " +
      "Through this project, the following key skills and concepts were learned and applied:"
    ),
    bullet("Building a RESTful API server using Node.js and Express.js"),
    bullet("Working with a file-based SQLite database using the sql.js library"),
    bullet("Using browser media APIs (getUserMedia, MediaRecorder, Geolocation) in a real-world application"),
    bullet("Integrating third-party HTTP APIs (ip-api.com, ipify.org) for IP geolocation"),
    bullet("Designing a rich, animated user interface using pure CSS with no external CSS frameworks"),
    bullet("Understanding the concept of social engineering and how it is used in phishing-style attacks"),
    bullet("Implementing an admin panel with basic authentication for reviewing captured data"),
    ...spacer(1),
    para(
      "This project serves as an important educational reference for understanding both the capabilities and the dangers of modern web browser APIs. " +
      "It highlights the importance of being cautious about granting permissions to websites, and demonstrates why browser security features like HTTPS enforcement " +
      "for sensitive APIs and the same-origin policy exist."
    ),
    ...spacer(1),
    infoBox(
      "🎓 Key Takeaway",
      "Modern browsers are powerful platforms. With user permission, a single web page can access a device's location, camera, and microphone. " +
      "This project shows how these capabilities work — and why it is critical to only grant such permissions to websites you fully trust."
    ),
    ...spacer(2),
    divider(),
    new Paragraph({
      children: [
        bold("Submitted by: ", 22, COLOR.primary), normal("Praveen Narni", 22),
      ],
      spacing: { before: 100, after: 60 },
    }),
    new Paragraph({
      children: [
        bold("Guided by: ", 22, COLOR.primary), normal("Sai Shanthan", 22),
      ],
      spacing: { before: 60, after: 100 },
    }),
  ];
}

// ── Build & Write Document ────────────────────────────────────
async function generateReport() {
  const doc = new Document({
    numbering: { config: [] },
    styles: {
      default: {
        document: {
          run: { font: "Calibri", size: 22, color: COLOR.black },
          paragraph: { spacing: { line: 320 } },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: { top: 1440, bottom: 1440, left: 1440, right: 1080 },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "NEXUS TRACE — Project Report", size: 18, color: COLOR.gray, font: "Calibri" }),
                  new TextRun({ text: "\t\t", size: 18 }),
                  new TextRun({ text: "Submitted by: Praveen Narni | Guided by: Sai Shanthan", size: 18, color: COLOR.gray, font: "Calibri" }),
                ],
                border: { bottom: { color: COLOR.accent, space: 1, size: 4, style: BorderStyle.SINGLE } },
                tabStops: [{ type: "right", position: 9072 }],
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "NEXUS TRACE | Cybersecurity Demonstration Project | Page ", size: 18, color: COLOR.gray }),
                  new PageNumberElement(),
                ],
                alignment: AlignmentType.CENTER,
                border: { top: { color: COLOR.accent, space: 1, size: 4, style: BorderStyle.SINGLE } },
              }),
            ],
          }),
        },
        children: [
          ...buildCoverPage(),
          ...buildToC(),
          ...buildIntro(),
          ...buildProblemStatement(),
          ...buildObjectives(),
          ...buildTechStack(),
          ...buildArchitecture(),
          ...buildDatabase(),
          ...buildModules(),
          ...buildAPI(),
          ...buildWorkflow(),
          ...buildUI(),
          ...buildSecurity(),
          ...buildLimitations(),
          ...buildConclusion(),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  const outPath = path.join(__dirname, "NEXUS_TRACE_Project_Report.docx");
  fs.writeFileSync(outPath, buffer);
  console.log(`\n✅ Report generated successfully!`);
  console.log(`📄 File saved at: ${outPath}\n`);
}

generateReport().catch(err => {
  console.error("❌ Error generating report:", err);
  process.exit(1);
});
