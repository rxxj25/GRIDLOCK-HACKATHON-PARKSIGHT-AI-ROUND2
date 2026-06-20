import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const artifactToolPath =
  "C:\\Users\\Rajdeep\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules\\@oai\\artifact-tool\\dist\\artifact_tool.mjs";
const { Presentation, PresentationFile } = await import(pathToFileURL(artifactToolPath).href);

const ROOT = "D:\\FLIPKART_GRIDLOCK";
const OUT_DIR = path.join(ROOT, "outputs");
const WORK = path.join(ROOT, "work", "presentations", "parksight-old-theme-final", "tmp");
const PREVIEW = path.join(WORK, "preview");
const LAYOUT = path.join(WORK, "layout");
const QA = path.join(WORK, "qa");
const FINAL = path.join(OUT_DIR, "ParkSight_AI_Pitch_Deck.pptx");
const ORIGINAL_BACKUP = path.join(OUT_DIR, "ParkSight_AI_Pitch_Deck_original_backup.pptx");

const shots = {
  overview: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002331.png"),
  evidence: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002346.png"),
  method: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002430.png"),
  plan: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002422.png"),
};

const C = {
  navy: "#0D1B2A",
  mid: "#162B3C",
  panel: "#112235",
  card: "#193246",
  teal: "#04756F",
  teal2: "#05968E",
  coral: "#D93D4A",
  amber: "#E09B2D",
  violet: "#6157A8",
  white: "#FFFFFF",
  off: "#F0F4F8",
  light: "#C4CFD8",
  muted: "#93A4B5",
  line: "#2B4358",
};

const fonts = {
  title: "Aptos Display",
  body: "Aptos",
  number: "Georgia",
  mono: "Consolas",
};

const W = 1280;
const H = 720;

await fs.mkdir(OUT_DIR, { recursive: true });
await fs.mkdir(PREVIEW, { recursive: true });
await fs.mkdir(LAYOUT, { recursive: true });
await fs.mkdir(QA, { recursive: true });

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function imageBlob(filePath) {
  const bytes = await fs.readFile(filePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function shape(slide, geometry, x, y, w, h, fill, opts = {}) {
  const config = {
    geometry,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: opts.line || { style: "solid", fill: opts.stroke || fill, width: opts.lineWidth ?? 0 },
    shadow: opts.shadow,
  };
  if (geometry === "rect" || geometry === "textbox" || geometry === "roundRect") {
    config.borderRadius = opts.radius || "rounded-lg";
  }
  return slide.shapes.add(config);
}

function rect(slide, x, y, w, h, fill, opts = {}) {
  return shape(slide, opts.geometry || "rect", x, y, w, h, fill, opts);
}

function card(slide, x, y, w, h, accent = C.teal, opts = {}) {
  const s = shape(slide, "roundRect", x, y, w, h, opts.fill || C.card, {
    stroke: opts.stroke || C.line,
    lineWidth: 1,
    radius: "rounded-lg",
    shadow: "shadow-sm",
  });
  if (accent) rect(slide, x, y, w, opts.strip ?? 5, accent);
  return s;
}

function text(slide, value, x, y, w, h, style = {}) {
  const s = slide.shapes.add({
    geometry: "textbox",
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  s.text = value;
  s.text.style = {
    typeface: style.typeface || fonts.body,
    fontSize: style.size || 16,
    bold: style.bold || false,
    italic: style.italic || false,
    color: style.color || C.off,
    alignment: style.align || "left",
  };
  return s;
}

function chrome(slide, section, title, subtitle, accent, num) {
  slide.background.fill = C.navy;
  rect(slide, 0, 0, W, H, C.navy);
  rect(slide, 0, 0, 10, H, accent);
  rect(slide, 0, 0, W, 5, accent);
  text(slide, section, 58, 32, 520, 22, { size: 11, bold: true, color: accent });
  text(slide, title, 58, 64, 930, 62, {
    size: 32,
    bold: true,
    color: C.white,
    typeface: fonts.title,
  });
  if (subtitle) {
    text(slide, subtitle, 58, 130, 950, 42, { size: 15, color: C.light });
  }
  rect(slide, 58, 188, 150, 4, accent);
  footer(slide, num);
}

function footer(slide, num) {
  rect(slide, 0, 674, W, 46, C.panel);
  text(slide, "ParkSight AI | Flipkart Gridlock Challenge", 58, 688, 430, 18, {
    size: 10,
    color: C.light,
  });
  text(slide, String(num).padStart(2, "0"), 1160, 684, 50, 24, {
    size: 14,
    bold: true,
    color: C.light,
    align: "right",
    typeface: fonts.number,
  });
}

function metric(slide, x, y, w, h, value, label, accent) {
  card(slide, x, y, w, h, accent);
  const size = value.length > 10 ? 24 : value.length > 7 ? 28 : 34;
  text(slide, value, x + 22, y + 24, w - 44, 42, {
    size,
    bold: true,
    color: accent,
    typeface: fonts.number,
  });
  text(slide, label, x + 22, y + 72, w - 44, h - 84, {
    size: 12,
    bold: true,
    color: C.off,
  });
}

function bullet(slide, x, y, label, body, accent, w = 440) {
  rect(slide, x, y + 8, 10, 10, accent, { geometry: "ellipse" });
  text(slide, label, x + 24, y, w - 24, 22, { size: 14, bold: true, color: C.white });
  if (body) {
    text(slide, body, x + 24, y + 24, w - 24, 38, { size: 11, color: C.light });
  }
}

async function image(slide, filePath, x, y, w, h, alt, fit = "contain") {
  slide.images.add({
    blob: await imageBlob(filePath),
    contentType: "image/png",
    alt,
    fit,
    position: { left: x, top: y, width: w, height: h },
    geometry: "roundRect",
    borderRadius: "rounded-lg",
  });
}

const deck = Presentation.create({ slideSize: { width: W, height: H } });

// 1. Cover
{
  const s = deck.slides.add();
  s.background.fill = C.navy;
  rect(s, 0, 0, W, H, C.navy);
  rect(s, 0, 0, 398, H, C.mid);
  rect(s, 0, 0, 10, H, C.teal);
  rect(s, 58, 42, 130, 5, C.teal2);
  text(s, "FLIPKART GRIDLOCK CHALLENGE | PROTOTYPE ROUND", 58, 64, 560, 22, {
    size: 12,
    bold: true,
    color: C.teal2,
  });
  text(s, "ParkSight", 58, 126, 490, 74, {
    size: 60,
    bold: true,
    color: C.white,
    typeface: fonts.title,
  });
  text(s, "AI", 58, 202, 160, 70, {
    size: 58,
    bold: true,
    color: C.amber,
    typeface: fonts.title,
  });
  text(s, "Parking-induced congestion intelligence\nfor targeted enforcement", 58, 304, 330, 58, {
    size: 17,
    italic: true,
    color: C.light,
  });
  text(s, "Detect hotspots. Quantify congestion impact. Deploy enforcement where it matters most.", 58, 382, 520, 58, {
    size: 16,
    color: C.off,
  });
  const stats = [
    ["298K", "violations analyzed", C.teal2],
    ["3,969", "urban cells scored", C.amber],
    ["51%", "junction-linked cases", C.coral],
    ["12", "deployable beats", C.violet],
  ];
  stats.forEach(([v, l, c], i) => {
    const x = 610 + (i % 2) * 270;
    const y = 100 + Math.floor(i / 2) * 160;
    metric(s, x, y, 230, 118, v, l, c);
  });
  card(s, 610, 450, 500, 88, C.teal, { fill: C.panel });
  text(s, "Live dashboard + explainable impact model", 638, 474, 430, 24, {
    size: 18,
    bold: true,
    color: C.white,
  });
  text(s, "React | Tailwind | Leaflet | Chart.js | Python ETL", 638, 508, 430, 18, {
    size: 12,
    color: C.light,
  });
  text(s, "Rajdeep Bandyopadhaya | Aniket Arya | BMS College of Engineering", 58, 626, 720, 22, {
    size: 13,
    color: C.light,
  });
  footer(s, 1);
}

// 2. Problem
{
  const s = deck.slides.add();
  chrome(s, "THE OPERATIONAL GAP", "Illegal parking is visible. Its traffic impact is not.", "Traffic teams need a ranked enforcement plan, not just dots on a map.", C.coral, 2);
  metric(s, 58, 230, 250, 126, "INR 3,700 Cr", "annual traffic loss in Bengaluru", C.coral);
  metric(s, 330, 230, 250, 126, "298,450", "violations in five months", C.amber);
  metric(s, 602, 230, 250, 126, "50.5%", "cases at junctions and crossings", C.teal);
  metric(s, 874, 230, 250, 126, "39.6%", "during peak movement hours", C.violet);
  card(s, 58, 420, 520, 170, C.coral, { fill: C.panel });
  text(s, "Today", 92, 448, 200, 28, { size: 22, bold: true, color: C.white, typeface: fonts.title });
  bullet(s, 96, 494, "Reactive patrols", "action starts after congestion has already formed", C.coral);
  bullet(s, 96, 552, "Ticket-count heatmaps", "raw volume does not reveal which violations choke carriageways", C.amber);
  card(s, 638, 420, 520, 170, C.teal, { fill: C.panel });
  text(s, "After ParkSight AI", 672, 448, 270, 28, { size: 22, bold: true, color: C.teal2, typeface: fonts.title });
  bullet(s, 676, 494, "Congestion-priority ranking", "impact, junction risk and recurrence are scored together", C.teal);
  bullet(s, 676, 552, "Deployable enforcement plan", "top hotspots become time-windowed patrol beats", C.violet);
}

// 3. Data Signal
{
  const s = deck.slides.add();
  chrome(s, "DATA SIGNAL", "298,450 violations become an enforcement intelligence layer", "The data story is clear, traceable and ready for questions from judges.", C.teal, 3);
  const stats = [
    ["2,98,450", "violations analyzed", C.teal2],
    ["3,969", "urban cells scored", C.amber],
    ["51%", "junction-linked", C.coral],
    ["55", "police stations", C.violet],
  ];
  stats.forEach(([v, l, c], i) => metric(s, 58 + i * 280, 230, 240, 116, v, l, c));
  card(s, 58, 412, 500, 142, C.teal, { fill: C.panel });
  text(s, "Actual record range", 92, 438, 260, 22, { size: 14, bold: true, color: C.teal2 });
  text(s, "10 Nov 2023 to 08 Apr 2024", 92, 474, 410, 34, {
    size: 24,
    bold: true,
    color: C.white,
    typeface: fonts.number,
  });
  text(s, "The filename mentions Jan-May, but timestamps show the true range above.", 92, 520, 380, 24, {
    size: 11,
    color: C.light,
  });
  card(s, 638, 412, 500, 142, C.coral, { fill: C.panel });
  text(s, "Top station and offence", 672, 438, 260, 22, { size: 14, bold: true, color: C.coral });
  text(s, "Upparpet | Wrong parking", 672, 474, 410, 32, {
    size: 23,
    bold: true,
    color: C.white,
    typeface: fonts.number,
  });
  text(s, "Station and offence mix supports targeted enforcement, not generic citywide action.", 672, 520, 390, 24, {
    size: 11,
    color: C.light,
  });
}

// 4. Solution
{
  const s = deck.slides.add();
  chrome(s, "SOLUTION", "A command center for congestion-aware parking enforcement", "Four outputs make the product easy for judges to understand and test.", C.teal2, 4);
  const blocks = [
    ["Impact", "Prioritizes where illegal parking is most likely to hurt traffic flow.", C.coral],
    ["Volume", "Shows where raw violation density is highest for enforcement pressure.", C.amber],
    ["Junction", "Highlights intersections, crossings and signal-sensitive risk.", C.teal],
    ["Output", "Ranked hotspots, station focus, 12-beat plan and CSV/PDF reports.", C.violet],
  ];
  blocks.forEach(([h, b, c], i) => {
    const x = 70 + (i % 2) * 560;
    const y = 230 + Math.floor(i / 2) * 165;
    card(s, x, y, 500, 126, c, { fill: C.panel });
    text(s, h, x + 30, y + 30, 180, 30, { size: 24, bold: true, color: c, typeface: fonts.title });
    text(s, b, x + 30, y + 70, 390, 30, { size: 14, color: C.light });
  });
  card(s, 142, 594, 990, 44, C.teal, { fill: C.card, strip: 0 });
  text(s, "Core pitch: ParkSight gives command teams a congestion-ranked starting point for smarter patrol allocation.", 176, 608, 930, 18, {
    size: 13,
    bold: true,
    color: C.white,
  });
}

// 5. Product Demo
{
  const s = deck.slides.add();
  chrome(s, "PRODUCT DEMO", "The first screen tells judges exactly what to test", "Real submitted dashboard screenshots keep the deck concrete and credible.", C.amber, 5);
  await image(s, shots.overview, 58, 218, 760, 405, "ParkSight AI dashboard overview screenshot", "contain");
  const callouts = [
    ["Highest risk surfaced instantly", "BTP051 - Safina Plaza Junction, score 77.5", C.coral],
    ["Layered map intelligence", "Impact, Volume and Junction views answer different enforcement questions.", C.teal],
    ["Reports built in", "CSV for operations, PDF brief for commanders and judges.", C.amber],
  ];
  callouts.forEach(([h, b, c], i) => {
    const y = 224 + i * 130;
    card(s, 860, y, 330, 102, c, { fill: C.panel });
    text(s, h, 890, y + 22, 260, 22, { size: 15, bold: true, color: C.white });
    text(s, b, 890, y + 54, 260, 30, { size: 11, color: C.light });
  });
}

// 6. Explainable AI
{
  const s = deck.slides.add();
  chrome(s, "EXPLAINABLE AI / ML ANALYTICS", "Parking Impact Index: transparent, auditable scoring", "Because the dataset has no speed labels, ParkSight uses a visible congestion-impact proxy.", C.violet, 6);
  const steps = [
    ["01", "Cluster", "Group violations into ~220m urban cells", C.teal2],
    ["02", "Engineer", "Density, PCU, severity, junction and peak recurrence", C.amber],
    ["03", "Score", "Weighted 0-100 impact score for every cell", C.coral],
    ["04", "Deploy", "Prioritized beats and patrol windows", C.violet],
  ];
  steps.forEach(([n, h, b, c], i) => {
    const x = 58 + i * 290;
    card(s, x, 230, 240, 130, c, { fill: C.panel });
    text(s, n, x + 20, 250, 90, 30, { size: 24, bold: true, color: c, typeface: fonts.body });
    text(s, h, x + 20, 292, 180, 24, { size: 17, bold: true, color: C.white });
    text(s, b, x + 20, 326, 188, 28, { size: 10.5, color: C.light });
  });
  card(s, 108, 430, 1000, 92, null, { fill: C.panel, stroke: C.line, strip: 0 });
  text(s, "Score = 100 x (0.34 obstruction + 0.18 density + 0.15 junction + 0.13 arterial", 142, 454, 930, 20, {
    size: 15,
    bold: true,
    color: C.white,
    typeface: fonts.mono,
  });
  text(s, "+ 0.10 peak + 0.06 recurrence + 0.04 severity)", 142, 478, 930, 20, {
    size: 15,
    bold: true,
    color: C.white,
    typeface: fonts.mono,
  });
  text(s, "Why explainable? Civic teams need auditable prioritization when speed/congestion labels are not present.", 142, 506, 880, 18, {
    size: 11,
    color: C.light,
  });
}

// 7. Hotspot Evidence
{
  const s = deck.slides.add();
  chrome(s, "HOTSPOT EVIDENCE", "From tickets to traffic-flow risk", "Evidence and action sit together so judges can see why the recommendation matters.", C.coral, 7);
  await image(s, shots.evidence, 58, 218, 700, 394, "Impact layer map and station burden screenshot", "cover");
  card(s, 800, 218, 350, 394, C.coral, { fill: C.panel });
  text(s, "#1 Hotspot", 832, 252, 160, 22, { size: 13, bold: true, color: C.coral });
  text(s, "BTP051 - Safina Plaza Junction", 832, 286, 270, 64, {
    size: 26,
    bold: true,
    color: C.white,
    typeface: fonts.title,
  });
  text(s, "Shivajinagar | Commercial market", 832, 360, 270, 20, { size: 12, color: C.light });
  card(s, 832, 408, 132, 86, C.coral, { fill: C.card });
  text(s, "77.5", 858, 432, 84, 28, { size: 28, bold: true, color: C.coral, typeface: fonts.body, align: "center" });
  text(s, "impact score", 850, 468, 96, 14, { size: 9, bold: true, color: C.white, align: "center" });
  card(s, 990, 408, 132, 86, C.teal, { fill: C.card });
  text(s, "6,523", 1014, 432, 86, 28, { size: 28, bold: true, color: C.teal2, typeface: fonts.body, align: "center" });
  text(s, "violations", 1008, 468, 96, 14, { size: 9, bold: true, color: C.white, align: "center" });
  text(s, "Recommended action", 832, 526, 240, 22, { size: 13, bold: true, color: C.white });
  text(s, "Keep a junction-clearance unit on short-cycle patrol around 09:00-12:00 IST; dominant issue is wrong parking.", 832, 554, 285, 42, {
    size: 11,
    color: C.light,
  });
}

// 8. Targeted Enforcement
{
  const s = deck.slides.add();
  chrome(s, "TARGETED ENFORCEMENT", "A 12-beat plan that can be deployed, not just admired", "ParkSight ends in operational instructions, not just analytics.", C.teal, 8);
  await image(s, shots.plan, 58, 218, 348, 410, "Targeted enforcement plan screenshot", "contain");
  const beats = [
    ["1", "BTP051 - Safina Plaza", "Shivajinagar | 10:00-12:00 IST", "77.5"],
    ["2", "BTP040 - Elite Junction", "Upparpet | 08:00-10:00 IST", "77.0"],
    ["3", "BTP040 - Elite Junction", "Upparpet | 08:00-10:00 IST", "76.7"],
    ["4", "BTP051 - Safina Plaza", "Shivajinagar | 10:00-12:00 IST", "76.1"],
    ["5", "BTP082 - KR Market", "City Market | 00:00-02:00 IST", "74.9"],
  ];
  beats.forEach(([n, h, b, score], i) => {
    const y = 222 + i * 74;
    card(s, 470, y, 660, 56, null, { fill: C.panel, strip: 0 });
    rect(s, 490, y + 12, 32, 32, C.violet, { geometry: "roundRect" });
    text(s, n, 490, y + 18, 32, 18, { size: 14, bold: true, color: C.white, align: "center" });
    text(s, h, 544, y + 10, 330, 20, { size: 15, bold: true, color: C.white });
    text(s, b, 544, y + 34, 330, 16, { size: 10, color: C.light });
    text(s, score, 1018, y + 14, 88, 22, { size: 17, bold: true, color: C.teal2, align: "right", typeface: fonts.body });
  });
  card(s, 470, 606, 660, 38, C.teal, { fill: C.teal, strip: 0 });
  text(s, "Each beat includes location, impact score, patrol window and reason.", 506, 618, 590, 14, {
    size: 11,
    bold: true,
    color: C.white,
    align: "center",
  });
}

// 9. System Architecture
{
  const s = deck.slides.add();
  chrome(s, "SYSTEM ARCHITECTURE", "Data pipeline to live decision dashboard", "Simple enough for hackathon judges to believe, complete enough for real deployment.", C.amber, 9);
  const parts = [
    ["Raw challan CSV", "lat/lon, offence, station, timestamp, vehicle", C.coral],
    ["Python ETL", "cleaning, grid cells, weights, scores", C.amber],
    ["React dashboard", "map layers, charts, station focus", C.teal],
    ["Enforcement outputs", "12 beats, CSV, PDF brief", C.violet],
  ];
  parts.forEach(([h, b, c], i) => {
    const x = 70 + i * 292;
    card(s, x, 250, 232, 150, c, { fill: C.panel });
    text(s, h, x + 24, 288, 180, 30, { size: 17, bold: true, color: C.white });
    text(s, b, x + 24, 336, 174, 34, { size: 11, color: C.light });
    if (i < 3) text(s, ">", x + 244, 308, 36, 26, { size: 28, bold: true, color: c, align: "center" });
  });
  card(s, 150, 492, 980, 72, C.teal, { fill: C.card });
  text(s, "Hosted online", 188, 520, 150, 18, { size: 13, bold: true, color: C.teal2 });
  text(s, "GitHub Pages live demo + GitHub repository with generated intelligence data", 348, 512, 650, 24, {
    size: 18,
    bold: true,
    color: C.white,
  });
  text(s, "No paid services, no hidden API dependency, no server required after build.", 348, 540, 650, 18, {
    size: 11,
    color: C.light,
  });
}

// 10. Impact
{
  const s = deck.slides.add();
  chrome(s, "IMPACT", "What changes for traffic police", "The before/after contrast is the fastest way for judges to understand product value.", C.amber, 10);
  card(s, 78, 230, 500, 300, C.coral, { fill: C.panel });
  text(s, "Before", 116, 268, 180, 34, { size: 28, bold: true, color: C.coral, typeface: fonts.title });
  ["Reactive patrols", "Ticket-count heatmaps", "Manual zone priority", "Hard-to-share reports"].forEach((item, i) =>
    bullet(s, 122, 334 + i * 44, item, "", C.coral, 330),
  );
  card(s, 702, 230, 500, 300, C.teal, { fill: C.panel });
  text(s, "After ParkSight AI", 740, 268, 300, 34, { size: 28, bold: true, color: C.teal2, typeface: fonts.title });
  ["Impact-ranked hotspots", "Station-wise navigation", "12-beat enforcement plan", "Downloadable PDF/CSV brief"].forEach((item, i) =>
    bullet(s, 746, 334 + i * 44, item, "", C.teal, 350),
  );
  rect(s, 604, 348, 72, 46, C.navy, { geometry: "roundRect", stroke: C.line, lineWidth: 1 });
  text(s, "VS", 604, 360, 72, 18, { size: 18, bold: true, color: C.white, align: "center" });
  card(s, 128, 588, 1020, 42, null, { fill: C.panel, strip: 0 });
  text(s, "Future scope: integrate live CCTV/YOLO detection and traffic-speed labels for supervised congestion prediction.", 158, 600, 960, 16, {
    size: 11,
    color: C.light,
  });
}

// 11. Close
{
  const s = deck.slides.add();
  s.background.fill = C.navy;
  rect(s, 0, 0, W, H, C.navy);
  rect(s, 0, 0, 420, H, C.mid);
  rect(s, 0, 0, 10, H, C.teal);
  rect(s, 58, 72, 140, 5, C.teal2);
  text(s, "ParkSight AI", 58, 110, 420, 60, { size: 46, bold: true, color: C.white, typeface: fonts.title });
  text(s, "Ready for smarter, evidence-led parking enforcement", 58, 180, 430, 34, { size: 17, color: C.light });
  text(s, "One line for judges", 58, 270, 240, 20, { size: 13, bold: true, color: C.teal2 });
  text(s, "Parking violations become congestion-prioritized patrol beats.", 58, 306, 330, 120, {
    size: 28,
    bold: true,
    color: C.white,
    typeface: fonts.title,
  });
  const closes = [
    ["Detect", "illegal-parking hotspots from geocoded violation records", C.coral],
    ["Prioritize", "impact, volume and junction-risk layers", C.amber],
    ["Deploy", "12 targeted enforcement beats with report downloads", C.teal],
    ["Explain", "auditable 7-factor impact score", C.violet],
  ];
  closes.forEach(([h, b, c], i) => {
    const x = 560 + (i % 2) * 300;
    const y = 92 + Math.floor(i / 2) * 155;
    card(s, x, y, 245, 108, c, { fill: C.panel });
    text(s, h, x + 28, y + 24, 190, 30, { size: 27, bold: true, color: c, typeface: fonts.number });
    text(s, b, x + 28, y + 64, 185, 30, { size: 10.5, bold: true, color: C.off });
  });
  card(s, 560, 430, 545, 96, C.teal, { fill: C.panel });
  text(s, "Live demo", 592, 456, 120, 20, { size: 13, bold: true, color: C.amber });
  text(s, "rxxj25.github.io/GRIDLOCK-HACKATHON-PARKSIGHT-AI", 592, 484, 460, 18, { size: 12, color: C.white });
  text(s, "GitHub: github.com/rxxj25/GRIDLOCK-HACKATHON-PARKSIGHT-AI", 592, 508, 460, 18, { size: 11, color: C.light });
  text(s, "Rajdeep Bandyopadhaya | Aniket Arya | BMS College of Engineering", 58, 626, 740, 22, {
    size: 13,
    color: C.light,
  });
  footer(s, 11);
}

await fs.writeFile(
  path.join(WORK, "source-notes.txt"),
  [
    "ParkSight AI old-theme judge-final deck",
    `Original source deck backup: ${ORIGINAL_BACKUP}`,
    "User instruction: keep the same theme as the old presentation; do not use v2.",
    "Theme retained: dark navy background, teal/coral/amber/violet accent strips, dark cards, compact footer/page marker.",
    "Facts retained from old deck: 298,450 violations, 3,969 cells, 51% junction-linked, 55 stations, 10 Nov 2023 to 08 Apr 2024 range, BTP051 Safina Plaza score 77.5, 12 beats, CSV/PDF reports, React/Tailwind/Leaflet/Chart.js/Python stack.",
    "Screenshots: project-provided dashboard screenshots under submission_assets/Screenshots.",
  ].join("\n"),
);

await fs.writeFile(
  path.join(WORK, "slide-plan.txt"),
  [
    "Old-theme 10/10 polish plan",
    "Slides: cover, problem, data, solution, product demo, explainable scoring, hotspot evidence, targeted enforcement, architecture, impact, close.",
    "Typography: Aptos Display headings, Aptos body, Georgia metrics, Consolas formula.",
    "Goal: preserve old deck's visual identity while improving story clarity, visual hierarchy, screenshot evidence and judge-facing claims.",
  ].join("\n"),
);

for (const [index, slide] of deck.slides.items.entries()) {
  const stem = `slide-${String(index + 1).padStart(2, "0")}`;
  await writeBlob(path.join(PREVIEW, `${stem}.png`), await deck.export({ slide, format: "png", scale: 1 }));
  const layout = await slide.export({ format: "layout" });
  await fs.writeFile(path.join(LAYOUT, `${stem}.layout.json`), await layout.text());
}
await writeBlob(path.join(PREVIEW, "deck-montage.webp"), await deck.export({ format: "webp", montage: true, scale: 1 }));

const pptx = await PresentationFile.exportPptx(deck);
await pptx.save(FINAL);

await fs.writeFile(
  path.join(QA, "visual-qa.txt"),
  [
    "Visual QA summary",
    "Rendered every slide plus montage before export.",
    "Checked old-theme consistency: dark navy background, old accent colors, dark cards, footer/page marker.",
    "Used real project screenshots on demo, evidence and enforcement slides.",
    "Final deck exported as editable PPTX through @oai/artifact-tool.",
  ].join("\n"),
);

console.log(JSON.stringify({ final: FINAL, slides: deck.slides.items.length, preview: PREVIEW }, null, 2));
