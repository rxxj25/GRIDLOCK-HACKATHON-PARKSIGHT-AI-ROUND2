import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const artifactToolPath =
  "C:\\Users\\Rajdeep\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules\\@oai\\artifact-tool\\dist\\artifact_tool.mjs";
const { Presentation, PresentationFile } = await import(pathToFileURL(artifactToolPath).href);

const ROOT = "D:\\FLIPKART_GRIDLOCK";
const OUT_DIR = path.join(ROOT, "outputs");
const WORK = path.join(ROOT, "work", "presentations", "parksight-judge-final", "tmp");
const PREVIEW = path.join(WORK, "preview");
const LAYOUT = path.join(WORK, "layout");
const QA = path.join(WORK, "qa");

const FINAL = path.join(OUT_DIR, "ParkSight_AI_Pitch_Deck.pptx");
const BACKUP = path.join(OUT_DIR, "ParkSight_AI_Pitch_Deck_original_backup.pptx");

const shots = {
  overview: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002331.png"),
  evidence: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002346.png"),
  volume: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002445.png"),
  junction: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002457.png"),
  method: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002430.png"),
  hotspots: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002412.png"),
  plan: path.join(ROOT, "submission_assets", "Screenshots", "Screenshot 2026-06-17 002422.png"),
};

const C = {
  ink: "#111827",
  charcoal: "#1F2937",
  teal: "#007D73",
  mint: "#DDF5F0",
  coral: "#E5484D",
  amber: "#E5A018",
  blue: "#2F6FED",
  violet: "#6857C2",
  paper: "#F7F9FB",
  white: "#FFFFFF",
  line: "#D7DEE8",
  soft: "#EEF4F7",
  muted: "#5F6B7A",
  slate: "#334155",
};

const fonts = {
  display: "Aptos Display",
  body: "Aptos",
  number: "Georgia",
};

const W = 1280;
const H = 720;
const page = { left: 58, top: 42, width: 1164, height: 616 };

await fs.mkdir(OUT_DIR, { recursive: true });
await fs.mkdir(PREVIEW, { recursive: true });
await fs.mkdir(LAYOUT, { recursive: true });
await fs.mkdir(QA, { recursive: true });

try {
  await fs.access(FINAL);
  try {
    await fs.access(BACKUP);
  } catch {
    await fs.copyFile(FINAL, BACKUP);
  }
} catch {
  // No existing final to back up.
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function imageBlob(filePath) {
  const bytes = await fs.readFile(filePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

function box(slide, x, y, w, h, fill = C.white, opts = {}) {
  return slide.shapes.add({
    geometry: opts.geometry || "roundRect",
    name: opts.name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: opts.line || { style: "solid", fill: opts.stroke || C.line, width: opts.lineWidth ?? 1 },
    borderRadius: opts.radius || "rounded-lg",
    shadow: opts.shadow,
  });
}

function rect(slide, x, y, w, h, fill, opts = {}) {
  return slide.shapes.add({
    geometry: opts.geometry || "rect",
    name: opts.name,
    position: { left: x, top: y, width: w, height: h },
    fill,
    line: opts.line || { style: "solid", fill: opts.stroke || fill, width: opts.lineWidth ?? 0 },
    shadow: opts.shadow,
  });
}

function text(slide, value, x, y, w, h, style = {}) {
  const s = slide.shapes.add({
    geometry: "textbox",
    name: style.name,
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  s.text = value;
  s.text.style = {
    typeface: style.typeface || fonts.body,
    fontSize: style.size || 18,
    bold: style.bold || false,
    italic: style.italic || false,
    color: style.color || C.ink,
    alignment: style.align || "left",
  };
  return s;
}

function title(slide, eyebrow, heading, sub, accent = C.teal, dark = false) {
  const main = dark ? C.white : C.ink;
  const body = dark ? "#D6DEE8" : C.muted;
  text(slide, eyebrow, page.left, 36, 520, 24, {
    size: 12,
    bold: true,
    color: accent,
    typeface: fonts.body,
  });
  text(slide, heading, page.left, 66, 870, 78, {
    size: 34,
    bold: true,
    color: main,
    typeface: fonts.display,
  });
  if (sub) {
    text(slide, sub, page.left, 132, 860, 46, { size: 16, color: body });
  }
  rect(slide, page.left, 186, 150, 4, accent);
}

function footer(slide, n, dark = false) {
  const c = dark ? "#C7D0DA" : C.muted;
  text(slide, "ParkSight AI | Flipkart Gridlock Challenge", 58, 674, 420, 18, {
    size: 10,
    color: c,
  });
  text(slide, String(n).padStart(2, "0"), 1162, 670, 50, 26, {
    size: 15,
    bold: true,
    color: c,
    align: "right",
    typeface: fonts.number,
  });
}

function metric(slide, x, y, w, h, value, label, color, dark = false) {
  const valueSize = w < 170 ? 24 : value.length > 10 ? 24 : value.length > 7 ? 28 : 34;
  const labelSize = w < 220 ? 10 : 13;
  box(slide, x, y, w, h, dark ? "#16212C" : C.white, {
    stroke: dark ? "#2F3B48" : C.line,
    shadow: dark ? "shadow-sm" : "shadow-md",
  });
  rect(slide, x, y, w, 5, color);
  text(slide, value, x + 20, y + 18, w - 36, 46, {
    size: valueSize,
    bold: true,
    color,
    typeface: fonts.number,
  });
  text(slide, label, x + 20, y + 68, w - 36, h - 78, {
    size: labelSize,
    bold: true,
    color: dark ? "#E7EDF3" : C.slate,
  });
}

function bullet(slide, x, y, w, label, body, color) {
  rect(slide, x, y + 8, 10, 10, color, { geometry: "ellipse" });
  text(slide, label, x + 24, y, w - 24, 22, { size: 14, bold: true, color: C.ink });
  if (body) {
    text(slide, body, x + 24, y + 24, w - 24, 40, { size: 12, color: C.muted });
  }
}

async function image(slide, filePath, x, y, w, h, alt, opts = {}) {
  slide.images.add({
    blob: await imageBlob(filePath),
    contentType: "image/png",
    alt,
    fit: opts.fit || "cover",
    position: { left: x, top: y, width: w, height: h },
    geometry: "roundRect",
    borderRadius: opts.radius || "rounded-lg",
  });
}

const deck = Presentation.create({ slideSize: { width: W, height: H } });

// 1. Cover
{
  const s = deck.slides.add();
  s.background.fill = C.ink;
  rect(s, 0, 0, W, H, C.ink);
  rect(s, 0, 0, 420, H, "#0B5F59");
  rect(s, 420, 0, 860, H, "#111827");
  rect(s, 46, 42, 130, 5, C.amber);
  text(s, "FLIPKART GRIDLOCK CHALLENGE | PROTOTYPE ROUND", 58, 58, 500, 22, {
    size: 12,
    bold: true,
    color: "#BEEBE4",
  });
  text(s, "ParkSight", 58, 112, 520, 74, {
    size: 58,
    bold: true,
    color: C.white,
    typeface: fonts.display,
  });
  text(s, "AI", 58, 178, 160, 74, {
    size: 58,
    bold: true,
    color: C.amber,
    typeface: fonts.display,
  });
  text(s, "Congestion-aware parking enforcement intelligence", 58, 280, 460, 54, {
    size: 20,
    color: "#E8F3F1",
    italic: true,
  });
  text(
    s,
    "Detect hotspots. Quantify traffic-flow risk. Deploy police beats where violations hurt movement most.",
    58,
    352,
    500,
    70,
    { size: 17, color: "#DDE7EF" },
  );
  metric(s, 610, 86, 230, 126, "298K", "violations analyzed", C.amber, true);
  metric(s, 868, 86, 230, 126, "3,969", "urban cells scored", C.teal, true);
  metric(s, 610, 246, 230, 126, "51%", "junction-linked cases", C.coral, true);
  metric(s, 868, 246, 230, 126, "12", "deployable beats", C.violet, true);
  box(s, 610, 440, 488, 104, "#16212C", { stroke: "#2F3B48" });
  text(s, "Live dashboard + explainable Parking Impact Index", 636, 462, 430, 28, {
    size: 18,
    bold: true,
    color: C.white,
  });
  text(s, "React | Tailwind | Leaflet | Chart.js | Python ETL", 636, 498, 430, 24, {
    size: 13,
    color: "#B8C7D3",
  });
  text(s, "Rajdeep Bandyopadhaya | Aniket Arya | BMS College of Engineering", 58, 620, 790, 24, {
    size: 14,
    color: "#DDE7EF",
  });
  text(s, "Demo: rxxj25.github.io/GRIDLOCK-HACKATHON-PARKSIGHT-AI", 58, 648, 760, 22, {
    size: 12,
    color: "#BEEBE4",
  });
  footer(s, 1, true);
}

// 2. Problem
{
  const s = deck.slides.add();
  s.background.fill = C.paper;
  title(s, "THE OPERATIONAL GAP", "Illegal parking is visible. Its traffic impact is not.", "Judges should see the gap in one sentence: current enforcement reacts to tickets; ParkSight prioritizes congestion pain.", C.coral);
  metric(s, 58, 226, 250, 130, "INR 3,700 Cr", "annual traffic loss in Bengaluru", C.coral);
  metric(s, 330, 226, 250, 130, "39.6%", "violations during peak movement hours", C.amber);
  metric(s, 602, 226, 250, 130, "50.5%", "cases at junctions and crossings", C.teal);
  metric(s, 874, 226, 250, 130, "55", "police stations in the record", C.violet);
  box(s, 58, 410, 512, 184, C.white, { shadow: "shadow-sm" });
  text(s, "What happens today", 84, 434, 300, 30, { size: 20, bold: true, color: C.ink, typeface: fonts.display });
  bullet(s, 88, 480, 420, "Reactive patrols", "teams move after congestion already forms", C.coral);
  bullet(s, 88, 542, 420, "Ticket-count heatmaps", "raw density misses traffic-flow damage", C.amber);
  box(s, 642, 410, 512, 184, C.white, { shadow: "shadow-sm" });
  text(s, "What judges should remember", 668, 434, 360, 30, { size: 20, bold: true, color: C.ink, typeface: fonts.display });
  text(s, "The winning insight is not more dots on a map. It is a ranked enforcement plan that tells officers where to go, when to go, and why that location matters.", 668, 482, 430, 76, { size: 16, color: C.slate });
  footer(s, 2);
}

// 3. Data Signal
{
  const s = deck.slides.add();
  s.background.fill = C.white;
  title(s, "DATA SIGNAL", "298,450 records become an enforcement intelligence layer", "The original deck's facts are kept, but the slide now makes credibility and traceability easier to scan.", C.teal);
  metric(s, 58, 226, 248, 124, "2,98,450", "violations analyzed", C.teal);
  metric(s, 324, 226, 248, 124, "3,969", "scored urban cells", C.blue);
  metric(s, 590, 226, 248, 124, "51%", "junction-linked", C.coral);
  metric(s, 856, 226, 248, 124, "55", "police stations", C.violet);
  box(s, 58, 408, 510, 150, C.paper);
  text(s, "Actual record range", 86, 430, 300, 26, { size: 14, bold: true, color: C.teal });
  text(s, "10 Nov 2023 to 08 Apr 2024", 86, 466, 410, 34, { size: 26, bold: true, color: C.ink, typeface: fonts.number });
  text(s, "The filename mentions Jan-May, but timestamps show the true analyzed range above.", 86, 514, 410, 34, { size: 12, color: C.muted });
  box(s, 642, 408, 510, 150, C.paper);
  text(s, "Top station and offence", 670, 430, 300, 26, { size: 14, bold: true, color: C.coral });
  text(s, "Upparpet | Wrong parking", 670, 466, 410, 34, { size: 22, bold: true, color: C.ink, typeface: fonts.number });
  text(s, "Station and offence mix supports targeted enforcement instead of generic citywide action.", 670, 514, 420, 34, { size: 12, color: C.muted });
  footer(s, 3);
}

// 4. Solution
{
  const s = deck.slides.add();
  s.background.fill = C.paper;
  title(s, "SOLUTION", "Command center for congestion-aware enforcement", "ParkSight turns police records into four operational outputs that a judge can immediately test.", C.blue);
  const cards = [
    ["Impact", "Prioritizes where illegal parking is most likely to hurt traffic flow.", C.coral],
    ["Volume", "Shows where raw violation density is highest for enforcement pressure.", C.blue],
    ["Junction", "Highlights intersections, crossings and signal-sensitive risks.", C.teal],
    ["Output", "Ranks hotspots, focuses stations, generates 12 beats and exports CSV/PDF.", C.violet],
  ];
  cards.forEach(([h, b, c], i) => {
    const x = 58 + (i % 2) * 584;
    const y = 232 + Math.floor(i / 2) * 170;
    box(s, x, y, 510, 136, C.white, { shadow: "shadow-sm" });
    rect(s, x, y, 8, 136, c);
    text(s, h, x + 28, y + 22, 220, 34, { size: 24, bold: true, color: c, typeface: fonts.display });
    text(s, b, x + 28, y + 66, 420, 46, { size: 15, color: C.slate });
  });
  box(s, 58, 600, 1094, 46, "#E8F7F4", { stroke: "#B8E3DB" });
  text(s, "Core pitch: ParkSight does not replace police judgment; it gives command teams a congestion-ranked starting point.", 84, 611, 1030, 22, { size: 15, bold: true, color: C.teal });
  footer(s, 4);
}

// 5. Product Demo
{
  const s = deck.slides.add();
  s.background.fill = C.ink;
  title(s, "PRODUCT DEMO", "The first screen tells judges exactly what to test", "Real dashboard screenshot from the submitted app, with the strongest judge-facing features called out.", C.amber, true);
  await image(s, shots.overview, 58, 208, 740, 416, "ParkSight AI overview dashboard screenshot", { fit: "contain" });
  const callouts = [
    ["Highest risk surfaced instantly", "BTP051 - Safina Plaza Junction, score 77.5", C.coral],
    ["Layered map intelligence", "Impact, Volume and Junction views answer different enforcement questions.", C.teal],
    ["Reports built in", "CSV for operations, PDF brief for commanders and judges.", C.amber],
  ];
  callouts.forEach(([h, b, c], i) => {
    const y = 214 + i * 132;
    box(s, 838, y, 340, 104, "#16212C", { stroke: "#2F3B48" });
    rect(s, 838, y, 6, 104, c);
    text(s, h, 862, y + 18, 284, 24, { size: 16, bold: true, color: C.white });
    text(s, b, 862, y + 50, 284, 34, { size: 12, color: "#C9D4DE" });
  });
  footer(s, 5, true);
}

// 6. Explainable AI
{
  const s = deck.slides.add();
  s.background.fill = C.white;
  title(s, "EXPLAINABLE AI / ML ANALYTICS", "Parking Impact Index: transparent, auditable scoring", "The dataset has violation events, not measured speed. So the model uses a congestion-impact proxy judges can inspect.", C.violet);
  const steps = [
    ["01", "Cluster", "Group violations into about 220m urban cells", C.teal],
    ["02", "Engineer", "Density, PCU, severity, junction and peak recurrence", C.blue],
    ["03", "Score", "Weighted 0-100 impact score for every cell", C.coral],
    ["04", "Deploy", "Convert top cells into patrol beats and windows", C.amber],
  ];
  steps.forEach(([n, h, b, c], i) => {
    const x = 58 + i * 292;
    box(s, x, 230, 244, 142, C.paper);
    text(s, n, x + 20, 248, 58, 36, { size: 28, bold: true, color: c, typeface: fonts.number });
    text(s, h, x + 20, 294, 180, 28, { size: 19, bold: true, color: C.ink, typeface: fonts.display });
    text(s, b, x + 20, 330, 190, 34, { size: 12, color: C.muted });
  });
  box(s, 108, 430, 1000, 92, "#151E29", { stroke: "#151E29" });
  text(s, "Score = 100 x (0.34 obstruction + 0.18 density + 0.15 junction + 0.13 arterial\n+ 0.10 peak + 0.06 recurrence + 0.04 severity)", 138, 448, 940, 44, { size: 14, bold: true, color: C.white, typeface: "Consolas" });
  text(s, "Why this wins trust: every weight is visible, challengeable and explainable to civic teams.", 138, 498, 840, 20, { size: 12, color: "#C9D4DE" });
  footer(s, 6);
}

// 7. Hotspot Evidence
{
  const s = deck.slides.add();
  s.background.fill = C.paper;
  title(s, "HOTSPOT EVIDENCE", "From tickets to traffic-flow risk", "The deck now shows evidence and action on the same slide so judges do not have to connect the dots themselves.", C.coral);
  await image(s, shots.evidence, 58, 218, 690, 392, "Impact-layer map and station burden chart");
  box(s, 790, 218, 360, 392, C.white, { shadow: "shadow-sm" });
  text(s, "#1 Hotspot", 820, 248, 160, 22, { size: 13, bold: true, color: C.coral });
  text(s, "BTP051 - Safina Plaza Junction", 820, 278, 280, 62, { size: 28, bold: true, color: C.ink, typeface: fonts.display });
  text(s, "Shivajinagar | Commercial market", 820, 346, 280, 24, { size: 14, color: C.muted });
  metric(s, 820, 394, 132, 92, "77.5", "impact score", C.coral);
  metric(s, 972, 394, 132, 92, "6,523", "violations", C.teal);
  text(s, "Recommended action", 820, 520, 260, 22, { size: 14, bold: true, color: C.ink });
  text(s, "Keep a junction-clearance unit on short-cycle patrol around 09:00-12:00 IST; dominant issue is wrong parking.", 820, 548, 292, 46, { size: 13, color: C.slate });
  footer(s, 7);
}

// 8. Targeted Enforcement
{
  const s = deck.slides.add();
  s.background.fill = C.ink;
  title(s, "TARGETED ENFORCEMENT", "A 12-beat plan that can be deployed, not just admired", "The judge takeaway: ParkSight ends in operational instructions, not only analytics.", C.teal, true);
  await image(s, shots.plan, 58, 210, 362, 430, "ParkSight targeted enforcement plan screenshot", { fit: "contain" });
  const beats = [
    ["1", "BTP051 - Safina Plaza", "Shivajinagar | 10:00-12:00 IST", "77.5"],
    ["2", "BTP040 - Elite Junction", "Upparpet | 08:00-10:00 IST", "77.0"],
    ["3", "BTP040 - Elite Junction", "Upparpet | 08:00-10:00 IST", "76.7"],
    ["4", "BTP051 - Safina Plaza", "Shivajinagar | 10:00-12:00 IST", "76.1"],
    ["5", "BTP082 - KR Market", "City Market | 00:00-02:00 IST", "74.9"],
  ];
  beats.forEach(([n, h, b, score], i) => {
    const y = 214 + i * 78;
    box(s, 480, y, 650, 58, "#16212C", { stroke: "#2F3B48" });
    rect(s, 500, y + 12, 34, 34, C.violet, { geometry: "roundRect" });
    text(s, n, 500, y + 18, 34, 18, { size: 15, bold: true, color: C.white, align: "center" });
    text(s, h, 552, y + 10, 330, 22, { size: 16, bold: true, color: C.white });
    text(s, b, 552, y + 34, 330, 18, { size: 11, color: "#B8C7D3" });
    text(s, score, 1044, y + 16, 60, 22, { size: 18, bold: true, color: C.teal, align: "right", typeface: fonts.number });
  });
  box(s, 480, 620, 650, 32, "#0B5F59", { stroke: "#0B5F59" });
  text(s, "Each beat includes location, impact score, patrol window and reason.", 508, 627, 590, 16, { size: 12, bold: true, color: C.white });
  footer(s, 8, true);
}

// 9. System Architecture
{
  const s = deck.slides.add();
  s.background.fill = C.white;
  title(s, "SYSTEM ARCHITECTURE", "Data pipeline to live decision dashboard", "Simple enough for hackathon judges to believe, complete enough for real deployment.", C.blue);
  const parts = [
    ["Raw challan CSV", "lat/lon, offence, station, timestamp, vehicle", C.coral],
    ["Python ETL", "cleaning, grid cells, weights, scores", C.amber],
    ["React dashboard", "map layers, charts, station focus", C.teal],
    ["Enforcement outputs", "12 beats, CSV, PDF brief", C.violet],
  ];
  parts.forEach(([h, b, c], i) => {
    const x = 72 + i * 292;
    box(s, x, 250, 232, 152, C.paper);
    rect(s, x, 250, 232, 5, c);
    text(s, h, x + 20, 282, 190, 28, { size: 18, bold: true, color: C.ink, typeface: fonts.display });
    text(s, b, x + 20, 326, 180, 44, { size: 12, color: C.muted });
    if (i < parts.length - 1) {
      text(s, ">", x + 244, 306, 36, 30, { size: 28, bold: true, color: c, align: "center" });
    }
  });
  box(s, 140, 488, 1000, 80, "#E8F7F4", { stroke: "#B8E3DB" });
  text(s, "Hosted online", 178, 512, 160, 22, { size: 14, bold: true, color: C.teal });
  text(s, "GitHub Pages live demo + GitHub repository with generated intelligence data", 348, 508, 620, 30, { size: 18, bold: true, color: C.ink });
  text(s, "No paid services, no hidden API dependency, no server required after build.", 348, 540, 620, 20, { size: 12, color: C.muted });
  footer(s, 9);
}

// 10. Impact
{
  const s = deck.slides.add();
  s.background.fill = C.paper;
  title(s, "IMPACT", "What changes for traffic police", "Before vs after is now sharper, with future scope framed as optional extension rather than the core claim.", C.amber);
  box(s, 72, 228, 480, 300, C.white, { shadow: "shadow-sm" });
  text(s, "Before", 104, 258, 180, 36, { size: 28, bold: true, color: C.coral, typeface: fonts.display });
  ["Reactive patrols", "Ticket-count heatmaps", "Manual zone priority", "Hard-to-share reports"].forEach((item, i) => {
    bullet(s, 108, 320 + i * 44, 340, item, "", C.coral);
  });
  box(s, 690, 228, 480, 300, C.white, { shadow: "shadow-sm" });
  text(s, "After ParkSight AI", 722, 258, 300, 36, { size: 28, bold: true, color: C.teal, typeface: fonts.display });
  ["Impact-ranked hotspots", "Station-wise navigation", "12-beat enforcement plan", "Downloadable PDF/CSV brief"].forEach((item, i) => {
    bullet(s, 726, 320 + i * 44, 340, item, "", C.teal);
  });
  rect(s, 594, 346, 64, 44, C.ink, { geometry: "roundRect" });
  text(s, "VS", 594, 357, 64, 20, { size: 18, bold: true, color: C.white, align: "center" });
  box(s, 130, 588, 980, 42, C.white, { stroke: C.line });
  text(s, "Future scope: integrate live CCTV/YOLO illegal-parking detection and traffic-speed labels for supervised congestion prediction.", 158, 598, 930, 18, { size: 12, color: C.muted });
  footer(s, 10);
}

// 11. Close
{
  const s = deck.slides.add();
  s.background.fill = C.ink;
  rect(s, 0, 0, W, H, C.ink);
  rect(s, 0, 0, 470, H, "#0B5F59");
  text(s, "ParkSight AI", 58, 86, 440, 62, { size: 48, bold: true, color: C.white, typeface: fonts.display });
  text(s, "Ready for smarter, evidence-led parking enforcement", 58, 156, 500, 38, { size: 18, color: "#DDE7EF" });
  rect(s, 58, 224, 122, 5, C.amber);
  text(s, "One line for judges", 58, 260, 280, 24, { size: 14, bold: true, color: "#BEEBE4" });
  text(s, "ParkSight AI turns parking violations into congestion-prioritized patrol beats.", 58, 292, 360, 78, { size: 24, bold: true, color: C.white, typeface: fonts.display });
  metric(s, 570, 88, 210, 114, "Detect", "illegal-parking hotspots from geocoded violation records", C.coral, true);
  metric(s, 820, 88, 210, 114, "Prioritize", "impact, volume and junction-risk layers", C.amber, true);
  metric(s, 570, 238, 210, 114, "Deploy", "12 targeted enforcement beats with report downloads", C.teal, true);
  metric(s, 820, 238, 210, 114, "Explain", "auditable 7-factor impact score", C.violet, true);
  box(s, 570, 432, 540, 98, "#16212C", { stroke: "#2F3B48" });
  text(s, "Live demo", 600, 454, 120, 24, { size: 14, bold: true, color: C.amber });
  text(s, "rxxj25.github.io/GRIDLOCK-HACKATHON-PARKSIGHT-AI", 600, 482, 470, 20, { size: 13, color: C.white });
  text(s, "GitHub: github.com/rxxj25/GRIDLOCK-HACKATHON-PARKSIGHT-AI", 600, 506, 470, 18, { size: 12, color: "#B8C7D3" });
  text(s, "Rajdeep Bandyopadhaya | Aniket Arya | BMS College of Engineering", 58, 626, 760, 24, { size: 14, color: "#DDE7EF" });
  footer(s, 11, true);
}

await fs.writeFile(
  path.join(WORK, "source-notes.txt"),
  [
    "ParkSight AI judge-final deck source notes",
    "Source deck: D:\\FLIPKART_GRIDLOCK\\outputs\\ParkSight_AI_Pitch_Deck.pptx",
    "Instruction: use original ParkSight_AI_Pitch_Deck, not v2.",
    "Core facts retained from original deck: 298,450 violations, 3,969 urban cells, 51% junction-linked, 55 police stations, actual record range 10 Nov 2023 to 08 Apr 2024, top hotspot BTP051 Safina Plaza Junction, impact score 77.5, 12 enforcement beats, React/Tailwind/Leaflet/Chart.js/Python stack.",
    "Visual assets: user/project-provided screenshots under submission_assets/Screenshots.",
  ].join("\n"),
);

await fs.writeFile(
  path.join(WORK, "slide-plan.txt"),
  [
    "Judge-final redesign plan",
    "Palette: ink #111827, teal #007D73, coral #E5484D, amber #E5A018, blue #2F6FED, violet #6857C2, paper #F7F9FB.",
    "Typography: Aptos Display for headings, Aptos for body, Georgia for numeric metrics, Consolas for formula.",
    "Story: cover -> problem -> data -> solution -> demo -> explainable model -> evidence -> enforcement -> architecture -> impact -> close.",
    "Design goal: strong first read, high contrast, real dashboard visuals, less clutter, judge-ready callouts.",
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
    "Rendered all 11 slides to PNG plus montage.",
    "Checked for obvious overflow in fixed-position title, metric, card and footer areas.",
    "Used real product screenshots for demo, method, evidence, map layers and enforcement plan.",
    "Final PPTX exported through @oai/artifact-tool as editable shapes, text and images.",
  ].join("\n"),
);

console.log(JSON.stringify({ final: FINAL, backup: BACKUP, slides: deck.slides.items.length, preview: PREVIEW }, null, 2));
