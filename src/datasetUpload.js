const formatter = new Intl.NumberFormat("en-IN");
const IST_OFFSET_MS = 5.5 * 60 * 60 * 1000;

const vehicleWeights = {
  SCOOTER: 0.5,
  "MOTOR CYCLE": 0.5,
  "TWO WHEELER": 0.5,
  CAR: 1,
  JEEP: 1.2,
  "PASSENGER AUTO": 1.2,
  AUTO: 1.2,
  TAXI: 1.1,
  BUS: 3.2,
  LORRY: 3,
  TRUCK: 3,
  TEMPO: 2,
  VAN: 1.4,
};

export function buildDatasetFromUpload(text, file) {
  const trimmed = text.trim();
  if (!trimmed) throw new Error("The uploaded file is empty.");
  if (trimmed.startsWith("{") || trimmed.startsWith("[")) {
    const parsed = JSON.parse(trimmed);
    const data = Array.isArray(parsed) ? buildDataFromRows(parsed, file.name) : parsed;
    return normalizeUploadedDataset(data, file.name);
  }
  return buildDataFromRows(parseCsvRows(text), file.name);
}

function normalizeUploadedDataset(data, fileName) {
  if (!data?.hotspots?.length) throw new Error("JSON must contain a non-empty hotspots array, or upload raw CSV records.");
  const hotspots = data.hotspots
    .map((hotspot, index) => ({
      ...hotspot,
      id: hotspot.id || `HS-${String(index + 1).padStart(4, "0")}`,
      rank: hotspot.rank || index + 1,
      lat: Number(hotspot.lat),
      lng: Number(hotspot.lng),
      violations: Number(hotspot.violations) || 0,
      impactScore: Number(hotspot.impactScore) || 0,
      weightedObstruction: Number(hotspot.weightedObstruction) || Number(hotspot.violations) || 0,
      pcu: Number(hotspot.pcu) || Number(hotspot.weightedObstruction) || Number(hotspot.violations) || 0,
      activeDays: Number(hotspot.activeDays) || 1,
      peakShare: Number(hotspot.peakShare) || 0,
      junctionShare: Number(hotspot.junctionShare) || 0,
      arterialShare: Number(hotspot.arterialShare) || 0,
      approvedShare: Number(hotspot.approvedShare) || 0,
      scitaShare: Number(hotspot.scitaShare) || 0,
      priority: hotspot.priority || priorityForScore(Number(hotspot.impactScore) || 0),
      station: hotspot.station || "Unknown station",
      area: hotspot.area || hotspot.location || "Uploaded hotspot",
      topViolation: hotspot.topViolation || "UNKNOWN",
      topVehicle: hotspot.topVehicle || "UNKNOWN",
      topHours: hotspot.topHours?.length ? hotspot.topHours : [{ hour: 10, count: Number(hotspot.violations) || 0 }],
      recommendation: hotspot.recommendation || recommendationForHotspot(hotspot),
    }))
    .filter((hotspot) => Number.isFinite(hotspot.lat) && Number.isFinite(hotspot.lng) && hotspot.violations > 0)
    .sort((a, b) => b.impactScore - a.impactScore)
    .map((hotspot, index) => ({ ...hotspot, rank: index + 1 }));
  if (!hotspots.length) throw new Error("No valid latitude/longitude hotspots were found.");
  return {
    summary: {
      ...buildSummaryFromHotspots(hotspots, fileName),
      ...data.summary,
      totalViolations: data.summary?.totalViolations || sum(hotspots, "violations"),
      generatedFrom: fileName,
      modelVersion: data.summary?.modelVersion || "Parking Impact Index v1.0",
    },
    hotspots,
    heatmap: data.heatmap?.length ? data.heatmap : buildHeatmapFromHotspots(hotspots),
    stations: data.stations?.length ? data.stations : buildStationsFromHotspots(hotspots),
    plates: data.plates?.length ? data.plates : [],
    enforcementPlan: data.enforcementPlan?.length ? data.enforcementPlan : buildEnforcementPlan(hotspots),
    charts: {
      hours: data.charts?.hours?.length ? data.charts.hours : buildHourChartFromHotspots(hotspots),
      vehicles: data.charts?.vehicles?.length ? data.charts.vehicles : buildVehicleChartFromHotspots(hotspots),
      stations: data.charts?.stations?.length ? data.charts.stations : buildStationChartFromHotspots(hotspots),
    },
    method: data.method || buildMethodNote(fileName),
  };
}

function buildDataFromRows(rows, fileName) {
  if (!rows.length) throw new Error("No rows were found in the uploaded dataset.");
  validateParkingSchema(rows);
  const groups = new Map();
  const hourCounts = Array.from({ length: 24 }, (_, hour) => ({ hour, count: 0 }));
  const vehicleCounts = new Map();
  const plateCounts = new Map();
  let validRows = 0;
  let firstTimestamp = Infinity;
  let lastTimestamp = -Infinity;

  rows.forEach((row) => {
    const lat = parseCoordinate(readField(row, ["latitude", "lat"]));
    const lng = parseCoordinate(readField(row, ["longitude", "lng", "lon", "long"]));
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return;
    validRows += 1;
    const station = titleCase(readField(row, ["policestation", "station", "centername"]) || "Unknown station");
    const location = readField(row, ["location", "address", "area"]) || `${station} hotspot`;
    const junctionName = readField(row, ["junctionname", "junction"]);
    const vehicle = cleanToken(readField(row, ["updatedvehicletype", "vehicletype", "vehicle"]) || "UNKNOWN");
    const plate = cleanPlate(readField(row, ["updatedvehiclenumber", "vehiclenumber", "numberplate", "registrationnumber", "vehicleid"]));
    const violations = parseListValue(readField(row, ["violationtype", "violation", "description", "offence"]));
    const topViolation = cleanToken(violations[0] || "UNKNOWN");
    const time = parseTrafficTime(readField(row, ["createddatetime", "createdat", "timestamp", "date"]));
    if (time) {
      firstTimestamp = Math.min(firstTimestamp, time.ms);
      lastTimestamp = Math.max(lastTimestamp, time.ms);
    }
    const hour = time ? time.hour : 10;
    const dateKey = time ? time.dateKey : "unknown";
    const groupKey = `${lat.toFixed(3)},${lng.toFixed(3)},${station}`;
    const group = groups.get(groupKey) || {
      lat: Number(lat.toFixed(3)),
      lng: Number(lng.toFixed(3)),
      station,
      location,
      junctionName,
      count: 0,
      pcu: 0,
      dates: new Set(),
      hourCounts: new Map(),
      vehicleCounts: new Map(),
      violationCounts: new Map(),
      junctionCount: 0,
      arterialCount: 0,
      approvedCount: 0,
      scitaCount: 0,
    };
    group.count += 1;
    group.pcu += vehicleWeights[vehicle] || 1;
    group.dates.add(dateKey);
    group.hourCounts.set(hour, (group.hourCounts.get(hour) || 0) + 1);
    group.vehicleCounts.set(vehicle, (group.vehicleCounts.get(vehicle) || 0) + 1);
    group.violationCounts.set(topViolation, (group.violationCounts.get(topViolation) || 0) + 1);
    if (isJunctionLinked(junctionName, topViolation, location)) group.junctionCount += 1;
    if (isArterialLinked(topViolation, location)) group.arterialCount += 1;
    if (/approved/i.test(readField(row, ["validationstatus", "status"]))) group.approvedCount += 1;
    if (/true|yes|1/i.test(readField(row, ["datasenttoscita", "scita"]))) group.scitaCount += 1;
    groups.set(groupKey, group);
    hourCounts[hour].count += 1;
    vehicleCounts.set(vehicle, (vehicleCounts.get(vehicle) || 0) + 1);
    if (plate) {
      const plateEntry = plateCounts.get(plate) || {
        plate,
        count: 0,
        pcu: 0,
        latestSeenMs: 0,
      };
      plateEntry.count += 1;
      plateEntry.pcu += vehicleWeights[vehicle] || 1;
      if (time) plateEntry.latestSeenMs = Math.max(plateEntry.latestSeenMs, time.ms);
      plateCounts.set(plate, plateEntry);
    }
  });

  if (!validRows || !groups.size) throw new Error("No usable latitude/longitude records were found.");

  const rawHotspots = [...groups.values()];
  const maxCount = Math.max(...rawHotspots.map((item) => item.count), 1);
  const maxPcu = Math.max(...rawHotspots.map((item) => item.pcu), 1);
  const maxDays = Math.max(...rawHotspots.map((item) => item.dates.size), 1);
  const scoredHotspots = rawHotspots
    .map((group, index) => {
      const topHours = topEntries(group.hourCounts, 4).map(([hour, count]) => ({ hour: Number(hour), count }));
      const peakCount = [...group.hourCounts].reduce((total, [hour, count]) => total + (isPeakHour(Number(hour)) ? count : 0), 0);
      const junctionShare = group.junctionCount / group.count;
      const arterialShare = group.arterialCount / group.count;
      const peakShare = peakCount / group.count;
      const density = Math.pow(group.count / maxCount, 0.45);
      const pcuShare = Math.pow(group.pcu / maxPcu, 0.5);
      const activeShare = group.dates.size / maxDays;
      const impactScore = round1((density * 34) + (pcuShare * 22) + (junctionShare * 18) + (arterialShare * 10) + (peakShare * 10) + (activeShare * 6));
      const hotspot = {
        id: `UP-${String(index + 1).padStart(4, "0")}`,
        lat: group.lat,
        lng: group.lng,
        impactScore,
        priority: priorityForScore(impactScore),
        violations: group.count,
        weightedObstruction: round1(group.pcu),
        pcu: round1(group.pcu),
        activeDays: group.dates.size,
        peakShare: round3(peakShare),
        junctionShare: round3(junctionShare),
        arterialShare: round3(arterialShare),
        approvedShare: round3(group.approvedCount / group.count),
        scitaShare: round3(group.scitaCount / group.count),
        station: group.station,
        area: deriveArea(group.location, group.junctionName, group.station),
        placeType: classifyPlace(group.location, group.junctionName),
        topViolation: topEntries(group.violationCounts, 1)[0]?.[0] || "UNKNOWN",
        topVehicle: topEntries(group.vehicleCounts, 1)[0]?.[0] || "UNKNOWN",
        topHours,
      };
      return { ...hotspot, recommendation: recommendationForHotspot(hotspot) };
    })
    .sort((a, b) => b.impactScore - a.impactScore || b.violations - a.violations)
    .map((hotspot, index) => ({ ...hotspot, id: `HS-${String(index + 1).padStart(4, "0")}`, rank: index + 1 }));
  const hotspots = scoredHotspots
    .slice(0, 220)
    .map((hotspot, index) => ({ ...hotspot, rank: index + 1 }));

  return {
    summary: {
      ...buildSummaryFromHotspots(scoredHotspots, fileName),
      totalViolations: validRows,
      dateRange: buildDateRangeFromMs(firstTimestamp, lastTimestamp),
      generatedFrom: fileName,
    },
    hotspots,
    heatmap: buildHeatmapFromHotspots(hotspots),
    stations: buildStationsFromHotspots(scoredHotspots),
    plates: buildPlateList(plateCounts, rows),
    enforcementPlan: buildEnforcementPlan(hotspots),
    charts: {
      hours: hourCounts,
      vehicles: [...vehicleCounts].sort((a, b) => b[1] - a[1]).slice(0, 12).map(([name, count]) => ({ name, count })),
      stations: buildStationChartFromHotspots(scoredHotspots),
    },
    method: buildMethodNote(fileName),
  };
}

function validateParkingSchema(rows) {
  const columns = new Set(Object.keys(rows[0] || {}));
  const hasLocation = hasAnyColumn(columns, ["latitude", "lat"]) && hasAnyColumn(columns, ["longitude", "lng", "lon", "long"]);
  const evidenceGroups = [
    ["violationtype", "violation", "description", "offence", "offencecode"],
    ["vehicletype", "updatedvehicletype", "vehicle"],
    ["policestation", "station", "centername"],
    ["createddatetime", "createdat", "timestamp", "date"],
    ["junctionname", "junction", "datasenttoscita", "validationstatus"],
  ];
  const matchedGroups = evidenceGroups.filter((aliases) => hasAnyColumn(columns, aliases)).length;
  if (!hasLocation) throw new Error("Dataset must include latitude and longitude columns.");
  if (matchedGroups < 2) {
    throw new Error("Dataset must include parking/enforcement fields such as violation_type, vehicle_type, police_station, or created_datetime.");
  }
}

function hasAnyColumn(columns, aliases) {
  return aliases.some((alias) => columns.has(normalizeColumn(alias)));
}

function parseCsvRows(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];
    if (char === '"') {
      if (quoted && next === '"') {
        cell += '"';
        index += 1;
      } else {
        quoted = !quoted;
      }
    } else if (char === "," && !quoted) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell);
      if (row.some((value) => value.trim())) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }
  row.push(cell);
  if (row.some((value) => value.trim())) rows.push(row);
  const headers = rows.shift()?.map(normalizeColumn) || [];
  return rows.map((values) => {
    const item = {};
    headers.forEach((header, index) => {
      if (!header) return;
      item[header] = values[index]?.trim() || "";
    });
    return item;
  });
}

function buildSummaryFromHotspots(hotspots, fileName) {
  const bounds = hotspots.reduce(
    (acc, hotspot) => ({
      minLat: Math.min(acc.minLat, hotspot.lat),
      maxLat: Math.max(acc.maxLat, hotspot.lat),
      minLng: Math.min(acc.minLng, hotspot.lng),
      maxLng: Math.max(acc.maxLng, hotspot.lng),
    }),
    { minLat: Infinity, maxLat: -Infinity, minLng: Infinity, maxLng: -Infinity },
  );
  const stationNames = new Set(hotspots.map((hotspot) => hotspot.station));
  return {
    totalViolations: sum(hotspots, "violations"),
    dateRange: { start: "Uploaded", end: "Now" },
    bounds,
    cellsAnalyzed: hotspots.length,
    stations: stationNames.size,
    junctionLinkedShare: weightedShare(hotspots, "junctionShare"),
    approvedShare: weightedShare(hotspots, "approvedShare"),
    peakHourShare: weightedShare(hotspots, "peakShare"),
    topStation: buildStationChartFromHotspots(hotspots)[0]?.name || "Unknown station",
    topViolation: topWeightedToken(hotspots, "topViolation"),
    modelVersion: "Parking Impact Index v1.0",
    generatedFrom: fileName,
  };
}

function buildHeatmapFromHotspots(hotspots) {
  const maxCount = Math.max(...hotspots.map((hotspot) => hotspot.violations), 1);
  return hotspots.map((hotspot) => [hotspot.lat, hotspot.lng, round3(Math.max(0.12, hotspot.impactScore / 100)), hotspot.violations || maxCount]);
}

function buildStationsFromHotspots(hotspots) {
  const byStation = new Map();
  hotspots.forEach((hotspot) => {
    const current = byStation.get(hotspot.station) || { station: hotspot.station, cases: 0, impactScore: 0, criticalHotspots: 0, junctionNumerator: 0, arterialNumerator: 0, peakNumerator: 0, latNumerator: 0, lngNumerator: 0, vehicleCounts: new Map(), violationCounts: new Map() };
    current.cases += hotspot.violations;
    current.impactScore = Math.max(current.impactScore, hotspot.impactScore);
    current.criticalHotspots += hotspot.priority === "Critical" ? 1 : 0;
    current.junctionNumerator += hotspot.junctionShare * hotspot.violations;
    current.arterialNumerator += hotspot.arterialShare * hotspot.violations;
    current.peakNumerator += hotspot.peakShare * hotspot.violations;
    current.latNumerator += hotspot.lat * hotspot.violations;
    current.lngNumerator += hotspot.lng * hotspot.violations;
    current.vehicleCounts.set(hotspot.topVehicle, (current.vehicleCounts.get(hotspot.topVehicle) || 0) + hotspot.violations);
    current.violationCounts.set(hotspot.topViolation, (current.violationCounts.get(hotspot.topViolation) || 0) + hotspot.violations);
    byStation.set(hotspot.station, current);
  });
  return [...byStation.values()].map((station) => ({
    station: station.station,
    cases: station.cases,
    impactScore: round1(station.impactScore),
    criticalHotspots: station.criticalHotspots,
    junctionShare: round3(station.junctionNumerator / station.cases),
    arterialShare: round3(station.arterialNumerator / station.cases),
    peakShare: round3(station.peakNumerator / station.cases),
    topViolation: topEntries(station.violationCounts, 1)[0]?.[0] || "UNKNOWN",
    topVehicle: topEntries(station.vehicleCounts, 1)[0]?.[0] || "UNKNOWN",
    lat: round5(station.latNumerator / station.cases),
    lng: round5(station.lngNumerator / station.cases),
  })).sort((a, b) => b.impactScore - a.impactScore);
}

function buildEnforcementPlan(hotspots) {
  return hotspots.slice(0, 12).map((hotspot, index) => ({
    rank: index + 1,
    hotspotId: hotspot.id,
    station: hotspot.station,
    area: hotspot.area,
    impactScore: hotspot.impactScore,
    window: hotspot.topHours?.[0] ? `${formatHourLabel(hotspot.topHours[0].hour)} field cycle` : "Peak field cycle",
    action: hotspot.recommendation,
    why: `${formatter.format(hotspot.violations)} cases, ${Math.round(hotspot.junctionShare * 100)}% junction exposure, ${Math.round(hotspot.arterialShare * 100)}% arterial obstruction`,
    lat: hotspot.lat,
    lng: hotspot.lng,
  }));
}

function buildHourChartFromHotspots(hotspots) {
  const counts = Array.from({ length: 24 }, (_, hour) => ({ hour, count: 0 }));
  hotspots.forEach((hotspot) => {
    (hotspot.topHours || []).forEach(({ hour, count }) => {
      if (Number.isInteger(hour) && counts[hour]) counts[hour].count += Number(count) || 0;
    });
  });
  return counts;
}

function buildVehicleChartFromHotspots(hotspots) {
  const counts = new Map();
  hotspots.forEach((hotspot) => counts.set(hotspot.topVehicle, (counts.get(hotspot.topVehicle) || 0) + hotspot.violations));
  return [...counts].sort((a, b) => b[1] - a[1]).slice(0, 12).map(([name, count]) => ({ name, count }));
}

function buildStationChartFromHotspots(hotspots) {
  const counts = new Map();
  hotspots.forEach((hotspot) => counts.set(hotspot.station, (counts.get(hotspot.station) || 0) + hotspot.violations));
  return [...counts].sort((a, b) => b[1] - a[1]).slice(0, 12).map(([name, count]) => ({ name, count }));
}

function buildPlateList(plateCounts, rows = []) {
  const topPlates = [...plateCounts.values()]
    .sort((a, b) => b.count - a.count || b.pcu - a.pcu)
    .slice(0, 160)
    .map((item) => ({
      plate: item.plate,
      count: item.count,
      pcu: round1(item.pcu),
      latestSeenMs: item.latestSeenMs,
      stationCounts: new Map(),
      vehicleCounts: new Map(),
      violationCounts: new Map(),
      hours: new Map(),
    }));
  const byPlate = new Map(topPlates.map((item) => [item.plate, item]));
  rows.forEach((row) => {
    const plate = cleanPlate(readField(row, ["updatedvehiclenumber", "vehiclenumber", "numberplate", "registrationnumber", "vehicleid"]));
    const item = byPlate.get(plate);
    if (!item) return;
    const station = titleCase(readField(row, ["policestation", "station", "centername"]) || "Unknown station");
    const vehicle = cleanToken(readField(row, ["updatedvehicletype", "vehicletype", "vehicle"]) || "UNKNOWN");
    const violation = cleanToken(parseListValue(readField(row, ["violationtype", "violation", "description", "offence"]))[0] || "UNKNOWN");
    const time = parseTrafficTime(readField(row, ["createddatetime", "createdat", "timestamp", "date"]));
    const hour = time ? time.hour : 10;
    item.stationCounts.set(station, (item.stationCounts.get(station) || 0) + 1);
    item.vehicleCounts.set(vehicle, (item.vehicleCounts.get(vehicle) || 0) + 1);
    item.violationCounts.set(violation, (item.violationCounts.get(violation) || 0) + 1);
    item.hours.set(hour, (item.hours.get(hour) || 0) + 1);
  });
  return topPlates.map((item) => ({
    plate: item.plate,
    count: item.count,
    pcu: item.pcu,
    station: topEntries(item.stationCounts, 1)[0]?.[0] || "Unknown station",
    vehicle: topEntries(item.vehicleCounts, 1)[0]?.[0] || "UNKNOWN",
    violation: topEntries(item.violationCounts, 1)[0]?.[0] || "UNKNOWN",
    peakHour: Number(topEntries(item.hours, 1)[0]?.[0] ?? 10),
    lastSeen: formatDateTimeShort(item.latestSeenMs),
  }));
}

function buildMethodNote(fileName) {
  return {
    index: "Impact score = density + PCU obstruction + junction exposure + arterial obstruction + peak recurrence + active-day persistence. Uploaded datasets are normalized into hotspot cells before scoring.",
    notes: [
      `Generated in-browser from ${fileName}; no server processing is required for the prototype.`,
      "Rows without usable latitude and longitude are ignored so the map stays auditable.",
      "The same dashboard, map, deployment plan, action queue, and report exports update from the uploaded dataset.",
    ],
  };
}

function readField(row, aliases) {
  for (const alias of aliases) {
    const value = row[normalizeColumn(alias)];
    if (value !== undefined && value !== null && String(value).trim() !== "" && String(value).toUpperCase() !== "NULL") return String(value).trim();
  }
  return "";
}

function normalizeColumn(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function parseListValue(value) {
  if (!value || value.toUpperCase() === "NULL") return [];
  const trimmed = value.trim();
  if (trimmed.startsWith("[") && trimmed.endsWith("]")) {
    try {
      const parsed = JSON.parse(trimmed);
      return Array.isArray(parsed) ? parsed.map(cleanToken).filter(Boolean) : [cleanToken(parsed)];
    } catch {
      return trimmed.slice(1, -1).split(",").map(cleanToken).filter(Boolean);
    }
  }
  return trimmed.split(/[|;,]/).map(cleanToken).filter(Boolean);
}

function cleanToken(value) {
  return String(value || "").replace(/^["'\s]+|["'\s]+$/g, "").trim().toUpperCase();
}

function cleanPlate(value) {
  const plate = cleanToken(value);
  if (!plate || plate === "NULL" || plate === "UNKNOWN") return "";
  return plate;
}

function parseTrafficTime(value) {
  if (!value || value.toUpperCase() === "NULL") return null;
  const raw = String(value).trim();
  const match = raw.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}:\d{2})(\.\d+)?([zZ]|[+-]\d{2}(?::?\d{2})?)?$/);
  if (!match) {
    const ms = Date.parse(raw.replace(" ", "T"));
    return Number.isNaN(ms) ? null : trafficParts(ms);
  }
  const [year, month, day] = match[1].split("-").map(Number);
  const [hour, minute, second] = match[2].split(":").map(Number);
  const millisecond = match[3] ? Number(match[3].slice(1, 4).padEnd(3, "0")) : 0;
  const offsetMinutes = parseOffsetMinutes(match[4] || "");
  const ms = Date.UTC(year, month - 1, day, hour, minute, second, millisecond) - offsetMinutes * 60 * 1000;
  return trafficParts(ms);
}

function parseOffsetMinutes(offset) {
  if (!offset || /z/i.test(offset)) return 0;
  const sign = offset.startsWith("-") ? -1 : 1;
  const digits = offset.slice(1).replace(":", "");
  const hours = Number(digits.slice(0, 2)) || 0;
  const minutes = Number(digits.slice(2, 4)) || 0;
  return sign * (hours * 60 + minutes);
}

function trafficParts(ms) {
  const shifted = new Date(ms + IST_OFFSET_MS);
  const year = shifted.getUTCFullYear();
  const month = pad2(shifted.getUTCMonth() + 1);
  const day = pad2(shifted.getUTCDate());
  return {
    ms,
    hour: shifted.getUTCHours(),
    dateKey: `${year}-${month}-${day}`,
  };
}

function parseCoordinate(value) {
  if (!String(value || "").trim()) return NaN;
  const coordinate = Number(value);
  return Number.isFinite(coordinate) ? coordinate : NaN;
}

function buildDateRangeFromMs(firstTimestamp, lastTimestamp) {
  if (!Number.isFinite(firstTimestamp) || !Number.isFinite(lastTimestamp)) return { start: "Uploaded", end: "Now" };
  return { start: formatDateShort(new Date(firstTimestamp)), end: formatDateShort(new Date(lastTimestamp)) };
}

function formatDateShort(date) {
  return date.toLocaleDateString("en-IN", { timeZone: "Asia/Kolkata", day: "2-digit", month: "short", year: "numeric" });
}

function formatDateTimeShort(ms) {
  if (!Number.isFinite(ms) || ms <= 0) return "Unknown";
  return `${new Intl.DateTimeFormat("en-IN", {
    timeZone: "Asia/Kolkata",
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  }).format(new Date(ms))} IST`;
}

function topEntries(map, count) {
  return [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, count);
}

function priorityForScore(score) {
  if (score >= 72) return "Critical";
  if (score >= 58) return "High";
  if (score >= 40) return "Watch";
  return "Routine";
}

function recommendationForHotspot(hotspot) {
  const hour = hotspot.topHours?.[0]?.hour;
  const window = Number.isInteger(hour) ? `${formatHourLabel(hour)}-${formatHourLabel((hour + 2) % 24)}` : "peak";
  if ((hotspot.junctionShare || 0) > 0.65) return `Keep a junction-clearance unit on ${window} cycle; dominant issue is ${hotspot.topViolation?.toLowerCase?.() || "parking obstruction"}.`;
  if ((hotspot.arterialShare || 0) > 0.28) return `Prioritize tow-ready arterial patrols during the ${window} recurrence window.`;
  if ((hotspot.peakShare || 0) > 0.45) return `Schedule repeat patrols in the ${window} peak window before congestion spills over.`;
  return "Assign station patrol follow-up and monitor recurrence after clearance.";
}

function deriveArea(location, junctionName, station) {
  if (junctionName && !/no junction|null/i.test(junctionName)) return titleCase(junctionName);
  const cleaned = String(location || "").split(",").map((part) => part.trim()).filter(Boolean);
  return titleCase(cleaned.slice(0, 2).join(", ") || `${station} hotspot`);
}

function classifyPlace(location, junctionName) {
  const text = `${location} ${junctionName}`.toLowerCase();
  if (/market|bazaar|mall|commercial/.test(text)) return "Commercial market";
  if (/hospital|clinic|medical/.test(text)) return "Hospital edge";
  if (/school|college|university/.test(text)) return "Institutional frontage";
  if (/junction|cross|circle|signal/.test(text)) return "Junction approach";
  if (/main road|highway|ring road|arterial/.test(text)) return "Arterial corridor";
  return "Street segment";
}

function isJunctionLinked(junctionName, violation, location) {
  return Boolean(junctionName && !/no junction|null/i.test(junctionName)) || /CROSSING|JUNCTION|CROSS|SIGNAL/.test(`${violation} ${location}`.toUpperCase());
}

function isArterialLinked(violation, location) {
  return /MAIN ROAD|ARTERIAL|HIGHWAY|RING ROAD|BUS|LANE|ROAD CROSSING/.test(`${violation} ${location}`.toUpperCase());
}

function isPeakHour(hour) {
  return (hour >= 8 && hour <= 11) || (hour >= 17 && hour <= 20);
}

function weightedShare(hotspots, key) {
  const total = sum(hotspots, "violations") || 1;
  return round3(hotspots.reduce((acc, hotspot) => acc + (Number(hotspot[key]) || 0) * hotspot.violations, 0) / total);
}

function topWeightedToken(hotspots, key) {
  const counts = new Map();
  hotspots.forEach((hotspot) => counts.set(hotspot[key], (counts.get(hotspot[key]) || 0) + hotspot.violations));
  return topEntries(counts, 1)[0]?.[0] || "UNKNOWN";
}

function titleCase(value) {
  return String(value || "").toLowerCase().replace(/\b[a-z]/g, (letter) => letter.toUpperCase());
}

function formatHourLabel(hour) {
  if (hour === 0) return "12 AM";
  if (hour === 12) return "12 PM";
  return hour < 12 ? `${hour} AM` : `${hour - 12} PM`;
}

function sum(items, key) {
  return items.reduce((total, item) => total + (Number(item[key]) || 0), 0);
}

function round1(value) {
  return Math.round((Number(value) || 0) * 10) / 10;
}

function round3(value) {
  return Math.round((Number(value) || 0) * 1000) / 1000;
}

function round5(value) {
  return Math.round((Number(value) || 0) * 100000) / 100000;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}
