import React, { useEffect, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import L from "leaflet";
import "leaflet/dist/leaflet.css";
import "leaflet.heat";
import { AnimatePresence, motion } from "framer-motion";
import LiquidGlass from "liquid-glass-react";
import {
  Activity,
  BadgeIndianRupee,
  BarChart3,
  CalendarDays,
  Car,
  CheckCircle2,
  ClipboardList,
  Download,
  FileText,
  FileUp,
  Flame,
  Gauge,
  GitBranch,
  Layers3,
  LocateFixed,
  LockKeyhole,
  LogIn,
  LogOut,
  Mail,
  MapPinned,
  MapPin,
  Menu,
  Radar,
  Route,
  Search,
  ShieldCheck,
  Siren,
  SlidersHorizontal,
  Sparkles,
  Target,
  TimerReset,
  Truck,
  X,
} from "lucide-react";
import {
  ArcElement,
  BarElement,
  CategoryScale,
  Chart as ChartJS,
  Filler,
  Legend,
  LinearScale,
  LineElement,
  PointElement,
  Tooltip,
} from "chart.js";
import { Bar, Doughnut, Line } from "react-chartjs-2";
import { buildDatasetFromUpload } from "./datasetUpload.js";
import { downloadCsvReport, downloadPdfReport } from "./reporting.js";
import "./index.css";

ChartJS.register(CategoryScale, LinearScale, PointElement, LineElement, BarElement, ArcElement, Tooltip, Legend, Filler);

const formatter = new Intl.NumberFormat("en-IN");
const MAPPLS_TOKEN = import.meta.env.VITE_MAPPLS_TOKEN || import.meta.env.VITE_MAPPLS_MAP_KEY || "";
const USE_MAPPLS_MAP = import.meta.env.VITE_MAP_PROVIDER === "mappls" && Boolean(MAPPLS_TOKEN);
const MAP_PROVIDER = USE_MAPPLS_MAP ? "Mappls" : "Leaflet fallback";
const AUTH_TOKEN_KEY = "parksight-auth-token";
const AUTH_USER_KEY = "parksight-auth-user";
const STATIC_DEMO_TOKEN = "parksight-static-demo";
const DEMO_EMAIL = "officer@parksight.ai";
const DEMO_PASSWORD = "Password1";

const priorityColors = {
  Critical: "#ff4d5e",
  High: "#ffb020",
  Watch: "#31d6b7",
  Routine: "#8d7cf6",
};

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

const layerMeta = {
  impact: {
    label: "Impact",
    hint: "Overall congestion-priority score",
    gradient: { 0.2: "#f8df7d", 0.45: "#ff9b54", 0.72: "#ff4d5e", 1: "#8b1e42" },
  },
  violations: {
    label: "Volume",
    hint: "Raw illegal-parking case density",
    gradient: { 0.2: "#c7ddff", 0.45: "#6ea8ff", 0.72: "#3267d6", 1: "#152c8f" },
  },
  junction: {
    label: "Junction",
    hint: "Intersection and crossing obstruction risk",
    gradient: { 0.2: "#d7f7ec", 0.45: "#56c9ad", 0.72: "#04756f", 1: "#073f45" },
  },
};

const views = [
  { id: "command", label: "Command", group: "Judge story", icon: Activity },
  { id: "map", label: "Hotspot map", group: "Evidence", icon: MapPinned },
  { id: "vehicles", label: "Vehicle intel", group: "Evidence", icon: Car },
  { id: "plates", label: "Number plates", group: "Evidence", icon: FileText },
  { id: "deployment", label: "Deployment", group: "Operations", icon: Route },
  { id: "simulator", label: "What-if lab", group: "Operations", icon: SlidersHorizontal },
  { id: "work", label: "Action queue", group: "Action", icon: ClipboardList },
  { id: "proof", label: "Audit trail", group: "Action", icon: ShieldCheck },
];

function App() {
  const [isAuthed, setIsAuthed] = useState(() => Boolean(sessionStorage.getItem(AUTH_TOKEN_KEY)));
  const [currentUser, setCurrentUser] = useState(() => readStoredUser());
  const [data, setData] = useState(null);
  const [uploadNotice, setUploadNotice] = useState("");
  const [uploadError, setUploadError] = useState("");
  const [activeView, setActiveView] = useState("command");
  const [station, setStation] = useState("all");
  const [priority, setPriority] = useState("all");
  const [layer, setLayer] = useState("impact");
  const [selected, setSelected] = useState(null);
  const [fleetSize, setFleetSize] = useState(12);
  const [rankMode, setRankMode] = useState("impact");
  const [mobileNav, setMobileNav] = useState(false);
  const [workStatus, setWorkStatus] = useState({});
  const mapApi = useRef(null);
  const previousStation = useRef("all");

  const handleDatasetUpload = async (file) => {
    if (!file) return;
    let uploadId = null;
    setUploadError("");
    setUploadNotice(`Preparing ${file.name}...`);
    setData(null);
    setSelected(null);
    setStation("all");
    setPriority("all");
    setLayer("impact");
    setWorkStatus({});
    try {
      let databaseOffline = false;
      try {
        const uploadSession = await startDatasetUpload(file);
        uploadId = uploadSession.upload.id;
      } catch (error) {
        if (!isDatabaseUnavailable(error)) throw error;
        databaseOffline = true;
      }
      const text = await file.text();
      await new Promise((resolve) => window.setTimeout(resolve, 0));
      setUploadNotice(`Analyzing ${file.name}...`);
      const nextData = await analyzeDatasetFile(text, file.name);
      if (uploadId) {
        setUploadNotice(`Saving ${formatter.format(nextData.summary.totalViolations)} analyzed records to database...`);
        try {
          await saveDatasetResults(uploadId, nextData);
        } catch (error) {
          if (!isDatabaseUnavailable(error)) throw error;
          databaseOffline = true;
        }
      }
      setData(nextData);
      setSelected(nextData.hotspots[0]);
      setUploadNotice(
        databaseOffline
          ? `Analyzed ${formatter.format(nextData.summary.totalViolations)} records from ${file.name}`
          : `Stored ${formatter.format(nextData.summary.totalViolations)} records from ${file.name} in PostgreSQL`,
      );
      window.setTimeout(() => mapApi.current?.reset(), 120);
    } catch (error) {
      if (uploadId) await markDatasetUploadFailed(uploadId, error.message).catch(() => {});
      setData(null);
      setSelected(null);
      setUploadNotice("");
      setUploadError(error.message || "Could not analyze that dataset.");
    }
  };

  const handleLoadDemoDataset = async () => {
    setUploadError("");
    setUploadNotice("Loading judge demo dataset...");
    setData(null);
    setSelected(null);
    try {
      const response = await fetch(`${import.meta.env.BASE_URL}data/parking_intelligence.json`);
      if (!response.ok) throw new Error("Demo dataset could not be loaded.");
      const nextData = await response.json();
      setData(nextData);
      setSelected(nextData.hotspots[0]);
      setUploadNotice(`Loaded ${formatter.format(nextData.summary.totalViolations)} precomputed records for the demo dashboard.`);
      window.setTimeout(() => mapApi.current?.reset(), 120);
    } catch (error) {
      setData(null);
      setSelected(null);
      setUploadNotice("");
      setUploadError(error.message || "Could not load the demo dataset.");
    }
  };

  const handleAuthenticated = (session) => {
    sessionStorage.setItem(AUTH_TOKEN_KEY, session.token);
    sessionStorage.setItem(AUTH_USER_KEY, JSON.stringify(session.user));
    setCurrentUser(session.user);
    setIsAuthed(true);
  };

  const handleLogout = () => {
    sessionStorage.removeItem(AUTH_TOKEN_KEY);
    sessionStorage.removeItem(AUTH_USER_KEY);
    sessionStorage.removeItem("parksight-auth");
    setIsAuthed(false);
    setCurrentUser(null);
    setData(null);
    setUploadNotice("");
    setUploadError("");
    setActiveView("command");
    setStation("all");
    setPriority("all");
    setLayer("impact");
    setSelected(null);
    setWorkStatus({});
    setMobileNav(false);
  };

  const filteredHotspots = useMemo(() => {
    if (!data?.hotspots) return [];
    return data.hotspots.filter((hotspot) => {
      const priorityMatch = priority === "all" || hotspot.priority === priority;
      const stationMatch = station === "all" || hotspot.station === station;
      return priorityMatch && stationMatch;
    });
  }, [data, priority, station]);

  const intelligence = useMemo(() => (data?.hotspots ? buildIntelligence(data, fleetSize, rankMode) : null), [data, fleetSize, rankMode]);

  useEffect(() => {
    if (!isAuthed || currentUser) return;
    let active = true;
    getCurrentUser()
      .then((session) => {
        if (!active) return;
        sessionStorage.setItem(AUTH_USER_KEY, JSON.stringify(session.user));
        setCurrentUser(session.user);
      })
      .catch(() => {
        if (!active) return;
        setCurrentUser(null);
      });
    return () => {
      active = false;
    };
  }, [currentUser, isAuthed]);

  useEffect(() => {
    if (!data?.hotspots || previousStation.current === station) return;
    previousStation.current = station;
    if (station === "all") {
      setSelected(data.hotspots[0]);
      mapApi.current?.reset();
      return;
    }
    const stationHotspots = data.hotspots.filter((hotspot) => hotspot.station === station);
    if (!stationHotspots.length) return;
    setSelected(stationHotspots[0]);
    mapApi.current?.fitHotspots(stationHotspots);
  }, [data, station]);

  if (!isAuthed) return <AuthGate onAuthenticated={handleAuthenticated} />;
  if (!data) {
    return (
      <UploadGate
        user={currentUser}
        uploadNotice={uploadNotice}
        uploadError={uploadError}
        onDatasetUpload={handleDatasetUpload}
        onLoadDemoDataset={handleLoadDemoDataset}
        onLogout={handleLogout}
      />
    );
  }

  const viewProps = {
    data,
    intelligence,
    selected,
    setSelected,
    filteredHotspots,
    station,
    setStation,
    priority,
    setPriority,
    layer,
    setLayer,
    mapApi,
    fleetSize,
    setFleetSize,
    rankMode,
    setRankMode,
    workStatus,
    setWorkStatus,
  };

  return (
    <div className="app-shell">
      <Sidebar activeView={activeView} setActiveView={setActiveView} data={data} open={mobileNav} setOpen={setMobileNav} />
      <div className="main-shell">
        <TopBar
          data={data}
          activeView={activeView}
          uploadNotice={uploadNotice}
          uploadError={uploadError}
          user={currentUser}
          onDatasetUpload={handleDatasetUpload}
          onMenu={() => setMobileNav(true)}
          onDownloadCsv={() => downloadCsvReport(data)}
          onDownloadPdf={() => downloadPdfReport(data)}
          onLogout={handleLogout}
        />
        <AnimatePresence mode="wait">
          <motion.main
            key={activeView}
            initial={{ opacity: 0, y: 12, filter: "blur(8px)" }}
            animate={{ opacity: 1, y: 0, filter: "blur(0px)" }}
            exit={{ opacity: 0, y: -8, filter: "blur(8px)" }}
            transition={{ duration: 0.22, ease: "easeOut" }}
            className="view-stage"
          >
            {activeView === "command" && <CommandView {...viewProps} />}
            {activeView === "map" && <MapView {...viewProps} />}
            {activeView === "vehicles" && <VehicleView {...viewProps} />}
            {activeView === "plates" && <PlateView {...viewProps} />}
            {activeView === "deployment" && <DeploymentView {...viewProps} />}
            {activeView === "simulator" && <SimulatorView {...viewProps} />}
            {activeView === "work" && <WorkOrdersView {...viewProps} />}
            {activeView === "proof" && <ProofView {...viewProps} />}
          </motion.main>
        </AnimatePresence>
        <SiteFooter />
      </div>
    </div>
  );
}

function SiteFooter() {
  return (
    <footer className="site-footer">
      Made by Rajdeep Bandyopadhaya and Aniket Arya. All rights reserved.
    </footer>
  );
}

function readStoredUser() {
  try {
    const raw = sessionStorage.getItem(AUTH_USER_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function isDatabaseUnavailable(error) {
  const message = String(error?.message || "");
  return /database is not reachable|static demo|econnrefused|enotfound|etimedout|ehostunreach/i.test(message);
}

async function submitAuthRequest(mode, form) {
  try {
    return await apiRequest(`/api/auth/${mode === "signup" ? "signup" : "login"}`, {
      method: "POST",
      body: JSON.stringify({
        name: form.name.trim(),
        email: form.email.trim(),
        password: form.password,
      }),
    });
  } catch (error) {
    if (mode === "login" && isStaticDemoHost() && isDemoCredential(form)) {
      const user = { id: "static-demo-user", full_name: "Judge Demo User", email: DEMO_EMAIL };
      return { user, token: STATIC_DEMO_TOKEN, offline: true };
    }
    throw error;
  }
}

async function getCurrentUser() {
  if (sessionStorage.getItem(AUTH_TOKEN_KEY) === STATIC_DEMO_TOKEN) {
    return { user: readStoredUser() || { id: "static-demo-user", full_name: "Judge Demo User", email: DEMO_EMAIL } };
  }
  return apiRequest("/api/auth/me", {
    method: "GET",
    auth: true,
  });
}

async function startDatasetUpload(file) {
  if (sessionStorage.getItem(AUTH_TOKEN_KEY) === STATIC_DEMO_TOKEN) {
    throw new Error("Static demo mode cannot persist uploads.");
  }
  return apiRequest("/api/uploads/start", {
    method: "POST",
    auth: true,
    body: JSON.stringify({
      fileName: file.name,
      fileSize: file.size,
      mimeType: file.type,
    }),
  });
}

async function saveDatasetResults(uploadId, data) {
  return apiRequest(`/api/uploads/${uploadId}/results`, {
    method: "PATCH",
    auth: true,
    body: JSON.stringify({ data }),
  });
}

async function markDatasetUploadFailed(uploadId, message) {
  return apiRequest(`/api/uploads/${uploadId}/fail`, {
    method: "PATCH",
    auth: true,
    body: JSON.stringify({ message }),
  });
}

async function apiRequest(path, options = {}) {
  const headers = {
    "Content-Type": "application/json",
    ...(options.headers || {}),
  };
  if (options.auth) {
    const token = sessionStorage.getItem(AUTH_TOKEN_KEY);
    if (!token) throw new Error("Please log in again before continuing.");
    headers.Authorization = `Bearer ${token}`;
  }
  const response = await fetch(path, { ...options, headers });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || payload.detail || "The server could not complete the request.");
  }
  return payload;
}

function isStaticDemoHost() {
  return window.location.hostname.endsWith("github.io") || window.location.protocol === "file:";
}

function isDemoCredential(form) {
  return form.email.trim().toLowerCase() === DEMO_EMAIL && form.password === DEMO_PASSWORD;
}

function analyzeDatasetFile(text, name) {
  if (typeof Worker === "undefined") return Promise.resolve(buildDatasetFromUpload(text, { name }));
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL("./datasetWorker.js", import.meta.url), { type: "module" });
    worker.onmessage = (event) => {
      worker.terminate();
      if (event.data?.ok) {
        resolve(event.data.data);
      } else {
        reject(new Error(event.data?.error || "Could not analyze that dataset."));
      }
    };
    worker.onerror = (error) => {
      worker.terminate();
      reject(new Error(error.message || "Dataset worker failed."));
    };
    worker.postMessage({ text, name });
  });
}

function Sidebar({ activeView, setActiveView, data, open, setOpen }) {
  const grouped = groupBy(views, "group");
  const [studyFloating, setStudyFloating] = useState(false);
  return (
    <>
      <aside className={`sidebar ${open ? "is-open" : ""}`}>
        <div className="brand-block">
          <div className="brand-mark">
            <Radar size={22} />
          </div>
          <div>
            <strong>ParkSight AI</strong>
            <span>Congestion Enforcement OS</span>
          </div>
          <button className="icon-button mobile-close" onClick={() => setOpen(false)} aria-label="Close navigation">
            <X size={18} />
          </button>
        </div>

        <div className="role-chip">
          <Sparkles size={15} />
          Judge walkthrough mode
        </div>

        <nav className="nav-groups">
          {Object.entries(grouped).map(([group, items]) => (
            <section key={group}>
              <p>{group}</p>
              {items.map((item) => {
                const Icon = item.icon;
                return (
                  <button
                    key={item.id}
                    className={activeView === item.id ? "active" : ""}
                    onClick={() => {
                      setActiveView(item.id);
                      setOpen(false);
                    }}
                  >
                    <Icon size={18} />
                    <span>{item.label}</span>
                  </button>
                );
              })}
            </section>
          ))}
        </nav>

        <button
          type="button"
          className={`sidebar-data ${studyFloating ? "is-floating" : ""}`}
          aria-pressed={studyFloating}
          onClick={() => setStudyFloating((current) => !current)}
        >
          <span>Study Data</span>
          <strong>{formatter.format(data.summary.totalViolations)}</strong>
          <small>{data.summary.dateRange.start} to {data.summary.dateRange.end}</small>
        </button>
      </aside>
      {open && <button className="nav-scrim" onClick={() => setOpen(false)} aria-label="Close navigation" />}
    </>
  );
}

function TopBar({ data, activeView, uploadNotice, uploadError, user, onDatasetUpload, onMenu, onDownloadCsv, onDownloadPdf, onLogout }) {
  const current = views.find((view) => view.id === activeView);
  const title = activeView === "command" ? "Dashboard" : current?.label || "Dashboard";
  return (
    <header className="topbar">
      <div className="topbar-title">
        <button className="icon-button mobile-menu" onClick={onMenu} aria-label="Open navigation">
          <Menu size={19} />
        </button>
        <div>
          <span>Flipkart Gridlock Challenge</span>
          <strong>{title}</strong>
        </div>
      </div>
      <div className="topbar-actions">
        {(uploadNotice || uploadError) && <span className={`upload-status ${uploadError ? "is-error" : ""}`}>{uploadError || uploadNotice}</span>}
        <UserChip user={user} />
        <Badge icon={<Layers3 size={15} />} label={data.summary.modelVersion} />
        <Badge icon={<CalendarDays size={15} />} label={`${data.summary.dateRange.start} - ${data.summary.dateRange.end}`} />
        <label className="soft-button upload-button">
          <FileUp size={15} />
          Upload data
          <input
            type="file"
            accept=".csv,.json,application/json,text/csv"
            onChange={(event) => {
              onDatasetUpload(event.target.files?.[0]);
              event.target.value = "";
            }}
          />
        </label>
        <button onClick={onDownloadCsv} className="soft-button">
          <Download size={15} />
          CSV
        </button>
        <button onClick={onDownloadPdf} className="primary-button">
          <FileText size={15} />
          PDF brief
        </button>
        <button onClick={onLogout} className="soft-button logout-button">
          <LogOut size={15} />
          Logout
        </button>
      </div>
    </header>
  );
}

function UserChip({ user }) {
  const name = user?.full_name || user?.fullName || user?.name || "Signed in";
  const email = user?.email || "";
  return (
    <div className="user-chip" title={email ? `${name} (${email})` : name}>
      <ShieldCheck size={16} />
      <span>
        <strong>{name}</strong>
        {email && <small>{email}</small>}
      </span>
    </div>
  );
}

function CommandView(props) {
  const { data, intelligence, selected, setSelected, filteredHotspots, layer, setLayer, mapApi } = props;
  return (
    <div className="command-grid">
      <MetricRail data={data} intelligence={intelligence} />

      <LiquidPanel className="span-8">
        <section className="hero-panel">
          <div>
            <p className="eyebrow">Civic signal, not ticket-count theatre</p>
            <h1>Illegal-parking intelligence that tells officers where traffic will choke next.</h1>
            <p className="hero-copy">
              ParkSight ranks each Bengaluru hotspot by obstruction, junction exposure, peak recurrence, and enforcement feasibility, then turns the evidence into dispatchable beats.
            </p>
          </div>
          <div className="hero-score">
            <span>Top impact score</span>
            <strong>{data.hotspots[0].impactScore}</strong>
            <small>{data.hotspots[0].area}</small>
          </div>
        </section>
      </LiquidPanel>

      <LiquidPanel className="span-4" variant="quiet">
        <section className="mission-card">
          <PanelHeader icon={<Target />} title="Judge signal" detail="What makes this different" />
          <div className="signal-list">
            <Signal value={`${Math.round(data.summary.junctionLinkedShare * 100)}%`} label="junction-linked evidence" />
            <Signal value={`${Math.round(data.summary.peakHourShare * 100)}%`} label="peak-window recurrence" />
            <Signal value={`${intelligence.planCoverage}%`} label="top-beat burden captured" />
          </div>
        </section>
      </LiquidPanel>

      <section className="map-card span-8">
        <PanelHeader icon={<MapPinned />} title="Live impact surface" detail={`${filteredHotspots.length} visible hotspots`} />
        <MapPanel data={data} layer={layer} hotspots={filteredHotspots} selected={selected} onSelect={setSelected} mapApi={mapApi} compact />
        <LayerSwitch layer={layer} setLayer={setLayer} />
      </section>

      <section className="span-4 stack">
        <SelectedHotspotCard hotspot={selected} />
        <PriorityList items={intelligence.deployment.slice(0, 7)} onFocus={(hotspot) => {
          setSelected(hotspot);
          mapApi.current?.focus(hotspot);
        }} />
      </section>

      <section className="panel span-4">
        <PanelHeader icon={<Truck />} title="Vehicle obstruction" detail="PCU-weighted view" />
        <Doughnut data={vehicleDoughnutData(data)} options={doughnutOptions} />
      </section>
      <section className="panel span-4">
        <PanelHeader icon={<TimerReset />} title="Violation rhythm" detail="hour of day" />
        <Line data={hourChartData(data)} options={lineOptions} />
      </section>
      <section className="panel span-4">
        <PanelHeader icon={<BarChart3 />} title="Station burden" detail="top police stations" />
        <Bar data={stationChartData(data)} options={barOptions} />
      </section>
    </div>
  );
}

function MapView(props) {
  const { data, selected, setSelected, filteredHotspots, station, setStation, priority, setPriority, layer, setLayer, mapApi, intelligence } = props;
  return (
    <div className="evidence-layout">
      <section className="panel map-workbench">
        <PanelHeader icon={<MapPinned />} title="Hotspot evidence map" detail="filter by police station, priority, and risk layer" />
        <ControlBar data={data} station={station} setStation={setStation} priority={priority} setPriority={setPriority} layer={layer} setLayer={setLayer} />
        <MapPanel data={data} layer={layer} hotspots={filteredHotspots} selected={selected} onSelect={setSelected} mapApi={mapApi} />
      </section>
      <aside className="stack">
        <SelectedHotspotCard hotspot={selected} />
        <StationLeague stations={intelligence.stationLeague.slice(0, 12)} onPick={(stationName) => setStation(stationName)} />
      </aside>
    </div>
  );
}

function VehicleView({ data, intelligence, setSelected, mapApi }) {
  return (
    <div className="vehicle-grid">
      <section className="panel span-5">
        <PanelHeader icon={<Car />} title="Vehicle mix" detail="raw violation records" />
        <Doughnut data={vehicleDoughnutData(data)} options={doughnutOptions} />
      </section>
      <section className="panel span-7">
        <PanelHeader icon={<Gauge />} title="Obstruction pressure by vehicle type" detail="PCU-weighted estimate" />
        <Bar data={vehiclePressureData(data)} options={barOptions} />
      </section>
      <section className="panel span-6">
        <PanelHeader icon={<Siren />} title="Vehicle-specific enforcement playbooks" detail="turn analytics into action" />
        <div className="playbook-list">
          {intelligence.vehiclePlaybooks.map((item) => (
            <article key={item.name}>
              <div>
                <strong>{item.name}</strong>
                <span>{formatter.format(item.count)} cases · {formatter.format(Math.round(item.pressure))} PCU pressure</span>
              </div>
              <p>{item.playbook}</p>
            </article>
          ))}
        </div>
      </section>
      <section className="panel span-6">
        <PanelHeader icon={<LocateFixed />} title="Top vehicle-led hotspots" detail="click to inspect on map" />
        <div className="dense-table">
          <div className="dense-row head">
            <span>Area</span>
            <span>Vehicle</span>
            <span>Score</span>
          </div>
          {intelligence.vehicleHotspots.slice(0, 12).map((hotspot) => (
            <button
              key={hotspot.id}
              className="dense-row"
              onClick={() => {
                setSelected(hotspot);
                mapApi.current?.focus(hotspot);
              }}
            >
              <span>{hotspot.area}</span>
              <span>{hotspot.topVehicle}</span>
              <strong>{hotspot.impactScore}</strong>
            </button>
          ))}
        </div>
      </section>
    </div>
  );
}

function PlateView({ data }) {
  const plates = data.plates || [];
  const [plateQuery, setPlateQuery] = useState("");
  const normalizedQuery = plateQuery.trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
  const visiblePlates = normalizedQuery ? plates.filter((item) => item.plate.includes(normalizedQuery)) : plates;
  const repeatPlates = plates.filter((item) => item.count > 1);
  const totalPlateEvents = sum(plates, "count");
  const topPlate = plates[0];
  return (
    <div className="plate-grid">
      <section className="panel span-12">
        <PanelHeader icon={<FileText />} title="Number plate violations" detail="derived from uploaded vehicle_number fields" />
        <div className="work-summary">
          <Kpi title="Unique plates" value={formatter.format(plates.length)} detail="visible in dataset" />
          <Kpi title="Plate-linked cases" value={formatter.format(totalPlateEvents)} detail="records with number plate" />
          <Kpi title="Repeat plates" value={formatter.format(repeatPlates.length)} detail="more than one record" />
          <Kpi title="Top plate" value={topPlate?.plate || "None"} detail={topPlate ? `${topPlate.count} records` : "no plate column found"} className="plate-kpi" />
        </div>
      </section>

      <section className="panel span-12">
        <PanelHeader icon={<Search />} title="Plate evidence table" detail={plates.length ? "search by number plate" : "no number plate field found"} />
        {plates.length ? (
          <>
            <label className="plate-search">
              <Search size={17} />
              <input value={plateQuery} onChange={(event) => setPlateQuery(event.target.value)} placeholder="Search plate number, e.g. KA01AA1111" />
              <span>{formatter.format(visiblePlates.length)} matches</span>
            </label>
            <div className="plate-table">
              <div className="plate-row head">
                <span>Number plate</span>
                <span>Records</span>
                <span>Vehicle</span>
                <span>Station</span>
                <span>Dominant violation</span>
                <span>Peak hour</span>
                <span>Last seen</span>
              </div>
              {visiblePlates.slice(0, 60).map((item) => (
                <div className="plate-row" key={item.plate}>
                  <strong>{item.plate}</strong>
                  <b>{formatter.format(item.count)}</b>
                  <span>{item.vehicle}</span>
                  <span>{item.station}</span>
                  <span>{item.violation}</span>
                  <span>{formatHourLabel(item.peakHour)}</span>
                  <LastSeen value={item.lastSeen} />
                </div>
              ))}
              {!visiblePlates.length && <p className="empty-note">No number plate matches found for "{plateQuery}".</p>}
            </div>
          </>
        ) : (
          <p className="empty-note">Upload data with vehicle_number or updated_vehicle_number columns to inspect number plate violations.</p>
        )}
      </section>
    </div>
  );
}

function LastSeen({ value }) {
  const [datePart, timePart] = splitLastSeen(value);
  return (
    <span className="last-seen-cell" title={value}>
      <span>{datePart}</span>
      {timePart && <small>{timePart}</small>}
    </span>
  );
}

function splitLastSeen(value) {
  const text = String(value || "Unknown").trim();
  const match = text.match(/^(.+?),\s*(.+)$/);
  return match ? [match[1], match[2]] : [text, ""];
}

function DeploymentView({ data, intelligence, fleetSize, setFleetSize, setSelected, mapApi }) {
  const shiftMix = countBy(intelligence.deployment, (item) => item.shift);
  return (
    <div className="deployment-grid">
      <section className="panel span-12">
        <PanelHeader icon={<Route />} title="Tow and enforcement deployment plan" detail="ranked by congestion impact, with a route-ready sequence" />
        <div className="fleet-strip">
          <Kpi title="Fleet size" value={fleetSize} detail="deployable beats" />
          <Kpi title="Impact captured" value={`${intelligence.simulation.impactPct}%`} detail="of ranked hotspot burden" />
          <Kpi title="Route length" value={`${intelligence.routeKm} km`} detail="centroid estimate" />
          <Kpi title="Peak units" value={shiftMix["Peak"] || 0} detail="highest traffic windows" />
          <label className="range-control">
            <span>Units available</span>
            <input type="range" min="6" max="36" value={fleetSize} onChange={(event) => setFleetSize(Number(event.target.value))} />
          </label>
        </div>
      </section>

      <section className="panel span-5">
        <PanelHeader icon={<MapPin />} title="Dispatch route" detail="straight-line planning approximation" />
        <div className="route-list">
          {intelligence.deployment.map((item) => (
            <button
              key={item.id}
              onClick={() => {
                setSelected(item);
                mapApi.current?.focus(item);
              }}
            >
              <b>{item.sequence}</b>
              <span>
                <strong>{item.station}</strong>
                {item.window} · +{item.legKm} km
              </span>
              <em>{item.impactScore}</em>
            </button>
          ))}
        </div>
      </section>

      <section className="panel span-7">
        <PanelHeader icon={<ClipboardList />} title="Unit cards" detail="who goes where and why" />
        <div className="unit-grid">
          {intelligence.deployment.map((item) => (
            <article key={item.id}>
              <span>{item.unit}</span>
              <strong>{item.area}</strong>
              <small>{item.station} · {item.shift}</small>
              <p>{item.recommendation}</p>
            </article>
          ))}
        </div>
      </section>
    </div>
  );
}

function SimulatorView({ intelligence, fleetSize, setFleetSize, rankMode, setRankMode }) {
  return (
    <div className="sim-grid">
      <section className="hero-panel compact-hero stable-hero span-12">
        <div>
          <p className="eyebrow">What-if lab</p>
          <h1>Show judges why impact-ranked enforcement beats ticket-count patrols.</h1>
          <p className="hero-copy">Move the fleet size and compare how much obstruction pressure the city captures when it ranks by impact, volume, or junction risk.</p>
        </div>
        <div className="hero-score">
          <span>Current lift</span>
          <strong>+{intelligence.simulation.liftPts}</strong>
          <small>impact points over count ranking</small>
        </div>
      </section>

      <section className="panel span-5">
        <PanelHeader icon={<SlidersHorizontal />} title="Planner controls" detail="interactive scenario" />
        <div className="sim-controls">
          <label className="range-control large">
            <span>Zones covered: {fleetSize}</span>
            <input type="range" min="6" max="80" value={fleetSize} onChange={(event) => setFleetSize(Number(event.target.value))} />
          </label>
          <div className="segmented">
            {["impact", "violations", "junction"].map((mode) => (
              <button key={mode} className={rankMode === mode ? "active" : ""} onClick={() => setRankMode(mode)}>
                {mode}
              </button>
            ))}
          </div>
        </div>
        <div className="sim-metrics">
          <Kpi title="Violations reached" value={`${intelligence.simulation.violPct}%`} detail={`${formatter.format(intelligence.simulation.violations)} cases`} />
          <Kpi title="Impact reached" value={`${intelligence.simulation.impactPct}%`} detail={`${formatter.format(Math.round(intelligence.simulation.impact))} score mass`} />
          <Kpi title="Chronic zones" value={intelligence.simulation.critical} detail="critical/high zones covered" />
        </div>
      </section>

      <section className="panel span-7">
        <PanelHeader icon={<BarChart3 />} title="Coverage curve" detail="cumulative benefit as zones are added" />
        <CoverageCurve points={intelligence.curve} active={fleetSize} />
      </section>

      <section className="panel span-12">
        <PanelHeader icon={<CheckCircle2 />} title="Judge-ready finding" detail="plain-language outcome" />
        <p className="finding">
          At {fleetSize} zones, the selected strategy captures {intelligence.simulation.impactPct}% of ranked congestion impact while covering {intelligence.simulation.violPct}% of violations.
          Impact ranking finds {intelligence.simulation.liftPts} percentage points more congestion pressure than a simple count-ranked patrol list on the same fleet size.
        </p>
      </section>
    </div>
  );
}

function WorkOrdersView({ intelligence, workStatus, setWorkStatus }) {
  const statusCounts = intelligence.workOrders.reduce(
    (counts, order) => {
      const status = workStatus[order.id] || "open";
      counts[status] = (counts[status] || 0) + 1;
      return counts;
    },
    { open: 0, assigned: 0, "field-check": 0, resolved: 0 },
  );

  return (
    <div className="work-grid">
      <section className="panel span-12">
        <PanelHeader icon={<ClipboardList />} title="Action queue" detail="convert hotspot evidence into agency-owned work" />
        <div className="work-summary">
          <Kpi title="Open actions" value={statusCounts.open} detail="ready for assignment" />
          <Kpi title="Assigned" value={statusCounts.assigned} detail="owned by agency team" />
          <Kpi title="Field checks" value={statusCounts["field-check"]} detail="awaiting ground validation" />
          <Kpi title="Resolved" value={statusCounts.resolved} detail="closed after intervention" />
        </div>
      </section>
      <section className="panel span-12">
        <div className="work-table">
          <div className="work-row head">
            <span>#</span>
            <span>Hotspot</span>
            <span>Action</span>
            <span>Owner</span>
            <span>Status</span>
            <span>Impact</span>
          </div>
          {intelligence.workOrders.map((order) => (
            <div className="work-row" key={order.id}>
              <span>{order.rank}</span>
              <strong>{order.area}</strong>
              <span>{order.action}</span>
              <span>{order.owner}</span>
              <select value={workStatus[order.id] || "open"} onChange={(event) => setWorkStatus((current) => ({ ...current, [order.id]: event.target.value }))}>
                <option value="open">Open</option>
                <option value="assigned">Assigned</option>
                <option value="field-check">Field check</option>
                <option value="resolved">Resolved</option>
              </select>
              <b>{order.impactScore}</b>
            </div>
          ))}
        </div>
      </section>
    </div>
  );
}

function ProofView({ data, intelligence }) {
  return (
    <div className="proof-grid">
      <section className="hero-panel compact-hero stable-hero span-12">
        <div>
          <p className="eyebrow">Audit trail</p>
          <h1>Transparent enough for police, transport planners, and judges to audit.</h1>
          <p className="hero-copy">The score avoids a black box because the supplied data has violations, not ground-truth speeds. Every hotspot is ranked from visible, explainable civic signals.</p>
        </div>
        <div className="hero-score">
          <span>Scored cells</span>
          <strong>{formatter.format(data.summary.cellsAnalyzed)}</strong>
          <small>220 m urban grid cells</small>
        </div>
      </section>

      <section className="panel span-7">
        <PanelHeader icon={<ShieldCheck />} title="Parking Impact Index" detail="auditable formula" />
        <div className="formula">{data.method.index}</div>
        <div className="proof-list">
          {data.method.notes.map((note) => <p key={note}>{note}</p>)}
        </div>
      </section>
      <section className="panel span-5">
        <PanelHeader icon={<Search />} title="Validation checks" detail="judge trust checklist" />
        <div className="check-list">
          <Check label="Uses supplied Bengaluru violation records" value={formatter.format(data.summary.totalViolations)} />
          <Check label="Separates density from obstruction pressure" value="PCU weighting" />
          <Check label="Calls out model limitation" value="No speed labels" />
          <Check label="Produces field-operational output" value={`${intelligence.deployment.length} beats`} />
        </div>
      </section>
      <section className="panel span-6">
        <PanelHeader icon={<GitBranch />} title="Feature contribution narrative" detail="why a zone rises" />
        <FeatureStack />
      </section>
      <section className="panel span-6">
        <PanelHeader icon={<BadgeIndianRupee />} title="Civic ROI estimate" detail="not revenue theatre" />
        <p className="finding">
          The app prioritizes capacity recovery: clear junction mouths, stop peak-window obstruction, and move repeat hotspots into structural fixes.
          For judging, this shows a closed loop from evidence to deployment to policy repair.
        </p>
      </section>
    </div>
  );
}

function MetricRail({ data, intelligence }) {
  const criticalCount = data.hotspots.filter((hotspot) => hotspot.priority === "Critical").length;
  return (
    <section className="metric-rail span-12">
      <Kpi title="Violations analyzed" value={formatter.format(data.summary.totalViolations)} icon={<Siren />} detail="geocoded parking records" />
      <Kpi title="Scored cells" value={formatter.format(data.summary.cellsAnalyzed)} icon={<Target />} detail="urban grid units" />
      <Kpi title="Critical hotspots" value={criticalCount} icon={<Flame />} detail="highest impact class" />
      <Kpi title="Deployable beats" value={data.enforcementPlan.length} icon={<Route />} detail={`${formatter.format(intelligence.planCases)} cases covered`} />
    </section>
  );
}

function ControlBar({ data, station, setStation, priority, setPriority, layer, setLayer }) {
  return (
    <div className="control-bar">
      <label>
        <span>Police station</span>
        <select value={station} onChange={(event) => setStation(event.target.value)}>
          <option value="all">All police stations</option>
          {data.stations.map((item) => (
            <option key={item.station} value={item.station}>
              {item.station} ({formatter.format(item.cases)})
            </option>
          ))}
        </select>
      </label>
      <div className="segmented">
        {["all", "Critical", "High", "Watch"].map((item) => (
          <button key={item} onClick={() => setPriority(item)} className={priority === item ? "active" : ""}>
            {item === "all" ? "All" : item}
          </button>
        ))}
      </div>
      <LayerSwitch layer={layer} setLayer={setLayer} />
    </div>
  );
}

function LayerSwitch({ layer, setLayer }) {
  const options = [
    ["impact", Flame, "Impact"],
    ["violations", MapPin, "Volume"],
    ["junction", GitBranch, "Junction"],
  ];
  return (
    <div className="layer-switch">
      {options.map(([id, Icon, label]) => (
        <button key={id} onClick={() => setLayer(id)} className={layer === id ? "active" : ""}>
          <Icon size={16} />
          {label}
        </button>
      ))}
    </div>
  );
}

function MapPanel(props) {
  if (USE_MAPPLS_MAP) return <MapplsMapPanel {...props} />;
  return <LeafletMapPanel {...props} />;
}

function MapplsMapPanel({ data, layer, hotspots, selected, onSelect, mapApi, compact = false }) {
  const mapRef = useRef(null);
  const mapInstance = useRef(null);
  const circleLayers = useRef([]);
  const [mapError, setMapError] = useState("");
  const [useFallback, setUseFallback] = useState(false);

  useEffect(() => {
    let cancelled = false;
    if (!mapRef.current || mapInstance.current) return undefined;
    loadMapplsSdk(MAPPLS_TOKEN)
      .then((api) => {
        if (cancelled || !mapRef.current || mapInstance.current) return;
        mapInstance.current = new api.Map(mapRef.current.id, {
          center: [12.975, 77.6],
          zoom: 12,
          fullscreenControl: false,
          traffic: false,
          hybrid: false,
        });
        mapApi.current = {
          focus: (hotspot) => focusMappls(mapInstance.current, hotspot, 16),
          fitHotspots: (items) => {
            if (!items?.length) return;
            const target = items.length === 1 ? items[0] : centerOf(items);
            focusMappls(mapInstance.current, target, items.length === 1 ? 15 : 12);
          },
          reset: () => focusMappls(mapInstance.current, { lat: 12.975, lng: 77.6 }, 12),
        };
        renderMapplsCircles(api, mapInstance.current, circleLayers, hotspots, selected, layer, data, onSelect);
      })
      .catch(() => {
        setMapError("Mappls rejected this credential. Using fallback map until a valid Web SDK token is added.");
        setUseFallback(true);
      });

    return () => {
      cancelled = true;
      clearMapplsLayers(circleLayers, mapInstance.current);
      mapInstance.current?.remove?.();
      mapInstance.current = null;
    };
  }, [mapApi]);

  useEffect(() => {
    if (!mapInstance.current) return;
    loadMapplsSdk(MAPPLS_TOKEN)
      .then((api) => renderMapplsCircles(api, mapInstance.current, circleLayers, hotspots, selected, layer, data, onSelect))
      .catch(() => {
        setMapError("Mappls rejected this credential. Using fallback map until a valid Web SDK token is added.");
        setUseFallback(true);
      });
  }, [data, hotspots, layer, onSelect, selected]);

  if (useFallback) {
    return (
      <div className="map-fallback-wrap">
        <LeafletMapPanel data={data} layer={layer} hotspots={hotspots} selected={selected} onSelect={onSelect} mapApi={mapApi} compact={compact} />
        <div className="map-provider-warning">{mapError}</div>
      </div>
    );
  }

  return (
    <div className={`map-shell ${compact ? "compact" : ""}`}>
      <div ref={mapRef} id={`mappls-${compact ? "compact" : "full"}`} className="map-canvas" />
      <div className="map-legend">
        <LegendDot color={layer === "impact" ? priorityColors.Critical : layer === "junction" ? "#04756f" : "#3267d6"} label={`${layerMeta[layer].label} on Mappls`} />
        <span>{hotspots.length} visible hotspots</span>
      </div>
      <div className="map-note">
        <strong>{mapError || `${MAP_PROVIDER} layer`}</strong>
        <span>{mapError || layerMeta[layer].hint}</span>
      </div>
    </div>
  );
}

function LeafletMapPanel({ data, layer, hotspots, selected, onSelect, mapApi, compact = false }) {
  const mapRef = useRef(null);
  const mapInstance = useRef(null);
  const heatRef = useRef(null);
  const markerLayerRef = useRef(null);

  useEffect(() => {
    if (!mapRef.current || mapInstance.current) return;
    mapInstance.current = L.map(mapRef.current, { zoomControl: false, preferCanvas: false }).setView([12.975, 77.6], 12);
    mapInstance.current.createPane("heatPane");
    mapInstance.current.getPane("heatPane").style.zIndex = 390;
    mapInstance.current.createPane("hotspotPane");
    mapInstance.current.getPane("hotspotPane").style.zIndex = 660;
    mapInstance.current.getPane("hotspotPane").style.pointerEvents = "auto";
    L.control.zoom({ position: "bottomright" }).addTo(mapInstance.current);
    L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19,
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
    }).addTo(mapInstance.current);

    heatRef.current = L.heatLayer(getHeatPoints(data, layer), {
      pane: "heatPane",
      radius: layer === "violations" ? 27 : layer === "junction" ? 20 : 23,
      blur: layer === "junction" ? 16 : 22,
      maxZoom: 19,
      gradient: layerMeta[layer].gradient,
    }).addTo(mapInstance.current);
    markerLayerRef.current = L.layerGroup([], { pane: "hotspotPane" }).addTo(mapInstance.current);

    mapApi.current = {
      focus: (hotspot) => mapInstance.current?.flyTo([hotspot.lat, hotspot.lng], 16, { duration: 0.8 }),
      fitHotspots: (items) => {
        if (!mapInstance.current || !items?.length) return;
        if (items.length === 1) {
          mapInstance.current.flyTo([items[0].lat, items[0].lng], 15, { duration: 0.8 });
          return;
        }
        const bounds = L.latLngBounds(items.slice(0, 45).map((item) => [item.lat, item.lng]));
        mapInstance.current.flyToBounds(bounds, { padding: [72, 72], maxZoom: 14, duration: 0.8 });
      },
      reset: () => mapInstance.current?.flyTo([12.975, 77.6], 12, { duration: 0.8 }),
    };

    return () => {
      mapInstance.current?.remove();
      mapInstance.current = null;
    };
  }, [data, layer, mapApi]);

  useEffect(() => {
    if (!heatRef.current) return;
    heatRef.current.setOptions({
      radius: layer === "violations" ? 27 : layer === "junction" ? 20 : 23,
      blur: layer === "junction" ? 16 : 22,
      gradient: layerMeta[layer].gradient,
    });
    heatRef.current.setLatLngs(getHeatPoints(data, layer));
  }, [data, layer]);

  useEffect(() => {
    if (!markerLayerRef.current) return;
    markerLayerRef.current.clearLayers();
    hotspots.forEach((hotspot) => {
      const selectedMarker = selected?.id === hotspot.id;
      L.circleMarker([hotspot.lat, hotspot.lng], {
        pane: "hotspotPane",
        className: `hotspot-dot ${selectedMarker ? "is-selected" : ""}`,
        radius: getMarkerRadius(hotspot, layer, data),
        color: selectedMarker ? "#ffffff" : "rgba(255,255,255,0.9)",
        weight: selectedMarker ? 4 : 2.5,
        opacity: 1,
        fillColor: getMarkerColor(hotspot, layer, data),
        fillOpacity: selectedMarker ? 0.96 : 0.88,
      })
        .bindPopup(
          `<strong>${hotspot.rank}. ${escapeHtml(hotspot.area)}</strong><br/>Impact score: <b>${hotspot.impactScore}</b><br/>${formatter.format(
            hotspot.violations,
          )} cases &middot; ${escapeHtml(hotspot.station)}`,
        )
        .on("click", () => onSelect(hotspot))
        .addTo(markerLayerRef.current);
    });
  }, [hotspots, layer, onSelect, selected?.id]);

  return (
    <div className={`map-shell ${compact ? "compact" : ""}`}>
      <div ref={mapRef} className="map-canvas" />
      <div className="map-legend">
        <LegendDot color={layer === "impact" ? priorityColors.Critical : layer === "junction" ? "#04756f" : "#3267d6"} label={layerMeta[layer].label} />
        <span>{hotspots.length} visible hotspots</span>
      </div>
      <div className="map-note">
        <strong>{layerMeta[layer].label} layer</strong>
        <span>{layerMeta[layer].hint}</span>
      </div>
    </div>
  );
}

function loadMapplsSdk(token) {
  if (window.mappls || window.Mappls) return Promise.resolve(window.mappls || window.Mappls);
  if (window.__parksightMapplsPromise) return window.__parksightMapplsPromise;
  window.__parksightMapplsPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = `https://apis.mappls.com/advancedmaps/api/${token}/map_sdk?layer=vector&v=3.0&callback=initMapplsSdk`;
    script.async = true;
    script.defer = true;
    window.initMapplsSdk = () => resolve(window.mappls || window.Mappls);
    script.onerror = reject;
    document.head.appendChild(script);
  });
  return window.__parksightMapplsPromise;
}

function renderMapplsCircles(api, map, layersRef, hotspots, selected, layer, data, onSelect) {
  clearMapplsLayers(layersRef, map);
  layersRef.current = hotspots.slice(0, 220).map((hotspot) => {
    const selectedMarker = selected?.id === hotspot.id;
    const circle = new api.Circle({
      map,
      center: { lat: hotspot.lat, lng: hotspot.lng },
      radius: Math.round(getMarkerRadius(hotspot, layer, data) * 13),
      fillColor: getMarkerColor(hotspot, layer, data),
      fillOpacity: selectedMarker ? 0.88 : 0.7,
      strokeColor: selectedMarker ? "#ffffff" : "rgba(255,255,255,0.86)",
      strokeOpacity: 1,
      strokeWeight: selectedMarker ? 4 : 2,
    });
    circle.addListener?.("click", () => onSelect(hotspot));
    circle.addEventListener?.("click", () => onSelect(hotspot));
    return circle;
  });
}

function clearMapplsLayers(layersRef, map) {
  const api = window.mappls || window.Mappls;
  layersRef.current.forEach((layer) => {
    try {
      if (api?.remove) api.remove({ map, layer });
      else layer?.remove?.();
    } catch {
      layer?.remove?.();
    }
  });
  layersRef.current = [];
}

function focusMappls(map, target, zoom) {
  if (!map || !target) return;
  const center = [target.lat, target.lng];
  if (map.setCenter) map.setCenter(center);
  if (map.setZoom) map.setZoom(zoom);
  if (map.flyTo) map.flyTo({ center, zoom });
  if (map.panTo) map.panTo(center);
}

function centerOf(items) {
  return {
    lat: items.reduce((total, item) => total + item.lat, 0) / items.length,
    lng: items.reduce((total, item) => total + item.lng, 0) / items.length,
  };
}

function SelectedHotspotCard({ hotspot }) {
  if (!hotspot) return null;
  const metrics = [
    [formatter.format(hotspot.violations), "records"],
    [formatter.format(Math.round(hotspot.weightedObstruction)), "PCU pressure"],
    [`${Math.round(hotspot.peakShare * 100)}%`, "peak recurrence"],
    [`${Math.round(hotspot.junctionShare * 100)}%`, "junction exposure"],
  ];
  return (
    <article className="selected-card">
      <div className="selected-top">
        <div>
          <span>Selected hotspot</span>
          <strong>{hotspot.rank}. {hotspot.area}</strong>
          <small>{hotspot.station} · {hotspot.placeType} · {hotspot.topVehicle}</small>
        </div>
        <ScorePill hotspot={hotspot} />
      </div>
      <div className="mini-metrics">
        {metrics.map(([value, label]) => (
          <div key={label}>
            <strong>{value}</strong>
            <span>{label}</span>
          </div>
        ))}
      </div>
      <p>{hotspot.recommendation}</p>
    </article>
  );
}

function PriorityList({ items, onFocus }) {
  return (
    <section className="panel priority-list">
      <PanelHeader icon={<Route />} title="Today first" detail="highest-impact beats" />
      {items.map((item) => (
        <button key={item.id} onClick={() => onFocus(item)}>
          <b>{item.sequence}</b>
          <span>
            <strong>{item.station}</strong>
            {item.area}
          </span>
          <em>{item.impactScore}</em>
        </button>
      ))}
    </section>
  );
}

function StationLeague({ stations, onPick }) {
  return (
    <section className="panel">
      <PanelHeader icon={<BarChart3 />} title="Station league" detail="burden and junction exposure" />
      <div className="dense-table">
        <div className="dense-row head">
          <span>Station</span>
          <span>Cases</span>
          <span>Risk</span>
        </div>
        {stations.map((station) => (
          <button key={station.station} className="dense-row" onClick={() => onPick(station.station)}>
            <span>{station.station}</span>
            <span>{formatter.format(station.cases)}</span>
            <strong>{station.impactScore}</strong>
          </button>
        ))}
      </div>
    </section>
  );
}

function PanelHeader({ icon, title, detail }) {
  return (
    <div className="panel-header">
      <div>
        <span>{React.cloneElement(icon, { size: 18 })}</span>
        <strong>{title}</strong>
      </div>
      {detail && <small>{detail}</small>}
    </div>
  );
}

function Kpi({ title, value, detail, icon, className = "" }) {
  return (
    <article className={`kpi-card ${className}`}>
      <div>
        <span>{title}</span>
        <strong>{value}</strong>
        {detail && <small>{detail}</small>}
      </div>
      {icon && <em>{React.cloneElement(icon, { size: 20 })}</em>}
    </article>
  );
}

function Badge({ icon, label }) {
  return (
    <div className="badge">
      {icon}
      {label}
    </div>
  );
}

function LiquidPanel({ children, className = "", variant = "hero" }) {
  return (
    <div className={`liquid-panel-slot liquid-${variant} ${className}`}>
      <LiquidGlass
        className="liquid-glass-wrap"
        displacementScale={variant === "hero" ? 54 : 38}
        blurAmount={0.13}
        saturation={155}
        aberrationIntensity={1.8}
        elasticity={0.22}
        cornerRadius={22}
        padding="0"
        overLight
        mode="standard"
        style={{ position: "absolute", top: "50%", left: "50%" }}
      >
        {children}
      </LiquidGlass>
    </div>
  );
}

function Signal({ value, label }) {
  return (
    <div>
      <strong>{value}</strong>
      <span>{label}</span>
    </div>
  );
}

function ScorePill({ hotspot }) {
  return (
    <div className="score-pill" style={{ "--score-color": priorityColors[hotspot.priority] || priorityColors.Routine }}>
      <strong>{hotspot.impactScore}</strong>
      <span>{hotspot.priority}</span>
    </div>
  );
}

function LegendDot({ color, label }) {
  return (
    <span className="legend-dot">
      <b style={{ backgroundColor: color }} />
      {label}
    </span>
  );
}

function CoverageCurve({ points, active }) {
  const width = 760;
  const height = 270;
  const maxN = Math.max(...points.map((point) => point.n), 1);
  const path = points
    .map((point, index) => {
      const x = (point.n / maxN) * width;
      const y = height - (point.impact / 100) * height;
      return `${index === 0 ? "M" : "L"} ${x.toFixed(1)} ${y.toFixed(1)}`;
    })
    .join(" ");
  const activePoint = points.reduce((best, point) => (Math.abs(point.n - active) < Math.abs(best.n - active) ? point : best), points[0]);
  const activeX = (activePoint.n / maxN) * width;
  const activeY = height - (activePoint.impact / 100) * height;
  return (
    <div className="curve-wrap">
      <svg viewBox={`0 0 ${width} ${height}`} role="img" aria-label="Coverage curve">
        <defs>
          <linearGradient id="curveFill" x1="0" x2="0" y1="0" y2="1">
            <stop offset="0%" stopColor="#31d6b7" stopOpacity="0.38" />
            <stop offset="100%" stopColor="#31d6b7" stopOpacity="0" />
          </linearGradient>
        </defs>
        <path d={`${path} L ${width} ${height} L 0 ${height} Z`} fill="url(#curveFill)" />
        <path d={path} fill="none" stroke="#31d6b7" strokeWidth="5" strokeLinecap="round" />
        <line x1={activeX} x2={activeX} y1="0" y2={height} stroke="#ffb020" strokeDasharray="7 8" />
        <circle cx={activeX} cy={activeY} r="9" fill="#ffb020" stroke="#fff" strokeWidth="3" />
      </svg>
    </div>
  );
}

function FeatureStack() {
  const features = [
    ["Weighted obstruction", 34, "#31d6b7"],
    ["Density", 18, "#6ea8ff"],
    ["Junction exposure", 15, "#ffb020"],
    ["Arterial obstruction", 13, "#ff4d5e"],
    ["Peak recurrence", 10, "#8d7cf6"],
    ["Active days + severity", 10, "#9ca3af"],
  ];
  return (
    <div className="feature-stack">
      {features.map(([label, value, color]) => (
        <div key={label}>
          <span>{label}</span>
          <b>{value}%</b>
          <em style={{ width: `${value * 2.4}%`, backgroundColor: color }} />
        </div>
      ))}
    </div>
  );
}

function Check({ label, value }) {
  return (
    <div>
      <CheckCircle2 size={18} />
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
}

function AuthGate({ onAuthenticated }) {
  const [mode, setMode] = useState("login");
  const [form, setForm] = useState({ name: "", email: "", password: "", confirmPassword: "" });
  const [errors, setErrors] = useState({});
  const [status, setStatus] = useState("");
  const [isSubmitting, setIsSubmitting] = useState(false);

  const updateField = (field, value) => {
    setForm((current) => ({ ...current, [field]: value }));
    setErrors((current) => ({ ...current, [field]: "", form: "" }));
    setStatus("");
  };

  const validateAuth = () => {
    const nextErrors = {};
    const name = form.name.trim();
    const email = form.email.trim();
    const password = form.password;
    const confirmPassword = form.confirmPassword;
    if (mode === "signup" && !name) nextErrors.name = "Name is required.";
    else if (mode === "signup" && name.length < 2) nextErrors.name = "Enter your full name.";
    if (!email) nextErrors.email = "Email is required.";
    else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) nextErrors.email = "Enter a valid email address.";
    if (!password) nextErrors.password = "Password is required.";
    else if (mode === "signup" && password.length < 8) nextErrors.password = "Use at least 8 characters for signup.";
    else if (mode === "signup" && !/\d/.test(password)) nextErrors.password = "Add at least one number.";
    else if (mode === "login" && password.length < 6) nextErrors.password = "Password must be at least 6 characters.";
    if (mode === "signup" && !confirmPassword) nextErrors.confirmPassword = "Re-enter your password.";
    else if (mode === "signup" && confirmPassword !== password) nextErrors.confirmPassword = "Passwords do not match.";
    setErrors(nextErrors);
    return Object.keys(nextErrors).length === 0;
  };

  const submitAuth = async (event) => {
    event.preventDefault();
    if (!validateAuth()) return;
    setIsSubmitting(true);
    setStatus(mode === "signup" ? "Creating secure workspace..." : "Checking database credentials...");
    try {
      const session = await submitAuthRequest(mode, form);
      setStatus(mode === "signup" ? "Workspace created. Opening ParkSight..." : "Credentials verified. Opening ParkSight...");
      window.setTimeout(() => onAuthenticated(session), 320);
    } catch (error) {
      setStatus("");
      setErrors((current) => ({ ...current, form: error.message || "Authentication failed." }));
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <main className="auth-gate">
      <video className="auth-video" src="/media/smart-city-login.mp4" autoPlay loop muted playsInline aria-hidden="true" />
      <div className="auth-video-scrim" />
      <section className="auth-card-shell" aria-label="ParkSight access">
        <form className="auth-card" onSubmit={submitAuth} noValidate>
          <div className="auth-brand">
            <span className="brand-mark"><Radar size={24} /></span>
            <div>
              <p className="eyebrow">ParkSight AI</p>
              <h1>{mode === "login" ? "Welcome back" : "Create access"}</h1>
            </div>
          </div>

          <div className="auth-switch" style={{ "--auth-mode": mode === "signup" ? 1 : 0 }}>
            <button type="button" className={mode === "login" ? "active" : ""} onClick={() => { setMode("login"); setErrors({}); setStatus(""); }} disabled={isSubmitting}>
              Login
            </button>
            <button type="button" className={mode === "signup" ? "active" : ""} onClick={() => { setMode("signup"); setErrors({}); setStatus(""); }} disabled={isSubmitting}>
              Signup
            </button>
          </div>

          {mode === "signup" && (
            <label className={`auth-field ${errors.name ? "has-error" : ""}`}>
              <span>Full name</span>
              <div>
                <ShieldCheck size={18} />
                <input
                  type="text"
                  value={form.name}
                  onChange={(event) => updateField("name", event.target.value)}
                  placeholder="Rajdeep Sharma"
                  autoComplete="name"
                  disabled={isSubmitting}
                />
              </div>
              {errors.name && <small>{errors.name}</small>}
            </label>
          )}

          <label className={`auth-field ${errors.email ? "has-error" : ""}`}>
            <span>Email address</span>
            <div>
              <Mail size={18} />
              <input
                type="email"
                value={form.email}
                onChange={(event) => updateField("email", event.target.value)}
                placeholder="officer@parksight.ai"
                autoComplete="email"
                disabled={isSubmitting}
              />
            </div>
            {errors.email && <small>{errors.email}</small>}
          </label>

          <label className={`auth-field ${errors.password ? "has-error" : ""}`}>
            <span>Password</span>
            <div>
              <LockKeyhole size={18} />
              <input
                type="password"
                value={form.password}
                onChange={(event) => updateField("password", event.target.value)}
                placeholder={mode === "signup" ? "8+ chars and a number" : "Enter password"}
                autoComplete={mode === "signup" ? "new-password" : "current-password"}
                disabled={isSubmitting}
              />
            </div>
            {errors.password && <small>{errors.password}</small>}
          </label>

          {mode === "signup" && (
            <label className={`auth-field ${errors.confirmPassword ? "has-error" : ""}`}>
              <span>Re-enter password</span>
              <div>
                <LockKeyhole size={18} />
                <input
                  type="password"
                  value={form.confirmPassword}
                  onChange={(event) => updateField("confirmPassword", event.target.value)}
                  placeholder="Confirm password"
                  autoComplete="new-password"
                  disabled={isSubmitting}
                />
              </div>
              {errors.confirmPassword && <small>{errors.confirmPassword}</small>}
            </label>
          )}

          <button className="auth-submit" type="submit" disabled={isSubmitting}>
            <span />
            <b>{isSubmitting ? "Please wait" : mode === "signup" ? "Create account" : "Login"}</b>
            <LogIn size={18} />
          </button>
          {errors.form && <p className="auth-status is-error">{errors.form}</p>}
          {status && <p className="auth-status">{status}</p>}
        </form>
      </section>
      <SiteFooter />
    </main>
  );
}

function UploadGate({ user, uploadNotice, uploadError, onDatasetUpload, onLoadDemoDataset, onLogout }) {
  return (
    <div className="upload-gate">
      <div className="upload-user-bar">
        <UserChip user={user} />
        <button onClick={onLogout} className="soft-button logout-button">
          <LogOut size={15} />
          Logout
        </button>
      </div>
      <section className="upload-gate-card">
        <div className="brand-mark">
          <Radar size={26} />
        </div>
        <p className="eyebrow">ParkSight AI prototype</p>
        <h1>Upload a parking dataset to generate the dashboard.</h1>
        <p>
          The command center, maps, deployment plan, action queue, and audit trail are created only after a CSV or JSON dataset is analyzed.
        </p>
        <label className="primary-button upload-button">
          <FileUp size={17} />
          Upload dataset
          <input
            type="file"
            accept=".csv,.json,application/json,text/csv"
            onChange={(event) => {
              onDatasetUpload(event.target.files?.[0]);
              event.target.value = "";
            }}
          />
        </label>
        <button type="button" className="soft-button" onClick={onLoadDemoDataset}>
          <Sparkles size={16} />
          Load judge demo
        </button>
        {(uploadNotice || uploadError) && <span className={`upload-status ${uploadError ? "is-error" : ""}`}>{uploadError || uploadNotice}</span>}
      </section>
      <SiteFooter />
    </div>
  );
}

function buildIntelligence(data, fleetSize, rankMode) {
  const totalHotspotCases = sum(data.hotspots, "violations");
  const totalHotspotImpact = sum(data.hotspots, "impactScore");
  const rankKey = rankMode === "junction" ? "junctionShare" : rankMode === "violations" ? "violations" : "impactScore";
  const ranked = [...data.hotspots].sort((a, b) => (b[rankKey] || 0) - (a[rankKey] || 0));
  const selected = ranked.slice(0, fleetSize);
  const countRanked = [...data.hotspots].sort((a, b) => b.violations - a.violations).slice(0, fleetSize);
  const selectedImpactPct = Math.round((sum(selected, "impactScore") / totalHotspotImpact) * 1000) / 10;
  const countImpactPct = Math.round((sum(countRanked, "impactScore") / totalHotspotImpact) * 1000) / 10;
  const planHotspots = data.enforcementPlan.map((item) => data.hotspots.find((hotspot) => hotspot.id === item.hotspotId)).filter(Boolean);
  const deployment = selected.map((hotspot, index) => {
    const previous = index ? selected[index - 1] : null;
    const legKm = previous ? haversineKm(previous, hotspot) : 0;
    return {
      ...hotspot,
      sequence: index + 1,
      legKm: Number(legKm.toFixed(2)),
      unit: `Unit ${String(index + 1).padStart(2, "0")}`,
      shift: classifyShift(hotspot),
      window: hotspot.topHours?.[0] ? `${formatHourLabel(hotspot.topHours[0].hour)} field cycle` : "Peak field cycle",
    };
  });
  const routeKm = deployment.reduce((total, item) => total + item.legKm, 0).toFixed(1);
  const curve = buildCurve(data.hotspots);
  const vehiclePressure = data.charts.vehicles.map((item) => ({
    ...item,
    pressure: item.count * (vehicleWeights[item.name] || 1),
  })).sort((a, b) => b.pressure - a.pressure);
  return {
    planCases: sum(planHotspots, "violations"),
    planCoverage: Math.round((sum(planHotspots, "violations") / totalHotspotCases) * 100),
    stationLeague: [...data.stations].sort((a, b) => b.impactScore - a.impactScore),
    deployment,
    routeKm,
    vehicleHotspots: [...data.hotspots].sort((a, b) => b.weightedObstruction - a.weightedObstruction),
    vehiclePlaybooks: vehiclePressure.slice(0, 5).map((item) => ({ ...item, playbook: playbookForVehicle(item.name) })),
    workOrders: data.hotspots.slice(0, 36).map(buildWorkOrder),
    curve,
    simulation: {
      violations: sum(selected, "violations"),
      impact: sum(selected, "impactScore"),
      violPct: Math.round((sum(selected, "violations") / totalHotspotCases) * 1000) / 10,
      impactPct: selectedImpactPct,
      critical: selected.filter((item) => item.priority === "Critical" || item.priority === "High").length,
      liftPts: Math.max(0, Math.round((selectedImpactPct - countImpactPct) * 10) / 10),
    },
  };
}

function buildWorkOrder(hotspot) {
  const structural = hotspot.junctionShare > 0.7 || hotspot.placeType?.toLowerCase().includes("market");
  const review = hotspot.approvedShare < 0.25;
  return {
    id: hotspot.id,
    rank: hotspot.rank,
    area: hotspot.area,
    impactScore: hotspot.impactScore,
    owner: review ? "Review Cell" : structural ? "DULT" : "Traffic Police",
    action: review
      ? "Audit booking quality and location tags"
      : structural
        ? "Add no-stopping protection, bay marking, and junction clearance"
        : "Schedule tow cycle and peak patrol",
  };
}

function buildCurve(hotspots) {
  const ranked = [...hotspots].sort((a, b) => b.impactScore - a.impactScore);
  const total = sum(ranked, "impactScore") || 1;
  const points = [{ n: 0, impact: 0 }];
  let running = 0;
  const steps = 36;
  for (let i = 1; i <= steps; i += 1) {
    const n = Math.max(1, Math.round((ranked.length * i) / steps));
    running = sum(ranked.slice(0, n), "impactScore");
    points.push({ n, impact: Math.round((running / total) * 1000) / 10 });
  }
  return points;
}

function classifyShift(hotspot) {
  const hour = hotspot.topHours?.[0]?.hour ?? 10;
  if (hour >= 7 && hour <= 11) return "Peak";
  if (hour >= 17 && hour <= 21) return "Evening";
  if (hour < 7) return "Early";
  return "Midday";
}

function playbookForVehicle(name) {
  if (/SCOOTER|MOTOR|TWO/.test(name)) return "Use bay marking, bollards at corners, and repeat-location warnings before tow escalation.";
  if (/CAR|JEEP|TAXI/.test(name)) return "Prioritize tow-ready patrols near markets, hospitals, and junction mouths during peak recurrence windows.";
  if (/AUTO/.test(name)) return "Create short-stay pick-up pockets and clear no-standing edges around signal approaches.";
  if (/BUS|TRUCK|LORRY|TEMPO/.test(name)) return "Escalate to obstruction removal because one vehicle consumes multiple lane-equivalent units.";
  return "Use station-level patrol reminders and location-specific signage.";
}

function hourChartData(data) {
  return {
    labels: data.charts.hours.map((item) => formatHourLabel(item.hour)),
    datasets: [
      {
        label: "Violations",
        data: data.charts.hours.map((item) => item.count),
        borderColor: "#31d6b7",
        backgroundColor: "rgba(49, 214, 183, 0.16)",
        fill: true,
        tension: 0.35,
        pointRadius: 0,
      },
    ],
  };
}

function vehicleDoughnutData(data) {
  const values = data.charts.vehicles.slice(0, 7);
  return {
    labels: values.map((item) => item.name),
    datasets: [{ data: values.map((item) => item.count), backgroundColor: ["#31d6b7", "#ffb020", "#6ea8ff", "#ff4d5e", "#8d7cf6", "#58c4dd", "#a3e635"], borderWidth: 0 }],
  };
}

function vehiclePressureData(data) {
  const values = data.charts.vehicles
    .map((item) => ({ ...item, pressure: item.count * (vehicleWeights[item.name] || 1) }))
    .sort((a, b) => b.pressure - a.pressure)
    .slice(0, 8)
    .reverse();
  return {
    labels: values.map((item) => item.name),
    datasets: [{ label: "PCU pressure", data: values.map((item) => Math.round(item.pressure)), backgroundColor: "#31d6b7", borderRadius: 8 }],
  };
}

function stationChartData(data) {
  const values = data.charts.stations.slice(0, 8).reverse();
  return {
    labels: values.map((item) => item.name),
    datasets: [{ label: "Cases", data: values.map((item) => item.count), backgroundColor: "#ff4d5e", borderRadius: 8 }],
  };
}

const axisColor = "#65707d";
const gridColor = "rgba(101,112,125,0.16)";

const lineOptions = {
  maintainAspectRatio: false,
  plugins: { legend: { display: false } },
  scales: {
    x: { grid: { display: false }, ticks: { color: axisColor, maxRotation: 0, autoSkip: true, maxTicksLimit: 7 } },
    y: { grid: { color: gridColor }, ticks: { color: axisColor } },
  },
};

const doughnutOptions = {
  maintainAspectRatio: false,
  plugins: { legend: { position: "bottom", labels: { boxWidth: 10, color: axisColor } } },
  cutout: "62%",
};

const barOptions = {
  indexAxis: "y",
  maintainAspectRatio: false,
  plugins: { legend: { display: false } },
  scales: {
    x: { grid: { color: gridColor }, ticks: { color: axisColor } },
    y: { grid: { display: false }, ticks: { color: axisColor } },
  },
};

function getHeatPoints(data, layer) {
  if (layer === "junction") {
    const byLocation = new Map(data.hotspots.map((hotspot) => [`${hotspot.lat},${hotspot.lng}`, hotspot.junctionShare]));
    return data.heatmap.map(([lat, lng, intensity]) => {
      const junctionShare = byLocation.get(`${lat},${lng}`) || 0.03;
      return [lat, lng, Math.min(1, Math.pow(junctionShare, 0.55) * 1.15 + intensity * 0.08)];
    });
  }
  if (layer === "violations") {
    const maxCount = Math.max(...data.heatmap.map((point) => point[3]));
    return data.heatmap.map(([lat, lng, , count]) => [lat, lng, Math.min(1, Math.pow(count / maxCount, 0.38) * 1.12)]);
  }
  return data.heatmap.map(([lat, lng, intensity]) => [lat, lng, intensity]);
}

function getMarkerRadius(hotspot, layer, data) {
  if (layer === "violations") {
    const maxViolations = data.hotspots[0]?.violations || 1;
    return 7 + Math.pow(hotspot.violations / maxViolations, 0.45) * 14;
  }
  if (layer === "junction") return 7 + Math.pow(hotspot.junctionShare, 0.55) * 13;
  return 8 + Math.pow(hotspot.impactScore / 100, 1.2) * 13;
}

function getMarkerColor(hotspot, layer, data) {
  if (layer === "violations") {
    const maxViolations = data.hotspots[0]?.violations || 1;
    const ratio = hotspot.violations / maxViolations;
    if (ratio > 0.72) return "#152c8f";
    if (ratio > 0.38) return "#3267d6";
    return "#6ea8ff";
  }
  if (layer === "junction") {
    if (hotspot.junctionShare > 0.8) return "#073f45";
    if (hotspot.junctionShare > 0.45) return "#04756f";
    return "#56c9ad";
  }
  return priorityColors[hotspot.priority] || priorityColors.Routine;
}

function haversineKm(a, b) {
  const toRad = (value) => (value * Math.PI) / 180;
  const earth = 6371;
  const dLat = toRad(b.lat - a.lat);
  const dLng = toRad(b.lng - a.lng);
  const lat1 = toRad(a.lat);
  const lat2 = toRad(b.lat);
  const h = Math.sin(dLat / 2) ** 2 + Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLng / 2) ** 2;
  return 2 * earth * Math.asin(Math.sqrt(h));
}

function formatHourLabel(hour) {
  if (hour === 0) return "12 AM";
  if (hour === 12) return "12 PM";
  return hour < 12 ? `${hour} AM` : `${hour - 12} PM`;
}

function groupBy(items, key) {
  return items.reduce((acc, item) => {
    const group = item[key];
    acc[group] = acc[group] || [];
    acc[group].push(item);
    return acc;
  }, {});
}

function countBy(items, fn) {
  return items.reduce((acc, item) => {
    const key = fn(item);
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function sum(items, key) {
  return items.reduce((total, item) => total + (Number(item[key]) || 0), 0);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

createRoot(document.getElementById("root")).render(<App />);
