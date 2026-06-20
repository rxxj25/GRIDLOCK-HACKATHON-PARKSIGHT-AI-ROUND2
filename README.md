# ParkSight AI

AI-driven parking intelligence for congestion-aware enforcement.

ParkSight AI turns raw illegal-parking violation records into a judge-ready command dashboard for traffic police. It identifies repeated illegal-parking hotspots, estimates congestion impact, ranks enforcement priority, and generates action-ready deployment and audit outputs.

## Live Demo

GitHub Pages:

```text
https://rxxj25.github.io/GRIDLOCK-HACKATHON-PARKSIGHT-AI-ROUND2/
```

## Recording

Recording:
```text
https://drive.google.com/file/d/1EGVFbq6H43lNT9DdP1bN9-uO_k3l3pxQ/view?pli=1
```

Judge demo login:

```text
Email: officer@parksight.ai
Password: Password1
```

After login, click **Load judge demo** to open the precomputed dashboard immediately. The hosted GitHub Pages build is static, so the demo uses the bundled `public/data/parking_intelligence.json` intelligence file. Local development can also run the Express + PostgreSQL backend for signup, login, and persisted dataset uploads.

## Project Screenshots

These screenshots are included here so judges can quickly inspect the working dashboard, enforcement workflow, and backend persistence design from the repository front page.

### Command Dashboard

![ParkSight AI command dashboard](docs/screenshots/dashboard-command.png)

### Hotspot Map and Congestion Evidence

![ParkSight AI hotspot map](docs/screenshots/hotspot-map.png)

### Vehicle Intelligence

![Vehicle intelligence dashboard](docs/screenshots/vehicle-intel.png)

### Deployment Plan

![Deployment plan dashboard](docs/screenshots/deployment-plan.png)

### Action Queue and Audit Trail

![Action queue](docs/screenshots/action-queue.png)

![Audit trail and model transparency](docs/screenshots/audit-trail.png)

### Backend and PostgreSQL

The full-stack mode uses a Node.js/Express API with PostgreSQL persistence. The database stores authenticated users, dataset upload sessions, generated hotspot rankings, station summaries, and repeated number-plate intelligence.

![Backend and PostgreSQL schema](docs/screenshots/backend-postgresql-schema.svg)

## Problem

Illegal on-street parking near markets, metro stations, hospitals, and junctions reduces usable road width and blocks turning or crossing movements. Enforcement teams often know where tickets are issued, but not which hotspots create the highest traffic-flow risk.

ParkSight AI answers:

1. Where is illegal parking repeating?
2. Which hotspots are most likely to affect congestion?
3. Which police station areas should be prioritized?
4. What action should enforcement teams take first?

## What Judges Should Try

1. Open the live demo.
2. Log in with the judge demo credentials.
3. Click **Load judge demo**.
4. Review the Command view for summary, impact index, and enforcement plan.
5. Open Hotspot map and switch between Impact, Volume, and Junction layers.
6. Check Vehicle intel, Number plates, Deployment, What-if lab, Action queue, and Audit trail.
7. Use the CSV/PDF export buttons in the top bar.

## Key Features

- Interactive illegal-parking hotspot map for Bengaluru.
- Heatmap modes for congestion impact, violation volume, and junction exposure.
- Police-station filtering with map focus.
- Explainable Parking Impact Index instead of black-box scoring.
- Vehicle obstruction weighting inspired by passenger-car-unit pressure.
- Hotspot ranking, enforcement beat generation, and action-owner assignment.
- What-if simulator for fleet size and ranking strategy.
- Number-plate recurrence intelligence.
- Downloadable CSV enforcement report.
- Downloadable PDF enforcement brief.
- Static GitHub Pages demo plus full local backend mode.

## AI and Analytics Method

The supplied source data contains violation records, not live traffic speed labels. ParkSight AI therefore uses an explainable geospatial analytics pipeline rather than claiming a supervised congestion model.

Pipeline:

1. Geospatial grouping into approximately 220 m urban cells.
2. Feature engineering for density, vehicle pressure, junction exposure, arterial obstruction, peak recurrence, active-day recurrence, and severity.
3. Vehicle obstruction weighting so buses, trucks, cars, autos, and two-wheelers contribute different road-space pressure.
4. Violation severity weighting for risky categories such as double parking, zebra-crossing obstruction, and traffic-light/junction parking.
5. Explainable Parking Impact Index ranking.
6. Enforcement plan generation from the top impact-ranked hotspots.

Parking Impact Index:

```text
100 * (
  0.34 weighted obstruction
  + 0.18 density
  + 0.15 junction exposure
  + 0.13 arterial obstruction
  + 0.10 peak recurrence
  + 0.06 active-day recurrence
  + 0.04 severity
)
```

This keeps the output auditable for civic and traffic-enforcement teams.

## Dataset

The original local CSV is not committed because it is larger than GitHub's normal file limit:

```text
jan to may police violation_anonymized791b166 (1).csv
```

The repository includes the generated intelligence artifact used by the hosted demo:

```text
public/data/parking_intelligence.json
```

Demo dataset summary:

- 298,450 geocoded parking-violation records.
- Bengaluru region.
- Actual record date range: 10 Nov 2023 to 08 Apr 2024.
- 3,969 scored urban cells.
- 160 dashboard hotspots.
- 12 deployable enforcement beats.

## Technology Stack

Frontend:

- React 19
- Vite 8
- Tailwind CSS
- CSS glassmorphism and responsive dashboard layout
- Framer Motion
- Liquid Glass React
- Lucide React icons

Maps and visualization:

- Leaflet
- Leaflet.heat
- Chart.js
- React Chart.js 2

Reports and exports:

- jsPDF
- CSV generation in browser

Backend and data persistence:

- Node.js
- Express 5
- PostgreSQL
- `pg`
- `bcryptjs`
- `dotenv`
- Session-token based auth helpers

Data processing:

- Python preprocessing scripts
- Client-side Web Worker dataset analysis
- JSON intelligence artifact generation

Deployment and tooling:

- GitHub Actions
- GitHub Pages
- Docker
- Docker Compose
- pgAdmin-compatible PostgreSQL setup

## Local Setup

Install dependencies:

```powershell
npm install
```

Start the frontend and API together:

```powershell
npm run dev
```

Open:

```text
http://localhost:8000
```

The dev script starts:

- Vite app on port `8000`
- Express API on port `8001`

## PostgreSQL Setup

Create `.env.local`:

```powershell
DATABASE_URL=postgresql://postgres:your_password@localhost:5432/parksight_ai
API_PORT=8001
APP_SESSION_SECRET=replace_with_a_long_random_secret
OFFLINE_LOGIN_EMAIL=officer@parksight.ai
OFFLINE_LOGIN_PASSWORD=Password1
```

Initialize the database:

```powershell
npm run db:setup
```

Tables created:

- `users`
- `dataset_uploads`
- `upload_hotspots`
- `upload_stations`
- `upload_plates`

## Docker

Build the app image:

```powershell
docker build -t parksight-ai .
```

Run the full stack with PostgreSQL:

```powershell
copy .env.docker.example .env.docker
docker compose --env-file .env.docker up --build
```

Open:

```text
http://localhost:8000
```

## Build

```powershell
npm run build
```

## GitHub Pages Deployment

The repository includes `.github/workflows/deploy.yml`. On every push to `main`, GitHub Actions installs dependencies, runs the Vite build, uploads `dist`, and deploys it to GitHub Pages.

The Vite config uses a relative base path so the app works under the GitHub Pages repository URL.

## Project Structure

```text
.
|-- .github/workflows/deploy.yml
|-- public/
|   |-- data/parking_intelligence.json
|   |-- media/smart-city-login.mp4
|   `-- parksight-icon.svg
|-- scripts/
|   `-- setup-db.mjs
|-- server/
|   |-- index.js
|   |-- db.js
|   `-- auth.js
|-- src/
|   |-- App.jsx
|   |-- datasetUpload.js
|   |-- datasetWorker.js
|   |-- index.css
|   `-- reporting.js
|-- docker-compose.yml
|-- Dockerfile
|-- package.json
|-- vite.config.js
`-- README.md
```

## Team

Made by Rajdeep Bandyopadhaya and Aniket Arya. All rights reserved.
