import http from 'node:http';
import { readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const PORT = Number(process.env.PORT || 3000);
const CSV_PATH = path.join(ROOT, 'upload', '01-_-5-_20250630-_-5-_20250630.csv');

// Load the local .env file without adding a runtime dependency.
const ENV_PATH = path.join(ROOT, '.env');
if (existsSync(ENV_PATH)) {
  const envText = await readFile(ENV_PATH, 'utf8');
  for (const line of envText.split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Z][A-Z0-9_]*)\s*=\s*["']?([^"']*)["']?\s*$/);
    if (match && !process.env[match[1]]) process.env[match[1]] = match[2];
  }
}

function parseCsv(text) {
  const rows = [];
  let row = [], cell = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i], n = text[i + 1];
    if (c === '"' && quoted && n === '"') { cell += '"'; i++; continue; }
    if (c === '"') { quoted = !quoted; continue; }
    if (c === ',' && !quoted) { row.push(cell); cell = ''; continue; }
    if ((c === '\n' || c === '\r') && !quoted) {
      if (c === '\r' && n === '\n') i++;
      row.push(cell); cell = '';
      if (row.some(Boolean)) rows.push(row);
      row = [];
      continue;
    }
    cell += c;
  }
  if (cell || row.length) { row.push(cell); rows.push(row); }
  const [header, ...data] = rows;
  return data.map(r => Object.fromEntries(header.map((h, i) => [h, r[i] ?? ''])));
}

const raw = existsSync(CSV_PATH) ? await readFile(CSV_PATH, 'utf8') : '';
const incidents = parseCsv(raw).filter(r => r['발생일'] && r['발생역']).map(r => ({
  date: r['발생일'], type: r['사고유형'], station: r['발생역'], time: r['발생시간']
}));
const stations = [...new Set(incidents.map(r => r.station))].sort((a, b) => a.localeCompare(b, 'ko'));

function hourOf(time) { return Number(String(time).split(':')[0]); }
function inWindow(time, target, windowHours = 1) {
  const h = hourOf(time), t = Number(String(target).split(':')[0]);
  const distance = Math.abs(h - t);
  return Math.min(distance, 24 - distance) <= windowHours;
}

function statsFor(routeStations, targetTime) {
  const route = routeStations.filter(Boolean);
  const selected = incidents.filter(r => route.includes(r.station));
  const timeMatched = selected.filter(r => inWindow(r.time, targetTime, 1));
  const byStation = Object.fromEntries(route.map(s => [s, {
    all: selected.filter(r => r.station === s).length,
    timeMatched: timeMatched.filter(r => r.station === s).length
  }]));
  const byType = Object.entries(selected.reduce((a, r) => {
    a[r.type] = (a[r.type] || 0) + 1; return a;
  }, {})).sort((a, b) => b[1] - a[1]);
  return {
    sourcePeriod: ['2021-01-02', '2025-06-20'],
    sourceIncidentCount: incidents.length,
    routeStations: route,
    routeIncidentCount: selected.length,
    timeWindow: `${targetTime} ± 1 hour`,
    timeMatchedIncidentCount: timeMatched.length,
    byStation,
    byType: Object.fromEntries(byType),
    overallHourlyDistribution: incidents.reduce((a, r) => {
      const h = hourOf(r.time); a[h] = (a[h] || 0) + 1; return a;
    }, {})
  };
}

async function askJev(state) {
  if (!process.env.TYPESAFE_API_KEY) {
    return { configured: false, error: 'TYPESAFE_API_KEY가 설정되지 않았습니다.' };
  }
  const response = await fetch('https://api.typesafe.ai/v1/systemone', {
    method: 'POST',
    headers: {
      authorization: `Bearer ${process.env.TYPESAFE_API_KEY}`,
      'content-type': 'application/json'
    },
    body: JSON.stringify({
      model: 'jev-latest',
      state,
      questions: {
        incident_risk: {
          type: 'noul',
          instructions: '주어진 이동경로와 시간대에 사고가 발생할 위험이 역사적 기준보다 높은가?',
          criteria: {
            true: '해당 경로의 역별 사고기록과 선택 시간대의 사고기록을 함께 고려했을 때 역사적 평균보다 높은 상대위험으로 볼 수 있음',
            false: '해당 경로의 사고기록과 선택 시간대의 사고기록이 역사적 평균보다 높다고 보기 어려움'
          }
        },
        severity: {
          type: 'score',
          instructions: '이 이동 조건에서 예상되는 사고 영향의 수준은 어느 정도인가?',
          criteria: [
            '낮음: 사고기록이 적고 시간대 신호도 약함',
            '중간: 사고기록 또는 시간대 신호 중 하나가 뚜렷함',
            '높음: 경로와 시간대 모두에서 사고 신호가 뚜렷함'
          ]
        }
      }
    })
  });
  const body = await response.text();
  if (!response.ok) throw new Error(`TypeSafe API ${response.status}: ${body}`);
  return { configured: true, result: JSON.parse(body) };
}

function json(res, status, body) {
  res.writeHead(status, { 'content-type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(body));
}

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === 'GET' && req.url === '/api/meta') {
      return json(res, 200, { stations, incidentCount: incidents.length, period: ['2021-01-02', '2025-06-20'] });
    }
    if (req.method === 'POST' && req.url === '/api/analyze') {
      let body = ''; for await (const chunk of req) body += chunk;
      const input = JSON.parse(body || '{}');
      const routeStations = input.routeStations?.length ? input.routeStations : [input.origin, input.destination];
      if (!input.origin || !input.destination || !input.time) return json(res, 400, { error: '출발역·도착역·시간을 입력하세요.' });
      const stats = statsFor(routeStations, input.time);
      const state = JSON.stringify({ query: input, historicalEvidence: stats });
      const jev = await askJev(state);
      return json(res, 200, { query: input, stats, jev });
    }
    const file = req.url === '/' ? '/public/index.html' : req.url;
    const filePath = path.join(ROOT, file);
    const content = await readFile(filePath);
    const type = filePath.endsWith('.html') ? 'text/html' : 'text/plain';
    res.writeHead(200, { 'content-type': `${type}; charset=utf-8` }); res.end(content);
  } catch (error) { json(res, 500, { error: error.message }); }
});

server.listen(PORT, () => console.log(`Jev subway risk app: http://localhost:${PORT}`));
