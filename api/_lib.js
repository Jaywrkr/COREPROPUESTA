// Utilidades compartidas de la API (los archivos con "_" no son rutas en Vercel).
const crypto = require('crypto');

const REPO = process.env.GITHUB_REPO || 'Jaywrkr/COREPROPUESTA';
const TOKEN = process.env.GITHUB_TOKEN || '';

function same(a, b) {
  const x = Buffer.from(String(a || '')), y = Buffer.from(String(b || ''));
  return x.length === y.length && x.length > 0 && crypto.timingSafeEqual(x, y);
}
const isApp = (req) => same(req.headers['x-clave'], process.env.APP_CLAVE);
const isAdmin = (req) => same(req.headers['x-admin'], process.env.ADMIN_CLAVE);
const configured = () => Boolean(TOKEN && process.env.APP_CLAVE);

async function gh(path, opts = {}) {
  const r = await fetch(`https://api.github.com/repos/${REPO}${path}`, {
    ...opts,
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'User-Agent': 'corepropuesta',
      ...(opts.body ? { 'Content-Type': 'application/json' } : {}),
    },
  });
  const text = await r.text();
  let data = null; try { data = text ? JSON.parse(text) : null; } catch (_) { data = text; }
  if (!r.ok) { const e = new Error(`GitHub ${r.status}: ${(data && data.message) || text}`); e.status = r.status; throw e; }
  return data;
}

const clip = (v, n) => String(v == null ? '' : v).slice(0, n);
const send = (res, code, obj) => { res.statusCode = code; res.setHeader('Content-Type', 'application/json; charset=utf-8'); res.setHeader('Cache-Control', 'no-store'); res.end(JSON.stringify(obj)); };
function body(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  try { return JSON.parse(req.body || '{}'); } catch (_) { return {}; }
}
// Datos de máquina dentro del issue: <!--cp:BASE64(JSON)-->
const pack = (o) => `<!--cp:${Buffer.from(JSON.stringify(o), 'utf8').toString('base64')}-->`;
function unpack(text) {
  const m = /<!--cp:([A-Za-z0-9+/=]+)-->/.exec(String(text || ''));
  if (!m) return null; try { return JSON.parse(Buffer.from(m[1], 'base64').toString('utf8')); } catch (_) { return null; }
}

module.exports = { REPO, gh, isApp, isAdmin, configured, clip, send, body, pack, unpack };
