// POST {numero, nota}: cierra un issue de la bandeja sin crear regla (requiere x-admin).
const { gh, isAdmin, configured, clip, send, body } = require('./_lib');
module.exports = async (req, res) => {
  try {
    if (req.method !== 'POST') return send(res, 405, { error: 'Método no permitido' });
    if (!configured()) return send(res, 503, { error: 'No configurado.' });
    if (!isAdmin(req)) return send(res, 401, { error: 'Clave de administrador inválida.' });
    const b = body(req); const n = parseInt(b.numero, 10); if (!n) return send(res, 400, { error: 'Falta el número.' });
    await gh(`/issues/${n}/comments`, { method: 'POST', body: JSON.stringify({ body: clip(b.nota || 'Revisado sin cambios.', 1000) }) });
    await gh(`/issues/${n}`, { method: 'PATCH', body: JSON.stringify({ state: 'closed', state_reason: 'completed' }) });
    return send(res, 200, { ok: true });
  } catch (e) { return send(res, 500, { error: 'No se pudo cerrar.', detalle: e.message }); }
};
