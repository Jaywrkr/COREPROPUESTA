// POST {codigo, texto, accion, desc, numero?, eliminar?}: crea/edita/borra una regla en reglas.json
// con un commit (Vercel publica solo). Si viene "numero", comenta y cierra ese issue (requiere x-admin).
const { gh, isAdmin, configured, clip, send, body } = require('./_lib');
module.exports = async (req, res) => {
  try {
    if (req.method !== 'POST') return send(res, 405, { error: 'Método no permitido' });
    if (!configured()) return send(res, 503, { error: 'No configurado.' });
    if (!isAdmin(req)) return send(res, 401, { error: 'Clave de administrador inválida.' });
    const b = body(req);
    const codigo = clip(b.codigo, 20).trim().toUpperCase();
    if (!/^[A-Z0-9][A-Z0-9-]{1,19}$/.test(codigo)) return send(res, 400, { error: 'Código Lenovo inválido.' });
    const accion = b.accion === 'ignorar' ? 'ignorar' : 'linea';
    const texto = clip(b.texto, 300);
    if (!b.eliminar && accion === 'linea' && !texto.trim()) return send(res, 400, { error: 'Falta el texto.' });
    for (let intento = 0; intento < 2; intento++) {
      const f = await gh('/contents/reglas.json?ref=main');
      const actual = JSON.parse(Buffer.from(f.content, 'base64').toString('utf8') || '{}');
      const reglas = actual.reglas || {};
      if (b.eliminar) delete reglas[codigo];
      else reglas[codigo] = { texto, accion, desc: clip(b.desc, 200), actualizado: new Date().toISOString().slice(0, 10), ...(b.numero ? { origen: `#${parseInt(b.numero, 10)}` } : {}) };
      const ordenadas = Object.fromEntries(Object.entries(reglas).sort((x, y) => x[0].localeCompare(y[0])));
      const contenido = JSON.stringify({ reglas: ordenadas }, null, 2) + '\n';
      try {
        await gh('/contents/reglas.json', { method: 'PUT', body: JSON.stringify({
          message: `${b.eliminar ? 'Elimina' : 'Regla'} ${codigo}${b.numero ? ` (desde #${parseInt(b.numero, 10)})` : ''}`,
          content: Buffer.from(contenido, 'utf8').toString('base64'), sha: f.sha, branch: 'main' }) });
        break;
      } catch (e) { if (e.status !== 409 || intento === 1) throw e; }
    }
    const n = parseInt(b.numero, 10);
    if (n) {
      await gh(`/issues/${n}/comments`, { method: 'POST', body: JSON.stringify({ body: `Convertido en regla \`${codigo}\`${accion === 'ignorar' ? ' (no mencionar)' : `: ${texto}`}` }) });
      await gh(`/issues/${n}`, { method: 'PATCH', body: JSON.stringify({ state: 'closed', state_reason: 'completed' }) });
    }
    return send(res, 200, { ok: true });
  } catch (e) { return send(res, 500, { error: 'No se pudo guardar la regla en GitHub.', detalle: e.message }); }
};
