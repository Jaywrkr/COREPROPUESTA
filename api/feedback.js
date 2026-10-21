// POST: crea un comentario o corrección como issue (requiere x-clave).
// GET: lista los abiertos para la bandeja de aprendizaje (requiere x-admin).
const { gh, isApp, isAdmin, configured, clip, send, body, pack, unpack } = require('./_lib');

function md(d) {
  const L = [];
  L.push(`**Tipo:** ${d.tipo === 'correccion' ? 'Corrección al descargar' : 'Comentario del vendedor'}`);
  L.push(`**Vendedor:** ${d.autor || '—'}  ·  **Archivo:** \`${d.archivo || '—'}\`  ·  **Configuración:** \`${d.cto || '—'}\` ${d.modelo || ''}`);
  L.push('');
  if (d.tipo === 'comentario') {
    L.push('**Línea generada**'); L.push('```'); L.push(d.linea || '(vacía)'); L.push('```');
    L.push('**Comentario**'); L.push(''); L.push('> ' + String(d.comentario || '').replace(/\n/g, '\n> '));
  } else {
    for (const c of d.cambios || []) {
      L.push(`- ${c.antes ? '`' + c.antes + '`' : '_(no existía)_'}  →  ${c.despues ? '`' + c.despues + '`' : '_(eliminada)_'}`);
    }
  }
  if ((d.codigos || []).length) {
    L.push(''); L.push('**Códigos Lenovo involucrados**');
    for (const c of d.codigos) L.push(`- \`${c.code}\` ${c.desc || ''}`);
  }
  L.push(''); L.push(pack(d));
  return L.join('\n');
}

module.exports = async (req, res) => {
  try {
    if (!configured()) return send(res, 503, { error: 'La bandeja de aprendizaje no está configurada en Vercel (faltan variables de entorno).' });
    if (req.method === 'POST') {
      if (!isApp(req)) return send(res, 401, { error: 'Clave inválida.' });
      const b = body(req);
      const tipo = b.tipo === 'correccion' ? 'correccion' : 'comentario';
      const d = {
        tipo, autor: clip(b.autor, 80), archivo: clip(b.archivo, 160), cto: clip(b.cto, 40), modelo: clip(b.modelo, 80),
        linea: clip(b.linea, 400), comentario: clip(b.comentario, 2000),
        cambios: (Array.isArray(b.cambios) ? b.cambios : []).slice(0, 40).map((c) => ({ antes: clip(c.antes, 400), despues: clip(c.despues, 400), codigos: (c.codigos || []).slice(0, 10).map((x) => clip(x, 30)) })),
        codigos: (Array.isArray(b.codigos) ? b.codigos : []).slice(0, 20).map((c) => ({ code: clip(c.code, 30), desc: clip(c.desc, 200) })),
        fecha: new Date().toISOString(),
      };
      if (tipo === 'comentario' && !d.comentario.trim()) return send(res, 400, { error: 'Escribe el comentario.' });
      if (tipo === 'correccion' && !d.cambios.length) return send(res, 400, { error: 'No hay cambios que reportar.' });
      const resumen = tipo === 'comentario' ? d.comentario : `${d.cambios.length} línea(s) cambiadas`;
      const title = clip(`[${tipo}] ${d.cto || ''} ${d.modelo || ''} · ${resumen.replace(/\s+/g, ' ')}`, 120);
      const issue = await gh('/issues', { method: 'POST', body: JSON.stringify({ title, body: md(d), labels: ['feedback', tipo] }) });
      return send(res, 201, { ok: true, numero: issue.number });
    }
    if (req.method === 'GET') {
      if (!isAdmin(req)) return send(res, 401, { error: 'Clave de administrador inválida.' });
      const list = await gh('/issues?labels=feedback&state=open&per_page=100&sort=created&direction=desc');
      const items = list.filter((i) => !i.pull_request).map((i) => ({ numero: i.number, url: i.html_url, titulo: i.title, creado: i.created_at, datos: unpack(i.body) }));
      return send(res, 200, { ok: true, items });
    }
    res.setHeader('Allow', 'GET, POST'); return send(res, 405, { error: 'Método no permitido' });
  } catch (e) { return send(res, e.status === 404 ? 502 : 500, { error: 'No se pudo hablar con GitHub. Revisa el token en Vercel.', detalle: e.message }); }
};
