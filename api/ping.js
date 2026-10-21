const { configured, send } = require('./_lib');
module.exports = async (req, res) => send(res, 200, { ok: true, feedback: configured(), admin: Boolean(process.env.ADMIN_CLAVE) });
