/* Cifrado Office "Agile" (ECMA-376, AES-256-CBC + SHA-512, 100000 iteraciones), igual que Excel "Cifrar con contraseña".
   encryptOffice(bytes, password, {CFB, sha512, dataspaces}) -> Promise<Uint8Array> (archivo .xlsx cifrado, contenedor CFB). */
(function (root) {
  'use strict';
  const subtle = (root.crypto || globalThis.crypto).subtle;
  const rnd = (n) => { const a = new Uint8Array(n); (root.crypto || globalThis.crypto).getRandomValues(a); return a; };
  const cat = (...arrs) => { const n = arrs.reduce((s, a) => s + a.length, 0); const o = new Uint8Array(n); let k = 0; for (const a of arrs) { o.set(a, k); k += a.length; } return o; };
  const u32 = (i) => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, i, true); return b; };
  const b64e = (u) => { let s = ''; for (let i = 0; i < u.length; i++) s += String.fromCharCode(u[i]); return btoa(s); };
  const b64d = (s) => { const bin = atob(s); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
  const utf16 = (s) => { const u = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); u[2 * i] = c & 255; u[2 * i + 1] = c >> 8; } return u; };
  const BK = {
    vhi: [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79], vhv: [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e],
    key: [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6], hmacKey: [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6], hmacVal: [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33],
  };
  async function aesCbc(key, iv, data) { // sin relleno: data múltiplo de 16
    const k = await subtle.importKey('raw', key, { name: 'AES-CBC' }, false, ['encrypt']);
    const out = new Uint8Array(await subtle.encrypt({ name: 'AES-CBC', iv }, k, data));
    return out.slice(0, data.length);
  }
  const pad16 = (u) => (u.length % 16 ? cat(u, new Uint8Array(16 - (u.length % 16))) : u);

  async function encryptOffice(plain, password, deps) {
    const { CFB, sha512, dataspaces } = deps;
    const H = (u) => new Uint8Array(sha512.arrayBuffer(u));
    const keyDataSalt = rnd(16), pwSalt = rnd(16), secretKey = rnd(32), spin = 100000;
    // Llave derivada de la contraseña
    let h = H(cat(pwSalt, utf16(password)));
    const buf = new Uint8Array(68);
    for (let i = 0; i < spin; i++) { buf.set(u32(i), 0); buf.set(h, 4); h = H(buf); }
    const dk = (blk) => H(cat(h, new Uint8Array(blk))).slice(0, 32);
    const vhi = rnd(16);
    const encVHI = await aesCbc(dk(BK.vhi), pwSalt, vhi);
    const encVHV = await aesCbc(dk(BK.vhv), pwSalt, H(vhi));
    const encKey = await aesCbc(dk(BK.key), pwSalt, secretKey);
    // Paquete cifrado por segmentos de 4096
    const parts = [];
    const size = new Uint8Array(8); new DataView(size.buffer).setUint32(0, plain.length, true);
    parts.push(size);
    for (let i = 0, seg = 0; i < plain.length; i += 4096, seg++) {
      const iv = H(cat(keyDataSalt, u32(seg))).slice(0, 16);
      parts.push(await aesCbc(secretKey, iv, pad16(plain.slice(i, i + 4096))));
    }
    const pkg = cat(...parts);
    // Integridad (HMAC-SHA512)
    const hmacKey = rnd(64);
    const encHmacKey = await aesCbc(secretKey, H(cat(keyDataSalt, new Uint8Array(BK.hmacKey))).slice(0, 16), hmacKey);
    const hk = await subtle.importKey('raw', hmacKey, { name: 'HMAC', hash: 'SHA-512' }, false, ['sign']);
    const hmacVal = new Uint8Array(await subtle.sign('HMAC', hk, pkg));
    const encHmacVal = await aesCbc(secretKey, H(cat(keyDataSalt, new Uint8Array(BK.hmacVal))).slice(0, 16), hmacVal);
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">'
      + `<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${b64e(keyDataSalt)}"/>`
      + `<dataIntegrity encryptedHmacKey="${b64e(encHmacKey)}" encryptedHmacValue="${b64e(encHmacVal)}"/>`
      + '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">'
      + `<p:encryptedKey spinCount="${spin}" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${b64e(pwSalt)}" encryptedVerifierHashInput="${b64e(encVHI)}" encryptedVerifierHashValue="${b64e(encVHV)}" encryptedKeyValue="${b64e(encKey)}"/>`
      + '</keyEncryptor></keyEncryptors></encryption>';
    const info = cat(new Uint8Array([4, 0, 4, 0, 0x40, 0, 0, 0]), new TextEncoder().encode(xml));
    // Contenedor CFB
    const cfb = CFB.utils.cfb_new();
    for (const [path, data] of Object.entries(dataspaces)) CFB.utils.cfb_add(cfb, '/' + path, b64d(data));
    CFB.utils.cfb_add(cfb, '/EncryptionInfo', info);
    CFB.utils.cfb_add(cfb, '/EncryptedPackage', pkg);
    try { CFB.utils.cfb_del(cfb, '/\u0001Sh33tJ5'); } catch (_) {}
    const out = CFB.write(cfb, { type: 'array', fileType: 'cfb' });
    return out instanceof Uint8Array ? out : new Uint8Array(out);
  }
  const api = { encryptOffice };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.OfficeCrypt = api;
})(typeof window !== 'undefined' ? window : globalThis);
