// Uso: CLAVE='...' node make-vault.js plantilla.xlsx config.json salida.json [ejemplo.xlsx]
// Cifra la plantilla, la configuración (firma) y opcionalmente un ejemplo con AES-GCM (PBKDF2-SHA256, 250000 iteraciones).
const fs=require('fs'); const {webcrypto:c}=require('crypto');
const b64=(u)=>Buffer.from(u).toString('base64');
(async()=>{
  const [tpl,cfg,out,sample]=process.argv.slice(2); const PW=process.env.CLAVE;
  if(!PW||!tpl||!cfg||!out) throw new Error('Uso: CLAVE=... node make-vault.js plantilla.xlsx config.json salida.json [ejemplo.xlsx]');
  const iter=250000, salt=c.getRandomValues(new Uint8Array(16));
  const base=await c.subtle.importKey('raw',new TextEncoder().encode(PW),'PBKDF2',false,['deriveKey']);
  const key=await c.subtle.deriveKey({name:'PBKDF2',salt,iterations:iter,hash:'SHA-256'},base,{name:'AES-GCM',length:256},false,['encrypt']);
  const enc=async(buf)=>{const iv=c.getRandomValues(new Uint8Array(12)); return {iv:b64(iv),ct:b64(new Uint8Array(await c.subtle.encrypt({name:'AES-GCM',iv},key,buf)))};};
  const v={iter,salt:b64(salt),template:await enc(fs.readFileSync(tpl)),config:await enc(fs.readFileSync(cfg))};
  if(sample) v.sample=await enc(fs.readFileSync(sample));
  fs.writeFileSync(out,JSON.stringify(v)); console.log('vault ->',out);
})().catch(e=>{console.error(e.message);process.exit(1)});
