// Uso: node build.js vault.json salida.html
// Arma la página única: src/app.src.html + motor + cifrado Office + vault cifrado + DataSpaces.
const fs=require('fs'), path=require('path');
const [vault,out]=process.argv.slice(2); const S=(f)=>fs.readFileSync(path.join(__dirname,f),'utf8');
let s=S('app.src.html'); const eng=S('engine.js'), oc=S('officecrypt.js');
if(/<\/script/i.test(eng+oc)) throw new Error('script tag in lib');
s=s.replace('/*ENGINE*/',()=>eng).replace('/*OFFICECRYPT*/',()=>oc).replace('/*VAULT*/',()=>fs.readFileSync(vault,'utf8')).replace('/*DATASPACES*/',()=>S('dataspaces.json'));
if(/\/\*(ENGINE|OFFICECRYPT|VAULT|DATASPACES)\*\//.test(s)) throw new Error('placeholder left');
for (const m of s.matchAll(/<script>([\s\S]*?)<\/script>/g)) new Function(m[1]);
fs.writeFileSync(out,s); console.log('built',out,s.length);
