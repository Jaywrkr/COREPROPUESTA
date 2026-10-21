/* Motor de propuestas Coresolutions: configurador Lenovo (archivo 1) -> propuesta (archivo 2).
   Funciona en navegador (window.CoreEngine) y en Node (module.exports). Requiere XLSX (SheetJS) para leer. */
(function (root) {
  'use strict';

  // ---------- utilidades ----------
  const COLS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const colIdx = (c) => COLS.indexOf(c);
  const r2 = (x) => Math.round((x + Number.EPSILON) * 100) / 100;
  const num = (x) => {
    if (x === null || x === undefined || x === '') return NaN;
    if (typeof x === 'number') return x;
    const s = String(x).trim().replace(/[$\s]/g, '');
    if (/^-?\d{1,3}(\.\d{3})*(,\d+)?$/.test(s) && s.includes(',')) return parseFloat(s.replace(/\./g, '').replace(',', '.'));
    return parseFloat(s.replace(/,/g, ''));
  };
  const isNum = (x) => !isNaN(num(x));
  const str = (x) => (x === null || x === undefined ? '' : String(x));
  const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const MON_EN = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
  function parseEnDate(s) { // "24-Sep-2026 00:00:00"
    const m = /(\d{1,2})-([A-Za-z]{3})-(\d{4})/.exec(str(s));
    if (!m) return null;
    return new Date(+m[3], MON_EN[m[2].toLowerCase()], +m[1]);
  }
  const fechaLarga = (d) => `${d.getDate()} de ${MESES[d.getMonth()]} de ${d.getFullYear()}`;
  const fechaNota = (d) => `${d.getDate()}-${MESES[d.getMonth()].toUpperCase()}-${d.getFullYear()}`;
  const fmtCap = (s) => s.replace(',', '.').replace(/\s+/g, '').replace(/\.0(?=[GT]B)/, '');
  const fmtGB = (gb) => (gb >= 1024 && gb % 1024 === 0 ? `${gb / 1024}TB` : `${gb}GB`);
  const NUMW = { 1: 'Un', 2: 'Dos', 3: 'Tres', 4: 'Cuatro' };

  // ---------- lectura del archivo ----------
  function gridOf(ws, XLSX) {
    const out = [];
    if (!ws || !ws['!ref']) return out;
    const rg = XLSX.utils.decode_range(ws['!ref']);
    for (let r = 0; r <= rg.e.r; r++) {
      const row = [];
      for (let c = 0; c <= Math.max(rg.e.c, 11); c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        row.push(cell ? { v: cell.v, f: cell.f || null, t: cell.t } : null);
      }
      out.push(row);
    }
    return out;
  }
  const gv = (g, r, c) => (g[r] && g[r][c] ? g[r][c].v : null);

  function parseWorkbook(buf, XLSX) {
    const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: false, cellStyles: false });
    // Bid Platform: alguna hoja con 'P/N' en A y 'Subtotal' en B
    for (const name of wb.SheetNames) {
      const g = gridOf(wb.Sheets[name], XLSX);
      const hr = g.findIndex((row) => row && row[0] && str(row[0].v).trim() === 'P/N');
      if (hr >= 0 && g.some((row) => row && row[1] && str(row[1].v).trim() === 'Subtotal')) return parseBid(g, hr, name);
    }
    for (const name of wb.SheetNames) {
      const g = gridOf(wb.Sheets[name], XLSX);
      const hr = g.findIndex((row) => row && row.some((c) => c && /N[úu]mero de pieza|Part Number/i.test(str(c.v))));
      if (hr >= 0) return parseDCSC(g, hr, name);
    }
    throw new Error('No reconozco el formato. Sube un configurador Lenovo: exportación de Bid Platform o cotización DCSC.');
  }

  function parseBid(g, hr, sheetName) {
    const meta = { format: 'bid', sheetName };
    const find = (re) => { for (let r = 0; r < hr; r++) for (const c of g[r] || []) if (c && re.test(str(c.v))) return str(c.v); return ''; };
    meta.title = str(gv(g, 0, 1));
    meta.preparedBy = find(/^Preparado Por:/i).replace(/^Preparado Por:\s*/i, '');
    meta.preparedFor = find(/^Preparado Para:/i).replace(/^Preparado Para:\s*/i, '');
    meta.endCustomer = find(/^Para el cliente final:/i).replace(/^Para el cliente final:\s*/i, '');
    meta.expires = parseEnDate(find(/^Fecha de expiraci/i));
    meta.created = parseEnDate(find(/^Fecha de creaci/i));
    meta.bidId = find(/^Bid Request ID:/i).replace(/^Bid Request ID:\s*/i, '');
    const configs = [];
    let cfg = null, owner = null;
    for (let r = hr + 1; r < g.length; r++) {
      const A = str(gv(g, r, 0)).trim(), B = str(gv(g, r, 1)).trim();
      const D = gv(g, r, 3), F = gv(g, r, 5);
      if (str(gv(g, r, 3)).trim() === 'Total:' || str(gv(g, r, 5)).trim() === 'Total:') break;
      if (/^Discount Codes/i.test(A)) break;
      if (/^Set from/i.test(A)) continue;
      if (B === 'Subtotal') {
        cfg = { subRow: r + 1, cto: A, listSub: num(D), finalSub: num(F), items: [], features: [] };
        configs.push(cfg); owner = null; continue;
      }
      if (!cfg) continue;
      if (/^C.digo De Caracter/i.test(A)) continue;
      if (!A && !B) continue;
      if ((isNum(D) || isNum(F)) && A) {
        if (cfg.ctoRow === undefined && A === cfg.cto) {
          Object.assign(cfg, { ctoRow: r + 1, name: B, qty: num(gv(g, r, 2)) || 1, listUnit: num(D), finalUnit: num(F), esd: str(gv(g, r, 7)) });
          owner = cfg.features;
        } else {
          const it = { pn: A, desc: B.replace(/^software\d+\s+/i, ''), qty: num(gv(g, r, 2)) || 1, listUnit: num(D), finalUnit: num(F), features: [] };
          cfg.items.push(it); owner = it.features;
        }
        continue;
      }
      if (owner && A) owner.push({ code: A, desc: B, qty: num(gv(g, r, 2)) || 1 });
    }
    // cantidades de items por unidad
    for (const c of configs) {
      if (isNaN(c.finalSub)) c.finalSub = r2((c.finalUnit || 0) + c.items.reduce((s, i) => s + (i.finalUnit || 0), 0));
      if (isNaN(c.listSub)) c.listSub = r2((c.listUnit || 0) + c.items.reduce((s, i) => s + (i.listUnit || 0), 0));
    }
    return { meta, configs, grid: g, hr };
  }

  function parseDCSC(g, hr, sheetName) {
    const meta = { format: 'dcsc', sheetName };
    const H = g[hr].map((c) => str(c && c.v));
    const ci = (re) => H.findIndex((h) => re.test(h));
    const cPN = ci(/N[úu]mero de pieza|Part Number/i), cDesc = ci(/Descripci|Description/i), cQty = ci(/^Cant|^Qty/i), cPrice = ci(/^Precio$|^Price$/i);
    for (let r = 0; r < hr; r++) (g[r] || []).forEach((c, k) => {
      const t = str(c && c.v);
      if (/Fecha del precio/i.test(t)) meta.priceDate = str(gv(g, r, k + 1));
    });
    meta.title = '';
    const configs = [];
    let cfg = null, owner = null;
    for (let r = hr + 1; r < g.length; r++) {
      const A = str(gv(g, r, cPN)).trim(), B = str(gv(g, r, cDesc)).trim();
      if (/^Total/i.test(A) || /^Total$/i.test(B)) break;
      if (!A || !/^[A-Z0-9][A-Z0-9-]{2,19}$/i.test(A)) continue;
      const qty = num(gv(g, r, cQty));
      const price = num(gv(g, r, cPrice));
      if (/^7D[A-Z0-9]{2}CTO\d?WW$/i.test(A) && B.includes(':')) {
        const name = B.split(':')[0].trim();
        cfg = { cto: A, name: `${name} ${B.split(':').slice(1).join(':').trim()}`, cfgName: name, qty: qty || 1, listUnit: price, finalUnit: NaN, items: [], features: [] };
        configs.push(cfg); owner = cfg.features; continue;
      }
      if (!cfg) continue;
      if (!isNaN(price)) {
        const it = { pn: A, desc: B, qty: qty || 1, listUnit: price, finalUnit: NaN, features: [] };
        cfg.items.push(it); owner = it.features; continue;
      }
      if (owner) owner.push({ code: A, desc: B, qty: qty || 1 });
    }
    for (const c of configs) {
      const q = c.qty || 1;
      for (const f of c.features) f.qty = f.qty / q;
      for (const it of c.items) { it.qty = it.qty; for (const f of it.features) f.qty = f.qty / q; }
      c.listSub = r2((c.listUnit || 0) + c.items.reduce((s, i) => s + (i.listUnit || 0), 0));
      c.finalSub = NaN; // DCSC no trae precio final (bid)
    }
    if (!configs.length) throw new Error('No encontré ninguna configuración de servidor (CTO) en el archivo DCSC.');
    return { meta, configs, grid: g, hr };
  }

  // ---------- descripción técnica en español ----------
  function modelOf(cfg) {
    const s = `${cfg.name} ${cfg.cfgName || ''}`;
    const m = /(ThinkAgile\s+)?(VX\d{3}\s?V\d|SR\d{3}a?\s?V\d|SR\d{3}|SD\d{3}\s?V\d|ST\d{3}\s?V\d|D4390|HX\d{3,4}\s?V\d)/i.exec(s);
    const model = m ? m[2].toUpperCase().replace(/(\d)\s?V/, '$1 V') : cfg.cto;
    const family = /^VX|^HX/.test(model) ? 'ThinkAgile' : 'ThinkSystem';
    let u = 1;
    if (/^(SR|VX|HX)\d5/.test(model)) u = 2;
    if (/^SR8[56]/.test(model)) u = /SR86/.test(model) ? 4 : 2;
    if (/D4390/.test(model)) u = 4;
    if (/^ST/.test(model)) u = 0;
    return { model, family, u };
  }

  function wrapLine(text, max = 96) {
    if (text.length <= max) return [text];
    const out = []; let rest = text;
    while (rest.length > max) {
      let cut = rest.lastIndexOf(', ', max); if (cut < max * 0.5) cut = rest.lastIndexOf(' ', max);
      if (cut <= 0) cut = max;
      out.push(rest.slice(0, cut + (rest[cut] === ',' ? 1 : 0)).trimEnd());
      rest = '  ' + rest.slice(cut + 1).trimStart();
    }
    out.push(rest);
    return out;
  }

  // Partes que nunca van en la descripción (chasis, cables, rellenos, etc.)
  const IGNORE_RE = /Chassis|Chasis|Heat ?sink|Disipador|\bFan|Ventilador|Cable|Cord|Cbl|Riser|elevadora|Backplane|Placa base|Filler|Relleno|Dummy|Rail|rieles|TPM|Media Bay|Processor Board|Tarjeta de procesador|I\/O board|Flag|Enablement|habilitaci|Environment|Entorno|Power Efficiency|Select Storage|Dispositivos de almacenamiento|Cage|Jaula|Tray|Bandeja|Interposer|SuperCap|Clip|Laser|indicador|MI for PXE|Feature Enable|habilitado en|Upgrade|Actualizaci|Configuration Instruction|Bezel|Bisel|Latch|Pestillo|IPMI|Slot|Ranura|Label|Etiqueta|Packag|Embalaje|Manual|Shipping|Language|Idioma|ESXi|Air Baffle|Deflector|Bracket|Screw|Tornillo|\bBay\b|Operating Mode|No OS|Sin sistema|Board|Compartimiento|Bah[ií]a|Baja tensi|Especificar|Precarga|LPK|Redundancia|Panel ciego|Placa de E\/S|Soporte|Paquete de configuraci|EXPANSION|Conducto|Air Duct|MI para PXE|Lista de materiales|\bMI-|Configuration ID|^Controller \d+$|Bandera|\bBP\b|^M\.2 NVMe$|^ThinkAgile [A-Z]{2}\d+ V\d$|^ThinkSystem [A-Z]{2}\d+ V\d$|Low voltage|Specify|Preload|Blank|Air duct|BOM/i;
  // Lo que las reglas incorporadas sí convierten en texto
  const KNOWN_RE = /Xeon|EPYC|TruDDR|RDIMM|DIMM|Ethernet|Fibre Channel|canal de fibra|\bSSD\b|\bHDD\b|Unidad|RAID\s*\d{3,4}-\d+i|NVIDIA|Power Supply|Fuente de alimentaci|XClarity|XCC|Premier|Foundation|Essential|CSR|KYD|Keep Your Drive|Base Warranty/i;

  function applyCustom(cfg, rules) {
    rules = rules || {};
    const hit = [];
    const keep = (f) => { const r = rules[String(f.code || f.pn || '').toUpperCase()]; if (r) { hit.push({ f, r }); return false; } return true; };
    const features = cfg.features.filter(keep);
    const items = cfg.items.filter(keep).map((it) => Object.assign({}, it, { features: it.features.filter(keep) }));
    return { cfg: Object.assign({}, cfg, { features, items }), hit };
  }

  function unknownParts(cfg, rules) {
    const { cfg: c } = applyCustom(cfg, rules);
    const out = [];
    const chk = (code, desc) => { desc = String(desc || '').trim(); if (!code || /^Total|^C.digo/i.test(code)) return; if (IGNORE_RE.test(desc) && !/SSD|HDD|NVIDIA|Ethernet|Fibre/i.test(desc)) return; if (KNOWN_RE.test(desc)) return; out.push({ code, desc }); };
    c.features.forEach((f) => chk(f.code, f.desc));
    c.items.forEach((it) => { chk(it.pn, it.desc); it.features.forEach((f) => chk(f.code, f.desc)); });
    const seen = new Set(); return out.filter((x) => !seen.has(x.code) && seen.add(x.code));
  }

  function describe(cfg0, rules) {
    const { cfg, hit } = applyCustom(cfg0, rules);
    const F = cfg.features, items = cfg.items;
    const all = F.concat(...items.map((i) => i.features));
    const has = (re) => all.filter((f) => re.test(f.desc));
    const warn = [];
    const { model, family, u } = modelOf(cfg);
    const lines = [], codes = [];
    let cur = [];
    const C = (arr) => (arr || []).filter(Boolean).map((f) => String(f.code || f.pn || f));
    const L = (t) => { lines.push(t); codes.push(cur.slice()); };
    const isMT = /\bMT\b|MEMORY\s*TIERING|TIERING/i.test(`${cfg.name} ${cfg.cfgName || ''}`);
    const isVX = /^VX|^HX/.test(model);
    const isStorage = /D4390/.test(model);
    cur = [cfg.cto];
    L(`Lenovo ${family} ${model}${u ? `, para rack (${u}U)` : ''},`);

    // CPU
    let cpus = 0;
    const cpu = has(/Xeon|EPYC/i).filter((f) => !/Heatsink|Disipador|Dummy|Clip/i.test(f.desc))[0];
    if (cpu) {
      cpus = cpu.qty; cur = C([cpu]);
      const m = /(Xeon|EPYC)\s+(\S+)\s+(\d+)C\b.*?([\d.,]+)\s*GHz/i.exec(cpu.desc);
      if (m) {
        const brand = m[1].toUpperCase() === 'EPYC' ? 'AMD EPYC' : 'Intel Xeon';
        const ghz = m[4].replace(',', '.');
        if (cpus === 1) L(`Un procesador ${brand} ${m[2]} (${m[3]} cores/${ghz}Ghz), expandible a 2 procesadores,`);
        else L(`${NUMW[cpus] || cpus} procesadores ${brand} ${m[2]} de ${m[3]} cores/${ghz}Ghz (total ${cpus * m[3]} cores),`);
      } else { L(`${cpus} x ${cpu.desc},`); warn.push(`Procesador no reconocido: "${cpu.desc}"`); }
    } else if (!isStorage) warn.push('No encontré el procesador en la configuración.');

    // Discos de datos
    const drives = has(/SSD|HDD|Unidad|Drive/i).filter((f) => /(2[.,]5|3[.,]5|U\.2|U\.3)/.test(f.desc) && !/M\.2|Filler|Dummy|Backplane|Placa|Relleno|Cage|Bay|Bah|Cable|Kit/i.test(f.desc));
    let mtDrives = [];
    // Memoria
    const mem = has(/TruDDR|RDIMM|DIMM/i).filter((f) => !/Filler|Relleno/i.test(f.desc));
    if (mem.length) {
      let n = 0, total = 0, size = 0, speed = '';
      for (const d of mem) {
        const m = /(\d+)\s*GB\s*TruDDR\d\s*(\d+)\s*MHz/i.exec(d.desc) || /(\d+)\s*GB/i.exec(d.desc);
        if (m) { size = +m[1]; speed = m[2] || speed; n += d.qty; total += d.qty * size; }
      }
      const slots = (/^SR6|^VX6|^HX6/.test(model) ? 16 : 12) * Math.max(cpus, 1);
      const maxSlots = (/^SR6|^VX6|^HX6/.test(model) ? 16 : 12) * 2;
      cur = C(mem);
      if (isMT) {
        mtDrives = drives.filter((d) => /Mixed Use|uso mixto/i.test(d.desc) && /NVMe/i.test(d.desc)); cur = C(mem.concat(mtDrives));
        const k = mtDrives.reduce((s, d) => s + d.qty, 0);
        const cap = mtDrives[0] ? fmtCap((/(\d+(?:[.,]\d+)?)\s*(TB|GB)\b/.exec(mtDrives[0].desc) || ['', '?', 'TB']).slice(1).join('')) : '';
        L(`${fmtGB(total * 2)} con Memory Tiering (${n} x ${size}GB TruDDR5${speed ? ` ${speed}Mhz` : ''} + ${k} x NVMe de ${cap}${k > 1 ? ' en RAID-1' : ''}),`);
        if (!k) warn.push('Configuración con Memory Tiering sin unidad NVMe Mixed Use detectada; revisa la línea de memoria.');
      } else {
        const free = slots - n;
        L(`${total}GB de memoria RAM (${n} x ${size}GB TruDDR5${speed ? ` ${speed}Mhz` : ''}), ${free} zócalos libres${cpus === 1 ? `, con el segundo procesador se habilitan ${maxSlots - slots} zócalos adicionales` : ''},`);
      }
    } else if (!isStorage) warn.push('No encontré memoria RAM en la configuración.');

    // Red
    const nics = has(/Ethernet|Network Adapter|Adaptador de red/i).filter((f) => /port|puerto/i.test(f.desc) && !/Riser|Cable|Filler|Relleno|PXE|Cage/i.test(f.desc));
    const byType = {}, byTypeC = {};
    for (const nc of nics) {
      const pm = /(\d+)\s*-?\s*port|(\d+)\s*puertos/i.exec(nc.desc);
      const ports = pm ? +(pm[1] || pm[2]) : 2;
      let label = 'Ethernet';
      if (/10\/25\s*GbE|10\/25 ?Gb/i.test(nc.desc)) label = 'Ethernet 10/25Gb SFP28';
      else if (/100\s*GbE|100\/200/i.test(nc.desc)) label = 'Ethernet 100Gb QSFP';
      else if (/10\s*GBASE-T|10\s*GbE.*RJ-?45/i.test(nc.desc)) label = 'Ethernet 10Gb RJ-45';
      else if (/10\s*GbE.*SFP\+/i.test(nc.desc)) label = 'Ethernet 10Gb SFP+';
      else if (/\b1\s*GbE/i.test(nc.desc)) label = 'Ethernet 1Gb RJ-45';
      byType[label] = (byType[label] || 0) + ports * nc.qty; (byTypeC[label] = byTypeC[label] || []).push(nc);
    }
    for (const [label, n] of Object.entries(byType)) (cur = C(byTypeC[label])), L(`${n} puertos ${label},`);
    const fcs = has(/Fibre Channel|canal de fibra/i).filter((f) => /port|puerto/i.test(f.desc));
    if (fcs.length) {
      cur = C(fcs);
      let n = 0, sp = '';
      for (const f of fcs) { const pm = /(\d+)\s*-?\s*port|(\d+)\s*puertos/i.exec(f.desc); n += (pm ? +(pm[1] || pm[2]) : 2) * f.qty; const s = /(\d+)\s*Gb/i.exec(f.desc); if (s) sp = s[1]; }
      L(`${n} puertos ${sp === '32' ? '16/32' : sp}Gb Fibre Channel,`);
    }

    // Arranque M.2
    const m2 = has(/M\.2/i).filter((f) => /SSD/i.test(f.desc) && !/HeatSink|Disipador|Cage|Tray|Interposer|Kit|Cable|Enablement|Habilitaci|Bandeja|Jaula/i.test(f.desc));
    if (m2.length) {
      cur = C(m2);
      const n = m2.reduce((s, d) => s + d.qty, 0);
      const cap = fmtCap((/(\d+(?:[.,]\d+)?)\s*(TB|GB)\b/.exec(m2[0].desc) || ['', '480', 'GB']).slice(1).join(''));
      const iface = /NVMe/i.test(m2[0].desc) ? 'NVMe ' : '';
      L(n > 1 ? `${n} unidades SSD ${iface}M.2 de ${cap} (en RAID-1) para ${isVX || /VMW/i.test(cfg.name) ? 'hipervisor' : 'sistema operativo o hipervisor'},`
        : `1 unidad SSD ${iface}M.2 de ${cap} para ${isVX || /VMW/i.test(cfg.name) ? 'hipervisor' : 'sistema operativo o hipervisor'},`);
    }

    // RAID
    const raid = has(/RAID\s*\d{3,4}-\d+i/i).filter((f) => !/SuperCap|Cable/i.test(f.desc))[0];
    if (raid) {
      cur = C([raid]);
      const c = /(\d+)\s*GB\s*Flash/i.exec(raid.desc);
      L(c ? `Controlador de discos con ${c[1]}GB de caché, maneja RAID-0/1/5/6/10,` : 'Controlador de discos RAID-0/1/10/5,');
    }

    // Discos de datos (excluye los de Memory Tiering)
    const dataDrives = drives.filter((d) => !mtDrives.includes(d));
    let nData = 0;
    for (const d of dataDrives) {
      const cap = fmtCap((/(\d+(?:[.,]\d+)?)\s*(TB|GB)\b/.exec(d.desc) || ['', '?', '']).slice(1).join(''));
      const iface = /NVMe/i.test(d.desc) ? 'NVMe' : /SAS/i.test(d.desc) ? 'SAS' : /SATA/i.test(d.desc) ? 'SATA' : '';
      const use = /Mixed Use|uso mixto/i.test(d.desc) ? ' tipo Mixed Use' : /Read Intensive|lectura intensiva/i.test(d.desc) ? ' tipo Read Intensive' : '';
      const kind = /HDD|rpm/i.test(d.desc) ? 'discos' : 'unidades SSD';
      nData += d.qty; cur = C([d]);
      L(`${d.qty} ${d.qty === 1 ? (kind === 'discos' ? 'disco' : 'unidad SSD') : kind} ${iface}${use} de ${cap}${d.qty > 1 ? ' c/u' : ''}${isVX ? ', para almacenamiento VSAN ESA' : ''},`.replace(/\s+/g, ' '));
    }
    const ch = has(/(\d+)x2[.,]5"/i)[0];
    if (ch && nData) {
      cur = C([ch]);
      const bays = +/(\d+)x2[.,]5"/i.exec(ch.desc)[1];
      const free = bays - nData - mtDrives.reduce((s, d) => s + d.qty, 0);
      if (free > 0) L(`${free} bahías libres para crecimiento en discos,`);
    }

    // GPU
    const gpus = has(/NVIDIA/i).filter((f) => !/Cable|Kit|Riser|Power|Bracket/i.test(f.desc));
    for (const gp of gpus) {
      cur = C([gp]);
      const m = /NVIDIA\s+([A-Z]*\d+[A-Za-z]*(?:\s?NVL)?)\s.*?(\d+)\s*GB/i.exec(gp.desc);
      L(m ? `${gp.qty} GPU NVIDIA ${m[1]} ${m[2]}GB PCIe,` : `${gp.qty} x ${gp.desc},`);
    }

    // Energía / ventilación
    const psu = has(/Power Supply|Fuente de alimentaci/i)[0];
    cur = C([psu]);
    if (psu) L(psu.qty >= 2 ? `${NUMW[psu.qty] || psu.qty} fuentes de poder redundantes y hot-swap, ventiladores redundantes y hot-swap,` : '1 fuente de poder, ventiladores redundantes y hot-swap,');

    // Gestión
    cur = C(items.filter((i) => /XClarity|XCC/i.test(i.desc)));
    const allItems = items.map((i) => i.desc).join(' | ') + ' | ' + all.map((f) => f.desc).join(' | ');
    if (/XClarity Controller Prem|XClarity Platinum|XCC\d? premier/i.test(allItems)) L('Puerto de gestión remota con XClarity Premier,');
    else L('Puerto de gestión remota con XClarity Controller,');
    if (/XClarity One/i.test(allItems)) (cur = C(items.filter((i) => /XClarity One/i.test(i.desc)))), L('Gestión centralizada con XClarity One,');

    // Garantía
    const wItems = items.map((i) => i.desc);
    const prem = wItems.map((d) => /(\d)\s*Yr\s+(Premier|Foundation|Essential)\s+(24x7|9x5)\s+(\d+)\s*Hr\s+Resp/i.exec(d)).find(Boolean);
    const nbd = wItems.map((d) => /(\d)\s*Yr\s+(Premier|Foundation|Essential)\s+(?:9x5\s+)?NBD/i.exec(d)).find(Boolean);
    const csr = wItems.map((d) => /(\d+)\s*Hr\s+CSR/i.exec(d)).find(Boolean);
    const kyd = wItems.some((d) => /KYD|Keep Your Drive/i.test(d));
    let w;
    if (prem) {
      w = `** Garantía de hardware: ${prem[1]} años en modalidad ${prem[3] === '24x7' ? '7x24' : '5x9'} con tiempo de ${csr ? `reparación de ${csr[1]}` : `respuesta de ${prem[4]}`} horas`;
    } else if (nbd) {
      w = `** Garantía de hardware: ${nbd[1]} años en modalidad 5x9 con tiempo de respuesta al siguiente día laborable`;
    } else {
      const y = /(\d)\s*yr\s+Base Warranty/i.exec(cfg.name);
      w = `** Garantía de hardware: ${y ? y[1] : 3} años en modalidad 5x9 con tiempo de respuesta de 24 horas laborables`;
    }
    if (kyd) w += ' y Keep Your Drive (KYD)';
    cur = C(items.filter((i) => /Premier|Foundation|Essential|CSR|KYD|Keep Your Drive/i.test(i.desc)));
    if (!cur.length) cur = [cfg.cto];
    L(w);

    // Reglas del diccionario: van antes de energía / gestión / garantía
    const custom = [], customC = [];
    const done = new Set();
    for (const { f, r } of hit) {
      const code = String(f.code || f.pn).toUpperCase();
      if (done.has(code) || r.accion === 'ignorar' || !String(r.texto || '').trim()) { done.add(code); continue; }
      done.add(code);
      const q = f.qty || 1;
      custom.push(String(r.texto).replace(/\{cant\}/gi, String(+q.toFixed(2)))); customC.push([code]);
    }
    if (custom.length) {
      let at = lines.findIndex((l) => /fuentes? de poder|Puerto de gestión|Garantía/i.test(l));
      if (at < 0) at = lines.length;
      lines.splice(at, 0, ...custom); codes.splice(at, 0, ...customC);
    }
    const out = [], outC = [];
    lines.forEach((l, i) => { const w = wrapLine(l); out.push(...w); w.forEach(() => outC.push(codes[i] || [])); });
    return { lines: out, codes: outC, warnings: warn, model, isVX, hasData: nData > 0, hasGPU: gpus.length > 0, isMT };
  }

  // ---------- modelo de propuesta por defecto ----------
  function sectionTitle(cfg, d) {
    if (d.isVX) return 'Nodos para VMware - almacenamiento VSAN';
    if (d.hasGPU || /\bIA\b|GPU/i.test(cfg.name)) return 'Servidor para Inteligencia Artificial';
    if (/D4390/.test(d.model)) return 'Gabinete de discos';
    if (/VMW|VMWARE|VCF|VSPHERE/i.test(cfg.name)) return 'Servidores para clúster de virtualización';
    return cfg.qty > 1 ? 'Servidores' : 'Servidor';
  }

  function buildDefault(parsed, today, fileName, rules, opts) {
    opts = opts || {};
    const F0 = opts.firma || {};
    const d0 = today || new Date();
    const { meta, configs } = parsed;
    const warnings = [];
    if (meta.format === 'dcsc') warnings.push({ level: 'alto', text: 'El archivo es una cotización DCSC: trae solo precio de lista, sin precio final de Bid. Ingresa el costo final por unidad en la hoja _Lenovo (celda amarilla) antes de enviar.' });
    const blocks = [];
    configs.forEach((cfg, i) => {
      const d = describe(cfg, rules);
      d.warnings.forEach((t) => warnings.push({ level: 'medio', text: `${cfg.cto}: ${t}` }));
      blocks.push({
        kind: 'item', id: 'b' + i + '_' + Math.random().toString(36).slice(2, 6), cfg: i,
        title: sectionTitle(cfg, d), qty: cfg.qty || 1, lines: d.lines, gen: d.lines.slice(), genCodes: d.codes,
        cost: [
          { kind: 'ref', factor: 25 },
          { kind: 'free', sku: 'SERVICIO', detalle: '', cant: 8, costo: 48, factor: 40 },
        ],
        _hasData: d.hasData,
      });
    });
    const hasData = blocks.some((b) => b._hasData);
    if (meta.expires) blocks.push({ kind: 'note', id: 'n1', lines: [`NOTA: ${configs.length > 1 ? 'LOS PRECIOS SON VÁLIDOS' : 'EL PRECIO ES VÁLIDO'} HASTA EL DÍA ${fechaNota(meta.expires)}.`] });
    blocks.push({
      kind: 'text', id: 't1', title: 'Servicios de implementación incluidos:',
      lines: [
        '- Instalación física del equipo en rack del cliente',
        '- Configuración de puerto de gestión remota',
        '- Actualización de firmware/microcódigos según últimas versiones liberadas al momento de instalar',
        ...(hasData ? ['- Creación de arreglos con discos internos'] : []),
        '- Soporte para iniciar la instalación del sistema operativo o hipervisor',
      ],
    });
    blocks.push({ kind: 'text', id: 't2', title: '', lines: ['** El cliente deberá contar con las licencias de sistema operativo y aplicaciones que requiera', '   ejecutar sobre el equipo propuesto.'] });
    const fromFile = fileName ? (/^(?:[0-9a-f]{8}-)?([^_]+)_/i.exec(fileName) || [])[1] : '';
    const prettyFile = fromFile ? fromFile.replace(/([a-z])([A-Z])/g, '$1 $2').toUpperCase() : '';
    const bidCli = (meta.endCustomer || meta.preparedFor || '').toUpperCase();
    const cliente = prettyFile || bidCli;
    if (prettyFile && bidCli && !bidCli.includes(prettyFile) && !prettyFile.includes(bidCli)) warnings.push({ level: 'medio', text: `El Bid de Lenovo está registrado para "${bidCli}", pero el archivo es de "${prettyFile}". Confirma el cliente antes de enviar.` });
    return {
      header: {
        cliente: cliente || 'NOMBRE DEL CLIENTE',
        titulo: configs.length > 1 || (configs[0] && configs[0].qty > 1) ? 'Servidores Lenovo' : 'Servidor Lenovo',
        atencion: 'Atención: Ing. ',
        fecha: `Cuenca, ${fechaLarga(d0)}`,
        cotizacion: `Cotización: HG-${String(d0.getFullYear()).slice(2)}___-1-1`,
      },
      blocks,
      notaIVA: 'Nota: El precio no incluye el IVA.',
      condiciones: [
        'Forma de pago: 100% contra entrega de los bienes y/o servicios ofrecidos.',
        'Plazo de entrega: Hasta 120 días contados desde la fecha de la orden de compra.',
        'Garantía técnica: CORESOLUTIONS transfiere la misma garantía técnica que el fabricante otorga a sus productos y servicios.',
        'Validez de la propuesta: según lo que se indica en cada caso, y/o mientras se mantengan las actuales condiciones comerciales.',
      ],
      firma: { nombre: F0.nombre || '', cel: F0.cel || '', email: F0.email || meta.preparedBy || '' },
      lenovoTitle: meta.title || (configs[0] ? (configs[0].cfgName || configs[0].name.split(' ThinkSystem')[0]) : ''),
      costOverride: {},
      warnings,
    };
  }

  // ---------- hoja _Lenovo ----------
  // Devuelve {rows:[{r,ht,custom,cells:{A:{v,f,s,t}}}], merges, refs:[{ctoRow,subRow,cto,name,finalSub}]}
  function buildLenovo(parsed, model) {
    const { meta, configs } = parsed;
    const rows = new Map();
    const put = (r, c, cell) => { if (!rows.has(r)) rows.set(r, { r, cells: {} }); rows.get(r).cells[c] = cell; };
    const setRow = (r, o) => { if (!rows.has(r)) rows.set(r, { r, cells: {} }); Object.assign(rows.get(r), o); };
    const refs = [];
    const S = { cto: 52, title: 73, titleSide: 44, hdrB: 74, plain: 45, notice: 46, colHdr: 47, setRow: 48, subA: 49, subVal: 50, subFinal: 51, item: 53, esd: 54, featHdr: 55, feat: 56, disc: 57 };
    const header = (title, by, forr, end, exp, created, bid, notice) => {
      put(1, 'A', { v: '', s: S.titleSide }); put(1, 'B', { v: title, s: S.title, edit: 'lenovoTitle' });
      'CDEFG'.split('').forEach((c) => put(1, c, { v: '', s: S.title })); 'HIJKL'.split('').forEach((c) => put(1, c, { v: '', s: S.titleSide }));
      setRow(1, { ht: 27.75 });
      put(2, 'A', { v: '', s: S.plain }); put(2, 'B', { v: `Preparado Por: ${by}`, s: S.hdrB }); 'CDEF'.split('').forEach((c) => put(2, c, { v: '', s: S.hdrB })); put(2, 'G', { v: `Preparado Para: ${forr}`, s: S.plain });
      put(3, 'A', { v: '', s: S.plain }); put(3, 'B', { v: `Para el cliente final: ${end}`, s: S.plain }); put(3, 'D', { v: '', s: S.hdrB }); put(3, 'E', { v: '', s: S.hdrB }); put(3, 'G', { v: `Fecha de expiración: ${exp}`, s: S.plain });
      put(4, 'A', { v: '', s: S.plain }); put(4, 'B', { v: `Fecha de creación: ${created}`, s: S.plain }); put(4, 'G', { v: `Bid Request ID: ${bid}`, s: S.plain }); put(4, 'I', { v: 'Versión#: 1', s: S.plain });
      put(5, 'A', { v: '', s: S.notice }); put(5, 'B', { v: notice || 'Tu configuración final puede contener hardware, software y servicios. Por lo tanto, las repercusiones de control de exportaciones podrían variar según la configuración final.', s: S.notice });
      setRow(5, { ht: 76.5, custom: true });
      ['P/N', 'Descripción', 'Cantidad Solicitada', 'Precio estimado del revendedor (USD)', 'Precio estimado del revendedor (USD)', 'Precio Final (USD)', 'Precio final (USD)', 'ESD'].forEach((t, k) => put(6, COLS[k], { v: t, s: S.colHdr }));
      setRow(6, { ht: 39.95, custom: true });
      ['', "'", "'", '(Unidad)', '(Unidad * Cantidad)', '(Unidad)', '(Unidad * Cantidad)', ''].forEach((t, k) => put(7, COLS[k], { v: t, s: S.plain }));
    };

    if (meta.format === 'bid') {
      const g = parsed.grid;
      const cell0 = (r, c) => g[r - 1] && g[r - 1][colIdx(c)];
      header(model.lenovoTitle, meta.preparedBy, meta.preparedFor, meta.endCustomer,
        str(gv(g, 2, 6)).replace(/^Fecha de expiraci.n:\s*/i, ''), str(gv(g, 3, 1)).replace(/^Fecha de creaci.n:\s*/i, ''), meta.bidId, str(gv(g, 4, 1)));
      // Copia literal de filas de datos (desde la fila siguiente a la de unidades) con estilos de plantilla
      const subRows = new Set(configs.map((c) => c.subRow)), ctoRows = new Set(configs.map((c) => c.ctoRow));
      let inFeat = false;
      for (let r = parsed.hr + 3; r <= g.length; r++) {
        const row = g[r - 1]; if (!row) continue;
        const A = str(row[0] && row[0].v).trim();
        const isTotal = row.some((c) => c && str(c.v).trim() === 'Total:');
        let role;
        if (/^Set from/i.test(A)) role = 'set';
        else if (subRows.has(r)) role = 'sub';
        else if (/^C.digo De Caracter/i.test(A)) { role = 'fhdr'; inFeat = true; }
        else if (isTotal) role = 'total';
        else if (/^Discount Codes/i.test(A)) role = 'disc';
        else if (row.slice(3, 7).some((c) => c && (isNum(c.v) || c.f))) { role = 'item'; inFeat = false; }
        else if (A) role = inFeat ? 'feat' : 'feat';
        else continue;
        for (let k = 0; k < 12; k++) {
          const c = row[k]; const L = COLS[k];
          if (!c && role !== 'item' && role !== 'sub') continue;
          let s = S.plain;
          if (role === 'set') s = S.setRow; else if (role === 'sub') s = k <= 2 ? S.subA : k === 5 ? S.subFinal : S.subVal;
          else if (role === 'item' || role === 'total') s = k === 7 ? S.esd : (role === 'item' && ctoRows.has(r) && k <= 1 ? S.cto : S.item);
          else if (role === 'fhdr') s = S.featHdr; else if (role === 'feat') s = S.feat; else if (role === 'disc') s = S.disc;
          if (role === 'sub' && k > 6) continue;
          if (role === 'item' && k > 8) continue;
          if (!c) { put(r, L, { v: '', s, blank: true }); continue; }
          const cell = { v: c.v === undefined ? '' : c.v, s };
          if (c.f) cell.f = c.f.replace(/^=/, '');
          put(r, L, cell);
        }
        if (role === 'set') setRow(r, { ht: 28.5, custom: true });
      }
      configs.forEach((c, i) => {
        const ov = model.costOverride[i];
        const subF = rows.get(c.subRow) && rows.get(c.subRow).cells.F;
        if (subF) { subF.edit = 'cost:' + i; if (ov !== undefined) { subF.v = ov; delete subF.f; } }
        refs.push({ ctoRow: c.ctoRow, subRow: c.subRow, cto: c.cto, name: c.name, finalSub: ov !== undefined ? ov : (isNaN(c.finalSub) ? 0 : c.finalSub) });
      });
    } else {
      // DCSC -> sintetiza el layout Bid
      header(model.lenovoTitle, model.firma.email, model.header.cliente, model.header.cliente, '', meta.priceDate || '', '(sin Bid - cotización DCSC)');
      let r = 11;
      configs.forEach((c, i) => {
        put(r, 'A', { v: `Set from Configurator(Config ${i + 1})`, s: S.setRow }); setRow(r, { ht: 28.5, custom: true }); r++;
        const sub = r, last0 = r + 1 + 1 + c.features.length + c.items.reduce((s, it) => s + 1 + (it.features.length ? 1 + it.features.length : 0), 0);
        const ov = model.costOverride[i];
        put(r, 'A', { v: c.cto, s: S.subA }); put(r, 'B', { v: 'Subtotal', s: S.subA }); put(r, 'C', { v: '', s: S.subA });
        put(r, 'D', { v: c.listSub, s: S.subVal }); put(r, 'E', { f: `SUM(E${r + 1}:E${last0})`, v: r2(c.listSub * c.qty), s: S.subVal });
        put(r, 'F', { v: ov !== undefined ? ov : '', s: S.subFinal, edit: 'cost:' + i }); put(r, 'G', { f: `SUM(G${r + 1}:G${last0})`, v: '', s: S.subVal });
        r++;
        const ctoRow = r;
        const itemRow = (pn, desc, q, lu, hi) => {
          put(r, 'A', { v: pn, s: hi ? S.cto : S.item }); put(r, 'B', { v: desc, s: hi ? S.cto : S.item }); put(r, 'C', { v: String(q), s: S.item }); put(r, 'D', { v: lu, s: S.item });
          put(r, 'E', { f: `C${r} * D${r}`, v: r2(q * lu), s: S.item }); put(r, 'F', { v: '', s: S.item }); put(r, 'G', { f: `C${r} * F${r}`, v: '', s: S.item }); put(r, 'H', { v: '', s: S.esd }); put(r, 'I', { v: '', s: S.item }); r++;
        };
        const feats = (list) => { if (!list.length) return; ['Código De Característica', 'Descripción', 'Cantidad'].forEach((t, k) => put(r, COLS[k], { v: t, s: S.featHdr })); r++; for (const f of list) { put(r, 'A', { v: f.code, s: S.feat }); put(r, 'B', { v: f.desc, s: S.feat }); put(r, 'C', { v: String(+f.qty.toFixed(2)), s: S.feat }); r++; } };
        itemRow(c.cto, c.name, c.qty, c.listUnit, true); feats(c.features);
        c.items.forEach((it) => { itemRow(it.pn, it.desc, it.qty, it.listUnit); feats(it.features); });
        refs.push({ ctoRow, subRow: sub, cto: c.cto, name: c.name, finalSub: ov !== undefined ? ov : 0 });
      });
      const subs = refs.map((x) => x.subRow);
      put(r, 'D', { v: 'Total:', s: S.item }); put(r, 'E', { f: `SUM(E${subs[0]}:E${r - 1}) - ${subs.map((x) => 'E' + x).join(' - ')}`, v: '', s: S.item });
      put(r, 'F', { v: 'Total:', s: S.item }); put(r, 'G', { f: `SUM(G${subs[0]}:G${r - 1}) - ${subs.map((x) => 'G' + x).join(' - ')}`, v: '', s: S.item });
    }
    const merges = ['B1:G1', 'B2:F2', 'D3:E3'];
    return { rows: [...rows.values()].sort((a, b) => a.r - b.r), merges, refs };
  }

  // ---------- hoja Propuesta ----------
  const ST = {
    cliente: 64, clienteFill: 65, conf2: 66, titulo: 67, tituloFill: 68, aten: 69, fecha: 70, cotiz: 71, conf5: 72,
    hdrL: 23, hdrR: 4, F: 2,
    firstB: 24, firstC: 27, firstD: 28, firstE: 25,
    B: 5, Bend: 6, D: 7, Dend: 8, Cfirst: 42, C: 32, Cwar: 35, Csec: 27, Cnote: 58, CsvcT: 59, Csvc: 60, Cend: 38,
    gRef: 26, gSvc: 9, I: 10, J: 11, K: 12, L: 12, M: 13, N: 14, O: 11,
    tG: 15, tH: 16, tI: 17, tJ: 18, tK: 19, tL: 20, tM: 21, tN: 22, tO: 3,
    iva: 61, total: 36, totalV: 37, condH: 62, cond: 63,
  };
  const eF = (r) => `IF(ISNUMBER($B${r}),IF($B${r}*$D${r}>0,$B${r}*$D${r},""),"")`;

  function layoutProposal(model, lenovo) {
    const rows = []; const merges = [];
    const row = (r, ht = 15) => { const o = { r, ht, custom: true, cells: {} }; rows.push(o); return o; };
    const H = model.header;
    // Encabezado
    let x = row(2, 27); x.cells.B = { v: H.cliente, s: ST.cliente, edit: 'h.cliente' }; 'CDE'.split('').forEach((c) => (x.cells[c] = { s: ST.clienteFill })); x.cells.F = { s: ST.F }; 'GHIJKLMNO'.split('').forEach((c) => (x.cells[c] = { s: ST.conf2 }));
    x = row(3, 21.95); x.cells.B = { v: H.titulo, s: ST.titulo, edit: 'h.titulo' }; 'CDE'.split('').forEach((c) => (x.cells[c] = { s: ST.tituloFill })); 'FGHIJKLMNO'.split('').forEach((c) => (x.cells[c] = { s: ST.F }));
    x = row(4, 18.95); x.cells.B = { v: H.atencion, s: ST.aten, edit: 'h.atencion' }; 'CDE'.split('').forEach((c) => (x.cells[c] = { s: ST.aten })); 'FGHIJKLMNO'.split('').forEach((c) => (x.cells[c] = { s: ST.F }));
    x = row(5, 18.95); x.cells.B = { v: H.fecha, s: ST.fecha, edit: 'h.fecha' }; x.cells.C = { s: ST.fecha }; x.cells.D = { v: H.cotizacion, s: ST.cotiz, edit: 'h.cotizacion' }; x.cells.E = { s: ST.cotiz }; x.cells.F = { s: ST.F };
    x.cells.G = { v: 'Sección clasificada como CONFIDENCIAL por CORESOLUTIONS', s: ST.conf5 }; 'HIJKLMNO'.split('').forEach((c) => (x.cells[c] = { s: ST.conf5 }));
    x = row(6, 18);
    ['Cant.', 'Descripción', 'Unitario', 'Total'].forEach((t, k) => (x.cells['BCDE'[k]] = { v: t, s: ST.hdrL }));
    x.cells.F = { s: ST.F };
    ['SKU', 'Detalle', 'Cantidad', 'Costo', 'Costo Tot.', 'Precio', 'Precio Tot.', 'Margen', 'Factor'].forEach((t, k) => (x.cells['GHIJKLMNO'[k]] = { v: t, s: ST.hdrR }));
    merges.push('B2:E2', 'G2:O2', 'B3:E3', 'B4:E4', 'B5:C5', 'D5:E5', 'G5:O5');

    let r = 7;
    const body = (opts = {}) => { const o = row(r); o.cells.B = { s: ST.B }; o.cells.C = { v: '', s: ST.C }; o.cells.D = { s: ST.D }; o.cells.E = { f: eF(r), v: '', s: ST.D, t: 'str' }; o.cells.F = { s: ST.F }; r++; return o; };
    const values = { totals: [] };
    let grand = 0;
    model.blocks.forEach((b, bi) => {
      if (b.kind === 'item') {
        if (b.title !== undefined && b.title !== null) { const o = body(); o.cells.C = { v: b.title, s: ST.Csec, edit: `b.${bi}.title` }; o.blockRef = bi; }
        const nCost = b.cost.length; const totIdx = nCost + 1; const N = Math.max(b.lines.length, totIdx + 1);
        const start = r; const totRow = start + totIdx;
        const ref = lenovo.refs[b.cfg] || null;
        let kSum = 0, mSum = 0, nSum = 0;
        for (let i = 0; i < N; i++) {
          const o = body(); o.blockRef = bi; o.lineIdx = i;
          const line = b.lines[i];
          if (i === 0) { o.cells.B = { v: b.qty, s: ST.B, edit: `b.${bi}.qty`, num: true }; o.cells.D = { f: `M${totRow}`, v: 0, s: ST.D }; }
          if (line !== undefined) o.cells.C = { v: line, s: i === 0 ? ST.Cfirst : /^\s*\*\*/.test(line) ? ST.Cwar : ST.C, edit: `b.${bi}.lines.${i}` };
          else o.cells.C = { v: '', s: ST.C, edit: `b.${bi}.lines.${i}`, phantom: true };
          if (i < nCost) {
            const c = b.cost[i];
            let sku, det, cant, costo;
            if (c.kind === 'ref' && ref) {
              sku = { f: `_Lenovo!A${ref.ctoRow}`, v: ref.cto, s: ST.gRef, t: 'str' };
              det = { f: `_Lenovo!B${ref.ctoRow}`, v: ref.name, s: ST.gRef, t: 'str' };
              cant = 1; costo = { f: `_Lenovo!F${ref.subRow}`, v: ref.finalSub, s: ST.J };
            } else {
              sku = { v: c.sku || '', s: ST.gSvc, edit: `b.${bi}.cost.${i}.sku` };
              det = { v: c.detalle || '', s: ST.gSvc, edit: `b.${bi}.cost.${i}.detalle` };
              cant = c.cant; costo = { v: c.costo === '' || c.costo === undefined ? '' : +c.costo, s: ST.J, edit: `b.${bi}.cost.${i}.costo`, num: true };
            }
            const I = num(cant), J = num(costo.v), O = num(c.factor);
            o.cells.G = sku; o.cells.H = det;
            o.cells.I = { v: cant === '' || cant === undefined ? '' : +cant, s: ST.I, edit: c.kind === 'ref' ? null : `b.${bi}.cost.${i}.cant`, num: true };
            o.cells.J = costo;
            const K = isNaN(I) || isNaN(J) ? 0 : (I * J > 0 ? I * J : 0);
            const Lp = isNaN(J) ? 0 : r2((J || 0) / (100 - (O || 0)) * 100);
            const M = isNaN(J) ? 0 : (I * Lp > 0 ? I * Lp : 0);
            const Nn = isNaN(J) ? 0 : (Lp - J > 0 ? (Lp - J) * I : 0);
            o.cells.K = { f: `IF(ISNUMBER($I${r - 1}),IF($I${r - 1}*$J${r - 1}>0,$I${r - 1}*$J${r - 1},0),0)`, v: K, s: ST.K };
            o.cells.L = { f: `ROUND((J${r - 1}/(100-O${r - 1})*100),2)`, v: Lp, s: ST.L };
            o.cells.M = { f: `IF(ISNUMBER($J${r - 1}),IF($I${r - 1}*$L${r - 1}>0,$I${r - 1}*$L${r - 1},0),0)`, v: M, s: ST.M };
            o.cells.N = { f: `IF(ISNUMBER($J${r - 1}),IF($L${r - 1}-$J${r - 1}>0,($L${r - 1}-$J${r - 1})*$I${r - 1},0),0)`, v: Nn, s: ST.N };
            o.cells.O = { v: c.factor === '' || c.factor === undefined ? '' : +c.factor, s: ST.O, edit: `b.${bi}.cost.${i}.factor`, num: true };
            o.costIdx = i;
            kSum += K; mSum += M; nSum += Nn;
          } else if (i < totIdx) {
            const rr = r - 1;
            o.cells.G = { s: ST.gSvc }; o.cells.H = { s: ST.gSvc }; o.cells.I = { s: ST.I }; o.cells.J = { s: ST.J };
            o.cells.K = { f: `IF(ISNUMBER($I${rr}),IF($I${rr}*$J${rr}>0,$I${rr}*$J${rr},0),0)`, v: 0, s: ST.K };
            o.cells.L = { f: `ROUND((J${rr}/(100-O${rr})*100),2)`, v: 0, s: ST.L };
            o.cells.M = { f: `IF(ISNUMBER($J${rr}),IF($I${rr}*$L${rr}>0,$I${rr}*$L${rr},0),0)`, v: 0, s: ST.M };
            o.cells.N = { f: `IF(ISNUMBER($J${rr}),IF($L${rr}-$J${rr}>0,($L${rr}-$J${rr})*$I${rr},0),0)`, v: 0, s: ST.N };
            o.cells.O = { s: ST.O };
          } else if (i === totIdx) {
            const mR = Math.round(mSum);
            o.cells.G = { s: ST.tG }; o.cells.H = { s: ST.tH }; o.cells.I = { s: ST.tI }; o.cells.J = { v: 'TOTALES', s: ST.tJ };
            o.cells.K = { f: `SUM(K${start}:K${totRow - 1})`, v: kSum, s: ST.tK }; o.cells.L = { s: ST.tL };
            o.cells.M = { f: `ROUND(SUM(M${start}:M${totRow - 1}),0)`, v: mR, s: ST.tM };
            o.cells.N = { f: `SUM(N${start}:N${totRow - 1})`, v: nSum, s: ST.tN };
            o.cells.O = { f: `IF(M${totRow}>0,ROUND(1-(K${totRow}/M${totRow}),4),0)`, v: mR > 0 ? Math.round((1 - kSum / mR) * 10000) / 10000 : 0, s: ST.tO };
            values.totals.push({ block: bi, cost: kSum, price: mR, margin: nSum });
          }
        }
        const unit = Math.round(mSum);
        const firstRow = rows.find((q) => q.r === start); firstRow.cells.D.v = unit;
        const qn = num(b.qty); firstRow.cells.E.v = !isNaN(qn) && qn * unit > 0 ? qn * unit : '';
        if (firstRow.cells.E.v !== '') { delete firstRow.cells.E.t; grand += firstRow.cells.E.v; }
        body().blockRef = bi; // separador
      } else if (b.kind === 'note') {
        b.lines.forEach((l, i) => { const o = body(); o.blockRef = bi; o.lineIdx = i; o.cells.C = { v: l, s: ST.Cnote, edit: `b.${bi}.lines.${i}` }; });
        body().blockRef = bi;
      } else if (b.kind === 'text') {
        if (b.title) { const o = body(); o.blockRef = bi; o.cells.C = { v: b.title, s: ST.CsvcT, edit: `b.${bi}.title` }; }
        b.lines.forEach((l, i) => { const o = body(); o.blockRef = bi; o.lineIdx = i; o.cells.C = { v: l, s: ST.Csvc, edit: `b.${bi}.lines.${i}` }; });
        body().blockRef = bi;
      }
    });
    // relleno hasta la altura de página, como la plantilla (cuerpo hasta fila 44)
    while (r < 44) body();
    // última fila del cuerpo con borde inferior
    const last = body(); last.cells.B = { s: ST.Bend }; last.cells.C = { v: '', s: ST.Cend }; last.cells.D = { s: ST.Dend }; last.cells.E = { f: eF(r - 1), v: '', s: ST.Dend, t: 'str' };
    // estilo de la primera fila del cuerpo (borde superior)
    const first = rows.find((q) => q.r === 7);
    if (first) {
      first.cells.B = Object.assign({}, first.cells.B, { s: ST.firstB });
      if (!first.cells.C.edit || /title$/.test(first.cells.C.edit)) first.cells.C = Object.assign({}, first.cells.C, { s: ST.firstC });
      first.cells.D = Object.assign({}, first.cells.D, { s: ST.firstD }); first.cells.E = Object.assign({}, first.cells.E, { s: ST.firstE });
    }
    const totRowN = r;
    x = row(r); x.cells.B = { v: model.notaIVA, s: ST.iva, edit: 'notaIVA' }; x.cells.C = { s: ST.iva }; x.cells.D = { v: 'TOTAL', s: ST.total };
    x.cells.E = { f: `SUM(E7:E${r - 1})`, v: grand, s: ST.totalV }; x.cells.F = { s: ST.F }; merges.push(`B${r}:C${r}`); r++;
    x = row(r); x.cells.F = { s: ST.F }; r++;
    x = row(r); x.cells.B = { v: 'Condiciones comerciales:', s: ST.condH }; 'CDE'.split('').forEach((c) => (x.cells[c] = { s: ST.condH })); x.cells.F = { s: ST.F }; merges.push(`B${r}:E${r}`); r++;
    model.condiciones.forEach((t, i) => {
      x = row(r, i === model.condiciones.length - 1 ? 24 : 15); x.condIdx = i;
      x.cells.B = { v: t, s: ST.cond, edit: `cond.${i}`, rich: true }; 'CDE'.split('').forEach((c) => (x.cells[c] = { s: ST.cond })); x.cells.F = { s: ST.F }; merges.push(`B${r}:E${r}`); r++;
    });
    return { rows, merges, lastRow: r - 1, totalRow: totRowN, grand, totals: values.totals };
  }

  // ---------- XML ----------
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  function cellXML(ref, c) {
    const s = c.s !== undefined ? ` s="${c.s}"` : '';
    if (c.f) {
      const isStr = c.t === 'str' || (typeof c.v === 'string');
      const v = c.v === '' || c.v === undefined || c.v === null ? '' : `<v>${esc(c.v)}</v>`;
      return `<c r="${ref}"${s}${isStr ? ' t="str"' : ''}><f>${esc(c.f)}</f>${v}</c>`;
    }
    if (c.v === undefined || c.v === null || c.v === '') return c.blank || c.v === '' && c.keepEmpty ? `<c r="${ref}"${s}/>` : `<c r="${ref}"${s}/>`;
    if (typeof c.v === 'number' && isFinite(c.v)) return `<c r="${ref}"${s}><v>${c.v}</v></c>`;
    if (c.rich) {
      const k = String(c.v).indexOf(':');
      if (k > 0 && k < 40) {
        const rp = (b) => `<rPr>${b ? '<b/>' : ''}<sz val="8"/><color theme="1"/><rFont val="Arial"/><family val="2"/></rPr>`;
        return `<c r="${ref}"${s} t="inlineStr"><is><r>${rp(true)}<t xml:space="preserve">${esc(String(c.v).slice(0, k + 1))}</t></r><r>${rp(false)}<t xml:space="preserve">${esc(String(c.v).slice(k + 1))}</t></r></is></c>`;
      }
    }
    return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(c.v)}</t></is></c>`;
  }
  function sheetDataXML(rows) {
    return rows.map((o) => {
      const cells = Object.keys(o.cells).sort((a, b) => colIdx(a) - colIdx(b)).map((c) => cellXML(c + o.r, o.cells[c])).join('');
      const ht = o.ht ? ` ht="${o.ht}"${o.custom ? ' customHeight="1"' : ''}` : '';
      return `<row r="${o.r}"${ht}>${cells}</row>`;
    }).join('');
  }
  const mergesXML = (m) => (m.length ? `<mergeCells count="${m.length}">${m.map((x) => `<mergeCell ref="${x}"/>`).join('')}</mergeCells>` : '');

  async function exportXLSX(templateBuf, JSZip, model, parsed) {
    const lenovo = buildLenovo(parsed, model);
    const prop = layoutProposal(model, lenovo);
    const zip = await JSZip.loadAsync(templateBuf);
    let s1 = await zip.file('xl/worksheets/sheet1.xml').async('string');
    s1 = s1.replace('{{DIM}}', `B2:O${prop.lastRow}`).replace('{{DATA}}', sheetDataXML(prop.rows)).replace('{{MERGES}}', mergesXML(prop.merges))
      .replace('{{FIRMA}}', esc(model.firma.nombre)).replace('{{CEL}}', esc(model.firma.cel)).replace('{{EMAIL}}', esc(model.firma.email));
    zip.file('xl/worksheets/sheet1.xml', s1);
    let s2 = await zip.file('xl/worksheets/sheet2.xml').async('string');
    const lastL = lenovo.rows.length ? lenovo.rows[lenovo.rows.length - 1].r : 1;
    s2 = s2.replace('{{DIM}}', `A1:L${lastL}`).replace('{{DATA}}', sheetDataXML(lenovo.rows)).replace('{{MERGES}}', mergesXML(lenovo.merges));
    zip.file('xl/worksheets/sheet2.xml', s2);
    let wbx = await zip.file('xl/workbook.xml').async('string');
    wbx = wbx.replace('{{PRINT}}', `Propuesta!$B$2:$E$${prop.lastRow}`);
    zip.file('xl/workbook.xml', wbx);
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }

  function fileName(model, today) {
    const d = today || new Date();
    const cli = (model.header.cliente || 'Cliente').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Za-z0-9]+/g, ' ').trim().split(' ').slice(0, 3).map((w) => w[0] + w.slice(1).toLowerCase()).join('');
    return `${cli}_${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}-propuesta.xlsx`;
  }

  const api = { parseWorkbook, describe, unknownParts, buildDefault, buildLenovo, layoutProposal, exportXLSX, fileName, num };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.CoreEngine = api;
})(typeof window !== 'undefined' ? window : globalThis);
