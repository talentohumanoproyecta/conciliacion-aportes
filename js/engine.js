/*
 * Motor de conciliación nómina (contabilidad) vs planilla PILA (operador).
 * Funciona en el navegador (window.Concilia) y en Node (require) para pruebas.
 * No envía datos a ningún servidor: todo se procesa en memoria.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Concilia = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ------------------------------------------------------------------ utilidades
  const SUBS = { sal: 'Salud', pen: 'Pensión', arl: 'ARL', ccf: 'Caja de compensación' };
  const ARL_CLASES = [0.00522, 0.01044, 0.02436, 0.0435, 0.0696];
  const TARIFA_REF = { sal: 0.04, pen: 0.16, ccf: 0.04 };

  function norm(s) {
    if (s === null || s === undefined) return '';
    return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[ñÑ]/g, 'N').toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();
  }
  function num(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    let s = String(v).trim().replace(/[^\d,.\-]/g, '');
    if (!s) return 0;
    if (s.includes(',') && s.includes('.')) s = s.replace(/,/g, '');
    else if (s.includes(',')) s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }
  function ced(v) {
    if (v === null || v === undefined) return '';
    const d = String(typeof v === 'number' ? Math.round(v) : v).replace(/\D/g, '').replace(/^0+/, '');
    return d;
  }
  function txt(v) { return v === null || v === undefined ? '' : String(v).trim(); }
  function tokens(s) { return norm(s).split(' ').filter(t => t.length > 1); }
  const round = n => Math.round(n);

  function subFromAccount(code) {
    const d = String(code || '').replace(/\D/g, '');
    if (d.startsWith('237005')) return 'sal';
    if (d.startsWith('237006')) return 'arl';
    if (d.startsWith('237010')) return 'ccf';
    if (d.startsWith('238030')) return 'pen';
    return null;
  }
  const NIT_SENA = '899999034', NIT_ICBF = '899999239';

  // Entidades conocidas (código PILA -> NIT). El catálogo y el resumen del operador las complementan.
  const ENTIDADES_BASE = {
    '230201': ['800229739', 'Protección'], '230301': ['800224808', 'Porvenir'],
    '231001': ['800227940', 'Colfondos'], '25-14': ['900336004', 'Colpensiones'],
    '230901': ['800253055', 'Skandia'],
    'EPS001': ['830113831', 'Aliansalud'], 'EPS002': ['800130907', 'Salud Total'],
    'EPS005': ['800251440', 'Sanitas'], 'EPS008': ['860066942', 'Compensar EPS'],
    'EPS010': ['800088702', 'EPS Sura'], 'EPS017': ['830003564', 'Famisanar'],
    'EPS037': ['900156264', 'Nueva EPS'], 'EPS041': ['900156264', 'Nueva EPS Movilidad'],
    'ESSC62': ['900935126', 'Asmet Salud'], 'EPS062': ['900935126', 'Asmet Salud'],
    'CCF22': ['860007336', 'Colsubsidio'], 'CCF24': ['860066942', 'Compensar'],
    'CCF32': ['891180008', 'Comfamiliar Huila'],
    '14-7': ['860002503', 'Seguros Bolívar ARL'],
  };

  // Palabras que no sirven para comparar nombres de entidades
  const STOP = new Set(['EPS', 'S', 'SA', 'SAS', 'DE', 'LA', 'EL', 'LOS', 'Y', 'CAJA', 'COMPENSACION', 'FAMILIAR',
    'FONDO', 'FONDOS', 'PENSIONES', 'CESANTIAS', 'ENTIDAD', 'PROMOTORA', 'SALUD', 'ADMINISTRADORA', 'COMPANIA',
    'SEGUROS', 'ARL', 'CCF', 'DEL', 'MOVILIDAD', 'ENTIDADES', 'PROMOTORAS', 'COLOMBIANA', 'FLIAR']);
  function entityWords(s) { return tokens(s).filter(t => !STOP.has(t) && t.length > 2); }

  // ------------------------------------------------------------------ lectura de archivos
  function sheetRows(XLSX, ws) {
    if (!ws || !ws['!ref']) return [];
    const r = XLSX.utils.decode_range(ws['!ref']);
    r.s.c = 0; r.s.r = 0;
    ws['!ref'] = XLSX.utils.encode_range(r);
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }

  function findRow(rows, pred, max) {
    const lim = Math.min(rows.length, max || rows.length);
    for (let i = 0; i < lim; i++) if (rows[i] && pred(rows[i].map(norm), rows[i])) return i;
    return -1;
  }
  function colOf(normRow, label, from, to) {
    const L = norm(label);
    for (let j = from || 0; j < (to === undefined ? normRow.length : to); j++) if (normRow[j] === L) return j;
    return -1;
  }
  function colStarts(normRow, label, from, to) {
    const L = norm(label);
    for (let j = from || 0; j < (to === undefined ? normRow.length : to); j++) if (normRow[j] && normRow[j].startsWith(L)) return j;
    return -1;
  }

  /** Lee un archivo y devuelve {tipo, ...datos}. data: ArrayBuffer/Uint8Array/Buffer. */
  function leerArchivo(XLSX, data, nombre) {
    const lower = (nombre || '').toLowerCase();
    if (lower.endsWith('.txt')) {
      const text = decodeText(data);
      if (/^01/.test(text) && /\n02\d{5}/.test(text.replace(/\r/g, ''))) return leerPlano(text, nombre);
      throw new Error('El archivo de texto no tiene formato de planilla PILA (Resolución 2388).');
    }
    const wb = XLSX.read(data, { type: 'array', cellDates: false });
    const sheets = wb.SheetNames.map(n => ({ name: n, rows: sheetRows(XLSX, wb.Sheets[n]) }));
    const first = sheets[0].rows;
    const head = first.slice(0, 40).map(r => (r || []).map(norm).join('|')).join('\n');

    if (sheets.some(s => norm(s.name) === 'CENTROS DE TRABAJO')) return leerCatalogo(sheets, nombre);
    if (head.includes('LIQUIDACION DETALLADA DE APORTES')) return leerAportesEnLinea(sheets, nombre);
    if (head.includes('NO DE IDENTIFICACION') && head.includes('CODIGO EPS')) return leerMiPlanilla(sheets, nombre);
    if (head.includes('CODIGO CONTABLE') && head.includes('DETALLE')) return leerSiigo(first, nombre);
    if (/(^|\|)TIPO\|/.test(head) && head.includes('EMPLEADO') && head.includes('CUENTA')) return leerMekano(first, nombre);
    throw new Error('No reconozco este archivo. Formatos soportados: auxiliar Siigo, auxiliar Mekano, plano PILA (.txt), planilla Aportes en Línea, planilla Mi Planilla y catálogo.');
  }

  function decodeText(data) {
    const u8 = data instanceof Uint8Array ? data : new Uint8Array(data);
    try { return new TextDecoder('utf-8', { fatal: true }).decode(u8); }
    catch (e) { return new TextDecoder('latin1').decode(u8); }
  }

  // ---------- Auxiliar Siigo
  function leerSiigo(rows, archivo) {
    const h = findRow(rows, r => r.includes('CODIGO CONTABLE') && r.includes('DETALLE'), 40);
    const H = rows[h].map(norm);
    const c = {
      cta: colOf(H, 'Código contable'), nit: colOf(H, 'Identificación'), ter: colOf(H, 'Nombre tercero'),
      comp: colOf(H, 'Comprobante'), fecha: colOf(H, 'Fecha elaboración'), det: colOf(H, 'Detalle'),
      deb: colOf(H, 'Débito'), cre: colOf(H, 'Crédito'),
    };
    const empresa = txt((rows[2] || [])[0]), nitEmp = ced(txt((rows[3] || [])[0]).split('-')[0]), rango = txt((rows[4] || [])[0]);
    const registros = [], pagos = [], fechas = new Set(), otrasCuentas = new Set();
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r) continue;
      const cta = txt(r[c.cta]); if (!/^\d{6,}$/.test(cta)) continue;
      const sub = subFromAccount(cta);
      if (!sub) { otrasCuentas.add(cta); continue; }
      const deb = num(r[c.deb]), cre = num(r[c.cre]);
      const det = txt(r[c.det]);
      const m = det.match(/^\s*(n[oó]mina|aporte\s+patronal)\s*-\s*(.+)$/i);
      const entNit = ced(r[c.nit]);
      if (m) {
        let s = sub;
        if (sub === 'ccf' && (entNit === NIT_SENA || entNit === NIT_ICBF)) s = 'par';
        registros.push({
          sub: s, entNit, entNombre: txt(r[c.ter]), nombre: m[2].trim(), cedula: null,
          tipo: /n[oó]mina/i.test(m[1]) ? 'empleado' : 'patronal', valor: cre - deb,
          fecha: txt(r[c.fecha]), comprobante: txt(r[c.comp]),
        });
        fechas.add(txt(r[c.fecha]));
      } else if (deb) {
        pagos.push({ sub, entNit, entNombre: txt(r[c.ter]), valor: deb, fecha: txt(r[c.fecha]), comprobante: txt(r[c.comp]) });
      }
    }
    return {
      tipo: 'nomina', fuente: 'Siigo', archivo, empresa, nitEmpresa: nitEmp, periodo: rango,
      registros, pagos, fechas: [...fechas], avisos: otrasCuentas.size ? ['Cuentas ignoradas: ' + [...otrasCuentas].join(', ')] : [],
    };
  }

  // ---------- Auxiliar Mekano
  function excelDate(v) {
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return d.toISOString().slice(0, 10);
    }
    return txt(v);
  }
  function leerMekano(rows, archivo) {
    const h = findRow(rows, r => r.includes('TIPO') && r.includes('EMPLEADO') && r.includes('CUENTA'), 40);
    const H = rows[h].map(norm);
    const c = {
      tipo: colOf(H, 'TIPO'), num: colOf(H, 'NUMERO'), fecha: colOf(H, 'FECHA'), cta: colOf(H, 'CUENTA'),
      ter: colOf(H, 'TERCERO'), emp: colOf(H, 'EMPLEADO'), deb: colOf(H, 'DEBITO'), cre: colOf(H, 'CREDITO'),
      mod: colOf(H, 'MODULO'), nota: colOf(H, 'NOTA'), nter: colOf(H, 'NOMBRE TERCERO'), nemp: colOf(H, 'NOMBRE_EMPLEADO'),
    };
    if (c.nemp < 0) c.nemp = colOf(H, 'NOMBRE EMPLEADO');
    const registros = [], pagos = [], notas = new Set(), otrasCuentas = new Set();
    for (let i = h + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r) continue;
      const tipo = norm(r[c.tipo]); if (!tipo) continue; // saldos iniciales
      const cta = txt(r[c.cta]);
      const sub = subFromAccount(cta);
      if (!sub) { if (cta) otrasCuentas.add(cta); continue; }
      const emp = ced(r[c.emp]);
      const deb = num(r[c.deb]), cre = num(r[c.cre]);
      const esNomina = emp && (tipo === 'NOM' || norm(r[c.mod]).includes('NOMINA'));
      const entNit = ced(r[c.ter]);
      if (esNomina) {
        let s = sub;
        if (sub === 'ccf' && (entNit === NIT_SENA || entNit === NIT_ICBF)) s = 'par';
        registros.push({
          sub: s, entNit, entNombre: txt(r[c.nter]), nombre: txt(r[c.nemp]), cedula: emp,
          tipo: 'mixto', valor: cre - deb, fecha: excelDate(r[c.fecha]), comprobante: tipo + ' ' + txt(r[c.num]),
        });
        notas.add(txt(r[c.nota]));
      } else if (deb) {
        pagos.push({ sub, entNit, entNombre: txt(r[c.nter]), valor: deb, fecha: excelDate(r[c.fecha]), comprobante: tipo + ' ' + txt(r[c.num]), nota: txt(r[c.nota]) });
      }
    }
    return {
      tipo: 'nomina', fuente: 'Mekano', archivo, empresa: '', nitEmpresa: '', periodo: [...notas].join(' / '),
      registros, pagos, fechas: [], avisos: otrasCuentas.size ? ['Cuentas ignoradas: ' + [...otrasCuentas].join(', ')] : [],
    };
  }

  // ---------- Plano PILA (Res. 2388) generado por Siigo
  function leerPlano(text, archivo) {
    const lines = text.replace(/\r/g, '').split('\n').filter(l => l.length > 10);
    const g = (l, a, b) => l.substring(a - 1, b);
    const h = lines[0];
    const empresa = g(h, 8, 207).trim();
    const mNit = h.match(/NI(\d{5,})\s+(\d)/);
    const mPer = h.match(/(\d{4}-\d{2})(\d{4}-\d{2})/);
    const NOV = ['ING', 'RET', 'TDE', 'TAE', 'TDP', 'TAP', 'VSP', 'COR', 'VST', 'SLN', 'IGE', 'LMA', 'VAC', 'AVP', 'VCT'];
    const lineas = [];
    for (const l of lines) {
      if (!l.startsWith('02')) continue;
      const nov = [];
      NOV.forEach((n, k) => { if (g(l, 137 + k, 137 + k).trim()) nov.push(n); });
      if (g(l, 152, 153).trim() && g(l, 152, 153).trim() !== '00') nov.push('IRL');
      const n = (a, b) => num(g(l, a, b));
      lineas.push({
        cedula: ced(g(l, 10, 25)), tipoDoc: g(l, 8, 9).trim(),
        nombre: [g(l, 37, 56), g(l, 57, 86), g(l, 87, 106), g(l, 107, 136)].map(s => s.trim()).filter(Boolean).join(' '),
        nombreCompleto: [g(l, 87, 106), g(l, 107, 136), g(l, 37, 56), g(l, 57, 86)].map(s => s.trim()).filter(Boolean).join(' '),
        centro: String(parseInt(g(l, 390, 398), 10) || 0), centroNombre: '',
        nov,
        sub: {
          pen: { cod: g(l, 154, 159).trim(), dias: n(184, 185), ibc: n(202, 210), val: n(245, 253) },
          sal: { cod: g(l, 166, 171).trim(), dias: n(186, 187), ibc: n(211, 219), val: n(315, 323) },
          arl: { cod: '', dias: n(188, 189), ibc: n(220, 228), val: n(399, 407), tarifa: n(381, 389) },
          ccf: { cod: g(l, 178, 183).trim(), dias: n(190, 191), ibc: n(229, 237), val: n(415, 423) },
        },
      });
    }
    const arl = (lines[0].match(/(14-\d+)\s/) || [])[1] || '';
    lineas.forEach(x => { x.sub.arl.cod = arl; });
    return {
      tipo: 'plano', fuente: 'Plano PILA Siigo', archivo, empresa, nitEmpresa: mNit ? mNit[1] : '',
      periodo: mPer ? mPer[1] : '', lineas,
    };
  }

  // ---------- Planilla Aportes en Línea (detalle "Planilla Resumen")
  function leerAportesEnLinea(sheets, archivo) {
    const rows = sheets[0].rows;
    let empresa = '', nit = '', periodo = '', numero = '', pagada = '';
    const iA = findRow(rows, r => r.includes('RAZON SOCIAL'), 30);
    if (iA >= 0) {
      const H = rows[iA].map(norm), V = rows[iA + 1] || [];
      empresa = txt(V[colOf(H, 'Razon Social')]);
      nit = ced(V[colOf(H, 'Identificación')]);
    }
    const iP = findRow(rows, (r, raw) => raw.some(v => /^\d{4}-\d{2}$/.test(txt(v))), 40);
    if (iP >= 0) {
      const vals = rows[iP].filter(v => v !== null && v !== '');
      periodo = txt(vals[0]); numero = txt(vals[3]);
      const pago = vals.find(v => v instanceof Date || (typeof v === 'number' && v > 40000));
      pagada = pago ? 'Sí' : '';
    }
    const ih = findRow(rows, r => r.includes('IDENTIFICACION') && r.includes('NOMBRE') && r.includes('ING'), 60);
    if (ih < 0) throw new Error('No encontré el encabezado del detalle de aportes.');
    const H = rows[ih].map(norm), G = rows[ih - 1].map(norm);
    const grupos = {};
    const orden = [['pen', 'PENSION'], ['sal', 'SALUD'], ['ccf', 'CCF'], ['arl', 'RIESGOS'], ['par', 'PARAFISCALES']];
    const starts = orden.map(([k, lbl]) => [k, colOf(G, lbl)]).filter(x => x[1] >= 0).sort((a, b) => a[1] - b[1]);
    starts.forEach(([k, s], i) => {
      const e = i + 1 < starts.length ? starts[i + 1][1] : H.length;
      grupos[k] = {
        cod: colOf(H, 'CODIGO', s, e), dias: colStarts(H, 'DIAS', s, e), ibc: colOf(H, 'IBC', s, e),
        val: colOf(H, 'APORTE', s, e), tarifa: colOf(H, 'TARIFA', s, e),
      };
    });
    const cId = colOf(H, 'IDENTIFICACION'), cNom = colOf(H, 'NOMBRE'), cNo = colOf(H, 'NO');
    const novCols = [];
    for (let j = 0; j < H.length; j++) {
      if (['ING', 'RET', 'TDE', 'TAE', 'TDP', 'TAP', 'VSP', 'COR', 'VST', 'SLN', 'IGE', 'LMA', 'VAC', 'AVP', 'VCT', 'IRL'].includes(H[j])) novCols.push([H[j], j]);
    }
    const lineas = [];
    let sucursal = '', centro = '', ciudad = '';
    for (let i = ih + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r) continue;
      const a = txt(r[cNo]);
      if (/^SUCURSAL/i.test(a)) { sucursal = a.replace(/^SUCURSAL:\s*/i, '').replace(/\(\d+ Afiliados\)/i, '').trim(); continue; }
      if (/^Centro de Trabajo/i.test(a)) { centro = a.replace(/^Centro de Trabajo:\s*/i, '').replace(/\(\d+ Afiliados\)/i, '').trim(); continue; }
      if (/^Ciudad/i.test(a)) { ciudad = a.replace(/\(\d+ Afiliados\)/i, '').trim(); continue; }
      if (typeof r[cNo] !== 'number' && !/^\d+$/.test(a)) continue;
      const doc = ced(r[cId + 1]); if (!doc) continue;
      const sub = {};
      for (const k of ['pen', 'sal', 'ccf', 'arl']) {
        const gk = grupos[k]; if (!gk) continue;
        sub[k] = {
          cod: txt(r[gk.cod]), dias: num(r[gk.dias]), ibc: num(r[gk.ibc]), val: num(r[gk.val]),
          tarifa: gk.tarifa >= 0 ? num(r[gk.tarifa]) : undefined,
        };
      }
      lineas.push({
        cedula: doc, tipoDoc: txt(r[cId]), nombre: txt(r[cNom]), centro: '', centroNombre: centro, sucursal, ciudad,
        nov: novCols.filter(([, j]) => txt(r[j])).map(([n]) => n), sub, planilla: numero,
      });
    }
    // Resumen de pago: código -> NIT
    const entidades = {};
    for (const s of sheets.slice(1)) {
      const ir = findRow(s.rows, r => r.includes('CODIGO') && r.includes('NIT'), 30);
      if (ir < 0) continue;
      const H2 = s.rows[ir].map(norm);
      const cc = colOf(H2, 'CODIGO'), cn = colOf(H2, 'NIT'), cr = colOf(H2, 'RIESGO');
      for (let i = ir + 1; i < s.rows.length; i++) {
        const r = s.rows[i]; if (!r) continue;
        const cod = txt(r[cc]), nitE = ced(r[cn]);
        if (cod && nitE) entidades[cod] = [nitE, txt(r[cr])];
      }
    }
    return { tipo: 'planilla', fuente: 'Aportes en Línea', archivo, empresa, nitEmpresa: nit, periodo, numero, pagada, lineas, entidades, correccion: false };
  }

  // ---------- Planilla Mi Planilla (reporte "ConEncabezado")
  function leerMiPlanilla(sheets, archivo) {
    const rows = sheets[0].rows;
    const kv = {};
    for (let i = 0; i < 15 && i < rows.length; i++) {
      const r = rows[i] || [];
      for (let j = 0; j + 1 < r.length; j += 2) if (txt(r[j])) kv[norm(r[j])] = r[j + 1];
    }
    const ih = findRow(rows, r => r.includes('NO DE IDENTIFICACION'), 40);
    const H = rows[ih].map(norm), G = (rows[ih - 1] || []).map(norm);
    const gs = [['pen', 'PENSION'], ['sal', 'SALUD'], ['arl', 'RIESGOS'], ['ccf', 'CAJAS'], ['par', 'SENA E ICBF']]
      .map(([k, l]) => [k, colOf(G, l)]).filter(x => x[1] >= 0).sort((a, b) => a[1] - b[1]);
    const rng = {};
    gs.forEach(([k, s], i) => { rng[k] = [s, i + 1 < gs.length ? gs[i + 1][1] : H.length]; });
    const f = (k, lbl, exact) => rng[k] ? (exact ? colOf(H, lbl, ...rng[k]) : colStarts(H, lbl, ...rng[k])) : -1;
    const c = {
      td: colOf(H, 'Tipo de Identificación'), id: colOf(H, 'No. de Identificación'),
      a1: colOf(H, 'Primer Apellido'), a2: colOf(H, 'Segundo Apellido'), n1: colOf(H, 'Primer Nombre'), n2: colOf(H, 'Segundo Nombre'),
      corr: colOf(H, 'Correcciones'),
      pen: { cod: f('pen', 'CODIGO ADM'), nom: f('pen', 'NOMBRE FONDO'), dias: f('pen', 'DIAS COTIZADOS'), ibc: f('pen', 'IBC'), val: f('pen', 'TOTAL A PAGAR') },
      sal: { cod: f('sal', 'CODIGO EPS'), nom: f('sal', 'NOMBRE EPS'), dias: f('sal', 'DIAS COTIZADOS'), ibc: f('sal', 'IBC'), val: f('sal', 'COTIZACION OBLIGATORIA') },
      arl: { cod: f('arl', 'COD ARL'), dias: f('arl', 'DIAS COTIZADOS'), ibc: f('arl', 'IBC'), val: f('arl', 'COTIZACION OBLIGATORIA'), tarifa: f('arl', 'TARIFA CENTRO'), centro: f('arl', 'CODIGO CENTRO'), clase: f('arl', 'CLASE DE RIESGO') },
      ccf: { cod: f('ccf', 'CODIGO CCF'), nom: f('ccf', 'NOMBRE CCF'), dias: f('ccf', 'DIAS COTIZADOS'), ibc: f('ccf', 'IBC'), val: f('ccf', 'VALOR APORTE') },
    };
    const NOVS = ['ING', 'RET', 'TDE', 'TAE', 'TDP', 'TAP', 'VSP', 'VST', 'SLN', 'IGE', 'LMA', 'VAC', 'AVP', 'VCT', 'IRL'];
    const novCols = NOVS.map(n => [n, colOf(H, n)]).filter(x => x[1] >= 0);
    let lineas = [];
    const numero = txt(kv['NO PLANILLA']);
    for (let i = ih + 1; i < rows.length; i++) {
      const r = rows[i]; if (!r) continue;
      const doc = ced(r[c.id]); if (!doc) continue;
      const sub = {};
      for (const k of ['pen', 'sal', 'arl', 'ccf']) {
        const ck = c[k];
        sub[k] = { cod: txt(r[ck.cod]), nomEnt: ck.nom >= 0 ? txt(r[ck.nom]) : '', dias: num(r[ck.dias]), ibc: num(r[ck.ibc]), val: num(r[ck.val]) };
        if (k === 'arl') { sub[k].tarifa = num(r[ck.tarifa]); sub[k].clase = num(r[ck.clase]); }
      }
      lineas.push({
        cedula: doc, tipoDoc: txt(r[c.td]),
        nombre: [r[c.a1], r[c.a2], r[c.n1], r[c.n2]].map(txt).filter(Boolean).join(' '),
        centro: txt(r[c.arl.centro]), centroNombre: '', nov: novCols.filter(([, j]) => txt(r[j])).map(([n]) => n),
        sub, planilla: numero, corr: norm(r[c.corr]),
      });
    }
    let correccion = false;
    if (lineas.some(l => l.corr === 'C')) { correccion = true; lineas = lineas.filter(l => l.corr !== 'A'); }
    return {
      tipo: 'planilla', fuente: 'Mi Planilla', archivo, empresa: txt(kv['EMPRESA']), nitEmpresa: ced(kv['NIT']),
      periodo: txt(kv['PERIODO DE PAGO']), numero, pagada: norm(kv['ESTADO']) === 'PAGADA' ? 'Sí' : '',
      tipoPlanilla: txt(kv['TIPO PLANILLA']), lineas, entidades: {}, correccion,
    };
  }

  // ---------- Catálogo de empresas / centros / entidades
  function leerCatalogo(sheets, archivo) {
    const get = n => sheets.find(s => norm(s.name) === n);
    const empresas = [], centros = [], entidades = {};
    const se = get('EMPRESAS');
    if (se) {
      const H = (se.rows[0] || []).map(norm);
      const cE = colOf(H, 'Empresa'), cN = colOf(H, 'NIT'), cS = colStarts(H, 'SOFTWARE'), cO = colStarts(H, 'OPERADOR');
      se.rows.slice(1).forEach(r => { if (r && txt(r[cE])) empresas.push({ empresa: txt(r[cE]), nit: ced(r[cN]), software: txt(r[cS]), operador: txt(r[cO]) }); });
    }
    const sc = get('CENTROS DE TRABAJO');
    if (sc) {
      const H = (sc.rows[0] || []).map(norm);
      const col = l => colStarts(H, l);
      const cE = colOf(H, 'Empresa'), cI = col('CENTRO DE TRABAJO'), cP = col('NOMBRE EN PLANILLA'), cC = col('CODIGO CENTRO'),
        cT = colOf(H, 'Tipo'), cF = col('PLANILLA PROPIA'), cR = col('CLASE DE RIESGO CONFIRMADA'), cO = col('RIESGO OBSERVADO'), cOb = col('OBSERVACIONES');
      sc.rows.slice(1).forEach(r => {
        if (!r || !(txt(r[cE]) || txt(r[cI]) || txt(r[cP]))) return;
        centros.push({
          empresa: txt(r[cE]), interno: txt(r[cI]), nombrePlanilla: txt(r[cP]), codigo: txt(r[cC]),
          tipo: txt(r[cT]), propia: txt(r[cF]), riesgo: num(r[cR]) || null, riesgoObservado: num(r[cO]) || null, obs: cOb >= 0 ? txt(r[cOb]) : '',
        });
      });
    }
    const sn = get('ENTIDADES');
    if (sn) {
      const H = (sn.rows[0] || []).map(norm);
      const cc = colStarts(H, 'CODIGO'), cn = colOf(H, 'NIT'), cm = colStarts(H, 'NOMBRE');
      sn.rows.slice(1).forEach(r => { if (r && txt(r[cc]) && ced(r[cn])) entidades[txt(r[cc])] = [ced(r[cn]), txt(r[cm])]; });
    }
    return { tipo: 'catalogo', archivo, empresas, centros, entidades };
  }

  // ------------------------------------------------------------------ conciliación
  const DEFAULTS = { tolTrabajador: 1000, tolEntidad: 5000 };

  /**
   * archivos: lista de objetos devueltos por leerArchivo (nomina / plano / planilla)
   * catalogo: objeto de leerCatalogo o null
   */
  function conciliar(archivos, catalogo, opciones) {
    const opt = Object.assign({}, DEFAULTS, opciones || {});
    const nominas = archivos.filter(a => a.tipo === 'nomina');
    const planos = archivos.filter(a => a.tipo === 'plano');
    const planillas = archivos.filter(a => a.tipo === 'planilla');
    const avisos = [];
    if (!planillas.length) throw new Error('Cargue al menos una planilla del operador (Aportes en Línea o Mi Planilla).');
    if (!nominas.length && !planos.length) throw new Error('Cargue el auxiliar contable (Siigo o Mekano) o el plano PILA.');

    // Mapa de entidades
    const ent = Object.assign({}, ENTIDADES_BASE);
    if (catalogo) Object.assign(ent, catalogo.entidades || {});
    planillas.forEach(p => Object.assign(ent, p.entidades || {}));
    const nitDe = cod => (ent[cod] || [])[0] || '';
    const nombreDe = cod => (ent[cod] || [])[1] || '';

    // Empresa
    const nitEmpresa = planillas[0].nitEmpresa || (nominas[0] || {}).nitEmpresa || (planos[0] || {}).nitEmpresa;
    if (planillas.some(p => p.nitEmpresa && p.nitEmpresa !== nitEmpresa)) avisos.push('Las planillas cargadas son de empresas (NIT) distintas.');
    const empCat = catalogo ? catalogo.empresas.find(e => e.nit && e.nit === nitEmpresa) : null;
    const empresaNombre = (empCat && empCat.empresa) || planillas[0].empresa || (nominas[0] || {}).empresa || '';
    if (catalogo && !empCat) avisos.push(`La empresa NIT ${nitEmpresa} no está en la hoja "Empresas" del catálogo.`);
    planillas.forEach(p => { if (p.pagada === 'Sí') avisos.push(`La planilla ${p.numero} (${p.archivo}) ya figura como PAGADA: la revisión ideal es antes de pagar.`); });
    planillas.forEach(p => { if (p.correccion) avisos.push(`La planilla ${p.numero} es de corrección (tipo ${p.tipoPlanilla || 'N'}): se toman solo las líneas corregidas (C).`); });

    // ---- Planilla por trabajador
    const P = new Map();
    const lineasPlanilla = [];
    planillas.forEach(p => p.lineas.forEach(l => { l.planilla = l.planilla || p.numero; l.fuente = p.fuente; lineasPlanilla.push(l); }));
    for (const l of lineasPlanilla) {
      let w = P.get(l.cedula);
      if (!w) { w = { cedula: l.cedula, nombre: l.nombre, lineas: [], centros: new Set(), planillas: new Set(), nov: new Set(), sub: {} }; P.set(l.cedula, w); }
      w.lineas.push(l); w.planillas.add(l.planilla);
      w.centros.add(l.centroNombre || l.centro);
      l.nov.forEach(n => w.nov.add(n));
      for (const k of Object.keys(SUBS)) {
        const s = l.sub[k]; if (!s) continue;
        const ws = w.sub[k] || (w.sub[k] = { total: 0, dias: 0, ibc: 0, ibcTarifa: 0, cods: new Set(), byNit: {}, tarifas: new Set(), nomEnt: '' });
        ws.total += s.val; ws.dias += s.dias; ws.ibc += s.ibc;
        if (s.cod) ws.cods.add(s.cod);
        if (s.nomEnt) ws.nomEnt = s.nomEnt;
        if (k === 'arl' && s.tarifa) { ws.tarifas.add(s.tarifa); ws.ibcTarifa += s.ibc; }
        if (s.cod) {
          const key = nitDe(s.cod) || ('COD:' + s.cod);
          ws.byNit[key] = (ws.byNit[key] || 0) + s.val;
        }
      }
    }

    // ---- Validación de centros contra catálogo
    const centrosNuevos = new Map();
    const hallCentro = new Map(); // cedula -> findings
    const centrosCat = catalogo ? catalogo.centros.filter(c => !empCat || !c.empresa || norm(c.empresa) === norm(empCat.empresa)) : [];
    function buscarCentro(l) {
      const nombreN = norm(l.centroNombre), cod = txt(l.centro);
      return centrosCat.find(c => (nombreN && norm(c.nombrePlanilla) === nombreN) || (cod && c.codigo !== '' && txt(c.codigo) === cod && !nombreN));
    }
    if (catalogo) {
      for (const l of lineasPlanilla) {
        const key = l.centroNombre || l.centro; if (!key) continue;
        const c = buscarCentro(l);
        const tarifa = l.sub.arl ? l.sub.arl.tarifa : 0;
        const clase = claseDe(tarifa);
        if (!c) {
          const k = norm(key);
          const cn = centrosNuevos.get(k) || { empresa: empresaNombre, nombrePlanilla: l.centroNombre, codigo: l.centroNombre ? '' : l.centro, riesgoObservado: clase, tarifa, trabajadores: new Set() };
          cn.trabajadores.add(l.cedula); if (clase) cn.riesgoObservado = clase;
          centrosNuevos.set(k, cn);
        } else if (c.riesgo && tarifa && clase && clase !== c.riesgo) {
          const arr = hallCentro.get(l.cedula) || [];
          arr.push({
            sev: 'rojo', sub: 'arl', tipo: 'Riesgo distinto al catálogo',
            detalle: `El centro "${key}" es riesgo ${c.riesgo} en el catálogo, pero la planilla ${l.planilla} liquida riesgo ${clase} (${pct(tarifa)}).`,
            resp: 'Talento Humano', accion: 'Corregir el centro de trabajo o la clase de riesgo del trabajador en el operador.',
          });
          hallCentro.set(l.cedula, arr);
        }
      }
    }

    // ---- Fuente de nómina
    let regs = [];
    let fuenteNomina = '';
    if (nominas.length) {
      fuenteNomina = [...new Set(nominas.map(n => n.fuente))].join(' + ');
      nominas.forEach(n => regs.push(...n.registros));
    } else {
      fuenteNomina = 'Plano PILA (sin auxiliar)';
      avisos.push('No se cargó auxiliar contable: se usa el plano PILA como dato de nómina.');
      planos.forEach(p => p.lineas.forEach(l => {
        for (const k of Object.keys(SUBS)) {
          const s = l.sub[k]; if (!s || !s.val) continue;
          regs.push({ sub: k, entNit: nitDe(s.cod), entNombre: nombreDe(s.cod) || s.cod, nombre: l.nombreCompleto, cedula: l.cedula, tipo: 'plano', valor: s.val });
        }
      }));
    }
    const ignoradosPar = regs.filter(r => r.sub === 'par');
    if (ignoradosPar.length) avisos.push(`${ignoradosPar.length} registros de SENA/ICBF no se comparan (aportante exonerado o no aplica).`);
    regs = regs.filter(r => r.sub !== 'par');

    // ---- Resolver cédula por nombre (Siigo no trae cédula)
    const candidatos = [];
    planos.forEach(p => p.lineas.forEach(l => candidatos.push({ cedula: l.cedula, toks: tokens(l.nombreCompleto), fuente: 'plano' })));
    for (const w of P.values()) candidatos.push({ cedula: w.cedula, toks: tokens(w.nombre), fuente: 'planilla' });
    const cacheNombre = new Map();
    const sinIdentificar = new Set();
    function resolver(nombre) {
      if (cacheNombre.has(nombre)) return cacheNombre.get(nombre);
      const T = new Set(tokens(nombre));
      let best = null, bestScore = 0, empate = false;
      for (const c of candidatos) {
        if (c.toks.length < 2) continue;
        if (!c.toks.every(t => T.has(t))) continue;
        const score = c.toks.length * 10 + (c.fuente === 'plano' ? 5 : 0) + (c.toks.length === T.size ? 3 : 0);
        if (score > bestScore) { best = c; bestScore = score; empate = false; }
        else if (score === bestScore && best && best.cedula !== c.cedula) empate = true;
      }
      const r = best && !empate ? best.cedula : null;
      cacheNombre.set(nombre, r);
      return r;
    }
    regs.forEach(r => {
      if (!r.cedula) {
        r.cedula = resolver(r.nombre);
        if (!r.cedula) { r.cedula = 'NOM:' + norm(r.nombre); sinIdentificar.add(r.nombre); }
      }
    });

    // ---- Nómina por trabajador
    const N = new Map();
    for (const r of regs) {
      let w = N.get(r.cedula);
      if (!w) { w = { cedula: r.cedula, nombre: r.nombre, sub: {} }; N.set(r.cedula, w); }
      if (r.nombre && r.nombre.length > (w.nombre || '').length) w.nombre = r.nombre;
      const s = w.sub[r.sub] || (w.sub[r.sub] = { total: 0, byNit: {}, nombres: {}, empleado: 0, patronal: 0 });
      s.total += r.valor;
      s.byNit[r.entNit] = (s.byNit[r.entNit] || 0) + r.valor;
      s.nombres[r.entNit] = r.entNombre;
      if (r.tipo === 'empleado') s.empleado += r.valor; else if (r.tipo === 'patronal') s.patronal += r.valor;
    }

    // Plano por cédula (para avisos de código EPS)
    const planoPor = new Map();
    planos.forEach(p => p.lineas.forEach(l => { const a = planoPor.get(l.cedula) || []; a.push(l); planoPor.set(l.cedula, a); }));

    // ---- Comparación por trabajador
    const trabajadores = [];
    const todos = new Set([...N.keys(), ...P.keys()]);
    for (const id of todos) {
      const n = N.get(id), p = P.get(id);
      const hall = [];
      const nombre = (n && n.nombre) || (p && p.nombre) || id;
      if (n && !p) {
        hall.push({
          sev: 'rojo', sub: '', tipo: id.startsWith('NOM:') ? 'Trabajador no identificado' : 'No está en la planilla',
          detalle: id.startsWith('NOM:')
            ? `"${n.nombre}" está en contabilidad pero no se encontró en la planilla por nombre.`
            : `Tiene aportes en contabilidad por ${money(sumTotales(n))} pero no aparece en las planillas cargadas.`,
          resp: 'Talento Humano', accion: 'Si se paga en otra planilla (p. ej. Profesionales), cárguela también; si no, incluirlo en la planilla o revisar la nómina.',
        });
      }
      if (p && !n) {
        hall.push({
          sev: 'rojo', sub: '', tipo: 'No está en contabilidad',
          detalle: `Aparece en la planilla ${[...p.planillas].join(', ')} por ${money(sumTotales(p))} pero no tiene registros en el auxiliar contable.`,
          resp: 'Contabilidad', accion: 'Verificar que la nómina del trabajador esté contabilizada en el período.',
        });
      }
      // Entidad configurada como el mismo trabajador
      if (n) {
        const malas = Object.keys(n.sub).filter(k => Object.keys(n.sub[k].byNit).some(nit => nit && nit === id));
        if (malas.length) {
          hall.push({
            sev: 'rojo', sub: '', tipo: 'Entidad mal configurada',
            detalle: `En contabilidad los aportes de ${malas.map(k => SUBS[k]).join(', ')} están a nombre del propio trabajador (NIT ${id}), no de la entidad.`,
            resp: 'Contabilidad', accion: 'Corregir las entidades en la ficha del empleado en el software contable y reclasificar los saldos.',
          });
        }
      }
      if (n && p) {
        for (const k of Object.keys(SUBS)) {
          const ns = n.sub[k], ps = p.sub[k];
          if (!ns && !ps) continue;
          const vN = ns ? ns.total : 0, vP = ps ? ps.total : 0;
          // Comparación de entidad
          if (ns && ps && Object.keys(ps.byNit).length) {
            const nitsN = Object.keys(ns.byNit).filter(x => Math.abs(ns.byNit[x]) >= 1 && x !== id);
            const nitsP = Object.keys(ps.byNit);
            const nomN = nitsN.map(x => ns.nombres[x] || x).join(', ');
            const nomP = nitsP.map(x => x.startsWith('COD:') ? (ps.nomEnt || x.slice(4)) : (nombreDe([...ps.cods].find(c => nitDe(c) === x)) || x)).join(', ');
            const codsP = [...ps.cods].join(', ');
            let distinto = false;
            for (const x of nitsP) {
              if (x.startsWith('COD:')) {
                const wordsP = entityWords(ps.nomEnt || '');
                const ok = wordsP.length && nitsN.some(y => entityWords(ns.nombres[y] || '').some(t => wordsP.includes(t)));
                if (!ok) distinto = true;
              } else if (!nitsN.includes(x)) distinto = true;
            }
            if (nitsN.length && distinto) {
              hall.push({
                sev: 'rojo', sub: k, tipo: 'Entidad distinta',
                detalle: `Contabilidad: ${nomN}. Planilla: ${nomP} (código ${codsP}).`,
                resp: 'Talento Humano', accion: `Actualizar la ${SUBS[k]} del trabajador en la ficha de nómina; Contabilidad reclasifica el valor a la entidad correcta.`,
                nomina: vN, planilla: vP,
              });
            }
          }
          // Aviso de código distinto con la misma entidad (requiere plano)
          const pl = planoPor.get(id);
          if (pl && ps && k !== 'arl') {
            const codPlano = [...new Set(pl.map(l => l.sub[k] && l.sub[k].cod).filter(Boolean))];
            const codPla = [...ps.cods];
            const difCod = codPlano.filter(c => !codPla.includes(c));
            if (difCod.length && codPla.length && difCod.every(c => codPla.some(cp => nitDe(cp) && nitDe(cp) === nitDe(c)))) {
              hall.push({
                sev: 'amarillo', sub: k, tipo: 'Código de entidad distinto (misma entidad)',
                detalle: `Nómina/plano usa ${difCod.join(', ')} y la planilla ${codPla.join(', ')} (${nombreDe(codPla[0])}).`,
                resp: 'Talento Humano', accion: 'Actualizar el código de la entidad en la ficha del trabajador.',
              });
            }
          }
          // Comparación de valor
          const dif = vP - vN;
          if (Math.abs(dif) > opt.tolTrabajador) {
            hall.push({
              sev: 'rojo', sub: k, tipo: !ns ? `Sin ${SUBS[k]} en contabilidad` : !ps ? `Sin ${SUBS[k]} en planilla` : 'Valor distinto',
              detalle: causa(k, ns, ps, p, dif), resp: 'Talento Humano',
              accion: k === 'arl' ? 'Revisar centro de trabajo, clase de riesgo y días de ARL.' : 'Revisar días, IBC y novedades del trabajador.',
              nomina: vN, planilla: vP, dif,
            });
          } else if (Math.abs(dif) >= 1 && ns && ps) {
            hall.push({ sev: 'ok', sub: k, tipo: 'Redondeo', detalle: `Diferencia de ${money(dif)} por redondeo.`, nomina: vN, planilla: vP, dif });
          }
        }
      }
      (hallCentro.get(id) || []).forEach(h => hall.push(h));
      const sev = hall.some(h => h.sev === 'rojo') ? 'rojo' : hall.some(h => h.sev === 'amarillo') ? 'amarillo' : 'ok';
      trabajadores.push({
        cedula: id.startsWith('NOM:') ? '' : id, nombre: id.startsWith('NOM:') ? n.nombre : nombre, sev, hallazgos: hall,
        centros: p ? [...p.centros].filter(Boolean) : [], planillas: p ? [...p.planillas] : [],
        novedades: p ? [...p.nov] : [],
        nomina: n ? mapSubs(n.sub, s => s.total) : {}, planilla: p ? mapSubs(p.sub, s => s.total) : {},
        dias: p ? mapSubs(p.sub, s => Math.min(s.dias, 30)) : {},
        entNomina: n ? mapSubs(n.sub, s => Object.keys(s.nombres).map(x => s.nombres[x]).join(', ')) : {},
        entPlanilla: p ? mapSubs(p.sub, s => [...s.cods].map(c => nombreDe(c) || s.nomEnt || c).join(', ')) : {},
      });
    }
    const ordenSev = { rojo: 0, amarillo: 1, ok: 2 };
    trabajadores.sort((a, b) => ordenSev[a.sev] - ordenSev[b.sev] || a.nombre.localeCompare(b.nombre));

    // ---- Resumen por entidad
    const entMap = new Map();
    const addEnt = (k, nit, nombre, campo, v) => {
      const key = k + '|' + nit;
      const e = entMap.get(key) || { sub: k, nit, nombre: '', nomina: 0, planilla: 0 };
      if (nombre && !e.nombre) e.nombre = nombre;
      e[campo] += v; entMap.set(key, e);
    };
    for (const w of N.values()) for (const k of Object.keys(w.sub)) for (const nit of Object.keys(w.sub[k].byNit)) addEnt(k, nit, w.sub[k].nombres[nit], 'nomina', w.sub[k].byNit[nit]);
    for (const l of lineasPlanilla) for (const k of Object.keys(SUBS)) {
      const s = l.sub[k]; if (!s || !s.cod) continue;
      const nit = nitDe(s.cod) || ('COD:' + s.cod);
      addEnt(k, nit, nombreDe(s.cod) || s.nomEnt || s.cod, 'planilla', s.val);
    }
    const entidades = [...entMap.values()].map(e => {
      const dif = e.planilla - e.nomina;
      return Object.assign(e, { dif, estado: Math.abs(dif) <= opt.tolEntidad ? 'ok' : 'rojo' });
    }).sort((a, b) => a.sub.localeCompare(b.sub) || b.planilla - a.planilla);

    const totales = {};
    for (const k of Object.keys(SUBS)) {
      const en = entidades.filter(e => e.sub === k);
      totales[k] = { nomina: sum(en.map(e => e.nomina)), planilla: sum(en.map(e => e.planilla)) };
    }

    return {
      empresa: empresaNombre, nitEmpresa, fuenteNomina,
      periodo: planillas.map(p => p.periodo).filter(Boolean).join(', '),
      planillas: planillas.map(p => ({ numero: p.numero, fuente: p.fuente, archivo: p.archivo, periodo: p.periodo, lineas: p.lineas.length, pagada: p.pagada })),
      trabajadores, entidades, totales, avisos,
      centrosNuevos: [...centrosNuevos.values()].map(c => Object.assign(c, { trabajadores: c.trabajadores.size })),
      sinIdentificar: [...sinIdentificar],
      conteo: {
        total: trabajadores.length, rojo: trabajadores.filter(t => t.sev === 'rojo').length,
        amarillo: trabajadores.filter(t => t.sev === 'amarillo').length, ok: trabajadores.filter(t => t.sev === 'ok').length,
      },
      opciones: opt,
    };
  }

  // ------------------------------------------------------------------ causas probables
  function claseDe(tarifa) {
    if (!tarifa) return null;
    let best = null, err = Infinity;
    ARL_CLASES.forEach((t, i) => { const e = Math.abs(tarifa - t) / t; if (e < err) { err = e; best = i + 1; } });
    return err < 0.08 ? best : null;
  }
  function causa(k, ns, ps, p, dif) {
    const partes = [];
    if (!ns) return `La planilla liquida ${money(ps.total)} de ${SUBS[k]} y en contabilidad no hay registro.`;
    if (!ps) return `Contabilidad registra ${money(ns.total)} de ${SUBS[k]} y la planilla no liquida este subsistema.`;
    partes.push(`Contabilidad ${money(ns.total)} vs planilla ${money(ps.total)} (dif. ${money(dif)}).`);
    if (k === 'arl') {
      const clP = [...ps.tarifas].map(claseDe).filter(Boolean);
      const ibc = ps.ibcTarifa || ps.ibc;
      const clN = ibc ? claseDe(ns.total / ibc) : null;
      if (clN && clP.length && !clP.includes(clN)) {
        partes.push(`Contabilidad liquida ARL como riesgo ${clN} (${pct(ARL_CLASES[clN - 1])}) y la planilla como riesgo ${clP.join('/')} (${[...ps.tarifas].map(pct).join('/')}).`);
      }
      const sinTarifa = p.lineas.filter(l => l.sub.arl && l.sub.arl.dias > 0 && !l.sub.arl.tarifa && !(l.nov.includes('IGE') || l.nov.includes('LMA') || l.nov.includes('SLN') || l.nov.includes('VAC')));
      if (sinTarifa.length) partes.push(`Hay ${sinTarifa.reduce((a, l) => a + l.sub.arl.dias, 0)} días con ARL en $0.`);
    } else if (TARIFA_REF[k] && ps.ibc) {
      const tasa = ns.total / ps.ibc;
      const ibcN = ns.total / (ps.total / ps.ibc || TARIFA_REF[k]);
      if (Math.abs(ibcN - ps.ibc) > 1000) partes.push(`El IBC implícito en contabilidad es ≈ ${money(ibcN)} y en planilla ${money(ps.ibc)}.`);
      void tasa;
    }
    if (ps.dias && ps.dias < 30) partes.push(`Planilla con ${Math.min(ps.dias, 30)} días.`);
    if (p.nov.size) partes.push(`Novedades en planilla: ${[...p.nov].join(', ')}.`);
    return partes.join(' ');
  }

  // ------------------------------------------------------------------ helpers de formato
  function sum(a) { return a.reduce((x, y) => x + y, 0); }
  function sumTotales(w) { return sum(Object.values(w.sub).map(s => s.total)); }
  function mapSubs(sub, fn) { const o = {}; Object.keys(sub).forEach(k => { o[k] = fn(sub[k]); }); return o; }
  function money(v) { return (v < 0 ? '-$' : '$') + Math.abs(Math.round(v)).toLocaleString('es-CO'); }
  function pct(t) { return (t * 100).toFixed(3).replace(/0+$/, '').replace(/\.$/, '') + '%'; }

  // ------------------------------------------------------------------ informe Excel
  function informeLibro(XLSX, res) {
    const wb = XLSX.utils.book_new();
    const r = res;
    const resumen = [
      ['Conciliación de aportes: nómina vs planilla'],
      ['Empresa', r.empresa], ['NIT', r.nitEmpresa], ['Período', r.periodo], ['Fuente nómina', r.fuenteNomina],
      ['Planillas', r.planillas.map(p => `${p.fuente} ${p.numero}`).join(', ')],
      ['Generado', new Date().toLocaleString('es-CO')], [],
      ['Trabajadores', r.conteo.total], ['Con diferencias', r.conteo.rojo], ['Con avisos', r.conteo.amarillo], ['OK', r.conteo.ok], [],
      ['Subsistema', 'Contabilidad', 'Planilla', 'Diferencia'],
      ...Object.keys(SUBS).map(k => [SUBS[k], round(r.totales[k].nomina), round(r.totales[k].planilla), round(r.totales[k].planilla - r.totales[k].nomina)]),
      [], ['Avisos'], ...r.avisos.map(a => [a]),
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumen), 'Resumen');
    const dif = [['Estado', 'Cédula', 'Trabajador', 'Centro de trabajo', 'Subsistema', 'Tipo', 'Detalle', 'Contabilidad', 'Planilla', 'Diferencia', 'Responsable', 'Acción sugerida', 'Corregido (S/N)']];
    r.trabajadores.forEach(t => t.hallazgos.filter(h => h.sev !== 'ok').forEach(h => dif.push([
      h.sev === 'rojo' ? 'Diferencia' : 'Aviso', t.cedula, t.nombre, t.centros.join(', '), SUBS[h.sub] || '', h.tipo, h.detalle,
      h.nomina !== undefined ? round(h.nomina) : '', h.planilla !== undefined ? round(h.planilla) : '', h.dif !== undefined ? round(h.dif) : '',
      h.resp || '', h.accion || '', '',
    ])));
    const wsd = XLSX.utils.aoa_to_sheet(dif);
    wsd['!cols'] = [10, 12, 32, 22, 14, 26, 70, 13, 13, 12, 16, 60, 12].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, wsd, 'Diferencias');
    const ent = [['Subsistema', 'NIT', 'Entidad', 'Contabilidad', 'Planilla', 'Diferencia', 'Estado']];
    r.entidades.forEach(e => ent.push([SUBS[e.sub], e.nit, e.nombre, round(e.nomina), round(e.planilla), round(e.dif), e.estado === 'ok' ? 'OK' : 'Revisar']));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ent), 'Por entidad');
    const all = [['Estado', 'Cédula', 'Trabajador', 'Centro', 'Planillas', 'Novedades',
      'Salud contab.', 'Salud planilla', 'Pensión contab.', 'Pensión planilla', 'ARL contab.', 'ARL planilla', 'CCF contab.', 'CCF planilla',
      'EPS contab.', 'EPS planilla', 'AFP contab.', 'AFP planilla']];
    r.trabajadores.forEach(t => all.push([
      t.sev === 'rojo' ? 'Diferencia' : t.sev === 'amarillo' ? 'Aviso' : 'OK', t.cedula, t.nombre, t.centros.join(', '), t.planillas.join(', '), t.novedades.join(', '),
      ...['sal', 'pen', 'arl', 'ccf'].flatMap(k => [round(t.nomina[k] || 0), round(t.planilla[k] || 0)]),
      t.entNomina.sal || '', t.entPlanilla.sal || '', t.entNomina.pen || '', t.entPlanilla.pen || '',
    ]));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(all), 'Todos');
    if (r.centrosNuevos.length) {
      const cn = [['Empresa', 'Nombre en planilla', 'Código centro', 'Riesgo observado', 'Trabajadores']];
      r.centrosNuevos.forEach(c => cn.push([c.empresa, c.nombrePlanilla, c.codigo, c.riesgoObservado || '', c.trabajadores]));
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cn), 'Centros nuevos');
    }
    return wb;
  }

  return { leerArchivo, conciliar, informeLibro, SUBS, money, pct, norm, claseDe, ARL_CLASES };
});
