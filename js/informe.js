/*
 * Informe Excel con formato (ExcelJS). Recibe el resultado de Concilia.conciliar().
 * Navegador: window.Informe.generar(ExcelJS, res, Concilia)  ·  Node: require('./informe')
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Informe = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const COL = {
    brand: 'FF1F4E78', brandLight: 'FFDCE6F1', ink: 'FF1B2430', muted: 'FF5F6B7A', white: 'FFFFFFFF',
    grey: 'FFF2F4F7', band: 'FFF8FAFC', line: 'FFD0D5DD',
    red: 'FFC0392B', redBg: 'FFFDECEA', amber: 'FF9A6700', amberBg: 'FFFFF4D6', green: 'FF1E7B4B', greenBg: 'FFE6F4EC',
  };
  const MONEY = '"$"#,##0;[Red]-"$"#,##0;"–"';
  const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
  const thin = { style: 'thin', color: { argb: COL.line } };
  const BORDER = { top: thin, left: thin, bottom: thin, right: thin };
  const SEV = {
    rojo: { txt: 'Diferencia', color: COL.red, bg: COL.redBg },
    amarillo: { txt: 'Aviso', color: COL.amber, bg: COL.amberBg },
    ok: { txt: 'OK', color: COL.green, bg: COL.greenBg },
  };
  const ESTADOS_CORR = ['Pendiente', 'Corregido', 'No aplica'];

  function titulo(ws, texto, sub, ancho) {
    ws.mergeCells(1, 1, 1, ancho);
    const t = ws.getCell(1, 1);
    t.value = texto; t.font = { bold: true, size: 15, color: { argb: COL.white } };
    t.fill = fill(COL.brand); t.alignment = { vertical: 'middle', indent: 1 };
    ws.getRow(1).height = 30;
    if (sub) {
      ws.mergeCells(2, 1, 2, ancho);
      const s = ws.getCell(2, 1);
      s.value = sub; s.font = { italic: true, size: 10, color: { argb: COL.muted } };
      s.alignment = { vertical: 'middle', indent: 1, wrapText: true };
      ws.getRow(2).height = 30;
    }
  }
  function encabezado(row, opts) {
    row.eachCell({ includeEmpty: true }, c => {
      c.font = { bold: true, color: { argb: COL.white }, size: 10 };
      c.fill = fill((opts && opts.color) || COL.brand);
      c.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      c.border = BORDER;
    });
    row.height = (opts && opts.height) || 30;
  }
  function celdaSev(c, sev) {
    const s = SEV[sev];
    c.value = s.txt; c.font = { bold: true, color: { argb: s.color } }; c.fill = fill(s.bg);
    c.alignment = { horizontal: 'center', vertical: 'top' };
  }
  function bordes(row, desde, hasta) {
    for (let i = desde; i <= hasta; i++) row.getCell(i).border = BORDER;
  }
  const r0 = v => Math.round(v || 0);

  function generar(ExcelJS, res, C) {
    const SUBS = C.SUBS;
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Conciliación de aportes';
    wb.created = new Date();
    const generado = new Date().toLocaleString('es-CO');
    const filasDif = [];
    res.trabajadores.forEach(t => t.hallazgos.filter(h => h.sev !== 'ok').forEach(h => filasDif.push({ t, h })));

    // ================================================================ RESUMEN
    const ws = wb.addWorksheet('Resumen', { properties: { tabColor: { argb: COL.brand } }, views: [{ showGridLines: false }] });
    ws.columns = [16, 16, 16, 16, 16, 16, 16, 16].map(width => ({ width }));
    titulo(ws, 'Conciliación de aportes · Nómina vs planilla PILA',
      `Generado: ${generado} · Las diferencias se calculan como planilla − contabilidad. Tolerancias: $${res.opciones.tolTrabajador.toLocaleString('es-CO')} por trabajador y $${res.opciones.tolEntidad.toLocaleString('es-CO')} por entidad.`, 8);

    let r = 4;
    [['Empresa', res.empresa], ['NIT', res.nitEmpresa], ['Período', res.periodo], ['Contabilidad', res.fuenteNomina],
      ['Planillas', res.planillas.map(p => `${p.fuente} N.º ${p.numero}${p.pagada === 'Sí' ? ' (pagada)' : ''}`).join(' · ')]].forEach(([k, v]) => {
      const a = ws.getCell(r, 1);
      a.value = k; a.font = { bold: true, color: { argb: COL.muted } }; a.fill = fill(COL.grey);
      a.alignment = { vertical: 'middle', indent: 1 };
      ws.mergeCells(r, 2, r, 8);
      const b = ws.getCell(r, 2);
      b.value = v; b.font = { bold: k === 'Empresa', size: k === 'Empresa' ? 12 : 11 }; b.alignment = { vertical: 'middle', indent: 1 };
      ws.getRow(r).height = 20;
      r++;
    });

    // Tarjetas
    r++;
    const kpis = [
      ['Trabajadores revisados', res.conteo.total, COL.brand, COL.brandLight],
      ['Con diferencias', res.conteo.rojo, COL.red, COL.redBg],
      ['Con avisos', res.conteo.amarillo, COL.amber, COL.amberBg],
      ['Sin diferencias', res.conteo.ok, COL.green, COL.greenBg],
    ];
    kpis.forEach(([lbl, val, color, bg], i) => {
      const c1 = 1 + i * 2;
      ws.mergeCells(r, c1, r, c1 + 1); ws.mergeCells(r + 1, c1, r + 1, c1 + 1);
      const v = ws.getCell(r, c1), l = ws.getCell(r + 1, c1);
      v.value = val; v.font = { bold: true, size: 22, color: { argb: color } }; v.fill = fill(bg);
      v.alignment = { horizontal: 'center', vertical: 'middle' };
      l.value = lbl; l.font = { size: 10, color: { argb: color } }; l.fill = fill(bg);
      l.alignment = { horizontal: 'center', vertical: 'top' };
    });
    ws.getRow(r).height = 36; ws.getRow(r + 1).height = 18;
    r += 3;

    // Avance de corrección (fórmulas vivas sobre la hoja Diferencias)
    if (filasDif.length) {
      const rng = `Diferencias!$N$4:$N$${3 + filasDif.length}`;
      ws.mergeCells(r, 1, r, 3);
      const a = ws.getCell(r, 1);
      a.value = 'Avance de corrección'; a.font = { bold: true, color: { argb: COL.muted } }; a.fill = fill(COL.grey);
      a.alignment = { vertical: 'middle', indent: 1 };
      const items = [
        ['Pendientes', { formula: `COUNTIF(${rng},"Pendiente")` }, COL.red],
        ['Corregidos', { formula: `COUNTIF(${rng},"Corregido")` }, COL.green],
        ['% avance', { formula: `IFERROR((COUNTIF(${rng},"Corregido")+COUNTIF(${rng},"No aplica"))/COUNTA(${rng}),0)` }, COL.brand],
      ];
      ws.getCell(r, 4).value = 'Pendientes'; ws.getCell(r, 5).value = items[0][1];
      ws.getCell(r, 6).value = 'Corregidos'; ws.getCell(r, 7).value = items[1][1];
      ws.getCell(r + 1, 4).value = '% avance'; ws.getCell(r + 1, 5).value = items[2][1];
      ws.getCell(r + 1, 5).numFmt = '0%';
      [[r, 4], [r, 6], [r + 1, 4]].forEach(([y, x]) => { const c = ws.getCell(y, x); c.font = { color: { argb: COL.muted } }; c.alignment = { horizontal: 'right' }; });
      [[r, 5, COL.red], [r, 7, COL.green], [r + 1, 5, COL.brand]].forEach(([y, x, col]) => { const c = ws.getCell(y, x); c.font = { bold: true, size: 12, color: { argb: col } }; c.alignment = { horizontal: 'center' }; });
      ws.mergeCells(r + 1, 1, r + 1, 3);
      const n = ws.getCell(r + 1, 1);
      n.value = 'Se actualiza al cambiar la columna "Estado corrección" en la hoja Diferencias.';
      n.font = { italic: true, size: 9, color: { argb: COL.muted } }; n.alignment = { indent: 1 };
      r += 3;
    }

    // Totales por subsistema
    const sec = (texto) => {
      ws.mergeCells(r, 1, r, 8);
      const c = ws.getCell(r, 1);
      c.value = texto; c.font = { bold: true, size: 12, color: { argb: COL.brand } };
      c.border = { bottom: { style: 'medium', color: { argb: COL.brand } } };
      ws.getRow(r).height = 22; r++;
    };
    sec('Totales por subsistema');
    const hdr = ws.getRow(r);
    ['Subsistema', '', 'Contabilidad', 'Planilla', 'Diferencia', 'Estado'].forEach((v, i) => { hdr.getCell(i + 1).value = v; });
    ws.mergeCells(r, 1, r, 2);
    for (let i = 1; i <= 6; i++) { const c = hdr.getCell(i); c.font = { bold: true, color: { argb: COL.white }, size: 10 }; c.fill = fill(COL.brand); c.alignment = { horizontal: 'center', vertical: 'middle' }; c.border = BORDER; }
    hdr.height = 20; r++;
    let tN = 0, tP = 0;
    Object.keys(SUBS).forEach(k => {
      const t = res.totales[k], dif = t.planilla - t.nomina;
      tN += t.nomina; tP += t.planilla;
      const row = ws.getRow(r);
      ws.mergeCells(r, 1, r, 2);
      row.getCell(1).value = SUBS[k];
      row.getCell(3).value = r0(t.nomina); row.getCell(4).value = r0(t.planilla); row.getCell(5).value = r0(dif);
      [3, 4, 5].forEach(i => { row.getCell(i).numFmt = MONEY; });
      celdaSev(row.getCell(6), Math.abs(dif) <= res.opciones.tolEntidad ? 'ok' : 'rojo');
      if (row.getCell(6).value === 'Diferencia') row.getCell(6).value = 'Revisar';
      bordes(row, 1, 6); r++;
    });
    const tot = ws.getRow(r);
    ws.mergeCells(r, 1, r, 2);
    tot.getCell(1).value = 'TOTAL'; tot.getCell(3).value = r0(tN); tot.getCell(4).value = r0(tP); tot.getCell(5).value = r0(tP - tN);
    for (let i = 1; i <= 6; i++) { const c = tot.getCell(i); c.font = { bold: true }; c.fill = fill(COL.grey); c.border = BORDER; if (i >= 3 && i <= 5) c.numFmt = MONEY; }
    r += 2;

    // Hallazgos por tipo
    if (filasDif.length) {
      sec('Hallazgos por tipo');
      const tipos = new Map();
      filasDif.forEach(({ h }) => {
        const key = h.tipo + '|' + h.sev;
        const o = tipos.get(key) || { tipo: h.tipo, sev: h.sev, n: 0, resp: new Set() };
        o.n++; if (h.resp) o.resp.add(h.resp); tipos.set(key, o);
      });
      const hh = ws.getRow(r);
      ['Tipo de hallazgo', '', '', 'Nivel', 'Cantidad', 'Responsable', ''].forEach((v, i) => { hh.getCell(i + 1).value = v; });
      ws.mergeCells(r, 1, r, 3); ws.mergeCells(r, 6, r, 7);
      for (let i = 1; i <= 7; i++) { const c = hh.getCell(i); c.font = { bold: true, color: { argb: COL.white }, size: 10 }; c.fill = fill(COL.brand); c.alignment = { horizontal: 'center', vertical: 'middle' }; c.border = BORDER; }
      hh.height = 20; r++;
      [...tipos.values()].sort((a, b) => (a.sev === b.sev ? b.n - a.n : a.sev === 'rojo' ? -1 : 1)).forEach(o => {
        const row = ws.getRow(r);
        ws.mergeCells(r, 1, r, 3); ws.mergeCells(r, 6, r, 7);
        row.getCell(1).value = o.tipo; row.getCell(1).alignment = { indent: 1 };
        celdaSev(row.getCell(4), o.sev);
        row.getCell(5).value = o.n; row.getCell(5).alignment = { horizontal: 'center' };
        row.getCell(6).value = [...o.resp].join(', ');
        bordes(row, 1, 7); r++;
      });
      r++;
    }

    // Avisos
    const avisos = res.avisos.concat(res.centrosNuevos.length ? [`${res.centrosNuevos.length} centro(s) de trabajo no están en el catálogo (ver hoja "Centros nuevos").`] : []);
    if (avisos.length) {
      sec('Avisos');
      avisos.forEach(a => {
        ws.mergeCells(r, 1, r, 8);
        const c = ws.getCell(r, 1);
        c.value = '⚠ ' + a; c.fill = fill(COL.amberBg); c.font = { color: { argb: COL.amber } };
        c.alignment = { wrapText: true, vertical: 'middle', indent: 1 };
        ws.getRow(r).height = 30; r++;
      });
      r++;
    }
    sec('Cómo usar este informe');
    [
      '1. En la hoja "Diferencias" está la lista de trabajo: un renglón por hallazgo, con su responsable y la acción sugerida.',
      '2. Talento Humano corrige en el operador o en la ficha del trabajador y cambia "Estado corrección" a Corregido (o No aplica, explicando en Observaciones).',
      '3. "Por entidad" es la conciliación clásica por NIT; "Detalle trabajadores" muestra todos los valores de cada persona.',
      '4. Guarde este archivo en la carpeta del mes en Teams (Aportes / AAAA-MM / Empresa) como soporte de la conciliación.',
    ].forEach(t => {
      ws.mergeCells(r, 1, r, 8);
      const c = ws.getCell(r, 1); c.value = t; c.alignment = { wrapText: true, indent: 1, vertical: 'middle' };
      ws.getRow(r).height = 28; r++;
    });
    ws.pageSetup = { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 } };

    // ================================================================ DIFERENCIAS
    const wd = wb.addWorksheet('Diferencias', { properties: { tabColor: { argb: COL.red } }, views: [{ state: 'frozen', ySplit: 3, xSplit: 4 }] });
    const colsD = [['N.º', 5], ['Nivel', 11], ['Cédula', 13], ['Trabajador', 30], ['Centro de trabajo', 18], ['Subsistema', 13],
      ['Tipo de hallazgo', 22], ['Detalle', 55], ['Contabilidad', 13], ['Planilla', 13], ['Diferencia', 13],
      ['Responsable', 15], ['Acción sugerida', 42], ['Estado corrección', 15], ['Observaciones', 30]];
    wd.columns = colsD.map(([, width]) => ({ width }));
    titulo(wd, `Diferencias por trabajador · ${res.empresa} · ${res.periodo}`,
      'Un renglón por hallazgo. Filtre por Nivel, Responsable o Estado corrección. Cambie "Estado corrección" a medida que se corrige.', colsD.length);
    const hD = wd.getRow(3);
    colsD.forEach(([h], i) => { hD.getCell(i + 1).value = h; });
    encabezado(hD);
    let rr = 4, nTrab = 0, prev = null;
    filasDif.forEach(({ t, h }) => {
      if (t !== prev) { nTrab++; prev = t; }
      const row = wd.getRow(rr);
      const vals = [nTrab, '', t.cedula, t.nombre, t.centros.join(', '), SUBS[h.sub] || 'General', h.tipo, h.detalle,
        h.nomina !== undefined ? r0(h.nomina) : null, h.planilla !== undefined ? r0(h.planilla) : null, h.dif !== undefined ? r0(h.dif) : null,
        h.resp || '', h.accion || '', 'Pendiente', ''];
      vals.forEach((v, i) => { row.getCell(i + 1).value = v; });
      const band = nTrab % 2 === 0 ? COL.band : COL.white;
      for (let i = 1; i <= colsD.length; i++) {
        const c = row.getCell(i);
        c.fill = fill(band); c.border = BORDER;
        c.alignment = { vertical: 'top', wrapText: [4, 5, 6, 7, 8, 13, 15].includes(i) };
      }
      celdaSev(row.getCell(2), h.sev);
      row.getCell(1).alignment = { horizontal: 'center', vertical: 'top' };
      row.getCell(4).font = { bold: true };
      [9, 10, 11].forEach(i => { row.getCell(i).numFmt = MONEY; });
      const ce = row.getCell(14);
      ce.dataValidation = { type: 'list', allowBlank: true, formulae: ['"' + ESTADOS_CORR.join(',') + '"'] };
      ce.alignment = { horizontal: 'center', vertical: 'top' }; ce.font = { bold: true };
      rr++;
    });
    if (!filasDif.length) {
      wd.mergeCells(4, 1, 4, colsD.length);
      const c = wd.getCell(4, 1); c.value = '✔ No hay diferencias por encima de la tolerancia.'; c.font = { bold: true, color: { argb: COL.green } };
    } else {
      wd.autoFilter = { from: { row: 3, column: 1 }, to: { row: rr - 1, column: colsD.length } };
      const ref = `N4:N${rr - 1}`;
      wd.addConditionalFormatting({
        ref, rules: [
          { type: 'expression', priority: 1, formulae: ['$N4="Corregido"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COL.greenBg } }, font: { color: { argb: COL.green }, bold: true } } },
          { type: 'expression', priority: 2, formulae: ['$N4="Pendiente"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COL.redBg } }, font: { color: { argb: COL.red }, bold: true } } },
          { type: 'expression', priority: 3, formulae: ['$N4="No aplica"'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COL.grey } }, font: { color: { argb: COL.muted }, bold: true } } },
        ],
      });
    }
    wd.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '3:3' };

    // ================================================================ POR ENTIDAD
    const we = wb.addWorksheet('Por entidad', { properties: { tabColor: { argb: COL.amber } }, views: [{ state: 'frozen', ySplit: 3, showGridLines: false }] });
    const colsE = [['Subsistema', 16], ['NIT', 13], ['Entidad', 38], ['Contabilidad', 15], ['Planilla', 15], ['Diferencia', 15], ['Estado', 11]];
    we.columns = colsE.map(([, width]) => ({ width }));
    titulo(we, `Conciliación por entidad · ${res.empresa} · ${res.periodo}`, `Diferencia = planilla − contabilidad. "Revisar" cuando supera $${res.opciones.tolEntidad.toLocaleString('es-CO')}.`, colsE.length);
    const hE = we.getRow(3); colsE.forEach(([h], i) => { hE.getCell(i + 1).value = h; }); encabezado(hE, { height: 22 });
    let re = 4;
    Object.keys(SUBS).forEach(k => {
      const ents = res.entidades.filter(e => e.sub === k);
      if (!ents.length) return;
      we.mergeCells(re, 1, re, colsE.length);
      const s = we.getCell(re, 1); s.value = SUBS[k].toUpperCase();
      s.font = { bold: true, color: { argb: COL.brand } }; s.fill = fill(COL.brandLight); s.alignment = { indent: 1 };
      re++;
      ents.forEach(e => {
        const row = we.getRow(re);
        [SUBS[k], e.nit.replace('COD:', 'cód. '), e.nombre, r0(e.nomina), r0(e.planilla), r0(e.dif)].forEach((v, i) => { row.getCell(i + 1).value = v; });
        [4, 5, 6].forEach(i => { row.getCell(i).numFmt = MONEY; });
        celdaSev(row.getCell(7), e.estado === 'ok' ? 'ok' : 'rojo');
        if (e.estado !== 'ok') row.getCell(7).value = 'Revisar';
        row.getCell(1).font = { color: { argb: COL.muted } };
        bordes(row, 1, 7); re++;
      });
      const t = res.totales[k], row = we.getRow(re);
      we.mergeCells(re, 1, re, 3);
      row.getCell(1).value = 'Total ' + SUBS[k];
      row.getCell(4).value = r0(t.nomina); row.getCell(5).value = r0(t.planilla); row.getCell(6).value = r0(t.planilla - t.nomina);
      for (let i = 1; i <= 7; i++) { const c = row.getCell(i); c.font = { bold: true }; c.fill = fill(COL.grey); c.border = BORDER; if (i >= 4 && i <= 6) c.numFmt = MONEY; }
      re += 2;
    });
    we.pageSetup = { orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

    // ================================================================ DETALLE TRABAJADORES
    const wt = wb.addWorksheet('Detalle trabajadores', { properties: { tabColor: { argb: COL.green } }, views: [{ state: 'frozen', ySplit: 4, xSplit: 3 }] });
    const subs = Object.keys(SUBS);
    const fijas = [['Nivel', 11], ['Cédula', 13], ['Trabajador', 30], ['Centro de trabajo', 18], ['Planillas', 13], ['Novedades', 14]];
    const fin = [['EPS contabilidad', 20], ['EPS planilla', 20], ['AFP contabilidad', 20], ['AFP planilla', 20]];
    const nCols = fijas.length + subs.length * 3 + fin.length;
    wt.columns = [...fijas.map(x => x[1]), ...subs.flatMap(() => [13, 13, 12]), ...fin.map(x => x[1])].map(width => ({ width }));
    titulo(wt, `Detalle por trabajador · ${res.empresa} · ${res.periodo}`, 'Valores de contabilidad y planilla por subsistema. Las diferencias que superan la tolerancia se resaltan en rojo.', nCols);
    const g = wt.getRow(3), h4 = wt.getRow(4);
    fijas.forEach(([h], i) => { g.getCell(i + 1).value = h; wt.mergeCells(3, i + 1, 4, i + 1); });
    subs.forEach((k, j) => {
      const c0 = fijas.length + 1 + j * 3;
      g.getCell(c0).value = SUBS[k]; wt.mergeCells(3, c0, 3, c0 + 2);
      h4.getCell(c0).value = 'Contabilidad'; h4.getCell(c0 + 1).value = 'Planilla'; h4.getCell(c0 + 2).value = 'Diferencia';
    });
    const cf = fijas.length + subs.length * 3 + 1;
    fin.forEach(([h], i) => { g.getCell(cf + i).value = h; wt.mergeCells(3, cf + i, 4, cf + i); });
    encabezado(g, { height: 20 }); encabezado(h4, { height: 20, color: 'FF2F6496' });
    let rt = 5;
    res.trabajadores.forEach((t, idx) => {
      const row = wt.getRow(rt);
      const vals = ['', t.cedula, t.nombre, t.centros.join(', '), t.planillas.join(', '), t.novedades.join(', ')];
      subs.forEach(k => { const n = r0(t.nomina[k]), p = r0(t.planilla[k]); vals.push(n, p, p - n); });
      vals.push(t.entNomina.sal || '', t.entPlanilla.sal || '', t.entNomina.pen || '', t.entPlanilla.pen || '');
      vals.forEach((v, i) => { row.getCell(i + 1).value = v; });
      const band = idx % 2 ? COL.band : COL.white;
      for (let i = 1; i <= nCols; i++) { const c = row.getCell(i); c.border = BORDER; c.fill = fill(band); c.alignment = { vertical: 'top', wrapText: i === 3 || i === 4 || i >= cf }; }
      celdaSev(row.getCell(1), t.sev);
      for (let i = fijas.length + 1; i < cf; i++) row.getCell(i).numFmt = MONEY;
      rt++;
    });
    if (rt > 5) {
      wt.autoFilter = { from: { row: 4, column: 1 }, to: { row: rt - 1, column: nCols } };
      subs.forEach((k, j) => {
        const col = colLetra(fijas.length + 3 + j * 3);
        wt.addConditionalFormatting({
          ref: `${col}5:${col}${rt - 1}`,
          rules: [{ type: 'expression', priority: 10 + j, formulae: [`ABS(${col}5)>${res.opciones.tolTrabajador}`], style: { font: { color: { argb: COL.red }, bold: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: COL.redBg } } } }],
        });
      });
    }
    wt.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '3:4' };

    // ================================================================ CENTROS NUEVOS
    if (res.centrosNuevos.length) {
      const wc = wb.addWorksheet('Centros nuevos', { properties: { tabColor: { argb: COL.amber } } });
      const colsC = [['Empresa', 40], ['Nombre en planilla', 26], ['Código centro', 14], ['Riesgo que liquida la planilla', 16], ['Trabajadores', 13]];
      wc.columns = colsC.map(([, width]) => ({ width }));
      titulo(wc, 'Centros de trabajo que no están en el catálogo', 'Agréguelos desde el aplicativo (pestaña Centros nuevos) y suba el catálogo actualizado a Teams.', colsC.length);
      const hc = wc.getRow(3); colsC.forEach(([h], i) => { hc.getCell(i + 1).value = h; }); encabezado(hc);
      res.centrosNuevos.forEach((c, i) => {
        const row = wc.getRow(4 + i);
        [c.empresa, c.nombrePlanilla, c.codigo, c.riesgoObservado ? 'Riesgo ' + c.riesgoObservado : '', c.trabajadores].forEach((v, j) => { row.getCell(j + 1).value = v; });
        bordes(row, 1, colsC.length);
      });
    }
    return wb;
  }

  function colLetra(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

  async function descargar(ExcelJS, res, C, nombre) {
    const wb = generar(ExcelJS, res, C);
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = nombre;
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  return { generar, descargar };
});
