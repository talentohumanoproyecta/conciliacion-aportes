/* Interfaz del aplicativo de conciliación. Requiere XLSX (SheetJS) y Concilia (engine.js). */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const C = window.Concilia;
  const SUBS = C.SUBS;
  const LS_CAT = 'conciliacion_catalogo_v1';
  const state = { catalogo: null, catInfo: null, archivos: [], res: null, tab: 'dif', filtro: '' };

  const esc = s => String(s === null || s === undefined ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const money = v => (v === undefined || v === null || v === '') ? '' : C.money(v);

  // ------------------------------------------------------------ catálogo
  function guardarCat() {
    try { localStorage.setItem(LS_CAT, JSON.stringify({ catalogo: state.catalogo, info: state.catInfo })); } catch (e) { /* sin almacenamiento */ }
  }
  function cargarCatGuardado() {
    try {
      const raw = localStorage.getItem(LS_CAT);
      if (raw) { const o = JSON.parse(raw); state.catalogo = o.catalogo; state.catInfo = o.info; }
    } catch (e) { /* sin almacenamiento */ }
  }
  function pintarCat() {
    const el = $('catEstado');
    if (!state.catalogo) {
      el.innerHTML = '<div class="msg warn">Sin catálogo: la conciliación funciona, pero no se validará la clase de riesgo por centro de trabajo.</div>';
      return;
    }
    const c = state.catalogo, i = state.catInfo || {};
    const sinRiesgo = c.centros.filter(x => !x.riesgo).length;
    el.innerHTML = `<b>${esc(i.archivo || 'Catálogo')}</b> · ${c.empresas.length} empresas · ${c.centros.length} centros` +
      (i.fecha ? ` · cargado ${esc(i.fecha)}` : '') + (i.editado ? ' · <span class="pill amarillo">editado en este equipo, falta subirlo a Teams</span>' : '') +
      (sinRiesgo ? `<div class="msg warn">${sinRiesgo} centros sin clase de riesgo confirmada: no se validan hasta completarlos.</div>` : '');
  }
  async function leerCat(file) {
    try {
      const buf = await file.arrayBuffer();
      const cat = C.leerArchivo(XLSX, new Uint8Array(buf), file.name);
      if (cat.tipo !== 'catalogo') throw new Error('Este archivo no es el catálogo (debe tener la hoja "Centros de trabajo").');
      state.catalogo = cat; state.catInfo = { archivo: file.name, fecha: new Date().toLocaleString('es-CO'), editado: false };
      guardarCat(); pintarCat(); pintarEditor();
      mensaje('info', 'Catálogo cargado.');
    } catch (e) { mensaje('err', e.message); }
  }

  // Editor de centros
  const RIESGOS = ['', '1', '2', '3', '4', '5'];
  function pintarEditor() {
    const t = $('tblCat');
    if (!state.catalogo) { t.innerHTML = '<tr><td class="empty">Cargue primero el catálogo.</td></tr>'; return; }
    const emps = state.catalogo.empresas.map(e => e.empresa);
    const opt = (arr, v) => arr.map(o => `<option${o === String(v || '') ? ' selected' : ''}>${esc(o)}</option>`).join('');
    t.innerHTML = '<thead><tr><th>Empresa</th><th>Centro (nombre interno)</th><th>Nombre en planilla</th><th>Código</th><th>Tipo</th><th>Riesgo</th><th></th></tr></thead><tbody>' +
      state.catalogo.centros.map((c, i) => `<tr data-i="${i}">
        <td><select data-f="empresa"><option></option>${opt(emps, c.empresa)}</select></td>
        <td><input type="text" data-f="interno" value="${esc(c.interno)}"></td>
        <td><input type="text" data-f="nombrePlanilla" value="${esc(c.nombrePlanilla)}"></td>
        <td><input type="text" data-f="codigo" value="${esc(c.codigo)}" style="width:70px"></td>
        <td><select data-f="tipo">${opt(['', 'Sede/Administración', 'Obra'], c.tipo)}</select></td>
        <td><select data-f="riesgo">${opt(RIESGOS, c.riesgo)}</select></td>
        <td><button class="x" data-del="${i}" title="Quitar">×</button></td></tr>`).join('') + '</tbody>';
  }
  $('tblCat').addEventListener('change', e => {
    const tr = e.target.closest('tr[data-i]'); if (!tr) return;
    const c = state.catalogo.centros[+tr.dataset.i], f = e.target.dataset.f;
    c[f] = f === 'riesgo' ? (parseInt(e.target.value, 10) || null) : e.target.value;
    marcarEditado();
  });
  $('tblCat').addEventListener('click', e => {
    const i = e.target.dataset.del; if (i === undefined) return;
    if (!confirm('¿Quitar este centro del catálogo?')) return;
    state.catalogo.centros.splice(+i, 1); marcarEditado(); pintarEditor();
  });
  function marcarEditado() { state.catInfo = Object.assign({}, state.catInfo, { editado: true }); guardarCat(); pintarCat(); }
  $('btnAddCentro').onclick = () => {
    if (!state.catalogo) return mensaje('warn', 'Cargue primero el catálogo.');
    state.catalogo.centros.push({ empresa: '', interno: '', nombrePlanilla: '', codigo: '', tipo: '', propia: '', riesgo: null, riesgoObservado: null, obs: 'Agregado desde el aplicativo' });
    marcarEditado(); pintarEditor();
    const rows = $('tblCat').querySelectorAll('tbody tr'); rows[rows.length - 1].querySelector('select').focus();
  };
  $('btnCatEditar').onclick = () => { $('catEditor').classList.toggle('hidden'); pintarEditor(); };
  $('btnCerrarCat').onclick = () => $('catEditor').classList.add('hidden');
  $('btnDescCat').onclick = descargarCatalogo;
  function descargarCatalogo() {
    if (!state.catalogo) return;
    const c = state.catalogo, wb = XLSX.utils.book_new();
    const tarifas = [0.00522, 0.01044, 0.02436, 0.0435, 0.0696];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['CATÁLOGO DE EMPRESAS Y CENTROS DE TRABAJO – Conciliación de aportes'],
      ['Generado desde el aplicativo el ' + new Date().toLocaleString('es-CO')],
      ['Para cambiar un riesgo o agregar un centro: edite la hoja "Centros de trabajo" (aquí o en el aplicativo) y guarde el archivo en la carpeta Catalogo de Teams.'],
      ['Riesgo 1 = 0,522% · 2 = 1,044% · 3 = 2,436% · 4 = 4,350% · 5 = 6,960%'],
    ]), 'Instrucciones');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Empresa', 'NIT', 'Software contable', 'Operador PILA'],
      ...c.empresas.map(e => [e.empresa, e.nit, e.software, e.operador])]), 'Empresas');
    const ws = XLSX.utils.aoa_to_sheet([['Empresa', 'Centro de trabajo (nombre interno)', 'Nombre en planilla (tal cual el operador)', 'Código centro en planilla', 'Tipo', 'Planilla propia', 'Tarifa ARL observada', 'Riesgo observado en archivos', 'Clase de riesgo CONFIRMADA (1-5)', 'Tarifa ARL esperada', 'Observaciones'],
      ...c.centros.map(x => [x.empresa, x.interno, x.nombrePlanilla, x.codigo, x.tipo, x.propia, '', x.riesgoObservado || '', x.riesgo || '', x.riesgo ? tarifas[x.riesgo - 1] : '', x.obs || ''])]);
    ws['!cols'] = [44, 24, 28, 12, 20, 12, 12, 12, 14, 12, 50].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, 'Centros de trabajo');
    const ents = Object.entries(c.entidades || {});
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Código PILA', 'NIT', 'Nombre'], ...ents.map(([k, v]) => [k, v[0], v[1]])]), 'Entidades');
    XLSX.writeFile(wb, 'Catalogo_Centros_Trabajo.xlsx');
    state.catInfo = Object.assign({}, state.catInfo, { editado: false }); guardarCat(); pintarCat();
    mensaje('info', 'Catálogo descargado. Reemplace el archivo en la carpeta Catalogo de Teams para que todos usen esta versión.');
  }
  $('btnCat').onclick = () => $('fileCat').click();
  $('fileCat').onchange = e => { if (e.target.files[0]) leerCat(e.target.files[0]); e.target.value = ''; };

  // ------------------------------------------------------------ archivos
  const drop = $('drop');
  drop.onclick = () => $('fileIn').click();
  drop.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('fileIn').click(); } };
  $('fileIn').onchange = e => { manejarArchivos([...e.target.files]); e.target.value = ''; };
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => manejarArchivos([...e.dataTransfer.files]));

  async function manejarArchivos(files) {
    for (const f of files) {
      const item = { name: f.name, data: null, error: null };
      try {
        const buf = await f.arrayBuffer();
        const d = C.leerArchivo(XLSX, new Uint8Array(buf), f.name);
        if (d.tipo === 'catalogo') {
          state.catalogo = d; state.catInfo = { archivo: f.name, fecha: new Date().toLocaleString('es-CO'), editado: false };
          guardarCat(); pintarCat(); pintarEditor(); continue;
        }
        item.data = d;
      } catch (e) { item.error = e.message; }
      state.archivos = state.archivos.filter(a => a.name !== f.name);
      state.archivos.push(item);
    }
    pintarLista();
  }
  const ETQ = { nomina: 'Contabilidad', planilla: 'Planilla', plano: 'Plano PILA' };
  function pintarLista() {
    $('lista').innerHTML = state.archivos.map((a, i) => {
      const d = a.data;
      const tag = a.error ? '<span class="tag err">No reconocido</span>' : `<span class="tag ${d.tipo}">${ETQ[d.tipo]} · ${esc(d.fuente)}</span>`;
      let det = '';
      if (a.error) det = esc(a.error);
      else if (d.tipo === 'nomina') det = `${d.registros.length} registros de nómina` + (d.periodo ? ` · ${esc(d.periodo)}` : '') + (d.empresa ? ` · ${esc(d.empresa)}` : '');
      else det = `${esc(d.empresa)} · período ${esc(d.periodo)}` + (d.numero ? ` · N.º ${esc(d.numero)}` : '') + ` · ${d.lineas.length} líneas` + (d.pagada === 'Sí' ? ' · <b>PAGADA</b>' : '');
      return `<li>${tag}<div class="meta"><div class="name">${esc(a.name)}</div><div class="det">${det}</div></div><button class="x" data-rm="${i}" aria-label="Quitar ${esc(a.name)}">×</button></li>`;
    }).join('');
    const ok = state.archivos.filter(a => a.data);
    $('btnRun').disabled = !(ok.some(a => a.data.tipo === 'planilla') && ok.some(a => a.data.tipo !== 'planilla'));
  }
  $('lista').addEventListener('click', e => {
    const i = e.target.dataset.rm; if (i === undefined) return;
    state.archivos.splice(+i, 1); pintarLista();
  });

  function mensaje(tipo, txt) {
    const d = document.createElement('div'); d.className = 'msg ' + tipo; d.textContent = txt;
    $('mensajes').prepend(d); setTimeout(() => d.remove(), tipo === 'err' ? 12000 : 6000);
  }

  // ------------------------------------------------------------ conciliar
  $('btnRun').onclick = () => {
    try {
      const archivos = state.archivos.filter(a => a.data).map(a => a.data);
      state.res = C.conciliar(archivos, state.catalogo, { tolTrabajador: +$('tolT').value || 0, tolEntidad: +$('tolE').value || 0 });
      state.tab = state.res.conteo.rojo + state.res.conteo.amarillo ? 'dif' : 'ent';
      pintarResultado();
      $('resultado').scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (e) { mensaje('err', e.message); }
  };

  function pintarResultado() {
    const r = state.res; $('resultado').classList.remove('hidden');
    $('resTitulo').textContent = `${r.empresa || 'Empresa'} · ${r.periodo}`;
    $('resSub').textContent = `NIT ${r.nitEmpresa} · Contabilidad: ${r.fuenteNomina} · Planillas: ${r.planillas.map(p => `${p.fuente} ${p.numero}`).join(', ')}`;
    $('kpis').innerHTML = [
      ['', r.conteo.total, 'Trabajadores revisados'], ['r', r.conteo.rojo, 'Con diferencias'],
      ['a', r.conteo.amarillo, 'Con avisos'], ['g', r.conteo.ok, 'Sin diferencias'],
    ].map(([c, v, l]) => `<div class="kpi ${c}"><div class="v">${v}</div><div class="l">${l}</div></div>`).join('');
    $('avisos').innerHTML = r.avisos.map(a => `<div class="msg warn">${esc(a)}</div>`).join('') +
      (r.centrosNuevos.length ? `<div class="msg warn">Hay ${r.centrosNuevos.length} centro(s) de trabajo que no están en el catálogo. Revise la pestaña "Centros nuevos".</div>` : '');
    const nDif = r.trabajadores.filter(t => t.sev !== 'ok').length;
    const nEnt = r.entidades.filter(e => e.estado !== 'ok').length;
    const tabs = [['dif', 'Diferencias', nDif], ['ent', 'Por entidad', nEnt], ['todos', 'Todos', r.trabajadores.length]];
    if (r.centrosNuevos.length) tabs.push(['cen', 'Centros nuevos', r.centrosNuevos.length]);
    $('tabs').innerHTML = tabs.map(([k, l, n]) => `<button class="tab${state.tab === k ? ' on' : ''}" role="tab" data-tab="${k}">${l}<span class="n">${n}</span></button>`).join('');
    pintarPanel();
  }
  $('tabs').addEventListener('click', e => {
    const b = e.target.closest('[data-tab]'); if (!b) return;
    state.tab = b.dataset.tab; pintarResultado();
  });

  function pintarPanel() {
    const r = state.res, p = $('panel');
    if (state.tab === 'dif') {
      const lista = r.trabajadores.filter(t => t.sev !== 'ok' && coincide(t));
      p.innerHTML = barraFiltro() + (lista.length ? lista.map(tarjeta).join('') : '<div class="empty">✅ No hay diferencias por encima de la tolerancia.</div>');
    } else if (state.tab === 'ent') {
      p.innerHTML = `<div class="tbl-wrap"><table><thead><tr><th>Subsistema</th><th>Entidad</th><th>NIT</th><th class="num">Contabilidad</th><th class="num">Planilla</th><th class="num">Diferencia</th><th>Estado</th></tr></thead><tbody>` +
        Object.keys(SUBS).map(k => r.entidades.filter(e => e.sub === k).map(e => `<tr><td>${SUBS[k]}</td><td>${esc(e.nombre)}</td><td>${esc(e.nit.replace('COD:', 'cód. '))}</td><td class="num">${money(e.nomina)}</td><td class="num">${money(e.planilla)}</td><td class="num">${money(e.dif)}</td><td><span class="pill ${e.estado}">${e.estado === 'ok' ? 'OK' : 'Revisar'}</span></td></tr>`).join('') +
          `<tr><td colspan="3"><b>Total ${SUBS[k]}</b></td><td class="num"><b>${money(r.totales[k].nomina)}</b></td><td class="num"><b>${money(r.totales[k].planilla)}</b></td><td class="num"><b>${money(r.totales[k].planilla - r.totales[k].nomina)}</b></td><td></td></tr>`).join('') +
        '</tbody></table></div>';
    } else if (state.tab === 'todos') {
      const lista = r.trabajadores.filter(coincide);
      p.innerHTML = barraFiltro() + `<div class="tbl-wrap"><table><thead><tr><th>Trabajador</th><th>Centro</th>` +
        Object.keys(SUBS).map(k => `<th class="num">${SUBS[k]}<br><span style="font-weight:400">contab. / planilla</span></th>`).join('') + '<th>Novedades</th></tr></thead><tbody>' +
        lista.map(t => `<tr><td><span class="sev ${t.sev}"></span>${esc(t.nombre)}<div class="wm">${esc(t.cedula)}</div></td><td>${esc(t.centros.join(', '))}</td>` +
          Object.keys(SUBS).map(k => `<td class="num">${money(t.nomina[k] || 0)}<br>${money(t.planilla[k] || 0)}</td>`).join('') + `<td>${esc(t.novedades.join(', '))}</td></tr>`).join('') +
        '</tbody></table></div>';
    } else if (state.tab === 'cen') {
      p.innerHTML = `<p class="hint">Estos centros aparecen en la planilla pero no en el catálogo. Defina su tipo y riesgo y agréguelos; luego descargue el catálogo y súbalo a Teams.</p>
      <div class="tbl-wrap"><table><thead><tr><th>Nombre en planilla</th><th>Código</th><th>Riesgo que liquida la planilla</th><th class="num">Trabajadores</th><th>Nombre interno</th><th>Tipo</th><th>Riesgo</th><th></th></tr></thead><tbody>` +
        r.centrosNuevos.map((c, i) => `<tr data-c="${i}"><td>${esc(c.nombrePlanilla)}</td><td>${esc(c.codigo)}</td><td>${c.riesgoObservado ? 'Riesgo ' + c.riesgoObservado : '—'}</td><td class="num">${c.trabajadores}</td>
          <td><input type="text" data-f="interno" placeholder="Ej. Obra X"></td>
          <td><select data-f="tipo"><option></option><option>Sede/Administración</option><option>Obra</option></select></td>
          <td><select data-f="riesgo">${RIESGOS.map(x => `<option${String(c.riesgoObservado || '') === x ? ' selected' : ''}>${x}</option>`).join('')}</select></td>
          <td><button class="btn small primary" data-add="${i}">Agregar al catálogo</button></td></tr>`).join('') + '</tbody></table></div>';
    }
    const f = $('filtro');
    if (f) { f.value = state.filtro; f.oninput = () => { state.filtro = f.value; const pos = f.selectionStart; pintarPanel(); const n = $('filtro'); n.focus(); n.setSelectionRange(pos, pos); }; }
  }
  $('panel').addEventListener('click', e => {
    const i = e.target.dataset.add; if (i === undefined) return;
    if (!state.catalogo) return mensaje('warn', 'Cargue primero el catálogo para poder agregar centros.');
    const tr = e.target.closest('tr'), c = state.res.centrosNuevos[+i];
    const val = f => tr.querySelector(`[data-f="${f}"]`).value;
    const riesgo = parseInt(val('riesgo'), 10) || null;
    if (!riesgo) return mensaje('warn', 'Seleccione la clase de riesgo del centro.');
    const emp = state.catalogo.empresas.find(x => x.nit === state.res.nitEmpresa);
    state.catalogo.centros.push({ empresa: emp ? emp.empresa : c.empresa, interno: val('interno'), nombrePlanilla: c.nombrePlanilla, codigo: c.codigo, tipo: val('tipo'), propia: '', riesgo, riesgoObservado: c.riesgoObservado, obs: 'Agregado desde el aplicativo ' + new Date().toLocaleDateString('es-CO') });
    marcarEditado(); pintarEditor();
    state.res.centrosNuevos.splice(+i, 1);
    if (!state.res.centrosNuevos.length) state.tab = 'dif';
    pintarResultado();
    mensaje('info', 'Centro agregado. Descargue el catálogo (paso 1 → Ver / editar centros) y reemplácelo en Teams.');
  });

  function barraFiltro() {
    return '<div class="toolbar"><input type="text" id="filtro" placeholder="Buscar por nombre, cédula o centro…" aria-label="Buscar"></div>';
  }
  function coincide(t) {
    const q = C.norm(state.filtro); if (!q) return true;
    return C.norm(t.nombre + ' ' + t.cedula + ' ' + t.centros.join(' ')).includes(q);
  }
  function tarjeta(t) {
    const h = t.hallazgos.filter(x => x.sev !== 'ok');
    const meta = [t.cedula && 'CC ' + t.cedula, t.centros.join(', '), t.planillas.length && 'Planilla ' + t.planillas.join(', '), t.novedades.length && 'Novedades: ' + t.novedades.join(', ')].filter(Boolean).join(' · ');
    return `<div class="worker"><div class="wh"><div><span class="sev ${t.sev}"></span><span class="wn">${esc(t.nombre)}</span><div class="wm">${esc(meta)}</div></div>
      <div><span class="pill ${t.sev}">${t.sev === 'rojo' ? 'Diferencia' : 'Aviso'}</span></div></div>` +
      h.map(x => `<div class="find"><div><div class="ft">${esc(x.tipo)}</div><div class="fs">${esc(SUBS[x.sub] || 'General')}${x.resp ? ' · ' + esc(x.resp) : ''}</div></div>
        <div>${esc(x.detalle)}</div>${x.accion ? `<div class="fa">➜ ${esc(x.accion)}</div>` : ''}</div>`).join('') + '</div>';
  }

  $('btnInforme').onclick = () => {
    const r = state.res; if (!r) return;
    const wb = C.informeLibro(XLSX, r);
    const per = (r.periodo || '').replace(/[^\w-]+/g, '_');
    const emp = C.norm(r.empresa).replace(/ /g, '_').slice(0, 40);
    XLSX.writeFile(wb, `Conciliacion_${emp}_${per}.xlsx`);
  };

  // ------------------------------------------------------------ inicio
  cargarCatGuardado(); pintarCat(); pintarLista();
  window.App = { manejarArchivos, leerCat, state };
})();
