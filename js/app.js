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
  // Estructura: Razón social > Sucursal principal > Sucursal > Centro de trabajo > Riesgo
  function migrar(cat) {
    if (!cat || !cat.centros) return cat;
    cat.empresas = cat.empresas || [];
    cat.centros = cat.centros.map(c => ('razon' in c) ? c : ({
      razon: c.empresa || '', sucPrincipal: '', sucursal: '', centro: c.nombrePlanilla || c.interno || '', codigo: c.codigo || '',
      operador: '', tipo: c.tipo || '', riesgo: c.riesgo || null, riesgoObservado: c.riesgoObservado || null, obs: c.obs || '',
    }));
    return cat;
  }
  function guardarCat() {
    try { localStorage.setItem(LS_CAT, JSON.stringify({ catalogo: state.catalogo, info: state.catInfo })); } catch (e) { /* sin almacenamiento */ }
  }
  function cargarCatGuardado() {
    try {
      const raw = localStorage.getItem(LS_CAT);
      if (raw) { const o = JSON.parse(raw); state.catalogo = migrar(o.catalogo); state.catInfo = o.info; }
    } catch (e) { /* sin almacenamiento */ }
  }
  function pintarCat() {
    const el = $('catEstado');
    if (!state.catalogo) {
      el.innerHTML = '<div class="msg warn">Sin catálogo: la conciliación funciona, pero no se validará la clase de riesgo por centro de trabajo.</div>';
      return;
    }
    const c = state.catalogo, i = state.catInfo || {};
    const empresas = new Set(c.centros.map(x => C.norm(x.razon)).filter(Boolean)).size;
    const sinRiesgo = c.centros.filter(x => !x.riesgo).length;
    el.innerHTML = `<b>${esc(i.archivo || 'Catálogo')}</b> · ${empresas} razones sociales · ${c.centros.length} centros` +
      (i.fecha ? ` · cargado ${esc(i.fecha)}` : '') + (i.editado ? ' · <span class="pill amarillo">editado en este equipo, falta subirlo a Teams</span>' : '') +
      (sinRiesgo ? `<div class="msg warn">${sinRiesgo} centros sin clase de riesgo confirmada: no se validan hasta completarlos.</div>` : '');
  }
  function usarCatalogo(cat, nombre) {
    state.catalogo = migrar(cat);
    state.catInfo = { archivo: nombre, fecha: new Date().toLocaleString('es-CO'), editado: false };
    guardarCat(); pintarCat(); pintarEditor(); pintarLista();
  }
  async function leerCat(file) {
    try {
      const buf = await file.arrayBuffer();
      const cat = C.leerArchivo(XLSX, new Uint8Array(buf), file.name);
      if (cat.tipo !== 'catalogo') throw new Error('Este archivo no es el catálogo (debe tener la hoja "Centros de trabajo").');
      usarCatalogo(cat, file.name);
      mensaje('info', 'Catálogo cargado.');
    } catch (e) { mensaje('err', e.message); }
  }

  // ---- Editor de centros
  const RIESGOS = ['', '1', '2', '3', '4', '5'];
  const TIPOS = ['', 'Sede/Administración', 'Obra'];
  const opt = (arr, v) => arr.map(o => `<option${o === String(v || '') ? ' selected' : ''}>${esc(o)}</option>`).join('');
  function pintarEditor() {
    const t = $('tblCat');
    if (!state.catalogo) { t.innerHTML = '<tr><td class="empty">Cargue primero el catálogo o use "Extraer centros de las planillas".</td></tr>'; return; }
    const orden = state.catalogo.centros.map((c, i) => [c, i])
      .sort((a, b) => C.norm(a[0].razon).localeCompare(C.norm(b[0].razon)) || C.norm(a[0].sucursal).localeCompare(C.norm(b[0].sucursal)));
    t.innerHTML = '<thead><tr><th>Razón social</th><th>Sucursal principal</th><th>Sucursal</th><th>Centro de trabajo</th><th>Tipo</th><th>Riesgo</th><th></th></tr></thead><tbody>' +
      orden.map(([c, i]) => `<tr data-i="${i}">
        <td><input type="text" data-f="razon" value="${esc(c.razon)}"></td>
        <td><input type="text" data-f="sucPrincipal" value="${esc(c.sucPrincipal)}"></td>
        <td><input type="text" data-f="sucursal" value="${esc(c.sucursal)}"></td>
        <td><input type="text" data-f="centro" value="${esc(c.centro)}"></td>
        <td><select data-f="tipo">${opt(TIPOS, c.tipo)}</select></td>
        <td><select data-f="riesgo">${opt(RIESGOS, c.riesgo)}</select></td>
        <td><button class="x" data-del="${i}" title="Quitar">×</button></td></tr>`).join('') + '</tbody>';
  }
  $('tblCat').addEventListener('change', e => {
    const tr = e.target.closest('tr[data-i]'); if (!tr) return;
    const c = state.catalogo.centros[+tr.dataset.i], f = e.target.dataset.f;
    c[f] = f === 'riesgo' ? (parseInt(e.target.value, 10) || null) : e.target.value.trim();
    marcarEditado(); pintarDetectados();
  });
  $('tblCat').addEventListener('click', e => {
    const i = e.target.dataset.del; if (i === undefined) return;
    if (!confirm('¿Quitar este centro del catálogo?')) return;
    state.catalogo.centros.splice(+i, 1); marcarEditado(); pintarEditor(); pintarDetectados();
  });
  function marcarEditado() { state.catInfo = Object.assign({}, state.catInfo, { editado: true }); guardarCat(); pintarCat(); }
  function asegurarCatalogo() {
    if (!state.catalogo) {
      state.catalogo = { tipo: 'catalogo', archivo: 'Catalogo_Centros_Trabajo.xlsx', empresas: [], centros: [], entidades: {} };
      state.catInfo = { archivo: 'Catálogo nuevo (sin guardar)', fecha: new Date().toLocaleString('es-CO'), editado: true };
    }
  }
  $('btnAddCentro').onclick = () => {
    asegurarCatalogo();
    state.catalogo.centros.push({ razon: '', sucPrincipal: '', sucursal: '', centro: '', codigo: '', operador: '', tipo: '', riesgo: null, riesgoObservado: null, obs: 'Agregado desde el aplicativo' });
    marcarEditado(); pintarEditor();
    const inp = $('tblCat').querySelector(`tr[data-i="${state.catalogo.centros.length - 1}"] input`); if (inp) inp.focus();
  };
  $('btnCatEditar').onclick = () => { $('catEditor').classList.toggle('hidden'); pintarEditor(); pintarDetectados(); };
  $('btnCerrarCat').onclick = () => $('catEditor').classList.add('hidden');
  $('btnDescCat').onclick = descargarCatalogo;

  // ---- Extraer centros de las planillas cargadas
  function planillasCargadas() { return state.archivos.filter(a => a.data && a.data.tipo === 'planilla').map(a => a.data); }
  $('btnExtraer').onclick = () => {
    if (!planillasCargadas().length) return mensaje('warn', 'Primero cargue una o varias planillas en el paso 2.');
    state.verDetectados = true;
    $('catEditor').classList.remove('hidden'); pintarEditor(); pintarDetectados();
    $('catEditor').scrollIntoView({ behavior: 'smooth', block: 'start' });
  };
  const ESTADO = {
    ok: ['ok', 'En catálogo'], 'sin riesgo': ['amarillo', 'En catálogo sin riesgo'], distinto: ['rojo', 'Riesgo distinto'], nuevo: ['amarillo', 'Nuevo'],
  };
  function filaEstructura(e, i, conAccion) {
    const [cls, txt] = ESTADO[e.estado];
    const accion = e.estado === 'nuevo'
      ? `<td><select data-f="tipo">${opt(TIPOS, e.riesgoObservado === 1 ? 'Sede/Administración' : e.riesgoObservado ? 'Obra' : '')}</select></td>
         <td><select data-f="riesgo">${opt(RIESGOS, e.riesgoObservado)}</select></td>
         <td><button class="btn small primary" data-addest="${i}">Agregar</button></td>`
      : `<td></td><td>${e.riesgoCatalogo ? 'Riesgo ' + e.riesgoCatalogo : '—'}</td><td></td>`;
    return `<tr data-e="${i}"><td>${esc(e.razon)}</td><td>${esc(e.sucPrincipal)}</td><td>${esc(e.sucursal)}</td><td><b>${esc(e.centro)}</b></td>
      <td>${e.clasesObservadas.length ? e.clasesObservadas.map(x => 'Riesgo ' + x).join(' / ') : '—'}</td><td class="num">${e.trabajadores}</td>
      <td><span class="pill ${cls}">${txt}</span></td>${conAccion ? accion : ''}</tr>`;
  }
  const CAB_EST = '<thead><tr><th>Razón social</th><th>Sucursal principal</th><th>Sucursal</th><th>Centro de trabajo</th><th>Riesgo en planilla</th><th class="num">Trab.</th><th>Estado</th><th>Tipo</th><th>Riesgo a confirmar</th><th></th></tr></thead>';
  function pintarDetectados() {
    const box = $('detectados');
    if (!state.verDetectados) { box.classList.add('hidden'); return; }
    const pl = planillasCargadas();
    state.estructura = C.extraerEstructura(pl, state.catalogo);
    const nuevos = state.estructura.filter(e => e.estado === 'nuevo').length;
    box.classList.remove('hidden');
    box.innerHTML = `<div class="res-head" style="margin-bottom:8px"><div><h2 style="font-size:15px;margin:0">Centros encontrados en las planillas (${state.estructura.length})</h2>
      <p class="hint" style="margin:4px 0 0">Leídos de ${pl.length} planilla(s). Para cada centro <b>Nuevo</b>, confirme tipo y riesgo y pulse <b>Agregar</b>. Lo que queda en el catálogo es lo que usted define.</p></div>
      <div class="row">${nuevos ? `<button class="btn small" id="btnAddTodos">Agregar los ${nuevos} nuevos</button>` : ''}<button class="btn small" id="btnOcultarDet">Ocultar</button></div></div>
      <div class="tbl-wrap"><table>${CAB_EST}<tbody>${state.estructura.map((e, i) => filaEstructura(e, i, true)).join('')}</tbody></table></div>`;
    const t = $('btnAddTodos');
    if (t) t.onclick = () => {
      box.querySelectorAll('tr[data-e]').forEach(tr => { const b = tr.querySelector('[data-addest]'); if (b) agregarDesdeFila(tr, state.estructura[+tr.dataset.e], true); });
      terminarAgregado();
    };
    $('btnOcultarDet').onclick = () => { state.verDetectados = false; pintarDetectados(); };
  }
  function agregarDesdeFila(tr, e, silencioso) {
    const v = f => tr.querySelector(`[data-f="${f}"]`).value;
    const riesgo = parseInt(v('riesgo'), 10) || null;
    if (!riesgo && !silencioso) { mensaje('warn', 'Seleccione la clase de riesgo del centro.'); return false; }
    asegurarCatalogo();
    state.catalogo.centros.push({
      razon: e.razon, sucPrincipal: e.sucPrincipal, sucursal: e.sucursal, centro: e.centro, codigo: e.codigo, operador: e.operador,
      tipo: v('tipo'), riesgo, riesgoObservado: e.riesgoObservado, obs: `Extraído de planilla ${e.planillas.join(', ')} el ${new Date().toLocaleDateString('es-CO')}`,
    });
    return true;
  }
  function terminarAgregado() {
    marcarEditado(); pintarEditor(); pintarDetectados();
    if (state.res) { state.res.centrosNuevos = C.extraerEstructura(planillasCargadas(), state.catalogo).filter(e => e.estado === 'nuevo'); pintarResultado(); }
    mensaje('info', 'Centro(s) agregado(s). Pulse "Descargar catálogo" y reemplace el archivo en la carpeta Catalogo de Teams.');
  }
  $('catEditor').addEventListener('click', e => {
    const i = e.target.dataset.addest; if (i === undefined) return;
    if (agregarDesdeFila(e.target.closest('tr'), state.estructura[+i], false)) terminarAgregado();
  });

  function descargarCatalogo() {
    if (!state.catalogo) return;
    const c = state.catalogo, wb = XLSX.utils.book_new();
    const tarifas = [0.00522, 0.01044, 0.02436, 0.0435, 0.0696];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['CATÁLOGO DE CENTROS DE TRABAJO – Conciliación de aportes'],
      ['Generado desde el aplicativo el ' + new Date().toLocaleString('es-CO')],
      ['Estructura: Razón social > Sucursal principal > Sucursal > Centro de trabajo > Clase de riesgo.'],
      ['Escriba los nombres tal como aparecen en la planilla del operador. "TODOS" o "n/a" sirve para cualquier sucursal o centro.'],
      ['Riesgo 1 = 0,522% · 2 = 1,044% · 3 = 2,436% · 4 = 4,350% · 5 = 6,960%'],
    ]), 'Instrucciones');
    const filas = [...c.centros].sort((a, b) => C.norm(a.razon).localeCompare(C.norm(b.razon)) || C.norm(a.sucursal).localeCompare(C.norm(b.sucursal)));
    const ws = XLSX.utils.aoa_to_sheet([['Razon Social', 'Sucursal Principal', 'SUCURSAL', 'CENTRO DE TRABAJO', 'Operador', 'Tipo', 'Riesgo en planilla', 'Clase de riesgo CONFIRMADA (1-5)', 'Tarifa ARL esperada', 'Observaciones'],
      ...filas.map(x => [x.razon, x.sucPrincipal, x.sucursal, x.centro, x.operador || '', x.tipo, x.riesgoObservado || '', x.riesgo || '', x.riesgo ? tarifas[x.riesgo - 1] : '', x.obs || ''])]);
    ws['!cols'] = [40, 34, 34, 26, 16, 20, 12, 14, 12, 45].map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, 'Centros de trabajo');
    if (c.empresas && c.empresas.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Empresa', 'NIT', 'Software contable', 'Operador PILA'], ...c.empresas.map(e => [e.empresa, e.nit, e.software, e.operador])]), 'Empresas');
    const ents = Object.entries(c.entidades || {});
    if (ents.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['Código PILA', 'NIT', 'Nombre'], ...ents.map(([k, v]) => [k, v[0], v[1]])]), 'Entidades');
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
        if (d.tipo === 'catalogo') { usarCatalogo(d, f.name); continue; }
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
    $('btnExtraer').disabled = !ok.some(a => a.data.tipo === 'planilla');
    if (state.verDetectados) pintarDetectados();
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
      p.innerHTML = `<p class="hint">Estos centros aparecen en la planilla pero no en el catálogo. Confirme tipo y riesgo y pulse <b>Agregar</b>; luego descargue el catálogo (paso 1 → Ver / editar centros) y súbalo a Teams.</p>
      <div class="tbl-wrap"><table>${CAB_EST}<tbody>${r.centrosNuevos.map((e, i) => filaEstructura(e, i, true).replace('data-addest=', 'data-addres=')).join('')}</tbody></table></div>`;
    }
    const f = $('filtro');
    if (f) { f.value = state.filtro; f.oninput = () => { state.filtro = f.value; const pos = f.selectionStart; pintarPanel(); const n = $('filtro'); n.focus(); n.setSelectionRange(pos, pos); }; }
  }
  $('panel').addEventListener('click', e => {
    const i = e.target.dataset.addres; if (i === undefined) return;
    if (agregarDesdeFila(e.target.closest('tr'), state.res.centrosNuevos[+i], false)) {
      terminarAgregado();
      if (!state.res.centrosNuevos.length) { state.tab = 'dif'; pintarResultado(); }
    }
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

  $('btnInforme').onclick = async () => {
    const r = state.res; if (!r) return;
    const per = (r.periodo || '').replace(/[^\w-]+/g, '_');
    const emp = C.norm(r.empresa).replace(/ /g, '_').slice(0, 40);
    const nombre = `Conciliacion_${emp}_${per}.xlsx`;
    const btn = $('btnInforme'); btn.disabled = true; btn.textContent = 'Generando…';
    try {
      if (window.ExcelJS && window.Informe) await window.Informe.descargar(window.ExcelJS, r, C, nombre);
      else XLSX.writeFile(C.informeLibro(XLSX, r), nombre); // respaldo sin formato
    } catch (e) {
      mensaje('err', 'No se pudo generar el informe con formato (' + e.message + '). Se descarga la versión simple.');
      XLSX.writeFile(C.informeLibro(XLSX, r), nombre);
    } finally { btn.disabled = false; btn.textContent = 'Descargar informe Excel'; }
  };

  // ------------------------------------------------------------ inicio
  cargarCatGuardado(); pintarCat(); pintarLista();
  window.App = { manejarArchivos, leerCat, state };
})();
