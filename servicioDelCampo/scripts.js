(function () {
    // --- CONSTANTES ---
    const KEYS = {
        CONDUCTORES: 'conductores_v1',
        LUGARES: 'lugares_v1',
        GRUPOS: 'grupos_v1',
        MAPA: 'mapa_coherencia_v1',
        MANUALES: 'manuales_v5',
        BLOQUEADOS: 'bloqueados_v1',
        HORARIO: 'horario_v2',
        CONFIG_COLS: 'config_cols_v1',
        CONFIG_ORDER: 'config_order_v2',
        CONFIG_DIAS: 'config_dias_v1',
        EXTRAS: 'filas_extras_v2'
    };

    const COL_DEFS = {
        lug: "Lugar de salida",
        ter: "Territorio",
        zon: "Zona",
        cond: "Conductor",
        cua: "Cuadra",
        gru: "Grupo"
    };

    const INITIAL_LUGARES = [];
    const INITIAL_MAPA = {};

    const getRangoTerritorios = () => {
        const todos = Object.values(state.mapaCoherencia).flatMap(m => m.ter || []);
        return [...new Set(todos)].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    };

    const getRangoZonas = () => {
        const todas = Object.values(state.mapaCoherencia).flatMap(m => m.zon || []);
        return [...new Set(todas)].sort();
    };

    const formatearGruposParaPDF = (val) => {
        if (!val || val === "-" || val === "") return "Todos";
        const lista = val.split(' + ').map(g => g.replace(/Grupo\s+/i, '').trim()).filter(g => g !== "");
        if (lista.length === 0) return "Todos";
        if (lista.length === 1) return "Grupo " + lista[0];
        const ultimo = lista.pop();
        return "Grupos " + lista.join(", ") + " y " + ultimo;
    };

    const DIAS_SEMANA = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

    // --- ESTADO ---
    let state = {
        conductores: [],
        lugares: [],
        grupos: [],
        mapaCoherencia: {},
        manuales: {},
        bloqueados: {},
        asignaciones: {},
        horario: { m: '', t: '', mFin: '', tFin: '' },
        filasExtras: {},
        colOrder: ['lug', 'ter', 'zon', 'gru', 'cond'],
        configDias: {}
    };

    // --- ELEMENTOS DOM ---
    const els = {
        fInicio: document.getElementById('fechaInicio'),
        fFinal: document.getElementById('fechaFinal'),
        cantFechas: document.getElementById('cantidadFechas'),
        nuevoCond: document.getElementById('nuevoConductorNombre'),
        listaCond: document.getElementById('listaConductores'),
        nuevoLugar: document.getElementById('nuevoLugarNombre'),
        listaLugares: document.getElementById('listaLugares'),
        nuevoGrupo: document.getElementById('nuevoGrupoNombre'),
        listaGrupos: document.getElementById('listaGrupos'),
        tablaBody: document.getElementById('cuerpoTabla'),
        headerRow: document.getElementById('headerRow'),
        btnBloquear: document.getElementById('btnBloquearTodo'),
        hManana: document.getElementById('hMananaInput'),
        hTarde: document.getElementById('hTardeInput'),
        hMananaFin: document.getElementById('hMananaFindeInput'),
        hTardeFin: document.getElementById('hTardeFindeInput'),
        diasContainer: document.getElementById('selectorDiasContainer')
    };

    // --- UTILIDADES ---
    const stringToDate = (d) => {
        if (!d) return new Date();
        const [y, m, day] = d.split('-').map(Number);
        return new Date(y, m - 1, day);
    };

    const generarFechasDiarias = () => {
        const inicio = els.fInicio.value;
        const final = els.fFinal.value;
        if (!inicio || !final) return [];
        const lista = [];
        let f = stringToDate(inicio);
        const fEnd = stringToDate(final);
        while (f <= fEnd) {
            const diaIndex = f.getDay();
            if (state.configDias[diaIndex]?.m || state.configDias[diaIndex]?.t) {
                lista.push(new Date(f));
            }
            f.setDate(f.getDate() + 1);
        }
        return lista;
    };

    const obtenerDatosDia = (id) => {
        const man = state.manuales[id] || {};
        const asig = state.asignaciones[id] || {};
        return {
            lugM: man.lugM || asig.lugM, lugT: man.lugT || asig.lugT,
            terM: man.terM || asig.terM, terT: man.terT || asig.terT,
            zonM: man.zonM || asig.zonM, zonT: man.zonT || asig.zonT,
            condM: man.condM || asig.condM, condT: man.condT || asig.condT,
            cuaM: man.cuaM || "", cuaT: man.cuaT || "",
            gruM: man.gruM || "", gruT: man.gruT || ""
        };
    };

    // --- PERSISTENCIA ---
    function loadFromStorage() {
        try {
            state.conductores = JSON.parse(localStorage.getItem(KEYS.CONDUCTORES)) || [];
            state.lugares = JSON.parse(localStorage.getItem(KEYS.LUGARES)) || INITIAL_LUGARES;
            state.grupos = JSON.parse(localStorage.getItem(KEYS.GRUPOS)) || [];
            state.mapaCoherencia = JSON.parse(localStorage.getItem(KEYS.MAPA)) || INITIAL_MAPA;
            state.manuales = JSON.parse(localStorage.getItem(KEYS.MANUALES)) || {};
            state.bloqueados = JSON.parse(localStorage.getItem(KEYS.BLOQUEADOS)) || {};
            state.horario = JSON.parse(localStorage.getItem(KEYS.HORARIO)) || { m: '', t: '', mFin: '', tFin: '' };
            state.filasExtras = JSON.parse(localStorage.getItem(KEYS.EXTRAS)) || {};
            const storedOrder = JSON.parse(localStorage.getItem(KEYS.CONFIG_ORDER));
            if (storedOrder) state.colOrder = storedOrder;
            const storedDias = JSON.parse(localStorage.getItem(KEYS.CONFIG_DIAS));
            state.configDias = storedDias || {};
            if (!storedDias) { for (let i = 0; i < 7; i++) state.configDias[i] = { m: true, t: (i !== 0) }; }
        } catch (e) { console.error("Error loading storage", e); }
    }

    function saveToStorage() {
        localStorage.setItem(KEYS.CONDUCTORES, JSON.stringify(state.conductores));
        localStorage.setItem(KEYS.LUGARES, JSON.stringify(state.lugares));
        localStorage.setItem(KEYS.GRUPOS, JSON.stringify(state.grupos));
        localStorage.setItem(KEYS.MAPA, JSON.stringify(state.mapaCoherencia));
        localStorage.setItem(KEYS.MANUALES, JSON.stringify(state.manuales));
        localStorage.setItem(KEYS.BLOQUEADOS, JSON.stringify(state.bloqueados));
        localStorage.setItem(KEYS.HORARIO, JSON.stringify(state.horario));
        localStorage.setItem(KEYS.CONFIG_ORDER, JSON.stringify(state.colOrder));
        localStorage.setItem(KEYS.CONFIG_DIAS, JSON.stringify(state.configDias));
        localStorage.setItem(KEYS.EXTRAS, JSON.stringify(state.filasExtras));
    }

    // --- ASIGNACIÓN ---
    function asignarGenerico(tipo, targetId = null) {
        if (tipo === 'conductor' && state.conductores.length === 0) return alert("Faltan conductores.");
        if (tipo === 'territorio' && state.lugares.length === 0) return alert("Faltan lugares.");
        const fechas = generarFechasDiarias();
        fechas.forEach(f => {
            const id = f.toISOString().split('T')[0];
            if ((targetId && id !== targetId) || state.bloqueados[id]) return;
            if (!state.asignaciones[id]) state.asignaciones[id] = {};
            const asig = state.asignaciones[id];
            if (tipo === 'conductor') { delete asig.condM; delete asig.condT; }
            else { ['lug', 'ter', 'zon'].forEach(k => { delete asig[k + 'M']; delete asig[k + 'T']; }); }
            const diaNum = f.getDay();
            const configDia = state.configDias[diaNum];
            const esFinde = (diaNum === 0 || diaNum === 6);
            const horaTarde = esFinde ? (state.horario.tFin || state.horario.t) : state.horario.t;
            const usarTarde = configDia.t && horaTarde !== '';
            const usadosHoy = [];
            ['M', 'T'].forEach(turno => {
                if (turno === 'T' && !usarTarde) return;
                if (turno === 'M' && !configDia.m) return;
                if (tipo === 'conductor') {
                    const manual = state.manuales[id]?.[`cond${turno}`];
                    if (manual) usadosHoy.push(manual);
                    else {
                        const pool = state.conductores.filter(c => !usadosHoy.includes(c));
                        const elegido = pool[Math.floor(Math.random() * pool.length)];
                        if (elegido) { asig[`cond${turno}`] = elegido; usadosHoy.push(elegido); }
                    }
                } else {
                    const manualLug = state.manuales[id]?.[`lug${turno}`];
                    let lugAct = manualLug;
                    if (!lugAct) {
                        const pool = state.lugares.filter(l => !usadosHoy.includes(l));
                        const elegido = pool[Math.floor(Math.random() * pool.length)];
                        if (elegido) { asig[`lug${turno}`] = elegido; usadosHoy.push(elegido); lugAct = elegido; }
                    } else usadosHoy.push(lugAct);
                    const cfg = state.mapaCoherencia[lugAct];
                    if (cfg) {
                        if (!state.manuales[id]?.[`ter${turno}`] && cfg.ter.length > 0) asig[`ter${turno}`] = cfg.ter[0];
                        if (!state.manuales[id]?.[`zon${turno}`] && cfg.zon.length > 0) {
                            const terAct = state.manuales[id]?.[`ter${turno}`] || asig[`ter${turno}`];
                            asig[`zon${turno}`] = cfg.zon.find(z => z.startsWith(terAct + '-')) || cfg.zon[0];
                        }
                    }
                }
            });
        });
        actualizarTodo();
    }

    function ajustarFechas(modo) {
        if (modo === 'combo') {
            if (els.cantFechas.value === "custom") return;
            const inicioStr = els.fInicio.value;
            if (!inicioStr) return;
            let fFin = stringToDate(inicioStr);
            fFin.setDate(fFin.getDate() + (parseInt(els.cantFechas.value, 10) - 1));
            els.fFinal.value = fFin.toISOString().split('T')[0];
        } else if (modo === 'final') {
            const inicio = stringToDate(els.fInicio.value);
            const final = stringToDate(els.fFinal.value);
            if (final < inicio) { alert("La fecha final no puede ser anterior."); ajustarFechas('combo'); return; }
            els.cantFechas.value = "custom";
        }
        actualizarTodo();
    }

    function modificarLista(tipo, accion, indice = null) {
        let lista, input;
        if (tipo === 'conductor') { lista = state.conductores; input = els.nuevoCond; }
        else if (tipo === 'lugar') { lista = state.lugares; input = els.nuevoLugar; }
        else if (tipo === 'grupo') { lista = state.grupos; input = els.nuevoGrupo; }

        if (accion === 'agregar') {
            const raw = input.value.trim(); if (!raw) return;
            if (tipo === 'grupo') {
                const range = raw.match(/^(\d+)-(\d+)$/);
                if (range) {
                    for (let i = parseInt(range[1]); i <= parseInt(range[2]); i++) {
                        if (!lista.includes(`Grupo ${i}`)) lista.push(`Grupo ${i}`);
                    }
                    input.value = ''; actualizarTodo(); return;
                }
            }
            if (!lista.includes(raw)) {
                if (tipo === 'lugar') {
                    const t = prompt("Territorios (separados por coma):") || "";
                    const z = prompt("Zonas (separadas por coma):") || "";
                    state.mapaCoherencia[raw] = { ter: t.split(',').map(s => s.trim()).filter(s => s), zon: z.split(',').map(s => s.trim()).filter(s => s) };
                }
                lista.push(raw); input.value = ''; actualizarTodo();
            }
        } else if (accion === 'eliminar') {
            if (confirm(`¿Eliminar "${lista[indice]}"?`)) { if (tipo === 'lugar') delete state.mapaCoherencia[lista[indice]]; lista.splice(indice, 1); actualizarTodo(); }
        } else if (accion === 'borrarTodo') { if (confirm("¿Borrar todos?")) { state.grupos = []; actualizarTodo(); } }
    }

    function guardarManual(fechaId, valor, rol) {
        if (state.bloqueados[fechaId]) return;
        if (!state.manuales[fechaId]) state.manuales[fechaId] = {};
        if (!valor) delete state.manuales[fechaId][rol];
        else {
            state.manuales[fechaId][rol] = valor;
            if (rol.startsWith('lug')) {
                const t = rol.slice(-1), cfg = state.mapaCoherencia[valor];
                if (cfg && cfg.ter.length > 0) {
                    state.manuales[fechaId][`ter${t}`] = cfg.ter[0];
                    if (cfg.zon.length > 0) state.manuales[fechaId][`zon${t}`] = cfg.zon.find(z => z.startsWith(cfg.ter[0] + '-')) || cfg.zon[0];
                }
            }
        }
        actualizarTodo();
    }

    function agregarFilaExtra(dateId, turno) {
        if (state.bloqueados[dateId]) return;
        if (!state.filasExtras[dateId]) state.filasExtras[dateId] = [];
        state.filasExtras[dateId].push({ id: Date.now(), turno, lug: '', ter: '', zon: '', cond: '', cua: '', gru: '' });
        actualizarTodo();
    }

    function eliminarFilaExtra(dateId, rowId) {
        if (state.bloqueados[dateId]) return;
        state.filasExtras[dateId] = (state.filasExtras[dateId] || []).filter(r => r.id !== rowId);
        actualizarTodo();
    }

    function guardarExtra(dateId, rowId, campo, valor) {
        if (state.bloqueados[dateId]) return;
        const row = state.filasExtras[dateId].find(r => r.id === rowId);
        if (row) {
            row[campo] = valor;
            if (campo === 'lug') {
                const cfg = state.mapaCoherencia[valor];
                if (cfg && cfg.ter.length > 0) {
                    row.ter = cfg.ter[0];
                    if (cfg.zon.length > 0) row.zon = cfg.zon.find(z => z.startsWith(row.ter + '-')) || cfg.zon[0];
                }
            }
            actualizarTodo();
        }
    }

    function crearSelectorGrupos(valorActual, locked, onChangeCallback) {
        const container = document.createElement('div');
        container.style.cssText = 'display:flex; flex-direction:column; gap:2px; width:100%';
        let seleccionados = valorActual ? valorActual.split(' + ') : [""];
        seleccionados.forEach((val, idx) => {
            const row = document.createElement('div'); row.style.display = 'flex'; row.style.gap = '2px';
            const sel = document.createElement('select'); sel.style.flex = '1'; sel.add(new Option("-", ""));
            state.grupos.forEach(g => sel.add(new Option(g, g, false, g === val)));
            sel.value = val; sel.disabled = locked;
            sel.onchange = (e) => { seleccionados[idx] = e.target.value; onChangeCallback(seleccionados.filter(v => v !== "").join(' + ')); };
            row.appendChild(sel);
            if (!locked && seleccionados.length > 1) {
                const btn = document.createElement('button'); btn.textContent = '×'; btn.style.color = 'red';
                btn.onclick = () => { seleccionados.splice(idx, 1); onChangeCallback(seleccionados.filter(v => v !== "").join(' + ')); };
                row.appendChild(btn);
            }
            container.appendChild(row);
        });
        if (!locked) {
            const btnAdd = document.createElement('button'); btnAdd.textContent = '+ Grupo'; btnAdd.style.fontSize = '0.75em';
            btnAdd.onclick = () => { seleccionados.push(""); onChangeCallback(seleccionados.join(' + ')); };
            container.appendChild(btnAdd);
        }
        return container;
    }

    function renderColumnSelector() {
        const grid = document.querySelector('.columns-grid'); if (!grid) return; grid.innerHTML = '';
        const all = [{ id: 'lug', lbl: 'Lugar' }, { id: 'ter', lbl: 'Territorio' }, { id: 'zon', lbl: 'Zona' }, { id: 'cond', lbl: 'Conductor' }, { id: 'cua', lbl: 'Cuadra' }, { id: 'gru', lbl: 'Grupo' }];
        all.forEach(c => {
            const label = document.createElement('label'); label.className = 'lb-col';
            const input = document.createElement('input'); input.type = 'checkbox'; input.value = c.id; input.checked = state.colOrder.includes(c.id);
            input.onchange = (e) => { if (e.target.checked) { if (!state.colOrder.includes(c.id)) state.colOrder.push(c.id); } else { state.colOrder = state.colOrder.filter(k => k !== c.id); } actualizarTodo(); };
            label.append(input, ' ' + c.lbl); grid.appendChild(label);
        });
    }

    function renderSelectorDias() {
        if (!els.diasContainer) return; els.diasContainer.innerHTML = '';
        DIAS_SEMANA.forEach((n, i) => {
            const row = document.createElement('div'); row.className = 'day-config-row';
            row.innerHTML = `<span style="width:80px">${n}</span>
                <label>Mañana <input type="checkbox" data-day="${i}" data-shift="m" ${state.configDias[i].m ? 'checked' : ''}></label>
                <label>Tarde <input type="checkbox" data-day="${i}" data-shift="t" ${state.configDias[i].t ? 'checked' : ''}></label>`;
            els.diasContainer.appendChild(row);
        });
    }

    function renderListas() {
        const draw = (l, container, tipo) => {
            container.innerHTML = '';
            l.forEach((item, i) => {
                const li = document.createElement('li'); li.innerHTML = `<span>${item}</span>`;
                const btn = document.createElement('button'); btn.textContent = '×'; btn.onclick = () => modificarLista(tipo, 'eliminar', i);
                li.appendChild(btn); container.appendChild(li);
            });
        };
        draw(state.conductores, els.listaCond, 'conductor');
        draw(state.lugares, els.listaLugares, 'lugar');
        draw(state.grupos, els.listaGrupos, 'grupo');
    }

    function renderTabla() {
        els.headerRow.innerHTML = '<th>Fecha</th><th>Hora</th>';
        state.colOrder.forEach(k => { const th = document.createElement('th'); th.textContent = COL_DEFS[k]; els.headerRow.appendChild(th); });
        els.headerRow.appendChild(Object.assign(document.createElement('th'), { className: "th-actions", textContent: "Acciones" }));
        els.tablaBody.innerHTML = '';
        const fechas = generarFechasDiarias();
        fechas.forEach(f => {
            const id = f.toISOString().split('T')[0], locked = state.bloqueados[id], dNum = f.getDay(), cfg = state.configDias[dNum];
            const esFinde = (dNum === 0 || dNum === 6), hM = esFinde ? (state.horario.mFin || state.horario.m) : state.horario.m, hT = esFinde ? (state.horario.tFin || state.horario.t) : state.horario.t;
            const showM = cfg.m, showT = cfg.t && hT !== '', extras = state.filasExtras[id] || [], datos = obtenerDatosDia(id);

            const tr = document.createElement('tr'); if (locked) tr.style.opacity = "0.7";
            const tdF = tr.insertCell(); tdF.innerHTML = `<b>${DIAS_SEMANA[dNum].slice(0, 3)}</b><br>${f.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit' })}`;
            tdF.rowSpan = 1 + extras.length; tdF.style.verticalAlign = 'top';
            tr.insertCell().innerHTML = `<div class="stacked-cell">${showM ? `<div class="time-text ${showT ? 'border-bottom' : ''}">${hM || '--:--'}</div>` : ''}${showT ? `<div class="time-text">${hT}</div>` : ''}</div>`;

            state.colOrder.forEach(tipo => {
                const td = tr.insertCell(); const cont = document.createElement('div'); cont.className = 'stacked-cell';
                ['M', 'T'].forEach(turno => {
                    if ((turno === 'M' && !showM) || (turno === 'T' && !showT)) return;
                    const r = `${tipo}${turno}`, val = datos[r];
                    if (tipo === 'gru') cont.appendChild(crearSelectorGrupos(val, locked, (v) => guardarManual(id, v, r)));
                    else if (tipo === 'cua') { const i = Object.assign(document.createElement('input'), { type: 'text', className: 'input-cuadra', value: val || "", disabled: locked }); i.onchange = (e) => guardarManual(id, e.target.value, r); cont.appendChild(i); }
                    else {
                        const s = document.createElement('select'); s.disabled = locked; s.add(new Option("-", ""));
                        let ops = (tipo === 'lug') ? state.lugares : (tipo === 'cond') ? state.conductores : (tipo === 'ter' ? (state.mapaCoherencia[datos[`lug${turno}`]]?.ter || getRangoTerritorios()) : (state.mapaCoherencia[datos[`lug${turno}`]]?.zon.filter(z => z.startsWith(datos[`ter${turno}`] + '-')) || getRangoZonas()));
                        ops.forEach(o => s.add(new Option(o, o, false, val === o))); s.onchange = (e) => guardarManual(id, e.target.value, r); cont.appendChild(s);
                    }
                });
                td.appendChild(cont);
            });
            const tdA = tr.insertCell(); tdA.innerHTML = `<div class="actions-wrapper"><div><button class="btn-lock">${locked ? '🔒' : '🔓'}</button><button class="btn-assign" ${locked ? 'disabled' : ''}>↻</button></div><div>${showM ? '<button class="btn-add-m">+M</button>' : ''}${showT ? '<button class="btn-add-t">+T</button>' : ''}</div></div>`;
            tdA.querySelector('.btn-lock').onclick = () => { state.bloqueados[id] = !state.bloqueados[id]; actualizarTodo(); };
            tdA.querySelector('.btn-assign').onclick = () => { asignarGenerico('conductor', id); asignarGenerico('territorio', id); };
            if (showM) tdA.querySelector('.btn-add-m').onclick = () => agregarFilaExtra(id, 'M');
            if (showT) tdA.querySelector('.btn-add-t').onclick = () => agregarFilaExtra(id, 'T');
            els.tablaBody.appendChild(tr);

            extras.forEach(ex => {
                const trx = document.createElement('tr'); if (locked) trx.style.opacity = "0.7";
                trx.insertCell().innerHTML = `<span style="border-left:3px solid ${ex.turno === 'M' ? '#2196F3' : '#FF9800'}; padding-left:5px; font-weight:bold;">${(ex.turno === 'M' ? hM : hT) || '--:--'}</span>`;
                state.colOrder.forEach(t => {
                    const td = trx.insertCell();
                    if (t === 'gru') td.appendChild(crearSelectorGrupos(ex[t], locked, (v) => guardarExtra(id, ex.id, t, v)));
                    else if (t === 'cua') { const i = Object.assign(document.createElement('input'), { type: 'text', className: 'input-cuadra', value: ex[t] || "", disabled: locked }); i.onchange = (e) => guardarExtra(id, ex.id, t, e.target.value); td.appendChild(i); }
                    else {
                        const s = document.createElement('select'); s.disabled = locked; s.add(new Option("-", ""));
                        let ops = (t === 'lug' ? state.lugares : t === 'cond' ? state.conductores : t === 'ter' ? (state.mapaCoherencia[ex.lug]?.ter || getRangoTerritorios()) : (state.mapaCoherencia[ex.lug]?.zon.filter(z => z.startsWith(ex.ter + '-')) || getRangoZonas()));
                        ops.forEach(o => s.add(new Option(o, o, false, ex[t] === o))); s.onchange = (e) => guardarExtra(id, ex.id, t, e.target.value); td.appendChild(s);
                    }
                });
                const tdAx = trx.insertCell(); const wrap = document.createElement('div'); wrap.className = 'actions-wrapper';
                const bDel = document.createElement('button'); bDel.textContent = '×'; bDel.style.color = 'red'; bDel.disabled = locked; bDel.onclick = () => eliminarFilaExtra(id, ex.id);
                wrap.appendChild(bDel); tdAx.appendChild(wrap); els.tablaBody.appendChild(trx);
            });
        });
        const ids = fechas.map(f => f.toISOString().split('T')[0]);
        els.btnBloquear.innerHTML = (ids.length > 0 && ids.every(i => state.bloqueados[i])) ? '🔓 Desbloquear Todo' : '🔒 Bloquear Todo';
    }

    function actualizarTodo() { renderListas(); renderColumnSelector(); renderSelectorDias(); renderTabla(); saveToStorage(); }

    // --- PDF EXPORT (COMPACTO EXTREMO - SIN ESPACIOS ENTRE SEMANAS) ---
    function exportarPDF() {
        if (!window.jspdf) return alert("jsPDF no cargado.");
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('p', 'mm', 'a4');
        const fechas = generarFechasDiarias();
        if (fechas.length === 0) return alert("No hay fechas.");

        const semanas = [];
        let semActual = [];
        fechas.forEach((f, i) => {
            semActual.push(f);
            const sig = fechas[i + 1];
            if (f.getDay() === 0 || !sig || sig.getDay() === 1) {
                semanas.push([...semActual]);
                semActual = [];
            }
        });

        const pdfHeaders = ['DÍA', 'HORA', ...state.colOrder.map(k => COL_DEFS[k].toUpperCase())];
        let currentY = 18; // Iniciar un poco más arriba
        let ultimoMesTitulado = "";
        const PAGE_LIMIT = 285; // Maximizar el uso del borde inferior

        semanas.forEach((semana, idx) => {
            const fI = semana[0];
            const fF = semana[semana.length - 1];
            const mesI = fI.toLocaleString('es-ES', { month: 'long' }).toUpperCase();
            const anioI = fI.getFullYear();
            const mesF = fF.toLocaleString('es-ES', { month: 'long' }).toUpperCase();
            const mesSemanaKey = `${mesI} ${anioI}`;

            // ESTIMACIÓN DE ALTURA DE LA SEMANA (Ajustada para mayor precisión)
            let totalFilasCuerpo = 1; // Fila del título de la semana
            semana.forEach(f => {
                const id = f.toISOString().split('T')[0];
                const cfg = state.configDias[f.getDay()];
                const hT = (f.getDay() === 0 || f.getDay() === 6) ? (state.horario.tFin || state.horario.t) : state.horario.t;
                const extras = (state.filasExtras[id] || []).length;
                totalFilasCuerpo += (cfg.m ? 1 : 0) + (cfg.t && hT !== '' ? 1 : 0) + extras;
            });
            const estimatedHeight = (totalFilasCuerpo * 7) + 8; // Altura estimada reducida

            const esNuevoMes = mesSemanaKey !== ultimoMesTitulado;

            // Salto de página si la semana completa no cabe
            if ((currentY + estimatedHeight > PAGE_LIMIT)) {
                doc.addPage();
                currentY = 18;
            }

            // TÍTULO DEL MES (Solo si es inicio de página o cambio de mes real)
            if (currentY === 18 || esNuevoMes) {
                doc.setFont("helvetica", "bold");
                doc.setFontSize(13);
                doc.text(`PROGRAMA DE SERVICIO - ${mesSemanaKey}`, 105, currentY - 5, { align: 'center' });
                ultimoMesTitulado = mesSemanaKey;
            }

            const tituloSemana = `SEMANA DEL ${fI.getDate()} DE ${mesI} AL ${fF.getDate()} DE ${mesF}`;
            const weekRows = [];
            weekRows.push([{
                content: tituloSemana,
                colSpan: pdfHeaders.length,
                styles: { halign: 'center', fillColor: [235, 240, 245], fontStyle: 'bold', fontSize: 8.5, cellPadding: 1 }
            }]);

            semana.forEach(f => {
                const id = f.toISOString().split('T')[0];
                const d = obtenerDatosDia(id);
                const dNum = f.getDay();
                const cfg = state.configDias[dNum];
                const esFinde = (dNum === 0 || dNum === 6);
                const hM = esFinde ? (state.horario.mFin || state.horario.m) : state.horario.m;
                const hT = esFinde ? (state.horario.tFin || state.horario.t) : state.horario.t;
                const extras = state.filasExtras[id] || [];
                const showM = cfg.m;
                const showT = cfg.t && hT !== '';

                let esPrimeraFilaDelDia = true;
                const totalR = (showM ? 1 : 0) + (showT ? 1 : 0) + extras.length;

                const addRow = (hora, source, turnoLetra = '') => {
                    const row = [];
                    if (esPrimeraFilaDelDia) {
                        row.push({
                            content: `${DIAS_SEMANA[dNum].slice(0, 3).toUpperCase()} ${f.getDate()}`,
                            rowSpan: totalR,
                            styles: { fontStyle: 'bold', valign: 'middle' }
                        });
                        esPrimeraFilaDelDia = false;
                    }
                    row.push(hora || '--:--');
                    state.colOrder.forEach(k => {
                        let val = turnoLetra ? source[`${k}${turnoLetra}`] : source[k];
                        row.push(k === 'gru' ? formatearGruposParaPDF(val) : (val || "-"));
                    });
                    weekRows.push(row);
                };

                if (showM) addRow(hM, d, 'M');
                extras.filter(e => e.turno === 'M').forEach(ex => addRow(hM, ex));
                if (showT) addRow(hT, d, 'T');
                extras.filter(e => e.turno === 'T').forEach(ex => addRow(hT, ex));
            });

            doc.autoTable({
                head: [pdfHeaders],
                body: weekRows,
                startY: currentY,
                theme: 'grid',
                styles: { fontSize: 7.5, cellPadding: 1.1, halign: 'center', overflow: 'linebreak' },
                headStyles: { fillColor: [44, 62, 80], textColor: 255 },
                margin: { left: 10, right: 10 },
                rowPageBreak: 'avoid'
            });

            // ELIMINACIÓN DEL ESPACIO: La siguiente semana empieza donde terminó la anterior
            currentY = doc.lastAutoTable.finalY;
        });

        doc.save(`Programa_Servicio.pdf`);
    }

    function init() {
        loadFromStorage();
        if (!els.fInicio.value) els.fInicio.value = new Date().toISOString().split('T')[0];
        els.hManana.value = state.horario.m; els.hTarde.value = state.horario.t; els.hMananaFin.value = state.horario.mFin; els.hTardeFin.value = state.horario.tFin;
        els.diasContainer.onchange = (e) => { if (e.target.type === 'checkbox') { state.configDias[e.target.dataset.day][e.target.dataset.shift] = e.target.checked; actualizarTodo(); } };
        els.fInicio.onchange = () => ajustarFechas('combo'); els.fFinal.onchange = () => ajustarFechas('final'); els.cantFechas.onchange = () => ajustarFechas('combo');
        document.getElementById('btnAgregarConductor').onclick = () => modificarLista('conductor', 'agregar');
        document.getElementById('btnAgregarLugar').onclick = () => modificarLista('lugar', 'agregar');
        document.getElementById('btnAgregarGrupo').onclick = () => modificarLista('grupo', 'agregar');
        document.getElementById('btnLimpiarGrupos').onclick = () => modificarLista('grupo', 'borrarTodo');
        document.getElementById('btnExportPDF').onclick = exportarPDF;
        document.getElementById('btnAsignarConductores').onclick = () => asignarGenerico('conductor');
        document.getElementById('btnAsignarTerritorios').onclick = () => asignarGenerico('territorio');
        document.getElementById('btnLimpiarAsignaciones').onclick = () => { if (confirm("¿Vaciar todo?")) { generarFechasDiarias().forEach(f => { const id = f.toISOString().split('T')[0]; if (!state.bloqueados[id]) delete state.asignaciones[id]; }); actualizarTodo(); } };
        els.btnBloquear.onclick = () => { const ids = generarFechasDiarias().map(f => f.toISOString().split('T')[0]), all = ids.every(i => state.bloqueados[i]); ids.forEach(i => state.bloqueados[i] = !all); actualizarTodo(); };
        document.getElementById('btnToggleTheme').onclick = () => { const el = document.documentElement; el.setAttribute('data-theme', el.getAttribute('data-theme') === 'dark' ? 'light' : 'dark'); };
        document.getElementById('toggleSidebar').onclick = () => { const l = document.querySelector('.main-layout'), h = l.classList.toggle('sidebar-hidden'); document.getElementById('sidebarIcon').textContent = h ? "➡️" : "⬅️"; document.getElementById('sidebarText').textContent = h ? "Mostrar" : "Ocultar"; setTimeout(actualizarTodo, 300); };
        document.getElementById('btnFijarHorario').onclick = () => { state.horario = { m: els.hManana.value, t: els.hTarde.value, mFin: els.hMananaFin.value, tFin: els.hTardeFin.value }; actualizarTodo(); };
        document.querySelectorAll('.tab-button').forEach(btn => btn.onclick = () => { document.querySelectorAll('.tab-button').forEach(b => b.classList.remove('active')); btn.classList.add('active'); document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active')); document.getElementById(btn.dataset.tab + 'Pane').classList.add('active'); });
        ajustarFechas('combo');
    }
    document.addEventListener('DOMContentLoaded', init);
})();