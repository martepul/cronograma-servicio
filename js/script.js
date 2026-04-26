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
        EXTRAS: 'filas_extras_v2',
        COMENTARIOS: 'comentarios_v1',
        START_WEEK: 'start_week_v1',
        THEME: 'theme_v1'
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

    // --- FORMATEO DE GRUPOS ---
    const formatearGruposParaPDF = (val) => {
        if (!val || val === "-" || val === "") return "";
        const listaRaw = val.split(' + ');
        const listaLimpia = listaRaw
            .map(g => g.replace(/Grupo\s+/gi, '').trim())
            .filter(g => g !== "");

        if (listaLimpia.length === 0) return "";
        if (listaLimpia.length === 1) return "Grupo " + listaLimpia[0];
        const ultimo = listaLimpia.pop();
        return "Grupos " + listaLimpia.join(", ") + " y " + ultimo;
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
        colOrder: ['lug', 'ter', 'zon', 'cond'],
        configDias: {},
        comentarios: {},
        startOfWeek: 1,
        uiState: { openGroups: {} }
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
            const rawCond = JSON.parse(localStorage.getItem(KEYS.CONDUCTORES)) || [];
            state.conductores = rawCond.map(c => {
                if (typeof c === 'string') return { name: c, tags: [] };
                return { ...c, tags: Array.isArray(c.tags) ? c.tags.map(Number) : [] };
            });

            state.lugares = JSON.parse(localStorage.getItem(KEYS.LUGARES)) || INITIAL_LUGARES;
            state.grupos = JSON.parse(localStorage.getItem(KEYS.GRUPOS)) || [];
            state.mapaCoherencia = JSON.parse(localStorage.getItem(KEYS.MAPA)) || INITIAL_MAPA;
            state.manuales = JSON.parse(localStorage.getItem(KEYS.MANUALES)) || {};
            state.bloqueados = JSON.parse(localStorage.getItem(KEYS.BLOQUEADOS)) || {};
            state.horario = JSON.parse(localStorage.getItem(KEYS.HORARIO)) || { m: '', t: '', mFin: '', tFin: '' };
            state.filasExtras = JSON.parse(localStorage.getItem(KEYS.EXTRAS)) || {};
            state.comentarios = JSON.parse(localStorage.getItem(KEYS.COMENTARIOS)) || {};

            const storedOrder = JSON.parse(localStorage.getItem(KEYS.CONFIG_ORDER));
            if (storedOrder) state.colOrder = storedOrder;
            const storedDias = JSON.parse(localStorage.getItem(KEYS.CONFIG_DIAS));
            state.configDias = storedDias || {};
            if (!storedDias) { for (let i = 0; i < 7; i++) state.configDias[i] = { m: true, t: (i !== 0) }; }

            const savedStart = localStorage.getItem(KEYS.START_WEEK);
            state.startOfWeek = savedStart !== null ? parseInt(savedStart) : 1;
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
        localStorage.setItem(KEYS.COMENTARIOS, JSON.stringify(state.comentarios));
        localStorage.setItem(KEYS.START_WEEK, state.startOfWeek);
    }

    // --- IMPORTACIÓN ---
    function importarDesdeArchivo(evento, tipo) {
        const archivo = evento.target.files[0];
        if (!archivo) return;
        const lector = new FileReader();
        lector.onload = function (e) {
            const contenido = e.target.result;
            const lineas = contenido.split(/\r?\n/);
            let agregados = 0;
            lineas.forEach(linea => {
                const lineaLimpia = linea.trim();
                if (lineaLimpia === "") return;
                const partes = lineaLimpia.split(',').map(p => p.trim()).filter(p => p !== "");
                if (tipo === 'conductor') {
                    const nombre = partes[0];
                    if (nombre && !state.conductores.some(c => c.name === nombre)) {
                        state.conductores.push({ name: nombre, tags: [] });
                        agregados++;
                    }
                } else if (tipo === 'lugar') {
                    const nombre = partes[0];
                    if (nombre) {
                        if (!state.lugares.includes(nombre)) { state.lugares.push(nombre); agregados++; }
                        const territorios = [], zonas = [];
                        partes.slice(1).forEach(item => {
                            if (item.includes('-')) { if (!zonas.includes(item)) zonas.push(item); }
                            else { if (!territorios.includes(item)) territorios.push(item); }
                        });
                        state.mapaCoherencia[nombre] = { ter: territorios, zon: zonas };
                    }
                }
            });
            if (agregados > 0) {
                state.conductores.sort((a, b) => a.name.localeCompare(b.name));
                state.lugares.sort((a, b) => a.localeCompare(b));
                actualizarTodo();
                alert(`¡Éxito! Se procesaron ${agregados} elementos.`);
            } else alert("No se añadieron elementos nuevos.");
            evento.target.value = "";
        };
        lector.readAsText(archivo);
    }

    // --- ASIGNACIÓN ---
    function asignarGenerico(tipo, targetId = null) {
        if (tipo === 'conductor' && state.conductores.length === 0) return alert("Faltan conductores.");
        if (tipo === 'territorio' && state.lugares.length === 0) return alert("Faltan lugares.");
        const fechas = generarFechasDiarias();
        const usaZona = state.colOrder.includes('zon');
        const obtenerUsoEnUltimoMes = (fechaTargetStr) => {
            const usoTer = {}, usoZon = {};
            const targetDate = stringToDate(fechaTargetStr);
            const limiteDate = new Date(targetDate);
            limiteDate.setDate(limiteDate.getDate() - 30);
            const todasLasFechas = new Set([...Object.keys(state.asignaciones), ...Object.keys(state.manuales), ...Object.keys(state.filasExtras)]);
            todasLasFechas.forEach(idDia => {
                const dDate = stringToDate(idDia);
                if (dDate >= limiteDate && dDate < targetDate) {
                    const datos = obtenerDatosDia(idDia);
                    ['M', 'T'].forEach(t => {
                        if (datos[`ter${t}`]) usoTer[datos[`ter${t}`]] = (usoTer[datos[`ter${t}`]] || 0) + 1;
                        if (datos[`zon${t}`]) usoZon[datos[`zon${t}`]] = (usoZon[datos[`zon${t}`]] || 0) + 1;
                    });
                    (state.filasExtras[idDia] || []).forEach(ex => {
                        if (ex.ter) usoTer[ex.ter] = (usoTer[ex.ter] || 0) + 1;
                        if (ex.zon) usoZon[ex.zon] = (usoZon[ex.zon] || 0) + 1;
                    });
                }
            });
            return { usoTer, usoZon };
        };
        fechas.forEach(f => {
            const id = f.toISOString().split('T')[0];
            if ((targetId && id !== targetId) || state.bloqueados[id]) return;
            if (!state.asignaciones[id]) state.asignaciones[id] = {};
            const asig = state.asignaciones[id];
            if (tipo === 'conductor') { delete asig.condM; delete asig.condT; }
            else { ['lug', 'ter', 'zon'].forEach(k => { delete asig[k + 'M']; delete asig[k + 'T']; }); }
            const diaNum = f.getDay(), configDia = state.configDias[diaNum];
            const esFinde = (diaNum === 0 || diaNum === 6), hT = esFinde ? (state.horario.tFin || state.horario.t) : state.horario.t;
            const usarTarde = configDia.t && hT !== '', usadosHoy = [];
            ['M', 'T'].forEach(turno => {
                if ((turno === 'M' && !configDia.m) || (turno === 'T' && !usarTarde)) return;
                if (tipo === 'conductor') {
                    const manual = state.manuales[id]?.[`cond${turno}`];
                    if (manual) usadosHoy.push(manual);
                    else {
                        const pool = state.conductores.filter(c => !usadosHoy.includes(c.name) && (c.tags.length === 0 || c.tags.some(t => Number(t) === diaNum)));
                        const elegido = pool[Math.floor(Math.random() * pool.length)];
                        if (elegido) { asig[`cond${turno}`] = elegido.name; usadosHoy.push(elegido.name); }
                    }
                } else {
                    const historial = obtenerUsoEnUltimoMes(id), manualLug = state.manuales[id]?.[`lug${turno}`];
                    let lugAct = manualLug;
                    if (!lugAct) {
                        const disponibles = state.lugares.filter(l => !usadosHoy.includes(l));
                        let combinaciones = [];
                        disponibles.forEach(l => {
                            const cfg = state.mapaCoherencia[l];
                            if (!cfg || cfg.ter.length === 0) { combinaciones.push({ lug: l, ter: null, zon: null, score: 0 }); return; }
                            cfg.ter.forEach(t => {
                                if (usaZona && cfg.zon.length > 0) {
                                    const zTer = cfg.zon.filter(z => z.startsWith(t + '-'));
                                    (zTer.length > 0 ? zTer : [null]).forEach(z => combinaciones.push({ lug: l, ter: t, zon: z, score: (historial.usoTer[t] || 0) + (z ? historial.usoZon[z] || 0 : 0) }));
                                } else combinaciones.push({ lug: l, ter: t, zon: null, score: historial.usoTer[t] || 0 });
                            });
                        });
                        if (combinaciones.length > 0) {
                            const min = Math.min(...combinaciones.map(c => c.score)), mejores = combinaciones.filter(c => c.score === min);
                            const el = mejores[Math.floor(Math.random() * mejores.length)];
                            asig[`lug${turno}`] = el.lug; if (el.ter) asig[`ter${turno}`] = el.ter; if (usaZona && el.zon) asig[`zon${turno}`] = el.zon;
                            lugAct = el.lug; usadosHoy.push(lugAct);
                        }
                    } else usadosHoy.push(lugAct);
                }
            });
        });
        actualizarTodo();
    }

    function ajustarFechas(modo) {
        if (modo === 'combo') {
            if (els.cantFechas.value === "custom") return;
            const inicioStr = els.fInicio.value; if (!inicioStr) return;
            let fFin = stringToDate(inicioStr); fFin.setDate(fFin.getDate() + (parseInt(els.cantFechas.value, 10) - 1));
            els.fFinal.value = fFin.toISOString().split('T')[0];
        } else if (modo === 'final') {
            const inicio = stringToDate(els.fInicio.value), final = stringToDate(els.fFinal.value);
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
            if (tipo === 'conductor') { if (!lista.some(c => c.name === raw)) { lista.push({ name: raw, tags: [] }); input.value = ''; actualizarTodo(); } return; }
            if (tipo === 'grupo') {
                const range = raw.match(/^(\d+)-(\d+)$/);
                if (range) { for (let i = parseInt(range[1]); i <= parseInt(range[2]); i++) { if (!lista.includes(`Grupo ${i}`)) lista.push(`Grupo ${i}`); } input.value = ''; actualizarTodo(); return; }
            }
            if (!lista.includes(raw)) {
                if (tipo === 'lugar') {
                    const t = prompt("Territorios:") || "", z = prompt("Zonas:") || "";
                    state.mapaCoherencia[raw] = { ter: t.split(',').map(s => s.trim()).filter(s => s), zon: z.split(',').map(s => s.trim()).filter(s => s) };
                }
                lista.push(raw); input.value = ''; actualizarTodo();
            }
        } else if (accion === 'eliminar') {
            const itemName = (tipo === 'conductor') ? lista[indice].name : lista[indice];
            if (confirm(`¿Eliminar "${itemName}"?`)) { if (tipo === 'lugar') delete state.mapaCoherencia[lista[indice]]; lista.splice(indice, 1); actualizarTodo(); }
        } else if (accion === 'borrarTodo') { if (confirm("¿Borrar todos?")) { state.grupos = []; actualizarTodo(); } }
    }

    function guardarManual(fechaId, valor, rol) {
        if (state.bloqueados[fechaId]) return;
        if (!state.manuales[fechaId]) state.manuales[fechaId] = {};
        if (!valor && valor !== 0) delete state.manuales[fechaId][rol];
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
        const container = document.createElement('div'); container.style.cssText = 'display:flex; flex-direction:column; gap:2px; width:100%';
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
            input.onchange = (e) => { if (e.target.checked) { if (!state.colOrder.includes(c.id)) state.colOrder.push(c.id); } else state.colOrder = state.colOrder.filter(k => k !== c.id); actualizarTodo(); };
            label.append(input, ' ' + c.lbl); grid.appendChild(label);
        });
    }

    function renderSelectorDias() {
        if (!els.diasContainer) return; els.diasContainer.innerHTML = '';
        let orden = []; let start = state.startOfWeek;
        for (let i = 0; i < 7; i++) orden.push((start + i) % 7);
        orden.forEach(i => {
            const row = document.createElement('div'); row.className = 'day-config-row';
            row.innerHTML = `<span style="width:80px">${DIAS_SEMANA[i]}</span>
                <label>Mañana <input type="checkbox" data-day="${i}" data-shift="m" ${state.configDias[i].m ? 'checked' : ''}></label>
                <label>Tarde <input type="checkbox" data-day="${i}" data-shift="t" ${state.configDias[i].t ? 'checked' : ''}></label>`;
            els.diasContainer.appendChild(row);
        });
    }

    function renderListas() {
        const draw = (l, container, tipo) => {
            container.innerHTML = '';
            l.forEach((item, i) => {
                const li = document.createElement('li'); if (tipo === 'conductor') li.className = 'participant-item-complex';
                const spanName = document.createElement('span'); spanName.textContent = (tipo === 'conductor') ? item.name : item;
                li.appendChild(spanName);
                if (tipo === 'conductor') {
                    const tagCont = document.createElement('div'); tagCont.className = 'tag-container';
                    ['D', 'L', 'M', 'Mi', 'J', 'V', 'S'].forEach((letra, idx) => {
                        const btn = document.createElement('button'); btn.textContent = letra; btn.className = `btn-tag ${item.tags.some(t => Number(t) === idx) ? 'active' : ''}`;
                        btn.onclick = () => { if (item.tags.some(t => Number(t) === idx)) item.tags = item.tags.filter(t => Number(t) !== idx); else item.tags.push(Number(idx)); actualizarTodo(); };
                        tagCont.appendChild(btn);
                    });
                    li.appendChild(tagCont);
                }
                const btn = document.createElement('button'); btn.textContent = '×'; btn.onclick = () => modificarLista(tipo, 'eliminar', i);
                li.appendChild(btn); container.appendChild(li);
            });
        };
        draw(state.conductores, els.listaCond, 'conductor'); draw(state.lugares, els.listaLugares, 'lugar'); draw(state.grupos, els.listaGrupos, 'grupo');
    }

    function renderTabla() {
        els.headerRow.innerHTML = '<th>Fecha</th><th>Hora</th>';
        state.colOrder.forEach(k => { const th = document.createElement('th'); th.textContent = COL_DEFS[k]; els.headerRow.appendChild(th); });
        els.headerRow.appendChild(Object.assign(document.createElement('th'), { className: "th-actions", textContent: "Acciones" }));
        els.tablaBody.innerHTML = '';
        generarFechasDiarias().forEach(f => {
            const id = f.toISOString().split('T')[0], locked = state.bloqueados[id], dNum = f.getDay(), cfg = state.configDias[dNum];
            const esFinde = (dNum === 0 || dNum === 6), hM = esFinde ? (state.horario.mFin || state.horario.m) : state.horario.m, hT = esFinde ? (state.horario.tFin || state.horario.t) : state.horario.t;
            const showM = cfg.m, showT = cfg.t && hT !== '', extras = state.filasExtras[id] || [], datos = obtenerDatosDia(id);
            if (state.comentarios[id]) {
                const trCom = document.createElement('tr'); trCom.className = 'comment-row';
                const tdCom = document.createElement('td'); tdCom.colSpan = 2 + state.colOrder.length + 1; tdCom.className = 'comment-cell';
                tdCom.innerHTML = `<span>${state.comentarios[id].toUpperCase()}</span>`;
                if (!locked) { const b = document.createElement('button'); b.className = 'btn-delete-note'; b.innerHTML = '&times;'; b.onclick = () => { if (confirm("¿Eliminar?")) { delete state.comentarios[id]; actualizarTodo(); } }; tdCom.appendChild(b); }
                trCom.appendChild(tdCom); els.tablaBody.appendChild(trCom);
            }
            const tr = document.createElement('tr'); if (locked) tr.style.opacity = "0.7";
            const tdF = tr.insertCell(); tdF.innerHTML = `<b>${DIAS_SEMANA[dNum].slice(0, 3)}</b><br>${f.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit' })}`;
            tdF.rowSpan = 1 + extras.length; tdF.style.verticalAlign = 'top';
            const tdH = tr.insertCell(); tdH.innerHTML = `<div class="stacked-cell">${showM ? `<div class="time-text ${showT ? 'border-bottom' : ''}">${hM || '--:--'}</div>` : ''}${showT ? `<div class="time-text">${hT}</div>` : ''}</div>`;
            state.colOrder.forEach(tipo => {
                const td = tr.insertCell(); const cont = document.createElement('div'); cont.className = 'stacked-cell';
                ['M', 'T'].forEach(turno => {
                    if ((turno === 'M' && !showM) || (turno === 'T' && !showT)) return;
                    if (tipo === 'lug') {
                        const w = document.createElement('div'); w.className = 'lug-gru-wrapper'; const key = `${id}-${turno}`;
                        const b = document.createElement('button'); b.className = 'btn-add-gru'; const s = document.createElement('select'); s.disabled = locked; s.add(new Option("-", ""));
                        state.lugares.forEach(o => s.add(new Option(o, o, false, datos[`lug${turno}`] === o)));
                        s.onchange = (e) => guardarManual(id, e.target.value, `lug${turno}`);
                        const gc = document.createElement('div'); gc.className = 'gru-container'; const open = !!state.uiState.openGroups[key]; gc.style.display = open ? 'block' : 'none'; b.textContent = open ? '-' : '+';
                        gc.appendChild(crearSelectorGrupos(datos[`gru${turno}`], locked, (v) => guardarManual(id, v, `gru${turno}`)));
                        b.onclick = () => { state.uiState.openGroups[key] = !state.uiState.openGroups[key]; actualizarTodo(); };
                        w.append(b, s, gc); cont.appendChild(w);
                    } else if (tipo === 'gru') cont.appendChild(crearSelectorGrupos(datos[`gru${turno}`], locked, (v) => guardarManual(id, v, `gru${turno}`)));
                    else if (tipo === 'cua') { const i = Object.assign(document.createElement('input'), { type: 'text', className: 'input-cuadra', value: datos[`cua${turno}`] || "", disabled: locked }); i.onchange = (e) => guardarManual(id, e.target.value, `cua${turno}`); cont.appendChild(i); }
                    else {
                        const s = document.createElement('select'); s.disabled = locked; s.add(new Option("-", ""));
                        let ops = (tipo === 'cond') ? state.conductores.filter(c => c.tags.length === 0 || c.tags.some(t => Number(t) === dNum)).map(c => c.name) : (tipo === 'ter') ? state.mapaCoherencia[datos[`lug${turno}`]]?.ter || getRangoTerritorios() : state.mapaCoherencia[datos[`lug${turno}`]]?.zon.filter(z => z.startsWith(datos[`ter${turno}`] + '-')) || getRangoZonas();
                        if (tipo === 'cond' && datos[`cond${turno}`] && !ops.includes(datos[`cond${turno}`])) ops.push(datos[`cond${turno}`]);
                        ops.forEach(o => s.add(new Option(o, o, false, datos[`${tipo}${turno}`] === o)));
                        s.onchange = (e) => guardarManual(id, e.target.value, `${tipo}${turno}`); cont.appendChild(s);
                    }
                });
                td.appendChild(cont);
            });
            const tdA = tr.insertCell(); tdA.innerHTML = `<div class="actions-wrapper"><div class="main-actions"><button class="btn-lock">${locked ? '🔒' : '🔓'}</button><button class="btn-assign" ${locked ? 'disabled' : ''}>↻</button><button class="btn-note">📝</button></div><div class="extra-actions">${showM ? '<button class="btn-add-m">+M</button>' : ''}${showT ? '<button class="btn-add-t">+T</button>' : ''}</div></div>`;
            tdA.querySelector('.btn-lock').onclick = () => { state.bloqueados[id] = !state.bloqueados[id]; actualizarTodo(); };
            tdA.querySelector('.btn-assign').onclick = () => { asignarGenerico('conductor', id); asignarGenerico('territorio', id); };
            tdA.querySelector('.btn-note').onclick = () => { const n = prompt("Nota:", state.comentarios[id] || ""); if (n !== null) { if (n.trim() === "") delete state.comentarios[id]; else state.comentarios[id] = n; actualizarTodo(); } };
            if (showM) tdA.querySelector('.btn-add-m').onclick = () => agregarFilaExtra(id, 'M');
            if (showT) tdA.querySelector('.btn-add-t').onclick = () => agregarFilaExtra(id, 'T');
            els.tablaBody.appendChild(tr);
            extras.forEach(ex => {
                const trx = document.createElement('tr'); if (locked) trx.style.opacity = "0.7";
                const tdExH = trx.insertCell(); tdExH.innerHTML = `<span style="border-left:3px solid ${ex.turno === 'M' ? '#2196F3' : '#FF9800'}; padding-left:5px; font-weight:bold;">${(ex.turno === 'M' ? hM : hT) || '--:--'}</span>`;
                state.colOrder.forEach(t => {
                    const tdEx = trx.insertCell();
                    if (t === 'lug') {
                        const w = document.createElement('div'); w.className = 'lug-gru-wrapper'; const b = document.createElement('button'); b.className = 'btn-add-gru'; b.textContent = !!state.uiState.openGroups[ex.id] ? '-' : '+';
                        const s = document.createElement('select'); s.disabled = locked; s.add(new Option("-", ""));
                        state.lugares.forEach(o => s.add(new Option(o, o, false, ex.lug === o)));
                        s.onchange = (e) => guardarExtra(id, ex.id, 'lug', e.target.value);
                        const gc = document.createElement('div'); gc.className = 'gru-container'; gc.style.display = !!state.uiState.openGroups[ex.id] ? 'block' : 'none';
                        gc.appendChild(crearSelectorGrupos(ex.gru, locked, (v) => guardarExtra(id, ex.id, 'gru', v)));
                        b.onclick = () => { state.uiState.openGroups[ex.id] = !state.uiState.openGroups[ex.id]; actualizarTodo(); };
                        w.append(b, s, gc); tdEx.appendChild(w);
                    } else if (t === 'gru') tdEx.appendChild(crearSelectorGrupos(ex[t], locked, (v) => guardarExtra(id, ex.id, t, v)));
                    else if (t === 'cua') { const i = Object.assign(document.createElement('input'), { type: 'text', className: 'input-cuadra', value: ex[t] || "", disabled: locked }); i.onchange = (e) => guardarExtra(id, ex.id, t, e.target.value); tdEx.appendChild(i); }
                    else {
                        const s = document.createElement('select'); s.disabled = locked; s.add(new Option("-", ""));
                        let ops = (t === 'cond') ? state.conductores.filter(c => c.tags.length === 0 || c.tags.some(tag => Number(tag) === dNum)).map(c => c.name) : (t === 'ter') ? state.mapaCoherencia[ex.lug]?.ter || getRangoTerritorios() : state.mapaCoherencia[ex.lug]?.zon.filter(z => z.startsWith(ex.ter + '-')) || getRangoZonas();
                        ops.forEach(o => s.add(new Option(o, o, false, ex[t] === o)));
                        s.onchange = (e) => guardarExtra(id, ex.id, t, e.target.value); tdEx.appendChild(s);
                    }
                });
                const tdAx = trx.insertCell(); const bDel = document.createElement('button'); bDel.innerHTML = 'Eliminar &times;'; bDel.style.color = 'red'; bDel.disabled = locked; bDel.onclick = () => eliminarFilaExtra(id, ex.id);
                tdAx.appendChild(bDel); els.tablaBody.appendChild(trx);
            });
        });
        const ids = generarFechasDiarias().map(f => f.toISOString().split('T')[0]);
        els.btnBloquear.innerHTML = (ids.length > 0 && ids.every(i => state.bloqueados[i])) ? ' Desbloquear Todo' : ' Bloquear Todo';
    }

    function actualizarTodo() { renderListas(); renderColumnSelector(); renderSelectorDias(); renderTabla(); saveToStorage(); }

    function exportarPDF() {
        const formatTime = (t) => { if (!t || !t.includes(':')) return '--:--'; let [h, m] = t.split(':').map(Number); const ampm = h >= 12 ? 'pm' : 'am'; h = h % 12 || 12; return `${h}:${m < 10 ? '0' + m : m} ${ampm}`; };
        const { jsPDF } = window.jspdf; const doc = new jsPDF('p', 'mm', 'a4'); const fechas = generarFechasDiarias(); if (fechas.length === 0) return alert("Sin fechas.");
        const semanas = []; let sem = []; fechas.forEach((f, i) => { sem.push(f); const sig = fechas[i + 1]; if (!sig || sig.getDay() === state.startOfWeek) { semanas.push([...sem]); sem = []; } });
        const headers = ['DÍA', 'HORA', ...state.colOrder.map(k => COL_DEFS[k].toUpperCase())];
        let currentY = 12; doc.setFont("helvetica", "bold"); doc.setFontSize(12); doc.text(`PROGRAMA DE SERVICIO - ${fechas[0].toLocaleString('es-ES', { month: 'long' }).toUpperCase()} ${fechas[0].getFullYear()}`, 105, currentY, { align: 'center' });
        currentY += 4; const rows = [];
        semanas.forEach(s => {
            rows.push([{ content: `SEMANA DEL ${s[0].getDate()} AL ${s[s.length - 1].getDate()} DE ${s[s.length - 1].toLocaleString('es-ES', { month: 'long' }).toUpperCase()}`, colSpan: headers.length, styles: { halign: 'center', fillColor: [235, 240, 245], fontStyle: 'bold', fontSize: 7.5 } }]);
            s.forEach(f => {
                const id = f.toISOString().split('T')[0]; if (state.comentarios[id]) rows.push([{ content: state.comentarios[id].toUpperCase(), colSpan: headers.length, styles: { halign: 'center', fillColor: [255, 253, 208], fontStyle: 'bolditalic', fontSize: 7 } }]);
                const d = obtenerDatosDia(id), dNum = f.getDay(), cfg = state.configDias[dNum], esF = (dNum === 0 || dNum === 6), hM = esF ? (state.horario.mFin || state.horario.m) : state.horario.m, hT = esF ? (state.horario.tFin || state.horario.t) : state.horario.t;
                const ex = state.filasExtras[id] || [], sM = cfg.m, sT = cfg.t && hT !== '';
                let first = true; const total = (sM ? 1 : 0) + (sT ? 1 : 0) + ex.length;
                const addR = (hora, src, turno = '') => {
                    const r = []; if (first) { r.push({ content: `${DIAS_SEMANA[dNum].slice(0, 3).toUpperCase()} ${f.getDate()}`, rowSpan: total, styles: { fontStyle: 'bold', valign: 'middle' } }); first = false; }
                    r.push(formatTime(hora));
                    state.colOrder.forEach(k => {
                        let val = turno ? src[`${k}${turno}`] : src[k];
                        if (k === 'lug') { const g = formatearGruposParaPDF(turno ? src[`gru${turno}`] : src['gru']); val = g ? `${g} | ${val || "-"}` : (val || "-"); }
                        else if (k === 'gru') val = formatearGruposParaPDF(val) || "-";
                        else val = val || "-"; r.push(val);
                    }); rows.push(r);
                };
                if (sM) addR(hM, d, 'M'); ex.filter(e => e.turno === 'M').forEach(x => addR(hM, x));
                if (sT) addR(hT, d, 'T'); ex.filter(e => e.turno === 'T').forEach(x => addR(hT, x));
            });
        });
        doc.autoTable({ head: [headers], body: rows, startY: currentY, theme: 'grid', styles: { fontSize: 8.5, fontStyle: 'bold', cellPadding: 0.8, halign: 'center' }, headStyles: { fillColor: [44, 62, 80], textColor: 255 }, margin: { left: 8, right: 8 }, rowPageBreak: 'avoid' });
        doc.save(`Programa_Servicio.pdf`);
    }

    function init() {
        loadFromStorage();
        const btnTheme = document.getElementById('btnToggleTheme'), systemDark = window.matchMedia('(prefers-color-scheme: dark)');
        const setTheme = (dark) => { if (dark) { document.documentElement.setAttribute('data-theme', 'dark'); localStorage.setItem(KEYS.THEME, 'dark'); } else { document.documentElement.removeAttribute('data-theme'); localStorage.setItem(KEYS.THEME, 'light'); } };
        const savedTheme = localStorage.getItem(KEYS.THEME); setTheme(savedTheme === 'dark' || (!savedTheme && systemDark.matches));
        if (btnTheme) btnTheme.onclick = () => setTheme(document.documentElement.getAttribute('data-theme') !== 'dark');
        systemDark.addEventListener('change', (e) => setTheme(e.matches));
        if (!els.fInicio.value) els.fInicio.value = new Date().toISOString().split('T')[0];
        els.hManana.value = state.horario.m; els.hTarde.value = state.horario.t; els.hMananaFin.value = state.horario.mFin; els.hTardeFin.value = state.horario.tFin;
        els.diasContainer.onchange = (e) => { if (e.target.type === 'checkbox') { state.configDias[e.target.dataset.day][e.target.dataset.shift] = e.target.checked; actualizarTodo(); } };
        els.fInicio.onchange = () => ajustarFechas('combo'); els.fFinal.onchange = () => ajustarFechas('final'); els.cantFechas.onchange = () => ajustarFechas('combo');
        document.getElementById('btnAgregarConductor').onclick = () => modificarLista('conductor', 'agregar');
        document.getElementById('btnAgregarLugar').onclick = () => modificarLista('lugar', 'agregar');
        document.getElementById('btnAgregarGrupo').onclick = () => modificarLista('grupo', 'agregar');
        document.getElementById('btnLimpiarGrupos').onclick = () => modificarLista('grupo', 'borrarTodo');
        document.querySelectorAll('.btn-export').forEach(btn => btn.onclick = exportarPDF);
        const inC = document.getElementById('fileConductores'); if (inC) inC.addEventListener('change', (e) => importarDesdeArchivo(e, 'conductor'));
        const inL = document.getElementById('fileLugares'); if (inL) inL.addEventListener('change', (e) => importarDesdeArchivo(e, 'lugar'));
        document.getElementById('btnAsignarConductores').onclick = () => asignarGenerico('conductor');
        document.getElementById('btnAsignarTerritorios').onclick = () => asignarGenerico('territorio');
        document.getElementById('btnLimpiarAsignaciones').onclick = () => { if (confirm("¿Vaciar?")) { generarFechasDiarias().forEach(f => { const id = f.toISOString().split('T')[0]; if (!state.bloqueados[id]) delete state.asignaciones[id]; }); actualizarTodo(); } };
        els.btnBloquear.onclick = () => { const ids = generarFechasDiarias().map(f => f.toISOString().split('T')[0]), all = ids.every(i => state.bloqueados[i]); ids.forEach(i => state.bloqueados[i] = !all); actualizarTodo(); };

        // --- LÓGICA DE SIDEBAR REFINADA ---
        const layout = document.querySelector('.main-layout');
        const sidebar = document.querySelector('.sidebar');
        const sIcon = document.getElementById('sidebarIcon');
        const sText = document.getElementById('sidebarText');

        const updateUI = () => {
            const isH = layout.classList.contains('sidebar-hidden'), isM = window.innerWidth <= 850;
            sIcon.textContent = isM ? (isH ? "⚙️" : "❌") : (isH ? "➡️" : "⬅️");
            sText.textContent = isM ? (isH ? "Configuración" : "Cerrar") : (isH ? "Mostrar" : "Ocultar");
        };

        // Botón principal (Abrir/Cerrar)
        document.getElementById('toggleSidebar').onclick = (e) => {
            e.stopPropagation();
            layout.classList.toggle('sidebar-hidden');
            updateUI();
            setTimeout(actualizarTodo, 300);
        };

        // Cerrar al tocar área vacía del panel (en móvil)
        if (sidebar) {
            sidebar.onclick = (e) => {
                const isMobile = window.innerWidth <= 850;
                // Si el clic NO fue en un botón, input, select o etiqueta de configuración
                const clickedControl = e.target.closest('button, input, select, label, .tab-button, .btn-tag');

                if (isMobile && !clickedControl) {
                    layout.classList.add('sidebar-hidden');
                    updateUI();
                }
            };
        }

        window.onresize = updateUI;
        if (window.innerWidth <= 850) layout.classList.add('sidebar-hidden');
        updateUI();

        document.getElementById('btnFijarHorario').onclick = () => { state.horario = { m: els.hManana.value, t: els.hTarde.value, mFin: els.hMananaFin.value, tFin: els.hTardeFin.value }; actualizarTodo(); };
        document.querySelectorAll('.tab-button').forEach(btn => btn.onclick = () => {
            document.querySelectorAll('.tab-button, .tab-pane').forEach(el => el.classList.remove('active'));
            btn.classList.add('active'); document.getElementById(btn.dataset.tab + 'Pane').classList.add('active');
        });
        const selS = document.getElementById('inicioSemanaSelect'); if (selS) { selS.value = state.startOfWeek; selS.onchange = (e) => { state.startOfWeek = parseInt(e.target.value); actualizarTodo(); }; }
        ajustarFechas('combo');
    }
    document.addEventListener('DOMContentLoaded', init);
})();
