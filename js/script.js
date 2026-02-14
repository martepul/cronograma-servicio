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

    // --- NUEVA LÓGICA DE FORMATEO DE GRUPOS ---
    const formatearGruposParaPDF = (val) => {
        if (!val || val === "-" || val === "") return "";

        // Separar por el delimitador de almacenamiento
        const listaRaw = val.split(' + ');

        // Limpiar la palabra "Grupo" para quedarse solo con el número/nombre
        // Ej: "Grupo 1" -> "1", "grupo A" -> "A"
        const listaLimpia = listaRaw
            .map(g => g.replace(/Grupo\s+/gi, '').trim())
            .filter(g => g !== "");

        if (listaLimpia.length === 0) return "";

        // Caso: Un solo grupo
        if (listaLimpia.length === 1) return "Grupo " + listaLimpia[0];

        // Caso: Múltiples grupos (1, 2 y 3)
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
            state.conductores = JSON.parse(localStorage.getItem(KEYS.CONDUCTORES)) || [];
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
            const selStart = document.getElementById('inicioSemanaSelect');
            if (selStart) selStart.value = state.startOfWeek;

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

        // Si el valor está vacío, eliminamos la entrada para permitir reasignación automática si aplica
        if (!valor && valor !== 0) {
            delete state.manuales[fechaId][rol];
        } else {
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

    // Permite seleccionar múltiples grupos
    function crearSelectorGrupos(valorActual, locked, onChangeCallback) {
        const container = document.createElement('div');
        container.style.cssText = 'display:flex; flex-direction:column; gap:2px; width:100%';
        let seleccionados = valorActual ? valorActual.split(' + ') : [""];

        // Renderizar selectores existentes
        seleccionados.forEach((val, idx) => {
            const row = document.createElement('div'); row.style.display = 'flex'; row.style.gap = '2px';
            const sel = document.createElement('select'); sel.style.flex = '1'; sel.add(new Option("-", ""));
            state.grupos.forEach(g => sel.add(new Option(g, g, false, g === val)));
            sel.value = val; sel.disabled = locked;
            sel.onchange = (e) => {
                seleccionados[idx] = e.target.value;
                onChangeCallback(seleccionados.filter(v => v !== "").join(' + '));
            };
            row.appendChild(sel);
            if (!locked && seleccionados.length > 1) {
                const btn = document.createElement('button'); btn.textContent = '×'; btn.style.color = 'red';
                btn.onclick = () => { seleccionados.splice(idx, 1); onChangeCallback(seleccionados.filter(v => v !== "").join(' + ')); };
                row.appendChild(btn);
            }
            container.appendChild(row);
        });

        // Botón para añadir otro grupo
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
        if (!els.diasContainer) return;
        els.diasContainer.innerHTML = '';

        let ordenDias = [];
        let start = state.startOfWeek;
        for (let i = 0; i < 7; i++) {
            ordenDias.push((start + i) % 7);
        }

        ordenDias.forEach((i) => {
            const n = DIAS_SEMANA[i];
            const row = document.createElement('div');
            row.className = 'day-config-row';
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
        state.colOrder.forEach(k => {
            const th = document.createElement('th');
            th.textContent = COL_DEFS[k];
            els.headerRow.appendChild(th);
        });
        els.headerRow.appendChild(Object.assign(document.createElement('th'), { className: "th-actions", textContent: "Acciones" }));

        els.tablaBody.innerHTML = '';
        const fechas = generarFechasDiarias();

        fechas.forEach(f => {
            const id = f.toISOString().split('T')[0];
            const locked = state.bloqueados[id];
            const dNum = f.getDay();
            const cfg = state.configDias[dNum];

            const esFinde = (dNum === 0 || dNum === 6);
            const hM = esFinde ? (state.horario.mFin || state.horario.m) : state.horario.m;
            const hT = esFinde ? (state.horario.tFin || state.horario.t) : state.horario.t;

            const showM = cfg.m;
            const showT = cfg.t && hT !== '';
            const extras = state.filasExtras[id] || [];
            const datos = obtenerDatosDia(id);

            const comentario = state.comentarios[id];
            if (comentario) {
                const trCom = document.createElement('tr');
                trCom.className = 'comment-row';
                const tdCom = document.createElement('td');
                tdCom.colSpan = 2 + state.colOrder.length + 1;
                tdCom.className = 'comment-cell';
                const spanTexto = document.createElement('span');
                spanTexto.textContent = comentario.toUpperCase();
                tdCom.appendChild(spanTexto);

                if (!locked) {
                    const btnDelNote = document.createElement('button');
                    btnDelNote.className = 'btn-delete-note';
                    btnDelNote.innerHTML = '&times;';
                    btnDelNote.title = "Eliminar nota";
                    btnDelNote.onclick = (e) => {
                        e.stopPropagation();
                        if (confirm("¿Eliminar esta nota?")) {
                            delete state.comentarios[id];
                            actualizarTodo();
                        }
                    };
                    tdCom.appendChild(btnDelNote);
                }
                trCom.appendChild(tdCom);
                els.tablaBody.appendChild(trCom);
            }

            const tr = document.createElement('tr');
            if (locked) tr.style.opacity = "0.7";

            const tdF = tr.insertCell();
            tdF.setAttribute('data-label', 'Fecha');
            tdF.innerHTML = `<b>${DIAS_SEMANA[dNum].slice(0, 3)}</b><br>${f.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit' })}`;
            tdF.rowSpan = 1 + extras.length;
            tdF.style.verticalAlign = 'top';

            const tdH = tr.insertCell();
            tdH.setAttribute('data-label', 'Horario');
            tdH.innerHTML = `<div class="stacked-cell">
            ${showM ? `<div class="time-text ${showT ? 'border-bottom' : ''}">${hM || '--:--'}</div>` : ''}
            ${showT ? `<div class="time-text">${hT}</div>` : ''}
        </div>`;

            state.colOrder.forEach(tipo => {
                const td = tr.insertCell();
                td.setAttribute('data-label', COL_DEFS[tipo]);
                const cont = document.createElement('div');
                cont.className = 'stacked-cell';

                ['M', 'T'].forEach(turno => {
                    if ((turno === 'M' && !showM) || (turno === 'T' && !showT)) return;

                    if (tipo === 'lug') {
                        const wrapper = document.createElement('div');
                        wrapper.className = 'lug-gru-wrapper';
                        const key = `${id}-${turno}`;

                        const btn = document.createElement('button');
                        btn.className = 'btn-add-gru';
                        // --- CAMBIO SOLICITADO: Tooltip ---
                        btn.title = "Agregar Grupo";

                        const s = document.createElement('select');
                        s.disabled = locked;
                        s.add(new Option("-", ""));
                        state.lugares.forEach(o => s.add(new Option(o, o, false, datos[`lug${turno}`] === o)));
                        s.onchange = (e) => guardarManual(id, e.target.value, `lug${turno}`);

                        const gruContainer = document.createElement('div');
                        gruContainer.className = 'gru-container';

                        const isOpen = !!state.uiState.openGroups[key];
                        gruContainer.style.display = isOpen ? 'block' : 'none';
                        btn.textContent = isOpen ? '-' : '+';

                        const gruSelector = crearSelectorGrupos(datos[`gru${turno}`], locked, (v) => guardarManual(id, v, `gru${turno}`));
                        gruContainer.appendChild(gruSelector);

                        btn.onclick = () => {
                            state.uiState.openGroups[key] = !state.uiState.openGroups[key];
                            actualizarTodo();
                        };

                        wrapper.appendChild(btn);
                        wrapper.appendChild(s);
                        wrapper.appendChild(gruContainer);
                        cont.appendChild(wrapper);
                    } else if (tipo === 'gru') {
                        cont.appendChild(crearSelectorGrupos(datos[`gru${turno}`], locked, (v) => guardarManual(id, v, `gru${turno}`)));
                    } else if (tipo === 'cua') {
                        const i = Object.assign(document.createElement('input'), {
                            type: 'text', className: 'input-cuadra', value: datos[`cua${turno}`] || "", disabled: locked
                        });
                        i.onchange = (e) => guardarManual(id, e.target.value, `cua${turno}`);
                        cont.appendChild(i);
                    } else {
                        const s = document.createElement('select');
                        s.disabled = locked;
                        s.add(new Option("-", ""));
                        let ops = (tipo === 'cond') ? state.conductores :
                            (tipo === 'ter' ? (state.mapaCoherencia[datos[`lug${turno}`]]?.ter || getRangoTerritorios()) :
                                (state.mapaCoherencia[datos[`lug${turno}`]]?.zon.filter(z => z.startsWith(datos[`ter${turno}`] + '-')) || getRangoZonas()));

                        ops.forEach(o => s.add(new Option(o, o, false, datos[`${tipo}${turno}`] === o)));
                        s.onchange = (e) => guardarManual(id, e.target.value, `${tipo}${turno}`);
                        cont.appendChild(s);
                    }
                });
                td.appendChild(cont);
            });

            const tdA = tr.insertCell();
            tdA.setAttribute('data-label', 'Acciones');
            tdA.innerHTML = `
            <div class="actions-wrapper">
                <div class="main-actions">
                    <button class="btn-lock" title="Bloquear">${locked ? '🔒' : '🔓'}</button>
                    <button class="btn-assign" title="Reasignar" ${locked ? 'disabled' : ''}>↻</button>
                    <button class="btn-note" title="Añadir Nota">📝</button>
                </div>
                <div class="extra-actions">
                    ${showM ? '<button class="btn-add-m" title="Añadir fila extra mañana">+M</button>' : ''}
                    ${showT ? '<button class="btn-add-t" title="Añadir fila extra tarde">+T</button>' : ''}
                </div>
            </div>`;

            tdA.querySelector('.btn-lock').onclick = () => { state.bloqueados[id] = !state.bloqueados[id]; actualizarTodo(); };
            tdA.querySelector('.btn-assign').onclick = () => { asignarGenerico('conductor', id); asignarGenerico('territorio', id); };
            tdA.querySelector('.btn-note').onclick = () => {
                const actual = state.comentarios[id] || "";
                const nuevo = prompt("Nota para este día (ej: Asamblea):", actual);
                if (nuevo !== null) {
                    if (nuevo.trim() === "") delete state.comentarios[id];
                    else state.comentarios[id] = nuevo;
                    actualizarTodo();
                }
            };
            if (showM) tdA.querySelector('.btn-add-m').onclick = () => agregarFilaExtra(id, 'M');
            if (showT) tdA.querySelector('.btn-add-t').onclick = () => agregarFilaExtra(id, 'T');

            els.tablaBody.appendChild(tr);

            extras.forEach(ex => {
                const trx = document.createElement('tr');
                if (locked) trx.style.opacity = "0.7";

                const tdExH = trx.insertCell();
                tdExH.setAttribute('data-label', 'Hora');
                tdExH.innerHTML = `<span style="border-left:3px solid ${ex.turno === 'M' ? '#2196F3' : '#FF9800'}; padding-left:5px; font-weight:bold;">${(ex.turno === 'M' ? hM : hT) || '--:--'}</span>`;

                state.colOrder.forEach(t => {
                    const tdEx = trx.insertCell();
                    tdEx.setAttribute('data-label', COL_DEFS[t]);
                    if (t === 'lug') {
                        const wrapper = document.createElement('div');
                        wrapper.className = 'lug-gru-wrapper';
                        const key = ex.id;

                        const btn = document.createElement('button');
                        btn.className = 'btn-add-gru';
                        // --- CAMBIO SOLICITADO: Tooltip ---
                        btn.title = "Agregar Grupo";

                        const s = document.createElement('select');
                        s.disabled = locked;
                        s.add(new Option("-", ""));
                        state.lugares.forEach(o => s.add(new Option(o, o, false, ex.lug === o)));
                        s.onchange = (e) => guardarExtra(id, ex.id, 'lug', e.target.value);

                        const gruContainer = document.createElement('div');
                        gruContainer.className = 'gru-container';

                        const isOpen = !!state.uiState.openGroups[key];
                        gruContainer.style.display = isOpen ? 'block' : 'none';
                        btn.textContent = isOpen ? '-' : '+';

                        const gruSelector = crearSelectorGrupos(ex.gru, locked, (v) => guardarExtra(id, ex.id, 'gru', v));
                        gruContainer.appendChild(gruSelector);

                        btn.onclick = () => {
                            state.uiState.openGroups[key] = !state.uiState.openGroups[key];
                            actualizarTodo();
                        };

                        wrapper.appendChild(btn);
                        wrapper.appendChild(s);
                        wrapper.appendChild(gruContainer);
                        tdEx.appendChild(wrapper);
                    } else if (t === 'gru') {
                        tdEx.appendChild(crearSelectorGrupos(ex[t], locked, (v) => guardarExtra(id, ex.id, t, v)));
                    } else if (t === 'cua') {
                        const i = Object.assign(document.createElement('input'), { type: 'text', className: 'input-cuadra', value: ex[t] || "", disabled: locked });
                        i.onchange = (e) => guardarExtra(id, ex.id, t, e.target.value);
                        tdEx.appendChild(i);
                    } else {
                        const s = document.createElement('select'); s.disabled = locked; s.add(new Option("-", ""));
                        let ops = (t === 'cond' ? state.conductores : t === 'ter' ? (state.mapaCoherencia[ex.lug]?.ter || getRangoTerritorios()) : (state.mapaCoherencia[ex.lug]?.zon.filter(z => z.startsWith(ex.ter + '-')) || getRangoZonas()));
                        ops.forEach(o => s.add(new Option(o, o, false, ex[t] === o)));
                        s.onchange = (e) => guardarExtra(id, ex.id, t, e.target.value);
                        tdEx.appendChild(s);
                    }
                });

                const tdAx = trx.insertCell();
                tdAx.setAttribute('data-label', 'Eliminar');
                const wrap = document.createElement('div'); wrap.className = 'actions-wrapper';
                const bDel = document.createElement('button');
                bDel.innerHTML = 'Eliminar fila &times;';
                bDel.style.color = 'red';
                bDel.style.width = '100%';
                bDel.disabled = locked;
                bDel.onclick = () => eliminarFilaExtra(id, ex.id);
                wrap.appendChild(bDel);
                tdAx.appendChild(wrap);
                els.tablaBody.appendChild(trx);
            });
        });

        const ids = fechas.map(f => f.toISOString().split('T')[0]);
        els.btnBloquear.innerHTML = (ids.length > 0 && ids.every(i => state.bloqueados[i])) ? '🔓 Desbloquear Todo' : '🔒 Bloquear Todo';
    }

    function actualizarTodo() { renderListas(); renderColumnSelector(); renderSelectorDias(); renderTabla(); saveToStorage(); }

    // --- PDF EXPORT ---
    function exportarPDF() {
        const formatTime12Hour = (time24) => {
            if (!time24 || !time24.includes(':')) return '--:--';
            const [hours, minutes] = time24.split(':');
            let h = parseInt(hours, 10);
            const ampm = h >= 12 ? 'pm' : 'am';
            h = h % 12;
            h = h ? h : 12;
            return `${h}:${minutes} ${ampm}`;
        };

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
            if (!sig || sig.getDay() === state.startOfWeek) {
                semanas.push([...semActual]);
                semActual = [];
            }
        });

        const pdfHeaders = ['DÍA', 'HORA', ...state.colOrder.map(k => COL_DEFS[k].toUpperCase())];
        let currentY = 18;
        let ultimoMesTitulado = "";
        const PAGE_LIMIT = 285;

        semanas.forEach((semana, idx) => {
            const fI = semana[0];
            const fF = semana[semana.length - 1];
            const mesI = fI.toLocaleString('es-ES', { month: 'long' }).toUpperCase();
            const anioI = fI.getFullYear();
            const mesF = fF.toLocaleString('es-ES', { month: 'long' }).toUpperCase();
            const mesSemanaKey = `${mesI} ${anioI}`;

            let totalFilasCuerpo = 1;
            semana.forEach(f => {
                const id = f.toISOString().split('T')[0];
                const cfg = state.configDias[f.getDay()];
                const esFindeNativo = (f.getDay() === 0 || f.getDay() === 6);
                const hT = esFindeNativo ? (state.horario.tFin || state.horario.t) : state.horario.t;
                const extras = (state.filasExtras[id] || []).length;
                const tieneComentario = !!state.comentarios[id];
                totalFilasCuerpo += (cfg.m ? 1 : 0) + (cfg.t && hT !== '' ? 1 : 0) + extras + (tieneComentario ? 1 : 0);
            });
            const estimatedHeight = (totalFilasCuerpo * 7) + 8;

            const esNuevoMes = mesSemanaKey !== ultimoMesTitulado;

            if ((currentY + estimatedHeight > PAGE_LIMIT)) {
                doc.addPage();
                currentY = 18;
            }

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

                if (state.comentarios[id]) {
                    weekRows.push([{
                        content: state.comentarios[id].toUpperCase(),
                        colSpan: pdfHeaders.length,
                        styles: {
                            halign: 'center',
                            fillColor: [255, 253, 208],
                            textColor: [50, 50, 50],
                            fontStyle: 'bolditalic',
                            fontSize: 7.5,
                            cellPadding: 1.5
                        }
                    }]);
                }

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
                    row.push(formatTime12Hour(hora));
                    state.colOrder.forEach(k => {
                        let val = turnoLetra ? source[`${k}${turnoLetra}`] : source[k];

                        // --- CAMBIO SOLICITADO: Inyectar grupo en Lugar de salida ---
                        if (k === 'lug') {
                            const rawGru = turnoLetra ? source[`gru${turnoLetra}`] : source['gru'];
                            const gruTxt = formatearGruposParaPDF(rawGru);
                            const lugTxt = val || "-";

                            if (gruTxt) {
                                // Simulamos "celda izquierda | celda derecha" en una sola columna con espaciado
                                val = `${gruTxt}   |   ${lugTxt}`;
                            } else {
                                val = lugTxt;
                            }
                        } else if (k === 'gru') {
                            val = formatearGruposParaPDF(val) || "-";
                        } else {
                            val = val || "-";
                        }

                        row.push(val);
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

            currentY = doc.lastAutoTable.finalY;
        });

        doc.save(`Programa_Servicio.pdf`);
    }

    function init() {
        loadFromStorage();

        const handleTheme = () => {
            const el = document.documentElement;
            const savedTheme = localStorage.getItem(KEYS.THEME);
            const applyTheme = (theme) => { el.setAttribute('data-theme', theme); };
            const toggleTheme = () => {
                const currentTheme = el.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
                applyTheme(currentTheme);
                localStorage.setItem(KEYS.THEME, currentTheme);
            };
            document.getElementById('btnToggleTheme').onclick = toggleTheme;
            const mediaQuery = window.matchMedia('(prefers-color-scheme: dark)');
            mediaQuery.addEventListener('change', (e) => {
                if (!localStorage.getItem(KEYS.THEME)) { applyTheme(e.matches ? 'dark' : 'light'); }
            });
            if (savedTheme) { applyTheme(savedTheme); } else { applyTheme(mediaQuery.matches ? 'dark' : 'light'); }
        };

        handleTheme();

        if (!els.fInicio.value) els.fInicio.value = new Date().toISOString().split('T')[0];
        els.hManana.value = state.horario.m; els.hTarde.value = state.horario.t; els.hMananaFin.value = state.horario.mFin; els.hTardeFin.value = state.horario.tFin;
        els.diasContainer.onchange = (e) => { if (e.target.type === 'checkbox') { state.configDias[e.target.dataset.day][e.target.dataset.shift] = e.target.checked; actualizarTodo(); } };
        els.fInicio.onchange = () => ajustarFechas('combo'); els.fFinal.onchange = () => ajustarFechas('final'); els.cantFechas.onchange = () => ajustarFechas('combo');
        document.getElementById('btnAgregarConductor').onclick = () => modificarLista('conductor', 'agregar');
        document.getElementById('btnAgregarLugar').onclick = () => modificarLista('lugar', 'agregar');
        document.getElementById('btnAgregarGrupo').onclick = () => modificarLista('grupo', 'agregar');
        document.getElementById('btnLimpiarGrupos').onclick = () => modificarLista('grupo', 'borrarTodo');
        document.querySelectorAll('.btn-export').forEach(btn => btn.onclick = exportarPDF);
        document.getElementById('btnAsignarConductores').onclick = () => asignarGenerico('conductor');
        document.getElementById('btnAsignarTerritorios').onclick = () => asignarGenerico('territorio');
        document.getElementById('btnLimpiarAsignaciones').onclick = () => { if (confirm("¿Vaciar todo?")) { generarFechasDiarias().forEach(f => { const id = f.toISOString().split('T')[0]; if (!state.bloqueados[id]) delete state.asignaciones[id]; }); actualizarTodo(); } };
        els.btnBloquear.onclick = () => { const ids = generarFechasDiarias().map(f => f.toISOString().split('T')[0]), all = ids.every(i => state.bloqueados[i]); ids.forEach(i => state.bloqueados[i] = !all); actualizarTodo(); };

        const layout = document.querySelector('.main-layout');
        const sidebarIcon = document.getElementById('sidebarIcon');
        const sidebarText = document.getElementById('sidebarText');

        const actualizarInterfazBoton = () => {
            const esMovil = window.innerWidth <= 850;
            const estaOculto = layout.classList.contains('sidebar-hidden');

            if (esMovil) {
                sidebarIcon.textContent = estaOculto ? "⚙️" : "❌";
                sidebarText.textContent = estaOculto ? "Configuración" : "Cerrar";
            } else {
                sidebarIcon.textContent = estaOculto ? "➡️" : "⬅️";
                sidebarText.textContent = estaOculto ? "Mostrar" : "Ocultar";
            }
        };

        document.getElementById('toggleSidebar').onclick = () => {
            layout.classList.toggle('sidebar-hidden');
            actualizarInterfazBoton();
            setTimeout(() => { if (typeof actualizarTodo === 'function') actualizarTodo(); }, 300);
        };

        const overlay = document.getElementById('sidebarOverlay');
        if (overlay) {
            overlay.onclick = () => {
                if (!layout.classList.contains('sidebar-hidden')) {
                    layout.classList.add('sidebar-hidden');
                    actualizarInterfazBoton();
                    setTimeout(() => { if (typeof actualizarTodo === 'function') actualizarTodo(); }, 300);
                }
            };
        }

        window.addEventListener('resize', actualizarInterfazBoton);
        actualizarInterfazBoton();

        if (window.innerWidth <= 850) {
            layout.classList.add('sidebar-hidden');
        } else {
            layout.classList.remove('sidebar-hidden');
        }
        actualizarInterfazBoton();

        document.getElementById('btnFijarHorario').onclick = () => { state.horario = { m: els.hManana.value, t: els.hTarde.value, mFin: els.hMananaFin.value, tFin: els.hTardeFin.value }; actualizarTodo(); };
        document.querySelectorAll('.tab-button').forEach(btn => btn.onclick = () => { document.querySelectorAll('.tab-button').forEach(b => b.classList.remove('active')); btn.classList.add('active'); document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active')); document.getElementById(btn.dataset.tab + 'Pane').classList.add('active'); });

        const selStart = document.getElementById('inicioSemanaSelect');
        if (selStart) {
            selStart.value = state.startOfWeek;
            selStart.onchange = (e) => {
                state.startOfWeek = parseInt(e.target.value);
                actualizarTodo();
            };
        }

        ajustarFechas('combo');
    }
    document.addEventListener('DOMContentLoaded', init);
})();
