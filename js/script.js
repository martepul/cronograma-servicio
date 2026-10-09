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
        CONFIG_ORDER: 'config_order_v3',
        CONFIG_DIAS: 'config_dias_v1',
        EXTRAS: 'filas_extras_v2',
        COMENTARIOS: 'comentarios_v1',
        START_WEEK: 'start_week_v1',
        THEME: 'theme_v1',
        CONFIG_TERRITORIOS: 'config_ter_v1',
        SEMANAS_CUSTOM: 'semanas_custom_v1', // Nueva llave para semanas personalizadas
        HORARIOS_CUSTOM: 'horarios_custom_v1' // Nueva llave para horarios personalizados por día
    };

    const COL_DEFS = {
        lug: "Lugar de salida",
        ter: "Territorio",
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
        colOrder: ['lug', 'ter', 'cond'],
        configDias: {},
        comentarios: {},
        startOfWeek: 1,
        configTerritorios: {},
        semanasCustom: {}, // Nuevo estado para semanas custom
        horariosCustom: {}, // Nuevo estado para horarios custom por día
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
        diasContainer: document.getElementById('selectorDiasContainer'),

        // Elementos del Modal
        btnConfigTer: document.getElementById('btnConfigTerritorios'),
        modalTer: document.getElementById('modalTerritorios'),
        btnCerrarModalTer: document.getElementById('btnCerrarModalTer'),
        listaConfigTer: document.getElementById('listaConfigTerritorios'),
        btnGenerarCronograma: document.getElementById('btnGenerarCronograma')
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
            condM: man.condM || asig.condM, condT: man.condT || asig.condT,
            cuaM: man.cuaM || "", cuaT: man.cuaT || "",
            gruM: man.gruM || "", gruT: man.gruT || ""
        };
    };

    const createDefaultDisp = () => {
        let d = {};
        for (let i = 0; i < 7; i++) d[i] = { m: false, t: false, pM: false, pT: false };
        return d;
    };

    const createDefaultTerDisp = () => {
        let d = {};
        for (let i = 0; i < 7; i++) d[i] = { m: true, t: true };
        return d;
    };

    // --- PERSISTENCIA ---
    function loadFromStorage() {
        try {
            const rawCond = JSON.parse(localStorage.getItem(KEYS.CONDUCTORES)) || [];
            state.conductores = rawCond.map(c => {
                if (typeof c === 'string') return { name: c, disponibilidad: createDefaultDisp(), cartas: false };

                let disp = c.disponibilidad;
                if (!disp) {
                    disp = createDefaultDisp();
                    if (Array.isArray(c.tags)) {
                        c.tags.forEach(t => {
                            disp[Number(t)] = { m: true, t: true, pM: false, pT: false };
                        });
                    }
                }
                return { name: c.name || c, disponibilidad: disp, cartas: c.cartas || false };
            });

            state.lugares = JSON.parse(localStorage.getItem(KEYS.LUGARES)) || INITIAL_LUGARES;
            state.grupos = JSON.parse(localStorage.getItem(KEYS.GRUPOS)) || [];
            state.mapaCoherencia = JSON.parse(localStorage.getItem(KEYS.MAPA)) || INITIAL_MAPA;
            state.manuales = JSON.parse(localStorage.getItem(KEYS.MANUALES)) || {};
            state.bloqueados = JSON.parse(localStorage.getItem(KEYS.BLOQUEADOS)) || {};
            state.horario = JSON.parse(localStorage.getItem(KEYS.HORARIO)) || { m: '', t: '', mFin: '', tFin: '' };
            state.filasExtras = JSON.parse(localStorage.getItem(KEYS.EXTRAS)) || {};
            state.comentarios = JSON.parse(localStorage.getItem(KEYS.COMENTARIOS)) || {};
            state.configTerritorios = JSON.parse(localStorage.getItem(KEYS.CONFIG_TERRITORIOS)) || {};
            state.semanasCustom = JSON.parse(localStorage.getItem(KEYS.SEMANAS_CUSTOM)) || {};
            state.horariosCustom = JSON.parse(localStorage.getItem(KEYS.HORARIOS_CUSTOM)) || {};

            const storedOrder = JSON.parse(localStorage.getItem(KEYS.CONFIG_ORDER));
            if (storedOrder) {
                state.colOrder = storedOrder.filter(c => c !== 'zon');
            } else {
                state.colOrder = ['lug', 'ter', 'cond'];
            }

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
        localStorage.setItem(KEYS.CONFIG_TERRITORIOS, JSON.stringify(state.configTerritorios));
        localStorage.setItem(KEYS.SEMANAS_CUSTOM, JSON.stringify(state.semanasCustom));
        localStorage.setItem(KEYS.HORARIOS_CUSTOM, JSON.stringify(state.horariosCustom));
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
                        state.conductores.push({ name: nombre, disponibilidad: createDefaultDisp(), cartas: false });
                        agregados++;
                    }
                } else if (tipo === 'lugar') {
                    const nombre = partes[0];
                    if (nombre) {
                        if (!state.lugares.includes(nombre)) { state.lugares.push(nombre); agregados++; }
                        const territorios = partes.slice(1);
                        state.mapaCoherencia[nombre] = { ter: territorios };
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

    function generarCronograma(targetId = null) {
        if (state.conductores.length === 0) return alert("Faltan conductores configurados.");
        if (state.lugares.length === 0) return alert("Faltan lugares configurados.");

        const fechas = generarFechasDiarias();
        const conteoMensual = {};
        const todosTer = getRangoTerritorios();

        // Inicializar conteo
        todosTer.forEach(t => conteoMensual[t] = 0);

        // 1. Limpiar automáticos y contar usos manuales/fijos en el mes
        fechas.forEach(f => {
            const id = f.toISOString().split('T')[0];

            if (!state.bloqueados[id] && (!targetId || targetId === id)) {
                if (!state.asignaciones[id]) state.asignaciones[id] = {};
                delete state.asignaciones[id].lugM;
                delete state.asignaciones[id].terM;
                delete state.asignaciones[id].condM;
                delete state.asignaciones[id].lugT;
                delete state.asignaciones[id].terT;
                delete state.asignaciones[id].condT;
            }

            ['M', 'T'].forEach(turno => {
                const terManual = state.manuales[id]?.[`ter${turno}`];
                const terAsig = (!targetId || targetId !== id || state.bloqueados[id]) ? state.asignaciones[id]?.[`ter${turno}`] : null;
                const ter = terManual || terAsig;
                if (ter && conteoMensual[ter] !== undefined) conteoMensual[ter]++;
            });

            (state.filasExtras[id] || []).forEach(ex => {
                if (ex.ter && conteoMensual[ex.ter] !== undefined) conteoMensual[ex.ter]++;
            });
        });

        // 2. Asignar dinámicamente según configuración
        fechas.forEach(f => {
            const id = f.toISOString().split('T')[0];
            if ((targetId && id !== targetId) || state.bloqueados[id]) return;

            if (!state.asignaciones[id]) state.asignaciones[id] = {};
            const asig = state.asignaciones[id];
            const diaNum = f.getDay();
            const configDia = state.configDias[diaNum];
            const esFinde = (diaNum === 0 || diaNum === 6);
            const customH = state.horariosCustom[id] || {};
            const hT = (customH.t !== undefined && customH.t !== '') ? customH.t : (esFinde ? (state.horario.tFin || state.horario.t) : state.horario.t);
            const usarTarde = configDia.t && hT !== '';

            // Evaluamos en qué mes cae el día actual (0 = Enero, 1 = Febrero, etc.)
            const mesActual = f.getMonth();

            const usadosHoyLug = [];
            const usadosHoyTer = [];

            ['M', 'T'].forEach(turno => {
                if ((turno === 'M' && !configDia.m) || (turno === 'T' && !usarTarde)) return;

                const manualCond = state.manuales[id]?.[`cond${turno}`];
                if (manualCond) usadosHoyLug.push(manualCond);

                const manualLug = state.manuales[id]?.[`lug${turno}`];
                const manualTer = state.manuales[id]?.[`ter${turno}`];
                let lugAct = manualLug;
                let terAct = manualTer;

                if (!lugAct || !terAct) {
                    const terDisponibles = todosTer.filter(t => {
                        if (usadosHoyTer.includes(t)) return false;

                        const cfg = state.configTerritorios[t] || { frecuencia: 4, disponibilidad: createDefaultTerDisp(), mesesActivos: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11] };
                        if (!cfg.mesesActivos) cfg.mesesActivos = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];

                        // Validar si el territorio está activo para el mes de la fecha actual
                        if (!cfg.mesesActivos.includes(mesActual)) return false;

                        if (conteoMensual[t] >= cfg.frecuencia) return false;

                        const dispDia = cfg.disponibilidad[diaNum];
                        if (turno === 'M' && !dispDia.m) return false;
                        if (turno === 'T' && !dispDia.t) return false;

                        return true;
                    });

                    const disponiblesLug = state.lugares.filter(l => !usadosHoyLug.includes(l));
                    let combinaciones = [];

                    disponiblesLug.forEach(l => {
                        const cfgLug = state.mapaCoherencia[l];
                        if (!cfgLug || cfgLug.ter.length === 0) return;

                        cfgLug.ter.forEach(t => {
                            if (terDisponibles.includes(t)) {
                                combinaciones.push({ lug: l, ter: t });
                            }
                        });
                    });

                    if (combinaciones.length > 0) {
                        const el = combinaciones[Math.floor(Math.random() * combinaciones.length)];
                        asig[`lug${turno}`] = el.lug;
                        asig[`ter${turno}`] = el.ter;
                        lugAct = el.lug;
                        terAct = el.ter;
                        usadosHoyLug.push(lugAct);
                        usadosHoyTer.push(terAct);
                        conteoMensual[terAct]++;
                    }
                } else {
                    usadosHoyLug.push(lugAct);
                    if (terAct) usadosHoyTer.push(terAct);
                }

                if (!manualCond) {
                    const esCartas = terAct && terAct.toLowerCase().includes('cartas');
                    const poolCond = state.conductores.filter(c => {
                        if (usadosHoyLug.includes(c.name)) return false;
                        if (esCartas && c.cartas) return true;

                        const hasAnyConfig = Object.values(c.disponibilidad).some(d => d.m || d.t);
                        if (!hasAnyConfig) return true;

                        return turno === 'M' ? c.disponibilidad[diaNum].m : c.disponibilidad[diaNum].t;
                    });

                    const prioritarios = poolCond.filter(c => turno === 'M' ? c.disponibilidad[diaNum].pM : c.disponibilidad[diaNum].pT);
                    const listaFinal = prioritarios.length > 0 ? prioritarios : poolCond;
                    const elegido = listaFinal[Math.floor(Math.random() * listaFinal.length)];

                    if (elegido) {
                        asig[`cond${turno}`] = elegido.name;
                        usadosHoyLug.push(elegido.name);
                    }
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
            if (tipo === 'conductor') { if (!lista.some(c => c.name === raw)) { lista.push({ name: raw, disponibilidad: createDefaultDisp(), cartas: false }); input.value = ''; actualizarTodo(); } return; }
            if (tipo === 'grupo') {
                const range = raw.match(/^(\d+)-(\d+)$/);
                if (range) { for (let i = parseInt(range[1]); i <= parseInt(range[2]); i++) { if (!lista.includes(`Grupo ${i}`)) lista.push(`Grupo ${i}`); } input.value = ''; actualizarTodo(); return; }
            }
            if (!lista.includes(raw)) {
                if (tipo === 'lugar') {
                    const t = prompt("Territorios (separados por coma):") || "";
                    state.mapaCoherencia[raw] = { ter: t.split(',').map(s => s.trim()).filter(s => s) };
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
                }
            } else if (rol.startsWith('ter')) {
                const t = rol.slice(-1);
                const currentData = obtenerDatosDia(fechaId);
                const currentLug = currentData[`lug${t}`];

                const lugEsValido = currentLug && state.mapaCoherencia[currentLug] && state.mapaCoherencia[currentLug].ter.includes(valor);

                if (!lugEsValido) {
                    const linkedLug = Object.keys(state.mapaCoherencia).find(l => state.mapaCoherencia[l].ter.includes(valor));
                    if (linkedLug) {
                        state.manuales[fechaId][`lug${t}`] = linkedLug;
                    }
                } else {
                    state.manuales[fechaId][`lug${t}`] = currentLug;
                }
            }
        }
        actualizarTodo();
    }

    function agregarFilaExtra(dateId, turno) {
        if (state.bloqueados[dateId]) return;
        if (!state.filasExtras[dateId]) state.filasExtras[dateId] = [];
        state.filasExtras[dateId].push({ id: Date.now(), turno, lug: '', ter: '', cond: '', cua: '', gru: '' });
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
                }
            } else if (campo === 'ter') {
                const lugEsValido = row.lug && state.mapaCoherencia[row.lug] && state.mapaCoherencia[row.lug].ter.includes(valor);
                if (!lugEsValido) {
                    const linkedLug = Object.keys(state.mapaCoherencia).find(l => state.mapaCoherencia[l].ter.includes(valor));
                    if (linkedLug) {
                        row.lug = linkedLug;
                    }
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
        const all = [{ id: 'lug', lbl: 'Lugar' }, { id: 'ter', lbl: 'Territorio' }, { id: 'cond', lbl: 'Conductor' }, { id: 'cua', lbl: 'Cuadra' }, { id: 'gru', lbl: 'Grupo' }];
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

    function renderConfigTerritorios() {
        if (!els.listaConfigTer) return;
        els.listaConfigTer.innerHTML = '';
        const todosTer = getRangoTerritorios();

        todosTer.forEach(t => {
            if (!state.configTerritorios[t]) {
                state.configTerritorios[t] = { frecuencia: 4, disponibilidad: createDefaultTerDisp(), mesesActivos: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11] };
            }
            const cfg = state.configTerritorios[t];
            if (!cfg.mesesActivos) cfg.mesesActivos = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11];

            const li = document.createElement('li');
            li.className = 'ter-config-item';

            const topDiv = document.createElement('div');
            topDiv.className = 'ter-config-main';

            const nameSpan = document.createElement('span');
            nameSpan.textContent = t;

            const selFreq = document.createElement('select');
            for (let i = 1; i <= 31; i++) {
                selFreq.add(new Option(`${i} ${i === 1 ? 'vez' : 'veces'}/mes`, i, false, cfg.frecuencia == i));
            }
            selFreq.onchange = (e) => { cfg.frecuencia = parseInt(e.target.value); saveToStorage(); };

            const centerGroup = document.createElement('div');
            centerGroup.style.display = 'flex';
            centerGroup.style.alignItems = 'center';
            centerGroup.style.justifyContent = 'space-between';

            const tagCont = document.createElement('div');
            tagCont.className = 'tag-container';

            ['D', 'L', 'M', 'Mi', 'J', 'V', 'S'].forEach((letra, idx) => {
                const disp = cfg.disponibilidad[idx];
                const isActive = disp.m || disp.t;
                const btn = document.createElement('button');
                btn.textContent = letra;
                btn.className = `btn-tag ${isActive ? 'active' : ''}`;
                btn.onclick = () => {
                    if (isActive) {
                        cfg.disponibilidad[idx] = { m: false, t: false };
                    } else {
                        cfg.disponibilidad[idx] = { m: true, t: true };
                    }
                    saveToStorage();
                    renderConfigTerritorios();
                };
                tagCont.appendChild(btn);
            });

            const uiKey = 'ter_cfg_' + t;
            const configBtn = document.createElement('button');
            configBtn.innerHTML = '⚙️';
            configBtn.style.background = 'transparent';
            configBtn.style.border = 'none';
            configBtn.style.cursor = 'pointer';
            configBtn.style.marginLeft = '5px';
            configBtn.onclick = () => {
                state.uiState[uiKey] = !state.uiState[uiKey];
                renderConfigTerritorios();
            };

            centerGroup.append(tagCont, configBtn);
            topDiv.append(nameSpan, selFreq, centerGroup);
            li.appendChild(topDiv);

            const detailsDiv = document.createElement('div');
            detailsDiv.style.display = state.uiState[uiKey] ? 'block' : 'none';
            detailsDiv.style.fontSize = '0.8em';
            detailsDiv.style.marginTop = '8px';
            detailsDiv.style.background = 'var(--bg)';
            detailsDiv.style.padding = '8px';
            detailsDiv.style.borderRadius = '2px';
            detailsDiv.style.border = '1px solid var(--border)';

            const header = document.createElement('div');
            header.style.display = 'grid';
            header.style.gridTemplateColumns = '50px 1fr 1fr';
            header.style.textAlign = 'center';
            header.style.fontWeight = 'bold';
            header.style.marginBottom = '5px';
            header.style.borderBottom = '1px solid var(--border)';
            header.style.paddingBottom = '3px';
            header.innerHTML = `<span>Día</span><span>☀️ Mañana</span><span>🌤️ Tarde</span>`;
            detailsDiv.appendChild(header);

            ['D', 'L', 'M', 'Mi', 'J', 'V', 'S'].forEach((letra, idx) => {
                const row = document.createElement('div');
                row.style.display = 'grid';
                row.style.gridTemplateColumns = '50px 1fr 1fr';
                row.style.textAlign = 'center';
                row.style.alignItems = 'center';
                row.style.marginBottom = '4px';

                const dSpan = document.createElement('span');
                dSpan.textContent = letra;
                dSpan.style.fontWeight = 'bold';

                const chkM = document.createElement('input'); chkM.type = 'checkbox'; chkM.checked = cfg.disponibilidad[idx].m;
                const chkT = document.createElement('input'); chkT.type = 'checkbox'; chkT.checked = cfg.disponibilidad[idx].t;

                const updateDisp = () => {
                    cfg.disponibilidad[idx] = { m: chkM.checked, t: chkT.checked };
                    saveToStorage();
                    const topBtn = tagCont.children[idx];
                    if (chkM.checked || chkT.checked) topBtn.classList.add('active');
                    else topBtn.classList.remove('active');
                };

                chkM.onchange = updateDisp; chkT.onchange = updateDisp;
                row.append(dSpan, chkM, chkT);
                detailsDiv.appendChild(row);
            });

            const mesesCont = document.createElement('div');
            mesesCont.style.marginTop = '8px';
            mesesCont.style.borderTop = '1px dashed var(--border)';
            mesesCont.style.paddingTop = '8px';

            const mesesHeader = document.createElement('div');
            mesesHeader.style.textAlign = 'center';
            mesesHeader.style.fontWeight = 'bold';
            mesesHeader.style.marginBottom = '5px';
            mesesHeader.textContent = 'Meses Activos';
            mesesCont.appendChild(mesesHeader);

            const gridMeses = document.createElement('div');
            gridMeses.style.display = 'grid';
            gridMeses.style.gridTemplateColumns = 'repeat(6, 1fr)';
            gridMeses.style.gap = '4px';

            const nombresMeses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];

            nombresMeses.forEach((nombre, mesIdx) => {
                const btnMes = document.createElement('button');
                btnMes.textContent = nombre;
                const isActive = cfg.mesesActivos.includes(mesIdx);
                btnMes.className = `btn-tag ${isActive ? 'active' : ''}`;
                btnMes.style.width = '100%';
                btnMes.style.padding = '2px 0';
                btnMes.onclick = () => {
                    if (cfg.mesesActivos.includes(mesIdx)) {
                        cfg.mesesActivos = cfg.mesesActivos.filter(m => m !== mesIdx);
                        btnMes.classList.remove('active');
                    } else {
                        cfg.mesesActivos.push(mesIdx);
                        btnMes.classList.add('active');
                    }
                    saveToStorage();
                };
                gridMeses.appendChild(btnMes);
            });

            mesesCont.appendChild(gridMeses);
            detailsDiv.appendChild(mesesCont);

            li.appendChild(detailsDiv);
            els.listaConfigTer.appendChild(li);
        });
    }

    function renderListas() {
        const draw = (l, container, tipo) => {
            container.innerHTML = '';
            l.forEach((item, i) => {
                const li = document.createElement('li');

                if (tipo === 'conductor') {
                    li.className = 'participant-item-complex';
                    li.style.flexDirection = 'column';
                    li.style.alignItems = 'stretch';

                    const topRow = document.createElement('div');
                    topRow.style.display = 'flex';
                    topRow.style.alignItems = 'center';
                    topRow.style.justifyContent = 'space-between';

                    const nameSpan = document.createElement('span');
                    nameSpan.textContent = item.name;
                    nameSpan.style.flex = '1';

                    const rightGroup = document.createElement('div');
                    rightGroup.style.display = 'flex';
                    rightGroup.style.alignItems = 'center';

                    const tagCont = document.createElement('div');
                    tagCont.className = 'tag-container';

                    ['D', 'L', 'M', 'Mi', 'J', 'V', 'S'].forEach((letra, idx) => {
                        const disp = item.disponibilidad[idx];
                        const isActive = disp.m || disp.t;
                        const btn = document.createElement('button');
                        btn.textContent = letra;
                        btn.className = `btn-tag ${isActive ? 'active' : ''}`;
                        btn.onclick = () => {
                            if (isActive) {
                                item.disponibilidad[idx] = { m: false, t: false, pM: false, pT: false };
                            } else {
                                item.disponibilidad[idx] = { m: true, t: true, pM: false, pT: false };
                            }
                            actualizarTodo();
                        };
                        tagCont.appendChild(btn);
                    });

                    const uiKey = 'cond_cfg_' + item.name;
                    const configBtn = document.createElement('button');
                    configBtn.innerHTML = '⚙️';
                    configBtn.style.background = 'transparent';
                    configBtn.style.border = 'none';
                    configBtn.style.cursor = 'pointer';
                    configBtn.style.marginLeft = '5px';
                    configBtn.onclick = () => {
                        state.uiState[uiKey] = !state.uiState[uiKey];
                        actualizarTodo();
                    };

                    const delBtn = document.createElement('button');
                    delBtn.textContent = '×';
                    delBtn.style.marginLeft = '5px';
                    delBtn.style.background = 'transparent';
                    delBtn.style.border = 'none';
                    delBtn.style.cursor = 'pointer';
                    delBtn.style.color = '#d32f2f';
                    delBtn.onclick = () => modificarLista(tipo, 'eliminar', i);

                    rightGroup.append(tagCont, configBtn, delBtn);
                    topRow.append(nameSpan, rightGroup);
                    li.appendChild(topRow);

                    const detailsDiv = document.createElement('div');
                    detailsDiv.style.display = state.uiState[uiKey] ? 'block' : 'none';
                    detailsDiv.style.fontSize = '0.8em';
                    detailsDiv.style.marginTop = '8px';
                    detailsDiv.style.background = 'var(--card-bg)';
                    detailsDiv.style.padding = '8px';
                    detailsDiv.style.borderRadius = '4px';
                    detailsDiv.style.border = '1px solid var(--border)';

                    const header = document.createElement('div');
                    header.style.display = 'grid';
                    header.style.gridTemplateColumns = '30px 1fr 1fr 1fr 1fr';
                    header.style.textAlign = 'center';
                    header.style.fontWeight = 'bold';
                    header.style.marginBottom = '5px';
                    header.style.borderBottom = '1px solid var(--border)';
                    header.style.paddingBottom = '3px';
                    header.innerHTML = `<span>Día</span><span title="Mañana">☀️ M</span><span title="Tarde">🌤️ T</span><span title="Prioridad Mañana">⭐ M</span><span title="Prioridad Tarde">⭐ T</span>`;
                    detailsDiv.appendChild(header);

                    ['D', 'L', 'M', 'Mi', 'J', 'V', 'S'].forEach((letra, idx) => {
                        const row = document.createElement('div');
                        row.style.display = 'grid';
                        row.style.gridTemplateColumns = '30px 1fr 1fr 1fr 1fr';
                        row.style.textAlign = 'center';
                        row.style.alignItems = 'center';
                        row.style.marginBottom = '4px';

                        const dSpan = document.createElement('span');
                        dSpan.textContent = letra;
                        dSpan.style.fontWeight = 'bold';

                        const chkM = document.createElement('input'); chkM.type = 'checkbox'; chkM.checked = item.disponibilidad[idx].m;
                        const chkT = document.createElement('input'); chkT.type = 'checkbox'; chkT.checked = item.disponibilidad[idx].t;
                        const chkPM = document.createElement('input'); chkPM.type = 'checkbox'; chkPM.checked = item.disponibilidad[idx].pM;
                        const chkPT = document.createElement('input'); chkPT.type = 'checkbox'; chkPT.checked = item.disponibilidad[idx].pT;

                        const updateDisp = (e) => {
                            if (e.target === chkPM && chkPM.checked) chkM.checked = true;
                            if (e.target === chkPT && chkPT.checked) chkT.checked = true;
                            if (e.target === chkM && !chkM.checked) chkPM.checked = false;
                            if (e.target === chkT && !chkT.checked) chkPT.checked = false;

                            item.disponibilidad[idx] = { m: chkM.checked, t: chkT.checked, pM: chkPM.checked, pT: chkPT.checked };

                            const topBtn = tagCont.children[idx];
                            if (chkM.checked || chkT.checked) topBtn.classList.add('active');
                            else topBtn.classList.remove('active');

                            saveToStorage();
                        };

                        chkM.onchange = updateDisp; chkT.onchange = updateDisp; chkPM.onchange = updateDisp; chkPT.onchange = updateDisp;
                        row.append(dSpan, chkM, chkT, chkPM, chkPT);
                        detailsDiv.appendChild(row);
                    });

                    const genOptions = document.createElement('div');
                    genOptions.style.marginTop = '8px';
                    genOptions.style.borderTop = '1px dashed var(--border)';
                    genOptions.style.paddingTop = '8px';
                    genOptions.innerHTML = `<label style="cursor:pointer; display:flex; align-items:center; gap:5px;"><input type="checkbox" ${item.cartas ? 'checked' : ''}> ✉️ Disponible siempre para territorio "Cartas"</label>`;

                    const chkCartas = genOptions.querySelector('input');
                    chkCartas.onchange = (e) => {
                        item.cartas = e.target.checked;
                        actualizarTodo();
                    };
                    detailsDiv.appendChild(genOptions);

                    li.appendChild(detailsDiv);

                } else {
                    const spanName = document.createElement('span');
                    spanName.textContent = item;
                    li.appendChild(spanName);
                    const btn = document.createElement('button');
                    btn.textContent = '×';
                    btn.onclick = () => modificarLista(tipo, 'eliminar', i);
                    li.appendChild(btn);
                }
                container.appendChild(li);
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
            const esFinde = (dNum === 0 || dNum === 6);
            const customH = state.horariosCustom[id] || {};
            const hM = (customH.m !== undefined && customH.m !== '') ? customH.m : (esFinde ? (state.horario.mFin || state.horario.m) : state.horario.m);
            const hT = (customH.t !== undefined && customH.t !== '') ? customH.t : (esFinde ? (state.horario.tFin || state.horario.t) : state.horario.t);
            const showM = cfg.m, showT = cfg.t && hT !== '', extras = state.filasExtras[id] || [], datos = obtenerDatosDia(id);

            // Render marcador de semana custom si existe
            if (state.semanasCustom[id]) {
                const trW = document.createElement('tr');
                trW.className = 'week-row';
                const tdW = document.createElement('td');
                tdW.colSpan = 2 + state.colOrder.length + 1;
                tdW.style.backgroundColor = 'rgba(33, 150, 243, 0.1)';
                tdW.style.textAlign = 'center';
                tdW.style.fontWeight = 'bold';
                tdW.style.padding = '5px';
                tdW.innerHTML = `<span>🏷️ BARRA DE SEMANA: ${state.semanasCustom[id].toUpperCase()}</span>`;
                if (!locked) {
                    const b = document.createElement('button');
                    b.innerHTML = '&times;';
                    b.style.marginLeft = '10px';
                    b.style.cursor = 'pointer';
                    b.style.border = 'none';
                    b.style.background = 'transparent';
                    b.style.color = 'red';
                    b.style.fontWeight = 'bold';
                    b.onclick = () => { delete state.semanasCustom[id]; actualizarTodo(); };
                    tdW.appendChild(b);
                }
                trW.appendChild(tdW);
                els.tablaBody.appendChild(trW);
            }

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

                    const wrap = document.createElement('div');
                    wrap.className = `cell-item turno-${turno.toLowerCase()}`;
                    wrap.setAttribute('data-label', tipo.charAt(0).toUpperCase());

                    if (tipo === 'lug') {
                        const w = document.createElement('div'); w.className = 'lug-gru-wrapper'; const key = `${id}-${turno}`;
                        const b = document.createElement('button'); b.className = 'btn-add-gru'; const s = document.createElement('select'); s.disabled = locked; s.add(new Option("-", ""));
                        state.lugares.forEach(o => s.add(new Option(o, o, false, datos[`lug${turno}`] === o)));
                        s.onchange = (e) => guardarManual(id, e.target.value, `lug${turno}`);
                        const gc = document.createElement('div'); gc.className = 'gru-container'; const open = !!state.uiState.openGroups[key]; gc.style.display = open ? 'block' : 'none'; b.textContent = open ? '-' : '+';
                        gc.appendChild(crearSelectorGrupos(datos[`gru${turno}`], locked, (v) => guardarManual(id, v, `gru${turno}`)));
                        b.onclick = () => { state.uiState.openGroups[key] = !state.uiState.openGroups[key]; actualizarTodo(); };
                        w.append(b, s, gc);
                        wrap.appendChild(w);
                    } else if (tipo === 'gru') {
                        wrap.appendChild(crearSelectorGrupos(datos[`gru${turno}`], locked, (v) => guardarManual(id, v, `gru${turno}`)));
                    } else if (tipo === 'cua') {
                        const i = Object.assign(document.createElement('input'), { type: 'text', className: 'input-cuadra', value: datos[`cua${turno}`] || "", disabled: locked });
                        i.onchange = (e) => guardarManual(id, e.target.value, `cua${turno}`);
                        wrap.appendChild(i);
                    } else {
                        const s = document.createElement('select'); s.disabled = locked; s.add(new Option("-", ""));
                        let ops = [];
                        if (tipo === 'cond') {
                            const currentTer = datos[`ter${turno}`];
                            const esCartas = currentTer && currentTer.toLowerCase().includes('cartas');
                            ops = state.conductores.filter(c => {
                                if (esCartas && c.cartas) return true;
                                const disp = c.disponibilidad;
                                const hasAnyConfig = Object.values(disp).some(d => d.m || d.t);
                                if (!hasAnyConfig) return true;
                                return turno === 'M' ? disp[dNum].m : disp[dNum].t;
                            }).map(c => c.name);
                        } else if (tipo === 'ter') {
                            ops = getRangoTerritorios();
                        }
                        if (tipo === 'cond' && datos[`cond${turno}`] && !ops.includes(datos[`cond${turno}`])) ops.push(datos[`cond${turno}`]);
                        ops.forEach(o => s.add(new Option(o, o, false, datos[`${tipo}${turno}`] === o)));
                        s.onchange = (e) => guardarManual(id, e.target.value, `${tipo}${turno}`);
                        wrap.appendChild(s);
                    }
                    cont.appendChild(wrap);
                });
                td.appendChild(cont);
            });
            const tdA = tr.insertCell(); tdA.innerHTML = `<div class="actions-wrapper"><div class="main-actions"><button class="btn-lock">${locked ? '🔒' : '🔓'}</button><button class="btn-assign" ${locked ? 'disabled' : ''} title="Generar para este día">↻</button><button class="btn-note" title="Añadir nota">📝</button><button class="btn-time" title="Editar horario de este día">⏰</button><button class="btn-week" title="Fijar inicio de semana y título de barra">🏷️</button></div><div class="extra-actions">${showM ? '<button class="btn-add-m">+M</button>' : ''}${showT ? '<button class="btn-add-t">+T</button>' : ''}</div></div>`;
            tdA.querySelector('.btn-lock').onclick = () => { state.bloqueados[id] = !state.bloqueados[id]; actualizarTodo(); };
            tdA.querySelector('.btn-assign').onclick = () => { generarCronograma(id); };
            tdA.querySelector('.btn-note').onclick = () => { const n = prompt("Nota:", state.comentarios[id] || ""); if (n !== null) { if (n.trim() === "") delete state.comentarios[id]; else state.comentarios[id] = n; actualizarTodo(); } };

            // Evento para editar horario específico del día
            tdA.querySelector('.btn-time').onclick = () => {
                const curM = hM || '';
                const curT = hT || '';
                const newM = prompt("Horario de Mañana para este día (ej. 08:30, o deja en blanco para usar por defecto):", curM);
                if (newM === null) return;
                const newT = prompt("Horario de Tarde para este día (ej. 16:00, o deja en blanco para usar por defecto):", curT);
                if (newT === null) return;

                if (newM.trim() === '' && newT.trim() === '') {
                    delete state.horariosCustom[id];
                } else {
                    state.horariosCustom[id] = {
                        m: newM.trim(),
                        t: newT.trim()
                    };
                }
                actualizarTodo();
            };

            // Evento para el botón de semana custom
            tdA.querySelector('.btn-week').onclick = () => {
                const current = state.semanasCustom[id] !== undefined ? state.semanasCustom[id] : `SEMANA DEL ${f.getDate()} AL ...`;
                const t = prompt("Título para la barra de semana que empieza este día (deja en blanco para cancelar/quitar):", current);
                if (t !== null) {
                    if (t.trim() === "") delete state.semanasCustom[id];
                    else state.semanasCustom[id] = t;
                    actualizarTodo();
                }
            };

            if (showM) tdA.querySelector('.btn-add-m').onclick = () => agregarFilaExtra(id, 'M');
            if (showT) tdA.querySelector('.btn-add-t').onclick = () => agregarFilaExtra(id, 'T');
            els.tablaBody.appendChild(tr);
            extras.forEach(ex => {
                const trx = document.createElement('tr');
                trx.className = 'fila-extra';
                if (locked) trx.style.opacity = "0.7";

                const tdExH = trx.insertCell();
                tdExH.innerHTML = `<span style="border-left:3px solid ${ex.turno === 'M' ? '#2196F3' : '#FF9800'}; padding-left:5px; font-weight:bold;">${(ex.turno === 'M' ? hM : hT) || '--:--'}</span>`;

                state.colOrder.forEach(t => {
                    const tdEx = trx.insertCell();
                    const wrap = document.createElement('div');
                    wrap.className = `cell-item turno-${ex.turno.toLowerCase()} is-extra`;
                    wrap.setAttribute('data-label', t.charAt(0).toUpperCase());

                    if (t === 'lug') {
                        const w = document.createElement('div');
                        w.className = 'lug-gru-wrapper';
                        const b = document.createElement('button');
                        b.className = 'btn-add-gru';
                        b.textContent = !!state.uiState.openGroups[ex.id] ? '-' : '+';
                        const s = document.createElement('select');
                        s.disabled = locked;
                        s.add(new Option("-", ""));
                        state.lugares.forEach(o => s.add(new Option(o, o, false, ex.lug === o)));
                        s.onchange = (e) => guardarExtra(id, ex.id, 'lug', e.target.value);
                        const gc = document.createElement('div');
                        gc.className = 'gru-container';
                        gc.style.display = !!state.uiState.openGroups[ex.id] ? 'block' : 'none';
                        gc.appendChild(crearSelectorGrupos(ex.gru, locked, (v) => guardarExtra(id, ex.id, 'gru', v)));
                        b.onclick = () => { state.uiState.openGroups[ex.id] = !state.uiState.openGroups[ex.id]; actualizarTodo(); };
                        w.append(b, s, gc);
                        wrap.appendChild(w);
                    } else if (t === 'gru') {
                        wrap.appendChild(crearSelectorGrupos(ex[t], locked, (v) => guardarExtra(id, ex.id, t, v)));
                    } else if (t === 'cua') {
                        const i = Object.assign(document.createElement('input'), { type: 'text', className: 'input-cuadra', value: ex[t] || "", disabled: locked });
                        i.onchange = (e) => guardarExtra(id, ex.id, t, e.target.value);
                        wrap.appendChild(i);
                    } else {
                        const s = document.createElement('select');
                        s.disabled = locked;
                        s.add(new Option("-", ""));
                        let ops = [];
                        if (t === 'cond') {
                            const currentTer = ex['ter'];
                            const esCartas = currentTer && currentTer.toLowerCase().includes('cartas');
                            ops = state.conductores.filter(c => {
                                if (esCartas && c.cartas) return true;
                                const disp = c.disponibilidad;
                                const hasAnyConfig = Object.values(disp).some(d => d.m || d.t);
                                if (!hasAnyConfig) return true;
                                return ex.turno === 'M' ? disp[dNum].m : disp[dNum].t;
                            }).map(c => c.name);
                        } else if (t === 'ter') {
                            ops = getRangoTerritorios();
                        }
                        ops.forEach(o => s.add(new Option(o, o, false, ex[t] === o)));
                        s.onchange = (e) => guardarExtra(id, ex.id, t, e.target.value);
                        wrap.appendChild(s);
                    }
                    tdEx.appendChild(wrap);
                });

                const tdAx = trx.insertCell();
                const bDel = document.createElement('button');
                bDel.innerHTML = 'Eliminar &times;';
                bDel.style.color = 'red';
                bDel.disabled = locked;
                bDel.onclick = () => eliminarFilaExtra(id, ex.id);

                const wrapAction = document.createElement('div');
                wrapAction.className = 'extra-delete-wrap';
                wrapAction.appendChild(bDel);
                tdAx.appendChild(wrapAction);

                els.tablaBody.appendChild(trx);
            });
        });
        const ids = generarFechasDiarias().map(f => f.toISOString().split('T')[0]);
        els.btnBloquear.innerHTML = (ids.length > 0 && ids.every(i => state.bloqueados[i])) ? ' Desbloquear Todo' : ' Bloquear Todo';
    }

    function actualizarTodo() { renderListas(); renderColumnSelector(); renderSelectorDias(); renderConfigTerritorios(); renderTabla(); saveToStorage(); }

    function exportarPDF() {
        const formatTime = (t) => { if (!t || !t.includes(':')) return '--:--'; let [h, m] = t.split(':').map(Number); const ampm = h >= 12 ? 'pm' : 'am'; h = h % 12 || 12; return `${h}:${m < 10 ? '0' + m : m} ${ampm}`; };
        const { jsPDF } = window.jspdf; const doc = new jsPDF('p', 'mm', 'a4'); const fechas = generarFechasDiarias(); if (fechas.length === 0) return alert("Sin fechas.");

        const semanas = [];
        const hasCustomWeeks = fechas.some(f => state.semanasCustom[f.toISOString().split('T')[0]]);
        const defaultTitleFunc = (arr) => `SEMANA DEL ${arr[0].getDate()} AL ${arr[arr.length - 1].getDate()} DE ${arr[arr.length - 1].toLocaleString('es-ES', { month: 'long' }).toUpperCase()}`;

        if (hasCustomWeeks) {
            let currentSem = [];
            let curTitle = "";

            fechas.forEach((f, i) => {
                const id = f.toISOString().split('T')[0];
                if (state.semanasCustom[id]) {
                    if (currentSem.length > 0) {
                        semanas.push({ dias: currentSem, titulo: curTitle || defaultTitleFunc(currentSem) });
                        currentSem = [];
                    }
                    curTitle = state.semanasCustom[id];
                }
                currentSem.push(f);
            });
            if (currentSem.length > 0) {
                semanas.push({ dias: currentSem, titulo: curTitle || defaultTitleFunc(currentSem) });
            }
        } else {
            let sem = [];
            fechas.forEach((f, i) => {
                sem.push(f);
                const sig = fechas[i + 1];
                if (!sig || sig.getDay() === state.startOfWeek) {
                    semanas.push({ dias: [...sem], titulo: defaultTitleFunc(sem) });
                    sem = [];
                }
            });
        }

        const headers = ['DÍA', 'HORA', ...state.colOrder.map(k => COL_DEFS[k].toUpperCase())];

        let currentY = 8;
        doc.setFont("helvetica", "bold");
        doc.setFontSize(11);
        doc.text(`PROGRAMA DE SERVICIO - ${fechas[0].toLocaleString('es-ES', { month: 'long' }).toUpperCase()} ${fechas[0].getFullYear()}`, 105, currentY, { align: 'center' });
        currentY += 3;

        const rows = [];
        semanas.forEach(weekObj => {
            const s = weekObj.dias;
            rows.push([{ content: weekObj.titulo, colSpan: headers.length, styles: { halign: 'center', fillColor: [235, 240, 245], fontStyle: 'bold', fontSize: 9, cellPadding: 0.6 } }]);
            s.forEach(f => {
                const id = f.toISOString().split('T')[0];
                if (state.comentarios[id]) rows.push([{ content: state.comentarios[id].toUpperCase(), colSpan: headers.length, styles: { halign: 'center', fillColor: [255, 253, 208], fontStyle: 'bolditalic', fontSize: 9, cellPadding: 0.6 } }]);
                const d = obtenerDatosDia(id), dNum = f.getDay(), cfg = state.configDias[dNum], esF = (dNum === 0 || dNum === 6);
                const customH = state.horariosCustom[id] || {};
                const hM = (customH.m !== undefined && customH.m !== '') ? customH.m : (esF ? (state.horario.mFin || state.horario.m) : state.horario.m);
                const hT = (customH.t !== undefined && customH.t !== '') ? customH.t : (esF ? (state.horario.tFin || state.horario.t) : state.horario.t);
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

        doc.autoTable({
            head: [headers],
            body: rows,
            startY: currentY,
            theme: 'grid',
            styles: { fontSize: 9, fontStyle: 'bold', cellPadding: 0.4, halign: 'center' },
            headStyles: { fillColor: [44, 62, 80], textColor: 255, cellPadding: 0.6 },
            margin: { left: 5, right: 5, top: 5, bottom: 5 },
            rowPageBreak: 'avoid'
        });
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

        if (els.btnGenerarCronograma) els.btnGenerarCronograma.onclick = () => generarCronograma();

        if (els.btnConfigTer) els.btnConfigTer.onclick = () => { renderConfigTerritorios(); els.modalTer.style.display = 'flex'; };
        if (els.btnCerrarModalTer) els.btnCerrarModalTer.onclick = () => els.modalTer.style.display = 'none';

        document.getElementById('btnLimpiarAsignaciones').onclick = () => { if (confirm("¿Vaciar?")) { generarFechasDiarias().forEach(f => { const id = f.toISOString().split('T')[0]; if (!state.bloqueados[id]) delete state.asignaciones[id]; }); actualizarTodo(); } };
        els.btnBloquear.onclick = () => { const ids = generarFechasDiarias().map(f => f.toISOString().split('T')[0]), all = ids.every(i => state.bloqueados[i]); ids.forEach(i => state.bloqueados[i] = !all); actualizarTodo(); };

        const layout = document.querySelector('.main-layout');
        const sidebar = document.querySelector('.sidebar');
        const sIcon = document.getElementById('sidebarIcon');
        const sText = document.getElementById('sidebarText');

        const updateUI = () => {
            const isH = layout.classList.contains('sidebar-hidden'), isM = window.innerWidth <= 850;
            sIcon.textContent = isM ? (isH ? "⚙️" : "❌") : (isH ? "➡️" : "⬅️");
            sText.textContent = isM ? (isH ? "Configuración" : "Cerrar") : (isH ? "Mostrar" : "Ocultar");
        };

        document.getElementById('toggleSidebar').onclick = (e) => {
            e.stopPropagation();
            layout.classList.toggle('sidebar-hidden');
            updateUI();
            setTimeout(actualizarTodo, 300);
        };

        if (sidebar) {
            sidebar.onclick = (e) => {
                const isMobile = window.innerWidth <= 850;
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
