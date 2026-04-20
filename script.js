const VALID_PIN = "1234"; 
let selectedTicketData = null; 
let currentTicketsList = []; 
let allTicketsList = []; 
let DATOS_HISTORICOS_MESES = []; // Los meses que vienen de Sheets (Ene, Feb...)
let DATOS_MANUAL_ACTUAL = null;  // Los datos que subes por CSV (Marzo...)
let DATOS_NUEVO_MES_HISTORICO = null;
// Variable global para almacenar el cruce de series/fechas
// Declaración única y blindada para todo el script
window.mapaFechasLab = {};

async function cargarDatosLaboratorio() {
    const urlOriginal = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRF7xX_Lf5Xjg6dYfV9y5e6arcmSPOVKmz6DGZAbQw8laxYRsC7V2FiifeoYpoiNd6S-h81aTvNwjq/pub?gid=337853765&single=true&output=csv";
    
    try {
        console.log("Intentando sincronización directa...");
        const response = await fetch(urlOriginal);
        const csvText = await response.text();
        
        if (csvText.includes("<!DOCTYPE")) throw new Error("HTML detectado");

        procesarContenidoCSV(csvText);
        console.log("✅ Sincronización exitosa:", Object.keys(window.mapaFechasLab).length, "series.");
    } catch (err) {
        console.warn("⚠️ Intentando vía Proxy de respaldo...");
        cargarDatosViaProxy(urlOriginal);
    }
}

function cargarDatosViaProxy(urlOriginal) {
    const urlProxy = `https://api.allorigins.win/get?url=${encodeURIComponent(urlOriginal)}`;
    fetch(urlProxy)
        .then(res => res.json())
        .then(data => {
            procesarContenidoCSV(data.contents);
            console.log("✅ Sincronización Proxy exitosa.");
        })
        .catch(err => console.error("❌ Error total de conexión", err));
}

function procesarContenidoCSV(texto) {
    const filas = texto.split(/\r?\n/);
    filas.forEach(fila => {
        const columnas = fila.includes(';') ? fila.split(';') : fila.split(',');
        if (columnas.length >= 2) {
            const serie = columnas[0].replace(/["\r]/g, "").trim().toUpperCase();
            const fecha = columnas[1].replace(/["\r]/g, "").trim();
            if (serie && serie !== "SERIE") {
                window.mapaFechasLab[serie] = fecha;
            }
        }
    });
}

// Ejecutar una sola vez al cargar
cargarDatosLaboratorio();

// === AGREGA ESTO AQUÍ ===
let PROYECTO_ACTIVO = "PJUD5";

const CONFIG_SHEETS = {
    "PJUD5": {
        gidHistorico: "2093520733",
        gidEquipos: "1959700363"
    },
    "PJUD4": {
        gidHistorico: "1085546001", // Reemplaza cuando lo tengas
        gidEquipos: "115813631"     // Reemplaza cuando lo tengas
    }
};

function cambiarProyecto(val) {
    PROYECTO_ACTIVO = val;
    showToast(`Cambiado a ${val}. Cargando datos históricos...`);
    
    // Al cambiar de proyecto, recargamos los datos de las hojas de Google
    cargarDatosDesdeSheets();
    cargarDatosTipos();
    
    // Limpiamos el informe anterior para no confundir datos
    const resultContainer = document.getElementById('mda-result-container');
    if (resultContainer) resultContainer.classList.add('hidden');
}
// =========================

// --- 0. LECTOR INTELIGENTE DE CSV (Faltaba esta función) ---
function parseCSV(text, delimiter = ';') {
    let rows = [];
    let row = [];
    let currentCell = '';
    let insideQuotes = false;

    for (let i = 0; i < text.length; i++) {
        let char = text[i];
        let nextChar = text[i + 1];

        if (char === '"' && insideQuotes && nextChar === '"') {
            currentCell += '"'; 
            i++;
        } else if (char === '"') {
            insideQuotes = !insideQuotes; 
        } else if (char === delimiter && !insideQuotes) {
            row.push(currentCell);
            currentCell = '';
        } else if ((char === '\n' || char === '\r') && !insideQuotes) {
            if (char === '\r' && nextChar === '\n') i++; 
            row.push(currentCell);
            rows.push(row);
            row = [];
            currentCell = '';
        } else {
            currentCell += char;
        }
    }
    if (currentCell || row.length > 0) {
        row.push(currentCell);
        rows.push(row);
    }
    return rows;
}

// --- 1. CONFIGURACIÓN DE TÉCNICOS POR JURISDICCIÓN ---
const TECNICOS_POR_JURISDICCION = {
    "Corte De Apelaciones De Arica": ["BENJAMIN  ZUÑIGA "],
    "Corte De Apelaciones De Iquique": ["MARCELO  CORTEZ "],
    "Corte De Apelaciones De Antofagasta": ["MARIO  MORENO", "ANDERSON  JARA "],
    "Corte De Apelaciones De Copiapo": ["ALVARO  SILVA", "JUAN  SEPULVEDA "],
    "Corte De Apelaciones De Copiapó": ["ALVARO  SILVA", "JUAN SEPULVEDA "],
    "Corte De Apelaciones De La Serena": ["BORIS  REINOSO ", "IGNACIO  SOTOMAYOR "],
    "Corte De Apelaciones De Valparaiso": ["ISMAEL CALQUIN","MAURICIO  TOLEDO ","JORGE  LOPEZ ","MATIAS  INOSTROZA ","JUAN  MANRIQUEZ ","GONZALO  CORTES ","MATIAS  JORQUERA ","BASTIAN  CARDENAS ","VICENTE  CULACIATI ","GERMAN PACHECO ","EDUARDO  DIAZ ","WILSCONIDEL DAUSTKY "],
    "Corte De Apelaciones De Santiago": ["ISMAEL CALQUIN","JORGE  LOPEZ ","MATIAS  INOSTROZA ","JUAN  MANRIQUEZ ","GONZALO  CORTES ","MATIAS  JORQUERA ","BASTIAN  CARDENAS ","VICENTE  CULACIATI ","GERMAN PACHECO ","EDUARDO  DIAZ ","WILSCONIDEL DAUSTKY "],
    "Corte De Apelaciones De San Miguel": ["ISMAEL CALQUIN","JORGE  LOPEZ ","MATIAS  INOSTROZA ","JUAN  MANRIQUEZ ","GONZALO  CORTES ","MATIAS  JORQUERA ","BASTIAN  CARDENAS ","VICENTE  CULACIATI ","GERMAN PACHECO ","EDUARDO  DIAZ ","WILSCONIDEL DAUSTKY "],
    "Corte De Apelaciones De Rancagua": ["ISMAEL CALQUIN","JORGE  LOPEZ ","MATIAS  INOSTROZA ","JUAN  MANRIQUEZ ","GONZALO  CORTES ","MATIAS  JORQUERA ","BASTIAN  CARDENAS ","VICENTE  CULACIATI ","GERMAN PACHECO ","EDUARDO  DIAZ ","WILSCONIDEL DAUSTKY "],
    "Corte De Apelaciones De Talca": [" FERNANDO  ALARCON ", "CRISTIAN  CORDOVA "],
    "Corte De Apelaciones De Chillan": ["MATIAS ALBURQUENQUE "],
    "Corte De Apelaciones De Concepcion": ["ROBINSON  BALLOQUI ", "PABLO  TORRES ","PABLO  RIOS "],
    "Corte De Apelaciones De Concepción": ["ROBINSON  BALLOQUI ", "PABLO  TORRES SALINAS","PABLO  RIOS "],
    "Corte De Apelaciones De Temuco": ["SAUL  BOLIVAR ", "JOSE ALIRO MORALES "],
    "Corte De Apelaciones De Valdivia": ["RODRIGO  SOTOMAYOR "],
    "Corte De Apelaciones De Puerto Montt": ["HECTOR  BARRIENTOS "," DANITZA NOVOA ","GABRIEL CALDERON "],
    "Corte De Apelaciones De Coyhaique": ["ALFREDO ALEJANDRO TORRES ORELLANA"],
    "Corte De Apelaciones De Punta Arenas": ["RICARDO LEONEL GARAY ALTAMIRANO"],
    "Corporacion Administrativa Central": ["JORGE IGNACIO LOPEZ TORRES","MATIAS GIOVANNI INOSTROZA SEPULVEDA","JUAN ANGEL MANRIQUEZ MONTALVA","GONZALO ANDRES CORTES PARRAGUEZ","MATIAS PATRICIO JORQUERA SANTANDER","BASTIAN ISRAEL CARDENAS LEMUÑIR","VICENTE AGUSTIN CULACIATI HERNANDEZ","GERMAN PACHECO MURILLO","EDUARDO JOSE DIAZ MORANTES","WILSCONIDEL DAUSTKY YANEZ LITWIN"],
    "Corte Suprema De Justicia": ["ANGELICA VACCA", "JOSE MARRUFO"],
    "GENERAL": ["Técnico Externo", "Soporte Nivel 2"]
};

// --- 2. CONFIGURACIÓN DE CORREOS ADICIONALES ---
const CORREOS_JURISDICCION = {
    "Corte De Apelaciones De Arica": "rmarancibiav@pjud.cl",
    "Corte De Apelaciones De Iquique": "myanez@pjud.cl",
    "Corte De Apelaciones De Antofagasta": "epenaa@pjud.cl",
    "Corte De Apelaciones De Copiapo": "vquevedoa@pjud.cl",
    "Corte De Apelaciones De Copiapó": "vquevedoa@pjud.cl",
    "Corte De Apelaciones De La Serena": "rsuarez@pjud.cl",
    "Corte De Apelaciones De Valparaiso": "chernandezg@pjud.cl",
    "Corte De Apelaciones De Santiago": "castudillo@pjud.cl",
    "Corte De Apelaciones De San Miguel": "hacevedo@pjud.cl",
    "Corte De Apelaciones De Rancagua": "matoro@pjud.cl;informatica_zonalrancagua@pjud.cl",
    "Corte De Apelaciones De Talca": "maescobar@pjud.cl",
    "Corte De Apelaciones De Chillan": "jparavena@pjud.cl",
    "Corte De Apelaciones De Concepcion": "rsagredo@pjud.cl",
    "Corte De Apelaciones De Concepción": "rsagredo@pjud.cl",
    "Corte De Apelaciones De Temuco": "ebohn@pjud.cl;cquijadah@pjud.cl",
    "Corte De Apelaciones De Valdivia": "adiazm@pjud.cl",
    "Corte De Apelaciones De Puerto Montt": "vhurrutia@pjud.cl",
    "Corte De Apelaciones De Coyhaique": "mmoraless@pjud.cl",
    "Corte De Apelaciones De Punta Arenas": "hhernandez@pjud.cl",
    "Corporacion Administrativa Central": "dvillarroels@pjud.cl",
    "Corte Suprema De Justicia": "dvillarroels@pjud.cl",
    "GENERAL": "soporte.central@fcom.cl"
};

const CORREOS_FIJOS = "c.zapata@fcom.cl; a.vacca@fcom.cl; s.guzman@fcom.cl; dvillarroels@pjud.cl; j.marrufo@fcom.cl;s.valbuena@fcom.cl"

// --- 3. LÓGICA DE ACCESO ---
function checkPin() {
    const pinInput = document.getElementById('pin-input').value;
    if (pinInput === VALID_PIN) {
        if(document.getElementById('login-screen')) document.getElementById('login-screen').classList.add('hidden');
        document.getElementById('main-interface').classList.remove('hidden');
    } else {
        if(document.getElementById('error-msg')) document.getElementById('error-msg').innerText = "PIN Incorrecto";
    }
}

function showModule(modName) {
    document.querySelectorAll('.module').forEach(m => m.classList.add('hidden'));
    document.querySelectorAll('.nav-links li').forEach(li => li.classList.remove('active'));
    document.getElementById('mod-' + modName).classList.remove('hidden');
}


// --- 4. PROCESAMIENTO DE DATOS PRINCIPAL ---
function processData() {

const finalizadosHTML = `
    <div class="card" style="border: none; padding: 0; background: transparent; margin-bottom: 20px;">
        <div class="pjud-header-toggle" onclick="toggleSection('finalizados-body', 'icon-fin')" style="border-top: 4px solid #014f8b; cursor: pointer; display: flex; justify-content: space-between; align-items: center; background: white; padding: 10px 15px; border: 1px solid #ddd; border-bottom: none; border-radius: 5px 5px 0 0;">
            <h3 style="color: #014f8b; margin: 0; font-size: 1.1rem;">
                <i class="fas fa-check-double"></i> Requerimientos Finalizados en el día
            </h3>
            <i id="icon-fin" class="fas fa-chevron-down rotate-icon"></i>
        </div>
        
        <div id="finalizados-body" style="background: white; border: 1px solid #ddd; padding: 0; border-radius: 0 0 5px 5px; display: none; 
             max-height: 350px; overflow-y: auto; overflow-x: hidden;">
            
            <table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 11px;">
                <thead style="position: sticky; top: 0; z-index: 10; background: #f8f9fa;">
                    <tr style="background-color: #f8f9fa; color: #333; text-align: left; text-transform: uppercase;">
                        <th style="width: 8%; padding: 10px; border: 1px solid #eee;">Proyecto</th>
                        <th style="width: 15%; padding: 10px; border: 1px solid #eee;">Ticket</th>
                        <th style="width: 12%; padding: 10px; border: 1px solid #eee; text-align: center;">SLA Total</th>
                        <th style="width: 10%; padding: 10px; border: 1px solid #eee;">Tipo</th>
                        <th style="width: 25%; padding: 10px; border: 1px solid #eee;">Dependencia</th>
                        <th style="width: 20%; padding: 10px; border: 1px solid #eee;">Técnico</th>
                        <th style="width: 10%; padding: 10px; border: 1px solid #eee; text-align: center;">Hora Fin</th>
                    </tr>
                </thead>
                <tbody id="bodyFinalizadosDia"></tbody>
            </table>
        </div>
    </div>
`;

// Inyectamos el cascarón en el contenedor (asegúrate de tener <div id="container-finalizados"></div> en tu HTML)
const container = document.getElementById('container-finalizados');
if(container) container.innerHTML = finalizadosHTML;

    const fileInput = document.getElementById('csv-file');
    const filterElement = document.getElementById('group-filter');
    const selectedGroup = filterElement ? filterElement.value : "SCO"; 

    if (!fileInput.files || fileInput.files.length === 0) {
        showToast("⚠️ Por favor adjunta un archivo CSV o TXT.");
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

   reader.onload = function(e) {
        const rawData = e.target.result;
        const rows = rawData.split('\n'); 
        
        const tableBody = document.querySelector('#tickets-table tbody');
        if(tableBody) tableBody.innerHTML = ''; 
        currentTicketsList = []; 
        allTicketsList = []; 

        let savedCoords = {};
        try {
            savedCoords = JSON.parse(localStorage.getItem('sco_coordinations') || '{}');
        } catch(err) {
            console.warn("Memoria caché bloqueada por el navegador.");
        }

rows.forEach((row) => {
    // 1. Primero dividimos la fila en columnas
    const cols = row.split(';');

    // 2. Validamos que la fila tenga datos
    if (cols.length < 5) return;

    // 3. Definimos la función de limpieza para poder usarla
    const cleanCol = (idx) => cols[idx] ? cols[idx].replace(/"/g, "").trim() : "";

    // 4. Extraemos datos básicos necesarios
    const minutosConsumidos = parseInt(cleanCol(26)) || 0; // Columna AA
    // --- NUEVO CÁLCULO DE TIEMPO HUMANO (Horas y Minutos) ---
const minutosTotales = parseInt(cleanCol(26)) || 0; 
const horasEnteras = Math.floor(minutosTotales / 60);
const minutosRestantes = minutosTotales % 60;

// Creamos un texto amigable: "1h 20m"
const tiempoFormateado = `${horasEnteras}h ${minutosRestantes}m`;

// Para los cálculos matemáticos del semáforo seguimos usando el valor decimal interno
const horasDecimales = minutosTotales / 60;
    const dependenciaTexto = cleanCol(12).toUpperCase();
    const grupoResolutor = cleanCol(3).toUpperCase(); // Columna Grupo

    // --- LÓGICA SEMÁFORO SLA CON FORMATO HUMANO ---
let limiteHorasSLA = 0;

if (grupoResolutor.includes("RESIDENTES")) {
    limiteHorasSLA = 2;
} else {
    const infoComuna = DATA_MATRIZ.find(m => dependenciaTexto.includes(m.comuna.toUpperCase()));
    limiteHorasSLA = infoComuna ? parseInt(infoComuna.slahardware) : 0;
}

let slaStatusHTML = "";

if (limiteHorasSLA > 0) {
    const porcentajeUso = (horasDecimales / limiteHorasSLA);

    if (horasDecimales >= limiteHorasSLA) {
        // VENCIDO
        slaStatusHTML = `
            <div style="color: #d32f2f; font-weight: bold;" title="Límite: ${limiteHorasSLA}h">
                <i class="fas fa-exclamation-triangle"></i> Fuera de SLA<br>
                <small>${tiempoFormateado} / ${limiteHorasSLA}h</small>
            </div>`;
    } else if (porcentajeUso >= 0.7) {
        // CRÍTICO
        slaStatusHTML = `
            <div style="color: #ef6c00; font-weight: bold;" title="Límite: ${limiteHorasSLA}h">
                <i class="fas fa-clock"></i> Crítico (70%)<br>
                <small>${tiempoFormateado} / ${limiteHorasSLA}h</small>
            </div>`;
    } else {
        // EN TIEMPO
        slaStatusHTML = `
            <div style="color: #2e7d32;" title="Límite: ${limiteHorasSLA}h">
                <i class="fas fa-check-circle"></i> En Tiempo<br>
                <small>${tiempoFormateado} / ${limiteHorasSLA}h</small>
            </div>`;
    }
} else {
    slaStatusHTML = `<span style="color: #9e9e9e;">N/A</span>`;
}

    // 5. Procesamiento de Solución (Tu lógica original)
    let rawSolucion = cols[23] || "REVISIÓN"; 
    let solucionLimpia = rawSolucion;
    const textoUpper = rawSolucion.toUpperCase();

    if (textoUpper.includes("MASTERIZAC") || textoUpper.includes("MAQUETADO") || textoUpper.includes("IMAGEN")) {
        solucionLimpia = "MASTERIZACIÓN";
    } else if (textoUpper.includes("CAMBIO") && (textoUpper.includes("EQUIPO") || textoUpper.includes("PC") || textoUpper.includes("NOTEBOOK"))) {
        solucionLimpia = "CAMBIO EQUIPO";
    } else if (textoUpper.includes("RETIRO")) {
        solucionLimpia = "RETIRO";
    } else if (textoUpper.includes("CONFIGURA") || textoUpper.includes("PERFIL")) {
        solucionLimpia = "CONFIGURACIÓN";
    } else if (rawSolucion.length > 50) {
        solucionLimpia = "SOPORTE TÉCNICO / REVISIÓN";
    }

    const ticketNum = cleanCol(1);
    let rawIP = (cols[11] && cols[11].includes('.')) ? cols[11].replace(/"/g, "").trim() : "-";

    // 6. Creación del objeto Ticket
    const ticket = {
        num: ticketNum,
        estado: cleanCol(2),
        grupo: cleanCol(3),
        finalizadoPor: cleanCol(5), 
        usuario: cleanCol(8),
        correo: cleanCol(10), 
        ip: rawIP,
        dependencia: cleanCol(12),
        jurisdiccion: cleanCol(13) || "GENERAL",
        direccion: cleanCol(14) || "Sin dirección",
        modelo: cleanCol(15) || "Modelo N/A", 
        proyecto: cleanCol(16), 
        tipo: cleanCol(17) || "Equipo",       
        serie: cleanCol(18) || "Serie N/A",
        fechaFin: cleanCol(24), 
        backup: cleanCol(27) ? cleanCol(27).toUpperCase() : "NO",
        guiaRetiro: cleanCol(28) ? cleanCol(28).toUpperCase() : "NO", 
        solucion: solucionLimpia, 
        despachosRaw: cleanCol(33) || cleanCol(27),
        solucionMDARaw: cleanCol(22),
        fechaCreacion: cleanCol(6),
        AsignadoA: cleanCol(4),
        slaHTML: slaStatusHTML,
        horaFinalizacion: cleanCol(25), // Columna Z
        fechaCreacion: cleanCol(6),
    };

    // 7. Lógica de tickets coordinados (Cache)
    if (savedCoords[ticketNum]) {
        ticket.fechaCoord = savedCoords[ticketNum].fechaCoord;
        ticket.horaCoord = savedCoords[ticketNum].horaCoord;
        ticket.tecnicoCoord = savedCoords[ticketNum].tecnicoCoord;
    }

    allTicketsList.push(ticket);

    // 8. Filtrado para mostrar en tabla
    const estadoUpper = ticket.estado.toUpperCase();
    const isPending = estadoUpper !== "CERRADO" && estadoUpper !== "FINALIZADO";
    let isTargetGroup = (selectedGroup === "TODOS") ? true : ticket.grupo.toUpperCase().includes(selectedGroup);

    if (isTargetGroup && isPending) {
        const ticketIndex = currentTicketsList.length;
        currentTicketsList.push(ticket);

        let statusStyle = "";
        if (estadoUpper.includes("COORDINADO")) statusStyle = "background-color: #d1e7dd; color: #0f5132; border: 1px solid #badbcc;"; 
        else if (estadoUpper.includes("EN PROCESO")) statusStyle = "background-color: #fff3cd; color: #664d03; border: 1px solid #ffecb5;"; 
        else statusStyle = "background-color: #f8d7da; color: #842029; border: 1px solid #f5c2c7;"; 

        let fechaDisplay = "-";
        if (ticket.fechaCoord) {
            const partes = ticket.fechaCoord.split('-');
            if (partes.length === 3) {
                fechaDisplay = `<span style="color:#007bff; font-weight:bold;">${partes[2]}-${partes[1]}-${partes[0]}</span><br><small>${ticket.horaCoord || ''}</small>`;
            }
        }
        const tecnicoDisplay = ticket.tecnicoCoord ? `<span style="font-size:0.85em; font-weight:600;">${ticket.tecnicoCoord}</span>` : "-";

        // 9. Pintado de la fila final
        // 9. Pintado de la fila corregido
const tr = document.createElement('tr');
tr.innerHTML = `
    <td style="color: var(--hp-blue); font-weight: bold;">${ticket.proyecto || 'N/A'}</td>
    <td><b>${ticket.num}</b></td>
    
    <td style="text-align:center; border-left: 1px solid #eee; background-color: #fcfcfc;">
        ${ticket.slaHTML}
    </td>

    <td><span class="status-pill" style="${statusStyle}">${ticket.estado}</span></td>
    
    <td style="font-size: 0.8rem; color: #666;">${ticket.grupo}</td>
    
    <td style="font-size: 0.8rem; color: #333; font-weight: 600;">${ticket.tipo}</td> 
    
    <td>
        ${ticket.dependencia} <br>
        <small style="color: #999;">(${ticket.jurisdiccion})</small>
    </td>
    
    <td style="text-align:center; background-color: #f9fcff;">${fechaDisplay}</td>
    
    <td style="text-align:center; background-color: #f9fcff;">${tecnicoDisplay}</td>
    
    <td style="text-align:center;">${ticket.backup}</td>
    
    <td>
        <div style="display: flex; gap: 5px; justify-content: center;">
            <button class="btn-primary" style="padding: 5px 10px; font-size: 0.8rem;" onclick="openCoordEditor(${ticketIndex})">Gestionar</button>
            <button class="btn-secondary" style="padding: 5px 10px; font-size: 0.8rem; background-color: #dc3545; color: white; border: none; border-radius: 4px; cursor: pointer;" onclick="clearTicketCoordination('${ticket.num}', this)" title="Limpiar Coordinación">
                <i class="fas fa-eraser"></i>
            </button>
        </div>
    </td>
`;
        const tableBody = document.querySelector('#tickets-table tbody');
        if(tableBody) tableBody.appendChild(tr);
    }
});

        if (allTicketsList.length > 0) {
            if (currentTicketsList.length > 0) showToast(`Cargados: ${currentTicketsList.length} tickets de ${selectedGroup}.`);
            else showToast(`Cargados ${allTicketsList.length} registros totales.`);
        } else {
            showToast("⚠️ No se encontraron datos válidos en el archivo.");
        }

// --- Llenado de la tabla de Finalizados ---
// --- LÓGICA PARA TABLA DE FINALIZADOS DEL DÍA (COMPLETA Y AJUSTADA) ---
const tbodyFinalizados = document.getElementById('bodyFinalizadosDia');
if (tbodyFinalizados) {
    tbodyFinalizados.innerHTML = ""; // Limpiar tabla antes de llenar

    // 1. Obtener la fecha de hoy en formato DD/MM/YYYY
    const d = new Date();
    const hoyFormateado = String(d.getDate()).padStart(2, '0') + "/" + 
                          String(d.getMonth() + 1).padStart(2, '0') + "/" + 
                          d.getFullYear();

    // 2. Obtener el grupo seleccionado actualmente en el filtro superior
    const grupoFiltroActual = document.getElementById('group-filter').value.toUpperCase();

    // 3. Filtrar tickets: Solo Hoy + Solo Cerrados/Finalizados + Solo el Grupo seleccionado
    const finalizadosHoy = allTicketsList.filter(t => {
        const est = t.estado.toUpperCase();
        const grp = t.grupo.toUpperCase();
        
        // Extraemos solo la fecha del campo fechaCreacion (asumiendo formato "DD/MM/YYYY HH:MM")
        const fechaTicket = t.fechaFin.split(' ')[0]; 

        const esDeHoy = (fechaTicket === hoyFormateado);
        const esCerrado = (est === "FINALIZADO" || est === "CERRADO");
        const cumpleGrupo = (grupoFiltroActual === "TODOS") ? true : grp.includes(grupoFiltroActual);
        
        return esDeHoy && esCerrado && cumpleGrupo;
    });

    // 4. Si no hay tickets, mostrar mensaje informativo
    if (finalizadosHoy.length === 0) {
        tbodyFinalizados.innerHTML = `
            <tr>
                <td colspan="7" style="text-align:center; padding:20px; color:#999; font-style: italic;">
                    No hay requerimientos finalizados el día de hoy (${hoyFormateado}) para el área seleccionada.
                </td>
            </tr>`;
    } else {
        // 5. Generar las filas de la tabla
        finalizadosHoy.forEach(t => {
            const tr = document.createElement('tr');
            tr.style.borderBottom = "1px solid #eee";
            
            tr.innerHTML = `
                <td style="padding: 8px 10px; border: 1px solid #eee; color: #014f8b; font-weight: bold;">${t.proyecto}</td>
                <td style="padding: 8px 10px; border: 1px solid #eee;">
                    <div style="display: flex; align-items: center; justify-content: space-between; gap: 5px;">
                        <span>${t.num}</span>
                        <button onclick="copiarTexto('${t.num}', this)" 
                                style="border: none; background: none; cursor: pointer; color: #aaa; padding: 2px;" 
                                title="Copiar Ticket">
                            <i class="far fa-copy"></i>
                        </button>
                    </div>
                </td>
                <td style="padding: 8px 10px; border: 1px solid #eee; text-align: center;">${t.slaHTML || '-'}</td>
                <td style="padding: 8px 10px; border: 1px solid #eee; color: #666;">${t.tipo}</td>
                <td style="padding: 8px 10px; border: 1px solid #eee; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${t.dependencia}">
                    ${t.dependencia}
                </td>
                <td style="padding: 8px 10px; border: 1px solid #eee; font-size: 0.85rem;">${t.finalizadoPor || t.tecnicoCoord || '-'}</td>
                <td style="padding: 8px 10px; border: 1px solid #eee; text-align: center; font-weight: bold; color: #28a745;">
                    ${t.horaFinalizacion || '-'}
                </td>
            `;
            tbodyFinalizados.appendChild(tr);
        });

        // Opcional: Auto-desplegar la sección si hay tickets finalizados
        const bodyFin = document.getElementById('finalizados-body');
        const iconFin = document.getElementById('icon-fin');
        if (bodyFin && iconFin) {
            bodyFin.style.display = "block";
            iconFin.style.transform = "rotate(180deg)";
        }
    }
}
        
    };

    reader.readAsText(file, "ISO-8859-1");





    
}

function openCoordEditor(index) {
    const data = currentTicketsList[index];
    selectedTicketData = data;
    
    document.getElementById('text-preview').innerHTML = ''; 
    document.getElementById('result-container').classList.add('hidden');
    document.getElementById('editing-title').innerText = `Gestión: ${data.num} (${data.dependencia})`;
    
    const selectTech = document.getElementById('tech-name');
    selectTech.innerHTML = '<option value="">Seleccione un técnico...</option>';
    
    const listaTecnicos = TECNICOS_POR_JURISDICCION[data.jurisdiccion] || TECNICOS_POR_JURISDICCION["GENERAL"];
    listaTecnicos.forEach(pro => {
        const option = document.createElement('option');
        option.value = pro;
        option.text = pro;
        selectTech.appendChild(option);
    });

    document.getElementById('coord-date').value = data.fechaCoord || "";
    document.getElementById('coord-time').value = data.horaCoord || "";
    if (data.tecnicoCoord) selectTech.value = data.tecnicoCoord;

    const editor = document.getElementById('coord-editor');
    editor.classList.remove('hidden');
    
    const body = document.getElementById('editor-body');
    if(body && body.classList.contains('hidden-content')) {
        body.classList.remove('hidden-content');
        document.getElementById('icon-editor').classList.remove('rotate-icon');
    }

    editor.scrollIntoView({ behavior: 'smooth' });
}


// --- 5. FUNCIÓN GUARDAR ACTUALIZADA ---
function saveTicketCoordination() {
    if (!selectedTicketData) {
        showToast("⚠️ No hay ningún ticket seleccionado.");
        return;
    }

    const fechaInput = document.getElementById('coord-date').value;
    const horaInput = document.getElementById('coord-time').value;
    const techInput = document.getElementById('tech-name').value;

    if (!fechaInput || !techInput) {
        showToast("⚠️ Debes seleccionar Fecha y Técnico para guardar.");
        return;
    }

    // Actualizamos el objeto en memoria
    selectedTicketData.fechaCoord = fechaInput;
    selectedTicketData.horaCoord = horaInput;
    selectedTicketData.tecnicoCoord = techInput;

    // 1. Guardar en LocalStorage para persistencia
    try {
        let savedCoords = JSON.parse(localStorage.getItem('sco_coordinations') || '{}');
        savedCoords[selectedTicketData.num] = {
            fechaCoord: fechaInput,
            horaCoord: horaInput,
            tecnicoCoord: techInput
        };
        localStorage.setItem('sco_coordinations', JSON.stringify(savedCoords));
    } catch (e) {
        console.warn("Memoria caché bloqueada, guardado solo por sesión.", e);
    }

    // 2. Actualizar la fila en la tabla visualmente
    try {
        const tableBody = document.querySelector('#tickets-table tbody');
        const filas = tableBody.querySelectorAll('tr');
        let encontrado = false;

        for (let i = 0; i < filas.length; i++) {
            const fila = filas[i];
            // El número de ticket está en la segunda celda (índice 1)
            const celdaTicket = fila.children[1].innerText.trim();
            
            if (celdaTicket === selectedTicketData.num.trim()) {
                const partes = fechaInput.split('-'); 
                const fechaFormateada = partes.length === 3 ? `${partes[2]}-${partes[1]}-${partes[0]}` : fechaInput;
                
                // --- ÍNDICES CORREGIDOS PARA LA NUEVA ESTRUCTURA ---
                // [7] es la columna FECHA/HORA
                // [8] es la columna TÉCNICO
                fila.children[7].innerHTML = `
                    <span style="color:#007bff; font-weight:bold;">${fechaFormateada}</span><br>
                    <small>${horaInput}</small>
                `;
                
                fila.children[8].innerHTML = `
                    <span style="font-size:0.85em; font-weight:600;">${techInput}</span>
                `;
                
                // Efecto visual de guardado exitoso
                fila.style.backgroundColor = "#d4edda";
                setTimeout(() => fila.style.backgroundColor = "", 800);
                encontrado = true;
                break;
            }
        }

        if (encontrado) {
            showToast(`✅ Guardado exitoso para TK ${selectedTicketData.num}`);
            // Opcional: Cerrar el modal automáticamente tras guardar
            closeCoordModal(); 
        } else {
            showToast("✅ Guardado en memoria.");
        }
    } catch (e) {
        console.error("Error pintando la tabla:", e);
        showToast("✅ Guardado exitosamente.");
    }
}


// --- 6. GENERACIÓN DE MENSAJES (CORREO / WHATSAPP) ---
function generateEmail() {
    if (!selectedTicketData) return;

    const tech = document.getElementById('tech-name').value || "PENDIENTE";
    const date = document.getElementById('coord-date').value || "POR DEFINIR";
    const time = document.getElementById('coord-time').value || "POR DEFINIR";
    
    const esCambio = (selectedTicketData.backup === "SI");
    const actividad = esCambio ? "CAMBIO DE EQUIPO" : "REVISIÓN / CONFIGURACION EQUIPO";

    const para = selectedTicketData.correo || "";
    const correoJurisdiccion = CORREOS_JURISDICCION[selectedTicketData.jurisdiccion] || CORREOS_JURISDICCION["GENERAL"];
    let cc = `${CORREOS_FIJOS}; ${correoJurisdiccion}`;
    
    if (esCambio) {
        const correosExtra = "alonconu@pjud.cl;rmanriquezl@pjud.cl;ariveros@pjud.cl;vperezc@pjud.cl;cescobarz@pjud.cl;lrcastro@pjud.cl";
        cc += `; ${correosExtra}`;
    }

    const proyecto = (selectedTicketData.proyecto || "").toUpperCase();
    if (proyecto.includes("PJUD 4")) cc += "; carol.oteiza@hp.com";
    else if (proyecto.includes("PJUD 5")) cc += "; christian.ojeda@hp.com";

    const asunto = `Requerimiento ${selectedTicketData.num} - Coordinación SCO - Proyecto ${selectedTicketData.proyecto}`;
    const nombreCompleto = selectedTicketData.usuario || "Usuario";
    const primerNombre = nombreCompleto.trim().split(' ')[0];

    let serieDespachada = "Pendiente / No registrada";
    const rawDespacho = selectedTicketData.despachosRaw || "";
    
    if (rawDespacho.includes(":")) {
        const parts = rawDespacho.split(":");
        if (parts.length >= 2) {
            serieDespachada = parts[1].trim(); 
        }
    }

    const tablaEstilizada = `
    <table style="border-collapse: collapse; width: 100%; max-width: 350px; font-family: Segoe UI, Calibri, Arial, sans-serif; border: 1px solid #014f8b; font-size: 11px;">
        <thead><tr><th colspan="2" style="background-color: #014f8b; color: #ffffff !important; padding: 5px; text-align: center; font-size: 12px;">DETALLE DE COORDINACIÓN</th></tr></thead>
        <tbody>
            <tr style="background-color: #f2f2f2;"><td style="padding: 4px 6px; border: 1px solid #ddd; font-weight: bold; width: 35%;">TICKET</td><td style="padding: 4px 6px; border: 1px solid #ddd;">${selectedTicketData.num}</td></tr>
            <tr><td style="padding: 4px 6px; border: 1px solid #ddd; font-weight: bold;">PROYECTO</td><td style="padding: 4px 6px; border: 1px solid #ddd;">${selectedTicketData.proyecto}</td></tr>
            <tr style="background-color: #f2f2f2;"><td style="padding: 4px 6px; border: 1px solid #ddd; font-weight: bold;">DEPENDENCIA</td><td style="padding: 4px 6px; border: 1px solid #ddd;">${selectedTicketData.dependencia}</td></tr>
            <tr><td style="padding: 4px 6px; border: 1px solid #ddd; font-weight: bold;">ACTIVIDAD</td><td style="padding: 4px 6px; border: 1px solid #ddd;">${actividad}</td></tr>
            
            <tr style="background-color: #f2f2f2;"><td style="padding: 4px 6px; border: 1px solid #ddd; font-weight: bold;">TIPO EQUIPO</td><td style="padding: 4px 6px; border: 1px solid #ddd;">${selectedTicketData.tipo}</td></tr>
            <tr><td style="padding: 4px 6px; border: 1px solid #ddd; font-weight: bold;">SERIE REPORTADA</td><td style="padding: 4px 6px; border: 1px solid #ddd;">${selectedTicketData.serie}</td></tr>
            
            <tr><td style="padding: 4px 6px; border: 1px solid #ddd; font-weight: bold;">FECHA</td><td style="padding: 4px 6px; border: 1px solid #ddd;">${date}</td></tr>
            <tr style="background-color: #f2f2f2;"><td style="padding: 4px 6px; border: 1px solid #ddd; font-weight: bold;">HORA</td><td style="padding: 4px 6px; border: 1px solid #ddd;">${time}</td></tr>
            <tr><td style="padding: 4px 6px; border: 1px solid #ddd; font-weight: bold;">TÉCNICO</td><td style="padding: 4px 6px; border: 1px solid #ddd;">${tech}</td></tr>
        </tbody>
    </table>`;

    const htmlResultado = `
        <div class="card" style="border: none; padding: 0; background: transparent;">
            <div class="pjud-header-toggle" onclick="toggleSection('email-body', 'icon-email')">
                <h3>Gestión de Correo</h3>
                <i id="icon-email" class="fas fa-chevron-up rotate-icon"></i>
            </div>
            <div id="email-body" style="background: white; border: 1px solid #ddd; padding: 20px; border-radius: 0 0 5px 5px;">
                <div class="mail-row" style="margin-bottom: 12px;">
                    <button class="btn-copy-small" onclick="copyId('mail-to')">¡Copiar!</button>
                    <span id="mail-to" style="font-size: 0.85rem; color: #444;">${para}</span>
                </div>
                <div class="mail-row" style="margin-bottom: 12px;">
                    <button class="btn-copy-small" onclick="copyId('mail-cc')">¡Copiar!</button>
                    <span id="mail-cc" style="font-size: 0.85rem; color: #444;">${cc}</span>
                </div>
                <div class="mail-row" style="margin-bottom: 12px;">
                    <button class="btn-copy-small" onclick="copyId('mail-sub')">¡Copiar!</button>
                    <span id="mail-sub" style="font-size: 0.85rem;">${asunto}</span>
                </div>
                
                <div id="email-full-content" style="background: white; padding: 15px; border: 1px solid #eee; margin-bottom: 15px;">
                    <p style="font-family: Calibri, sans-serif; font-size: 13px;">Estimado(a) ${primerNombre},<br><br>
                    Junto con saludar, informamos que se ha coordinado la atención de su requerimiento con el siguiente detalle:<br><br></p>
                    ${tablaEstilizada}
                    <br>
                    <p style="font-family: Calibri, sans-serif; font-size: 11px; color: #777; margin-top: 15px; line-height: 1.2;">
                        <strong>Nota:</strong> Durante los procesos de Cambio/Masterización de Notebook y/o PC pueden surgir imprevistos que impidan el respaldo parcial o completo de la data del equipo. Se sugiere realizar un respaldo previo. HP no realiza gestiones de recuperación de data.
                    </p>
                </div>

                <div class="button-actions-group">
                    <button onclick="enviarYCopiar('${encodeURIComponent(para)}','${encodeURIComponent(cc)}','${encodeURIComponent(asunto)}', 'email-full-content')" class="btn-outlook">
                        <i class="fas fa-paper-plane"></i> 1. Copiar y Enviar Outlook
                    </button>
                    <button onclick="window.open('https://outlook.office.com/mail/deeplink/compose', '_blank')" class="btn-secondary" style="flex:1; background: #0078d4; color: white; border:none;">
                        <i class="fas fa-globe"></i> 2. Ir a Outlook Web
                    </button>
                    <button onclick="copyTable('email-full-content')" class="btn-copy-body">
                        <i class="fas fa-copy"></i> 3. Solo Copiar Texto
                    </button>
                </div>
            </div>
        </div>`;

    mostrarResultado("Gestión de Correo", htmlResultado, true);
    setTimeout(() => copyTable('email-full-content'), 100);
}

function generateWhatsApp() {
    if (!selectedTicketData) return;

    const tech = document.getElementById('tech-name').value || "PENDIENTE";
    const date = document.getElementById('coord-date').value || "POR DEFINIR";
    const time = document.getElementById('coord-time').value || "POR DEFINIR";
    
    const esCambio = (selectedTicketData.backup === "SI");
    const actividad = esCambio ? "CAMBIO DE EQUIPO" : "REVISIÓN / CONFIGURACION EQUIPO";
    const dir = selectedTicketData.direccion || "No especificada";
    
    const tipoEquipo = selectedTicketData.tipo || "No especificado";
    const serieReportada = selectedTicketData.serie || "No registrada";
    
    let serieDespachada = "Pendiente Validar";
    if (selectedTicketData.despachosRaw && selectedTicketData.despachosRaw.trim() !== "") {
        serieDespachada = selectedTicketData.despachosRaw.trim();
    }

    const mensaje = `*TK:* ${selectedTicketData.num}\n🧰 *Tecnico:* ${tech}\n📀 *Proyecto:* ${selectedTicketData.proyecto}\n🏛️ *Tribunal:* ${selectedTicketData.dependencia}\n🏛️ *Direccion:* ${dir}\n🗓️ *Fecha Coordinada:* ${date}\n⏰ *Hora:* ${time}\n📝 *Actividad:* ${actividad}\n💻 *Tipo Equipo:* ${tipoEquipo}\n🏷️ *Serie Reportada:* ${serieReportada}\n📦 *Serie Despachada:* ${serieDespachada}\n🚨 *OBLIGATORIO:* COLOCAR LA IP EN TODAS LAS ATENCIONES\n\n⚠️ *Nota:* SI ES UN CAMBIO DE MULTIFUNCIONAL FAVOR COMUNICARSE CON EL ADM DE IMPRESION (JOSE CHAVEZ) ANTES DE DESCONECTAR LA MULTIFUNCIONAL SALIENTE. ENTREGAR DIRECCION IP.
`;
    
    const htmlWhatsApp = `
        <div class="card" style="border: none; padding: 0; background: transparent; margin-bottom: 20px;">
             <div class="pjud-header-toggle" onclick="toggleSection('wa-body', 'icon-wa')" style="border-top: 4px solid #25D366;">
                <h3 style="color: #0f5132;">Mensaje Técnico (WhatsApp)</h3>
                <i id="icon-wa" class="fas fa-chevron-up rotate-icon"></i>
            </div>
            <div id="wa-body" style="background: white; border: 1px solid #ddd; padding: 20px; border-radius: 0 0 5px 5px;">
                <pre id="copy-area-wa" style="background:#e9f7ef; border:1px solid var(--success-green); padding:15px; border-radius:8px; white-space:pre-wrap; min-height: 250px; max-height: none; font-family: inherit;">${mensaje}</pre>
                <div class="button-actions-group">
                    <button onclick="enviarWA('${encodeURIComponent(mensaje)}')" class="btn-whatsapp">
                        <i class="fab fa-whatsapp"></i> Enviar WhatsApp
                    </button>
                    <button onclick="copyId('copy-area-wa')" class="btn-copy-body">
                        <i class="fas fa-copy"></i> Copiar Texto
                    </button>
                </div>
            </div>
        </div>
    `;
    mostrarResultado("WhatsApp para Técnico", htmlWhatsApp, true);
}


// --- 7. UTILIDADES Y TOGGLE ---
function toggleSection(contentId, iconId) {
    const content = document.getElementById(contentId);
    const icon = document.getElementById(iconId);
    if(content) content.classList.toggle('hidden-content');
    if(icon) icon.classList.toggle('rotate-icon');
}

function enviarWA(msg) {
    window.open(`https://api.whatsapp.com/send?text=${msg}`, '_blank');
}

function copyId(id) {
    const text = document.getElementById(id).innerText;
    navigator.clipboard.writeText(text).then(() => showToast("¡Copiado con éxito!"));
}

function copyTable(id) {
    const el = document.getElementById(id);
    const range = document.createRange();
    range.selectNode(el);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    try {
        document.execCommand('copy');
        showToast("¡Tabla copiada con formato!");
    } catch (e) {
        showToast("Error al copiar");
    }
    sel.removeAllRanges();
}

function showToast(mensaje) {
    const toast = document.getElementById('toast-container');
    const msgSpan = document.getElementById('toast-message');
    if(!toast) return;
    msgSpan.innerText = mensaje;
    toast.classList.remove('toast-hidden');
    setTimeout(() => toast.classList.add('toast-hidden'), 2500);
}

function mostrarResultado(titulo, contenido, esHTML) {
    document.getElementById('result-type-title').style.display = 'none';
    const preview = document.getElementById('text-preview');
    if (esHTML) preview.innerHTML = contenido; 
    else preview.innerText = contenido;
    const container = document.getElementById('result-container');
    container.classList.remove('hidden');
    container.scrollIntoView({ behavior: 'smooth' });
}

function clearData() {
    document.getElementById('csv-file').value = ''; 
    const tb = document.querySelector('#tickets-table tbody');
    if(tb) tb.innerHTML = '';
    document.getElementById('coord-editor').classList.add('hidden');
    document.getElementById('result-container').classList.add('hidden');
    document.getElementById('text-preview').innerHTML = '';
    currentTicketsList = [];
    allTicketsList = [];
    showToast("Datos limpiados");
}

function intentarAbrirOutlook(para, cc, asunto) {
    const mailto = `mailto:${para}?cc=${cc}&subject=${asunto}`;
    const start = Date.now();
    window.location.href = mailto;
    setTimeout(() => {
        if (Date.now() - start < 1000) showToast("Abriendo Outlook.");
    }, 500);
}

// --- MÓDULO GUÍAS ---
function buscarTicketParaGuia() {
    const tkInput = document.getElementById('tk-guia-input').value.trim();
    if (allTicketsList.length === 0) {
        showToast("⚠️ Primero debes procesar los datos en Coordinación");
        return;
    }
    const ticket = allTicketsList.find(t => t.num.trim() === tkInput);
    if (ticket) {
        generarFormatoGuia(ticket);
        showToast("✅ Datos de guía cargados");
    } else {
        showToast("❌ No se encontró el TK: " + tkInput);
        document.getElementById('guia-result-container').classList.add('hidden');
    }
}

function generarFormatoGuia(ticket) {
    const container = document.getElementById('guia-result-container');
    
    const para = ticket.correo || "";
    const correoJurisdiccion = CORREOS_JURISDICCION[ticket.jurisdiccion] || CORREOS_JURISDICCION["GENERAL"];
    let cc = `${CORREOS_FIJOS}; ${correoJurisdiccion}`;
    
    const proyecto = (ticket.proyecto || "").toUpperCase();
    if (proyecto.includes("PJUD 4")) cc += "; carol.oteiza@hp.com";
    else if (proyecto.includes("PJUD 5")) cc += "; christian.ojeda@hp.com";

    const nombreUsuario = ticket.usuario ? ticket.usuario.split(' ')[0] : "Usuario";
    const asuntoRetiro = `Requerimiento ${ticket.num} - Retiro de Equipo Reportado, Proyecto ${ticket.proyecto}`;

    const boxStyle = "border: 2px solid #000; width: 48%; font-family: Arial, sans-serif; font-size: 14px; background: white;";
    const headerStyle = "background: #6699cc; color: black; font-weight: bold; text-align: center; padding: 8px; border-bottom: 2px solid #000; text-transform: uppercase;";
    const contentStyle = "padding: 15px; line-height: 1.8;";
    
    const formatoHTML = `
        <div class="card" style="margin-bottom: 20px;">
            <h3 style="margin-bottom: 15px; color: #014f8b;">Vistas Previas de Guías</h3>
            <div style="display: flex; justify-content: space-between; gap: 20px; flex-wrap: wrap;">
                <div id="guia-retiro-box" style="${boxStyle}">
                    <div style="${headerStyle}">GUIA DE RETIRO DE EQUIPO</div>
                    <div style="${contentStyle}">
                        <strong>RETIRO DE EQUIPAMIENTO CON FALLA</strong><br><br>
                        <strong>${ticket.tipo} - ${ticket.modelo}</strong><br>
                        TK ${ticket.num}<br>
                        ${ticket.serie}<br><br>
                        ${ticket.dependencia}<br>
                        ${ticket.direccion}
                    </div>
                </div>
                <div id="guia-despacho-box" style="${boxStyle}">
                    <div style="${headerStyle}">GUIA DESPACHO EQUIPOS</div>
                    <div style="${contentStyle}">
                        <strong>DESPACHO DE EQUIPAMIENTO</strong><br><br>
                        <strong>COMPUTADOR ${ticket.tipo} - ${ticket.modelo}</strong><br>
                        TK ${ticket.num}<br>
                        ${ticket.serie}<br><br>
                        ${ticket.dependencia}<br>
                        ${ticket.direccion}
                    </div>
                </div>
            </div>
            <div class="button-actions-group" style="margin-top: 20px; display: flex; gap: 10px;">
                <button onclick="copyTable('guia-retiro-box')" class="btn-copy-body" style="background: #d9534f;">
                    <i class="fas fa-copy"></i> Copiar Guía Retiro
                </button>
                <button onclick="copyTable('guia-despacho-box')" class="btn-copy-body" style="background: #5cb85c;">
                    <i class="fas fa-copy"></i> Copiar Guía Despacho
                </button>
            </div>
             <p style="font-size: 0.8rem; color: #666; margin-top: 10px;">* Recuerda verificar que las columnas Modelo y Serie se estén cargando correctamente desde el CSV.</p>
        </div>

        <div class="card" style="border: none; padding: 0; background: transparent;">
            <div class="pjud-header-toggle" onclick="toggleSection('email-retiro-body', 'icon-retiro')">
                <h3>Gestión de Correo: Retiro de Equipo</h3>
                <i id="icon-retiro" class="fas fa-chevron-up"></i>
            </div>
            <div id="email-retiro-body" class="hidden-content" style="background: white; border: 1px solid #ddd; padding: 20px; border-radius: 0 0 5px 5px;">
                <div class="mail-row" style="margin-bottom: 12px;">
                    <button class="btn-copy-small" onclick="copyId('retiro-to')">¡Copiar!</button>
                    <span id="retiro-to" style="font-size: 0.85rem; color: #444;">${para}</span>
                </div>
                <div class="mail-row" style="margin-bottom: 12px;">
                    <button class="btn-copy-small" onclick="copyId('retiro-cc')">¡Copiar!</button>
                    <span id="retiro-cc" style="font-size: 0.85rem; color: #444;">${cc}</span>
                </div>
                <div class="mail-row" style="margin-bottom: 12px;">
                    <button class="btn-copy-small" onclick="copyId('retiro-sub')">¡Copiar!</button>
                    <span id="retiro-sub" style="font-size: 0.85rem;">${asuntoRetiro}</span>
                </div>

                <div id="retiro-email-content" style="background: white; padding: 15px; border: 1px solid #eee; margin-bottom: 15px; font-family: Calibri, sans-serif;">
                    <p style="font-size: 0.9rem;">Estimado/a ${nombreUsuario};</p><br>
                    <p style="font-size: 0.9rem;">Buenas tardes, Junto con saludar, informarle que de acuerdo a la atencion de su requerimiento de asunto, el equipo reportado quedó en dependencia del tribunal por motivo de recuperacion de data u horario/disponibilidad, por lo cual se requiere efectuar el retiro mediante la asistencia de un tecnico al tribunal para dicha finalidad.</p><br>
                    
                    <table style="border-collapse: collapse; width: 100%; max-width: 350px; font-size: 0.85rem; border: 1px solid #000; margin: 15px 0;">
                        <thead>
                            <tr style="background-color: #b4c6e7;"> 
                                <th colspan="2" style="border: 1px solid #000; padding: 5px; text-align: left;">Datos De Equipo.</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr><td style="border: 1px solid #000; padding: 5px; font-weight: bold; width: 40%;">Tipo</td><td style="border: 1px solid #000; padding: 5px;">${ticket.tipo}</td></tr>
                            <tr><td style="border: 1px solid #000; padding: 5px; font-weight: bold;">Serie</td><td style="border: 1px solid #000; padding: 5px;">${ticket.serie}</td></tr>
                            <tr><td style="border: 1px solid #000; padding: 5px; font-weight: bold;">TK Asociado</td><td style="border: 1px solid #000; padding: 5px;">${ticket.num}</td></tr>
                        </tbody>
                    </table>

                    <p style="font-size: 0.9rem;"><br>Favor confirmar disponibilidad para proceder al envio de la guia y la visita tecnica para el retiro.</p><br>
                    <p style="font-size: 0.85rem;">
                        <strong>Nota:</strong> en caso que el retiro ya se haya efectuado, favor confirmar para actualizar registros.<br>
                        Si aparte del equipo mencionado, cuenta con algún computador, monitor o dispositivo que este disponible para ser retirado, se agradece poder complementar la información.
                    </p>
                </div>

                <div class="button-actions-group">
                    <button onclick="enviarYCopiar('${encodeURIComponent(para)}','${encodeURIComponent(cc)}','${encodeURIComponent(asuntoRetiro)}', 'retiro-email-content')" class="btn-outlook">
                        <i class="fas fa-paper-plane"></i> 1. Copiar y Enviar Outlook
                    </button>
                    <button onclick="window.open('https://outlook.office.com/mail/deeplink/compose', '_blank')" class="btn-secondary" style="flex:1; background: #0078d4; color: white; border:none;">
                        <i class="fas fa-globe"></i> 2. Ir a Outlook Web
                    </button>
                    <button onclick="copyTable('retiro-email-content')" class="btn-copy-body">
                        <i class="fas fa-copy"></i> 3. Solo Copiar Texto
                    </button>
                </div>
            </div>
        </div>
    `;
    container.innerHTML = formatoHTML;
    container.classList.remove('hidden');
}

// --- MÓDULO REPORTES (PESTAÑAS) ---
function switchReportTab(tabId, btnElement) {
    const contents = document.querySelectorAll('#mod-reportes .tab-content');
    contents.forEach(content => {
        content.classList.add('hidden-content');
        content.classList.remove('active-content');
    });
    const buttons = document.querySelectorAll('#mod-reportes .tab-btn');
    buttons.forEach(btn => btn.classList.remove('active'));

    document.getElementById(tabId).classList.remove('hidden-content');
    document.getElementById(tabId).classList.add('active-content');
    btnElement.classList.add('active');

    // ESTO ES LO QUE FALTA: Ejecutar la lógica de gráficos cuando entras a la pestaña
   // Dentro de tu lógica de cambio de pestañas, busca el caso de 'tab-informe-mda'
if (tabId === 'tab-informe-mda') {
    setTimeout(() => { 
        renderizarHistorico(); 
        renderizarGraficoCanales(); // <-- ESTA ES LA QUE DIBUJA EL NUEVO GRÁFICO
    }, 200); 
}
}

// --- REPORTE: INFORMATIVO TÉCNICO (SÓLO SIGUIENTE DÍA HÁBIL) ---
function generarInformativoTecnico() {
    if (currentTicketsList.length === 0) {
        showToast("⚠️ No hay tickets procesados para generar el reporte.");
        return;
    }

    const container = document.getElementById('info-result-container');
    const asunto = `Informativo Técnico - Actividades Coordinadas SCO PJUD`;
    const para = "j.santos@fcom.cl";
    const cc = "c.zapata@fcom.cl; e.suarez@fcom.cl; e.socorro@fcom.cl; l.torres@fcom.cl; juan.diaz@fcom.cl; sandrade_fcom@pjud.cl; jchavez_hp@pjud.cl; s.guzman@fcom.cl; jmarrufo_hp@pjud.cl; f.solar@fcom.cl; svaldivieso_hp@pjud.cl; myabrudez_fcom@pjud.cl; j.riffo@fcom.cl; a.vacca@fcom.cl";

    const hoy = new Date();
    const diaObjetivo = new Date(hoy);
    
    if (hoy.getDay() === 5) { 
        diaObjetivo.setDate(diaObjetivo.getDate() + 3);
    } else if (hoy.getDay() === 6) { 
        diaObjetivo.setDate(diaObjetivo.getDate() + 2);
    } else {
        diaObjetivo.setDate(diaObjetivo.getDate() + 1);
    }
    
    const yyyy = diaObjetivo.getFullYear();
    const mm = String(diaObjetivo.getMonth() + 1).padStart(2, '0');
    const dd = String(diaObjetivo.getDate()).padStart(2, '0');
    const fechaObjetivoStr = `${yyyy}-${mm}-${dd}`; 
    const fechaVisual = `${dd}/${mm}/${yyyy}`; 

    let filasTabla = "";
    let ticketsCount = 0;
    
    currentTicketsList.forEach(ticket => {
        if (ticket.fechaCoord && ticket.fechaCoord === fechaObjetivoStr) {
            ticketsCount++;
            
            const esCambio = (ticket.backup && ticket.backup.toUpperCase().trim() === "SI");
            const solucion = esCambio ? "CAMBIO DE EQUIPO" : "EVALUACIÓN / CONFIGURACIÓN DE EQUIPO";
            
            const [y, m, d] = ticket.fechaCoord.split('-');
            let fechaShow = `${d}/${m}/${y}`;
            if (ticket.horaCoord) fechaShow += ` ${ticket.horaCoord}`;

            const tecnicoShow = ticket.tecnicoCoord || "POR ASIGNAR";

            filasTabla += `
                <tr style="font-size: 11px; border-bottom: 1px solid #ddd;">
                    <td style="padding: 4px; border: 1px solid #ccc;">${ticket.num}</td>
                    <td style="padding: 4px; border: 1px solid #ccc;">${ticket.proyecto}</td>
                    <td style="padding: 4px; border: 1px solid #ccc;">${ticket.modelo}</td>
                    <td style="padding: 4px; border: 1px solid #ccc;">${ticket.serie}</td>
                    <td style="padding: 4px; border: 1px solid #ccc;">${ticket.dependencia}</td>
                    <td style="padding: 4px; border: 1px solid #ccc;">${ticket.tipo}</td>
                    <td style="padding: 4px; border: 1px solid #ccc; background-color: #d1e7dd;">${fechaShow}</td>
                    <td style="padding: 4px; border: 1px solid #ccc; background-color: #d1e7dd;">${tecnicoShow}</td>
                    <td style="padding: 4px; border: 1px solid #ccc;">${solucion}</td>
                </tr>
            `;
        }
    });

    if (ticketsCount === 0) {
        showToast(`ℹ️ No hay actividades coordinadas para el próximo día hábil (${fechaVisual}).`);
        container.classList.add('hidden');
        return;
    }

    const htmlReporte = `
        <div class="pjud-header-toggle" onclick="toggleSection('info-body', 'icon-info')">
            <h3>Vista Previa: Informativo Técnico (${ticketsCount} registros para el ${fechaVisual})</h3>
            <i id="icon-info" class="fas fa-chevron-up rotate-icon"></i>
        </div>
        <div id="info-body" style="background: white; border: 1px solid #ddd; padding: 20px; border-radius: 0 0 5px 5px;">
            <div class="mail-row" style="margin-bottom: 12px;">
                <button class="btn-copy-small" onclick="copyId('info-to')">¡Copiar!</button>
                <span style="font-weight:bold; margin-right:5px; color:#014f8b;">Para:</span>
                <span id="info-to" style="font-size: 0.85rem; color: #444;">${para}</span>
            </div>
            <div class="mail-row" style="margin-bottom: 12px;">
                <button class="btn-copy-small" onclick="copyId('info-cc')">¡Copiar!</button>
                <span style="font-weight:bold; margin-right:5px; color:#014f8b;">CC:</span>
                <span id="info-cc" style="font-size: 0.85rem; color: #444;">${cc}</span>
            </div>
            <div class="mail-row" style="margin-bottom: 12px;">
                <button class="btn-copy-small" onclick="copyId('info-sub')">¡Copiar!</button>
                <span style="font-weight:bold; margin-right:5px; color:#014f8b;">Asunto:</span>
                <span id="info-sub" style="font-size: 0.85rem;">${asunto}</span>
            </div>
            <div id="info-email-content" style="background: white; padding: 15px; border: 1px solid #eee; margin-bottom: 15px; font-family: Calibri, sans-serif; color: #000;">
                <p style="margin-bottom: 10px;">Jorge<br>Buenas tardes</p><br>
                <p style="margin-bottom: 10px;">
                Junto con saludar, informo de las actividades que se realizarán el próximo día hábil por el área de SCO PJUD.
                Se brinda información del técnico, la dependencia y la actividad a realizar. Favor revisar para poder apoyar a los técnicos que estarán en dichos procesos, de esta forma puedan contar con las herramientas, imágenes y procedimientos actualizados.
                <br><br>Actividades coordinadas para el ${fechaVisual}:<br>
                </p><br>
                <table style="border-collapse: collapse; width: 100%; border: 1px solid #999; font-family: Arial, sans-serif;">
                    <thead>
                        <tr style="background-color: #dbe5f1; font-size: 11px; text-align: left;">
                            <th style="padding: 5px; border: 1px solid #999;">TICKET</th>
                            <th style="padding: 5px; border: 1px solid #999;">PROYECTO</th>
                            <th style="padding: 5px; border: 1px solid #999;">MODELO EQUIPO</th>
                            <th style="padding: 5px; border: 1px solid #999;">SERIE REPORTADA</th>
                            <th style="padding: 5px; border: 1px solid #999;">DEPENDENCIA</th>
                            <th style="padding: 5px; border: 1px solid #999;">TIPO</th>
                            <th style="padding: 5px; border: 1px solid #999;">FECHA COORDINACIÓN</th>
                            <th style="padding: 5px; border: 1px solid #999;">TECNICO COORDINADO</th>
                            <th style="padding: 5px; border: 1px solid #999;">SOLUCION TERRENO</th>
                        </tr>
                    </thead>
                    <tbody>${filasTabla}</tbody>
                </table>
                <br>
            </div>
            <div class="button-actions-group">
                <button onclick="enviarYCopiar('${encodeURIComponent(para)}', '${encodeURIComponent(cc)}', '${encodeURIComponent(asunto)}', 'info-email-content')" class="btn-outlook">
                    <i class="fas fa-paper-plane"></i> 1. Copiar y Enviar Outlook
                </button>
                <button onclick="window.open('https://outlook.office.com/mail/deeplink/compose', '_blank')" class="btn-secondary" style="flex:1; background: #0078d4; color: white; border:none;">
                    <i class="fas fa-globe"></i> 2. Ir a Outlook Web
                </button>
                <button onclick="copyTable('info-email-content')" class="btn-copy-body">
                    <i class="fas fa-copy"></i> 3. Solo Copiar Texto
                </button>
            </div>
        </div>
    `;
    container.innerHTML = htmlReporte;
    container.classList.remove('hidden');
}

function exportarInformativoExcel() {
    if (currentTicketsList.length === 0) {
        showToast("⚠️ No hay datos para exportar.");
        return;
    }

    const hoy = new Date();
    const diaObjetivo = new Date(hoy);
    if (hoy.getDay() === 5) diaObjetivo.setDate(diaObjetivo.getDate() + 3);
    else if (hoy.getDay() === 6) diaObjetivo.setDate(diaObjetivo.getDate() + 2);
    else diaObjetivo.setDate(diaObjetivo.getDate() + 1);

    const yyyy = diaObjetivo.getFullYear();
    const mm = String(diaObjetivo.getMonth() + 1).padStart(2, '0');
    const dd = String(diaObjetivo.getDate()).padStart(2, '0');
    const fechaObjetivoStr = `${yyyy}-${mm}-${dd}`; 

    let csvContent = "TICKET;PROYECTO;MODELO EQUIPO;SERIE REPORTADA;DEPENDENCIA;TIPO;FECHA COORDINACIÓN;TECNICO COORDINADO;SOLUCION TERRENO\n";
    let datosExportados = 0;

    currentTicketsList.forEach(ticket => {
        if (ticket.fechaCoord && ticket.fechaCoord === fechaObjetivoStr) {
            datosExportados++;
            const esCambio = (ticket.backup && ticket.backup.toUpperCase().trim() === "SI");
            const solucion = esCambio ? "CAMBIO DE EQUIPO" : "EVALUACIÓN / CONFIGURACIÓN DE EQUIPO";
            
            const [y, m, d] = ticket.fechaCoord.split('-');
            let fechaShow = `${d}/${m}/${y}`;
            if (ticket.horaCoord) fechaShow += ` ${ticket.horaCoord}`;
            
            const tecnicoShow = ticket.tecnicoCoord || "POR ASIGNAR";

            const row = [ticket.num, ticket.proyecto, ticket.modelo, ticket.serie, ticket.dependencia, ticket.tipo, fechaShow, tecnicoShow, solucion];
            csvContent += row.join(";") + "\n";
        }
    });

    if (datosExportados === 0) {
        showToast("⚠️ No hay tickets coordinados para el próximo día hábil.");
        return;
    }

    const blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `Informativo_Tecnico_${fechaObjetivoStr}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast("✅ Archivo Excel descargado");
}

// --- REPORTE: EN BLANCO ---
function generarReporteEnBlanco() {
    const container = document.getElementById('blanco-result-container');
    const fInicio = document.getElementById('blanco-fecha-inicio').value;
    const fFin = document.getElementById('blanco-fecha-fin').value;
    
    const selectProyecto = document.getElementById('blanco-filtro-proyecto');
    const filtroProyecto = selectProyecto ? selectProyecto.value : "TODOS";

    if (allTicketsList.length === 0) return showToast("⚠️ Carga datos primero.");

    let dDesde = fInicio ? new Date(fInicio + "T00:00:00") : null;
    let dHasta = fFin ? new Date(fFin + "T23:59:59") : null;

    const casosEnBlanco = allTicketsList.filter(t => {
        if (!t.num || t.num.trim() === "") return false;

        const estadoUpper = t.estado ? t.estado.toUpperCase() : "";
        if (!estadoUpper.includes("FINALIZADO") && !estadoUpper.includes("CERRADO")) return false;

        const proyUpper = t.proyecto ? t.proyecto.toUpperCase() : "";
        if (filtroProyecto !== "TODOS" && !proyUpper.includes(filtroProyecto)) return false;

        const modUpper = t.modelo ? t.modelo.toUpperCase().trim() : "";
        const sinModelo = modUpper === "" || modUpper === "MODELO N/A" || modUpper === "N/A" || modUpper === "-" || modUpper === "S/N" || modUpper === "." || modUpper === "0";

        const tipoUpper = t.tipo ? t.tipo.toUpperCase().trim() : "";
        const sinTipo = tipoUpper === "" || tipoUpper === "EQUIPO" || tipoUpper === "N/A" || tipoUpper === "-" || tipoUpper === "S/N";

        const depUpper = t.dependencia ? t.dependencia.toUpperCase().trim() : "";
        const sinDependencia = depUpper === "" || depUpper === "N/A" || depUpper === "-";

        const jurUpper = t.jurisdiccion ? t.jurisdiccion.toUpperCase().trim() : "";
        const sinJurisdiccion = jurUpper === "" || jurUpper === "N/A" || jurUpper === "-" || jurUpper === "GENERAL";

        if (!(sinModelo || sinTipo || sinDependencia || sinJurisdiccion)) return false;

        if (dDesde && dHasta) {
            if (!t.fechaFin) return false;
            const parts = t.fechaFin.split(' ')[0].split(/[-/]/);
            if(parts.length === 3) {
                 const fechaT = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T12:00:00`);
                 if (fechaT < dDesde || fechaT > dHasta) return false;
            } else {
                return false;
            }
        }
        return true;
    });

    if (casosEnBlanco.length === 0) {
        container.innerHTML = `<div style="padding:15px; background:#d4edda; color:#155724; border-radius:5px;">✅ No se encontraron requerimientos incompletos para los filtros seleccionados.</div>`;
        container.classList.remove('hidden');
        return;
    }

    let tablaHTML = `
        <div style="padding:15px; background:#fff3cd; color:#856404; border-radius:5px; margin-bottom:15px;">⚠️ Se encontraron <b>${casosEnBlanco.length}</b> requerimientos sin información técnica completa.</div>
        <div style="overflow-x:auto;">
            <table id="tabla-en-blanco" style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px;">
                <thead style="background-color: #6c757d; color: white;">
                    <tr>
                        <th style="padding:8px; border:1px solid #ddd; text-align: left;">Proyecto</th>
                        <th style="padding:8px; border:1px solid #ddd; text-align: left; white-space: nowrap;">
                            Ticket 
                            <button onclick="copiarColumnaTicketsBlanco()" title="Copiar solo Tickets" style="margin-left: 5px; padding: 2px 6px; font-size: 10px; background-color: #495057; color: white; border: 1px solid #ccc; border-radius: 3px; cursor: pointer;">
                                <i class="fas fa-copy"></i>
                            </button>
                        </th>
                        <th style="padding:8px; border:1px solid #ddd; text-align: left;">Grupo Resolutor</th>
                        
                        <th style="padding:8px; border:1px solid #ddd; text-align: left; white-space: nowrap;">
                            Asignado A
                            <button onclick="copiarColumnaAsignadoBlanco()" title="Copiar solo Nombres" style="margin-left: 5px; padding: 2px 6px; font-size: 10px; background-color: #495057; color: white; border: 1px solid #ccc; border-radius: 3px; cursor: pointer;">
                                <i class="fas fa-copy"></i>
                            </button>
                        </th>
                        
                        <th style="padding:8px; border:1px solid #ddd; text-align: center;">Fecha Finalizado</th>
                        <th style="padding:8px; border:1px solid #ddd; text-align: left;">Modelo Detectado</th>
                        <th style="padding:8px; border:1px solid #ddd; text-align: left;">Tipo Detectado</th>
                        <th style="padding:8px; border:1px solid #ddd; text-align: left;">Jurisdicción</th>
                        <th style="padding:8px; border:1px solid #ddd; text-align: left;">Dependencia</th>
                    </tr>
                </thead>
                <tbody>
    `;

    casosEnBlanco.forEach(t => {
        const modUpper = t.modelo ? t.modelo.toUpperCase().trim() : "";
        const sinModelo = modUpper === "" || modUpper === "MODELO N/A" || modUpper === "N/A" || modUpper === "-" || modUpper === "S/N" || modUpper === "." || modUpper === "0";

        const tipoUpper = t.tipo ? t.tipo.toUpperCase().trim() : "";
        const sinTipo = tipoUpper === "" || tipoUpper === "EQUIPO" || tipoUpper === "N/A" || tipoUpper === "-" || tipoUpper === "S/N";

        const depUpper = t.dependencia ? t.dependencia.toUpperCase().trim() : "";
        const sinDependencia = depUpper === "" || depUpper === "N/A" || depUpper === "-";

        const jurUpper = t.jurisdiccion ? t.jurisdiccion.toUpperCase().trim() : "";
        const sinJurisdiccion = jurUpper === "" || jurUpper === "N/A" || jurUpper === "-" || jurUpper === "GENERAL";

        const modeloShow = sinModelo ? `FALTA MODELO ${modUpper && modUpper !== "MODELO N/A" ? `("${t.modelo}")` : ''}` : t.modelo;
        const tipoShow = sinTipo ? `FALTA TIPO ${tipoUpper && tipoUpper !== "EQUIPO" ? `("${t.tipo}")` : ''}` : t.tipo;
        const jurShow = sinJurisdiccion ? `FALTA JURISDICCIÓN ${jurUpper === "GENERAL" ? '' : `("${t.jurisdiccion}")`}` : t.jurisdiccion;
        const depShow = sinDependencia ? "FALTA DEPENDENCIA" : t.dependencia;

        const asignadoShow = t.AsignadoA || "-";

        tablaHTML += `
            <tr style="border-bottom: 1px solid #eee;">
                <td style="padding:8px; border:1px solid #ddd; color: #014f8b; font-weight: bold;">${t.proyecto || "N/A"}</td>
                <td style="padding:8px; border:1px solid #ddd;"><b>${t.num}</b></td>
                <td style="padding:8px; border:1px solid #ddd;">${t.grupo}</td>
                <td style="padding:8px; border:1px solid #ddd;">${asignadoShow}</td>
                <td style="padding:8px; border:1px solid #ddd; text-align: center;">${t.fechaFin}</td>
                <td style="padding:8px; border:1px solid #ddd; color: ${sinModelo ? '#dc3545' : '#333'}; font-weight: ${sinModelo ? 'bold' : 'normal'};">${modeloShow}</td>
                <td style="padding:8px; border:1px solid #ddd; color: ${sinTipo ? '#dc3545' : '#333'}; font-weight: ${sinTipo ? 'bold' : 'normal'};">${tipoShow}</td>
                <td style="padding:8px; border:1px solid #ddd; color: ${sinJurisdiccion ? '#dc3545' : '#333'}; font-weight: ${sinJurisdiccion ? 'bold' : 'normal'};">${jurShow}</td>
                <td style="padding:8px; border:1px solid #ddd; color: ${sinDependencia ? '#dc3545' : '#333'}; font-weight: ${sinDependencia ? 'bold' : 'normal'};">${depShow}</td>
            </tr>
        `;
    });

    tablaHTML += `</tbody></table></div>`;
    container.innerHTML = tablaHTML;
    container.classList.remove('hidden');
}

function copiarColumnaTicketsBlanco() {
    const tabla = document.getElementById('tabla-en-blanco');
    if (!tabla) return showToast("⚠️ No hay tabla generada.");

    let tickets = [];
    const filas = tabla.querySelectorAll('tbody tr');
    
    filas.forEach(fila => {
        const celdaTicket = fila.children[1]; 
        if (celdaTicket) tickets.push(celdaTicket.innerText.trim());
    });

    if (tickets.length === 0) return showToast("⚠️ No hay tickets para copiar.");
    
    const textoCopiar = tickets.join('\n');
    copiarAlPortapapeles(textoCopiar, `✅ ${tickets.length} tickets copiados`);
}

function copiarColumnaAsignadoBlanco() {
    const tabla = document.getElementById('tabla-en-blanco');
    if (!tabla) return showToast("⚠️ No hay tabla generada.");

    let asignados = [];
    const filas = tabla.querySelectorAll('tbody tr');
    
    filas.forEach(fila => {
        const celdaAsignado = fila.children[3]; 
        if (celdaAsignado) asignados.push(celdaAsignado.innerText.trim());
    });

    if (asignados.length === 0) return showToast("⚠️ No hay datos para copiar.");
    
    const textoCopiar = asignados.join('\n');
    copiarAlPortapapeles(textoCopiar, `✅ ${asignados.length} nombres copiados`);
}

function copiarAlPortapapeles(texto, mensajeExito) {
    if (navigator.clipboard) {
        navigator.clipboard.writeText(texto).then(() => {
            showToast(mensajeExito);
        }).catch(err => {
            showToast("⚠️ Error al copiar");
        });
    } else {
        const textArea = document.createElement("textarea");
        textArea.value = texto;
        document.body.appendChild(textArea);
        textArea.select();
        try {
            document.execCommand('copy');
            showToast(mensajeExito);
        } catch (err) {
            showToast("⚠️ Error al copiar");
        }
        document.body.removeChild(textArea);
    }
}

function exportarReporteBlancoExcel() {
    const table = document.getElementById("tabla-en-blanco");
    if (!table) {
        showToast("⚠️ Genera el reporte antes de exportar.");
        return;
    }
    let csv = [];
    const rows = table.querySelectorAll("tr");
    for (let i = 0; i < rows.length; i++) {
        const row = [], cols = rows[i].querySelectorAll("td, th");
        for (let j = 0; j < cols.length; j++) row.push(cols[j].innerText);
        csv.push(row.join(";"));
    }
    const csvContent = "data:text/csv;charset=utf-8,\uFEFF" + csv.join("\n");
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", `Reporte_En_Blanco_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}


// --- REPORTE: CONTROL DE CAMBIOS ---
async function generarReporteCambios() {
    const container = document.getElementById('cambios-result-container');
    const inputInicio = document.getElementById('filtro-fecha-inicio').value; 
    const inputFin = document.getElementById('filtro-fecha-fin').value;        
    const inputProyecto = document.getElementById('filtro-proyecto').value;

    function parseDateFromCSV(dateStr) {
        if (!dateStr) return null;
        const cleanDateStr = dateStr.trim().split(' ')[0]; 
        const parts = cleanDateStr.split(/[-/]/);
        if (parts.length !== 3) return null;
        return new Date(parts[2], parts[1] - 1, parts[0]);
    }

    let fechaDesde = inputInicio ? new Date(inputInicio.split('-')[0], inputInicio.split('-')[1] - 1, inputInicio.split('-')[2]) : null;
    let fechaHasta = inputFin ? new Date(inputFin.split('-')[0], inputFin.split('-')[1] - 1, inputFin.split('-')[2], 23, 59, 59) : null;
    
    const ticketsFiltrados = allTicketsList.filter(t => {
        const estadoUpper = t.estado ? t.estado.toUpperCase() : "";
        if (!estadoUpper.includes("FINALIZADO")) return false;
        const backupNorm = t.backup ? t.backup.toUpperCase().trim() : "";
        if (backupNorm !== "SI") return false;
        if (inputProyecto !== "TODOS" && (!t.proyecto || !t.proyecto.toUpperCase().includes(inputProyecto))) return false;
        if (fechaDesde && fechaHasta) {
            const fechaTicket = parseDateFromCSV(t.fechaFin); 
            if (!fechaTicket || fechaTicket < fechaDesde || fechaTicket > fechaHasta) return false;
        }
        return true; 
    });

    if (ticketsFiltrados.length === 0) {
        showToast("⚠️ No se encontraron tickets para el reporte.");
        container.innerHTML = "<p style='text-align:center; color:#666;'>Sin resultados.</p>";
        container.classList.remove('hidden');
        return;
    }

    const fechaFinFormato = inputFin ? inputFin.split('-').reverse().join('-') : "";
    let para = "", cc = "", asunto = "";
    const cuerpo = "Buen día Estimados,<br><br> Adjunto reporte acumulativo de control de cambios del Proyecto.";

    if (inputProyecto.includes("PJUD 4")) {
        para = "Jorge.ceballos.de.la.carrera@hp.com;carol.oteiza@hp.com";
        cc = "j.marrufo@fcom.cl; juan.diaz@fcom.cl; jchavez_hp@pjud.cl;j.riffo@fcom.cl;s.valbuena@fcom.cl; s.guzman@fcom.cl; jmarrufo_hp@pjud.cl; svaldivieso_hp@pjud.cl; avacca_fcom@pjud.cl; a.vacca@fcom.cl; c.zapata@fcom.cl";
        asunto = `Control de cambios SCO PJUD4-2 hasta el ${fechaFinFormato}`;
    } else {
        para = "jorge.ceballos.de.la.carrera@hp.com; christian.ojeda@hp.com";
        cc = "j.marrufo@fcom.cl; juan.diaz@fcom.cl; jchavez_hp@pjud.cl; s.guzman@fcom.cl; jmarrufo_hp@pjud.cl; svaldivieso_hp@pjud.cl; avacca_fcom@pjud.cl; a.vacca@fcom.cl; christian.ojeda@hp.com; c.zapata@fcom.cl; j.riffo@fcom.cl";
        asunto = `Control de cambios SCO PJUD5 hasta el ${fechaFinFormato}`;
    }

    let filas = "";
    ticketsFiltrados.forEach(t => {
        const dependenciaLimpia = t.dependencia ? t.dependencia.trim() : "";
        const modeloLimpio = t.modelo ? t.modelo.trim() : "";
        
        const skuEncontrado = (typeof CATALOGO_SKUS !== 'undefined') ? (CATALOGO_SKUS[modeloLimpio] || "#N/A") : "#N/A";
        const ciudadEncontrada = (typeof CATALOGO_CIUDADES !== 'undefined') ? (CATALOGO_CIUDADES[dependenciaLimpia] || "#N/A") : "#N/A";

        filas += `
            <tr>
                <td>${t.proyecto || "#N/A"}</td>
                <td>${t.grupo || "RESIDENTES"}</td>
                <td style="background-color: #e8f5e9; font-weight: 500;">${t.num}</td>
                <td>${t.tipo || "EQUIPO"}</td>
                <td>${modeloLimpio}</td>
                <td>${skuEncontrado}</td>
               <td style="color: #e74c3c; font-weight: bold;">${t.serie || "#N/A"}</td>
                <td>${modeloLimpio}</td>
                <td>${skuEncontrado}</td>
                <td>${t.fechaFin || "#N/A"}</td>
                <td>${t.fechaFin || "#N/A"}</td>
                <td style="background-color: #e8f5e9; font-weight: 500;">${t.despachosRaw || ""}</td>
                <td style="background-color: #e8f5e9; font-weight: 500;">${t.despachosRaw || ""}</td>
                <td>${t.dependencia}</td>
                <td>${ciudadEncontrada}</td>
                <td>${t.jurisdiccion || "#N/A"}</td>
                <td>${t.usuario || "Sin Nombre"}</td>
                <td>-</td>
                <td>CHILE</td>
                <td>${t.correo || ""}</td>
                <td style="font-weight: bold; color: #007db3;">${t.ip || "-"}</td>
            </tr>
        `;
    });

    container.innerHTML = `
        <div class="card" style="border: 1px solid #ddd; margin-bottom: 20px;">
            <div class="pjud-header-toggle" style="background: #e3f2fd; border-top: 4px solid #007db3;">
                <h3 style="color: #014f8b;"><i class="fas fa-envelope"></i> Gestión de Correo: ${inputProyecto}</h3>
            </div>
            <div style="padding: 15px; font-size: 0.85rem;">
                <div class="mail-row" style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
                    <button class="btn-copy-small" style="min-width: 60px;" onclick="copiarTextoDirecto('${para}')">Copiar</button>
                    <strong>Para:</strong> <span id="cc-to">${para}</span>
                </div>
                <div class="mail-row" style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
                    <button class="btn-copy-small" style="min-width: 60px;" onclick="copiarTextoDirecto('${cc}')">Copiar</button>
                    <strong>CC:</strong> <span id="cc-copy">${cc}</span>
                </div>
                <div class="mail-row" style="display: flex; align-items: center; gap: 10px; margin-bottom: 12px;">
                    <button class="btn-copy-small" style="min-width: 60px;" onclick="copiarTextoDirecto('${asunto}')">Copiar</button>
                    <strong>Asunto:</strong> <span id="cc-sub">${asunto}</span>
                </div>
                <div id="cuerpo-correo-reporte" style="margin-top:10px; padding:12px; border:1px dashed #ccc; background:#fff; font-family: Calibri, sans-serif;">
                    ${cuerpo}
                </div>
                <div style="margin-top: 15px;">
                    <button class="btn-outlook" style="padding: 8px 16px; font-size: 0.85rem; height: auto !important; width: auto !important; color: white !important;" 
    onclick="enviarReporteOutlook('${encodeURIComponent(para)}', '${encodeURIComponent(cc)}', '${encodeURIComponent(asunto)}', 'cuerpo-correo-reporte')">
    <i class="fas fa-paper-plane"></i> Enviar y Copiar Cuerpo
</button>
                </div>
            </div>
        </div>
        <div class="pjud-header-toggle" onclick="toggleSection('cambios-table-body', 'icon-cambios')">
            <h3><i class="fas fa-table"></i> Reporte: ${ticketsFiltrados.length} registros encontrados</h3>
            <i id="icon-cambios" class="fas fa-chevron-up rotate-icon"></i>
        </div>
        <div id="cambios-table-body" style="overflow-x: auto; border: 1px solid #eee; border-radius: 0 0 5px 5px;">
            <table id="tabla-cambios-excel" style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px; white-space: nowrap;">
                <thead style="background-color: #0088cc; color: white;">
                    <tr>
                        <th>Proyecto</th><th>GRUPO</th><th>TICKET</th><th>TIPO</th><th>MODELO</th><th>SKU</th><th>SALIENTE</th>
                        <th>NEW DESCRIPTION</th><th>NEW SKU</th><th>FECHA FIN</th><th>FECHA FIN</th>
                        <th>ENTRANTE</th><th>DESPACHOS</th>
                        <th>DEPENDENCIA</th><th>CIUDAD</th><th>REGION</th><th>USUARIO</th><th>DNI</th><th>PAIS</th><th>CORREO</th>
                        <th>DIRECCIÓN IP</th>
                    </tr>
                </thead>
                <tbody>${filas}</tbody>
            </table>
        </div>`;
    
    container.classList.remove('hidden');
}

function enviarReporteOutlook(para, cc, asunto, idCuerpo) {
    const textoCuerpo = document.getElementById(idCuerpo).innerText;
    navigator.clipboard.writeText(textoCuerpo);
    const mailtoUrl = `mailto:${para}?cc=${cc}&subject=${asunto}`;
    window.location.href = mailtoUrl;
    showToast("✅ Destinatarios cargados y cuerpo copiado.");
}

function copiarTablaCambios() {
    const tabla = document.getElementById('tabla-cambios-excel');
    if (!tabla) {
        showToast("⚠️ Primero genera la tabla.");
        return;
    }
    const range = document.createRange();
    range.selectNode(tabla);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(range);
    document.execCommand("copy");
    window.getSelection().removeAllRanges();
    showToast("✅ Tabla copiada. Pégala en Excel (Ctrl+V)");
}

function toggleSidebarLinks() {
    const container = document.getElementById('sii-container');
    const icon = document.getElementById('sii-icon');
    if(container) container.classList.toggle('hidden');
    if(icon) icon.classList.toggle('rotate-icon');
}

function copiarTextoDirecto(texto) {
    if (navigator.clipboard) {
        navigator.clipboard.writeText(texto).then(() => {
            showToast("¡Copiado!");
        }).catch(err => {
            showToast("Error al copiar");
        });
    } else {
        const textArea = document.createElement("textarea");
        textArea.value = texto;
        document.body.appendChild(textArea);
        textArea.select();
        try {
            document.execCommand('copy');
            showToast("¡Copiado!");
        } catch (err) {
            showToast("Error al copiar");
        }
        document.body.removeChild(textArea);
    }
}

function enviarYCopiar(para, cc, asunto, idElemento) {
    copyTable(idElemento); 
    setTimeout(() => { intentarAbrirOutlook(para, cc, asunto); }, 300);
}

// --- REPORTE: ANTIVIRUS ---
function generarReporteAntivirus() {
    const container = document.getElementById('antivirus-result-container');
    const inputFecha = document.getElementById('av-fecha-unica').value; 

    if (!inputFecha) return showToast("⚠️ Por favor selecciona una fecha.");

    const para = "myabrudez_fcom@pjud.cl"; 
    const cc = "c.zapata@fcom.cl; s.guzman@fcom.cl; jmarrufo_hp@pjud.cl; roberto.miranda@fcom.cl; a.vacca@fcom.cl";
    
    const [y, m, d] = inputFecha.split('-');
    const fechaFormat = `${d}/${m}/${y}`;
    const asunto = `Actividades de Cambios y Masterizaciones PC - ${fechaFormat}`;
    const fechaSeleccionada = new Date(y, m - 1, d, 0, 0, 0);

    function parseDateSimple(dateStr) {
        if (!dateStr) return null;
        const parts = dateStr.trim().split(' ')[0].split(/[-/]/);
        if (parts.length !== 3) return null;
        return new Date(parts[2], parts[1] - 1, parts[0], 0, 0, 0);
    }
    
    const ticketsFiltrados = allTicketsList.filter(t => {
        const estadoUpper = t.estado ? t.estado.toUpperCase() : "";
        if (!estadoUpper.includes("FINALIZADO")) return false;
        
        const grupoUpper = t.grupo ? t.grupo.toUpperCase() : "";
        if (!grupoUpper.includes("SCO") && !grupoUpper.includes("RESIDENTES")) return false;
        
        const tipoUpper = t.tipo ? t.tipo.toUpperCase() : "";
        if (!tipoUpper.includes("COMPUTADOR") && !tipoUpper.includes("NOTEBOOK")) return false;

        const fechaTicket = parseDateSimple(t.fechaFin);
        if (!fechaTicket) return false;
        return fechaTicket.getTime() === fechaSeleccionada.getTime();
    });

    if (ticketsFiltrados.length === 0) {
        showToast(`⚠️ No hay tickets de PC/Notebook finalizados para el ${fechaFormat}.`);
        container.innerHTML = "<p style='text-align:center; color:#666; padding: 20px;'>Sin resultados para la fecha seleccionada.</p>";
        container.classList.remove('hidden');
        return;
    }

    let filasHTML = "";
    ticketsFiltrados.forEach(t => {
        let serieDespachada = "";
        const esCambio = (t.solucion === "CAMBIO EQUIPO" && t.backup === "SI");
        if (esCambio) {
            serieDespachada = t.despachosRaw ? t.despachosRaw.trim() : "Pendiente Validar";
        }

        // Estilo de color para el Backup
        const backupStyle = t.backup === "SI" ? "color: #dc3545; font-weight: bold;" : "color: #28a745;";

        filasHTML += `
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 5px; border: 1px solid #ccc;">${t.proyecto || ""}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.num}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.grupo || ""}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.tipo || ""}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.solucion || ""}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.serie || ""}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${serieDespachada}</td>
                <td style="padding: 5px; border: 1px solid #ccc; text-align:center; ${backupStyle}">${t.backup || "NO"}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.ip || ""}</td>
            </tr>
        `;
    });

    // Encabezado de tabla (agregamos columna BACKUP)
    const headersHTML = `
        <th style="padding: 5px; border: 1px solid #ddd;">PROYECTO</th>
        <th style="padding: 5px; border: 1px solid #ddd;">TK</th>
        <th style="padding: 5px; border: 1px solid #ddd;">GRUPO RESOLUTOR</th>
        <th style="padding: 5px; border: 1px solid #ddd;">TIPO</th>
        <th style="padding: 5px; border: 1px solid #ddd;">SOLUCION TERRENO</th>
        <th style="padding: 5px; border: 1px solid #ddd;">SERIE REPORTADA</th>
        <th style="padding: 5px; border: 1px solid #ddd;">SERIE DESPACHADA</th>
        <th style="padding: 5px; border: 1px solid #ddd;">BACKUP</th>
        <th style="padding: 5px; border: 1px solid #ddd;">IP</th>
    `;

    const tablaVisual = `
        <div class="card" style="border: none; padding: 0; background: transparent; margin-bottom: 20px;">
            <div class="pjud-header-toggle" onclick="toggleSection('av-table-body', 'icon-av-table')" style="border-top: 4px solid #0088cc; cursor: pointer;">
                <h3 style="color: #0088cc; margin: 0;">Reporte Antivirus: ${ticketsFiltrados.length} registros (${fechaFormat})</h3>
                <i id="icon-av-table" class="fas fa-chevron-up"></i>
            </div>
            <div id="av-table-body" style="transition: all 0.3s ease;">
                <div style="max-height: 300px; overflow-y: auto; border: 1px solid #ccc; border-top: none;">
                    <table id="tabla-antivirus" style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px;">
                        <thead style="background-color: #4a6fa5; color: white; position: sticky; top: 0; z-index: 10;">
                            <tr>${headersHTML}</tr>
                        </thead>
                        <tbody>${filasHTML}</tbody>
                    </table>
                </div>
            </div>
        </div>
    `;

    const correoHTML = `
        <div class="card" style="border: none; padding: 0; background: transparent;">
            <div class="pjud-header-toggle" onclick="toggleSection('av-email-body', 'icon-av-email')" style="border-top: 4px solid #6c757d;">
                <h3 style="color: #444;">Gestión de Correo (Formato Antivirus)</h3>
                <i id="icon-av-email" class="fas fa-chevron-up"></i>
            </div>
            <div id="av-email-body" style="background: white; border: 1px solid #ddd; padding: 20px; border-radius: 0 0 5px 5px;">
                <div class="mail-row" style="margin-bottom: 8px;">
                    <button class="btn-copy-small" onclick="copyId('av-to')">Copiar</button>
                    <span style="font-weight:bold; width: 60px; display:inline-block;">Para:</span>
                    <span id="av-to" style="color: #444;">${para}</span>
                </div>
                <div class="mail-row" style="margin-bottom: 8px;">
                    <button class="btn-copy-small" onclick="copyId('av-cc')">Copiar</button>
                    <span style="font-weight:bold; width: 60px; display:inline-block;">CC:</span>
                    <span id="av-cc" style="color: #444;">${cc}</span>
                </div>
                <div class="mail-row" style="margin-bottom: 15px;">
                    <button class="btn-copy-small" onclick="copyId('av-sub')">Copiar</button>
                    <span style="font-weight:bold; width: 60px; display:inline-block;">Asunto:</span>
                    <span id="av-sub">${asunto}</span>
                </div>
                <div id="av-email-content" style="background: white; padding: 15px; border: 1px solid #ccc; font-family: Calibri, sans-serif; font-size: 11pt; color: #000;">
                    <p>Miguel<br>Buenos días</p><br>
                    <p>Envío listado de los requerimientos gestionados el día de ayer, que involucraron Cambio o masterización de equipo por las areas de SCO o residencias.</p>
                    <br><br>
                    <table style="border-collapse: collapse; width: 100%; border: 1px solid #999; font-family: Calibri, sans-serif; font-size: 10pt;">
                        <thead style="background-color: #e6e6e6;">
                            <tr>${headersHTML.replace(/#ddd/g, '#999').replace(/padding: 5px/g, 'padding: 4px')}</tr>
                        </thead>
                        <tbody>${filasHTML}</tbody>
                    </table>
                    <br>
                </div>
                <div class="button-actions-group" style="margin-top: 15px;">
                    <button onclick="enviarYCopiar('${encodeURIComponent(para)}', '${encodeURIComponent(cc)}', '${encodeURIComponent(asunto)}', 'av-email-content')" class="btn-outlook">
                        <i class="fas fa-paper-plane"></i> Copiar y Enviar Outlook
                    </button>
                    <button onclick="copyTable('av-email-content')" class="btn-copy-body">
                        <i class="fas fa-copy"></i> Solo Copiar Texto
                    </button>
                </div>
            </div>
        </div>
    `;

    container.innerHTML = tablaVisual + correoHTML;
    container.classList.remove('hidden');
}

function copiarTablaAntivirus() {
    const tabla = document.getElementById('tabla-antivirus');
    if (!tabla) {
        showToast("⚠️ Primero genera la tabla.");
        return;
    }
    const range = document.createRange();
    range.selectNode(tabla);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(range);
    document.execCommand("copy");
    window.getSelection().removeAllRanges();
    showToast("✅ Tabla copiada al portapapeles");
}

// --- CONFIGURACIÓN DE LA BIBLIOTECA DE ENLACES ---
const BIBLIOTECA_LINKS = [
    {
        titulo: "Generar Guías de Despacho (SII)",
        desc: "Plataforma para emisión de documentos tributarios.",
        url: "https://zeus.sii.cl/AUT2000/InicioAutenticacion/IngresoRutClave.html",
        credenciales: [
            { label: "RUT", value: "102478460", tipo: "text" },
            { label: "Clave", value: "Leonera7", tipo: "password" },
            { label: "Firma", value: "Leon3ra7#", tipo: "password" }
        ]
    },
     {
        titulo: "Control De Cambios PJUD5",
        desc: "Registro acumulado de cambios de equipos del proyecto PJUD5",
        url: "https://docs.google.com/spreadsheets/d/1qqDsCVN00kAnWVNUZfImf3CuqowyN1z2vTN-DXWeINM/edit?gid=0#gid=0",
        credenciales: []
    },
    {
        titulo: "Control De Cambios PJUD4",
        desc: "Registro acumulado de cambios de equipos del proyecto PJUD5 - planilla",
        url: "https://docs.google.com/spreadsheets/d/1lE8qh6PjsjNIPBNWxYOuyV8yTARmR2Zt9CSSxiA2qZk/edit?gid=0#gid=0",
        credenciales: [
             { label: "PIN", value: "pjud5upg", tipo: "text" }
        ]
    },
    {
        titulo: "Biblioteca de Comentarios",
        desc: "Accede a comentarios estandarizados y automatizados.",
        url: "https://mdapjud.netlify.app/#comentarios",
        credenciales: []
    },
    {
        titulo: "Consulta de Equipos",
        desc: "Revisa el estado de una serie, puedes ver proyecto, modelo, tipo.",
        url: "https://jmarrufo-gh.github.io/equiposfcom-web.github.io/",
        credenciales: []
    },
    {
        titulo: "Procedimientos Técnicos",
        desc: "Accede a todos los procedimientos tecnico del proyecto",
        url: "https://procedimientosmda.netlify.app/",
        credenciales: []
    },
    {
        titulo: "Registros de Bonos extraordinarios",
        desc: "Acceso para registro de bonos y alojamiento de personal.",
        url: "https://docs.google.com/spreadsheets/d/1AgFLn7CRFcykHris4CVU0T9awJX1brxIpPGYS8IXap0/edit?pli=1&gid=475953692#gid=475953692",
        credenciales: []
    }
];

function cargarBibliotecaEnlaces() {
    const tbody = document.getElementById('body-enlaces');
    if (!tbody) return;
    
    tbody.innerHTML = ""; 

    BIBLIOTECA_LINKS.forEach((item, indexLink) => {
        const tr = document.createElement('tr');
        tr.style.borderBottom = "1px solid #eee";
        
        let credsHTML = "";
        if (item.credenciales && item.credenciales.length > 0) {
            credsHTML = `<div style="display: flex; flex-direction: column; gap: 6px;">`;
            item.credenciales.forEach((cred, indexCred) => {
                const uniqueID = `secret-${indexLink}-${indexCred}`;
                const isPassword = cred.tipo === 'password';
                const displayValue = isPassword ? '••••••••' : cred.value;
                const eyeButton = isPassword 
                    ? `<button onclick="toggleSecret('${uniqueID}', '${cred.value}')" title="Mostrar/Ocultar" style="border: none; background: transparent; color: #888; cursor: pointer; margin-left: 5px;">
                         <i id="icon-${uniqueID}" class="fas fa-eye"></i>
                       </button>` 
                    : '';

                credsHTML += `
                    <div style="display: flex; justify-content: space-between; align-items: center; background: #fff; border: 1px solid #ddd; padding: 5px 10px; border-radius: 4px;">
                        <div style="display: flex; align-items: center; gap: 8px;">
                            <strong style="color:#014f8b; font-size: 0.85rem;">${cred.label}:</strong>
                            <span id="${uniqueID}" style="font-family: monospace; color: #444; font-size: 0.9rem;">${displayValue}</span>
                            ${eyeButton}
                        </div>
                        <button onclick="copiarTextoDirecto('${cred.value}')" title="Copiar ${cred.label}" style="border: 1px solid #ddd; background: #f8f9fa; color: #555; cursor: pointer; padding: 3px 8px; border-radius: 3px; font-size: 0.8rem;">
                            <i class="fas fa-copy"></i>
                        </button>
                    </div>
                `;
            });
            credsHTML += `</div>`;
        } else {
            credsHTML = `<span style="color: #999; font-size: 0.85rem;">Sin credenciales requeridas</span>`;
        }

        tr.innerHTML = `
            <td style="padding: 15px;">
                <strong style="color: #333; font-size: 0.95rem;">${item.titulo}</strong>
                <div style="font-size: 0.85rem; color: #666; margin-top: 4px;">${item.desc}</div>
            </td>
            <td style="padding: 15px; text-align: center; vertical-align: middle;">
                <a href="${item.url}" target="_blank" class="btn-primary" style="text-decoration: none; padding: 6px 12px; font-size: 0.85rem; background-color: #007bff;">
                    <i class="fas fa-external-link-alt"></i> Abrir
                </a>
            </td>
            <td style="padding: 10px; background-color: #f9f9f9; vertical-align: middle; width: 320px;">
                ${credsHTML}
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function toggleSecret(elementId, realValue) {
    const span = document.getElementById(elementId);
    const icon = document.getElementById('icon-' + elementId);
    if (span.innerText.includes('•')) {
        span.innerText = realValue; 
        icon.className = 'fas fa-eye-slash'; 
        icon.style.color = '#d9534f'; 
    } else {
        span.innerText = '••••••••'; 
        icon.className = 'fas fa-eye'; 
        icon.style.color = '#888';
    }
}

document.addEventListener('DOMContentLoaded', () => {
    cargarBibliotecaEnlaces();
});

// --- REPORTE: GUÍAS PENDIENTES ---
async function generarReporteGuiasPendientes() {
    const container = document.getElementById('guias-pendientes-result');
    const grupoSeleccionado = document.getElementById('filtro-grupo-guias').value;

    if (allTicketsList.length === 0) {
        showToast("⚠️ No hay datos cargados. Carga un archivo primero.");
        return;
    }

    // Si el mapa está vacío, intentamos cargarlo antes de seguir
    if (Object.keys(window.mapaFechasLab).length === 0) {
        await cargarDatosLaboratorio(); 
    }

    // ELIMINADO: await cargarDatosDesdeSheets() ya no es necesario aquí
    // porque los datos se cargan al abrir la página.

    const ticketsEnRecuperacion = [
        "7493206","7494261","7496812","7496968","7524642","7494261","7496969","7539950","7544696","7560826","7581956","7598371","7614098","7605150","7637221","7662905","7635910","7690441","7692876","7492515","7494920","7495317","7495546","7495776","7802224","7842786","7841094"
    ]; 

    const ticketsFiltrados = allTicketsList.filter(t => {
        const backupNorm = (t.backup || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
        if (backupNorm !== "SI") return false;

        const guiaNorm = (t.guiaRetiro || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
        if (guiaNorm !== "NO" && guiaNorm !== "") return false;

        const tipoUpper = t.tipo ? t.tipo.toUpperCase() : "";
        if (tipoUpper.includes("MOUSE") || tipoUpper.includes("TECLADO")) return false;

        const grupoUpper = (t.grupo || "").toUpperCase();
        if (grupoSeleccionado !== "TODOS") {
            if (!grupoUpper.includes(grupoSeleccionado)) return false;
        } else {
            if (!grupoUpper.includes("SCO") && !grupoUpper.includes("RESIDENTES")) return false;
        }
        return true; 
    });

    if (ticketsFiltrados.length === 0) {
        showToast("✅ No hay guías pendientes.");
        container.innerHTML = "<p style='text-align:center; color:#28a745; padding: 20px;'>No hay pendientes.</p>";
        container.classList.remove('hidden');
        return;
    }

    let filasHTML = "";
    
    // Usamos una referencia local segura de los datos del lab
    const labData = window.mapaFechasLab || mapaFechasLab || {};

    ticketsFiltrados.forEach(t => {
        const estaEnRecuperacion = ticketsEnRecuperacion.includes(t.num.trim());
        const badgeEstado = estaEnRecuperacion 
            ? `<span style="background:#fff3cd; color:#856404; padding:2px 6px; border-radius:4px; font-weight:bold; font-size:10px; border:1px solid #ffeeba;">CON USUARIO - REVISAR</span>`
            : `<span style="color:#666; font-size:10px;">PENDIENTE RETIRO</span>`;

        // CRUCE DINÁMICO
       // Dentro del forEach de generarReporteGuiasPendientes:
const serieTicket = (t.serie || "").toString().trim().toUpperCase();
// Usamos window para asegurar acceso a la memoria global
const fechaLab = window.mapaFechasLab[serieTicket] ? window.mapaFechasLab[serieTicket] : "-";

        filasHTML += `
            <tr style="border-bottom: 1px solid #ddd; ${estaEnRecuperacion ? 'background-color: #fffdf5;' : ''}">
                <td style="padding: 8px; border: 1px solid #ccc; text-align: center; font-weight:bold;">${t.num}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.fechaFin}</td>
                <td style="padding: 8px; border: 1px solid #ccc; text-align:center;">${badgeEstado}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.finalizadoPor || t.usuario}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.dependencia}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.tipo}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.serie}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.despachosRaw}</td>
                <td style="padding: 8px; border: 1px solid #ccc; font-weight:bold; color:#014f8b; text-align:center;">${fechaLab}</td>
            </tr>
        `;
    });

    container.innerHTML = `
        <h4 style="margin-bottom:10px; color:#b48b02;">
            <i class="fas fa-exclamation-triangle"></i> Guías Pendientes: ${ticketsFiltrados.length} casos
        </h4>
        <table id="tabla-guias-pendientes" style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px;">
            <thead style="background-color: #ffc107; color: #333;">
                <tr>
                    <th>Ticket</th>
                    <th>Finalización</th>
                    <th>Estado Físico</th>
                    <th>Finalizado por</th>
                    <th>Dependencia</th>
                    <th>Tipo</th>
                    <th>Serie</th>
                    <th>Despachos</th>
                    <th style="background-color: #014f8b; color: white;">FECHA LAB</th>
                </tr>
            </thead>
            <tbody>${filasHTML}</tbody>
        </table>
    `;

    container.classList.remove('hidden');
    showToast(`⚠️ Cruce de datos finalizado.`);
}

function copiarTablaGuiasPendientes() {
    const tabla = document.getElementById('tabla-guias-pendientes');
    if (!tabla) {
        showToast("⚠️ Primero genera el reporte.");
        return;
    }
    const range = document.createRange();
    range.selectNode(tabla);
    window.getSelection().removeAllRanges();
    window.getSelection().addRange(range);
    document.execCommand("copy");
    window.getSelection().removeAllRanges();
    showToast("✅ Tabla copiada al portapapeles");
}

function switchSheetTab(tabId, btnElement) {
    document.querySelectorAll('#mod-registros .sheet-content').forEach(content => {
        content.classList.add('hidden-content');
        content.classList.remove('active-content');
    });
    document.querySelectorAll('#mod-registros .tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    document.getElementById(tabId).classList.remove('hidden-content');
    document.getElementById(tabId).classList.add('active-content');
    btnElement.classList.add('active');
}

function clearTicketCoordination(ticketNum, btnElement) {
    try {
        let savedCoords = JSON.parse(localStorage.getItem('sco_coordinations') || '{}');
        if (savedCoords[ticketNum]) {
            delete savedCoords[ticketNum];
            localStorage.setItem('sco_coordinations', JSON.stringify(savedCoords));
        }
    } catch(e) {
        console.warn("No se pudo acceder a localStorage", e);
    }

    const ticketInCurrent = currentTicketsList.find(t => t.num === ticketNum);
    if (ticketInCurrent) {
        ticketInCurrent.fechaCoord = "";
        ticketInCurrent.horaCoord = "";
        ticketInCurrent.tecnicoCoord = "";
    }
    
    const ticketInAll = allTicketsList.find(t => t.num === ticketNum);
    if (ticketInAll) {
        ticketInAll.fechaCoord = "";
        ticketInAll.horaCoord = "";
        ticketInAll.tecnicoCoord = "";
    }

    const fila = btnElement.closest('tr');
    if (fila) {
        fila.children[7].innerHTML = "-";
        fila.children[8].innerHTML = "-";
        
        fila.style.backgroundColor = "#f8d7da";
        setTimeout(() => fila.style.backgroundColor = "", 800);
    }

    if (selectedTicketData && selectedTicketData.num === ticketNum) {
        document.getElementById('coord-date').value = "";
        document.getElementById('coord-time').value = "";
        document.getElementById('tech-name').value = "";
    }

    showToast(`🧹 Coordinación eliminada para TK ${ticketNum}`);
}

function ejecutarScriptPython() {
    alert("Esta función requiere un servidor backend y no está disponible en la versión estática de GitHub Pages. Por favor adjunta el archivo manualmente usando el botón 'Adjuntar Reporte'.");
}

function switchFormatTab(tabId, btnElement) {
    document.querySelectorAll('#mod-guias .tab-content').forEach(content => {
        content.classList.add('hidden-content');
        content.classList.remove('active-content');
    });
    document.querySelectorAll('#mod-guias .tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    document.getElementById(tabId).classList.remove('hidden-content');
    document.getElementById(tabId).classList.add('active-content');
    btnElement.classList.add('active');
}

function logout() {
    window.location.reload();
}

function copiarTextoRendicion() {
    const texto = `Para rendir Actividades, 
Emitir Boleta 
Con los datos de:
Fcom Spa 76.741.749-2
Dirección: Huerfanos 670 Santiago

Debe considerar en su emisión, la rendición de pasajes, actividad y algún servicio que deba rendirse asociado a la actividad realizada. Detallar el numero de requerimiento.

Una Vez emitida la boleta debe enviarla a los siguientes correos:
Ricardo Cornejo <rcornejo@fcom.cl> Carlos Zapata <c.zapata@fcom.cl>

Adjuntar la boleta, y la OT (Orden de trabajo) de la actividad, esta OT es la que Firma el usuario al finalizar el proceso.`;

    if (navigator.clipboard) {
        navigator.clipboard.writeText(texto).then(() => {
            showToast("Texto de Rendición copiado");
        }).catch(err => {
            console.error("Error al copiar", err);
        });
    } else {
         const textArea = document.createElement("textarea");
         textArea.value = texto;
         document.body.appendChild(textArea);
         textArea.select();
         try {
             document.execCommand('copy');
             showToast("Texto de Rendición copiado");
         } catch (err) {
             console.error("Error Fallback", err);
         }
         document.body.removeChild(textArea);
    }
}

// =========================================================
// === MÓDULO: INFORME MENSUAL MDA (CON AUDITORÍA QA Y DESGLOSE DE ANULADOS LADO A LADO) ===
// =========================================================

// --- 1. DICCIONARIO DE ESTADOS ---
const MAPPING_SOLUCIONES_INFORME = {
    'CONFIGURACIÓN': 'Solucionados en MDA',
    'DERIVADO TERRENO': 'Derivado a Residencia',
    'DERIVADO SCO': 'Derivado a SCO',
    'DERIVADO OTRA AREA': 'Derivado a Otra Área',
    'ANULADO NO CONTACTO': 'Anulado no Contacto',
    'ANULADO USUARIO': 'Anulado Usuario',
    'ANULADO': 'Anulado no Contacto', 
    'HABILITACION SCO': 'Habilitacion x SCO'
};

// --- 2. CATÁLOGO DE RESIDENCIAS (Para Control de Calidad) ---
const CATALOGO_RESIDENCIAS = [
    // CAPJ
    "CORPORACION ADMINISTRATIVA CENTRAL", "CORPORACION ADMINISTRATIVA CENTRAL - DEPARTAMENTO DE PLANIFICACION", "CORPORACION ADMINISTRATIVA CENTRAL - JUSTICIA MOVIL", "CORPORACION ADMINISTRATIVA CENTRAL - FINANZAS", "CORPORACION ADMINISTRATIVA CENTRAL - UNIDAD JURIDICA", "CORPORACION ADMINISTRATIVA CENTRAL - INFORMATICA", "CORPORACION ADMINISTRATIVA CENTRAL - ADQUISICIONES Y MANTENIMIENTO", "CORPORACION ADMINISTRATIVA CENTRAL - CENTRO DOCUMENTAL", "CORPORACION ADMINISTRATIVA CENTRAL - RECURSOS HUMANOS", "CORPORACION ADMINISTRATIVA CENTRAL - CENTRO DE NOTIFICACIONES", "CORPORACION ADMINISTRATIVA CENTRAL - BIENESTAR", "CORPORACION ADMINISTRATIVA DE SANTIAGO",
    // CENTRO DE JUSTICIA
    "2° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "10° JUZGADO DE GARANTIA DE SANTIAGO", "15° JUZGADO DE GARANTIA DE SANTIAGO", "11° JUZGADO DE GARANTIA DE SANTIAGO", "12° JUZGADO DE GARANTIA DE SANTIAGO", "7° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "8° JUZGADO DE GARANTIA DE SANTIAGO", "2° JUZGADO DE GARANTIA DE SANTIAGO", "6° JUZGADO DE GARANTIA DE SANTIAGO", "6° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "14°JUZGADO DE GARANTIA DE SANTIAGO", "1° JUZGADO DE GARANTIA DE SANTIAGO", "34° JUZGADO DEL CRIMEN DE SANTIAGO", "7° JUZGADO DE GARANTIA DE SANTIAGO", "14° JUZGADO DE GARANTIA DE SANTIAGO", "1° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "13° JUZGADO DE GARANTIA DE SANTIAGO", "3° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "4° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "5° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "9° JUZGADO DE GARANTIA DE SANTIAGO", "4° JUZGADO DE GARANTIA DE SANTIAGO", "3° JUZGADO DE GARANTIA DE SANTIAGO", "5° JUZGADO DE GARANTIA DE SANTIAGO", "CENTRO DE JUSTICIA DE SANTIAGO", "CENTRO DE JUSTICIA DE SANTIAGO - ADMINISTRACION CJS", "CENTRO DE NOTIFICACIONES JUDICIALES C.J. SANTIAGO", "ADMINISTRACION EXTERNA CENTRO DE JUSTICIA DE STGO.",
    // CIVILES
     "27° Juzgado Civil De Santiago","27º Juzgado Civil De Santiago","4° JUZGADO CIVIL DE SANTIAGO","4º JUZGADO CIVIL DE SANTIAGO","10° JUZGADO CIVIL DE SANTIAGO", "11° JUZGADO CIVIL DE SANTIAGO", "18° JUZGADO CIVIL DE SANTIAGO", "22° JUZGADO CIVIL DE SANTIAGO", "23° JUZGADO CIVIL DE SANTIAGO", "6° JUZGADO CIVIL DE SANTIAGO", "9° JUZGADO CIVIL DE SANTIAGO", "15° JUZGADO CIVIL DE SANTIAGO", "16° JUZGADO CIVIL DE SANTIAGO", "20° JUZGADO CIVIL DE SANTIAGO", "21° JUZGADO CIVIL DE SANTIAGO", "24° JUZGADO CIVIL DE SANTIAGO", "7° JUZGADO CIVIL DE SANTIAGO", "12° JUZGADO CIVIL DE SANTIAGO", "13° JUZGADO CIVIL DE SANTIAGO", "14° JUZGADO CIVIL DE SANTIAGO", "17° JUZGADO CIVIL DE SANTIAGO", "19° JUZGADO CIVIL DE SANTIAGO", "1° JUZGADO CIVIL DE SANTIAGO", "9°JUZGADO CIVIL DE SANTIAGO", "30° JUZGADO CIVIL DE SANTIAGO", "8° JUZGADO CIVIL DE SANTIAGO", "26° JUZGADO CIVIL DE SANTIAGO", "25° JUZGADO CIVIL DE SANTIAGO", "27° JUZGADO CIVIL DE SANTIAGO", "28° JUZGADO CIVIL DE SANTIAGO", "29° JUZGADO CIVIL DE SANTIAGO", "2° JUZGADO CIVIL DE SANTIAGO", "5° JUZGADO CIVIL DE SANTIAGO", "4° JUZGADO CIVIL DE SANTIAGO", "3° JUZGADO CIVIL DE SANTIAGO", "APOYO JUZGADOS CIVILES", "CENTRO APOYO JUZGADOS CIVILES Y LABORALES SANTIAGO",
    // CONCEPCION
     "Juzgado De Cobranza Laboral Y Previsional De Concepcion","Juzgado De Cobranza Laboral Y Previsional De ConcepciÓn","1° JUZGADO CIVIL DE CONCEPCION", "2° JUZGADO CIVIL DE CONCEPCION", "JUZGADO DE FAMILIA DE CONCEPCION", "CENTRO DE NOTIFICACIONES DE CONCEPCION", "CORPORACION ADMINISTRATIVA DE CONCEPCION", "JUZGADO DE LETRAS DEL TRABAJO DE CONCEPCION", "CORTE DE APELACIONES DE CONCEPCION", "3° JUZGADO CIVIL DE CONCEPCION",
    // CORTE SUPREMA
    "CORTE SUPREMA DE JUSTICIA", "CORTE DE APELACIONES DE SANTIAGO", "CORTE SUPREMA DE JUSTICIA - BIBLIOTECA CS", "CORTE SUPREMA DE JUSTICIA - OFICINA DE EXTRADICION CS", "CORTE SUPREMA DE JUSTICIA - MINISTROS CS", "CORTE SUPREMA DE JUSTICIA - RELATORES CS", "CORTE SUPREMA DE JUSTICIA - FISCALIA CS", "UNIDAD DE PROTECCIONES I.C.A. DE SANTIAGO",
    // FAMILIA
    "JUZGADO DE COBRANZA LABORAL Y PREVISIONAL DE SANTI","2° JUZGADO DE FAMILIA DE SANTIAGO", "4° JUZGADO DE FAMILIA DE SANTIAGO", "3° JUZGADO DE FAMILIA DE SANTIAGO", "1° JUZGADO DE FAMILIA DE SANTIAGO", "1° JUZGADO DE LETRAS DEL TRABAJO DE SANTIAGO", "CENTRO ATENCION ASUNTOS DE FAMILIA DE SANTIAGO", "UNIDAD ADMTVA. TRAMIT. LAB. Y COBR. PREV. SANTIAGO", "JUZGADO DE COBRANZA LABORAL Y PREVISIONAL DE SANTIAGO", "CENTRO MEDIDAS CAUTELARES JDOS. FAMILIA SANTIAGO",
    // SAN MIGUEL
    "CORTE DE APELACIONES DE SAN MIGUEL", "JUZGADO DE LETRAS DEL TRABAJO DE SAN MIGUEL", "CORPORACION ADMINISTRATIVA DE SAN MIGUEL", "1° JUZGADO DE FAMILIA DE SAN MIGUEL", "2° JUZGADO DE FAMILIA DE SAN MIGUEL"
];

// Variables globales para copiar tickets
window.ticketsOtrosMDA = [];
window.ticketsQAErrores = [];

// ==========================================
// --- FUNCIONES DE COPIADO PARA FALLAS ---
// ==========================================
function copiarColumnaHardware() {
    if (!window.ultimaDataFallas) return showToast("⚠️ No hay datos para copiar.");
    const text = window.ultimaDataFallas.orden.map(j => window.ultimaDataFallas.datos[j].hardware).join('\n');
    copiarTextoDirecto(text);
}

function copiarColumnaSoftware() {
    if (!window.ultimaDataFallas) return showToast("⚠️ No hay datos para copiar.");
    const text = window.ultimaDataFallas.orden.map(j => window.ultimaDataFallas.datos[j].software).join('\n');
    copiarTextoDirecto(text);
}

function copiarColumnaTotalFallas() {
    if (!window.ultimaDataFallas) return showToast("⚠️ No hay datos para copiar.");
    const text = window.ultimaDataFallas.orden.map(j => window.ultimaDataFallas.datos[j].hardware + window.ultimaDataFallas.datos[j].software).join('\n');
    copiarTextoDirecto(text);
}
// ==========================================

function generarInformeMDA() {
    const fileInput = document.getElementById('csv-informe-mda');
    
    if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
        return showToast("⚠️ Por favor carga el archivo CSV limpio primero.");
    }

    const resultContainer = document.getElementById('mda-result-container');
    resultContainer.classList.add('hidden'); 
    resultContainer.innerHTML = '<p style="color:#666; padding: 15px;">Procesando datos y generando calendario completo...</p>'; 

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = function(e) {
        const rawData = e.target.result;
        const countComas = (rawData.match(/,/g) || []).length;
        const countPuntoComas = (rawData.match(/;/g) || []).length;
        const delimiter = countPuntoComas > countComas ? ';' : ',';
        const rows = parseCSV(rawData, delimiter);

        const contadoresFinales = {
            'Anulado no Contacto': 0, 'Anulado Usuario': 0, 'Solucionados en MDA': 0,
            'Derivado a Otra Área': 0, 'Derivado a SCO': 0, 'Derivado a Residencia': 0,
            'Habilitacion x SCO': 0, 'No es Residencia': 0, 'Otros / Sin Clasificar': 0 
        };

        const contadoresResidencias = {
            'CAPJ': 0, 'CENTRO DE JUSTICIA': 0, 'CIVILES': 0, 'CONCEPCION': 0, 'CORTE SUPREMA': 0, 'FAMILIA': 0, 'SAN MIGUEL': 0
        };
        let totalResidencias = 0;

        const reporteFallas = {};

        const catCAPJ = ["CORPORACION ADMINISTRATIVA CENTRAL", "CORPORACION ADMINISTRATIVA CENTRAL - DEPARTAMENTO DE PLANIFICACION", "CORPORACION ADMINISTRATIVA CENTRAL - JUSTICIA MOVIL", "CORPORACION ADMINISTRATIVA CENTRAL - FINANZAS", "CORPORACION ADMINISTRATIVA CENTRAL - UNIDAD JURIDICA", "CORPORACION ADMINISTRATIVA CENTRAL - INFORMATICA", "CORPORACION ADMINISTRATIVA CENTRAL - ADQUISICIONES Y MANTENIMIENTO", "CORPORACION ADMINISTRATIVA CENTRAL - CENTRO DOCUMENTAL", "CORPORACION ADMINISTRATIVA CENTRAL - RECURSOS HUMANOS", "CORPORACION ADMINISTRATIVA CENTRAL - CENTRO DE NOTIFICACIONES", "CORPORACION ADMINISTRATIVA CENTRAL - BIENESTAR", "CORPORACION ADMINISTRATIVA DE SANTIAGO"];
        const catCJ = ["2° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "10° JUZGADO DE GARANTIA DE SANTIAGO", "15° JUZGADO DE GARANTIA DE SANTIAGO", "11° JUZGADO DE GARANTIA DE SANTIAGO", "12° JUZGADO DE GARANTIA DE SANTIAGO", "7° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "8° JUZGADO DE GARANTIA DE SANTIAGO", "2° JUZGADO DE GARANTIA DE SANTIAGO", "6° JUZGADO DE GARANTIA DE SANTIAGO", "6° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "14°JUZGADO DE GARANTIA DE SANTIAGO", "1° JUZGADO DE GARANTIA DE SANTIAGO", "34° JUZGADO DEL CRIMEN DE SANTIAGO", "7° JUZGADO DE GARANTIA DE SANTIAGO", "14° JUZGADO DE GARANTIA DE SANTIAGO", "1° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "13° JUZGADO DE GARANTIA DE SANTIAGO", "3° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "4° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "5° TRIBUNAL DE JUICIO ORAL EN LO PENAL DE SANTIAGO", "9° JUZGADO DE GARANTIA DE SANTIAGO", "4° JUZGADO DE GARANTIA DE SANTIAGO", "3° JUZGADO DE GARANTIA DE SANTIAGO", "5° JUZGADO DE GARANTIA DE SANTIAGO", "CENTRO DE JUSTICIA DE SANTIAGO", "CENTRO DE JUSTICIA DE SANTIAGO - ADMINISTRACION CJS", "CENTRO DE NOTIFICACIONES JUDICIALES C.J. SANTIAGO", "ADMINISTRACION EXTERNA CENTRO DE JUSTICIA DE STGO."];
        const catCiviles = ["4º JUZGADO CIVIL DE SANTIAGO","10° JUZGADO CIVIL DE SANTIAGO", "11° JUZGADO CIVIL DE SANTIAGO", "18° JUZGADO CIVIL DE SANTIAGO", "22° JUZGADO CIVIL DE SANTIAGO", "23° JUZGADO CIVIL DE SANTIAGO", "6° JUZGADO CIVIL DE SANTIAGO", "9° JUZGADO CIVIL DE SANTIAGO", "15° JUZGADO CIVIL DE SANTIAGO", "16° JUZGADO CIVIL DE SANTIAGO", "20° JUZGADO CIVIL DE SANTIAGO", "21° JUZGADO CIVIL DE SANTIAGO", "24° JUZGADO CIVIL DE SANTIAGO", "7° JUZGADO CIVIL DE SANTIAGO", "12° JUZGADO CIVIL DE SANTIAGO", "13° JUZGADO CIVIL DE SANTIAGO", "14° JUZGADO CIVIL DE SANTIAGO", "17° JUZGADO CIVIL DE SANTIAGO", "19° JUZGADO CIVIL DE SANTIAGO", "1° JUZGADO CIVIL DE SANTIAGO", "9°JUZGADO CIVIL DE SANTIAGO", "30° JUZGADO CIVIL DE SANTIAGO", "8° JUZGADO CIVIL DE SANTIAGO", "26° JUZGADO CIVIL DE SANTIAGO", "25° JUZGADO CIVIL DE SANTIAGO", "27° JUZGADO CIVIL DE SANTIAGO", "28° JUZGADO CIVIL DE SANTIAGO", "29° JUZGADO CIVIL DE SANTIAGO", "2° JUZGADO CIVIL DE SANTIAGO", "5° JUZGADO CIVIL DE SANTIAGO", "4° JUZGADO CIVIL DE SANTIAGO", "3° JUZGADO CIVIL DE SANTIAGO", "APOYO JUZGADOS CIVILES", "CENTRO APOYO JUZGADOS CIVILES Y LABORALES SANTIAGO"];
        const catConcepcion = ["1° JUZGADO CIVIL DE CONCEPCION", "2° JUZGADO CIVIL DE CONCEPCION", "JUZGADO DE FAMILIA DE CONCEPCION", "CENTRO DE NOTIFICACIONES DE CONCEPCION", "CORPORACION ADMINISTRATIVA DE CONCEPCION", "JUZGADO DE LETRAS DEL TRABAJO DE CONCEPCION", "CORTE DE APELACIONES DE CONCEPCION", "3° JUZGADO CIVIL DE CONCEPCION"];
        const catSuprema = ["CORTE SUPREMA DE JUSTICIA", "CORTE DE APELACIONES DE SANTIAGO", "CORTE SUPREMA DE JUSTICIA - BIBLIOTECA CS", "CORTE SUPREMA DE JUSTICIA - OFICINA DE EXTRADICION CS", "CORTE SUPREMA DE JUSTICIA - MINISTROS CS", "CORTE SUPREMA DE JUSTICIA - RELATORES CS", "CORTE SUPREMA DE JUSTICIA - FISCALIA CS", "UNIDAD DE PROTECCIONES I.C.A. DE SANTIAGO"];
        const catFamilia = ["JUZGADO DE COBRANZA LABORAL Y PREVISIONAL DE SANTI","2° JUZGADO DE FAMILIA DE SANTIAGO", "4° JUZGADO DE FAMILIA DE SANTIAGO", "3° JUZGADO DE FAMILIA DE SANTIAGO", "1° JUZGADO DE FAMILIA DE SANTIAGO", "1° JUZGADO DE LETRAS DEL TRABAJO DE SANTIAGO", "CENTRO ATENCION ASUNTOS DE FAMILIA DE SANTIAGO", "UNIDAD ADMTVA. TRAMIT. LAB. Y COBR. PREV. SANTIAGO", "JUZGADO DE COBRANZA LABORAL Y PREVISIONAL DE SANTIAGO", "CENTRO MEDIDAS CAUTELARES JDOS. FAMILIA SANTIAGO"];
        const catSanMiguel = ["CORTE DE APELACIONES DE SAN MIGUEL", "JUZGADO DE LETRAS DEL TRABAJO DE SAN MIGUEL", "CORPORACION ADMINISTRATIVA DE SAN MIGUEL", "1° JUZGADO DE FAMILIA DE SAN MIGUEL", "2° JUZGADO DE FAMILIA DE SAN MIGUEL"];

        const getCategoriaResidencia = (depClean) => {
            if (catCAPJ.some(c => limpiarTexto(c) === depClean)) return 'CAPJ';
            if (catCJ.some(c => limpiarTexto(c) === depClean)) return 'CENTRO DE JUSTICIA';
            if (catCiviles.some(c => limpiarTexto(c) === depClean)) return 'CIVILES';
            if (catConcepcion.some(c => limpiarTexto(c) === depClean)) return 'CONCEPCION';
            if (catSuprema.some(c => limpiarTexto(c) === depClean)) return 'CORTE SUPREMA';
            if (catFamilia.some(c => limpiarTexto(c) === depClean)) return 'FAMILIA';
            if (catSanMiguel.some(c => limpiarTexto(c) === depClean)) return 'SAN MIGUEL';
            return null;
        };
        
        const ingresoDiario = {}; 
        let totalTicketsMesFiltro = 0;
        let mesDetectado = null;
        let anioDetectado = null;

        window.ticketsOtrosMDA = []; 
        window.ticketsQAErrores = []; 
        window.ticketsSinJurisdiccion = []; // NUEVO
        window.ticketsSinFalla = []; // NUEVO

        const limpiarTexto = (texto) => texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();


        const conteoEquiposParaTabla = {};
        rows.forEach((cols, index) => {
            // Agrega esto al inicio del rows.forEach
        const tipoEquipoCSV = cols[10] ? limpiarTexto(cols[10]) : "";
        if (tipoEquipoCSV !== "") {
            conteoEquiposParaTabla[tipoEquipoCSV] = (conteoEquiposParaTabla[tipoEquipoCSV] || 0) + 1;
        }
            if (index === 0 || cols.length < 5) return; 

            const numTicket = cols[1] ? cols[1].trim() : "";
            const fechaFull = cols[5] ? cols[5].split(' ')[0] : "Sin Fecha"; 
            
            // --- JURISDICCION Y FALLAS ---
            let jurisdiccionRaw = cols[7] ? cols[7].replace(/"/g, "").trim() : "En blanco"; 
            if (jurisdiccionRaw === "" || jurisdiccionRaw === "-" || jurisdiccionRaw.toUpperCase() === "N/A") {
                jurisdiccionRaw = "En blanco";
            }
            
            // NUEVO: Atrapa los tickets sin jurisdicción para auditar
            if (jurisdiccionRaw === "En blanco" && numTicket !== "") {
                window.ticketsSinJurisdiccion.push(numTicket);
            }
            
            const tipoFallaRaw = cols[11] ? cols[11].replace(/"/g, "").trim().toUpperCase() : ""; 
            
            if (!reporteFallas[jurisdiccionRaw]) {
                reporteFallas[jurisdiccionRaw] = { hardware: 0, software: 0, vacia: 0 };
            }
            
            if (tipoFallaRaw.includes("HARDWARE")) {
                reporteFallas[jurisdiccionRaw].hardware++;
            } else if (tipoFallaRaw.includes("SOFTWARE")) {
                reporteFallas[jurisdiccionRaw].software++;
            } else {
                reporteFallas[jurisdiccionRaw].vacia++;
                // NUEVO: Atrapa los tickets sin tipo de falla para auditar
                if (numTicket !== "") {
                    window.ticketsSinFalla.push(numTicket);
                }
            }

            if (fechaFull !== "Sin Fecha") {
                // 1. Detectamos el separador que traiga el CSV (- o /)
                const separadorOriginal = fechaFull.includes('/') ? '/' : '-';
                const partesFecha = fechaFull.split(separadorOriginal);
                
                const d = partesFecha[0].padStart(2, '0');
                const m = partesFecha[1].padStart(2, '0');
                const y = partesFecha[2];

                // 2. NORMALIZAMOS SIEMPRE A SLASH (/) 
                // Esto es vital para que coincida con las llaves de la tabla diaria
                const fechaNormalizada = `${d}/${m}/${y}`;

                if (!mesDetectado) { 
                    mesDetectado = parseInt(m); 
                    anioDetectado = parseInt(y); 
                }
                
                const dependencia = cols[6] ? limpiarTexto(cols[6]) : ""; 
                const rawSolMDA = cols[14] ? cols[14].trim().toUpperCase() : ""; 
                let estadoMapped = MAPPING_SOLUCIONES_INFORME[rawSolMDA] || 'Otros / Sin Clasificar';
                const esResidencia = (rawSolMDA === "DERIVADO TERRENO");

                if (estadoMapped === 'Derivado a Residencia') {
                    if (!CATALOGO_RESIDENCIAS.includes(dependencia)) {
                        estadoMapped = 'No es Residencia';
                        if (numTicket !== "") window.ticketsQAErrores.push(numTicket);
                    } else {
                        const catResidencia = getCategoriaResidencia(dependencia);
                        if (catResidencia) {
                            contadoresResidencias[catResidencia]++;
                            totalResidencias++;
                        }
                    }
                }

                if (estadoMapped === 'Otros / Sin Clasificar' && numTicket !== "") {
                    window.ticketsOtrosMDA.push(numTicket);
                }

                contadoresFinales[estadoMapped]++;
                totalTicketsMesFiltro++;

                // 3. USAMOS LA FECHA NORMALIZADA CON SLASH
                if (!ingresoDiario[fechaNormalizada]) ingresoDiario[fechaNormalizada] = { mda: 0, resi: 0 };
                if (esResidencia) {
                    ingresoDiario[fechaNormalizada].resi++;
                } else {
                    ingresoDiario[fechaNormalizada].mda++;
                }
            }
        });

        // === NUEVO: PROCESAMIENTO TABLA AGENTES ===
        const agentesData = {};
        const totalesPorEstado = { 'ANULADO': 0, 'SOLUCIONADO': 0, 'DERIVADO AREA': 0, 'DERIVADO SCO': 0, 'HABILITACION SCO': 0, 'RESIDENTES': 0 };

        rows.forEach((cols, idx) => {
            if (idx === 0 || cols.length < 5) return;
            const agenteNom = cols[cols.length - 1] ? cols[cols.length - 1].trim() : "SIN AGENTE";
            const solMDA = cols[14] ? cols[14].trim().toUpperCase() : "";

            let eKey = "";
            if (solMDA.includes("ANULADO")) eKey = "ANULADO";
            else if (solMDA === "CONFIGURACIÓN") eKey = "SOLUCIONADO";
            else if (solMDA === "DERIVADO OTRA AREA") eKey = "DERIVADO AREA";
            else if (solMDA === "DERIVADO SCO") eKey = "DERIVADO SCO";
            else if (solMDA === "HABILITACION SCO") eKey = "HABILITACION SCO";
            else if (solMDA === "DERIVADO TERRENO") eKey = "RESIDENTES";

            if (eKey !== "" && agenteNom !== "SIN AGENTE" && agenteNom !== "TECNICO") {
                if (!agentesData[agenteNom]) agentesData[agenteNom] = { 'ANULADO': 0, 'SOLUCIONADO': 0, 'DERIVADO AREA': 0, 'DERIVADO SCO': 0, 'HABILITACION SCO': 0, 'RESIDENTES': 0, 'TOTAL': 0 };
                agentesData[agenteNom][eKey]++;
                agentesData[agenteNom]['TOTAL']++;
                totalesPorEstado[eKey]++;
            }
        });

        let filasAgentesHTML = "";
        const sumaTotalGralAgentes = Object.values(totalesPorEstado).reduce((a, b) => a + b, 0);
        Object.keys(agentesData).sort().forEach(ag => {
            const d = agentesData[ag];
            const p = (val, tot) => tot > 0 ? ((val / tot) * 100).toFixed(2) + '%' : '0.00%';
            filasAgentesHTML += `
                <tr style="border-bottom: 1px solid #ddd; font-size: 12px; text-align: center;">
                    <td style="text-align: left; padding: 5px 10px; font-weight: bold; background: #f9f9f9;">${ag}</td>
                    <td>${p(d['ANULADO'], totalesPorEstado['ANULADO'])}</td>
                    <td>${p(d['SOLUCIONADO'], totalesPorEstado['SOLUCIONADO'])}</td>
                    <td>${p(d['DERIVADO AREA'], totalesPorEstado['DERIVADO AREA'])}</td>
                    <td>${p(d['DERIVADO SCO'], totalesPorEstado['DERIVADO SCO'])}</td>
                    <td>${p(d['HABILITACION SCO'], totalesPorEstado['HABILITACION SCO'])}</td>
                    <td>${p(d['RESIDENTES'], totalesPorEstado['RESIDENTES'])}</td>
                    <td style="font-weight: bold; background: #f0f4f8;">${p(d['TOTAL'], sumaTotalGralAgentes)}</td>
                </tr>`;
        });
const totalAnuladosCorrecto = contadoresFinales['Anulado no Contacto'] + contadoresFinales['Anulado Usuario'];
        const totalGeneralCorrecto = totalTicketsMesFiltro;

// === VARIABLE CON EL DISEÑO DE LA TABLA ===
        const tablaComparativaAgentesHTML = `
            <div class="card" style="border: 1px solid #ddd; padding: 0; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-top: 20px;">
                <div class="pjud-header-toggle" onclick="toggleSection('mda-agentes-body', 'icon-mda-agentes')" style="background: #f8f9fa; border-top: 4px solid #315e9a; padding: 12px 20px; display: flex; justify-content: space-between; align-items: center; cursor: pointer; border-radius: 4px 4px 0 0;">
                    <h4 style="color: #315e9a; margin: 0; font-family: Segoe UI, sans-serif; font-weight: bold;">
                        <i class="fas fa-users" style="margin-right: 8px;"></i> Comparativa de agentes de MDA
                    </h4>
                    <i id="icon-mda-agentes" class="fas fa-chevron-up rotate-icon" style="color: #315e9a;"></i>
                </div>
                
                <div id="mda-agentes-body" class="hidden-content" style="padding: 20px; background: #fff; border-radius: 0 0 4px 4px;">
                    <p style="margin-bottom: 15px; color: #555; font-size: 14px;">
                        Detalle porcentual de atenciones por agente según el estado del ticket en el mes seleccionado.
                    </p>
                    <div style="overflow-x: auto;">
                        <table style="width: 100%; border-collapse: collapse; font-family: Calibri, sans-serif; font-size: 12px; border: 1px solid #1a3b6c;">
                           <thead>
    <tr style="background-color: #315e9a; color: white !important; text-align: center;">
        <th style="padding: 8px; text-align: left; background-color: #315e9a; color: white !important;">AGENTE MESA</th>
        <th style="background-color: #315e9a; color: white !important;">ANULADO</th>
        <th style="background-color: #315e9a; color: white !important;">SOLUCIONADO</th>
        <th style="background-color: #315e9a; color: white !important;">DERIVADO AREA</th>
        <th style="background-color: #315e9a; color: white !important;">DERIVADO SCO</th>
        <th style="background-color: #315e9a; color: white !important;">HABILITACION SCO</th>
        <th style="background-color: #315e9a; color: white !important;">DERIVADO RESIDENTES</th>
        <th style="background-color: #1a3b6c; color: white !important;">SUMA TOTAL</th>
    </tr>
</thead>
                            <tbody>${filasAgentesHTML}</tbody>

<tfoot style="background-color: #315e9a; color: white; font-weight: bold; text-align: center;">
    <tr>
        <td style="text-align: left; padding: 8px;">Cantidad total</td>
        <td>${totalAnuladosCorrecto}</td>
        <td>${contadoresFinales['Solucionados en MDA']}</td>
        <td>${contadoresFinales['Derivado a Otra Área']}</td>
        <td>${contadoresFinales['Derivado a SCO']}</td>
        <td>${contadoresFinales['Habilitacion x SCO']}</td>
        <td>${contadoresFinales['Derivado a Residencia']}</td>
        <td style="background: #1a3b6c;">${totalGeneralCorrecto}</td>
    </tr>
</tfoot>    
                        </table>
                    </div>
                </div>
            </div>`;

        // === INYECTAR EN EL CONTENEDOR DEL HTML ===
        const containerAgentes = document.getElementById('mda-agentes-container');
        if (containerAgentes) {
            containerAgentes.innerHTML = tablaComparativaAgentesHTML;
        }


        // === FIN BLOQUE PROCESAMIENTO ===

        window.ultimoConteoResidencias = contadoresResidencias;
        window.ultimoTotalResidencias = totalResidencias;

        // --- LÓGICA DE CALENDARIO COMPLETO ---
        let dailyRowsHTML = "";
        const colFechas = [], colMDA = [], colResi = [], colTotal = [];
        const mesesNombres = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
        const labelsGrafico = [];

        if (mesDetectado && anioDetectado) {
            const ultimoDia = new Date(anioDetectado, mesDetectado, 0).getDate();
            for (let d = 1; d <= ultimoDia; d++) {
                const diaStr = d.toString().padStart(2, '0');
                const mesStr = mesDetectado.toString().padStart(2, '0');
                const fechaClave = `${diaStr}/${mesStr}/${anioDetectado}`;
                const datosDia = ingresoDiario[fechaClave] || { mda: 0, resi: 0 };
                const sum = datosDia.mda + datosDia.resi;
                const fechaObj = new Date(anioDetectado, mesDetectado - 1, d);
                const numeroDiaSemana = fechaObj.getDay(); 

                let colorFondo = "#ffffff"; 
                if (numeroDiaSemana === 6) colorFondo = "#d3efff"; 
                if (numeroDiaSemana === 0) colorFondo = "#FFEBEE"; 

                colFechas.push(fechaClave);
                colMDA.push(datosDia.mda);
                colResi.push(datosDia.resi);
                colTotal.push(sum);
                labelsGrafico.push(`${diaStr}-${mesesNombres[mesDetectado-1]}`);

                dailyRowsHTML += `<tr style="border-bottom: 1px solid #ddd; font-size: 13px;">
                    <td style="padding: 2px 8px; background-color: ${colorFondo} !important;">${fechaClave}</td>
                    <td style="padding: 2px 8px; text-align:center; color:#014f8b; background-color: ${colorFondo} !important;">${datosDia.mda}</td>
                    <td style="padding: 2px 8px; text-align:center; color:#2e7d32; background-color: ${colorFondo} !important;">${datosDia.resi}</td>
                    <td style="padding: 2px 8px; text-align:center; font-weight:bold; background-color: ${colorFondo} !important;">${sum}</td>
                </tr>`;
            }
        }

        // --- TABLA 1: ESTADOS ---
        const ordenEstados = ['Anulado no Contacto', 'Anulado Usuario', 'Solucionados en MDA', 'Derivado a Otra Área', 'Derivado a SCO', 'Derivado a Residencia', 'Habilitacion x SCO', 'No es Residencia', 'Otros / Sin Clasificar'];
        let rowsHTML = "";
        ordenEstados.forEach(estado => {
            const foliosCount = contadoresFinales[estado];
            if (estado === 'No es Residencia' && foliosCount === 0) return;
            const porcentaje = ((foliosCount / totalTicketsMesFiltro) * 100).toFixed(2); 
            const colorTexto = (estado.includes('Residencia') && estado === 'No es Residencia' || estado.includes('Otros')) && foliosCount > 0 ? '#dc3545' : '#333';
            const colorNegrita = (estado.includes('Residencia') && estado === 'No es Residencia' || estado.includes('Otros')) && foliosCount > 0 ? 'bold' : 'normal';

            let botonCopiarInfo = "";
            if (estado === 'Otros / Sin Clasificar' && foliosCount > 0) {
                botonCopiarInfo = `<button onclick="copiarArray(window.ticketsOtrosMDA)" class="btn-copy-small" style="background-color: #dc3545; color: white; border:none; border-radius:3px; padding:2px 6px; font-size:10px; cursor:pointer;"><i class="fas fa-copy"></i> Copiar TK</button>`;
            } else if (estado === 'No es Residencia' && foliosCount > 0) {
                botonCopiarInfo = `<button onclick="copiarArray(window.ticketsQAErrores)" class="btn-copy-small" style="background-color: #ff8c00; color: white; border:none; border-radius:3px; padding:2px 6px; font-size:10px; cursor:pointer;"><i class="fas fa-search"></i> Auditar TK</button>`;
            }
            rowsHTML += `<tr style="border-bottom: 1px solid #ddd; font-size: 13px;">
                <td style="padding: 3px 10px; color: ${colorTexto}; font-weight: ${colorNegrita};">${estado} ${botonCopiarInfo}</td>
                <td style="padding: 3px 10px; text-align: center; color: #444; font-weight: bold;">${foliosCount}</td>
                <td style="padding: 3px 10px; text-align: center; color: #666;">${porcentaje}%</td>
            </tr>`;
        });

        // --- TABLA 2: ANULADOS ---
        const totalAnulados = contadoresFinales['Anulado no Contacto'] + contadoresFinales['Anulado Usuario'];
        const valoresTotalesAnulados = `${contadoresFinales['Anulado no Contacto']}\t${contadoresFinales['Anulado Usuario']}\t${totalAnulados}`;
        const tablaAnuladosHTML = `
            <div class="card-child" style="border: 1px solid #ddd; padding: 15px; background: #fff; border-radius: 4px; flex: 1; max-width: 480px;">
                <h5 style="color: #014f8b; margin-top: 0; margin-bottom: 15px; font-family: Segoe UI, sans-serif; font-weight: bold;">Desglose de anulados por MDA:</h5>
                <table style="width: 100%; border-collapse: collapse; font-family: Calibri, sans-serif; border: 1px solid #c0c0c0;">
                    <thead><tr style="background-color: #315e9a; color: white; font-size: 13px; font-weight: bold;"><th style="padding: 5px 10px; text-align: left;">Estado</th><th style="padding: 5px 10px; text-align: center;">Sin contacto</th><th style="padding: 5px 10px; text-align: center;">Usuario</th><th style="padding: 5px 10px; text-align: center;">Total</th></tr></thead>
                    <tbody><tr style="font-size: 13px;"><td style="padding: 3px 10px;">Anulados</td><td style="padding: 3px 10px; text-align:center;">${contadoresFinales['Anulado no Contacto']}</td><td style="padding: 3px 10px; text-align:center;">${contadoresFinales['Anulado Usuario']}</td><td style="padding: 3px 10px; text-align:center; font-weight:bold;">${totalAnulados}</td></tr></tbody>
                    <tfoot><tr style="background-color: #315e9a; color: white; font-weight: bold; font-size: 13px;"><td style="padding: 5px 10px;">Total</td><td style="padding: 5px 10px; text-align:center;">${contadoresFinales['Anulado no Contacto']}</td><td style="padding: 5px 10px; text-align:center;">${contadoresFinales['Anulado Usuario']}</td><td style="padding: 5px 10px; text-align:center;">${totalAnulados} <button onclick="copiarTextoDirecto('${valoresTotalesAnulados}')" class="btn-copy-small" style="cursor:pointer;"><i class="fas fa-copy"></i></button></td></tr></tfoot>
                </table>
            </div>`;

        // --- TABLA DE RESIDENCIAS ---
        const ordenResidencias = ['CAPJ', 'CENTRO DE JUSTICIA', 'CIVILES', 'CONCEPCION', 'CORTE SUPREMA', 'FAMILIA', 'SAN MIGUEL'];
        let rowsResidenciasHTML = "";
        
        ordenResidencias.forEach(res => {
            const count = contadoresResidencias[res];
            const porcentaje = totalResidencias > 0 ? ((count / totalResidencias) * 100).toFixed(2) : "0.00";
            rowsResidenciasHTML += `
                <tr style="border-bottom: 1px solid #ddd; font-size: 13px;">
                    <td style="padding: 3px 10px;">${res}</td>
                    <td style="padding: 3px 10px; text-align: center; font-weight: bold;">${count}</td>
                    <td style="padding: 3px 10px; text-align: center;">${porcentaje}%</td>
                </tr>`;
        });

        const tablaResidenciasHTML = `
            <div class="card" style="border: 1px solid #ddd; padding: 0; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div class="pjud-header-toggle" onclick="toggleSection('mda-residencias-body', 'icon-mda-residencias')" style="background: #f8f9fa; border-top: 4px solid #17a2b8; padding: 12px 20px; display: flex; justify-content: space-between; align-items: center; cursor: pointer;">
                    <h4 style="color: #17a2b8; margin: 0; font-family: Segoe UI; font-weight: bold;"><i class="fas fa-building"></i> Atenciones de Residencias</h4>
                    <i id="icon-mda-residencias" class="fas fa-chevron-up rotate-icon"></i>
                </div>
                <div id="mda-residencias-body">
                    <div style="padding: 20px; background: #fff;">
                        <p style="margin-top: 0; margin-bottom: 20px; color: #555; font-size: 14px;">La siguiente tabla, muestra el total de requerimientos creados en cada una de las residencias y su respectivo estado de solución, el cual es el actual a la fecha de presentación de este informe.</p>
                        <div class="card-child" style="max-width: 400px; border: 1px solid #014f8b;">
                            <table id="tabla-mda-residencias" style="width: 100%; border-collapse: collapse; font-family: Calibri, sans-serif;">
                                <thead>
                                    <tr style="background-color: #014f8b; color: white;">
                                        <th style="padding: 5px 10px; border: 1px solid #014f8b; text-align: left;">ESTADOS</th>
                                        <th style="padding: 5px 10px; border: 1px solid #014f8b; text-align: center;">FOLIOS <button onclick="copiarColumnaFoliosResidencias()" class="btn-copy-small" style="margin-left: 5px;"><i class="fas fa-copy"></i></button></th>
                                        <th style="padding: 5px 10px; border: 1px solid #014f8b; text-align: center;">% <button onclick="copiarColumnaPorcentajeResidencias()" class="btn-copy-small" style="margin-left: 5px;"><i class="fas fa-copy"></i></button></th>
                                    </tr>
                                </thead>
                                <tbody>${rowsResidenciasHTML}</tbody>
                                <tfoot>
                                    <tr style="background-color: #014f8b; color: white; font-weight: bold;">
                                        <td style="padding: 5px 10px; border: 1px solid #014f8b; text-align: left;">Suma total</td>
                                        <td style="padding: 5px 10px; border: 1px solid #014f8b; text-align: center;">${totalResidencias}</td>
                                        <td style="padding: 5px 10px; border: 1px solid #014f8b; text-align: center;">100.00%</td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        `;

        // --- TABLA DE FALLAS ---
        const ordenJurisdicciones = Object.keys(reporteFallas).sort();
        window.ultimaDataFallas = { orden: ordenJurisdicciones, datos: reporteFallas };

        let rowsFallasHTML = "";
        let sumaTotalHard = 0, sumaTotalSoft = 0, sumaTotalVacia = 0;

        ordenJurisdicciones.forEach(jur => {
            const data = reporteFallas[jur];
            const totalFila = data.hardware + data.software + data.vacia;
            
            // CREA EL BOTON SI ES "En blanco" Y TIENE TICKETS
            let botonAuditarJur = "";
            if (jur === "En blanco" && totalFila > 0) {
                botonAuditarJur = ` <button onclick="copiarArray(window.ticketsSinJurisdiccion)" class="btn-copy-small" style="background-color: #ff8c00; color: white; border:none; border-radius:3px; padding:2px 6px; font-size:10px; cursor:pointer;" title="Auditar TKs sin jurisdicción"><i class="fas fa-search"></i> Auditar TK</button>`;
            }
            
            const jurFormat = jur === "En blanco" ? "En blanco" + botonAuditarJur : jur.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');

            rowsFallasHTML += `
                <tr style="border-bottom: 1px solid #ddd; font-size: 13px; text-align: center;">
                    <td style="padding: 3px 10px; text-align: left; font-weight: bold; white-space: nowrap;">${jurFormat}</td>
                    <td style="padding: 3px 10px;">${data.hardware}</td>
                    <td style="padding: 3px 10px;">${data.software}</td>
                    <td style="padding: 3px 10px;">${data.vacia}</td>
                    <td style="padding: 3px 10px; font-weight: bold;">${totalFila}</td>
                </tr>`;
                
            sumaTotalHard += data.hardware;
            sumaTotalSoft += data.software;
            sumaTotalVacia += data.vacia;
        });

        const granTotalFallas = sumaTotalHard + sumaTotalSoft + sumaTotalVacia;

        const tablaFallasHTML = `
            <div class="card" style="border: 1px solid #ddd; padding: 0; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
                <div class="pjud-header-toggle" onclick="toggleSection('mda-fallas-body', 'icon-mda-fallas')" style="background: #f8f9fa; border-top: 4px solid #014f8b; padding: 12px 20px; display: flex; justify-content: space-between; align-items: center; cursor: pointer;">
                    <h4 style="color: #014f8b; margin: 0; font-family: Segoe UI; font-weight: bold;"><i class="fas fa-tools"></i> Reporte de Fallas Divididas por Hardware y Software Según Jurisdicción</h4>
                    <i id="icon-mda-fallas" class="fas fa-chevron-up rotate-icon"></i>
                </div>
                <div id="mda-fallas-body">
                    <div style="padding: 20px; background: #fff;">
                        <p style="margin-top: 0; margin-bottom: 20px; color: #555; font-size: 14px;">
                            Este reporte ofrece un análisis comparativo de las fallas registradas en diferentes jurisdicciones, clasificadas en categorías: hardware, software y sin clasificar.
                        </p>
                        <div class="card-child" style="max-width: 850px; border: 1px solid #014f8b;">
                            <table style="width: 100%; border-collapse: collapse; font-family: Calibri, sans-serif;">
                                <thead>
                                    <tr style="background-color: #315e9a; color: white;">
                                        <th style="padding: 5px 10px; border: 1px solid #014f8b; text-align: left; white-space: nowrap;">JURISDICCION <button onclick="copiarTextoDirecto(window.ultimaDataFallas.orden.map(j => j === 'En blanco' ? 'En blanco' : j.split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ')).join('\\n'))" class="btn-copy-small" style="margin-left: 5px;"><i class="fas fa-copy"></i></button></th>
                                        <th style="padding: 5px 10px; border: 1px solid #014f8b; text-align: center;">HARDWARE <button onclick="copiarTextoDirecto(window.ultimaDataFallas.orden.map(j => window.ultimaDataFallas.datos[j].hardware).join('\\n'))" class="btn-copy-small" style="margin-left: 5px;"><i class="fas fa-copy"></i></button></th>
                                        <th style="padding: 5px 10px; border: 1px solid #014f8b; text-align: center;">SOFTWARE <button onclick="copiarTextoDirecto(window.ultimaDataFallas.orden.map(j => window.ultimaDataFallas.datos[j].software).join('\\n'))" class="btn-copy-small" style="margin-left: 5px;"><i class="fas fa-copy"></i></button></th>
                                        <th style="padding: 5px 10px; border: 1px solid #014f8b; text-align: center;">SIN CLASIFICAR <button onclick="copiarTextoDirecto(window.ultimaDataFallas.orden.map(j => window.ultimaDataFallas.datos[j].vacia).join('\\n'))" class="btn-copy-small" style="margin-left: 5px;"><i class="fas fa-copy"></i></button></th>
                                        <th style="padding: 5px 10px; border: 1px solid #014f8b; text-align: center;">SUMA TOTAL <button onclick="copiarTextoDirecto(window.ultimaDataFallas.orden.map(j => window.ultimaDataFallas.datos[j].hardware + window.ultimaDataFallas.datos[j].software + window.ultimaDataFallas.datos[j].vacia).join('\\n'))" class="btn-copy-small" style="margin-left: 5px;"><i class="fas fa-copy"></i></button></th>
                                    </tr>
                                </thead>
                                <tbody>${rowsFallasHTML}</tbody>
                                <tfoot>
                                    <tr style="background-color: #315e9a; color: white; font-weight: bold; text-align: center;">
                                        <td style="padding: 5px 10px; border: 1px solid #014f8b; text-align: left;">Suma total</td>
                                        <td style="padding: 5px 10px; border: 1px solid #014f8b;">${sumaTotalHard}</td>
                                        <td style="padding: 5px 10px; border: 1px solid #014f8b;">${sumaTotalSoft}</td>
                                        <td style="padding: 5px 10px; border: 1px solid #014f8b;">
                                            ${sumaTotalVacia}
                                            ${sumaTotalVacia > 0 ? `<button onclick="copiarArray(window.ticketsSinFalla)" class="btn-copy-small" style="background-color: #ff8c00; color: white; border:none; border-radius:3px; padding:2px 6px; font-size:10px; cursor:pointer; margin-left: 5px;" title="Copiar TKs sin clasificar"><i class="fas fa-search"></i> TK</button>` : ''}
                                        </td>
                                        <td style="padding: 5px 10px; border: 1px solid #014f8b;">${granTotalFallas}</td>
                                    </tr>
                                </tfoot>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        `;

// --- CONSTRUCCIÓN DEL HTML BASE FINAL (ORDEN PERSONALIZADO CORREGIDO) ---
resultContainer.innerHTML = `
    <div class="card" style="border: 1px solid #ddd; padding: 0; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
        <div class="pjud-header-toggle" onclick="toggleSection('mda-atencion-body', 'icon-mda-atencion')" style="background: #f8f9fa; border-top: 4px solid #014f8b; padding: 12px 20px; display: flex; justify-content: space-between; align-items: center; cursor: pointer;">
            <h4 style="color: #014f8b; margin: 0; font-family: Segoe UI; font-weight: bold;"><i class="fas fa-list-alt"></i> Atención MDA</h4>
            <i id="icon-mda-atencion" class="fas fa-chevron-up rotate-icon"></i>
        </div>
        <div id="mda-atencion-body">
            <div style="padding: 20px; background: #fff; display: flex; gap: 20px; flex-wrap: wrap;">
                <div class="card-child" style="flex: 1; max-width: 480px;">
                    <table style="width: 100%; border-collapse: collapse; font-family: Calibri; border: 1px solid #c0c0c0;">
                        <thead style="background:#315e9a; color:white;">
                            <tr>
                                <th style="padding: 5px 10px; text-align: left;">ESTADOS</th>
                                <th style="padding: 5px 10px; text-align:center;">
                                    FOLIOS <button onclick="copiarColumnaFoliosMDA()" class="btn-copy-small" title="Copiar Folios"><i class="fas fa-copy"></i></button>
                                </th>
                                <th style="padding: 5px 10px; text-align:center;">
                                    % <button onclick="copiarColumnaPorcentajesMDA()" class="btn-copy-small" title="Copiar Porcentajes"><i class="fas fa-copy"></i></button>
                                </th>
                            </tr>
                        </thead>
                        <tbody>${rowsHTML}</tbody>
                        <tfoot style="background:#315e9a; color:white; font-weight:bold;">
                            <tr>
                                <td style="padding: 5px 10px;">Total general</td>
                                <td style="padding: 5px 10px; text-align:center;">${totalTicketsMesFiltro}</td>
                                <td style="padding: 5px 10px; text-align:center;">100.00%</td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
                ${tablaAnuladosHTML}
            </div>
        </div>
    </div>

    ${typeof bloqueAcumuladosHP !== 'undefined' ? bloqueAcumuladosHP : ''}

    <div class="card" style="border: 1px solid #ddd; padding: 0; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
        <div class="pjud-header-toggle" onclick="toggleSection('mda-diario-body', 'icon-mda-diario')" style="background: #f8f9fa; border-top: 4px solid #28a745; padding: 12px 20px; display: flex; justify-content: space-between; align-items: center; cursor: pointer;">
            <h4 style="color: #2e7d32; margin: 0; font-family: Segoe UI; font-weight: bold;"><i class="fas fa-calendar-alt"></i> Ingreso Diario Completo</h4>
            <i id="icon-mda-diario" class="fas fa-chevron-up rotate-icon"></i>
        </div>
        <div id="mda-diario-body">
            <div style="padding: 20px; background: #fff;">
                <table style="width: 100%; max-width: 650px; border-collapse: collapse; font-family: Calibri; border: 1px solid #c0c0c0;">
                    <thead>
                        <tr style="background-color: #2e7d32; color: white; font-size: 13px;">
                            <th style="padding: 4px 8px; text-align: left;">Fecha <button onclick="copiarTextoDirecto('${colFechas.join('\\n')}')" class="btn-copy-small"><i class="fas fa-copy"></i></button></th>
                            <th style="padding: 4px 8px; text-align:center;">Tickets MDA <button onclick="copiarTextoDirecto('${colMDA.join('\\n')}')" class="btn-copy-small"><i class="fas fa-copy"></i></button></th>
                            <th style="padding: 4px 8px; text-align:center;">Residencias <button onclick="copiarTextoDirecto('${colResi.join('\\n')}')" class="btn-copy-small"><i class="fas fa-copy"></i></button></th>
                            <th style="padding: 4px 8px; text-align:center;">Total Día <button onclick="copiarTextoDirecto('${colTotal.join('\\n')}')" class="btn-copy-small"><i class="fas fa-copy"></i></button></th>
                        </tr>
                    </thead>
                    <tbody>${dailyRowsHTML}</tbody>
                    <tfoot style="background-color: #e8f5e9; font-weight: bold; font-size: 13px; border-top: 2px solid #2e7d32;">
                        <tr>
                            <td style="padding: 4px 8px;">TOTALES</td>
                            <td style="padding: 4px 8px; text-align:center; color:#014f8b;">${colMDA.reduce((a, b) => a + b, 0)}</td>
                            <td style="padding: 4px 8px; text-align:center; color:#2e7d32;">${colResi.reduce((a, b) => a + b, 0)}</td>
                            <td style="padding: 4px 8px; text-align:center;">${totalTicketsMesFiltro}</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        </div>
    </div>

    <div class="card" style="border: 1px solid #ddd; padding: 0; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
        <div class="pjud-header-toggle" onclick="toggleSection('mda-evolucion-body', 'icon-mda-evolucion')" style="background: #f8f9fa; border-top: 4px solid #28a745; padding: 12px 20px; display: flex; justify-content: space-between; align-items: center; cursor: pointer;">
            <h4 style="color: #2e7d32; margin: 0; font-family: Segoe UI; font-weight: bold;"><i class="fas fa-chart-line"></i> Evolución de Ingreso Diario</h4>
            <i id="icon-mda-evolucion" class="fas fa-chevron-up rotate-icon"></i>
        </div>
        <div id="mda-evolucion-body">
            <div style="padding: 20px; background: #fff;">
                <div style="width: 100%; height: 380px;">
                    <canvas id="canvasGraficoEvolucion"></canvas>
                </div>
            </div>
        </div>
    </div>

    ${tablaResidenciasHTML}

    ${typeof bloqueCanalAtencion !== 'undefined' ? bloqueCanalAtencion : ''}

    <div id="mda-agentes-container"></div>

    ${typeof bloqueFallasEquipos !== 'undefined' ? bloqueFallasEquipos : ''}

    ${tablaFallasHTML}

    <div id="conclusiones-container"></div>
`;

        const drawNumbersPlugin = {
            id: 'drawNumbers',
            afterDatasetsDraw(chart) {
                const { ctx, data } = chart;
                ctx.save();
                ctx.font = "bold 11px 'Segoe UI', Arial, sans-serif";
                ctx.textAlign = 'center';
                
                const meta0 = chart.getDatasetMeta(0);
                if (meta0 && !meta0.hidden) {
                    meta0.data.forEach((point, index) => {
                        const val = data.datasets[0].data[index];
                        if (val > 0) {
                            ctx.fillStyle = '#014f8b';
                            ctx.textBaseline = 'bottom';
                            ctx.fillText(val, point.x, point.y - 6);
                        }
                    });
                }

                const meta1 = chart.getDatasetMeta(1);
                if (meta1 && !meta1.hidden) {
                    meta1.data.forEach((point, index) => {
                        const val = data.datasets[1].data[index];
                        if (val > 0) {
                            ctx.fillStyle = '#dc3545';
                            ctx.textBaseline = 'top';
                            ctx.fillText(val, point.x, point.y + 6);
                        }
                    });
                }
                ctx.restore();
            }
        };

        try {
            const ctxL = document.getElementById('canvasGraficoEvolucion').getContext('2d');
            const modernFont = "'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif";
            
            new Chart(ctxL, {
                type: 'line',
                data: {
                    labels: labelsGrafico,
                    datasets: [
                        { label: 'MDA', data: colMDA, borderColor: '#014f8b', backgroundColor: '#014f8b', borderWidth: 2, tension: 0.3, pointRadius: 4 },
                        { label: 'Residencias', data: colResi, borderColor: '#dc3545', backgroundColor: '#dc3545', borderWidth: 2, tension: 0.3, pointRadius: 4 }
                    ]
                },
                plugins: [drawNumbersPlugin], 
                options: {
                    responsive: true, 
                    maintainAspectRatio: false,
                    plugins: { 
                        title: { display: true, text: 'Requerimientos creados en el Mes', font: { size: 16, family: modernFont, weight: 'normal' }, padding: 15 },
                        legend: { position: 'bottom', labels: { font: { family: modernFont } } }
                    },
                    scales: { 
                        x: { ticks: { font: { size: 12, family: modernFont }, maxRotation: 60, minRotation: 45 }, grid: { display: false } },
                        y: { beginAtZero: true, grid: { color: '#f0f0f0' } }
                    }
                }
            });

        } catch (error) {
            console.error("Error cargando el gráfico:", error);
        }

        resultContainer.classList.remove('hidden'); 
        showToast(`✅ Informe generado con éxito.`);


     // --- 1. PRIMERO DETECTAMOS EL NOMBRE DEL MES ---
        let nombreMesDetectado = "MES NUEVO";
        if (rows.length > 1 && rows[1][5]) {
            nombreMesDetectado = extraerMesDeCSV(rows[1][5].split(' ')[0]);
        }

        // --- 2. PREPARAMOS LOS DATOS PARA LA CONCLUSIÓN ---
        const infoResParaConclusion = {
            total: totalResidencias,
            capj: contadoresResidencias['CAPJ'] || 0,
            cj: contadoresResidencias['CENTRO DE JUSTICIA'] || 0
        };

        const infoMDAParaConclusion = {
            solucionados: contadoresFinales['Solucionados en MDA'] || 0,
            mesaHP: (contadoresFinales['Solucionados en MDA'] || 0) + 
                    (contadoresFinales['Derivado a Residencia'] || 0) + 
                    (contadoresFinales['Derivado a SCO'] || 0)
        };

        // --- 3. LLAMADA QUE GENERA EL TEXTO DE CONCLUSIONES ---
        generarConclusionesDinamicas(nombreMesDetectado, infoMDAParaConclusion, infoResParaConclusion, totalTicketsMesFiltro);

        // --- 4. PROCESO DE INYECCIÓN DE DATOS DEL CSV A LA TABLA ---
        try {
            // Llenamos la variable global del histórico
            window.DATOS_NUEVO_MES_HISTORICO = {
                mes: nombreMesDetectado,
                anulado: (contadoresFinales['Anulado no Contacto'] || 0) + (contadoresFinales['Anulado Usuario'] || 0),
                derivadaOtraArea: contadoresFinales['Derivado a Otra Área'] || 0,
                derivadaResidencia: contadoresFinales['Derivado a Residencia'] || 0,
                derivadaSCO: contadoresFinales['Derivado a SCO'] || 0,
                habilitacionSCO: contadoresFinales['Habilitacion x SCO'] || 0,
                solucionadoMDA: contadoresFinales['Solucionados en MDA'] || 0,
                dimensionado: 1300,
                mesaHP: infoMDAParaConclusion.mesaHP
            };

            // Busca esta parte y déjala así:
            DATOS_MANUAL_ACTUAL = { 
                mes: nombreMesDetectado, 
                equipos: conteoEquiposParaTabla // <--- Antes estaba como {}
            };

            console.log("✅ Datos e Informe generados para:", nombreMesDetectado);
            
            // Refrescamos visualmente todo
            renderizarHistorico();
            renderizarTablaTipos(DATOS_HISTORICOS_MESES);
            renderizarGraficoCanales(); // <--- NUEVO: Esta línea actualiza el gráfico de líneas (MDA vs Residentes)

        } catch (err) {
            console.error("Error al actualizar la interfaz:", err);
        }

        resultContainer.classList.remove('hidden'); 
        showToast(`✅ Informe y Conclusiones generados con éxito.`);

    }; // Cierre del reader.onload
    reader.readAsText(file, "ISO-8859-1"); 
}
// ... (Resto de funciones copiarArray, copiarColumnaFoliosMDA y renderizarHistorico se mantienen igual)

// Función multiuso para copiar cualquier lista
function copiarArray(arrayTickets) {
    if (!arrayTickets || arrayTickets.length === 0) {
        return showToast("⚠️ No hay tickets para copiar.");
    }
    const textoCopiar = arrayTickets.join('\n');
    copiarAlPortapapeles(textoCopiar, `✅ ${arrayTickets.length} tickets copiados`);
}

// Función específica para copiar solo la columna de números (Folios) de la tabla principal
function copiarColumnaFoliosMDA() {
    const tabla = document.querySelector("#mda-atencion-body table");
    if (!tabla) return showToast("⚠️ No se encontró la tabla de Atención MDA");

    const filas = tabla.querySelectorAll("tbody tr");
    let texto = "";
    
    filas.forEach(f => {
        if(f.cells[1]) texto += f.cells[1].innerText.trim() + "\n";
    });
    
    const total = tabla.querySelector("tfoot tr td:nth-child(2)");
    if (total) texto += total.innerText.trim();

    copiarTextoDirecto(texto);
    showToast("✅ Folios MDA copiados");
}

// Función para copiar PORCENTAJES de la primera tabla
function copiarColumnaPorcentajesMDA() {
    const tabla = document.querySelector("#mda-atencion-body table");
    if (!tabla) return showToast("⚠️ No se encontró la tabla");

    const filas = tabla.querySelectorAll("tbody tr");
    let texto = "";
    
    filas.forEach(f => {
        if(f.cells[2]) {
            // Reemplaza punto por coma para Excel Chile
            let valor = f.cells[2].innerText.trim().replace('.', ',');
            texto += valor + "\n";
        }
    });
    
    texto += "100,00%";
    copiarTextoDirecto(texto);
    showToast("✅ Porcentajes MDA copiados");
}
// Función para copiar la columna de Folios de Residencias
function copiarColumnaFoliosResidencias() {
    const tabla = document.getElementById('tabla-mda-residencias');
    if (!tabla) return showToast("⚠️ No hay tabla generada.");

    let folios = [];
    const filasBody = tabla.querySelectorAll('tbody tr');
    const filaFoot = tabla.querySelector('tfoot tr');
    
    filasBody.forEach(fila => {
        const celdaFolio = fila.children[1]; 
        if (celdaFolio) folios.push(celdaFolio.innerText.trim());
    });

    if (filaFoot) {
        const celdaTotal = filaFoot.children[1];
        if (celdaTotal) folios.push(celdaTotal.innerText.trim());
    }

    if (folios.length === 0) return showToast("⚠️ No hay datos para copiar.");
    
    const textoCopiar = folios.join('\n');
    copiarAlPortapapeles(textoCopiar, `✅ ${folios.length} valores copiados`);
}

// Función para copiar la columna de Porcentajes de Residencias
function copiarColumnaPorcentajeResidencias() {
    const tabla = document.getElementById('tabla-mda-residencias');
    if (!tabla) return showToast("⚠️ No hay tabla generada.");

    let porcentajes = [];
    const filasBody = tabla.querySelectorAll('tbody tr');
    const filaFoot = tabla.querySelector('tfoot tr');
    
    filasBody.forEach(fila => {
        const celdaPorcentaje = fila.children[2]; 
        if (celdaPorcentaje) porcentajes.push(celdaPorcentaje.innerText.trim());
    });

    if (filaFoot) {
        const celdaTotal = filaFoot.children[2];
        if (celdaTotal) porcentajes.push(celdaTotal.innerText.trim());
    }

    if (porcentajes.length === 0) return showToast("⚠️ No hay datos para copiar.");
    
    const textoCopiar = porcentajes.join('\n');
    copiarAlPortapapeles(textoCopiar, `✅ ${porcentajes.length} valores copiados`);
}

// =========================================================
// === MÓDULO: REGISTRO HISTÓRICO Y GRÁFICOS (PJUD 5) ===
// =========================================================

// 1. Variable global
let HISTORICO_PJUD5 = []; 

// 2. ENLACE DE TU GOOGLE SHEETS
const URL_SHEETS_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRF7xX_Lf5Xjg6dYfV9y5e6arcmSP0VKmz6DGZAbQw8laxYRsC7V2FiifeorYpoiNd6S-h81aTvNwjq/pub?gid=2093520733&single=true&output=csv"; 
let DATOS_TIPOS_EQUIPO = [];
const URL_TIPOS_JSONP = "https://docs.google.com/spreadsheets/d/1kFRvMmyHom4APfhVqmdvHg_wSkKPH8Y0eJ5Z6HnVeq0/gviz/tq?tqx=responseHandler:manejarRespuestaTipos&gid=1959700363";

// 3. FUNCIÓN MOTOR (CON PROXY PARA EVITAR ERROR CORS)
function cargarDatosDesdeSheets() {
    const docId = '1kFRvMmyHom4APfhVqmdvHg_wSkKPH8Y0eJ5Z6HnVeq0';
    // Usamos el GID dinámico según el proyecto activo
    const gid = CONFIG_SHEETS[PROYECTO_ACTIVO].gidHistorico;
    
    const url = `https://docs.google.com/spreadsheets/d/${docId}/gviz/tq?tqx=responseHandler:manejarRespuestaSheets&gid=${gid}`;
    console.log(`Conectando a Sheets: Historico ${PROYECTO_ACTIVO} (GID: ${gid})`);
    
    const script = document.createElement('script');
    script.src = url;
    document.body.appendChild(script);
}

// Esta función recibe los datos mágicamente desde Google
window.manejarRespuestaSheets = function(response) {
    try {
        const filas = response.table.rows;
        
        // --- NUEVA LÓGICA DE FORMATEO DE FECHAS ---
        const mesesNombres = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sept", "oct", "nov", "dic"];
        
        const formatearFechaGoogle = (valorFecha) => {
            if (!valorFecha || typeof valorFecha !== 'string') return "S/N";
            
            // Si viene como "Date(2026,0,25)"
            if (valorFecha.includes("Date(")) {
                const coincidencia = valorFecha.match(/\(([^)]+)\)/);
                if (coincidencia) {
                    const partes = coincidencia[1].split(',');
                    const año = partes[0].trim().slice(-2); 
                    const mesIdx = parseInt(partes[1].trim()); 
                    
                    return `${mesesNombres[mesIdx]} ${año}`;
                }
            }
            return valorFecha.replace(/"/g, '').trim();
        };

        // --- MAPEO DE DATOS Y RECORTE A 12 MESES ---
        HISTORICO_PJUD5 = filas.map(row => {
            const c = row.c;
            const getV = (idx) => (c[idx] && c[idx].v !== null) ? c[idx].v : 0;
            const getF = (idx) => (c[idx] && c[idx].f !== null) ? c[idx].f : null;
            
            let fechaRaw = getF(0) || getV(0);
            let fechaLimpia = formatearFechaGoogle(String(fechaRaw));

            return {
                mes: fechaLimpia,
                anulado: parseInt(getV(1)) || 0,
                derivadaOtraArea: parseInt(getV(2)) || 0,
                derivadaResidencia: parseInt(getV(3)) || 0,
                derivadaSCO: parseInt(getV(4)) || 0,
                habilitacionSCO: parseInt(getV(5)) || 0,
                solucionadoMDA: parseInt(getV(6)) || 0,
                dimensionado: parseInt(getV(8)) || 0,
                mesaHP: parseInt(getV(9)) || 0
            };
        }).slice(-12); // <--- ESTE ES EL CAMBIO: Mantiene solo los últimos 12 meses

        console.log("✅ Datos procesados (limitados a últimos 12 meses):", HISTORICO_PJUD5);
        
        // --- DIBUJAMOS TODO ---
        if (typeof renderizarHistorico === "function") renderizarHistorico();
        if (typeof renderizarGraficoCanales === "function") renderizarGraficoCanales();
        
    } catch (error) {
        console.error("❌ Error procesando la respuesta:", error);
    }
};
// ... (Aquí termina tu función manejarRespuestaSheets)

document.addEventListener('DOMContentLoaded', () => {
    // 1. Cargamos la base de datos de Guías (Sincronización con LAB)
    cargarDatosLaboratorio(); 
    
    // 2. Cargamos los datos del Informe MDA
    cargarDatosDesdeSheets();
    cargarDatosTipos();
    
    // 3. Cargamos la tabla de links
    cargarBibliotecaEnlaces();
    
    console.log("🚀 Sistema FCOM Inicializado.");
});

// 4. Iniciar carga al cargar la web
document.addEventListener('DOMContentLoaded', cargarDatosDesdeSheets);

let chartAcumulado = null;
let chartGestiones = null;

// 5. Función de Tabla y Gráficos de Barras
function renderizarHistorico() {
    const tableContainer = document.getElementById('tabla-historica-mda');
    if (!tableContainer || HISTORICO_PJUD5.length === 0) return;

    // --- 1. UNIÓN DINÁMICA DE DATOS ---
    let datosCompletos = JSON.parse(JSON.stringify(HISTORICO_PJUD5));
    if (window.DATOS_NUEVO_MES_HISTORICO) {
        const existeIdx = datosCompletos.findIndex(d => d.mes === window.DATOS_NUEVO_MES_HISTORICO.mes);
        if (existeIdx !== -1) {
            datosCompletos[existeIdx] = window.DATOS_NUEVO_MES_HISTORICO;
        } else {
            datosCompletos.push(window.DATOS_NUEVO_MES_HISTORICO);
        }
    }

    let totAnulado = 0, totOtraArea = 0, totResidencia = 0, totSCO = 0, totHab = 0, totSol = 0, totMesaHP = 0, totGlobal = 0;

    // --- 2. PLUGIN PARA TOTALES EN GRÁFICOS ---
    const pluginTotalesBarras = {
        id: 'pluginTotalesBarras',
        afterDatasetsDraw(chart) {
            const { ctx, data } = chart;
            ctx.save();
            ctx.font = "bold 11px Arial";
            ctx.textAlign = 'center';
            ctx.fillStyle = '#002b5c';
            chart.data.datasets.forEach((dataset, i) => {
                if (chart.getDatasetMeta(i).type === 'bar') {
                    const meta = chart.getDatasetMeta(i);
                    meta.data.forEach((bar, index) => {
                        const val = dataset.data[index];
                        if (val > 0) ctx.fillText(val.toLocaleString('es-CL'), bar.x, bar.y - 5);
                    });
                }
            });
            ctx.restore();
        }
    };

    // --- 3. CONSTRUCCIÓN DE CABECERA UNIFICADA (AZUL) ---
    const azulPJUD = "#2a579a";
    let headHTML = `<tr style="background-color: ${azulPJUD};">`; 
    
    // Celda ESTADO
    headHTML += `<th style="padding: 4px 8px; border: 1px solid #1a3b6c; text-align: center; background-color: ${azulPJUD}; color: white !important;">ESTADO</th>`;
    
    datosCompletos.forEach((d, i) => {
        let botonCopiar = "";
        if (i === datosCompletos.length - 1) {
            const totalMes = d.anulado + d.derivadaOtraArea + d.derivadaResidencia + d.derivadaSCO + d.habilitacionSCO + d.solucionadoMDA;
            const dataStr = [d.mes, d.anulado, d.derivadaOtraArea, d.derivadaResidencia, d.derivadaSCO, d.habilitacionSCO, d.solucionadoMDA, totalMes, d.dimensionado, d.mesaHP].join('\t');
            botonCopiar = `<br><button onclick="copiarTextoDirecto('${dataStr}')" style="margin-top: 2px; cursor: pointer; color: ${azulPJUD}; background: #fff; border:none; border-radius:3px; padding: 1px 4px; font-size: 9px;"><i class="fas fa-copy"></i> Copiar</button>`;
        }
        
        // Celdas de Meses
        headHTML += `<th style="padding: 4px 8px; border: 1px solid #1a3b6c; text-align: center; vertical-align: middle; background-color: ${azulPJUD}; color: white !important;">${d.mes.toUpperCase()}${botonCopiar}</th>`;
        
        totAnulado += d.anulado; totOtraArea += d.derivadaOtraArea; totResidencia += d.derivadaResidencia;
        totSCO += d.derivadaSCO; totHab += d.habilitacionSCO; totSol += d.solucionadoMDA; totMesaHP += d.mesaHP;
    });

    // Celda TOTAL ACUMULADO
    headHTML += `<th style="padding: 4px 8px; border: 1px solid #1a3b6c; text-align: center; background-color: ${azulPJUD}; color: white !important;">TOTAL ACUM.</th></tr>`;

    // --- 4. FILAS DE DATOS ---
    const buildRow = (label, key, totalVal) => {
        let rowHTML = `<tr><td style="padding: 3px 8px; border: 1px solid #ccc; text-align: left; font-weight: bold; background: #f8f9fa;">${label}</td>`;
        datosCompletos.forEach(d => { rowHTML += `<td style="padding: 3px 8px; border: 1px solid #ccc; text-align: center;">${d[key].toLocaleString('es-CL')}</td>`; });
        rowHTML += `<td style="padding: 3px 8px; border: 1px solid #ccc; font-weight: bold; text-align: center; background-color: #eee;">${totalVal.toLocaleString('es-CL')}</td></tr>`;
        return rowHTML;
    };

    let bodyHTML = buildRow('Anulado', 'anulado', totAnulado) +
                   buildRow('Derivado a Otra Área', 'derivadaOtraArea', totOtraArea) +
                   buildRow('Derivado a Residencia', 'derivadaResidencia', totResidencia) +
                   buildRow('Derivado a SCO', 'derivadaSCO', totSCO) +
                   buildRow('Habilitación x SCO', 'habilitacionSCO', totHab) +
                   buildRow('Solucionado en MDA', 'solucionadoMDA', totSol);

    let rowTotales = `<tr style="background-color: #e2eaf4; font-weight: bold;"><td style="padding: 4px 8px; border: 1px solid #ccc;">Total Mensual</td>`;
    datosCompletos.forEach(d => {
        let mesT = d.anulado + d.derivadaOtraArea + d.derivadaResidencia + d.derivadaSCO + d.habilitacionSCO + d.solucionadoMDA;
        totGlobal += mesT;
        rowTotales += `<td style="padding: 4px 8px; border: 1px solid #ccc; text-align: center;">${mesT.toLocaleString('es-CL')}</td>`;
    });
    rowTotales += `<td style="padding: 4px 8px; border: 1px solid #ccc; text-align: center; background: ${azulPJUD}; color:white;">${totGlobal.toLocaleString('es-CL')}</td></tr>`;

    // --- 5. FILAS DE GESTIÓN ---
    let rowDim = `<tr><td style="padding: 3px 8px; border: 1px solid #ccc; text-align: left; font-weight: bold;">Dimensionado</td>`;
    let rowMesa = `<tr><td style="padding: 3px 8px; border: 1px solid #ccc; text-align: left; font-weight: bold;">Mesa HP</td>`;
    let rowRes = `<tr><td style="padding: 3px 8px; border: 1px solid #ccc; text-align: left; font-weight: bold;">% Resolución mesa HP</td>`;

    datosCompletos.forEach(d => {
        let mesT = d.anulado + d.derivadaOtraArea + d.derivadaResidencia + d.derivadaSCO + d.habilitacionSCO + d.solucionadoMDA;
        let porcentaje = ((d.mesaHP / (mesT || 1)) * 100).toFixed(1).replace('.', ',');
        rowDim += `<td style="padding: 3px 8px; border: 1px solid #ccc; text-align: center;">${d.dimensionado.toLocaleString('es-CL')}</td>`;
        rowMesa += `<td style="padding: 3px 8px; border: 1px solid #ccc; text-align: center; font-weight: bold; color: #014f8b;">${d.mesaHP.toLocaleString('es-CL')}</td>`;
        rowRes += `<td style="padding: 3px 8px; border: 1px solid #ccc; text-align: center; font-weight: bold; color: #166534;">${porcentaje}%</td>`;
    });
    rowDim += `<td style="padding: 3px 8px; border: 1px solid #ccc; background: #eee; text-align:center;">-</td></tr>`;
    rowMesa += `<td style="padding: 3px 8px; border: 1px solid #ccc; font-weight: bold; text-align: center; background: #eee;">${totMesaHP.toLocaleString('es-CL')}</td></tr>`;
    rowRes += `<td style="padding: 3px 8px; border: 1px solid #ccc; font-weight: bold; text-align: center; background: #eee;">${((totMesaHP / (totGlobal || 1)) * 100).toFixed(1).replace('.', ',')}%</td></tr>`;

    tableContainer.innerHTML = `<thead>${headHTML}</thead><tbody>${bodyHTML}${rowTotales}${rowDim}${rowMesa}${rowRes}</tbody>`;

    // --- 6. GRÁFICOS ---
    const labelsGraficos = datosCompletos.map(d => d.mes);
    if (window.chartAcumulado instanceof Chart) window.chartAcumulado.destroy();
    const ctxA = document.getElementById('chartAcumulado');
    if(ctxA) {
        window.chartAcumulado = new Chart(ctxA, {
            type: 'bar', data: { labels: labelsGraficos, datasets: [
                { label: 'Mesa HP', data: datosCompletos.map(d => d.mesaHP), backgroundColor: '#8ea4d2' },
                { label: 'Derivado a Otra Área', data: datosCompletos.map(d => d.derivadaOtraArea), backgroundColor: '#4b74b1' },
                { label: 'Anulado', data: datosCompletos.map(d => d.anulado), backgroundColor: '#1f4287' },
                { label: '% Resolución', data: datosCompletos.map(d => {
                    let t = d.anulado + d.derivadaOtraArea + d.derivadaResidencia + d.derivadaSCO + d.habilitacionSCO + d.solucionadoMDA;
                    return ((d.mesaHP / (t || 1)) * 100).toFixed(1);
                }), type: 'line', borderColor: '#b0c4de', yAxisID: 'y1', pointRadius: 3, tension: 0.4 }
            ]},
            plugins: [pluginTotalesBarras],
            options: { responsive: true, maintainAspectRatio: false, scales: { y: { beginAtZero: true }, y1: { type: 'linear', position: 'right', min: 0, max: 100 } } }
        });
    }

    if (window.chartGestiones instanceof Chart) window.chartGestiones.destroy();
    const ctxG = document.getElementById('chartGestiones');
    if(ctxG) {
        window.chartGestiones = new Chart(ctxG, {
            type: 'bar', data: { labels: labelsGraficos, datasets: [
                { label: 'Total General', data: datosCompletos.map(d => d.anulado + d.derivadaOtraArea + d.derivadaResidencia + d.derivadaSCO + d.habilitacionSCO + d.solucionadoMDA), backgroundColor: '#4b85c5' },
                { label: 'Dimensionado', data: datosCompletos.map(d => d.dimensionado), type: 'line', borderColor: '#c00000', pointRadius: 0, fill: false }
            ]},
            plugins: [pluginTotalesBarras],
            options: { responsive: true, maintainAspectRatio: false }
        });
    }
}
// 6. Función Gráfico Comparativa Canales
function renderizarGraficoCanales() {
    const ctx = document.getElementById('chartCanalAtencion');
    const tableDiv = document.getElementById('tabla-inferior-canales');
    if (!ctx || !tableDiv || HISTORICO_PJUD5.length === 0) return;

    if (window.miGraficoCanales) window.miGraficoCanales.destroy();

    // --- 1. UNIÓN DE DATOS ---
    let datosCompletos = JSON.parse(JSON.stringify(HISTORICO_PJUD5));
    if (window.DATOS_NUEVO_MES_HISTORICO) {
        const existeIdx = datosCompletos.findIndex(d => d.mes === window.DATOS_NUEVO_MES_HISTORICO.mes);
        if (existeIdx !== -1) {
            datosCompletos[existeIdx] = window.DATOS_NUEVO_MES_HISTORICO;
        } else {
            datosCompletos.push(window.DATOS_NUEVO_MES_HISTORICO);
        }
    }

    // --- 2. PREPARACIÓN DE LABELS Y DATA ---
    const labelsMeses = datosCompletos.map(d => d.mes);
    
    // MDA: Suma de categorías remotas
    const dataMDA = datosCompletos.map(d => 
        (d.anulado || 0) + (d.derivadaOtraArea || 0) + (d.derivadaSCO || 0) + (d.habilitacionSCO || 0) + (d.solucionadoMDA || 0)
    );
    
    // Residentes: Valor directo
    const dataResidentes = datosCompletos.map(d => d.derivadaResidencia || 0);

    // --- 3. PLUGIN DE VALORES ---
    const pluginValoresYPuntos = {
        id: 'pluginValoresYPuntos',
        afterDatasetsDraw(chart) {
            const { ctx } = chart;
            ctx.save();
            ctx.font = "bold 12px 'Segoe UI', Arial, sans-serif";
            ctx.textAlign = 'center';
            ctx.textBaseline = 'bottom';
            const colorTextoGlobal = '#002b5c';

            chart.data.datasets.forEach((dataset, i) => {
                const meta = chart.getDatasetMeta(i);
                const colorDeLaLinea = dataset.borderColor;
                if (!meta.hidden) {
                    meta.data.forEach((point, index) => {
                        const val = dataset.data[index];
                        if (val !== undefined) {
                            ctx.fillStyle = colorTextoGlobal;
                            ctx.fillText(val, point.x, point.y - 14);
                            ctx.beginPath();
                            ctx.arc(point.x, point.y, 4.5, 0, 2 * Math.PI); 
                            ctx.fillStyle = colorDeLaLinea;
                            ctx.fill();
                            ctx.strokeStyle = 'white';
                            ctx.lineWidth = 1;
                            ctx.stroke();
                            ctx.closePath();
                        }
                    });
                }
            });
            ctx.restore();
        }
    };

    // --- 4. CREACIÓN DEL CHART ---
    window.miGraficoCanales = new Chart(ctx.getContext('2d'), {
        type: 'line',
        data: {
            labels: labelsMeses,
            datasets: [
                {
                    label: 'MDA',
                    data: dataMDA,
                    borderColor: '#4b74b1',
                    tension: 0.35,
                    fill: false,
                    pointRadius: 0, 
                    pointHitRadius: 10
                },
                {
                    label: 'Residentes',
                    data: dataResidentes,
                    borderColor: '#b0c4de',
                    tension: 0.35,
                    fill: false,
                    pointRadius: 0,
                    pointHitRadius: 10
                }
            ]
        },
        plugins: [pluginValoresYPuntos],
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { 
                legend: { display: false }, 
                title: { display: true, text: 'Comparativa histórica por canal de atención', font: { size: 16, weight: 'bold' } } 
            },
            scales: { 
                y: { beginAtZero: true, max: 1250 }, 
                x: { ticks: { display: true } } // Aseguramos que se vean los nombres de los meses
            }
        }
    });

    // --- 5. ACTUALIZACIÓN DE LA TABLA INFERIOR DEL GRÁFICO ---
    // --- ACTUALIZACIÓN DE LA TABLA INFERIOR (FILAS MÁS CHICAS) ---
    let tablaHTML = `<table style="width: 100%; border-collapse: collapse; font-family: 'Segoe UI'; font-size: 11px; border: 1px solid #ccc; background: white;">
        <thead>
            <tr style="background: #f8f9fa;">
                <th style="border: 1px solid #ccc; width: 120px; padding: 4px 8px;">CANAL</th>
                ${labelsMeses.map(m => `<th style="border: 1px solid #ccc; text-align:center; padding: 4px 2px;">${m}</th>`).join('')}
            </tr>
        </thead>
        <tbody>
            <tr>
                <td style="border: 1px solid #ccc; padding: 4px 8px; color:#4b74b1; font-weight:bold;"><i class="fas fa-headset"></i> MDA</td>
                ${dataMDA.map(v => `<td style="border: 1px solid #ccc; text-align:center; padding: 4px 2px;">${v}</td>`).join('')}
            </tr>
            <tr>
                <td style="border: 1px solid #ccc; padding: 4px 8px; color:#8ba0ba; font-weight:bold;"><i class="fas fa-user-tie"></i> Residentes</td>
                ${dataResidentes.map(v => `<td style="border: 1px solid #ccc; text-align:center; padding: 4px 2px;">${v}</td>`).join('')}
            </tr>
        </tbody>
    </table>`;
    
    tableDiv.innerHTML = tablaHTML;
}

function cargarDatosTipos() {
    const docId = '1kFRvMmyHom4APfhVqmdvHg_wSkKPH8Y0eJ5Z6HnVeq0';
    // Usamos el GID dinámico según el proyecto activo
    const gid = CONFIG_SHEETS[PROYECTO_ACTIVO].gidEquipos;
    
    const url = `https://docs.google.com/spreadsheets/d/${docId}/gviz/tq?tqx=responseHandler:manejarRespuestaTipos&gid=${gid}`;
    console.log(`Conectando a Sheets: Equipos ${PROYECTO_ACTIVO} (GID: ${gid})`);
    
    const script = document.createElement('script');
    script.src = url;
    document.body.appendChild(script);
}

window.manejarRespuestaTipos = function(response) {
    try {
        const rows = response.table.rows;
        const cols = response.table.cols;
        const mesesNombres = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sept", "oct", "nov", "dic"];
        
        const formatearFechaSimple = (valor) => {
            if (!valor || typeof valor !== 'string') return "S/N";
            if (valor.includes("Date(")) {
                const coincidencia = valor.match(/\(([^)]+)\)/);
                if (coincidencia) {
                    const partes = coincidencia[1].split(',');
                    const año = partes[0].trim().slice(-2); 
                    const mesIdx = parseInt(partes[1].trim()); 
                    return `${mesesNombres[mesIdx]} ${año}`;
                }
            }
            return valor;
        };

// Usamos slice(1) para que tome DESDE la segunda columna hasta el final (incluyendo COMPUTADOR)
const nombresEquipos = cols.slice(1).filter(c => c.label && c.label !== "TOTAL" && c.label !== "Suma total").map(c => c.label);        
        // --- CORRECCIÓN DE FILAS: Filtramos para ignorar filas que digan "Suma total" o estén vacías al final ---
        const filasValidas = rows.filter(row => {
            const primerCelda = row.c[0] ? String(row.c[0].v || row.c[0].f) : "";
            return primerCelda !== "" && !primerCelda.toLowerCase().includes("suma");
        });

        // Tomamos los últimos 12 meses reales
        const ultimasFilas = filasValidas.slice(-12);
        
        const mesesCabecera = ultimasFilas.map(row => {
            const val = row.c[0] ? (row.c[0].f || row.c[0].v) : "S/N";
            return formatearFechaSimple(String(val));
        });

      DATOS_TIPOS_EQUIPO = nombresEquipos.map((nombre, equipoIdx) => {
    const idxReal = equipoIdx + 1; // La columna 0 es la fecha, la 1 es el primer equipo...
    const valoresPorMes = ultimasFilas.map(row => {
        return (row.c[idxReal] && row.c[idxReal].v !== null) ? row.c[idxReal].v : 0;
    });
    const totalEquipo = valoresPorMes.reduce((a, b) => a + b, 0);
    return { tipo: nombre.toUpperCase(), valores: valoresPorMes, total: totalEquipo };
});

        // Fila de TOTAL (Suma vertical de la tabla)
        const sumaTotalesMensuales = mesesCabecera.map((_, mesIdx) => {
            let sumaMes = 0;
            DATOS_TIPOS_EQUIPO.forEach(equipo => {
                sumaMes += equipo.valores[mesIdx];
            });
            return sumaMes;
        });

        const granTotalGlobal = sumaTotalesMensuales.reduce((a, b) => a + b, 0);

        DATOS_TIPOS_EQUIPO.push({
            tipo: "TOTAL",
            valores: sumaTotalesMensuales,
            total: granTotalGlobal,
            esFilaTotal: true
        });

        renderizarTablaTipos(mesesCabecera);
    } catch (e) {
        console.error("Error en tipos:", e);
    }
};



function renderizarTablaTipos(meses) {
    const contenedor = document.getElementById('tabla-fallas-equipos');
    if (!contenedor) return;
    DATOS_HISTORICOS_MESES = meses;

    let columnasAMostrar = [...meses];
    if (DATOS_MANUAL_ACTUAL) {
        columnasAMostrar.push(DATOS_MANUAL_ACTUAL.mes);
    }

    const normalizarKey = (t) => t.trim().toUpperCase();
    const azulPJUD = "#2a579a"; // Color azul estándar

    let html = `<table id="tabla-equipos-datos" style="width:100%; border-collapse: collapse; font-family: 'Segoe UI', sans-serif; font-size: 11px; border: 1px solid #ccc;">
        <thead>
            <tr style="background: ${azulPJUD}; color: white; font-weight: bold;">
                <th style="padding: 4px 8px; border: 1px solid #1a3b6c; text-align: left; color: white; background-color: ${azulPJUD};">TIPO</th>
                ${columnasAMostrar.map((m, i) => {
                    // Quitamos el color verde, ahora siempre es azul
                    let boton = (i === columnasAMostrar.length - 1) ? `<br><button onclick="copiarColumnaVertical(${i + 1})" style="margin-top: 2px; cursor: pointer; color: ${azulPJUD}; background: #fff; border:none; border-radius:3px; padding: 0px 4px; font-size: 9px;"><i class="fas fa-copy"></i> Copiar</button>` : "";
                    return `<th style="padding: 4px 8px; border: 1px solid #1a3b6c; text-align:center; background: ${azulPJUD}; color: white;">${m.toUpperCase()}${boton}</th>`;
                }).join('')}
                <th style="padding: 4px 8px; border: 1px solid #1a3b6c; text-align:center; background: #1a3b6c; color: white;">SUMA TOTAL</th>
            </tr>
        </thead>
        <tbody>`;

    DATOS_TIPOS_EQUIPO.forEach(item => {
        const esFilaTotal = item.esFilaTotal || item.tipo === "TOTAL";
        let valoresFila = [...item.valores];
        let totalFila = item.total; 

       // --- BUSCA ESTE BLOQUE DENTRO DE DATOS_TIPOS_EQUIPO.forEach ---
if (DATOS_MANUAL_ACTUAL) {
    let valorExtra = 0;
    if (esFilaTotal) {
        // CAMBIO AQUÍ: En lugar de sumar todo el objeto, sumamos solo lo que coincide 
        // con los tipos de equipo definidos (excluyendo la palabra "TOTAL")
        valorExtra = DATOS_TIPOS_EQUIPO
            .filter(t => t.tipo !== "TOTAL") // No sumamos la fila total a sí misma
            .reduce((suma, t) => {
                const sKey = normalizarKey(t.tipo);
                const match = Object.keys(DATOS_MANUAL_ACTUAL.equipos).find(k => normalizarKey(k) === sKey);
                return suma + (match ? DATOS_MANUAL_ACTUAL.equipos[match] : 0);
            }, 0);
    } else {
        const searchKey = normalizarKey(item.tipo);
        const matchKey = Object.keys(DATOS_MANUAL_ACTUAL.equipos).find(k => normalizarKey(k) === searchKey);
        valorExtra = matchKey ? DATOS_MANUAL_ACTUAL.equipos[matchKey] : 0;
    }
    
    // El resto sigue igual...
    valoresFila.push(valorExtra);
    // IMPORTANTE: Asegúrate de que totalFila sume el valorExtra correctamente
    totalFila = item.valores.reduce((a, b) => a + b, 0) + valorExtra; 
}

        html += `<tr style="${esFilaTotal ? 'background: #f2f2f2; font-weight: bold;' : ''}">
            <td style="padding: 3px 8px; border: 1px solid #ccc; font-weight: bold; background: ${esFilaTotal ? '#f2f2f2' : '#f8f9fa'};">${item.tipo}</td>
            ${valoresFila.map((v, idx) => {
                // Quitamos el fondo verde de la celda de datos
                return `<td style="padding: 3px 8px; border: 1px solid #ccc; text-align:center;">${v.toLocaleString('es-CL')}</td>`;
            }).join('')}
            <td style="padding: 3px 8px; border: 1px solid #ccc; text-align:center; font-weight: bold; background: #f2f2f2;">${totalFila.toLocaleString('es-CL')}</td>
        </tr>`;
    });

    html += `</tbody></table>`;
    contenedor.innerHTML = html;
}

// Nueva función para copiar la columna de arriba hacia abajo
function copiarColumnaVertical(colIdx) {
    const tabla = document.getElementById('tabla-equipos-datos');
    if (!tabla) return;

    let valores = [];
    
    // 1. Capturamos la cabecera (el nombre del mes) de la columna seleccionada
    const cabecera = tabla.querySelector(`thead tr th:nth-child(${colIdx + 1})`);
    if (cabecera) {
        // Limpiamos el texto para quitar la palabra "Copiar" del botón
        let nombreMes = cabecera.innerText.split('\n')[0].trim(); 
        valores.push(nombreMes);
    }

    // 2. Capturamos los datos de cada fila de esa misma columna
    const filasBody = tabla.querySelectorAll('tbody tr');
    filasBody.forEach(fila => {
        const celda = fila.children[colIdx];
        if (celda) {
            // Quitamos puntos de miles para que Excel lo pegue como número
            let valorLimpio = celda.innerText.replace(/\./g, '').trim();
            valores.push(valorLimpio);
        }
    });

    // 3. Unimos con salto de línea para que al pegar en Excel sea una columna
    const textoACopiar = valores.join('\n');
    
    navigator.clipboard.writeText(textoACopiar).then(() => {
        showToast(`✅ Columna ${valores[0]} copiada con éxito`);
    }).catch(err => {
        console.error('Error al copiar: ', err);
    });
}

function ejecutarCopiaTipos(colIdx) {
    try {
        // 1. Extraer valores del mes seleccionado
        const valoresEquipos = DATOS_TIPOS_EQUIPO
            .filter(item => !item.esFilaTotal && item.tipo !== "TOTAL")
            .map(item => item.valores[colIdx]);

        const filaTotal = DATOS_TIPOS_EQUIPO.find(item => item.esFilaTotal || item.tipo === "TOTAL");
        const valorTotalMes = filaTotal ? filaTotal.valores[colIdx] : 0;

        // 2. Unir con tabuladores para Sheets (Horizontal)
        const textoFinal = [...valoresEquipos, valorTotalMes].join('\t');

        // 3. Crear elemento temporal y copiar
        const el = document.createElement('textarea');
        el.value = textoFinal;
        document.body.appendChild(el);
        el.select();
        document.execCommand('copy');
        document.body.removeChild(el);

        // --- SOLUCIÓN PARA LA NOTIFICACIÓN ---
        // Buscamos el elemento del toast en tu HTML
        const toast = document.getElementById('toast-container');
        const toastMsg = document.getElementById('toast-message');

        if (toast && toastMsg) {
            toastMsg.textContent = "¡Copiado horizontalmente!";
            toast.classList.remove('toast-hidden');
            toast.classList.add('toast-show');

            // Lo ocultamos después de 2 segundos
            setTimeout(() => {
                toast.classList.remove('toast-show');
                toast.classList.add('toast-hidden');
            }, 2000);
        } else {
            // Si no encuentra los IDs, intentamos llamar a la función global
            if (typeof mostrarToast === "function") {
                mostrarToast("¡Copiado!");
            }
        }

    } catch (err) {
        console.error("Error al copiar:", err);
    }
}
function copiarColumnaTiposHorizontal(colIdx) {
    try {
        // 1. Extraer valores
        const valoresEquipos = DATOS_TIPOS_EQUIPO
            .filter(item => !item.esFilaTotal && item.tipo !== "TOTAL")
            .map(item => item.valores[colIdx]);

        const filaTotal = DATOS_TIPOS_EQUIPO.find(item => item.esFilaTotal || item.tipo === "TOTAL");
        const valorTotalMes = filaTotal ? filaTotal.valores[colIdx] : 0;

        // 2. Unir con tabuladores para Sheets (Horizontal)
        const textoFinal = [...valoresEquipos, valorTotalMes].join('\t');

        // 3. Intento de copia moderna
        navigator.clipboard.writeText(textoFinal).then(() => {
            lanzarNotificacionCopiado();
        }).catch(err => {
            // 4. Método de respaldo (Fallback) si el anterior falla
            const textArea = document.createElement("textarea");
            textArea.value = textoFinal;
            document.body.appendChild(textArea);
            textArea.select();
            document.execCommand('copy');
            document.body.removeChild(textArea);
            lanzarNotificacionCopiado();
        });
    } catch (err) {
        console.error("Error al copiar:", err);
    }
}

// Función auxiliar para no repetir código del aviso
function lanzarNotificacionCopiado() {
    if (typeof mostrarToast === "function") {
        mostrarToast("¡Copiado horizontalmente!");
    } else if (typeof copiarTextoDirecto === "function") {
        // En algunos de tus scripts el toast se dispara así
        mostrarToast("¡Copiado!");
    } else {
        console.log("Datos copiados al portapapeles.");
    }
}


// --- FUNCIÓN DE UTILIDAD AL FINAL DEL SCRIPT ---
function extraerMesDeCSV(fechaRaw) {
    const mesesNombres = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sept", "oct", "nov", "dic"];
    
    // Limpiamos espacios o comillas que puedan venir del CSV
    const fechaLimpia = String(fechaRaw).replace(/"/g, '').trim();
    
    // Intentamos detectar el separador (puede ser / o -)
    const partes = fechaLimpia.split(/[-/]/);
    
    // Si no tiene el formato esperado, devolvemos un genérico
    if (partes.length < 2) return "MES NUEVO";

    let mesIdx;
    let año;

    // Lógica dinámica para detectar orden de fecha:
    if (partes[0].length === 4) {
        // Formato YYYY-MM-DD
        mesIdx = parseInt(partes[1]) - 1;
        año = partes[0].slice(-2);
    } else {
        // Formato DD/MM/YYYY
        mesIdx = parseInt(partes[1]) - 1;
        año = partes[2] ? partes[2].slice(-2) : "26";
    }

    // Validamos que el índice del mes sea correcto (0 a 11)
    if (isNaN(mesIdx) || mesIdx < 0 || mesIdx > 11) return "MES NUEVO";

    return `${mesesNombres[mesIdx]} ${año}`.toUpperCase();
}

function generarConclusionesDinamicas(mes, datosMDA, datosResi, totalMes) {
    const contenedor = document.getElementById('conclusiones-dinamicas-body');
    if (!contenedor) return;

    // Cálculos
    const dimensionado = 1300;
    const porcDiferenciaDim = (((dimensionado - totalMes) / dimensionado) * 100).toFixed(1);
    const porcMDA = ((datosMDA.solucionados / (totalMes || 1)) * 100).toFixed(1);
    const porcGlobal = ((datosMDA.mesaHP / (totalMes || 1)) * 100).toFixed(1);
    const porcResiTotal = ((datosResi.total / (totalMes || 1)) * 100).toFixed(1);
    const porcCAPJ = datosResi.total > 0 ? ((datosResi.capj / datosResi.total) * 100).toFixed(1) : 0;
    const porcCJ = datosResi.total > 0 ? ((datosResi.cj / datosResi.total) * 100).toFixed(1) : 0;

    // Redacción alineada a la izquierda (text-align: left)
    contenedor.innerHTML = `
        <div style="width: 100%; text-align: left; font-size: 16px;">
            <p>El informe de gestión del servicio de Mesa de Ayuda y Soporte Computacional para el proyecto PJUD 5 durante el mes de <strong>${mes}</strong> presenta un desempeño efectivo y alineado con los objetivos del servicio. Este mes hubo un flujo de requerimientos constante, demostrando una alta capacidad de resolución y una gestión eficiente de los recursos.</p>
            
            <ul style="margin-top: 20px; list-style-type: disc; padding-left: 25px; text-align: left;">
                <li style="margin-bottom: 15px;">
                    <strong>Volumen de Requerimientos:</strong> Se gestionó un total de <strong>${totalMes.toLocaleString('es-CL')}</strong> requerimientos, lo que se sitúa un <strong>${porcDiferenciaDim}%</strong> por debajo de los requerimientos dimensionados para el servicio.
                </li>
                <li style="margin-bottom: 15px;">
                    <strong>Tasa de Resolución:</strong> El servicio mantuvo una sólida tasa de resolución. La Mesa de Ayuda solucionó directamente el <strong>${porcMDA}%</strong> de los requerimientos, mientras que la tasa de resolución de la mesa HP (incluyendo gestiones de residentes y SCO) alcanzó el <strong>${porcGlobal}%</strong>.
                </li>
                <li style="margin-bottom: 15px;">
                    <strong>Gestión de Canales de Atención:</strong> La Mesa de Ayuda resolvió la mayoría de los casos de forma remota, mientras que un <strong>${porcResiTotal}%</strong> de los requerimientos fueron derivados a los técnicos residentes. Las residencias de <strong>CAPJ</strong> y <strong>Centro de Justicia</strong> generaron un <strong>${porcCAPJ}%</strong> y un <strong>${porcCJ}%</strong> respectivamente del total de atenciones en Residencia.
                </li>
            </ul>
        </div>
    `;
}

// --- BASE DE DATOS DE MATRIZ DE TRÁNSITO CORREOS DE CHILE (CONDENSADA) ---
const DATA_MATRIZ = [{comuna:"ACHAO",domiciliaria:"2",sucursal:"2",obs:"",slasoftware:"3",slahardware:"12"},{comuna:"ALERCE",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"ALGARROBO",domiciliaria:"1",sucursal:"2",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"ALHUE",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"ALTO BIOBIO",domiciliaria:"Sin Distribución",sucursal:"3",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"ALTO DEL CARMEN",domiciliaria:"Sin Distribución",sucursal:"3",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"ALTO HOSPICIO",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"7"},{comuna:"ANCUD",domiciliaria:"2",sucursal:"2",obs:"",slasoftware:"3",slahardware:"10"},{comuna:"ANDACOLLO",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"8"},{comuna:"ANGOL",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"8"},{comuna:"ANTARTICA",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"ANTOFAGASTA",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"6"},{comuna:"ANTUCO",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"ARAUCO",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"10"},{comuna:"ARICA",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"6"},{comuna:"AYACARA",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"AYSEN",domiciliaria:"2",sucursal:"2",obs:"",slasoftware:"3",slahardware:"8"},{comuna:"BALMACEDA",domiciliaria:"3",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"BAQUEDANO",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"BARRANCAS (pudahuel)",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"BARROS ARANA",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"BATUCO",domiciliaria:"1",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"BOLLENAR",domiciliaria:"1",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"BUCHUPUREO",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"BUIN",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"2",slahardware:"6"},{comuna:"BULNES",domiciliaria:"1",sucursal:"1",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"3",slahardware:"8"},{comuna:"BUTACHEUQUES",domiciliaria:"Sin Distribución",sucursal:"6",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CABILDO",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CABO DE HORNOS",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"6",slahardware:"12"},{comuna:"CABRERO",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"8"},{comuna:"CACHAPOAL",domiciliaria:"Sin Distribución",sucursal:"4",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CAJON",domiciliaria:"1",sucursal:"Sin Distribución",obs:"Si, solo Radio Urbano",slasoftware:"0",slahardware:"0"},{comuna:"CALAMA",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"6"},{comuna:"CALBUCO",domiciliaria:"1",sucursal:"2",obs:"",slasoftware:"3",slahardware:"10"},{comuna:"CALDERA",domiciliaria:"1",sucursal:"2",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"6"},{comuna:"CALERA DE TANGO",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CALEU",domiciliaria:"5",sucursal:"5",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CALLE LARGA",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CAMARICO",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CAMARONES",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CAMINA",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CAMPANARIO",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CANELA",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CANETE",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"10"},{comuna:"CANITAS",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CAPITAN PASTEN",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CARAHUE",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"8"},{comuna:"CARAMPANGUE",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CARTAGENA",domiciliaria:"1",sucursal:"1",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"0",slahardware:"0"},{comuna:"CASABLANCA",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"6"},{comuna:"CASAS DE CHACABUCO",domiciliaria:"1",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CASMA",domiciliaria:"Sin Distribución",sucursal:"5",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CASTRO",domiciliaria:"2",sucursal:"2",obs:"",slasoftware:"3",slahardware:"10"},{comuna:"CATAPILCO",domiciliaria:"Sin Distribución",sucursal:"4",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CATEMU",domiciliaria:"1",sucursal:"1",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"0",slahardware:"0"},{comuna:"CAUQUENES",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"8"},{comuna:"CERRILLOS",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CERRO NAVIA",domiciliaria:"1",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CERRO SOMBRERO",domiciliaria:"Sin Distribución",sucursal:"2",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CHACAO",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CHAITEN",domiciliaria:"Sin Distribución",sucursal:"6",obs:"",slasoftware:"3",slahardware:"18"},{comuna:"CHANARAL",domiciliaria:"2",sucursal:"4",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"3",slahardware:"10"},{comuna:"CHANARAL ALTO",domiciliaria:"Sin Distribución",sucursal:"2",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CHANCO",domiciliaria:"1",sucursal:"1",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"3",slahardware:"10"},{comuna:"CHAULINEC",domiciliaria:"Sin Distribución",sucursal:"2",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CHEPICA",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CHIGUAYANTE",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"6"},{comuna:"CHILE CHICO",domiciliaria:"Sin Distribución",sucursal:"2",obs:"",slasoftware:"3",slahardware:"18"},{comuna:"CHILLAN",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"3",slahardware:"6"},{comuna:"CHILLAN VIEJO",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CHIMBARONGO",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"0",slahardware:"0"},{comuna:"CHINCOLCO",domiciliaria:"Sin Distribución",sucursal:"3",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CHOLCHOL",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CHONCHI",domiciliaria:"2",sucursal:"2",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"0",slahardware:"0"},{comuna:"CISNES",domiciliaria:"Sin Distribución",sucursal:"2",obs:"",slasoftware:"3",slahardware:"18"},{comuna:"COBQUECURA",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"COCHAMO",domiciliaria:"Sin Distribución",sucursal:"3",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"COCHRANE",domiciliaria:"Sin Distribución",sucursal:"2",obs:"",slasoftware:"3",slahardware:"18"},{comuna:"CODEGUA",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"COELEMU",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"10"},{comuna:"COIHUECO",domiciliaria:"3",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"0",slahardware:"0"},{comuna:"COINCO",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"COLBUN",domiciliaria:"3",sucursal:"1",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"0",slahardware:"0"},{comuna:"COLCHANE",domiciliaria:"5",sucursal:"5",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"COLINA",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"2",slahardware:"4"},{comuna:"COLLIPULLI",domiciliaria:"1",sucursal:"1",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"3",slahardware:"8"},{comuna:"COLTAUCO",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"COMBARBALA",domiciliaria:"1",sucursal:"Sin Distribución",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"8"},{comuna:"CONARIPE",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CONCEPCION",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"4"},{comuna:"CONCHALI",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CONCON",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"0",slahardware:"0"},{comuna:"CONFLUENCIA",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CONSTITUCION",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"8"},{comuna:"CONTULMO",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"COPIAPO",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"6"},{comuna:"COQUIMBO",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"6"},{comuna:"CORONEL",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"8"},{comuna:"CORRAL",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"COYHAIQUE",domiciliaria:"2",sucursal:"2",obs:"",slasoftware:"3",slahardware:"6"},{comuna:"CUNCO",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"0",slahardware:"0"},{comuna:"CURACAUTIN",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"10"},{comuna:"CURACAVI",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"2",slahardware:"8"},{comuna:"CURACO DE VELEZ",domiciliaria:"2",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CURANILAHUE",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"10"},{comuna:"CURANIPE",domiciliaria:"4",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CURARREHUE",domiciliaria:"Sin Distribución",sucursal:"6",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"CUREPTO",domiciliaria:"1",sucursal:"1",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"3",slahardware:"10"},{comuna:"CURICO",domiciliaria:"1",sucursal:"1",obs:"Si, solo Radio Urbano",slasoftware:"3",slahardware:"6"},{comuna:"DALCAHUE",domiciliaria:"2",sucursal:"2",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"DICHATO",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"DIEGO DE ALMAGRO",domiciliaria:"1",sucursal:"2",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"3",slahardware:"10"},{comuna:"DONIHUE",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"EL BOSQUE",domiciliaria:"1",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"EL CARMEN",domiciliaria:"Sin Distribución",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"EL GUADAL",domiciliaria:"Sin Distribución",sucursal:"Sin Distribución",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"EL MELON",domiciliaria:"1",sucursal:"1",obs:"Distribución domiciliaria solo Radio Urbano",slasoftware:"0",slahardware:"0"},{comuna:"EL MONTE",domiciliaria:"1",sucursal:"1",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"EL MURTA",domiciliaria:"Sin Distribución",sucursal:"5",obs:"",slasoftware:"0",slahardware:"0"},{comuna:"EL PALQUI", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "EL QUISCO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "EL RECURSO", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "EL SALADO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "EL SALVADOR", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "12" }, { comuna: "EL TABO", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "EMPEDRADO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "ENTRE LAGOS", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "ERCILLA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "ESTACION CENTRAL", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "ESTACION DE VILLA ALEGRE", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "ESTACION DOMEIKO", domiciliaria: "Sin Distribución", sucursal: "3", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "FARELLONES", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "FLORIDA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "3", slahardware: "10" }, { comuna: "FREIRE", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "FREIRINA", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "FRESIA", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "FRUTILLAR", domiciliaria: "1", sucursal: "2", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "FUTALEUFU", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "3", slahardware: "18" }, { comuna: "FUTRONO", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "GALVARINO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "GENERAL CRUZ", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "GENERAL LAGOS", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "GORBEA", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "GRANEROS", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "GUAITECAS", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "GUANAQUEROS", domiciliaria: "3", sucursal: "3", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "GULTRO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "HIJUELAS", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "HOSPITAL", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "HUALAIHUE", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "3", slahardware: "12" }, { comuna: "HUALANE", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "HUALPEN", domiciliaria: "1", sucursal: "Sin Distribución", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "HUALPIN", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "HUALQUI", domiciliaria: "1", sucursal: "Sin Distribución", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "HUARA", domiciliaria: "1", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "HUASCO", domiciliaria: "3", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "HUECHURABA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "HUEPIL", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "HUERTO FAMILIARES", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "ILLAPEL", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "INDEPENDENCIA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "IQUIQUE", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "ISLA DE MAIPO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "ISLA DE PASCUA", domiciliaria: "Sin Distribución", sucursal: "5", obs: "", slasoftware: "6", slahardware: "12" }, { comuna: "ISLA NEGRA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "ISLA SANTA MARIA", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "JUAN FERNANDEZ", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA CALERA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "6" }, { comuna: "LA CISTERNA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA CRUZ", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA DEHESA (Lo Barnechea)", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA ESTRELLA", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA FLORIDA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA GRANJA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA HIGUERA", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA JUNTA", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA LIGUA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "6" }, { comuna: "LA PARVA", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA PINTANA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA REINA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LA SERENA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "LA UNION", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "LABRANZA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "LAGO RANCO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LAGO VERDE", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LAGUNA BLANCA", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LAGUNA VERDE", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LAJA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "LAMPA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LANCO", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "LARAQUETE", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LAS CABRAS", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "LAS CONDES", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LAUTARO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "LEBU", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "LICANRAY", domiciliaria: "Sin Distribución", sucursal: "4", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LICANTEN", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "LIMACHE", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "LINARES", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "LITUECHE", domiciliaria: "Sin Distribución", sucursal: "5", obs: "", slasoftware: "3", slahardware: "10" }, { comuna: "LIUCURA", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LLANQUIHUE", domiciliaria: "1", sucursal: "2", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "LLAY LLAY", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LLICO", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LLIFEN", domiciliaria: "Sin Distribución", sucursal: "5", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LLO LLEO", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LO BARNECHEA", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LO ESPEJO", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LO MIRANDA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LO PRADO", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LOLOL", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LONCOCHE", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "LONGAVI", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "LONQUIMAY", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LONTUE", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LOS ALAMOS", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "LOS ANDES", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "6" }, { comuna: "LOS ANGELES", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "LOS LAGOS", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "LOS LAURELES", domiciliaria: "Sin Distribución", sucursal: "4", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LOS MUERMOS", domiciliaria: "1", sucursal: "2", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "12" }, { comuna: "LOS SAUCES", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "LOS VILOS", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "6" }, { comuna: "LOTA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "LUMACO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MACHALI", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "MACUL", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MAFIL", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MAIPU", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MAITENCILLO", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MALALHUE", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MALLARAUCO", domiciliaria: "5", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MALLOA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MALLOCO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MANQUEHUA", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MARCHIHUE", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MARIA ELENA", domiciliaria: "3", sucursal: "2", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "MARIA PINTO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MARIQUINA", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "MAULE", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MAULLIN", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "3", slahardware: "10" }, { comuna: "MECHUQUE", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MEJILLONES", domiciliaria: "1", sucursal: "2", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "MELIMOYU", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MELINKA", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MELIPEUCO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MELIPILLA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "2", slahardware: "8" }, { comuna: "MOLINA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "MONTE AGUILA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MONTE PATRIA", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MONTEGRANDE", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MONTENEGRO", domiciliaria: "5", sucursal: "5", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MOSTAZAL", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "MULCHEN", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "NACIMIENTO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "NANCAGUA", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "NATALES", domiciliaria: "2", sucursal: "2", obs: "", slasoftware: "3", slahardware: "12" }, { comuna: "NAVIDAD", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "NEGRETE", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "NINHUE", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "NIPAS", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "NIQUEN", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "NOGALES", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "NUEVA ALDEA", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "NUEVA IMPERIAL", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "NUNOA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "OLIVAR", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "OLLAGUE", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "OLMUE", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "OSORNO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "8" }, { comuna: "OVALLE", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "PADRE HURTADO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PADRE LAS CASAS", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "PAIHUANO", domiciliaria: "Sin Distribución", sucursal: "3", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PAILLACO", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "PAINE", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PALENA", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PALMILLA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PANGUIPULLI", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "PANQUEHUE", domiciliaria: "3", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "PAPUDO", domiciliaria: "1", sucursal: "Sin Distribución", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "PAREDONES", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PARGUA", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PARRAL", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "PEDRO AGUIRRE CERDA", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PELARCO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PELEQUEN", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PELLUHUE", domiciliaria: "4", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PEMUCO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PENAFLOR", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "2", slahardware: "8" }, { comuna: "PENALOLEN", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PENCAHUE", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PENCO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "PERALILLO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "3", slahardware: "10" }, { comuna: "PERQUENCO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PETORCA", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "PEUMO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "3", slahardware: "10" }, { comuna: "PICA", domiciliaria: "1", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PICHIDEGUA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PICHILEMU", domiciliaria: "1", sucursal: "3", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "PINTO", domiciliaria: "4", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PIRQUE", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PISCO ELQUI", domiciliaria: "Sin Distribución", sucursal: "4", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PITRUFQUEN", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "PLACILLA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PLACILLA DE PENUELAS", domiciliaria: "1", sucursal: "Sin Distribución", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "POCONCHILE", domiciliaria: "Sin Distribución", sucursal: "5", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "POLCURA", domiciliaria: "Sin Distribución", sucursal: "4", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "POLPAICO", domiciliaria: "5", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PORTEZUELO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PORTILLO", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PORVENIR", domiciliaria: "3", sucursal: "2", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "12" }, { comuna: "POZO ALMONTE", domiciliaria: "1", sucursal: "2", obs: "", slasoftware: "3", slahardware: "6" }, { comuna: "PRIMAVERA", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PROVIDENCIA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUCHUNCAVI", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "PUCON", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "PUDAHUEL", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "2", slahardware: "4" }, { comuna: "PUEBLO SECO", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUENTE ALTO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "2", slahardware: "4" }, { comuna: "PUENTE NEGRO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUERTO AGUIRRE", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUERTO CHACABUCO", domiciliaria: "2", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUERTO DOMINGUEZ", domiciliaria: "Sin Distribución", sucursal: "5", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUERTO INGENIERO IBANEZ", domiciliaria: "Sin Distribución", sucursal: "4", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUERTO MONTT", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "6" }, { comuna: "PUERTO OCTAY", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUERTO RAUL MARIN BALMACEDA", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUERTO RIO TRANQUILO", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUERTO VARAS", domiciliaria: "1", sucursal: "2", obs: "", slasoftware: "3", slahardware: "8" }, { comuna: "PUERTO WILLIAMS", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUMANQUE", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUNITAQUI", domiciliaria: "3", sucursal: "Sin Distribución", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "PUNTA ARENAS", domiciliaria: "2", sucursal: "2", obs: "", slasoftware: "3", slahardware: "6" }, { comuna: "PUQUELDON", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUREN", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "PURRANQUE", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "PUTAENDO", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "PUTRE", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUTU", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUYEHUE", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "PUYUHUAPI", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUEHUI", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUEILEN", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUELLON", domiciliaria: "2", sucursal: "2", obs: "", slasoftware: "3", slahardware: "12" }, { comuna: "QUEMCHI", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUENAC", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUEPE", domiciliaria: "Sin Distribución", sucursal: "5", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUILACAHUIN", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUILACO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUILICURA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUILIMARI", domiciliaria: "Sin Distribución", sucursal: "3", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUILLECO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUILLON", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "QUILLOTA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "QUILPUE", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "4" }, { comuna: "QUINCHAMALI", domiciliaria: "Sin Distribución", sucursal: "3", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUINCHAO", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUINTA DE TILCOCO", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUINTA NORMAL", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "QUINTERO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "QUIRIHUE", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "RAFAEL", domiciliaria: "Sin Distribución", sucursal: "4", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RANCAGUA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "RANQUIL", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RAUCO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RECINTO", domiciliaria: "Sin Distribución", sucursal: "4", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RECOLETA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RENACA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "RENAICO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RENCA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RENGO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "REQUINOA", domiciliaria: "1", sucursal: "3", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "RETIRO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RINCONADA", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RIO BUENO", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "RIO CLARO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RIO CLARO (VIII)", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RIO HURTADO", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RIO IBANEZ", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RIO NEGRO", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "RIO PUELO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RIO VERDE", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "ROMERAL", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "ROSARIO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "RUNGUE", domiciliaria: "5", sucursal: "5", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAAVEDRA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAGRADA FAMILIA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SALAMANCA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "SAN ANTONIO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "6" }, { comuna: "SAN BERNARDO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "2", slahardware: "4" }, { comuna: "SAN CARLOS", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "SAN CLEMENTE", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "SAN ESTEBAN", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN FABIAN", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN FELIPE", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "6" }, { comuna: "SAN FELIX", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN FERNANDO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "SAN GREGORIO", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN GREGORIO (XVI)", domiciliaria: "Sin Distribución", sucursal: "3", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN IGNACIO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN JAVIER", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "SAN JOAQUIN", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN JOSE DE MAIPO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "SAN JUAN DE LA COSTA", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN MIGUEL", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "2", slahardware: "4" }, { comuna: "SAN MIGUEL DE AZAPA", domiciliaria: "Sin Distribución", sucursal: "5", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN NICOLAS", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN PABLO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN PEDRO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN PEDRO DE ATACAMA", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "SAN PEDRO DE LA PAZ", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "SAN PEDRO DE QUILLOTA", domiciliaria: "1", sucursal: "Sin Distribución", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "SAN RAFAEL", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN RAMON", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN ROSENDO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SAN VICENTE", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "SANTA BARBARA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "10" }, { comuna: "SANTA CLARA", domiciliaria: "Sin Distribución", sucursal: "4", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SANTA CRUZ", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "SANTA FE DE LAJA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SANTA JUANA", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "SANTA MARIA", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "SANTIAGO", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "2", slahardware: "4" }, { comuna: "SANTO DOMINGO", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "SIERRA GORDA", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "SOTAQUI", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "TALAGANTE", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "2", slahardware: "8" }, { comuna: "TALCA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "TALCAHUANO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "TALTAL", domiciliaria: "3", sucursal: "4", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "12" }, { comuna: "TEMUCO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "TENO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "TEODORO SCHMIDT", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "TIERRA AMARILLA", domiciliaria: "1", sucursal: "Sin Distribución", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "TILTIL", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "TIMAUKEL", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "TIRUA", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "TOCOPILLA", domiciliaria: "1", sucursal: "2", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "TOLTEN", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "3", slahardware: "10" }, { comuna: "TOME", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "TONGOY", domiciliaria: "3", sucursal: "3", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "TORRES DEL PAINE", domiciliaria: "Sin Distribución", sucursal: "2", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "TORTEL", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "TRAIGUEN", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "TREGUACO", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "TROVOLHUE", domiciliaria: "Sin Distribución", sucursal: "5", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "TUCAPEL", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "VALDIVIA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "6" }, { comuna: "VALDIVIA DE PAINE", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "VALLENAR", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "VALPARAISO", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "4" }, { comuna: "VICHUQUEN", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "VICTORIA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "VICUNA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "6" }, { comuna: "VILCUN", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "VILLA ALEGRE", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }, { comuna: "VILLA ALEMANA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "4" }, { comuna: "VILLA AMENGUAL", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "VILLA LA TAPERA", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "VILLA MANIHUALES", domiciliaria: "Sin Distribución", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "VILLA O'HIGGINS", domiciliaria: "Sin Distribución", sucursal: "6", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "VILLARRICA", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "8" }, { comuna: "VILUCO", domiciliaria: "1", sucursal: "Sin Distribución", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "VINA DEL MAR", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "4" }, { comuna: "VITACURA", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "YERBAS BUENAS", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "YUMBEL", domiciliaria: "1", sucursal: "1", obs: "", slasoftware: "3", slahardware: "10" }, { comuna: "YUMBEL ESTACION", domiciliaria: "Sin Distribución", sucursal: "1", obs: "", slasoftware: "0", slahardware: "0" }, { comuna: "YUNGAY", domiciliaria: "1", sucursal: "1", obs: "Si, solo Radio Urbano", slasoftware: "3", slahardware: "10" }, { comuna: "ZAPALLAR", domiciliaria: "1", sucursal: "1", obs: "Distribución domiciliaria solo Radio Urbano", slasoftware: "0", slahardware: "0" }];

function renderizarMatriz() {
    const input = document.getElementById('inputBusquedaMatriz');
    const tbody = document.getElementById('bodyMatriz');
    if (!input || !tbody) return;

    const query = input.value.toUpperCase();
    let html = "";

    // Filtramos la data por comuna
    const filtrados = DATA_MATRIZ.filter(item => item.comuna && item.comuna.includes(query));

    if (filtrados.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align:center; padding:20px; color:#666;">No se encontró la comuna "${query}".</td></tr>`;
        return;
    }

    // Función auxiliar para poner "Día" o "Días" o mantener el texto original
    const formatearDia = (valor) => {
        if (!valor || valor === "Sin Distribución") return `<span style="color: #999;">Sin Distribución</span>`;
        if (valor === "1") return `<b>1 Día</b>`;
        return `<b>${valor} Días</b>`;
    };

    filtrados.forEach(item => {
        // Estilo visual: si no hay distribución domiciliaria, se marca en naranja/gris
        const styleDomicilio = item.domiciliaria === "Sin Distribución" ? "color: #999;" : "color: #28a745;";
        const styleSucursal = item.sucursal === "Sin Distribución" ? "color: #999;" : "color: #014f8b;";

        html += `
            <tr style="border-bottom: 1px solid #eee;">
                <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold; background-color: #fcfcfc;">
                    ${item.comuna}
                </td>
                <td style="padding: 10px; border: 1px solid #ddd; text-align: center; ${styleDomicilio}">
                    ${formatearDia(item.domiciliaria)}
                </td>
                <td style="padding: 10px; border: 1px solid #ddd; text-align: center; ${styleSucursal}">
                    ${formatearDia(item.sucursal)}
                </td>
                <td style="padding: 10px; border: 1px solid #ddd; color: #d32f2f; font-size: 0.85rem; font-weight: 500;">
                    ${item.obs || "-"}
                </td>
            </tr>
        `;
    });

    tbody.innerHTML = html;
}

// --- FUNCIÓN PARA RENDERIZAR LA TABLA DE MATRIZ SLA ---
function renderizarSLA() {
    const input = document.getElementById('inputBusquedaSLA');
    const tbody = document.getElementById('bodySLA');
    if (!input || !tbody) return;

    const query = input.value.toUpperCase();
    let html = "";

    // Filtramos la data: 
    // 1. Que coincida con la búsqueda.
    // 2. Que al menos uno de los dos SLA sea distinto de "0", "" o "N/A" (Para no mostrar filas vacías).
    const filtrados = DATA_MATRIZ.filter(item => {
        const coincidenciaNombre = item.comuna && item.comuna.includes(query);
        const tieneSLA = (item.slasoftware !== "0" && item.slasoftware !== "" && item.slasoftware !== "N/A") || 
                         (item.slahardware !== "0" && item.slahardware !== "" && item.slahardware !== "N/A");
        return coincidenciaNombre && tieneSLA;
    });

    if (filtrados.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" style="text-align:center; padding:20px; color:#666;">No hay datos de SLA disponibles para esta búsqueda.</td></tr>`;
        return;
    }

    // Función interna para mostrar directamente el valor como Horas
    const formatearDirectoAHoras = (valor) => {
        if (!valor || valor === "0" || valor === "N/A" || valor === "") {
            return `<span style="color: #999;">N/A</span>`;
        }
        // Agregamos directamente "Hrs" sin multiplicar
        return `<b>${valor} Hrs</b>`;
    };

    filtrados.forEach(item => {
        html += `
            <tr style="border-bottom: 1px solid #eee;">
                <td style="padding: 10px; border: 1px solid #ddd; font-weight: bold; background-color: #fcfcfc;">
                    ${item.comuna}
                </td>
                <td style="padding: 10px; border: 1px solid #ddd; text-align: center; color: #014f8b;">
                    ${formatearDirectoAHoras(item.slasoftware)}
                </td>
                <td style="padding: 10px; border: 1px solid #ddd; text-align: center; color: #d32f2f;">
                    ${formatearDirectoAHoras(item.slahardware)}
                </td>
            </tr>
        `;
    });

    tbody.innerHTML = html;
}


// Llamar a la función al cargar la página para que la tabla inicie con datos
document.addEventListener('DOMContentLoaded', () => {
    renderizarMatriz();
    renderizarSLA();    // Carga la tabla de Tiempos de Solución
});


function copiarTexto(texto, btn) {
    navigator.clipboard.writeText(texto).then(() => {
        const iconoOriginal = btn.innerHTML;
        btn.innerHTML = '<i class="fas fa-check" style="color: #28a745;"></i>';
        setTimeout(() => {
            btn.innerHTML = iconoOriginal;
        }, 1500);
    }).catch(err => {
        console.error('Error al copiar: ', err);
    });
}

// Función genérica para abrir/cerrar secciones (la misma que usas en WhatsApp)
function toggleSection(bodyId, iconId) {
    const body = document.getElementById(bodyId);
    const icon = document.getElementById(iconId);
    if (body.style.display === "none" || body.style.display === "") {
        body.style.display = "block";
        icon.style.transform = "rotate(180deg)";
    } else {
        body.style.display = "none";
        icon.style.transform = "rotate(0deg)";
    }
}

