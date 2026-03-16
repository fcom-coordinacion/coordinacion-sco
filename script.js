const VALID_PIN = "1234"; 
let selectedTicketData = null; 
let currentTicketsList = []; 
let allTicketsList = []; 

// --- 1. CONFIGURACIÓN DE TÉCNICOS POR JURISDICCIÓN ---
const TECNICOS_POR_JURISDICCION = {
    "Corte De Apelaciones De Arica": ["BENJAMIN  ZUÑIGA "],
    "Corte De Apelaciones De Iquique": ["MARCELO  CORTEZ "],
    "Corte De Apelaciones De Antofagasta": ["MARIO  MORENO", "ANDERSON  JARA "],
    "Corte De Apelaciones De Copiapo": ["ALVARO  SILVA", "JUAN GABRIEL SEPULVEDA "],
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

const CORREOS_FIJOS = "c.zapata@fcom.cl; a.vacca@fcom.cl; s.guzman@fcom.cl; dvillarroels@pjud.cl; j.marrufo@fcom.cl";

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


// --- 4. PROCESAMIENTO DE DATOS ---
function processData() {
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

        // Recuperar memoria local (Con bloque TRY para evitar que se rompa si está bloqueado)
        let savedCoords = {};
        try {
            savedCoords = JSON.parse(localStorage.getItem('sco_coordinations') || '{}');
        } catch(err) {
            console.warn("Memoria caché bloqueada por el navegador.");
        }

        rows.forEach((row) => {
            const delimiter = row.includes('\t') ? '\t' : ';';
            const cols = row.split(delimiter);

            if (cols.length < 5) return; 

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

            const cleanCol = (idx) => cols[idx] ? cols[idx].replace(/"/g, "").trim() : "";
            const rawGrupo = cleanCol(3);

            let rawIP = "-";
            if (cols[11] && cols[11].includes('.')) {
                rawIP = cols[11]; 
            }

            const ticketNum = cleanCol(1);
            
            const ticket = {
                num: ticketNum,
                estado: cleanCol(2),
                grupo: rawGrupo,
                finalizadoPor: cleanCol(5), 
                usuario: cleanCol(8),
                correo: cleanCol(10), 
                ip: rawIP.replace(/"/g, "").trim(),
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
                despachosRaw: cleanCol(33) || cleanCol(27) 
            };

            // Inyectar datos guardados si existen en memoria local
            if (savedCoords[ticketNum]) {
                ticket.fechaCoord = savedCoords[ticketNum].fechaCoord;
                ticket.horaCoord = savedCoords[ticketNum].horaCoord;
                ticket.tecnicoCoord = savedCoords[ticketNum].tecnicoCoord;
            }

            allTicketsList.push(ticket);

            // FILTRADO VISUAL
            const estadoUpper = ticket.estado.toUpperCase();
            const grupoUpper = ticket.grupo.toUpperCase();
            const isPending = !estadoUpper.includes("FINALIZADO") && !estadoUpper.includes("CERRADO");
            let isTargetGroup = (selectedGroup === "TODOS") ? true : grupoUpper.includes(selectedGroup);

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

                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td style="color: var(--hp-blue); font-weight: bold;">${ticket.proyecto || 'N/A'}</td>
                    <td><b>${ticket.num}</b></td>
                    <td><span class="status-pill" style="${statusStyle}">${ticket.estado}</span></td>
                    <td style="font-size: 0.8rem; color: #666;">${ticket.grupo}</td>
                    <td>${ticket.dependencia} <br><small style="color: #999;">(${ticket.jurisdiccion})</small></td>
                    <td style="text-align:center; background-color: #f9fcff;">${fechaDisplay}</td>
                    <td style="text-align:center; background-color: #f9fcff;">${tecnicoDisplay}</td>
                    <td style="text-align:center;">${ticket.backup}</td>
                    <td><button class="btn-primary" style="padding: 5px 10px; font-size: 0.8rem;" onclick="openCoordEditor(${ticketIndex})">Gestionar</button></td>
                `;
                if(tableBody) tableBody.appendChild(tr);
            }
        });

        if (allTicketsList.length > 0) {
            if (currentTicketsList.length > 0) showToast(`Cargados: ${currentTicketsList.length} tickets de ${selectedGroup}.`);
            else showToast(`Cargados ${allTicketsList.length} registros totales.`);
        } else {
            showToast("⚠️ No se encontraron datos válidos en el archivo.");
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


// --- 5. FUNCIÓN GUARDAR COMPLETAMENTE REESCRITA Y A PRUEBA DE FALLOS ---
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

    // 1. Guardar en memoria RAM
    selectedTicketData.fechaCoord = fechaInput;
    selectedTicketData.horaCoord = horaInput;
    selectedTicketData.tecnicoCoord = techInput;

    // 2. Guardar en memoria de Navegador (Con try-catch para evitar bloqueos locales)
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

    // 3. Pintar en la tabla usando el Número de Ticket directamente del HTML
    try {
        const tableBody = document.querySelector('#tickets-table tbody');
        const filas = tableBody.querySelectorAll('tr');
        let encontrado = false;

        for (let i = 0; i < filas.length; i++) {
            const fila = filas[i];
            const celdaTicket = fila.children[1].innerText.trim();
            
            if (celdaTicket === selectedTicketData.num.trim()) {
                const partes = fechaInput.split('-'); 
                const fechaFormateada = partes.length === 3 ? `${partes[2]}-${partes[1]}-${partes[0]}` : fechaInput;
                
                fila.children[5].innerHTML = `<span style="color:#007bff; font-weight:bold;">${fechaFormateada}</span><br><small>${horaInput}</small>`;
                fila.children[6].innerHTML = `<span style="font-size:0.85em; font-weight:600;">${techInput}</span>`;
                
                fila.style.backgroundColor = "#d4edda"; // Color verde de éxito
                setTimeout(() => fila.style.backgroundColor = "", 800);
                encontrado = true;
                break;
            }
        }

        if (encontrado) {
            showToast(`✅ Guardado exitoso para TK ${selectedTicketData.num}`);
        } else {
            showToast("✅ Guardado en memoria (Fila visual no encontrada).");
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

    const tablaEstilizada = `
    <table style="border-collapse: collapse; width: 100%; max-width: 400px; font-family: Segoe UI, Calibri, Arial, sans-serif; border: 1px solid #014f8b; font-size: 0.85rem;">
        <thead><tr><th colspan="2" style="background-color: #014f8b; color: #ffffff !important; padding: 8px; text-align: center; font-size: 0.9rem;">DETALLE DE COORDINACIÓN</th></tr></thead>
        <tbody>
            <tr style="background-color: #f2f2f2;"><td style="padding: 6px; border: 1px solid #ddd; font-weight: bold; width: 35%;">TICKET</td><td style="padding: 6px; border: 1px solid #ddd;">${selectedTicketData.num}</td></tr>
            <tr><td style="padding: 6px; border: 1px solid #ddd; font-weight: bold;">PROYECTO</td><td style="padding: 6px; border: 1px solid #ddd;">${selectedTicketData.proyecto}</td></tr>
            <tr style="background-color: #f2f2f2;"><td style="padding: 6px; border: 1px solid #ddd; font-weight: bold;">DEPENDENCIA</td><td style="padding: 6px; border: 1px solid #ddd;">${selectedTicketData.dependencia}</td></tr>
            <tr><td style="padding: 6px; border: 1px solid #ddd; font-weight: bold;">ACTIVIDAD</td><td style="padding: 6px; border: 1px solid #ddd;">${actividad}</td></tr>
            <tr style="background-color: #f2f2f2;"><td style="padding: 6px; border: 1px solid #ddd; font-weight: bold;">FECHA</td><td style="padding: 6px; border: 1px solid #ddd;">${date}</td></tr>
            <tr><td style="padding: 6px; border: 1px solid #ddd; font-weight: bold;">HORA</td><td style="padding: 6px; border: 1px solid #ddd;">${time}</td></tr>
            <tr style="background-color: #f2f2f2;"><td style="padding: 6px; border: 1px solid #ddd; font-weight: bold;">TÉCNICO</td><td style="padding: 6px; border: 1px solid #ddd;">${tech}</td></tr>
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
                    <p style="font-family: Calibri, sans-serif; font-size: 0.9rem;">Estimado(a) ${primerNombre},<br><br>
                    Junto con saludar, informamos que se ha coordinado la atención de su requerimiento con el siguiente detalle:<br><br></p>
                    ${tablaEstilizada}
                    <br>
                    <p style="font-family: Calibri, sans-serif; font-size: 0.75rem; color: #777; margin-top: 15px; line-height: 1.2;">
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

    const mensaje = `*TK:* ${selectedTicketData.num}\n🧰 *Tecnico:* ${tech}\n📀 *Proyecto:* ${selectedTicketData.proyecto}\n🏛️ *Tribunal:* ${selectedTicketData.dependencia}\n🏛️ *Direccion:* ${dir}\n🗓️ *Fecha:* ${date}\n⏰ *Hora:* ${time}\n📝 *Actividad:* ${actividad}\n🚨 *OBLIGATORIO:* COLOCAR LA IP EN TODAS LAS ATENCIONES\n\n⚠️ *Nota:* Revisar que la direccion sea la correcta. Descrita en la Mauweb`;
    
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
        if (Date.now() - start < 1000) showToast("Outlook bloqueado. Use copiar.");
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
}

// --- REPORTE: INFORMATIVO TÉCNICO ---
function generarInformativoTecnico() {
    if (currentTicketsList.length === 0) {
        showToast("⚠️ No hay tickets procesados para generar el reporte.");
        return;
    }

    const container = document.getElementById('info-result-container');
    const asunto = `Informativo Técnico - Actividades Coordinadas SCO PJUD`;
    const para = "j.santos@fcom.cl";
    const cc = "c.zapata@fcom.cl; e.suarez@fcom.cl; e.socorro@fcom.cl; l.torres@fcom.cl; juan.diaz@fcom.cl; sandrade_fcom@pjud.cl; jchavez_hp@pjud.cl; s.guzman@fcom.cl; jmarrufo_hp@pjud.cl; f.solar@fcom.cl; svaldivieso_hp@pjud.cl; myabrudez_fcom@pjud.cl; j.riffo@fcom.cl; a.vacca@fcom.cl";

    let filasTabla = "";
    
    currentTicketsList.forEach(ticket => {
        const esCambio = (ticket.backup && ticket.backup.toUpperCase().trim() === "SI");
        const solucion = esCambio ? "CAMBIO DE EQUIPO" : "EVALUACIÓN / CONFIGURACIÓN DE EQUIPO";
        
        let fechaShow = "POR DEFINIR";
        if (ticket.fechaCoord) {
            const [y, m, d] = ticket.fechaCoord.split('-');
            fechaShow = `${d}/${m}/${y}`;
            if (ticket.horaCoord) fechaShow += ` ${ticket.horaCoord}`;
        }

        const tecnicoShow = ticket.tecnicoCoord || "POR ASIGNAR";

        filasTabla += `
            <tr style="font-size: 11px; border-bottom: 1px solid #ddd;">
                <td style="padding: 4px; border: 1px solid #ccc;">${ticket.num}</td>
                <td style="padding: 4px; border: 1px solid #ccc;">${ticket.proyecto}</td>
                <td style="padding: 4px; border: 1px solid #ccc;">${ticket.modelo}</td>
                <td style="padding: 4px; border: 1px solid #ccc;">${ticket.serie}</td>
                <td style="padding: 4px; border: 1px solid #ccc;">${ticket.dependencia}</td>
                <td style="padding: 4px; border: 1px solid #ccc;">${ticket.tipo}</td>
                <td style="padding: 4px; border: 1px solid #ccc; background-color: ${fechaShow === "POR DEFINIR" ? '#fff' : '#d1e7dd'};">${fechaShow}</td>
                <td style="padding: 4px; border: 1px solid #ccc; background-color: ${tecnicoShow === "POR ASIGNAR" ? '#fff' : '#d1e7dd'};">${tecnicoShow}</td>
                <td style="padding: 4px; border: 1px solid #ccc;">${solucion}</td>
            </tr>
        `;
    });

    const htmlReporte = `
        <div class="pjud-header-toggle" onclick="toggleSection('info-body', 'icon-info')">
            <h3>Vista Previa: Informativo Técnico</h3>
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
                Junto con saludar, informo de las actividades que se realizarán el día de mañana por el área de SCO PJUD.
                Se brinda información del técnico, la dependencia y la actividad a realizar. Favor revisar para poder apoyar a los técnicos que estarán en dichos procesos, de esta forma puedan contar con las herramientas, imágenes y procedimientos actualizados.
                <br><br>Actividades coordinadas Para mañana:<br>
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
    let csvContent = "TICKET;PROYECTO;MODELO EQUIPO;SERIE REPORTADA;DEPENDENCIA;TIPO;FECHA COORDINACIÓN;TECNICO COORDINADO;SOLUCION TERRENO\n";
    currentTicketsList.forEach(ticket => {
        const esCambio = (ticket.backup && ticket.backup.toUpperCase().trim() === "SI");
        const solucion = esCambio ? "CAMBIO DE EQUIPO" : "EVALUACIÓN / CONFIGURACIÓN DE EQUIPO";
        const row = [ticket.num, ticket.proyecto, ticket.modelo, ticket.serie, ticket.dependencia, ticket.tipo, "POR DEFINIR", "POR ASIGNAR", solucion];
        csvContent += row.join(";") + "\n";
    });
    const blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `Informativo_Tecnico_${new Date().toISOString().slice(0,10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast("✅ Archivo Excel descargado");
}

// --- REPORTE: EN BLANCO (AÑADIDO NUEVAMENTE) ---
function generarReporteEnBlanco() {
    const container = document.getElementById('blanco-result-container');
    const fInicio = document.getElementById('blanco-fecha-inicio').value;
    const fFin = document.getElementById('blanco-fecha-fin').value;

    if (allTicketsList.length === 0) {
        showToast("⚠️ Carga datos en el módulo de Coordinación primero.");
        return;
    }

    let dDesde = fInicio ? new Date(fInicio + "T00:00:00") : null;
    let dHasta = fFin ? new Date(fFin + "T23:59:59") : null;

    const casosEnBlanco = allTicketsList.filter(t => {
        const sinModelo = t.modelo === "Modelo N/A" || !t.modelo || t.modelo.trim() === "";
        const sinTipo = t.tipo === "Equipo" || !t.tipo || t.tipo.trim() === "";
        if (!(sinModelo || sinTipo)) return false;

        if (dDesde && dHasta) {
            const parts = t.fechaFin.split(' ')[0].split(/[-/]/);
            if(parts.length === 3) {
                 const fechaT = new Date(`${parts[2]}-${parts[1]}-${parts[0]}T12:00:00`);
                 if (fechaT < dDesde || fechaT > dHasta) return false;
            }
        }
        return true;
    });

    if (casosEnBlanco.length === 0) {
        container.innerHTML = `<div style="padding:15px; background:#d4edda; color:#155724; border-radius:5px;">✅ No se encontraron requerimientos en blanco en este periodo.</div>`;
        container.classList.remove('hidden');
        return;
    }

    let tablaHTML = `
        <div style="padding:15px; background:#fff3cd; color:#856404; border-radius:5px; margin-bottom:15px;">⚠️ Se encontraron <b>${casosEnBlanco.length}</b> requerimientos sin información técnica.</div>
        <div style="overflow-x:auto;">
            <table id="tabla-en-blanco" style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif; font-size: 11px;">
                <thead style="background-color: #6c757d; color: white;">
                    <tr>
                        <th style="padding:8px; border:1px solid #ddd;">Ticket</th>
                        <th style="padding:8px; border:1px solid #ddd;">Grupo Resolutor</th>
                        <th style="padding:8px; border:1px solid #ddd;">Fecha Finalizado</th>
                        <th style="padding:8px; border:1px solid #ddd;">Modelo Detectado</th>
                        <th style="padding:8px; border:1px solid #ddd;">Tipo Detectado</th>
                        <th style="padding:8px; border:1px solid #ddd;">Dependencia</th>
                    </tr>
                </thead>
                <tbody>
    `;

    casosEnBlanco.forEach(t => {
        tablaHTML += `
            <tr style="border-bottom: 1px solid #eee;">
                <td style="padding:8px; border:1px solid #ddd;"><b>${t.num}</b></td>
                <td style="padding:8px; border:1px solid #ddd;">${t.grupo}</td>
                <td style="padding:8px; border:1px solid #ddd;">${t.fechaFin}</td>
                <td style="padding:8px; border:1px solid #ddd; color: ${t.modelo === 'Modelo N/A' ? '#dc3545' : 'inherit'}; font-weight: ${t.modelo === 'Modelo N/A' ? 'bold' : 'normal'};">${t.modelo}</td>
                <td style="padding:8px; border:1px solid #ddd; color: ${t.tipo === 'Equipo' ? '#dc3545' : 'inherit'}; font-weight: ${t.tipo === 'Equipo' ? 'bold' : 'normal'};">${t.tipo}</td>
                <td style="padding:8px; border:1px solid #ddd;">${t.dependencia}</td>
            </tr>
        `;
    });

    tablaHTML += `</tbody></table></div>`;
    container.innerHTML = tablaHTML;
    container.classList.remove('hidden');
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

    let catalogos = { skus: {}, ciudades: {} };

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
        const skuEncontrado = catalogos.skus[modeloLimpio] || "#N/A";
        const ciudadEncontrada = catalogos.ciudades[dependenciaLimpia] || "#N/A";

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

    if (!inputFecha) {
        showToast("⚠️ Por favor selecciona una fecha.");
        return;
    }

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
        const fechaTicket = parseDateSimple(t.fechaFin);
        if (!fechaTicket) return false;
        return fechaTicket.getTime() === fechaSeleccionada.getTime();
    });

    if (ticketsFiltrados.length === 0) {
        showToast(`⚠️ No hay tickets finalizados para el ${fechaFormat}.`);
        container.innerHTML = "<p style='text-align:center; color:#666; padding: 20px;'>Sin resultados para la fecha seleccionada.</p>";
        container.classList.remove('hidden');
        return;
    }

    let filasHTML = "";
    ticketsFiltrados.forEach(t => {
        let serieDespachada = "";
        if (t.despachosRaw && t.despachosRaw.includes(":")) {
            const parts = t.despachosRaw.split(":");
            if(parts.length > 1) serieDespachada = parts[1].replace("|", "").trim();
        }
        filasHTML += `
            <tr style="border-bottom: 1px solid #ddd;">
                <td style="padding: 5px; border: 1px solid #ccc;">${t.proyecto || ""}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.num}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.grupo || ""}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.tipo || ""}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.solucion || ""}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.serie || ""}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${serieDespachada}</td>
                <td style="padding: 5px; border: 1px solid #ccc;">${t.ip || ""}</td>
            </tr>
        `;
    });

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
                            <tr>
                                <th style="padding: 5px; border: 1px solid #ddd;">PROYECTO</th>
                                <th style="padding: 5px; border: 1px solid #ddd;">TK</th>
                                <th style="padding: 5px; border: 1px solid #ddd;">GRUPO RESOLUTOR</th>
                                <th style="padding: 5px; border: 1px solid #ddd;">TIPO</th>
                                <th style="padding: 5px; border: 1px solid #ddd;">SOLUCION TERRENO</th>
                                <th style="padding: 5px; border: 1px solid #ddd;">SERIE REPORTADA</th>
                                <th style="padding: 5px; border: 1px solid #ddd;">SERIE DESPACHADA</th>
                                <th style="padding: 5px; border: 1px solid #ddd;">IP</th>
                            </tr>
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
                    <p>Envío listado de los requerimientos gestionados el día de ayer, que involucraron Cambio o masterización de equipo.</p>
                    <br><br>
                    <table style="border-collapse: collapse; width: 100%; border: 1px solid #999; font-family: Calibri, sans-serif; font-size: 10pt;">
                        <thead style="background-color: #e6e6e6;">
                            <tr>
                                <th style="border: 1px solid #999; padding: 4px; text-align: left;">PROYECTO</th>
                                <th style="border: 1px solid #999; padding: 4px; text-align: left;">TK</th>
                                <th style="border: 1px solid #999; padding: 4px; text-align: left;">GRUPO RESOLUTOR</th>
                                <th style="border: 1px solid #999; padding: 4px; text-align: left;">TIPO</th>
                                <th style="border: 1px solid #999; padding: 4px; text-align: left;">SOLUCION TERRENO</th>
                                <th style="border: 1px solid #999; padding: 4px; text-align: left;">SERIE REPORTADA</th>
                                <th style="border: 1px solid #999; padding: 4px; text-align: left;">SERIE DESPACHADA</th>
                                <th style="border: 1px solid #999; padding: 4px; text-align: left;">IP</th>
                            </tr>
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
function generarReporteGuiasPendientes() {
    const container = document.getElementById('guias-pendientes-result');
    const grupoSeleccionado = document.getElementById('filtro-grupo-guias').value;

    if (allTicketsList.length === 0) {
        showToast("⚠️ No hay datos cargados. Carga un archivo primero.");
        return;
    }

    const ticketsEnRecuperacion = [
        "7493206","7494261","7496812","7496968","7524642","7494261","7496969","7539950","7544696","7560826","7581956","7598371","7614098","7605150","7637221","7662905","7635910","7690441","7692876","7492515","7494920","7495317","7495546","7495776","7802224","7842786","7841094"
    ]; 

    const ticketsFiltrados = allTicketsList.filter(t => {
        const backupNorm = t.backup.normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
        if (backupNorm !== "SI") return false;

        const guiaNorm = t.guiaRetiro.normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toUpperCase();
        if (guiaNorm !== "NO" && guiaNorm !== "") return false;

        const tipoUpper = t.tipo ? t.tipo.toUpperCase() : "";
        if (tipoUpper.includes("MOUSE") || tipoUpper.includes("TECLADO")) return false;

        const grupoUpper = t.grupo.toUpperCase();
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
    ticketsFiltrados.forEach(t => {
        const estaEnRecuperacion = ticketsEnRecuperacion.includes(t.num.trim());
        const badgeEstado = estaEnRecuperacion 
            ? `<span style="background:#fff3cd; color:#856404; padding:2px 6px; border-radius:4px; font-weight:bold; font-size:10px; border:1px solid #ffeeba;">CON USUARIO - REVISAR</span>`
            : `<span style="color:#666; font-size:10px;">PENDIENTE RETIRO</span>`;

        filasHTML += `
            <tr style="border-bottom: 1px solid #ddd; ${estaEnRecuperacion ? 'background-color: #fffdf5;' : ''}">
                <td style="padding: 8px; border: 1px solid #ccc; text-align: center; font-weight:bold;">${t.num}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.fechaFin}</td>
                <td style="padding: 8px; border: 1px solid #ccc; text-align:center;">${badgeEstado}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.finalizadoPor || t.usuario}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.tipo}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.serie}</td>
                <td style="padding: 8px; border: 1px solid #ccc;">${t.despachosRaw}</td>
                <td style="padding: 8px; border: 1px solid #ccc; font-weight:bold; color:#014f8b;">${t.ip || "-"}</td>
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
                    <th>Tipo</th>
                    <th>Serie</th>
                    <th>Despachos</th>
                    <th style="background-color: #014f8b; color: white;">Dirección IP</th>
                </tr>
            </thead>
            <tbody>${filasHTML}</tbody>
        </table>
    `;

    container.classList.remove('hidden');
    showToast(`⚠️ Se encontraron ${ticketsFiltrados.length} guías pendientes.`);
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

// BOTÓN DESCARGAR PROMANAGER DESHABILITADO PARA MODO ESTÁTICO
function ejecutarScriptPython() {
    alert("Esta función requiere un servidor backend (como el anterior en Render) y no está disponible en la versión estática de GitHub Pages. Por favor adjunta el archivo manualmente usando el botón 'Adjuntar Reporte'.");
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