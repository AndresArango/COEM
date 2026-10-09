/* ================================================
   Seguimiento Actividades MS
   Reutiliza la misma lógica de roster y barras
   segmentadas que la tarjeta de Programas del home.
   ================================================ */

const EXCEL_PATH = "../ventas_occidente.xlsx";
const BIRTHDAY_EXCEL_PATH = "../cumpleanos.xlsx"; // Programas + IN30 viven aquí

let cumplimientoMesData = [];
let programasData = [];
let in30Data = [];

/* ── Helpers básicos (iguales al dashboard principal) ── */

function normalize(s) {
  return String(s||"").toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .replace(/[^a-z0-9]/g,"");
}

function esc(s){ return String(s||"—").replace(/[&<>]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[m]||m)); }

function esFilaTotalGeneral(nombreRaw){
    const n = normalize(nombreRaw);
    return n === "totalgeneral" || n === "grandtotal" || n === "total";
}

function esNombrePersona(nombreRaw){
    const nombre = String(nombreRaw || "").trim();
    if(!nombre) return false;
    const partes = nombre.split(/\s+/);
    if(partes.length < 2) return false;
    if(nombre.includes("_")) return false;
    if(partes.some(p => normalize(p) === "nn")) return false;
    if(partes.length === 2 && normalize(partes[0]) === normalize(partes[1])) return false;
    const nombreNormalizado = normalize(nombre);
    if(nombreNormalizado.includes("total")) return false;
    if(nombreNormalizado.includes("resumen")) return false;
    return true;
}

function shortName(nombre){
    const partes = nombre.trim().split(/\s+/);
    if(partes.length === 2) return `${partes[1]} ${partes[0]}`;
    if(partes.length >= 3){
        const primerNombre = partes[partes.length - 2];
        const primerApellido = partes[0];
        return `${primerNombre} ${primerApellido}`;
    }
    return nombre;
}

// Lee una columna de una fila sin importar cómo esté escrito el
// encabezado en el Excel (Cuenta / CUENTA / Cliente / CLIENTE, etc.)
function campo(obj, ...nombresPosibles){
    const keys = Object.keys(obj || {});
    for(const nombre of nombresPosibles){
        const found = keys.find(k => normalize(k) === normalize(nombre));
        if(found) return obj[found];
    }
    return "";
}

function palabrasNombre(nombre){
    return String(nombre || "").trim().split(/\s+/).map(normalize).filter(w => w.length > 2);
}

function mejorIndiceRoster(nombreComercial, roster){
    const palabrasComercial = palabrasNombre(nombreComercial);
    let mejorIdx = -1;
    let mejorScore = 0;
    roster.forEach((r, idx) => {
        const palabrasRoster = palabrasNombre(r.nombreCompleto);
        const score = palabrasComercial.filter(w => palabrasRoster.includes(w)).length;
        if(score > mejorScore){
            mejorScore = score;
            mejorIdx = idx;
        }
    });
    return mejorIdx;
}

const PROGRAMAS_ROSTER_EXCLUIR = ["espinosa", "collazos"];

function getRosterComerciales(){
    return cumplimientoMesData
        .map(r => String(r["Etiquetas de fila"] || "").trim())
        .filter(nombre =>
            nombre !== "" &&
            nombre !== "VALLE" &&
            !esFilaTotalGeneral(nombre) &&
            esNombrePersona(nombre) &&
            !PROGRAMAS_ROSTER_EXCLUIR.includes(normalize(nombre.split(/\s+/)[0]))
        )
        .map(nombreCompleto => {
            const partes = nombreCompleto.split(/\s+/);
            return {
                nombreCompleto,
                apellido: normalize(partes[0]),
                display: shortName(nombreCompleto)
            };
        });
}

function pintarBarraSegmentada(wrapEl, aprobado, postulado, ofrecido, maxTotal){
    if(!wrapEl) return;
    const escala = Math.max(1, maxTotal);
    const pctA = Math.min(100, Math.round((aprobado  / escala) * 100));
    const pctP = Math.min(100, Math.round((postulado / escala) * 100));
    const pctO = Math.min(100, Math.round((ofrecido  / escala) * 100));
    wrapEl.innerHTML = `
        <div class="seg seg-aprobado"  style="left:0%; width:${pctA}%"></div>
        <div class="seg seg-postulado" style="left:${pctA}%; width:${pctP}%"></div>
        <div class="seg seg-ofrecido"  style="left:${pctA + pctP}%; width:${pctO}%"></div>
    `;
}

function estadoActividad(p){
    const e = normalize(String(campo(p, "Estado") || ""));
    if(e === "aprobado") return "aprobado";
    if(e === "postulado") return "postulado";
    if(e === "ofrecido") return "ofrecido";
    return "otro";
}

/* ── Config: dominios de Programas y actividades IN30 ── */

const PROGRAMAS_DOMINIOS = {
    seguridad: { label: "Seguridad", responsable: "Corella" },
    copilot:   { label: "Copilot",   responsable: "Mario" },
    azure:     { label: "Azure",     responsable: "Omar" }
};
const PROGRAMAS_POSTULA_GENERAL = "Mario";

function dominioPrograma(p){
    const explicito = normalize(String(campo(p, "Dominio") || ""));
    if(PROGRAMAS_DOMINIOS[explicito]) return explicito;
    const n = normalize(String(campo(p, "Programa") || ""));
    if(n.includes("copilot")) return "copilot";
    if(n.includes("security") || n.includes("seguridad")) return "seguridad";
    if(n.includes("azure")) return "azure";
    return "otro";
}

// Clave que debe tener la columna "Actividad" en la hoja IN30 del Excel
// para que cada fila caiga en la tarjeta correcta.
const ACTIVIDADES_MS = {
    copilotIn30: {
        clave: "copilotin30",
        nombre: "Copilot in 30",
        icono: "🤖",
        descripcion: "Prueba de un mes para Copilot 365 (25 licencias), con entrenamiento especializado por usuario y compromiso de compra.",
        postula: "Jhony y Mario",
        ofrece: "Jhony y Mario",
        cumpleObjetivoMS: true
    },
    coemBackupIn30: {
        clave: "coembackupin30",
        nombre: "COEM Backup in 30",
        icono: "💾",
        descripcion: "Prueba de un mes para COEM Backup, sin compromiso de compra.",
        postula: "Mario y Nerly",
        ofrece: "Corella",
        cumpleObjetivoMS: false
    },
    assessment: {
        clave: "assessment",
        nombre: "Assessment",
        icono: "🔍",
        descripcion: "Evaluación realizada por Microsoft donde entrega un informe del estado actual.",
        postula: "Nerly o Zully",
        ofrece: "Corella y Nerly",
        cumpleObjetivoMS: false
    }
};

/* ── Carga de datos ── */

Promise.all([
    fetch(`${BIRTHDAY_EXCEL_PATH}?v=${Date.now()}`).then(r => r.ok ? r.arrayBuffer() : null).catch(() => null),
    fetch(`${EXCEL_PATH}?v=${Date.now()}`).then(r => r.ok ? r.arrayBuffer() : null).catch(() => null)
]).then(([bufBirthday, bufVentas]) => {

    if(bufBirthday){
        const wb = XLSX.read(bufBirthday, { type:"array", cellDates:true });

        const wsProgramas = wb.Sheets["Programas"];
        if(wsProgramas){
            programasData = XLSX.utils.sheet_to_json(wsProgramas, { defval:"" })
                .filter(r => String(r["Cuenta"] || "").trim() !== "");
        }

        const wsIn30 = wb.Sheets["IN30"];
        if(wsIn30){
            in30Data = XLSX.utils.sheet_to_json(wsIn30, { defval:"" })
                .filter(r => String(campo(r, "Cuenta", "Cliente") || "").trim() !== "");
        }
    }

    if(bufVentas){
        const wb = XLSX.read(bufVentas, { type:"array" });
        const wsCumpl = wb.Sheets["Cump_mes_a_mes"];
        if(wsCumpl){
            cumplimientoMesData = XLSX.utils.sheet_to_json(wsCumpl, { defval:0 });
        }
    }

    renderPagina();

});

/* ── Render ── */

function renderPagina(){
    renderProgramasDominios();
    renderCardPrograma("programas", programasData.filter(p => true), {
        nombre: "Programas", icono: "🚀",
        descripcion: "Postulación y aprobación de programas de Seguridad, Copilot y Azure ante Microsoft.",
        postula: PROGRAMAS_POSTULA_GENERAL
    }, "programas");

    Object.entries(ACTIVIDADES_MS).forEach(([key, cfg]) => {
        const datos = in30Data.filter(r => normalize(String(campo(r, "Actividad") || "")).includes(cfg.clave));
        renderCardPrograma(key, datos, cfg, "actividad");
    });
}

function renderProgramasDominios(){
    const el = document.getElementById("programasDominiosDetalle");
    if(!el) return;

    const conteo = {};
    Object.keys(PROGRAMAS_DOMINIOS).forEach(d => conteo[d] = 0);
    programasData.forEach(p => {
        const estado = estadoActividad(p);
        if(estado !== "aprobado" && estado !== "postulado") return;
        const dom = dominioPrograma(p);
        if(dom in conteo) conteo[dom]++;
    });

    el.innerHTML = Object.entries(PROGRAMAS_DOMINIOS).map(([key, d]) => `
        <div class="ms-dominio-item">
            <span class="ms-dominio-nombre">${esc(d.label)}</span>
            <span class="ms-dominio-responsable">Responsable: ${esc(d.responsable)}</span>
            <span class="ms-dominio-count">${conteo[key]}</span>
        </div>
    `).join("");
}

function renderCardPrograma(key, datos, cfg, tipo){

    const totalEl        = document.getElementById(`ms_${key}_total`);
    const aprobadosEl    = document.getElementById(`ms_${key}_aprobados`);
    const postuladosEl   = document.getElementById(`ms_${key}_postulados`);
    const ofrecidosEl    = document.getElementById(`ms_${key}_ofrecidos`);
    const comercialesEl = document.getElementById(`ms_${key}_comerciales`);
    const rankingEl       = document.getElementById(`ms_${key}_ranking`);
    const descEl          = document.getElementById(`ms_${key}_desc`);
    const respEl          = document.getElementById(`ms_${key}_resp`);

    if(descEl) descEl.textContent = cfg.descripcion;
    if(respEl){
        respEl.innerHTML = tipo === "programas"
            ? `<div><strong>Postula (general):</strong> ${esc(cfg.postula)}</div>`
            : `<div><strong>Postula:</strong> ${esc(cfg.postula)}</div><div><strong>Ofrece:</strong> ${esc(cfg.ofrece)}</div>
               ${cfg.cumpleObjetivoMS ? `<div class="ms-objetivo-tag">🎯 Cuenta para objetivos MS en controles</div>` : ""}`;
    }

    const aprobados  = datos.filter(d => estadoActividad(d) === "aprobado").length;
    const postulados = datos.filter(d => estadoActividad(d) === "postulado").length;
    const ofrecidos   = datos.filter(d => estadoActividad(d) === "ofrecido").length;

    if(totalEl) totalEl.textContent = aprobados; // solo Aprobado cuenta para el objetivo
    if(aprobadosEl) aprobadosEl.textContent = aprobados;
    if(postuladosEl) postuladosEl.textContent = postulados;
    if(ofrecidosEl) ofrecidosEl.textContent = ofrecidos;

    const roster = getRosterComerciales();

    if(!roster.length && !datos.length){
        if(rankingEl) rankingEl.innerHTML = `<div class="programas-empty">Cargando datos…</div>`;
        return;
    }

    const conteoPorRoster = roster.map(() => ({ aprobado: 0, postulado: 0, ofrecido: 0 }));

    datos.forEach(d => {
        const estado = estadoActividad(d);
        if(!["aprobado","postulado","ofrecido"].includes(estado)) return;
        const idx = mejorIndiceRoster(campo(d, "Comercial"), roster);
        if(idx === -1) return;
        conteoPorRoster[idx][estado]++;
    });

    const ranking = roster.map((r, idx) => ({
        nombre: r.display,
        aprobado: conteoPorRoster[idx].aprobado,
        postulado: conteoPorRoster[idx].postulado,
        ofrecido: conteoPorRoster[idx].ofrecido
    }));

    ranking.forEach(r => r.total = r.aprobado + r.postulado + r.ofrecido);
    ranking.sort((a, b) => b.total - a.total || a.nombre.localeCompare(b.nombre));

    const comercialesVan = ranking.filter(r => r.total > 0).length;
    if(comercialesEl) comercialesEl.textContent = `${comercialesVan} / ${ranking.length || "—"}`;

    if(!rankingEl) return;

    if(!datos.length){
        rankingEl.innerHTML = `<div class="programas-empty">Aún no hay registros para "${esc(cfg.nombre)}" en el Excel (hoja IN30, columna Actividad = "${esc(cfg.nombre)}").</div>`;
        return;
    }

    const maxTotal = Math.max(1, ...ranking.map(r => r.total));

    rankingEl.innerHTML = ranking.map((r, i) => {
        const sinDatos = r.total === 0;
        return `
            <div class="ranking-item ${sinDatos ? "ranking-item-vacio" : ""}">
                <span class="ranking-pos">${i + 1}</span>
                <span class="ranking-name">${esc(r.nombre)}</span>
                <div class="ranking-bar-wrap" id="ms_${key}_bar${i}"></div>
                <span class="ranking-count">${r.total}</span>
            </div>
        `;
    }).join("");

    ranking.forEach((r, i) => {
        pintarBarraSegmentada(
            document.getElementById(`ms_${key}_bar${i}`),
            r.aprobado, r.postulado, r.ofrecido, maxTotal
        );
    });

}

/* Flip de las 4 tarjetas */
document.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll(".ms-flip-card").forEach(card => {
        card.addEventListener("click", () => card.classList.toggle("is-flipped"));
    });
});
