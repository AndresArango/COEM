/* ================================================
   Dashboard Comercial Occidente — APP.JS v4 (limpio)
   Parser robusto para ventas_occidente.xlsx
   ================================================ */

const EXCEL_PATH =
  window.location.pathname.includes("/pages/")
    ? "../ventas_occidente.xlsx"
    : "ventas_occidente.xlsx";

const BIRTHDAY_EXCEL_PATH =
  window.location.pathname.includes("/pages/")
    ? "../cumpleanos.xlsx"
    : "cumpleanos.xlsx";

// Variables globales
let tableData  = [];
let yearlyData = [];
let domainMonthData = [];
let domainYearData = [];
let domainTotalPct = 0;
let domainTotalUB = 0;
let totalAcumPct = 0;
let totalAcumUB = 0;
let cumplimientoMesData = [];
let cumpleanosData = [];
let paretoData = [];
const MESES_CUMP = ["2026 - ENERO","2026 - FEBRERO","2026 - MARZO","2026 - ABRIL","2026 - MAYO","2026 - JUNIO","2026 - JULIO","2026 - AGOSTO","2026 - SEPTIEMBRE","2026 - OCTUBRE","2026 - NOVIEMBRE","2026 - DICIEMBRE"];

document.addEventListener("DOMContentLoaded", loadExcelAuto);

document.addEventListener("DOMContentLoaded", () => {
    const flipCard = document.getElementById("cumplimientoFlipCard");
    if(flipCard){
        flipCard.addEventListener("click", () => {
            flipCard.classList.toggle("is-flipped");
        });
    }

    const dominiosFlip = document.getElementById("dominiosFlipCard");
    if(dominiosFlip){
        dominiosFlip.addEventListener("click", () => {
            dominiosFlip.classList.toggle("is-flipped");
        });
    }  

        const paretoFlip = document.getElementById("paretoFlipCard");
    if(paretoFlip){
        paretoFlip.addEventListener("click", () => {
            paretoFlip.classList.toggle("is-flipped");
        });
    }
});

loadBirthdaysIfNeeded();

function loadExcelAuto() {
  showLoading();

  fetch(`${EXCEL_PATH}?v=${Date.now()}`)
    .then(r => {
      if (!r.ok) throw new Error("no encontrado");
      return r.arrayBuffer();
    })
    .then(buf => parseExcel(buf))
    .catch(() => {
      useDemoData();
      renderAll();
    });
}

function showLoading() {
  const tb = document.getElementById("tableBody");
  if (tb) {
  tb.innerHTML = `
    <tr>
      <td colspan="7" style="text-align:center;padding:36px;color:rgba(220,240,220,0.5);">
        Cargando datos…
      </td>
    </tr>
  `;
}
}

/* ════════════════════════════
   DEMO DATA (fallback)
   ════════════════════════════ */
function useDemoData(){
   console.warn("No se encontró archivo Excel");
}

/* ════════════════════════════
   Función para obtener los meses visibles en la tabla de cumplimiento mes a mes
   ════════════════════════════ */
function getMesesVisiblesCumplimiento(){

    const mesActualIndex =
        new Date().getMonth();

    return MESES_CUMP.slice(
        0,
        mesActualIndex + 1
    );

}

function getCumplimientoClass(valor){

    const pct =
        Number(valor || 0) * 100;

    if (pct > 100)
        return "cump-supera";

    if (pct >= 100)
        return "cump-cumple";

    if (pct >= 70)
        return "cump-cerca";

    if (pct > 0)
        return "cump-no-cumple";

    return "cump-sin-dato";

}

function renderCumplimientoMesAMes(){

    const head =
        document.getElementById("cumplimientoMatrixHead");

    const body =
        document.getElementById("cumplimientoMatrixBody");

    if(!head || !body)
        return;

    head.innerHTML = "";
    body.innerHTML = "";

    const mesesVisibles =
        getMesesVisiblesCumplimiento();

const filas =
        cumplimientoMesData
            .filter(r => {

                const nombre =
                    String(
                        r["Etiquetas de fila"] || ""
                    ).trim();

                const esNombreValido =
                    nombre.split(/\s+/).length >= 2 &&
                    !normalize(nombre).startsWith("nn");

                return (
                    nombre !== "" &&
                    nombre !== "VALLE" &&
                    nombre !== "Total general" &&
                    esNombreValido
                );

            })
            .sort((a, b) => {

                const mesesA = Number(a["Cant Meses Cumplió cuota"] || 0);
                const mesesB = Number(b["Cant Meses Cumplió cuota"] || 0);

                return mesesB - mesesA;

            });

    head.innerHTML = `
        <tr>
            <th>Comercial</th>
            ${mesesVisibles.map(mes => `
                <th>
                    ${mes.replace("2026 - ", "").substring(0,3)}
                </th>
            `).join("")}
            <th>Meses cumplió</th>
        </tr>
    `;

    filas.forEach(row => {

        const nombre =
            String(
                row["Etiquetas de fila"] || ""
            ).trim();

        const mesesCumplio =
            Number(
                row["Cant Meses Cumplió cuota"] || 0
            );

        const tr =
            document.createElement("tr");

        tr.innerHTML = `
            <td class="cumplimiento-nombre">
                ${esc(shortName(nombre))}
            </td>

            ${mesesVisibles.map(mes => {

                const valor =
                    Number(row[mes] || 0);

                const clase =
                    getCumplimientoClass(valor);

                const pct =
                    Math.round(valor * 100);

                return `
                    <td title="${pct}%">
                        <span class="cump-dot ${clase}">${pct > 0 ? pct + "%" : ""}</span>
                    </td>
                `;

            }).join("")}

            <td class="cumplimiento-total">
                ${mesesCumplio}
            </td>
        `;

        body.appendChild(tr);

    });
}

function renderCumplimientoAcumulado(){

    const topEl = document.getElementById("topCumplimientoHistorico");
    const cumpliMesEl = document.getElementById("cumplieronMesVal");
    const todosEl = document.getElementById("todosCumplieronVal");
    const todosKpi = document.getElementById("todosCumplieronKpi");

    if(!topEl && !cumpliMesEl && !todosEl)
        return;

    const filas = cumplimientoMesData.filter(r => {
        const nombre = String(r["Etiquetas de fila"] || "").trim();
        const esNombreValido = nombre.split(/\s+/).length >= 2 && !normalize(nombre).startsWith("nn");
        return nombre !== "" && nombre !== "VALLE" && nombre !== "Total general" && esNombreValido;
    });

    const filasConDatos = filas.filter(r => Number(r["Total general"] || 0) > 0);

    const topAcumulado = [...filasConDatos].sort((a, b) => {
        const mesesA = Number(a["Cant Meses Cumplió cuota"] || 0);
        const mesesB = Number(b["Cant Meses Cumplió cuota"] || 0);
        return mesesB - mesesA;
    });

    if(topEl){
        topEl.innerHTML = "";
        topAcumulado.forEach((row, index) => {
            const nombre = String(row["Etiquetas de fila"] || "").trim();
            const meses = Number(row["Cant Meses Cumplió cuota"] || 0);
            const div = document.createElement("div");
            div.className = "top-cump-item";
            div.innerHTML = `
                <span class="top-cump-name">${index + 1}. ${esc(shortName(nombre))}</span>
                <span class="top-cump-value">${meses}</span>
            `;
            topEl.appendChild(div);
        });
    }

    const mesesVisibles = getMesesVisiblesCumplimiento();
    const mesActualKey = mesesVisibles[mesesVisibles.length - 1];

    const cumplieronMes = filas.filter(r => Number(r[mesActualKey] || 0) >= 1).length;
    const totalVendedores = filas.length;

    if(cumpliMesEl){
        cumpliMesEl.textContent = `${cumplieronMes} de ${totalVendedores}`;
    }

    if(todosEl){

        const notaEl = document.getElementById("todosCumplieronNota");

        const totalGeneralRow = cumplimientoMesData.find(r =>
            String(r["Etiquetas de fila"] || "").trim() === "Total general"
        );

        if(!totalGeneralRow){

            todosEl.textContent = "—";
            if(notaEl) notaEl.textContent = "No se encontró la fila 'Total general' en el Excel";

        } else {

            let vecesCumplido = 0;
            let mejorPct = -1;
            let mejorMes = "";

            mesesVisibles.forEach(mes => {
                const valor = Number(totalGeneralRow[mes] || 0);
                if(valor >= 1) vecesCumplido++;
                if(valor > mejorPct){
                    mejorPct = valor;
                    mejorMes = mes;
                }
            });

            if(vecesCumplido > 0){

                todosEl.textContent = `Sí (${vecesCumplido})`;
                if(notaEl) notaEl.textContent = "";
                if(todosKpi) todosKpi.style.borderLeftColor = "#00ff88";

            } else {

                todosEl.textContent = "Aún no";

                if(notaEl && mejorMes){
                    const mesLabel = mejorMes.replace("2026 - ", "");
                    const pctLabel = Math.round(mejorPct * 100);
                    notaEl.textContent = `Lo más cerca: ${mesLabel} (${pctLabel}%)`;
                }

                if(todosKpi) todosKpi.style.borderLeftColor = "#ff4d4d";

            }

        }

    }

}

/* ════════════════════════════
   PARSE EXCEL
   ════════════════════════════ */
function parseExcel(buffer) {
  try {
    const wb = XLSX.read(buffer, { type:"array" });

    /* ── Hoja 1: Tabla principal ── */
    const ws1  = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws1, { defval: 0, raw: true });

    if (rows.length > 0) {
      const keys = Object.keys(rows[0]);

      const findKey = (...candidates) => {
        for (const cand of candidates) {
          const found = keys.find(k => normalize(k) === normalize(cand));
          if (found) return found;
        }
        for (const cand of candidates) {
          const found = keys.find(k => normalize(k).includes(normalize(cand)));
          if (found) return found;
        }
        return null;
      };

      const kPos    = findKey("Pos","pos","posicion","posición","#");
      const kArea   = findKey("Area","Comercial","Área","area","área","nombre","departamento","equipo");
      const kPct    = findKey("% de cumplimiento","% Usuarios Activos","pct","porcentaje","usuarios","activos","user");
      const kInter  = findKey("Interacciones Screenshot","Cuota","interacciones screenshot","inter screenshot","screenshot interactions");
      const kGoles  = findKey("Goles","Utilidad Bruta","Oportunidades","goals","puntos","pts");
      const kEstado = findKey("Estado","Compromiso","status","clasificacion","clasificación");
      const kFalta  = findKey("Falta","falta","restante","diferencia");

       tableData = rows
        .filter(r => kArea && String(r[kArea]||"").trim() !== "")
        .filter(r => esNombrePersona(r[kArea]))
        .filter(r => {

      const pct = Number(r[kPct] || 0);
      const cuota = Number(r[kInter] || 0);

    return !(pct === 0 && cuota === 0);
})
        .map((r, i) => {
          let pct = kPct ? toPct(r[kPct]) : 0;

          return {
            pos:    kPos    ? (parseInt(r[kPos])  || i+1) : i+1,
            area:   kArea   ? String(r[kArea]).trim()      : `Área ${i+1}`,
            pct:    Math.round(pct),
            inter:  kInter  ? toNum(r[kInter])             : 0,
            goles:  kGoles  ? toNum(r[kGoles])             : 0,
            compromiso: kEstado ? toNum(r[kEstado]):0,
            falta: kFalta ? toNum(r[kFalta]) : 0
          };
        });

    }

    /* ── Hoja 2: Acumulado del año ── */
    if (wb.SheetNames[1]) {
      const ws2  = wb.Sheets[wb.SheetNames[1]];
      const raw2 = XLSX.utils.sheet_to_json(ws2, { defval:0 });
      const ultimaFila = raw2[raw2.length - 1];

totalAcumPct = toNum(ultimaFila["% Cump"]);

totalAcumUB = toNum(ultimaFila["Utilidad Bruta"]);

yearlyData = raw2
  .filter(r => String(r["Vendedor"] || "").trim() !== "")
  .filter(r => String(r["Vendedor"] || "").trim() !== "Total general")
  .filter(r => esNombrePersona(r["Vendedor"]))
  .filter(r => {

   const pct = toNum(r["% Cump"]);
   const cuota = toNum(r["Cuota"]);

   return !(pct === 0 && cuota === 0);

})
  .map(r => ({

      vendedor: String(r["Vendedor"] || ""),

      pct:
        Math.round(
            toNum(r["% Cump"]) * 100
        ),

      cuota:
        toNum(r["Cuota"]),

      utilidad:
        toNum(r["Utilidad Bruta"]),

      gap:
        toNum(r["Gap"])

  }));
  }

    /* ── Hoja 4: Dominios Mes ── */

if (wb.Sheets["Dominio_Mes"]) {

   const ws4 = wb.Sheets["Dominio_Mes"];
   const raw4 = XLSX.utils.sheet_to_json(ws4,{defval:0});

   domainMonthData = raw4
      .filter(r =>
         String(r["Etiquetas de fila"] || "").trim() !== ""
      )
      .filter(r =>
         !String(r["Etiquetas de fila"] || "").includes("Total")
      );
}

    /* ── Hoja 5: Dominio ACUMULADO ── */

if (wb.Sheets["Dominios_Acum"]) {

   const ws5 = wb.Sheets["Dominios_Acum"];
   const raw5 = XLSX.utils.sheet_to_json(ws5,{defval:0});

   const totalRow5 = raw5.find(r =>
      String(r["Etiquetas de fila"] || "").includes("Total")
   );

   domainTotalPct = totalRow5 ? toNum(totalRow5["% UB"]) : 0;
   domainTotalUB  = totalRow5 ? toNum(totalRow5["Utilidad Bruta"]) : 0;

   domainYearData = raw5
      .filter(r =>
         String(r["Etiquetas de fila"] || "").trim() !== ""
      )
      .filter(r =>
         !String(r["Etiquetas de fila"] || "").includes("Total")
      );
}

    /* ── Hoja: Pareto Acumulado ── */
    const paretoSheetName = wb.SheetNames.find(n => normalize(n).includes("pareto"));

    if(paretoSheetName){
        const wsP = wb.Sheets[paretoSheetName];
        const rawP = XLSX.utils.sheet_to_json(wsP, { defval: 0 });
        paretoData = rawP
            .filter(r => String(r["NOMBRE CLIENTE LEASING"] || "").trim() !== "")
            .filter(r => normalize(String(r["NOMBRE CLIENTE LEASING"] || "")) !== "totalgeneral");
    }

    /* ── Hoja 6: MES A MES ── */
if (wb.Sheets["Cump_mes_a_mes"]) {

    const wsCumpl =
        wb.Sheets["Cump_mes_a_mes"];

    const rawCumpl =
        XLSX.utils.sheet_to_json(
            wsCumpl,
            { defval: 0 }
        );

    cumplimientoMesData = rawCumpl;
}

  } catch(e) {
    console.error("Error parseExcel:", e);
    useDemoData();
  }

  renderAll();
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

function normalize(s) {
  return String(s||"").toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .replace(/[^a-z0-9]/g,"");
}

function findYtoYKey(rows){
    if(!rows || !rows.length) return "Y to Y Ub";
    const keys = Object.keys(rows[0]);
    const found = keys.find(k => normalize(k).includes("ytoy"));
    return found || "Y to Y Ub";
}

const DOMINIO_RESPONSABLES = {
    "cloudaiplatforms": "Omar Sanchez",
    "cybersecuritynetworking": "Santiago Vasquez",
    "datacentercloud": "Jaime Roman",
    "securityms": "Fernando Corella"
};

function domainDisplayName(nombre){
    const responsable = DOMINIO_RESPONSABLES[normalize(nombre)];
    if(responsable){
        return `${esc(nombre)}<span class="dominio-responsable">(${esc(responsable)})</span>`;
    }
    return esc(nombre);
}

function toNum(v) {

  if (v === null || v === undefined || v === "")
      return 0;

  if (typeof v === "number")
      return v;

  let str = String(v);

  str = str
      .replace(/\$/g,"")
      .replace(/\s/g,"")
      .replace(/\./g,"")
      .replace(",", ".");

  const n = Number(str);

  return isNaN(n) ? 0 : n;

}

function toPct(v) {

  if (v === null || v === undefined || v === "")
    return 0;

  const isString = typeof v === "string";
  const hasPercent = isString && v.includes("%");

  let n = toNum(v);

  if (isNaN(n))
    return 0;

  // Si viene como texto "228,60%", ya está en escala 228.60
  if (hasPercent)
    return n;

  // Si viene desde Excel como 0.4666 o 2.286, es porcentaje en formato decimal
  if (n > 0 && n <= 10)
    return n * 100;

  // Si ya viene como 46.66 o 228.60, se deja igual
  return n;

}

function formatCurrency(value){
  return new Intl.NumberFormat(
    "es-CO",
    {
      style: "currency",
      currency: "COP",
      maximumFractionDigits: 0
    }
  ).format(value);
}

/* ════════════════════════════
   RENDER ALL
   ════════════════════════════ */
function renderAll() {
  renderTable();
  renderYearlyTable();
  renderYearlySubtitle();
  renderDomainSubtitle();
  renderDomainTables();
  renderDominiosRankingFlip();
  renderDomainCardSummary();
  renderMainTitle();
  renderExecutiveSummary();
  renderCumplimientoMesAMes();
  renderCumplimientoAcumulado();
  renderParetoCuentas();
  renderKPIs();
  //adjustLayoutByTeamCount();
}

/* ════════════════════════════
   TABLA
   ════════════════════════════ */
function renderTable() {
  const tbody = document.getElementById("tableBody");
  if (!tbody) return;
  tbody.innerHTML = "";

  const sorted = [...tableData].sort((a, b) => {

    // 1. % Cumplimiento
    if ((b.pct || 0) !== (a.pct || 0))
        return (b.pct || 0) - (a.pct || 0);

    // 2. Utilidad Bruta
    if ((b.goles || 0) !== (a.goles || 0))
        return (b.goles || 0) - (a.goles || 0);

    // 3. Compromiso
    if ((b.compromiso || 0) !== (a.compromiso || 0))
        return (b.compromiso || 0) - (a.compromiso || 0);

    // 4. Falta (más cercano a cero gana)
    return Math.abs(a.falta || 0) - Math.abs(b.falta || 0);

});

  sorted.forEach((row, idx) => {
    const tr = document.createElement("tr");
    const pct = row.pct;

   let barClass = "fill-red";
   if (pct >= 100)
    barClass = "fill-green-dark";
   else if (pct >= 70)
    barClass = "fill-green";
   else if (pct >= 35)
    barClass = "fill-yellow";
   else if (pct >= 15)
    barClass = "fill-orange";

  if (idx === 0){

    tr.classList.add("rank-1");

}
else if (pct >= 30){

    tr.classList.add("rank-green");

}
else if (idx === 1){

    tr.classList.add("rank-2");

}
else if (idx === 2){

    tr.classList.add("rank-3");

}
else if (idx <= 4){

    tr.classList.add("rank-4");

}
else if (idx <= 6){

    tr.classList.add("rank-5");

}
else{

    tr.classList.add("rank-6");

}

    let medal = "";
    if (idx === 0) { medal="🥇";}
    else if (idx === 1) { medal="🥈";}
    else if (idx === 2) { medal="🥉";}

    const posNum = idx + 1;

    tr.innerHTML = `
      <td><div class="pos-badge">${posNum}</div></td>
      <td class="area-name">${esc(shortName(row.area))}</td>
      <td>
        <div class="pct-bar-wrap">
          <div class="pct-bar-bg">
            <div class="pct-bar-fill ${barClass}" style="width:0%" data-target="${pct}%"></div>
          </div>
          <span class="pct-num">${medal} ${pct}%</span>
        </div>
      </td>
      <td>
    <span class="currency-cell">
        ${formatCurrency(row.inter)}
    </span>
</td>

<td>
    <span class="currency-cell">
        ${formatCurrency(row.goles)}
    </span>
</td>

<td>
    ${formatCurrency(row.compromiso)}
</td>

<td>
    <span class="falta-pendiente">
        ${formatCurrency(row.falta)}
    </span>
</td>

    `;
    tbody.appendChild(tr);
  });

  requestAnimationFrame(() => setTimeout(() => {
    document.querySelectorAll(".pct-bar-fill").forEach(b => b.style.width = b.dataset.target);
  }, 180));
}

/* ════════════════════════════
   TABLA ACUMULADO DEL AÑO
   ════════════════════════════ */

function renderYearlyTable() {

  const tbody = document.getElementById("yearlyTableBody");

  if (!tbody) return;

  tbody.innerHTML = "";

  const sorted = [...yearlyData].sort((a, b) => {

    // 1. % Cumplimiento acumulado
    if ((b.pct || 0) !== (a.pct || 0))
      return (b.pct || 0) - (a.pct || 0);

    // 2. Utilidad acumulada
    if ((b.utilidad || 0) !== (a.utilidad || 0))
      return (b.utilidad || 0) - (a.utilidad || 0);

    // 3. Cuota acumulada
    if ((b.cuota || 0) !== (a.cuota || 0))
      return (b.cuota || 0) - (a.cuota || 0);

    // 4. Gap positivo gana sobre gap negativo
    return (b.gap || 0) - (a.gap || 0);

  });

  sorted.forEach((row, idx) => {

    const tr = document.createElement("tr");

    const pct = row.pct || 0;

    let barClass = "fill-red";

    if (pct >= 100)
      barClass = "fill-green-dark";
    else if (pct >= 70)
      barClass = "fill-green";
    else if (pct >= 35)
      barClass = "fill-yellow";
    else if (pct >= 15)
      barClass = "fill-orange";

    if (idx === 0) {

      tr.classList.add("rank-1");

    } else if (pct >= 30) {

      tr.classList.add("rank-green");

    } else if (idx === 1) {

      tr.classList.add("rank-2");

    } else if (idx === 2) {

      tr.classList.add("rank-3");

    } else if (idx <= 4) {

      tr.classList.add("rank-4");

    } else if (idx <= 6) {

      tr.classList.add("rank-5");

    } else {

      tr.classList.add("rank-6");

    }

    const medal =
      idx === 0 ? "🥇" :
      idx === 1 ? "🥈" :
      idx === 2 ? "🥉" : "";

const gapClass =
   row.gap >= 0
      ? "gap-positivo"
      : "gap-negativo";

    tr.innerHTML = `
      <td>
        <div class="pos-badge">
          ${idx + 1}
        </div>
      </td>

      <td class="area-name">
        ${esc(shortName(row.vendedor))}
      </td>

      <td>
        <div class="pct-bar-wrap">
          <div class="pct-bar-bg">
            <div
              class="pct-bar-fill ${barClass}"
              style="width:0%"
              data-target="${Math.min(pct, 100)}%">
            </div>
          </div>

          <span class="pct-num">
            ${medal} ${pct}%
          </span>
        </div>
      </td>

      <td>
        <span class="currency-cell">
          ${formatCurrency(row.cuota)}
        </span>
      </td>

      <td>
        <span class="currency-cell">
          ${formatCurrency(row.utilidad)}
        </span>
      </td>

      <td>
        <span class="${row.gap >= 0 ? 'gap-positivo' : 'gap-negativo'}">
          <span class="${gapClass}">
    ${formatCurrency(row.gap)}</span>
        </span>
      </td>
    `;

    tbody.appendChild(tr);

  });

  requestAnimationFrame(() =>
    setTimeout(() => {

      document
        .querySelectorAll("#yearlyTable .pct-bar-fill")
        .forEach(b => {
          b.style.width = b.dataset.target;
        });

    }, 180)
  );

}

function renderYearlySubtitle() {

  const meses = [
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE"
  ];

  const el = document.getElementById("yearlySubtitle");

  if (!el) return;

  const mesActual = meses[new Date().getMonth()];

  el.textContent =
  `🏆 ACUMULADO ENERO - ${mesActual}`;

}

function renderDomainSubtitle(){

    const meses = [
      "ENERO",
      "FEBRERO",
      "MARZO",
      "ABRIL",
      "MAYO",
      "JUNIO",
      "JULIO",
      "AGOSTO",
      "SEPTIEMBRE",
      "OCTUBRE",
      "NOVIEMBRE",
      "DICIEMBRE"
    ];

    const el =
      document.getElementById(
        "domainSubtitle"
      );

    if(!el) return;

    el.textContent =
      `Dominios Enero - ${
        meses[new Date().getMonth()]
      }`;

}

function renderDomainTables(){

  const monthBody =
    document.getElementById("domainMonthBody");

  const yearBody =
    document.getElementById("domainYearBody");

  if(!monthBody && !yearBody)
    return;

  if(monthBody){

    monthBody.innerHTML = "";

    const sortedMonth =
      [...domainMonthData]
        .sort((a,b) => {
          const pctA = Number(a["% UB"] || 0);
          const pctB = Number(b["% UB"] || 0);

          if (pctB !== pctA) return pctB - pctA;

          const ubA = Number(a["Utilidad Bruta"] || 0);
          const ubB = Number(b["Utilidad Bruta"] || 0);

          return ubB - ubA;
        });

    const yToYKeyMonth = findYtoYKey(domainMonthData);

    sortedMonth.forEach((row,idx) => {

      const tr =
        document.createElement("tr");

      const medal =
        idx === 0 ? "🥇" :
        idx === 1 ? "🥈" :
        idx === 2 ? "🥉" : "";

      tr.innerHTML = `
        <td>${idx + 1}</td>

        <td class="domain-name-cell">
          ${medal}
          ${domainDisplayName(row["Etiquetas de fila"])}
        </td>

        <td class="${Number(row["% UB"] || 0) < 0 ? "gap-negativo" : ""}">
          ${Math.round(Number(row["% UB"] || 0) * 100)}%
        </td>

        <td class="${Number(row["Utilidad Bruta"] || 0) < 0 ? "gap-negativo" : ""}">
          ${formatCurrency(Number(row["Utilidad Bruta"] || 0))}
        </td>

        <td class="${
          toNum(row[yToYKeyMonth]) >= 0
            ? "gap-positivo"
            : "gap-negativo"
        }">
          ${formatCurrency(toNum(row[yToYKeyMonth]))}
        </td>
      `;

      monthBody.appendChild(tr);

    });

  }

  if(yearBody){

    yearBody.innerHTML = "";

    const sortedYear =
      [...domainYearData]
        .sort((a,b) => {
          const pctA = Number(a["% UB"] || 0);
          const pctB = Number(b["% UB"] || 0);

          if (pctB !== pctA) return pctB - pctA;

          const ubA = Number(a["Utilidad Bruta"] || 0);
          const ubB = Number(b["Utilidad Bruta"] || 0);

          return ubB - ubA;
        });

    const yToYKeyYear = findYtoYKey(domainYearData);

    sortedYear.forEach((row,idx) => {

      const tr =
        document.createElement("tr");

      const medal =
        idx === 0 ? "🥇" :
        idx === 1 ? "🥈" :
        idx === 2 ? "🥉" : "";

      tr.innerHTML = `
        <td>${idx + 1}</td>

        <td class="domain-name-cell">
          ${medal}
          ${domainDisplayName(row["Etiquetas de fila"])}
        </td>

        <td class="${Number(row["% UB"] || 0) < 0 ? "gap-negativo" : ""}">
          ${Math.round(Number(row["% UB"] || 0) * 100)}%
        </td>

        <td class="${Number(row["Utilidad Bruta"] || 0) < 0 ? "gap-negativo" : ""}">
          ${formatCurrency(Number(row["Utilidad Bruta"] || 0))}
        </td>

        <td class="${
          toNum(row[yToYKeyYear]) >= 0
            ? "gap-positivo"
            : "gap-negativo"
        }">
          ${formatCurrency(toNum(row[yToYKeyYear]))}
        </td>
      `;

      yearBody.appendChild(tr);

    });

  }

}

function renderExecutiveSummary(){

  const monthEl =
    document.getElementById("monthSummary");

  const yearEl =
    document.getElementById("yearSummary");

  if(!monthEl && !yearEl)
    return;

  const cuotaMes =
    tableData.reduce(
      (s, r) => s + (r.inter || 0),
      0
    );

  const ventasMes =
    tableData.reduce(
      (s, r) => s + (r.goles || 0),
      0
    );

  const cumplimientoMes =
    cuotaMes > 0
      ? Math.round((ventasMes / cuotaMes) * 100)
      : 0;

  const topMes =
    [...tableData]
      .sort((a, b) => {

        if ((b.pct || 0) !== (a.pct || 0))
          return (b.pct || 0) - (a.pct || 0);

        return (b.goles || 0) - (a.goles || 0);

      })[0];

  const ventasYear =
    totalAcumUB;

  const cumplimientoYear =
    Math.round(totalAcumPct * 100);

  const topYear =
    [...yearlyData]
      .sort((a, b) => {

        if ((b.pct || 0) !== (a.pct || 0))
          return (b.pct || 0) - (a.pct || 0);

        return (b.utilidad || 0) - (a.utilidad || 0);

      })[0];

  if(monthEl && topMes){

    monthEl.innerHTML = `

      <div class="summary-main">

        <div class="summary-pct">
          ${cumplimientoMes}%
        </div>

        <div class="summary-text">
          Cumplimiento
        </div>

      </div>

      <div class="summary-sales">

        <div class="summary-sales-value">
          ${formatCurrency(ventasMes)}
        </div>

        <div class="summary-sales-label">
          Ventas del Mes
        </div>

      </div>

      <div class="summary-mvp">

        <div class="summary-name">
          ${shortName(topMes.area)}
        </div>

        <div class="summary-name-pct">
          ${topMes.pct}%
        </div>

      </div>

    `;

  }

  if(yearEl && topYear){

    yearEl.innerHTML = `

      <div class="summary-main">

        <div class="summary-pct">
          ${cumplimientoYear}%
        </div>

        <div class="summary-text">
          Cumplimiento
        </div>

      </div>

      <div class="summary-sales">

        <div class="summary-sales-value">
          ${formatCurrency(ventasYear)}
        </div>

        <div class="summary-sales-label">
          Ventas Acumuladas
        </div>

      </div>

      <div class="summary-mvp">

        <div class="summary-name">
          ${shortName(topYear.vendedor)}
        </div>

        <div class="summary-name-pct">
          ${topYear.pct}%
        </div>

      </div>

    `;

  }

}

function renderMainTitle(){

    const meses = [
        "ENERO",
        "FEBRERO",
        "MARZO",
        "ABRIL",
        "MAYO",
        "JUNIO",
        "JULIO",
        "AGOSTO",
        "SEPTIEMBRE",
        "OCTUBRE",
        "NOVIEMBRE",
        "DICIEMBRE"
    ];

    const el =
        document.getElementById(
            "mainTitle"
        );

    if(!el) return;

    const mesActual =
        meses[new Date().getMonth()];

    el.innerHTML =
        `VENTAS <span class="accent">${mesActual}</span> SUROCCIDENTE`;

}

function shortName(nombre){

    const partes = nombre.trim().split(/\s+/);

    if(partes.length === 2){
        return `${partes[1]} ${partes[0]}`;
    }

    if(partes.length >= 3){

        const primerNombre = partes[partes.length - 2];
        const primerApellido = partes[0];

        return `${primerNombre} ${primerApellido}`;

    }

    return nombre;

}

/* ════════════════════════════
   KPIs
   ════════════════════════════ */

function renderKPIs() {
  const el = document.getElementById("kpiCards");
  if (!el) return;
  el.innerHTML = "";
  
  el.style.display = "grid";
  el.style.gridTemplateColumns = "repeat(auto-fit, minmax(180px, 1fr))";
  el.style.gap = "0.7rem";
  el.style.alignItems = "stretch";

  if (!tableData.length) {
    el.innerHTML = '<div class="kpi"><span class="kpi-label">Sin datos</span><span class="kpi-val">—</span></div>';
    return;
  }

const ventasMes =
    tableData.reduce(
        (s, r) => s + (r.goles || 0),
        0
    );

const cumplimientoPromedio =
    Math.round(
        tableData.reduce(
            (s, r) => s + (r.pct || 0),
            0
        ) / tableData.length
    );

const top =
    [...tableData]
        .sort((a, b) => (b.pct || 0) - (a.pct || 0))[0];

const n = tableData.length;

  const kpisData = [

  {
    label: "🏆 Cumplimiento del Mes",
    val: `${cumplimientoPromedio}%`
  },

  {
    label: "💰 Ventas del Mes",
    val: formatCurrency(ventasMes)
  },

  {
    label: "🥇 MVP DEL MES",
    val: shortName(top.area)
  },

  {
    label: "👥 Vendedores",
    val: `${n}`
  }

];

  kpisData.forEach(k => {
    const d = document.createElement("div");
    d.className = "kpi";
    d.innerHTML = `
      <span class="kpi-label">${k.label}</span>
      <span class="kpi-val">${k.val}</span>
    `;
    el.appendChild(d);
  });
}

function renderDominiosRankingFlip(){

    const listEl = document.getElementById("topDominiosAcumulado");

    if(!listEl)
        return;

    listEl.innerHTML = "";

    const ranking = [...domainYearData].sort((a,b) => {

        const pctA = Number(a["% UB"] || 0);
        const pctB = Number(b["% UB"] || 0);

        if (pctB !== pctA) return pctB - pctA;

        const ubA = Number(a["Utilidad Bruta"] || 0);
        const ubB = Number(b["Utilidad Bruta"] || 0);

        return ubB - ubA;

    });

    ranking.forEach((row, index) => {

        const nombre = String(row["Etiquetas de fila"] || "").trim();

        const pct = Math.round(Number(row["% UB"] || 0) * 100);

        const div = document.createElement("div");
        div.className = "top-cump-item";
        div.innerHTML = `
            <span class="top-cump-name">${index + 1}. ${domainDisplayName(nombre)}</span>
            <span class="top-cump-value">${pct}%</span>
        `;
        listEl.appendChild(div);

    });

}

function renderDomainCardSummary(){

    const el = document.getElementById("domainCardSummary");
    if(!el) return;

    const top = [...domainYearData].sort((a,b) => {

        const pctA = Number(a["% UB"] || 0);
        const pctB = Number(b["% UB"] || 0);

        if (pctB !== pctA) return pctB - pctA;

        return Number(b["Utilidad Bruta"] || 0) - Number(a["Utilidad Bruta"] || 0);

    })[0];

    if(!top){
        el.innerHTML = "";
        return;
    }

    const nombre = String(top["Etiquetas de fila"] || "").trim();
    const pct = Math.round(Number(top["% UB"] || 0) * 100);

    el.innerHTML = `
      <div class="summary-main">
        <div class="summary-pct">${Math.round(domainTotalPct * 100)}%</div>
        <div class="summary-text">Cumplimiento</div>
      </div>
      <div class="summary-sales">
        <div class="summary-sales-value">${formatCurrency(domainTotalUB)}</div>
        <div class="summary-sales-label">Utilidad Bruta Acumulada</div>
      </div>
      <div class="summary-mvp">
        <div class="summary-name">${domainDisplayName(nombre)}</div>
        <div class="summary-name-pct">${pct}%</div>
      </div>
    `;

}

function loadBirthdaysIfNeeded(){

    const el = document.getElementById("cumpleanosCardList");
    if(!el) return;

    fetch(`${BIRTHDAY_EXCEL_PATH}?v=${Date.now()}`)
        .then(r => {
            if(!r.ok) throw new Error("no encontrado");
            return r.arrayBuffer();
        })
        .then(buf => {
            const wb = XLSX.read(buf, { type:"array", cellDates:true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });
            cumpleanosData = rows;
            renderCumpleanosCard();
        })
        .catch(() => {
            console.warn("No se encontró el archivo de cumpleaños");
            el.innerHTML = "";
        });

}

function extractDayMonth(value){

    if(value instanceof Date && !isNaN(value)){
        return { day: value.getDate(), month: value.getMonth() + 1 };
    }

    if(typeof value === "number"){
        const date = new Date(Math.round((value - 25569) * 86400 * 1000));
        if(!isNaN(date)) return { day: date.getDate(), month: date.getMonth() + 1 };
    }

    if(typeof value === "string"){
        const match = value.match(/(\d{1,2})[\/\-](\d{1,2})/);
        if(match){
            return { day: parseInt(match[1]), month: parseInt(match[2]) };
        }
    }

    return null;

}

const ZODIAC_SIGNS = {
    aries:       { name:"Aries",       emoji:"♈", element:"fuego"  },
    tauro:       { name:"Tauro",       emoji:"♉", element:"tierra" },
    geminis:     { name:"Géminis",     emoji:"♊", element:"aire"   },
    cancer:      { name:"Cáncer",      emoji:"♋", element:"agua"   },
    leo:         { name:"Leo",         emoji:"♌", element:"fuego"  },
    virgo:       { name:"Virgo",       emoji:"♍", element:"tierra" },
    libra:       { name:"Libra",       emoji:"♎", element:"aire"   },
    escorpio:    { name:"Escorpio",    emoji:"♏", element:"agua"   },
    sagitario:   { name:"Sagitario",   emoji:"♐", element:"fuego"  },
    capricornio: { name:"Capricornio", emoji:"♑", element:"tierra" },
    acuario:     { name:"Acuario",     emoji:"♒", element:"aire"   },
    piscis:      { name:"Piscis",      emoji:"♓", element:"agua"   }
};

const ELEMENTOS_AFINES = {
    fuego:  ["fuego","aire"],
    aire:   ["aire","fuego"],
    tierra: ["tierra","agua"],
    agua:   ["agua","tierra"]
};

const ELEMENTOS_CHOQUE = {
    fuego:  ["agua"],
    agua:   ["fuego"],
    tierra: ["aire"],
    aire:   ["tierra"]
};

function getZodiacSign(day, month){

    const md = month * 100 + day;

    if (md >= 321 && md <= 419) return ZODIAC_SIGNS.aries;
    if (md >= 420 && md <= 520) return ZODIAC_SIGNS.tauro;
    if (md >= 521 && md <= 620) return ZODIAC_SIGNS.geminis;
    if (md >= 621 && md <= 722) return ZODIAC_SIGNS.cancer;
    if (md >= 723 && md <= 822) return ZODIAC_SIGNS.leo;
    if (md >= 823 && md <= 922) return ZODIAC_SIGNS.virgo;
    if (md >= 923 && md <= 1022) return ZODIAC_SIGNS.libra;
    if (md >= 1023 && md <= 1121) return ZODIAC_SIGNS.escorpio;
    if (md >= 1122 && md <= 1221) return ZODIAC_SIGNS.sagitario;
    if (md >= 1222 || md <= 119) return ZODIAC_SIGNS.capricornio;
    if (md >= 120 && md <= 218) return ZODIAC_SIGNS.acuario;
    return ZODIAC_SIGNS.piscis;

}

function renderCumpleanosCard(){

    const el = document.getElementById("cumpleanosCardList");
    if(!el) return;

    const keys = cumpleanosData.length ? Object.keys(cumpleanosData[0]) : [];

    const findKey = (...candidates) => {
        for (const cand of candidates) {
            const found = keys.find(k => normalize(k) === normalize(cand));
            if (found) return found;
        }
        for (const cand of candidates) {
            const found = keys.find(k => normalize(k).includes(normalize(cand)));
            if (found) return found;
        }
        return null;
    };

    const kNombre = findKey("Nombre", "Nombre Completo", "Empleado", "Comercial");
    const kFecha  = findKey("Fecha de Nacimiento", "Fecha Nacimiento", "Cumpleaños", "Fecha");

    // Todo el equipo, con su signo ya calculado (para buscar afinidades)
    const equipoCompleto = cumpleanosData
        .map(r => {
            const nombre = kNombre ? String(r[kNombre] || "").trim() : "";
            const fecha = kFecha ? extractDayMonth(r[kFecha]) : null;
            if(!nombre || !fecha) return null;
            return { nombre, signo: getZodiacSign(fecha.day, fecha.month) };
        })
        .filter(Boolean);

    const mesActual = new Date().getMonth() + 1;

    const cumples = equipoCompleto === null ? [] :
        cumpleanosData
            .map(r => {
                const nombre = kNombre ? String(r[kNombre] || "").trim() : "";
                const fecha = kFecha ? extractDayMonth(r[kFecha]) : null;
                return { nombre, fecha };
            })
            .filter(c => c.nombre && c.fecha && c.fecha.month === mesActual)
            .sort((a, b) => a.fecha.day - b.fecha.day);

    if(!cumples.length){
        el.innerHTML = `<div class="cumple-empty">Nadie cumple años este mes 🎉</div>`;
        return;
    }

    el.innerHTML = cumples.map(c => {

        const signo = getZodiacSign(c.fecha.day, c.fecha.month);

        const elementosAfines = ELEMENTOS_AFINES[signo.element] || [];
        const elementosChoque = ELEMENTOS_CHOQUE[signo.element] || [];

        const afines = equipoCompleto
            .filter(p => p.nombre !== c.nombre && elementosAfines.includes(p.signo.element))
            .map(p => shortName(p.nombre));

        const choques = equipoCompleto
            .filter(p => p.nombre !== c.nombre && elementosChoque.includes(p.signo.element))
            .map(p => shortName(p.nombre));

        const lineaAfin = afines.length
            ? `<div class="cumple-afinidad">❤️ Afinidad: ${esc(afines.join(", "))}</div>`
            : `<div class="cumple-fun">No conecta con nadie 😅</div>`;

        const lineaChoque = choques.length
            ? `<div class="cumple-noafinidad">⚡ No conecta con: ${esc(choques.join(", "))}</div>`
            : `<div class="cumple-fun">Tranquilo, no tiene con quién no la lleve bien 😌</div>`;

        return `
            <div class="top-cump-item cumple-item">
                <span class="top-cump-name">${signo.emoji} ${c.fecha.day} — ${esc(shortName(c.nombre))}</span>
                ${lineaAfin}
                ${lineaChoque}
            </div>
        `;

    }).join("");

}

function renderParetoCuentas(){

    const body = document.getElementById("paretoTableBody");
    const perdidasEl = document.getElementById("cuentasPerdidasVal");
    const nuevasEl = document.getElementById("cuentasNuevasVal");
    const vendidasEl = document.getElementById("cuentasVendidasVal");

    if(!body && !perdidasEl && !nuevasEl && !vendidasEl) return;

    const cuentas = paretoData
        .map(r => ({
            nombre: String(r["NOMBRE CLIENTE LEASING"] || "").trim(),
            v2025: toNum(r["2025"]),
            v2026: toNum(r["2026"])
        }))
        .filter(c => c.nombre !== "");

    if(body){

        body.innerHTML = "";

        const top15 = [...cuentas]
            .sort((a, b) => b.v2026 - a.v2026)
            .slice(0, 15);

        top15.forEach((c, idx) => {

            let difHtml;

            if(c.v2025 <= 0 && c.v2026 > 0){
                difHtml = `<span class="gap-positivo">Nueva</span>`;
            } else if(c.v2025 > 0){
                const pct = Math.round(((c.v2026 - c.v2025) / c.v2025) * 100);
                difHtml = `<span class="${pct >= 0 ? "gap-positivo" : "gap-negativo"}">${pct}%</span>`;
            } else {
                difHtml = `<span class="gap-negativo">—</span>`;
            }

            const tr = document.createElement("tr");
            tr.innerHTML = `
                <td>${idx + 1}</td>
                <td class="domain-name-cell">${esc(c.nombre)}</td>
                <td class="currency-cell">${formatCurrency(c.v2026)}</td>
                <td>${difHtml}</td>
            `;
            body.appendChild(tr);

        });

    }

    const perdidas = cuentas.filter(c => c.v2025 > 0 && c.v2026 <= 0).length;
    const nuevas = cuentas.filter(c => c.v2025 <= 0 && c.v2026 > 0).length;
    const vendidas = cuentas.filter(c => c.v2026 > 0).length;

    if(perdidasEl) perdidasEl.textContent = perdidas;
    if(nuevasEl) nuevasEl.textContent = nuevas;
    if(vendidasEl) vendidasEl.textContent = vendidas;

}

function esc(s){ return String(s||"—").replace(/[&<>]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[m]||m)); }

/* ════════════════════════════
   AJUSTAR LAYOUT SEGÚN CANTIDAD DE EQUIPOS
   (Función desactivada — su llamada está comentada en renderAll().
    Se conserva por si se decide reactivar más adelante.)
   ════════════════════════════ */
function adjustLayoutByTeamCount() {
  const dashboard = document.querySelector('.dashboard-bottom');
  if (!dashboard) return;
  
  const teamCount = tableData.length;
  const activityCard = document.querySelector('.card-activity');
  const golesCard = document.querySelector('.card-goles');
  const kpisCard = document.querySelector('.card-kpis');
  
  if (!activityCard || !golesCard || !kpisCard) return;
  
  if (teamCount > 5) {
    dashboard.style.display = "grid";
    dashboard.style.gridTemplateColumns = "1fr";
    dashboard.style.gap = "1rem";
    
    let wrapper = document.querySelector('.goles-kpis-wrapper');
    if (!wrapper) {
      wrapper = document.createElement('div');
      wrapper.className = 'goles-kpis-wrapper';
      wrapper.style.display = "grid";
      wrapper.style.gridTemplateColumns = "1fr 1fr";
      wrapper.style.gap = "1rem";
      
      const golesParent = golesCard.parentNode;
      const kpisParent = kpisCard.parentNode;
      
      if (golesParent === dashboard && kpisParent === dashboard) {
        dashboard.removeChild(golesCard);
        dashboard.removeChild(kpisCard);
        wrapper.appendChild(golesCard);
        wrapper.appendChild(kpisCard);
        dashboard.appendChild(wrapper);
      }
    }
  } else {
    dashboard.style.display = "grid";
    dashboard.style.gridTemplateColumns = "1fr 1fr 1fr";
    dashboard.style.gap = "1rem";
    
    const wrapper = document.querySelector('.goles-kpis-wrapper');
    if (wrapper && wrapper.parentNode === dashboard) {
      const golesCardInside = wrapper.querySelector('.card-goles');
      const kpisCardInside = wrapper.querySelector('.card-kpis');
      if (golesCardInside && kpisCardInside) {
        dashboard.insertBefore(golesCardInside, wrapper);
        dashboard.insertBefore(kpisCardInside, wrapper);
      }
      wrapper.remove();
    }
  }
}
