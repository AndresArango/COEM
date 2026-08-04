/* ================================================
   Dashboard Comercial Occidente — APP.JS v3
   Parser robusto para copa_copilot_datos.xlsx
   ================================================ */

const AREA_COLORS = ["#e8c84a","#5ab4d6","#a87cbe","#d4863a","#c05555","#6ec87a","#4ab8c8","#f39c12","#1abc9c","#e74c3c","#3498db","#9b59b6","#2ecc71"];

// Variables globales
let tableData  = [];
let yearlyData = [];
let domainMonthData = [];
let domainYearData = [];
let totalAcumPct = 0;
let totalAcumUB = 0;
let cumplimientoMesData = [];
// LEGACY - REVISAR ELIMINACION
let weeklyData = { labels:[], series:[] };

document.addEventListener("DOMContentLoaded", loadExcelAuto);

function loadExcelAuto() {
  showLoading();
  fetch(`ventas_occidente.xlsx?v=${Date.now()}`)
    .then(r => { if (!r.ok) throw new Error("no encontrado"); return r.arrayBuffer(); })
    .then(buf => parseExcel(buf))
    .catch(() => { useDemoData(); renderAll(); });
}

function showLoading() {
  const tb = document.getElementById("tableBody");
  if (tb) tb.innerHTML = `<tr><td colspan="6" style="text-align:center;padding:36px;color:rgba(220,240,220,0.5);">Cargando datos…</td></tr>`;
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

    const topEl =
        document.getElementById("topCumplimientoHistorico");

    if(!head || !body || !topEl)
        return;

    head.innerHTML = "";
    body.innerHTML = "";
    topEl.innerHTML = "";

    const mesesVisibles =
        getMesesVisiblesCumplimiento();

    const filas =
        cumplimientoMesData
            .filter(r => {

                const nombre =
                    String(
                        r["Etiquetas de fila"] || ""
                    ).trim();

                return (
                    nombre !== "" &&
                    nombre !== "VALLE" &&
                    nombre !== "Total general"
                );

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
                ${esc(nombre)}
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
                        <span class="cump-dot ${clase}">
                        </span>
                    </td>
                `;

            }).join("")}

            <td class="cumplimiento-total">
                ${mesesCumplio}
            </td>
        `;

        body.appendChild(tr);

    });

    const top5 =
        [...filas]
            .sort((a, b) => {

                const mesesB =
                    Number(
                        b["Cant Meses Cumplió cuota"] || 0
                    );

                const mesesA =
                    Number(
                        a["Cant Meses Cumplió cuota"] || 0
                    );

                if(mesesB !== mesesA)
                    return mesesB - mesesA;

                const totalB =
                    Number(
                        b["Total general"] || 0
                    );

                const totalA =
                    Number(
                        a["Total general"] || 0
                    );

                return totalB - totalA;

            })
            .slice(0, 5);

    top5.forEach((row, index) => {

        const nombre =
            String(
                row["Etiquetas de fila"] || ""
            ).trim();

        const mesesCumplio =
            Number(
                row["Cant Meses Cumplió cuota"] || 0
            );

        const div =
            document.createElement("div");

        div.className =
            "top-cump-item";

        div.innerHTML = `
            <span class="top-cump-name">
                ${index + 1}. ${esc(shortName(nombre))}
            </span>

            <span class="top-cump-value">
                ${mesesCumplio}
            </span>
        `;

        topEl.appendChild(div);

    });

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

      /* ── Constantes para mes a mes ── */
      const MESES_CUMP = ["2026 - ENERO","2026 - FEBRERO","2026 - MARZO","2026 - ABRIL","2026 - MAYO","2026 - JUNIO","2026 - JULIO","2026 - AGOSTO","2026 - SEPTIEMBRE","2026 - OCTUBRE","2026 - NOVIEMBRE","2026 - DICIEMBRE"];
      
      tableData = rows
        .filter(r => kArea && String(r[kArea]||"").trim() !== "")
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

    /* ── Hoja 2: Actividad Semanal ── */
    if (wb.SheetNames[1]) {
      const ws2  = wb.Sheets[wb.SheetNames[1]];
      const raw2 = XLSX.utils.sheet_to_json(ws2, { defval:0 });
      const ultimaFila = raw2[raw2.length - 1];

totalAcumPct =
   Number(
      ultimaFila["% Cump"] || 0
   );

totalAcumUB =
   Number(
      ultimaFila["Utilidad Bruta"] || 0
   );

yearlyData = raw2
  .filter(r => String(r["Vendedor"] || "").trim() !== "")
  .filter(r => String(r["Vendedor"] || "").trim() !== "Total general")
  .filter(r => {

   const pct = Number(r["% Cump"] || 0);
   const cuota = Number(r["Cuota"] || 0);

   return !(pct === 0 && cuota === 0);

})
  .map(r => ({

      vendedor: String(r["Vendedor"] || ""),

      pct:
        Math.round(
            (Number(r["% Cump"]) || 0) * 100
        ),

      cuota:
        Number(r["Cuota"]) || 0,

      utilidad:
        Number(r["Utilidad Bruta"]) || 0,

      gap:
        Number(r["Gap"]) || 0

  }));

      if (raw2.length > 0) {
        const semKey = Object.keys(raw2[0]).find(k => normalize(k).includes("semana") || normalize(k) === "sem") || Object.keys(raw2[0])[0];
        const areas  = Object.keys(raw2[0]).filter(k => k !== semKey);
        if (areas.length > 0) {
          weeklyData = {
            labels: raw2.map(r => String(r[semKey]||"").replace("Semana ","S").replace("Semana","S")||"—"),
            series: areas.map((a,i) => ({
              name:   a,
              color:  AREA_COLORS[i % AREA_COLORS.length],
              values: raw2.map(r => toNum(r[a])),
            }))
          };
        } else {
          // PENDIENTE ELIMINAR
          buildWeeklyFromTable();
        }
      }
    } else {
      buildWeeklyFromTable();
    }

    /* ── Hoja 4: Dominios Mes ── */

if (wb.Sheets["Dominio_Mes"]) {

   const ws4 = wb.Sheets[wb.SheetNames[3]];
   const raw4 = XLSX.utils.sheet_to_json(ws4,{defval:0});

   domainMonthData = raw4
      .filter(r =>
         String(r["Etiquetas de fila"] || "").trim() !== ""
      )
      .filter(r =>
         !String(r["Etiquetas de fila"] || "").includes("Total")
      )
      .filter(r =>
   Number(r["% UB"] || 0) > 0
);
}

    /* ── Hoja 5: Dominio ACUMULADO ── */

if (wb.Sheets["Dominios_Acum"]) {

   const ws5 = wb.Sheets[wb.SheetNames[4]];
   const raw5 = XLSX.utils.sheet_to_json(ws5,{defval:0});

   domainYearData = raw5
      .filter(r =>
         String(r["Etiquetas de fila"] || "").trim() !== ""
      )
      .filter(r =>
         !String(r["Etiquetas de fila"] || "").includes("Total")
      )
      .filter(r =>
   Number(r["% UB"] || 0) > 0
);
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
   
console.log(
   "CUMPLIMIENTO MES A MES",
   cumplimientoMesData
);

  } catch(e) {
    console.error("Error parseExcel:", e);
    useDemoData();
  }

  renderAll();
}


function buildWeeklyFromTable() {
  if (!tableData.length) return;
  const labels = ["S1","S2","S3","S4"];
  weeklyData = {
    labels,
    series: tableData.map((row, i) => ({
      name:   row.area,
      color:  AREA_COLORS[i % AREA_COLORS.length],
      values: labels.map((_, wi) => Math.round(row.pct * (0.65 + wi * 0.12)))
    }))
  };
}

function normalize(s) {
  return String(s||"").toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .replace(/[^a-z0-9]/g,"");
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
  renderMainTitle();
  renderExecutiveSummary();
  renderCumplimientoMesAMes();
  renderWeeklyChart();
  renderDonut();
  renderKPIs();
  adjustLayoutByTeamCount();
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

    /*const est = normalize(row.estado);

    if (est.includes("clasif")) tr.classList.add("row-top");
    else if (est.includes("elim")) tr.classList.add("row-eliminated");
    else tr.classList.add("row-warning");

    let statusClass = "status-race", statusIcon = "🔥";
    if (est.includes("clasif")) { statusClass="status-classified"; statusIcon="🏆"; }
    else if (est.includes("elim")) { statusClass="status-eliminated"; statusIcon="❌"; }*/

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
      <td class="area-name">${esc(row.area)}</td>
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
        ${esc(row.vendedor)}
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

    if(!monthBody || !yearBody)
        return;

    monthBody.innerHTML = "";
    yearBody.innerHTML = "";

    /* ==========================
       DOMINIO MES
       ========================== */

    const sortedMonth =
        [...domainMonthData]
        .sort((a,b)=>{

            const pctA =
                Number(a["% UB"] || 0);

            const pctB =
                Number(b["% UB"] || 0);

            return pctB - pctA;

        });

    sortedMonth.forEach((row,idx)=>{

        const tr =
            document.createElement("tr");

        const medal =
            idx===0 ? "🥇" :
            idx===1 ? "🥈" :
            idx===2 ? "🥉" : "";

        tr.innerHTML = `
            <td>${idx + 1}</td>

            <td>
                ${medal}
                ${row["Etiquetas de fila"]}
            </td>

            <td>
                ${(
                    Number(row["% UB"] || 0) * 100
                ).toFixed(1)}%
            </td>

            <td>
                ${formatCurrency(
                    Number(row["Utilidad Bruta"] || 0)
                )}
            </td>

            <td class="${
                Number(row["Y to Y Ub"] || 0) >= 0
                    ? "gap-positivo"
                    : "gap-negativo"
            }">
                ${formatCurrency(
                    Number(row["Y to Y Ub"] || 0)
                )}
            </td>
        `;

        monthBody.appendChild(tr);

    });

    /* ==========================
       DOMINIO ACUMULADO
       ========================== */

    const sortedYear =
        [...domainYearData]
        .sort((a,b)=>{

            const pctA =
                Number(a["% UB"] || 0);

            const pctB =
                Number(b["% UB"] || 0);

            return pctB - pctA;

        });

    sortedYear.forEach((row,idx)=>{

        const tr =
            document.createElement("tr");

        const medal =
            idx===0 ? "🥇" :
            idx===1 ? "🥈" :
            idx===2 ? "🥉" : "";

        tr.innerHTML = `
            <td>${idx + 1}</td>

            <td>
                ${medal}
                ${row["Etiquetas de fila"]}
            </td>

            <td>
                ${(
                    Number(row["% UB"] || 0) * 100
                ).toFixed(1)}%
            </td>

            <td>
                ${formatCurrency(
                    Number(row["Utilidad Bruta"] || 0)
                )}
            </td>

            <td class="${
                Number(row["Y to Y Ub"] || 0) >= 0
                    ? "gap-positivo"
                    : "gap-negativo"
            }">
                ${formatCurrency(
                    Number(row["Y to Y Ub"] || 0)
                )}
            </td>
        `;

        yearBody.appendChild(tr);

    });

}

function renderExecutiveSummary(){

    const monthEl =
        document.getElementById("monthSummary");

    const yearEl =
        document.getElementById("yearSummary");

    if(!monthEl || !yearEl)
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

// TODO:
// reemplazar por totales oficiales
// provenientes del Excel
const ventasYear =
    totalAcumUB;

const cumplimientoYear =
    Math.round(
        totalAcumPct * 100
    );


const topYear =
    [...yearlyData]
        .sort((a, b) => {

            if ((b.pct || 0) !== (a.pct || 0))
                return (b.pct || 0) - (a.pct || 0);

            return (b.utilidad || 0) - (a.utilidad || 0);

        })[0];

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

    const partes = nombre.trim().split(" ");

    if(partes.length >= 2){

        const primerNombre = partes[partes.length - 2];
        const primerApellido = partes[0];

        return `${primerNombre} ${primerApellido}`;

    }

    return nombre;

}

/* ════════════════════════════
   GRÁFICA SEMANAL
   ════════════════════════════ */
function renderWeeklyChart() {
  const barsEl   = document.getElementById("weeklyChart");
  const legendEl = document.getElementById("areaLegend");
  const lineSVG  = document.getElementById("lineSVG");
  const labelsEl = document.getElementById("chartLabels");
  if (!barsEl) return;

  barsEl.innerHTML = legendEl.innerHTML = lineSVG.innerHTML = labelsEl.innerHTML = "";

  const { labels, series } = weeklyData;
  if (!labels.length || !series.length) {
    barsEl.innerHTML = '<div style="text-align:center;padding:22px;color:rgba(220,240,220,0.35);font-size:0.82rem;">Sin datos semanales</div>';
    return;
  }

  const allVals = series.flatMap(s => s.values);
  const maxVal  = Math.max(...allVals, 1);
  const CHART_H = 138;

  labels.forEach((lbl, wIdx) => {
    const grp = document.createElement("div");
    grp.className = "chart-group";
    series.forEach(s => {
      const bar = document.createElement("div");
      bar.className = "bar";
      bar.style.background = s.color;
      bar.style.height = "0px";
      bar.title = `${s.name}: ${s.values[wIdx]}`;
      grp.appendChild(bar);
      setTimeout(() => {
        bar.style.height = ((s.values[wIdx]||0) / maxVal * CHART_H) + "px";
      }, 120 + wIdx * 55);
    });
    barsEl.appendChild(grp);

    const sp = document.createElement("span");
    sp.textContent = String(lbl);
    labelsEl.appendChild(sp);
  });

  const lead = series[0];
  const W=400, nw=labels.length;
  const step = W / (nw - 1 || 1);
  const pts  = lead.values.map((v,i) => [
    i * step,
    CHART_H - ((v||0) / maxVal * CHART_H * 0.92)
  ]);

  const poly = document.createElementNS("http://www.w3.org/2000/svg","polyline");
  poly.setAttribute("points", pts.map(p=>p.join(",")).join(" "));
  poly.setAttribute("fill","none");
  poly.setAttribute("stroke", lead.color);
  poly.setAttribute("stroke-width","2");
  poly.setAttribute("stroke-linejoin","round");
  poly.setAttribute("opacity","0.8");
  lineSVG.appendChild(poly);

  pts.forEach(([x,y],i) => {
    const c = document.createElementNS("http://www.w3.org/2000/svg","circle");
    c.setAttribute("cx",x); c.setAttribute("cy",y);
    c.setAttribute("r","4"); c.setAttribute("fill",lead.color);
    c.setAttribute("stroke","rgba(4,13,5,0.9)"); c.setAttribute("stroke-width","1.2");
    const t = document.createElementNS("http://www.w3.org/2000/svg","title");
    t.textContent = `${lead.name} ${labels[i]}: ${lead.values[i]}`;
    c.appendChild(t); lineSVG.appendChild(c);
  });

  series.forEach(s => {
    const el = document.createElement("div");
    el.className = "al-item";
    el.innerHTML = `<div class="al-dot" style="background:${s.color}"></div><span>${esc(s.name)}</span>`;
    legendEl.appendChild(el);
  });
}

/* ════════════════════════════
   DONUT - GOLES POR ÁREA (CON SCROLL PARA TODOS LOS EQUIPOS)
   ════════════════════════════ */
function renderDonut() {
  const svg    = document.getElementById("donutSVG");
  const legend = document.getElementById("donutLegend");
  if (!svg||!legend) return;
  svg.innerHTML = legend.innerHTML = "";

  // Calcular altura dinámica según cantidad de equipos
  const teamCount = tableData.length;
  let maxHeight = "180px";
  
  if (teamCount > 12) {
    maxHeight = "340px";
  } else if (teamCount > 8) {
    maxHeight = "280px";
  } else if (teamCount > 5) {
    maxHeight = "220px";
  }
  
  // Aplicar estilos a la leyenda - SCROLL VERTICAL
  legend.style.maxHeight = maxHeight;
  legend.style.overflowY = "auto";
  legend.style.paddingRight = "8px";
  legend.style.display = "flex";
  legend.style.flexDirection = "column";
  legend.style.gap = "0.5rem";
  legend.style.scrollbarWidth = "thin";
  legend.style.scrollbarColor = `#FFD700 rgba(255,255,255,0.1)`;

  const total = tableData.reduce((s,r)=>s+(r.goles||0),0);
  const cx=100,cy=100,R=78,ri=48;
  let start = -Math.PI/2;

  if (total === 0) {
    const circle = document.createElementNS("http://www.w3.org/2000/svg","circle");
    circle.setAttribute("cx",cx); circle.setAttribute("cy",cy);
    circle.setAttribute("r",(R+ri)/2); circle.setAttribute("fill","none");
    circle.setAttribute("stroke","rgba(255,255,255,0.1)");
    circle.setAttribute("stroke-width", R-ri);
    svg.appendChild(circle);
    legend.innerHTML = '<div style="color:rgba(220,240,220,0.4);font-size:0.8rem;">Sin goles registrados</div>';
    return;
  }

  // Solo dibujar segmentos para equipos con goles > 0
  const teamsWithGoals = [...tableData].filter(row => (row.goles||0) > 0).sort((a,b) => (b.goles||0) - (a.goles||0));
  
  teamsWithGoals.forEach((row,i) => {
    const slice = ((row.goles||0)/total)*Math.PI*2;
    if (slice === 0) { start+=slice; return; }
    const end   = start+slice;
    const color = AREA_COLORS[i%AREA_COLORS.length];
    const large = slice > Math.PI ? 1:0;
    const x1=cx+R*Math.cos(start),  y1=cy+R*Math.sin(start);
    const x2=cx+R*Math.cos(end),    y2=cy+R*Math.sin(end);
    const ix1=cx+ri*Math.cos(end),  iy1=cy+ri*Math.sin(end);
    const ix2=cx+ri*Math.cos(start),iy2=cy+ri*Math.sin(start);

    const path=document.createElementNS("http://www.w3.org/2000/svg","path");
    path.setAttribute("d",`M${x1},${y1} A${R},${R},0,${large},1,${x2},${y2} L${ix1},${iy1} A${ri},${ri},0,${large},0,${ix2},${iy2} Z`);
    path.setAttribute("fill",color); path.setAttribute("stroke","#040d05");
    path.setAttribute("stroke-width","2"); path.setAttribute("opacity","0.88");
    path.addEventListener("mouseenter",()=>{path.setAttribute("opacity","1");path.setAttribute("stroke-width","3");});
    path.addEventListener("mouseleave",()=>{path.setAttribute("opacity","0.88");path.setAttribute("stroke-width","2");});
    const t=document.createElementNS("http://www.w3.org/2000/svg","title");
    t.textContent=`${row.area}: ${row.goles} goles`; path.appendChild(t); svg.appendChild(path);

    if (row.goles>0){
      const mid=start+slice/2;
      const tx=cx+(R+ri)/2*Math.cos(mid), ty=cy+(R+ri)/2*Math.sin(mid);
      const txt=document.createElementNS("http://www.w3.org/2000/svg","text");
      txt.setAttribute("x",tx); txt.setAttribute("y",ty+4);
      txt.setAttribute("text-anchor","middle"); txt.setAttribute("font-size","10.5");
      txt.setAttribute("fill","#fff"); txt.setAttribute("font-weight","bold");
      txt.textContent=row.goles; svg.appendChild(txt);
    }
    start=end;
  });

  // LEYENDA: MOSTRAR TODOS LOS EQUIPOS (incluyendo los que tienen 0 goles)
  const allTeamsSorted = [...tableData].sort((a,b) => (b.goles||0) - (a.goles||0));
  
  allTeamsSorted.forEach((row, i) => {
    const color = AREA_COLORS[i % AREA_COLORS.length];
    const li = document.createElement("div"); 
    li.className = "dl-item";
    li.style.display = "flex";
    li.style.alignItems = "center";
    li.style.gap = "0.6rem";
    li.style.justifyContent = "space-between";
    li.style.padding = "4px 0";
    li.style.borderBottom = "1px solid rgba(255,255,255,0.05)";
    
    li.innerHTML = `
      <div style="display:flex; align-items:center; gap:0.5rem;">
        <div class="dl-color" style="background:${color}; width:12px; height:12px; border-radius:3px;"></div>
        <span class="dl-label" style="font-size:0.85rem; color:var(--text-dim);">${esc(row.area)}</span>
      </div>
      <span class="dl-val" style="font-weight:bold; font-size:0.9rem; color:var(--gold);">${row.goles}</span>
    `;
    legend.appendChild(li);
  });
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
    val: top.area
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

function esc(s){ return String(s||"—").replace(/[&<>]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[m]||m)); }

/* ════════════════════════════
   AJUSTAR LAYOUT SEGÚN CANTIDAD DE EQUIPOS
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
    // Reordenar: Actividad Semanal arriba, Goles y KPIs abajo en 2 columnas
    dashboard.style.display = "grid";
    dashboard.style.gridTemplateColumns = "1fr";
    dashboard.style.gap = "1rem";
    
    // Crear contenedor para Goles y KPIs
    let wrapper = document.querySelector('.goles-kpis-wrapper');
    if (!wrapper) {
      wrapper = document.createElement('div');
      wrapper.className = 'goles-kpis-wrapper';
      wrapper.style.display = "grid";
      wrapper.style.gridTemplateColumns = "1fr 1fr";
      wrapper.style.gap = "1rem";
      
      // Mover las tarjetas al wrapper
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
    // Restaurar layout original (3 columnas)
    dashboard.style.display = "grid";
    dashboard.style.gridTemplateColumns = "1fr 1fr 1fr";
    dashboard.style.gap = "1rem";
    
    // Eliminar wrapper si existe y restaurar orden
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