/* ================================================
   COPA COPILOT — APP.JS v2
   Parser robusto para copa_copilot_datos.xlsx
   ================================================ */

const AREA_COLORS = ["#e8c84a","#5ab4d6","#a87cbe","#d4863a","#c05555","#6ec87a","#4ab8c8","#f39c12","#1abc9c","#e74c3c","#3498db","#9b59b6","#2ecc71"];

let tableData  = [];
let weeklyData = { labels:[], series:[] };

document.addEventListener("DOMContentLoaded", loadExcelAuto);

function loadExcelAuto() {
  showLoading();
  fetch("copa_copilot_datos.xlsx")
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
function useDemoData() {
  tableData = [
    {pos:1,area:"Ventas",    pct:61,inter:890,goles:13,estado:"En Carrera"},
    {pos:2,area:"Marketing", pct:55,inter:730,goles:9, estado:"En Carrera"},
    {pos:3,area:"Finanzas",  pct:47,inter:605,goles:8, estado:"En Carrera"},
    {pos:4,area:"RRHH",      pct:36,inter:420,goles:5, estado:"En Carrera"},
    {pos:5,area:"Admin",     pct:28,inter:280,goles:2, estado:"En Carrera"},
  ];
  weeklyData = {
    labels:["S1","S2","S3","S4"],
    series:[
      {name:"Ventas",    color:"#e8c84a",values:[40,48,55,61]},
      {name:"Marketing", color:"#5ab4d6",values:[35,40,50,55]},
      {name:"Finanzas",  color:"#a87cbe",values:[30,38,45,47]},
      {name:"RRHH",      color:"#d4863a",values:[25,30,34,36]},
      {name:"Admin",     color:"#c05555",values:[20,22,26,28]},
    ]
  };
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

    console.log("Columnas encontradas:", rows.length > 0 ? Object.keys(rows[0]) : "ninguna");
    console.log("Primera fila:", rows[0]);

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

      console.log("Columnas mapeadas →", {kPos,kArea,kPct,kInter,kGoles,kEstado});

      tableData = rows
        .filter(r => kArea && String(r[kArea]||"").trim() !== "")
        .map((r, i) => {
          let pct = kPct ? toNum(r[kPct]) : 0;
          if (pct > 0 && pct <= 1) pct = pct * 100;

          return {
            pos:    kPos    ? (parseInt(r[kPos])  || i+1) : i+1,
            area:   kArea   ? String(r[kArea]).trim()      : `Área ${i+1}`,
            pct:    Math.round(pct),
            inter:  kInter  ? toNum(r[kInter])             : 0,
            goles:  kGoles  ? toNum(r[kGoles])             : 0,
            compromiso: kEstado ? toNum(r[kEstado]):0,
          };
        });

      console.log("tableData procesado:", tableData);
    }

    /* ── Hoja 2: Actividad Semanal ── */
    if (wb.SheetNames[1]) {
      const ws2  = wb.Sheets[wb.SheetNames[1]];
      const raw2 = XLSX.utils.sheet_to_json(ws2, { defval:0 });
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
          buildWeeklyFromTable();
        }
      }
    } else {
      buildWeeklyFromTable();
    }

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
  if (v === null || v === undefined || v === "") return 0;
  const n = parseFloat(String(v).replace(",",".").replace("%",""));
  return isNaN(n) ? 0 : n;
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
  ;
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
    if ((b.pct || 0) !== (a.pct || 0)) return (b.pct || 0) - (a.pct || 0);
    return (b.goles || 0) - (a.goles || 0);
  });

  sorted.forEach((row, idx) => {
    const tr = document.createElement("tr");
    const pct = row.pct;
    const est = normalize(row.estado);

    if (est.includes("clasif")) tr.classList.add("row-top");
    else if (est.includes("elim")) tr.classList.add("row-eliminated");
    else tr.classList.add("row-warning");

    let barClass = "fill-gold";
    if (pct < 30) barClass = "fill-red";
    else if (pct < 50) barClass = "fill-amber";

    let statusClass = "status-race", statusIcon = "🔥";
    if (est.includes("clasif")) { statusClass="status-classified"; statusIcon="🏆"; }
    else if (est.includes("elim")) { statusClass="status-eliminated"; statusIcon="❌"; }

    let medal = "";
    if (idx === 0) { statusIcon="🥇";}
    else if (idx === 1) { statusIcon="🥈";}
    else if (idx === 2) { statusIcon="🥉";}

    const posNum = idx + 1;

    tr.innerHTML = `
      <td><div class="pos-badge${topClass}">${posNum}</div></td>
      <td class="area-name">${esc(row.area)}</td>
      <td>
        <div class="pct-bar-wrap">
          <div class="pct-bar-bg">
            <div class="pct-bar-fill ${barClass}" style="width:0%" data-target="${pct}%"></div>
          </div>
          <span class="pct-num">${medal} ${pct}%</span>
        </div>
      </td>
      <td>${Number(row.inter).toLocaleString()}</td>
      <td><span class="currency-cell">${formatCurrency(row.goles)}</span></td>
      <td>${formatCurrency(row.compromiso)}</td>
    `;
    tbody.appendChild(tr);
  });

  requestAnimationFrame(() => setTimeout(() => {
    document.querySelectorAll(".pct-bar-fill").forEach(b => b.style.width = b.dataset.target);
  }, 180));
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

  const totalG = tableData.reduce((s, r) => s + (r.goles || 0), 0);
  const totalI = tableData.reduce((s, r) => s + (r.inter || 0), 0);
  const top = [...tableData].sort((a, b) => (b.goles || 0) - (a.goles || 0))[0];
  const n = tableData.length;
    const totalSemanal = weeklyData.series.reduce(
    (sumSeries, serie) =>
      sumSeries +
      serie.values.reduce((sumVals, val) => sumVals + (val || 0), 0),
    0
  );

  const kpisData = [
    { label: "🏆 % Cumplimiento del mes", val: `${totalG} ⚽` },
    { label: "📈 Ventas del mes", val: totalI.toLocaleString() },
    { label: "💬 Total Interacciones Actividad Semanal", val: totalSemanal.toLocaleString() },
    { label: "🥇 MVP DEL MES", val: top.area },
    { label: "🏟️ Vendedores en carrera", val: `${n} equipos` },
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