const HOSTED_EXCEL_URLS = [
  "https://dl.dropboxusercontent.com/scl/fi/xv93qy8t293au22unojov/Por-venda_2026.xlsm?rlkey=gtyk5at1kta87k5rlioktgwxb&st=yaw1gyip&dl=1",
  "https://www.dropbox.com/scl/fi/xv93qy8t293au22unojov/Por-venda_2026.xlsm?rlkey=gtyk5at1kta87k5rlioktgwxb&st=yaw1gyip&raw=1"
];

const DIAS = ["Segunda", "Terca", "Quarta", "Quinta", "Sexta", "Sabado", "Domingo"];
const MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
const DRINKS = new Set(["DRINKS", "DRINKS ESPECIAIS", "DRINK SEM ALCOOL", "DOUBLE DRINKS"]);
const COLORS = ["#2f5f9e", "#24756b", "#7056a3", "#b66a1d", "#a83d52", "#526171", "#0e7490", "#854d0e"];

let allData = [];
let filteredData = [];
let matrizData = [];
let charts = {};

const $ = id => document.getElementById(id);
const text = value => value == null ? "" : String(value).trim();

function num(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  let clean = String(value ?? "").trim().replace(/R\$|\s/g, "");
  if (!clean) return 0;
  if (clean.includes(",")) clean = clean.replace(/\./g, "").replace(",", ".");
  const parsed = parseFloat(clean);
  return Number.isFinite(parsed) ? parsed : 0;
}

function money(value) {
  return "R$ " + Number(value || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function shortMoney(value) {
  const n = Number(value || 0);
  if (Math.abs(n) >= 1e6) return "R$ " + (n / 1e6).toFixed(1) + "M";
  if (Math.abs(n) >= 1e3) return "R$ " + (n / 1e3).toFixed(0) + "K";
  return money(n);
}

function nb(value) {
  return Math.round(value || 0).toLocaleString("pt-BR");
}

function fmtPct(value) {
  return (Number(value || 0) * 100).toFixed(1) + "%";
}

function fmtX(value) {
  return value ? "x" + Number(value).toFixed(2) : "-";
}

function norm(value) {
  return text(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/ç/g, "c");
}

function asPctValue(value) {
  const n = Number(value || 0);
  return Math.abs(n) > 1.5 ? n / 100 : n;
}

function monthIndex(mes) {
  const parts = norm(mes).split("/");
  const month = (parts[0] || "").slice(0, 3);
  const yy = parseInt(parts[1] || 0, 10);
  const year = yy ? yy + (yy < 100 ? 2000 : 0) : 0;
  return year * 12 + MESES.map(norm).indexOf(month);
}

function yearFromMes(mes) {
  const parts = text(mes).split("/");
  const yy = parseInt(parts[1] || 0, 10);
  return yy ? String(yy + (yy < 100 ? 2000 : 0)) : "";
}

function sortMonths(values) {
  return [...values].sort((a, b) => monthIndex(a) - monthIndex(b));
}

function setProgress(percent, label) {
  $("progressBar").style.width = percent + "%";
  $("progressText").textContent = label || "";
}

function showLoading(on) {
  $("loading").classList.toggle("hidden", !on);
  $("upload").classList.add("hidden");
}

function showUpload() {
  document.body.classList.remove("busy");
  $("upload").classList.remove("hidden");
  $("app").classList.add("hidden");
  $("loading").classList.add("hidden");
}

async function loadHostedExcel() {
  showLoading(true);
  setProgress(10, "Baixando planilha do Dropbox...");
  for (const url of HOSTED_EXCEL_URLS) {
    try {
      const response = await fetch(url, { cache: "no-store" });
      if (!response.ok) throw new Error("HTTP " + response.status);
      const blob = await response.blob();
      await loadFile(new File([blob], "Por venda_2026.xlsm"));
      return;
    } catch (error) {
      console.warn(error);
    }
  }
  showUpload();
  $("uploadHint").textContent = "Dropbox nao carregou. Selecione o arquivo local.";
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = event => resolve(event.target.result);
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

async function loadFile(file) {
  try {
    showLoading(true);
    $("fileName").textContent = file.name;
    setProgress(22, "Lendo arquivo...");
    const buffer = await readFileAsArrayBuffer(file);
    setProgress(38, "Decodificando abas...");
    const workbook = XLSX.read(buffer, { type: "array", dense: true, cellFormula: false, cellHTML: false, cellStyles: false, cellText: false });
    parseWorkbook(workbook);
    setProgress(82, "Montando filtros...");
    populateFilters();
    setProgress(94, "Renderizando painel...");
    $("app").classList.remove("hidden");
    $("upload").classList.add("hidden");
    applyFilters();
    setProgress(100, "Pronto");
    setTimeout(() => $("loading").classList.add("hidden"), 250);
  } catch (error) {
    console.error(error);
    showUpload();
    $("uploadHint").textContent = "Erro ao ler planilha: " + error.message;
  }
}

function parseWorkbook(workbook) {
  const vendasSheet = workbook.SheetNames.find(name => norm(name) === "vendasoperador") || workbook.SheetNames.find(name => norm(name).includes("venda"));
  if (!vendasSheet) throw new Error("Aba VendasOperador nao encontrada");

  const raw = XLSX.utils.sheet_to_json(workbook.Sheets[vendasSheet], { header: 1, defval: "" });
  allData = raw.slice(1).map(row => ({
    tipo: text(row[6]),
    produto: text(row[8]),
    qtd: num(row[9]),
    total: num(row[12]),
    pgto: text(row[13]),
    grupo: text(row[16]),
    dia: cleanDia(text(row[17])),
    semana: num(row[18]),
    hora: text(row[19]),
    mes: text(row[20]),
    liquido: num(row[21]),
    qtven: num(row[22]),
    quinzena: text(row[24]),
    qt_double: num(row[25]),
    qt_final: num(row[26])
  })).filter(row => row.mes && row.produto && (row.liquido || row.total || row.qtd));

  const matrizSheet = workbook.SheetNames.find(name => norm(name).includes("matrizanalitico") || norm(name).includes("matrizanaltico"));
  matrizData = [];
  if (matrizSheet) {
    const rawMatriz = XLSX.utils.sheet_to_json(workbook.Sheets[matrizSheet], { header: 1, defval: "" });
    for (let index = 7; index < rawMatriz.length; index++) {
      const row = rawMatriz[index];
      const mes = text(row[1]);
      if (!mes) continue;
      const item = {
        mes,
        receita: num(row[7]),
        custoFixo: num(row[30]),
        custoFixoVar: num(row[31]),
        totalFixo: num(row[32]),
        custoGeral: num(row[35]),
        cmv: num(row[36]),
        mkp: num(row[37]),
        margem: asPctValue(num(row[38])),
        pctCmv: asPctValue(num(row[42])),
        lucPrej: num(row[43]),
        pctLuc: asPctValue(num(row[44]))
      };
      if (item.receita || item.custoGeral || item.lucPrej) matrizData.push(item);
    }
  }
}

function cleanDia(value) {
  const n = norm(value);
  if (n.startsWith("seg")) return "Segunda";
  if (n.startsWith("ter")) return "Terca";
  if (n.startsWith("qua")) return "Quarta";
  if (n.startsWith("qui")) return "Quinta";
  if (n.startsWith("sex")) return "Sexta";
  if (n.startsWith("sab")) return "Sabado";
  if (n.startsWith("dom")) return "Domingo";
  return value;
}

function unique(key, rows = allData) {
  return [...new Set(rows.map(row => row[key]).filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), "pt-BR"));
}

function fillSelect(id, values, label = "Todos") {
  const element = $(id);
  const oldValue = element.value;
  element.innerHTML = "";
  element.append(new Option(label, ""));
  values.forEach(value => element.append(new Option(value, value)));
  if (values.includes(oldValue)) element.value = oldValue;
}

function populateFilters() {
  fillSelect("fAno", [...new Set(allData.map(row => yearFromMes(row.mes)).filter(Boolean))].sort());
  fillSelect("fMes", sortMonths(unique("mes")));
  fillSelect("fGrupo", unique("grupo"));
  fillSelect("fCanal", unique("tipo"));
  fillSelect("fPgto", unique("pgto"));
  fillSelect("fSemana", [...new Set(allData.map(row => String(Math.round(row.semana))).filter(value => value && value !== "0"))].sort((a, b) => +a - +b), "Todas");
  fillSelect("fDia", DIAS.filter(dia => allData.some(row => row.dia === dia)));

  const dataList = $("produtosList");
  dataList.innerHTML = "";
  unique("produto").forEach(value => {
    const option = document.createElement("option");
    option.value = value;
    dataList.append(option);
  });
}

function currentFilters() {
  return {
    ano: $("fAno").value,
    mes: $("fMes").value,
    grupo: $("fGrupo").value,
    tipo: $("fCanal").value,
    pgto: $("fPgto").value,
    semana: $("fSemana").value,
    dia: $("fDia").value,
    produto: norm($("fProduto").value)
  };
}

function passes(row, filters) {
  if (filters.ano && yearFromMes(row.mes) !== filters.ano) return false;
  if (filters.mes && row.mes !== filters.mes) return false;
  if (filters.grupo && row.grupo !== filters.grupo) return false;
  if (filters.tipo && row.tipo !== filters.tipo) return false;
  if (filters.pgto && row.pgto !== filters.pgto) return false;
  if (filters.semana && String(Math.round(row.semana)) !== filters.semana) return false;
  if (filters.dia && row.dia !== filters.dia) return false;
  if (filters.produto && !norm(row.produto).includes(filters.produto)) return false;
  return true;
}

function applyFilters() {
  document.body.classList.add("busy");
  requestAnimationFrame(() => {
    const filters = currentFilters();
    filteredData = allData.filter(row => passes(row, filters));
    renderAll();
    document.body.classList.remove("busy");
  });
}

function matrizFiltered() {
  const filters = currentFilters();
  return matrizData.filter(row => {
    if (filters.ano && yearFromMes(row.mes) !== filters.ano) return false;
    if (filters.mes && row.mes !== filters.mes) return false;
    return true;
  });
}

function summarize(rows) {
  const products = new Set();
  const days = new Set();
  let liquido = 0;
  let total = 0;
  let itens = 0;
  let vendas = 0;
  let drinks = 0;
  let doubleQty = 0;
  let drinkQty = 0;

  rows.forEach(row => {
    liquido += row.liquido;
    total += row.total;
    itens += row.qtd;
    vendas += row.qtven > 0 ? row.qtven : 0;
    products.add(row.produto);
    days.add([row.mes, Math.round(row.semana), row.dia].join("|"));
    doubleQty += row.qt_double;
    if (DRINKS.has(String(row.grupo).toUpperCase())) {
      drinks += row.liquido;
      drinkQty += row.qt_final || row.qtd;
    }
  });

  return {
    liquido,
    total,
    itens,
    vendas,
    produtos: products.size,
    dias: days.size,
    drinks,
    doubleQty,
    drinkQty,
    ticket: vendas ? liquido / vendas : 0,
    mix: liquido ? drinks / liquido : 0
  };
}

function sum(rows, key) {
  return rows.reduce((acc, row) => acc + (Number(row[key]) || 0), 0);
}

function avg(rows, key) {
  return rows.length ? sum(rows, key) / rows.length : 0;
}

function weighted(rows, key, weight = "receita") {
  const base = sum(rows, weight);
  return base ? rows.reduce((acc, row) => acc + (Number(row[key]) || 0) * (Number(row[weight]) || 0), 0) / base : avg(rows, key);
}

function financeSummary(summary) {
  const rows = matrizFiltered();
  const receita = summary.liquido;
  const custoFixo = sum(rows, "custoFixo");
  const custoFixoVar = sum(rows, "custoFixoVar");
  const totalFixo = sum(rows, "totalFixo");
  const custoGeral = sum(rows, "custoGeral");
  const cmv = sum(rows, "cmv");
  const custoVariavel = Math.max(0, custoGeral - totalFixo);
  const lucro = rows.length ? sum(rows, "lucPrej") : receita - custoGeral;
  const lucroPct = receita ? lucro / receita : 0;
  const pctCmv = rows.length ? weighted(rows, "pctCmv") : receita ? cmv / receita : 0;
  const margem = rows.length ? weighted(rows, "margem") : receita ? (receita - cmv) / receita : 0;
  const mkp = rows.length ? weighted(rows, "mkp") : 0;
  const cobertura = totalFixo ? receita / totalFixo : 0;
  const eficiencia = custoGeral ? receita / custoGeral : 0;
  const pe = margem ? totalFixo / margem : 0;
  return { rows, receita, custoFixo, custoFixoVar, totalFixo, custoGeral, custoVariavel, cmv, lucro, lucroPct, pctCmv, margem, mkp, cobertura, eficiencia, pe };
}

function renderAll() {
  const summary = summarize(filteredData);
  const finance = financeSummary(summary);
  renderStatus(summary);
  renderKPIs(summary, finance);
  renderFinanceiro(summary, finance);
  renderMediaSerie();
  renderCharts();
  renderWeekly();
  renderProducts();
}

function renderStatus(summary) {
  $("sRows").textContent = nb(filteredData.length) + " / " + nb(allData.length);
  const months = sortMonths([...new Set(filteredData.map(row => row.mes).filter(Boolean))]);
  $("sPeriod").textContent = months.length ? months[0] + " - " + months[months.length - 1] : "-";
  $("sProducts").textContent = nb(summary.produtos);
}

function renderKPIs(summary, finance) {
  const cards = [
    ["Receita liquida", money(summary.liquido), money(summary.total) + " bruto", "--blue"],
    ["Receita bruta", money(summary.total), "valor total antes do liquido", "--slate"],
    ["N de vendas", nb(summary.vendas), nb(filteredData.length) + " registros", "--green"],
    ["Ticket medio", money(summary.ticket), "liquido / vendas", "--purple"],
    ["Itens vendidos", nb(summary.itens), nb(summary.produtos) + " produtos distintos", "--amber"],
    ["Receita / dia", summary.dias ? shortMoney(summary.liquido / summary.dias) : "-", summary.dias + " dias de operacao", "--blue"],
    ["Lucro", money(finance.lucro), fmtPct(finance.lucroPct) + " sobre receita", "--green"],
    ["Custo geral", shortMoney(finance.custoGeral), "total da matriz", "--rose"],
    ["Mix drinks", fmtPct(summary.mix), money(summary.drinks) + " em drinks", "--rose"],
    ["Double drinks", nb(summary.doubleQty) + " un", "unidades double vendidas", "--slate"],
    ["Qtd vendas drinks", nb(summary.drinkQty) + " un", "drinks + especiais", "--green"],
    ["Ponto equilibrio", finance.pe ? shortMoney(finance.pe) : "-", finance.pe && summary.liquido >= finance.pe ? "atingido" : "meta contabil", "--amber"]
  ];

  $("kpis").innerHTML = cards.map(card => `
    <article class="card" style="--card-accent:var(${card[3]})">
      <div class="card-label">${card[0]}</div>
      <div class="card-value">${card[1]}</div>
      <div class="card-sub">${card[2]}</div>
    </article>`).join("");
}

function mini(label, value, sub, cls = "") {
  return `<div class="mini"><div class="mini-label">${label}</div><div class="mini-value ${cls}">${value}</div><div class="mini-sub">${sub}</div></div>`;
}

function renderFinanceiro(summary, finance) {
  $("receitaGrid").innerHTML =
    mini("Receita total", money(finance.receita), "coluna liquida", "pos") +
    mini(finance.lucro >= 0 ? "Lucro" : "Prejuizo", money(finance.lucro), fmtPct(finance.lucroPct) + " sobre receita", finance.lucro >= 0 ? "pos" : "neg") +
    mini("Receita / dia", summary.dias ? shortMoney(finance.receita / summary.dias) : "-", summary.dias + " dias de operacao") +
    mini("Cobertura custo fixo", finance.cobertura ? fmtPct(finance.cobertura) : "-", finance.cobertura >= 1 ? "custo fixo coberto" : "abaixo do fixo", finance.cobertura >= 1 ? "pos" : "neg") +
    mini("Eficiencia operacional", finance.eficiencia ? fmtX(finance.eficiencia) : "-", "receita / custo geral", finance.eficiencia >= 1 ? "pos" : "neg") +
    mini("Indice de lucratividade", fmtPct(finance.lucroPct), "lucro / receita", finance.lucroPct >= 0 ? "pos" : "neg") +
    mini("Mix drinks", fmtPct(summary.mix), money(summary.drinks) + " em drinks") +
    mini("Ticket medio", money(summary.ticket), "liquido / vendas");

  $("custosGrid").innerHTML =
    mini("Custo geral", shortMoney(finance.custoGeral), "total geral", "neg") +
    mini("Custo fixo", shortMoney(finance.custoFixo), "fixo puro") +
    mini("Fixo / variavel", shortMoney(finance.custoFixoVar), "componente misto") +
    mini("Total fixo", shortMoney(finance.totalFixo), "fixo + misto") +
    mini("Custo variavel", shortMoney(finance.custoVariavel), "custo geral - total fixo", "neg") +
    mini("CMV", shortMoney(finance.cmv), "custo merc. vendida", "neg") +
    mini("CMV %", fmtPct(finance.pctCmv), "custo / receita", finance.pctCmv <= .4 ? "pos" : "neg") +
    mini("MKP", finance.mkp ? fmtX(finance.mkp) : "-", "markup medio") +
    mini("Margem contrib.", fmtPct(finance.margem), "margem de contribuicao", finance.margem >= .5 ? "pos" : "") +
    mini("Ponto equilibrio", finance.pe ? shortMoney(finance.pe) : "-", finance.pe && finance.receita >= finance.pe ? "atingido" : "necessario", finance.pe && finance.receita >= finance.pe ? "pos" : "");
}

function renderMediaSerie() {
  const rows = matrizData.filter(row => row.receita || row.custoGeral || row.lucPrej);
  if (!rows.length) {
    $("mediaSerie").innerHTML = '<div class="hint">A matriz analitica ainda nao foi encontrada nesta planilha.</div>';
    return;
  }

  const lucro = avg(rows, "lucPrej");
  const receita = avg(rows, "receita");
  const pctLuc = receita ? lucro / receita : 0;
  const items = [
    ["Receita", shortMoney(receita)],
    ["Custo fixo", shortMoney(avg(rows, "custoFixo"))],
    ["Fixo / var", shortMoney(avg(rows, "custoFixoVar"))],
    ["Fixo total", shortMoney(avg(rows, "totalFixo"))],
    ["Custo geral", shortMoney(avg(rows, "custoGeral"))],
    ["Lucro liquido", shortMoney(lucro)],
    ["Lucratividade", fmtPct(pctLuc)],
    ["CMV", shortMoney(avg(rows, "cmv"))],
    ["CMV %", fmtPct(weighted(rows, "pctCmv"))],
    ["MKP", fmtX(weighted(rows, "mkp"))],
    ["Margem", fmtPct(weighted(rows, "margem"))]
  ];

  $("mediaSerie").innerHTML = items.map(item => `<div class="metric-pill-card"><small>${item[0]}</small><strong>${item[1]}</strong></div>`).join("");
}

function group(rows, keyFn, valFn) {
  const map = new Map();
  rows.forEach(row => {
    const key = keyFn(row) || "Sem info";
    map.set(key, (map.get(key) || 0) + valFn(row));
  });
  return map;
}

function rowsFromMap(map, limit) {
  return [...map.entries()].sort((a, b) => b[1] - a[1]).slice(0, limit || 999);
}

function chartBase() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 180 },
    plugins: {
      legend: { labels: { color: "#334155", boxWidth: 10, boxHeight: 10, usePointStyle: true, font: { weight: "750" } } },
      tooltip: { backgroundColor: "rgba(15,23,42,.96)", titleColor: "#fff", bodyColor: "#f8fafc", padding: 12, cornerRadius: 10 }
    },
    scales: {
      x: { ticks: { color: "#526171", font: { weight: "700" } }, grid: { color: "rgba(148,163,184,.18)" } },
      y: { ticks: { color: "#526171", font: { weight: "700" }, callback: value => shortMoney(value) }, grid: { color: "rgba(148,163,184,.2)" } }
    }
  };
}

function makeChart(id, type, data, options = {}) {
  if (charts[id]) charts[id].destroy();
  charts[id] = new Chart($(id), { type, data, options: { ...chartBase(), ...options } });
}

function renderCharts() {
  const byMes = sortMonths([...new Set(filteredData.map(row => row.mes).filter(Boolean))]);
  const mensal = byMes.map(mes => filteredData.filter(row => row.mes === mes));

  makeChart("chartMensal", "bar", {
    labels: byMes,
    datasets: [
      { label: "Receita liquida", data: mensal.map(rows => rows.reduce((acc, row) => acc + row.liquido, 0)), backgroundColor: "rgba(47,95,158,.78)", borderColor: "#2f5f9e", borderWidth: 1, borderRadius: 7, yAxisID: "y" },
      { label: "Ticket medio", type: "line", data: mensal.map(rows => summarize(rows).ticket), borderColor: "#7056a3", backgroundColor: "rgba(112,86,163,.13)", pointBackgroundColor: "#fff", pointBorderColor: "#7056a3", borderWidth: 3, tension: .35, yAxisID: "y1" }
    ]
  }, {
    scales: {
      x: chartBase().scales.x,
      y: { ...chartBase().scales.y },
      y1: { position: "right", ticks: { color: "#7056a3", font: { weight: "800" }, callback: value => money(value) }, grid: { drawOnChartArea: false } }
    }
  });

  const pgtoRows = rowsFromMap(group(filteredData, row => row.pgto, row => row.liquido), 8);
  makeChart("chartPgto", "doughnut", {
    labels: pgtoRows.map(row => row[0]),
    datasets: [{ data: pgtoRows.map(row => row[1]), backgroundColor: COLORS, borderColor: "#fff", borderWidth: 3, hoverOffset: 8 }]
  }, { cutout: "62%", scales: {} });

  const grupo = rowsFromMap(group(filteredData, row => row.grupo, row => row.liquido), 10).reverse();
  makeChart("chartGrupo", "bar", {
    labels: grupo.map(row => row[0]),
    datasets: [{ label: "Receita", data: grupo.map(row => row[1]), backgroundColor: "rgba(36,117,107,.75)", borderColor: "#24756b", borderWidth: 1, borderRadius: 7 }]
  }, { indexAxis: "y" });

  const dia = DIAS.map(diaNome => [diaNome, filteredData.filter(row => row.dia === diaNome).reduce((acc, row) => acc + row.liquido, 0)]);
  makeChart("chartDia", "bar", {
    labels: dia.map(row => row[0]),
    datasets: [{ label: "Receita", data: dia.map(row => row[1]), backgroundColor: "rgba(182,106,29,.75)", borderColor: "#b66a1d", borderWidth: 1, borderRadius: 7 }]
  });

  const hora = rowsFromMap(group(filteredData, row => row.hora, row => row.liquido), 14).sort((a, b) => String(a[0]).localeCompare(String(b[0])));
  makeChart("chartHora", "line", {
    labels: hora.map(row => row[0]),
    datasets: [{ label: "Receita", data: hora.map(row => row[1]), borderColor: "#2f5f9e", backgroundColor: "rgba(47,95,158,.12)", fill: true, borderWidth: 3, pointBackgroundColor: "#fff", pointBorderColor: "#2f5f9e", pointBorderWidth: 2, tension: .38 }]
  });
}

function renderWeekly() {
  const months = sortMonths([...new Set(filteredData.map(row => row.mes).filter(Boolean))]);
  let html = '<thead><tr><th>Mes</th><th class="num">1a sem</th><th class="num">2a sem</th><th class="num">3a sem</th><th class="num">4a sem</th><th class="num">5a sem</th><th class="num">Total</th></tr></thead><tbody>';
  months.forEach(mes => {
    const vals = [1, 2, 3, 4, 5].map(week => filteredData.filter(row => row.mes === mes && Math.round(row.semana) === week).reduce((acc, row) => acc + row.liquido, 0));
    const positive = vals.filter(value => value > 0);
    const max = positive.length ? Math.max(...positive) : 0;
    const min = positive.length ? Math.min(...positive) : 0;
    html += '<tr><td><strong>' + mes + '</strong></td>' + vals.map(value => `<td class="num ${value && value === max ? 'pos' : value && value === min ? 'neg' : ''}">${money(value)}</td>`).join("") + `<td class="num"><strong>${money(vals.reduce((acc, value) => acc + value, 0))}</strong></td></tr>`;
  });
  html += "</tbody>";
  $("weeklyTable").innerHTML = html;
}

function renderProducts() {
  const query = norm($("tableSearch").value);
  const total = filteredData.reduce((acc, row) => acc + row.liquido, 0) || 1;
  const map = new Map();
  filteredData.forEach(row => {
    if (query && !norm(row.produto).includes(query)) return;
    const item = map.get(row.produto) || { produto: row.produto, grupo: row.grupo, qtd: 0, liquido: 0 };
    item.qtd += row.qtd;
    item.liquido += row.liquido;
    map.set(row.produto, item);
  });

  const rows = [...map.values()].sort((a, b) => b.liquido - a.liquido).slice(0, 80);
  $("prodCount").textContent = nb(map.size) + " produtos";
  $("prodTable").innerHTML = rows.map((row, index) => `<tr><td><span class="rank">${index + 1}</span></td><td>${row.produto}</td><td><span class="badge">${row.grupo || '-'}</span></td><td class="num">${nb(row.qtd)}</td><td class="num"><strong>${money(row.liquido)}</strong></td><td class="num">${(row.liquido / total * 100).toFixed(1)}%</td></tr>`).join("");
}

function resetFilters() {
  ["fAno", "fMes", "fGrupo", "fCanal", "fPgto", "fSemana", "fDia"].forEach(id => $(id).value = "");
  $("fProduto").value = "";
  $("tableSearch").value = "";
  applyFilters();
}

function bindEvents() {
  ["fAno", "fMes", "fGrupo", "fCanal", "fPgto", "fSemana", "fDia", "fProduto"].forEach(id => $(id).addEventListener(id === "fProduto" ? "input" : "change", applyFilters));
  $("tableSearch").addEventListener("input", renderProducts);
  $("clearBtn").addEventListener("click", resetFilters);
  $("swapBtn").addEventListener("click", showUpload);
  $("pickBtn").addEventListener("click", () => $("fileInput").click());
  $("hostedBtn").addEventListener("click", loadHostedExcel);
  $("fileInput").addEventListener("change", event => { if (event.target.files[0]) loadFile(event.target.files[0]); });
  $("drop").addEventListener("dragover", event => { event.preventDefault(); $("drop").style.borderColor = "var(--blue)"; });
  $("drop").addEventListener("dragleave", () => $("drop").style.borderColor = "#b9c6d8");
  $("drop").addEventListener("drop", event => {
    event.preventDefault();
    $("drop").style.borderColor = "#b9c6d8";
    if (event.dataTransfer.files[0]) loadFile(event.dataTransfer.files[0]);
  });
  window.addEventListener("resize", () => Object.values(charts).forEach(chart => { try { chart.resize(); } catch (error) {} }));
}

bindEvents();
setTimeout(loadHostedExcel, 350);
