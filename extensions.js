/* Extensoes seguras para a base limpa do JANELA-TESTE.
   Este arquivo acrescenta blocos de KPIs sem alterar a leitura original. */
(function () {
  function injectStyle() {
    if ($("extendedStyles")) return;
    const style = document.createElement("style");
    style.id = "extendedStyles";
    style.textContent = `
      .insight-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:16px}
      .evo-cell{display:grid;gap:3px}
      .delta{font-size:.68rem;font-weight:900}
      .delta.good{color:var(--green)}
      .delta.bad{color:var(--rose)}
      .delta.neutral{color:var(--muted)}
      .subtle-total{background:#f8fafc}
      .table-card.compact .table-wrap{max-height:460px}
      @media(max-width:1180px){.insight-grid{grid-template-columns:1fr}}
      @media(max-width:720px){.insight-grid{gap:12px}.table-card.compact .table-wrap{max-height:390px}}
    `;
    document.head.append(style);
  }

  function ensureExtendedLayout() {
    injectStyle();
    if ($("projectionGrid")) return;

    const media = $("mediaSerie") && $("mediaSerie").closest(".table-card");
    if (media) {
      media.insertAdjacentHTML("afterend", `
        <div class="section-title"><h2>Projecao de Receita</h2><span>Estimativa</span></div>
        <section class="finance-grid">
          <div class="panel">
            <div class="panel-head">
              <div class="panel-title"><span class="dot" style="background:var(--blue)"></span>Cenarios do mes</div>
              <div class="panel-note">base: quinzena + historico</div>
            </div>
            <div class="mini-grid" id="projectionGrid"></div>
          </div>
          <section class="table-card">
            <div class="table-toolbar">
              <div class="table-title">Faixas de projecao</div>
              <div class="hint">pessimista, media e otimista</div>
            </div>
            <div id="projectionDetail" class="metric-strip"></div>
          </section>
        </section>
      `);
    }

    const weekly = $("weeklyTable") && $("weeklyTable").closest(".table-card");
    if (weekly) {
      weekly.insertAdjacentHTML("beforebegin", `
        <div class="section-title"><h2>Demonstrativo de Drinks</h2><span>Resumo</span></div>
        <section class="insight-grid">
          <div class="panel">
            <div class="panel-head">
              <div class="panel-title"><span class="dot" style="background:var(--purple)"></span>Mix de bebidas</div>
              <div class="panel-note">faturamento e unidades</div>
            </div>
            <div class="mini-grid" id="drinkGrid"></div>
          </div>
          <section class="table-card compact">
            <div class="table-toolbar">
              <div class="table-title">Drinks por mes</div>
              <div class="hint">participacao no faturamento</div>
            </div>
            <div class="table-wrap"><table id="drinkTable"></table></div>
          </section>
        </section>
        <div class="section-title"><h2>Evolucao Financeira</h2><span>Mes a mes</span></div>
        <section class="insight-grid">
          <section class="table-card compact">
            <div class="table-toolbar">
              <div class="table-title">Quadro de custos</div>
              <div class="hint">valores e variacao</div>
            </div>
            <div class="table-wrap"><table id="costEvolutionTable"></table></div>
          </section>
          <section class="table-card compact">
            <div class="table-toolbar">
              <div class="table-title">Indicadores de performance</div>
              <div class="hint">KPIs financeiros</div>
            </div>
            <div class="table-wrap"><table id="performanceEvolutionTable"></table></div>
          </section>
        </section>
      `);
    }
  }

  function drinkCategory(group) {
    const g = norm(group);
    if (g.includes("double")) return "Double drinks";
    if (g.includes("especial")) return "Drinks especiais";
    if (g.includes("sem alcool") || g.includes("semalcool")) return "Sem alcool";
    if (g === "drinks" || g === "drink" || g.startsWith("drinks ")) return "Drinks";
    return "";
  }

  function drinkQty(row) {
    return row.qt_double || row.qt_final || row.qtd || 0;
  }

  summarize = function (rows) {
    const products = new Set();
    const days = new Set();
    let liquido = 0;
    let total = 0;
    let itens = 0;
    let vendas = 0;
    let drinks = 0;
    let doubleQty = 0;
    let drinkQtyTotal = 0;

    rows.forEach(row => {
      liquido += row.liquido;
      total += row.total;
      itens += row.qtd;
      vendas += row.qtven > 0 ? row.qtven : 0;
      products.add(row.produto);
      days.add([row.mes, Math.round(row.semana), row.dia].join("|"));
      doubleQty += row.qt_double;
      if (drinkCategory(row.grupo)) {
        drinks += row.liquido;
        drinkQtyTotal += drinkQty(row);
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
      drinkQty: drinkQtyTotal,
      ticket: vendas ? liquido / vendas : 0,
      mix: liquido ? drinks / liquido : 0
    };
  };

  function rowsIgnoringMonth() {
    const filters = currentFilters();
    const loose = { ...filters, mes: "" };
    return allData.filter(row => passes(row, loose));
  }

  function quinzena(row) {
    const q = norm(row.quinzena);
    if (q.includes("1")) return 1;
    if (q.includes("2")) return 2;
    const week = Math.round(row.semana || 0);
    if (week > 0 && week <= 2) return 1;
    if (week >= 3) return 2;
    return 0;
  }

  function buildProjection() {
    const filters = currentFilters();
    const baseRows = rowsIgnoringMonth();
    const months = sortMonths([...new Set(baseRows.map(row => row.mes).filter(Boolean))]);
    if (!months.length) return null;
    const anchor = filters.mes || months[months.length - 1];
    const currentRows = baseRows.filter(row => row.mes === anchor);
    const current = currentRows.reduce((acc, row) => acc + row.liquido, 0);
    const q1 = currentRows.filter(row => quinzena(row) === 1).reduce((acc, row) => acc + row.liquido, 0);
    const q2 = currentRows.filter(row => quinzena(row) === 2).reduce((acc, row) => acc + row.liquido, 0);

    const historical = months
      .filter(month => monthIndex(month) < monthIndex(anchor))
      .map(month => {
        const rows = baseRows.filter(row => row.mes === month);
        const total = rows.reduce((acc, row) => acc + row.liquido, 0);
        const first = rows.filter(row => quinzena(row) === 1).reduce((acc, row) => acc + row.liquido, 0);
        return { month, total, first, factor: first > 0 ? total / first : 0 };
      })
      .filter(row => row.factor > 1);

    if (!historical.length) {
      return { mes: anchor, q1, q2, current, min: current, avg: current, max: current, minFactor: 1, avgFactor: 1, maxFactor: 1, samples: 0 };
    }

    const factors = historical.map(row => row.factor);
    const minFactor = Math.min(...factors);
    const maxFactor = Math.max(...factors);
    const avgFactor = factors.reduce((acc, value) => acc + value, 0) / factors.length;
    const base = q1 || current;
    return {
      mes: anchor,
      q1,
      q2,
      current,
      min: base * minFactor,
      avg: base * avgFactor,
      max: base * maxFactor,
      minFactor,
      avgFactor,
      maxFactor,
      samples: historical.length
    };
  }

  function renderProjection() {
    if (!$("projectionGrid") || !$("projectionDetail")) return;
    const p = buildProjection();
    if (!p) {
      $("projectionGrid").innerHTML = '<div class="hint">Sem dados suficientes para projecao.</div>';
      $("projectionDetail").innerHTML = "";
      return;
    }
    const progress = p.avg > 0 ? p.current / p.avg : 0;
    $("projectionGrid").innerHTML =
      mini("Mes em projecao", p.mes, "periodo usado como referencia") +
      mini("1a quinzena", shortMoney(p.q1), "base da estimativa") +
      mini("2a quinzena", shortMoney(p.q2), p.q2 > 0 ? "valor ja realizado" : "ainda sem valor") +
      mini("Acumulado real", shortMoney(p.current), fmtPct(progress) + " da projecao media", progress >= 1 ? "pos" : "") +
      mini("Projecao media", shortMoney(p.avg), "fator historico x" + p.avgFactor.toFixed(2), "pos") +
      mini("Cenario otimista", shortMoney(p.max), "melhor fator x" + p.maxFactor.toFixed(2));

    const items = [
      ["Pessimista", shortMoney(p.min), "x" + p.minFactor.toFixed(2)],
      ["Media", shortMoney(p.avg), "x" + p.avgFactor.toFixed(2)],
      ["Otimista", shortMoney(p.max), "x" + p.maxFactor.toFixed(2)],
      ["Historico", nb(p.samples) + " meses", "base comparavel"],
      ["Falta p/ media", shortMoney(Math.max(0, p.avg - p.current)), p.current >= p.avg ? "meta atingida" : "saldo necessario"],
      ["Cobertura", fmtPct(progress), "real / projecao media"]
    ];
    $("projectionDetail").innerHTML = items.map(item => `<div class="metric-pill-card"><small>${item[0]}</small><strong>${item[1]}</strong><div class="card-sub">${item[2]}</div></div>`).join("");
  }

  function drinkSummary(rows) {
    const categories = new Map([
      ["Drinks", { value: 0, qty: 0 }],
      ["Drinks especiais", { value: 0, qty: 0 }],
      ["Sem alcool", { value: 0, qty: 0 }],
      ["Double drinks", { value: 0, qty: 0 }]
    ]);
    let total = 0;
    rows.forEach(row => {
      total += row.liquido;
      const cat = drinkCategory(row.grupo);
      if (!cat) return;
      const item = categories.get(cat) || { value: 0, qty: 0 };
      item.value += row.liquido;
      item.qty += drinkQty(row);
      categories.set(cat, item);
    });
    const drinkTotal = [...categories.values()].reduce((acc, item) => acc + item.value, 0);
    return { categories, total, drinkTotal, mix: total ? drinkTotal / total : 0 };
  }

  function renderDrinkDashboard() {
    if (!$("drinkGrid") || !$("drinkTable")) return;
    const data = drinkSummary(filteredData);
    const cards = [...data.categories.entries()].map(([label, item]) =>
      mini(label, shortMoney(item.value), nb(item.qty) + " unidades", item.value ? "" : "neg")
    );
    cards.push(mini("Participacao", fmtPct(data.mix), "drinks / faturamento", data.mix >= .2 ? "pos" : ""));
    $("drinkGrid").innerHTML = cards.join("");

    const months = sortMonths([...new Set(filteredData.map(row => row.mes).filter(Boolean))]);
    let html = '<thead><tr><th>Mes</th><th class="num">Drinks</th><th class="num">Especiais</th><th class="num">Sem alcool</th><th class="num">Double</th><th class="num">Total drinks</th><th class="num">% Fat.</th></tr></thead><tbody>';
    months.forEach(month => {
      const rows = filteredData.filter(row => row.mes === month);
      const monthData = drinkSummary(rows);
      const val = label => monthData.categories.get(label)?.value || 0;
      html += `<tr>
        <td><strong>${month}</strong></td>
        <td class="num">${money(val("Drinks"))}</td>
        <td class="num">${money(val("Drinks especiais"))}</td>
        <td class="num">${money(val("Sem alcool"))}</td>
        <td class="num">${money(val("Double drinks"))}</td>
        <td class="num subtle-total"><strong>${money(monthData.drinkTotal)}</strong></td>
        <td class="num"><strong>${fmtPct(monthData.mix)}</strong></td>
      </tr>`;
    });
    $("drinkTable").innerHTML = html + "</tbody>";
  }

  function valueFor(row, metric) {
    return metric.calc ? metric.calc(row) : Number(row[metric.key] || 0);
  }

  function deltaHtml(value, previous, goodUp) {
    if (!previous) return '<span class="delta neutral">base</span>';
    const diff = (value - previous) / Math.abs(previous);
    if (Math.abs(diff) < .005) return '<span class="delta neutral">0.0%</span>';
    const up = diff > 0;
    const good = up === goodUp;
    return `<span class="delta ${good ? "good" : "bad"}">${up ? "+" : "-"}${Math.abs(diff * 100).toFixed(1)}%</span>`;
  }

  function evolutionTable(rows, metrics) {
    const ordered = [...rows].sort((a, b) => monthIndex(a.mes) - monthIndex(b.mes));
    let html = '<thead><tr><th>Indicador</th>' + ordered.map(row => `<th class="num">${row.mes}</th>`).join("") + "</tr></thead><tbody>";
    metrics.forEach(metric => {
      html += `<tr><td><strong>${metric.label}</strong></td>`;
      ordered.forEach((row, index) => {
        const value = valueFor(row, metric);
        const previous = index > 0 ? valueFor(ordered[index - 1], metric) : 0;
        html += `<td class="num"><div class="evo-cell"><strong>${metric.format(value)}</strong>${deltaHtml(value, previous, metric.goodUp)}</div></td>`;
      });
      html += "</tr>";
    });
    return html + "</tbody>";
  }

  function renderEvolutionTables() {
    if (!$("costEvolutionTable") || !$("performanceEvolutionTable")) return;
    const rows = matrizFiltered().filter(row => row.receita || row.custoGeral || row.lucPrej);
    if (!rows.length) {
      $("costEvolutionTable").innerHTML = '<tbody><tr><td>Sem matriz analitica para o periodo.</td></tr></tbody>';
      $("performanceEvolutionTable").innerHTML = "";
      return;
    }
    const pct = value => fmtPct(asPctValue(value));
    const costMetrics = [
      { label: "Receita", key: "receita", format: shortMoney, goodUp: true },
      { label: "Custo geral", key: "custoGeral", format: shortMoney, goodUp: false },
      { label: "Custo fixo", key: "custoFixo", format: shortMoney, goodUp: false },
      { label: "Fixo / variavel", key: "custoFixoVar", format: shortMoney, goodUp: false },
      { label: "Total fixo", key: "totalFixo", format: shortMoney, goodUp: false },
      { label: "Custo variavel", calc: row => Math.max(0, row.custoGeral - row.totalFixo), format: shortMoney, goodUp: false },
      { label: "CMV", key: "cmv", format: shortMoney, goodUp: false }
    ];
    const performanceMetrics = [
      { label: "Lucro / prejuizo", key: "lucPrej", format: shortMoney, goodUp: true },
      { label: "Lucratividade", calc: row => row.receita ? row.lucPrej / row.receita : 0, format: pct, goodUp: true },
      { label: "CMV %", key: "pctCmv", format: pct, goodUp: false },
      { label: "MKP", key: "mkp", format: fmtX, goodUp: true },
      { label: "Margem contrib.", key: "margem", format: pct, goodUp: true },
      { label: "Cobertura fixo", calc: row => row.totalFixo ? row.receita / row.totalFixo : 0, format: pct, goodUp: true }
    ];
    $("costEvolutionTable").innerHTML = evolutionTable(rows, costMetrics);
    $("performanceEvolutionTable").innerHTML = evolutionTable(rows, performanceMetrics);
  }

  renderAll = function () {
    const summary = summarize(filteredData);
    const finance = financeSummary(summary);
    renderStatus(summary);
    renderKPIs(summary, finance);
    renderFinanceiro(summary, finance);
    renderMediaSerie();
    ensureExtendedLayout();
    renderProjection();
    renderDrinkDashboard();
    renderEvolutionTables();
    renderCharts();
    renderWeekly();
    renderProducts();
  };
})();