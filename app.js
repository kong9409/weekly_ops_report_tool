const BANK = Array.isArray(window.REASON_BANK) ? window.REASON_BANK : [];
const CATEGORIES = ["市场", "竞品", "产品", "广告", "售后", "其他"];
const WEEK_NAMES = ["第一周", "第二周", "第三周", "第四周"];
const ISSUE_KEYS = BANK.map((item, index) => ({
  index,
  key: `${item.dimension}||${item.issue}`,
  label: `${item.dimension} / ${item.issue}`,
}));

const state = {
  fileName: "",
  latestDate: null,
  monthLabel: "",
  current: null,
  previous: null,
  month: null,
  products: [],
  metricCounter: 0,
  analysisCounter: 0,
};

const $ = (id) => document.getElementById(id);

document.addEventListener("DOMContentLoaded", () => {
  $("file").addEventListener("change", handleFile);
  $("loadSample").addEventListener("click", loadSampleState);
  $("addMetric").addEventListener("click", () => addManualMetric());
  $("addAnalysis").addEventListener("click", () => addAnalysisItem());
  $("generate").addEventListener("click", generateReport);
  $("downloadDoc").addEventListener("click", downloadWord);
  $("downloadTxt").addEventListener("click", downloadText);
  $("copy").addEventListener("click", copyReport);
  $("print").addEventListener("click", () => {
    generateReport();
    window.print();
  });
});

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[char]));
}

function normalizeHeader(value) {
  return String(value ?? "")
    .replace(/[\s\n\r（）()：:%％$￥¥_\-–—]/g, "")
    .toLowerCase();
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const number = Number.parseFloat(String(value ?? "").replace(/[$￥¥,，%\s]/g, ""));
  return Number.isFinite(number) ? number : 0;
}

function parseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && window.XLSX?.SSF) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const text = String(value ?? "").trim().replace(/\//g, "-");
  const match = text.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!match) return null;
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  return Number.isNaN(date.getTime()) ? null : date;
}

function weekIndex(date) {
  if (date.getDate() <= 7) return 0;
  if (date.getDate() <= 14) return 1;
  if (date.getDate() <= 21) return 2;
  return 3;
}

function formatDate(date) {
  if (!date) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function formatMoney(value) {
  return `$${Number(value || 0).toLocaleString("zh-CN", { maximumFractionDigits: 2 })}`;
}

function formatNumber(value, digits = 0) {
  return Number(value || 0).toLocaleString("zh-CN", { maximumFractionDigits: digits });
}

function formatPercent(value, digits = 2) {
  return `${(Number(value || 0) * 100).toFixed(digits)}%`;
}

function formatChange(value) {
  if (value === null || !Number.isFinite(value)) return "无可比数据";
  const sign = value > 0 ? "+" : "";
  return `${sign}${(value * 100).toFixed(1)}%`;
}

function findHeaderRow(rows) {
  let bestIndex = 0;
  let bestScore = -1;
  const keys = ["日期", "分类", "实际销量", "加运费销售额", "二级利润", "广告费", "会话次数", "展示量"];
  rows.slice(0, 12).forEach((row, index) => {
    const text = row.map(normalizeHeader).join("|");
    const score = keys.filter((key) => text.includes(normalizeHeader(key))).length;
    if (score > bestScore) {
      bestIndex = index;
      bestScore = score;
    }
  });
  return bestIndex;
}

function findColumn(header, aliases) {
  const normalized = header.map(normalizeHeader);
  for (const alias of aliases) {
    const target = normalizeHeader(alias);
    const exact = normalized.findIndex((item) => item === target);
    if (exact >= 0) return exact;
  }
  for (const alias of aliases) {
    const target = normalizeHeader(alias);
    const partial = normalized.findIndex((item) => item.includes(target) || target.includes(item));
    if (partial >= 0) return partial;
  }
  return -1;
}

function emptyAggregate(label, index) {
  return {
    label,
    index,
    dates: new Set(),
    units: 0,
    sales: 0,
    profit: 0,
    adFee: 0,
    sessions: 0,
    impressions: 0,
    clicks: 0,
    adOrders: 0,
    adSpend: 0,
    adSales: 0,
    stock: 0,
    transit: 0,
    products: new Map(),
  };
}

function addValues(target, values, productName) {
  target.dates.add(values.dateKey);
  [
    "units",
    "sales",
    "profit",
    "adFee",
    "sessions",
    "impressions",
    "clicks",
    "adOrders",
    "adSpend",
    "adSales",
    "stock",
    "transit",
  ].forEach((key) => {
    target[key] += values[key];
  });

  if (!target.products.has(productName)) {
    target.products.set(productName, emptyAggregate(productName, -1));
  }
  const product = target.products.get(productName);
  product.dates.add(values.dateKey);
  [
    "units",
    "sales",
    "profit",
    "adFee",
    "sessions",
    "impressions",
    "clicks",
    "adOrders",
    "adSpend",
    "adSales",
  ].forEach((key) => {
    product[key] += values[key];
  });
}

function finalizeAggregate(aggregate) {
  const result = { ...aggregate };
  result.dayCount = aggregate.dates.size || 1;
  result.aov = aggregate.units ? aggregate.sales / aggregate.units : 0;
  result.margin = aggregate.sales ? aggregate.profit / aggregate.sales : 0;
  result.tacos = aggregate.sales ? aggregate.adFee / aggregate.sales : 0;
  result.cvr = aggregate.sessions ? aggregate.units / aggregate.sessions : 0;
  result.ctr = aggregate.impressions ? aggregate.clicks / aggregate.impressions : 0;
  result.adCvr = aggregate.clicks ? aggregate.adOrders / aggregate.clicks : 0;
  result.acos = aggregate.adSales ? aggregate.adSpend / aggregate.adSales : 0;
  result.salesDaily = aggregate.sales / result.dayCount;
  result.unitsDaily = aggregate.units / result.dayCount;
  result.sessionsDaily = aggregate.sessions / result.dayCount;
  result.impressionsDaily = aggregate.impressions / result.dayCount;
  result.products = [...aggregate.products.entries()]
    .map(([name, item]) => ({ name, ...finalizeAggregate({ ...item, products: new Map() }) }))
    .sort((a, b) => b.sales - a.sales);
  return result;
}

async function handleFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  try {
    if (!window.XLSX) throw new Error("Excel 解析组件未加载。");
    setStatus("正在读取绩效复盘表，请稍候……", "");
    const workbook = XLSX.read(await file.arrayBuffer(), {
      type: "array",
      cellDates: true,
    });
    const sheetName = workbook.SheetNames.find((name) => name.includes("每日数据"));
    if (!sheetName) throw new Error("未找到“每日数据”工作表，请确认上传的是绩效复盘表。");
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: null,
      raw: true,
    });
    const result = analyzeRows(rows);
    state.fileName = file.name;
    applyAnalysisResult(result);
    setStatus(`读取成功：${file.name}`, "success");
  } catch (error) {
    console.error(error);
    setStatus(`读取失败：${error.message}`, "error");
  }
}

function analyzeRows(rows) {
  const headerRow = findHeaderRow(rows);
  const header = rows[headerRow] || [];
  const columns = {
    product: findColumn(header, ["分类", "产品", "型号"]),
    date: findColumn(header, ["日期", "Date", "开始日期"]),
    units: findColumn(header, ["实际销量", "销量"]),
    sales: findColumn(header, ["加运费销售额", "销售额"]),
    profit: findColumn(header, ["二级利润"]),
    adFee: findColumn(header, ["广告费"]),
    sessions: findColumn(header, ["会话次数 – 总计", "会话次数总计", "访客", "Sessions"]),
    impressions: findColumn(header, ["展示量", "Impressions"]),
    clicks: findColumn(header, ["点击量", "Clicks"]),
    adOrders: findColumn(header, ["订单数", "广告订单"]),
    adSpend: findColumn(header, ["花费", "Spend"]),
    adSales: findColumn(header, ["广告销售额", "AdSales"]),
    stock: findColumn(header, ["在售"]),
    transit: findColumn(header, ["在途"]),
  };

  const required = ["product", "date", "units", "sales"];
  const missing = required.filter((key) => columns[key] < 0);
  if (missing.length) {
    throw new Error(`缺少必要字段：${missing.join("、")}。`);
  }

  const parsedRows = rows.slice(headerRow + 1).map((row) => {
    const date = parseDate(row[columns.date]);
    const product = String(row[columns.product] ?? "").trim();
    if (!date || !product || product === "合计" || product === "总计") return null;
    return { row, date, product };
  }).filter(Boolean);

  if (!parsedRows.length) throw new Error("“每日数据”中没有可识别的日期和产品数据。");

  const latestDate = parsedRows.reduce((latest, item) => item.date > latest ? item.date : latest, parsedRows[0].date);
  const monthRows = parsedRows.filter((item) =>
    item.date.getFullYear() === latestDate.getFullYear()
    && item.date.getMonth() === latestDate.getMonth()
  );

  const weeks = WEEK_NAMES.map((label, index) => emptyAggregate(label, index));
  const month = emptyAggregate(`${latestDate.getFullYear()}年${latestDate.getMonth() + 1}月`, -1);

  monthRows.forEach(({ row, date, product }) => {
    const read = (key) => columns[key] >= 0 ? toNumber(row[columns[key]]) : 0;
    const values = {
      dateKey: formatDate(date),
      units: read("units"),
      sales: read("sales"),
      profit: read("profit"),
      adFee: read("adFee"),
      sessions: read("sessions"),
      impressions: read("impressions"),
      clicks: read("clicks"),
      adOrders: read("adOrders"),
      adSpend: read("adSpend"),
      adSales: read("adSales"),
      stock: read("stock"),
      transit: read("transit"),
    };
    addValues(weeks[weekIndex(date)], values, product);
    addValues(month, values, product);
  });

  const populatedWeeks = weeks.filter((week) => week.dates.size);
  const currentRaw = populatedWeeks[populatedWeeks.length - 1];
  const previousRaw = populatedWeeks.length > 1 ? populatedWeeks[populatedWeeks.length - 2] : null;

  return {
    latestDate,
    monthLabel: `${latestDate.getFullYear()}年${latestDate.getMonth() + 1}月`,
    current: finalizeAggregate(currentRaw),
    previous: previousRaw ? finalizeAggregate(previousRaw) : null,
    month: finalizeAggregate(month),
  };
}

function applyAnalysisResult(result) {
  state.latestDate = result.latestDate;
  state.monthLabel = result.monthLabel;
  state.current = result.current;
  state.previous = result.previous;
  state.month = result.month;
  state.products = result.month.products.map((item) => item.name);

  renderSummary();
  renderMetrics();
  renderRecommendedAnalysis();
  $("periodInfo").classList.remove("hidden");
  $("periodInfo").textContent = state.previous
    ? `数据期间：${state.monthLabel}；最新周为${state.current.label}（${state.current.dayCount}天），对比${state.previous.label}（${state.previous.dayCount}天）。总量指标按日均环比。`
    : `数据期间：${state.monthLabel}；当前仅识别到${state.current.label}，暂不计算周环比。`;
}

function setStatus(message, type) {
  $("status").className = `status ${type || "muted"}`;
  $("status").textContent = message;
}

function change(currentValue, previousValue) {
  if (!state.previous || !Number.isFinite(previousValue) || previousValue === 0) return null;
  return (currentValue - previousValue) / Math.abs(previousValue);
}

function pointChange(currentValue, previousValue) {
  if (!state.previous || !Number.isFinite(previousValue)) return null;
  return currentValue - previousValue;
}

function trendClass(value, reverse = false) {
  if (value === null || Math.abs(value) < 0.001) return "";
  const positive = reverse ? value < 0 : value > 0;
  return positive ? "trend-up" : "trend-down";
}

function renderSummary() {
  const current = state.current;
  const previous = state.previous;
  const cards = [
    {
      label: `${current.label} GMV`,
      value: formatMoney(current.sales),
      delta: previous ? change(current.salesDaily, previous.salesDaily) : null,
      note: "日均环比",
    },
    {
      label: `${current.label}销量`,
      value: formatNumber(current.units),
      delta: previous ? change(current.unitsDaily, previous.unitsDaily) : null,
      note: "日均环比",
    },
    {
      label: "二级利润率",
      value: formatPercent(current.margin),
      delta: previous ? pointChange(current.margin, previous.margin) : null,
      note: "较上周",
      points: true,
    },
    {
      label: "TACOS",
      value: formatPercent(current.tacos),
      delta: previous ? pointChange(current.tacos, previous.tacos) : null,
      note: "较上周",
      points: true,
      reverse: true,
    },
    {
      label: "访客 / 会话",
      value: formatNumber(current.sessions),
      delta: previous ? change(current.sessionsDaily, previous.sessionsDaily) : null,
      note: "日均环比",
    },
    {
      label: "综合转化率",
      value: formatPercent(current.cvr),
      delta: previous ? change(current.cvr, previous.cvr) : null,
      note: "环比",
    },
    {
      label: "广告CTR",
      value: formatPercent(current.ctr),
      delta: previous ? change(current.ctr, previous.ctr) : null,
      note: "环比",
    },
    {
      label: "广告转化率",
      value: formatPercent(current.adCvr),
      delta: previous ? change(current.adCvr, previous.adCvr) : null,
      note: "环比",
    },
  ];

  $("summary").innerHTML = cards.map((card) => {
    const displayDelta = card.delta === null
      ? "无可比数据"
      : card.points
        ? `${card.delta > 0 ? "+" : ""}${(card.delta * 100).toFixed(2)}个百分点`
        : formatChange(card.delta);
    return `
      <div class="kpi">
        <span>${escapeHtml(card.label)}</span>
        <strong>${escapeHtml(card.value)}</strong>
        <small class="${trendClass(card.delta, card.reverse)}">${escapeHtml(card.note)} ${escapeHtml(displayDelta)}</small>
      </div>
    `;
  }).join("");
}

function metricDefinitions() {
  const current = state.current;
  const previous = state.previous;
  const definitions = [
    {
      name: "GMV",
      value: formatMoney(current.sales),
      detail: previous
        ? `日均 ${formatMoney(current.salesDaily)}，较${previous.label}${formatChange(change(current.salesDaily, previous.salesDaily))}`
        : "暂无上周数据",
      delta: previous ? change(current.salesDaily, previous.salesDaily) : null,
      direction: "normal",
    },
    {
      name: "销量",
      value: `${formatNumber(current.units)}件`,
      detail: previous
        ? `日均 ${formatNumber(current.unitsDaily, 1)}件，较${previous.label}${formatChange(change(current.unitsDaily, previous.unitsDaily))}`
        : "暂无上周数据",
      delta: previous ? change(current.unitsDaily, previous.unitsDaily) : null,
      direction: "normal",
    },
    {
      name: "平均客单价",
      value: formatMoney(current.aov),
      detail: previous ? `较${previous.label}${formatChange(change(current.aov, previous.aov))}` : "暂无上周数据",
      delta: previous ? change(current.aov, previous.aov) : null,
      direction: "normal",
    },
    {
      name: "二级利润率",
      value: formatPercent(current.margin),
      detail: previous
        ? `较${previous.label}${(pointChange(current.margin, previous.margin) * 100).toFixed(2)}个百分点`
        : "暂无上周数据",
      delta: previous ? change(current.margin, previous.margin) : null,
      direction: "normal",
    },
    {
      name: "访客 / 会话",
      value: formatNumber(current.sessions),
      detail: previous
        ? `日均 ${formatNumber(current.sessionsDaily)}，较${previous.label}${formatChange(change(current.sessionsDaily, previous.sessionsDaily))}`
        : "暂无上周数据",
      delta: previous ? change(current.sessionsDaily, previous.sessionsDaily) : null,
      direction: "normal",
    },
    {
      name: "综合转化率",
      value: formatPercent(current.cvr),
      detail: previous ? `较${previous.label}${formatChange(change(current.cvr, previous.cvr))}` : "暂无上周数据",
      delta: previous ? change(current.cvr, previous.cvr) : null,
      direction: "normal",
    },
    {
      name: "广告点击率",
      value: formatPercent(current.ctr),
      detail: previous ? `较${previous.label}${formatChange(change(current.ctr, previous.ctr))}` : "暂无上周数据",
      delta: previous ? change(current.ctr, previous.ctr) : null,
      direction: "normal",
    },
    {
      name: "广告转化率",
      value: formatPercent(current.adCvr),
      detail: previous ? `较${previous.label}${formatChange(change(current.adCvr, previous.adCvr))}` : "暂无上周数据",
      delta: previous ? change(current.adCvr, previous.adCvr) : null,
      direction: "normal",
    },
    {
      name: "广告TACOS",
      value: formatPercent(current.tacos),
      detail: previous
        ? `较${previous.label}${(pointChange(current.tacos, previous.tacos) * 100).toFixed(2)}个百分点`
        : "暂无上周数据",
      delta: previous ? change(current.tacos, previous.tacos) : null,
      direction: "reverse",
    },
  ];
  return definitions;
}

function renderMetrics() {
  const metrics = metricDefinitions();
  $("metrics").className = "";
  $("metrics").innerHTML = metrics.map((metric) => {
    const abnormal = metric.delta === null
      || Math.abs(metric.delta) >= 0.1
      || (metric.direction === "reverse" && metric.delta >= 0.08);
    const bad = metric.delta !== null
      && (metric.direction === "reverse" ? metric.delta > 0 : metric.delta < 0);
    return `
      <div class="metric-row" data-generated="true">
        <input type="checkbox" class="metric-check" ${abnormal ? "checked" : ""}>
        <div class="metric-main">
          <strong class="${bad ? "trend-down" : ""}">${escapeHtml(metric.name)}：${escapeHtml(metric.value)}</strong>
          <span>${escapeHtml(metric.detail)}</span>
        </div>
        <textarea class="metric-note" rows="2" placeholder="补充具体产品、活动、竞品或异常背景"></textarea>
      </div>
    `;
  }).join("");
}

function addManualMetric() {
  if ($("metrics").classList.contains("empty")) {
    $("metrics").className = "";
    $("metrics").innerHTML = "";
  }
  state.metricCounter += 1;
  $("metrics").insertAdjacentHTML("beforeend", `
    <div class="metric-row" data-generated="false">
      <input type="checkbox" class="metric-check" checked>
      <div class="metric-main">
        <input type="text" class="manual-metric-name" value="自定义指标" aria-label="自定义指标名称">
        <input type="text" class="manual-metric-value" placeholder="例如：下降12.5%" aria-label="自定义指标变化">
      </div>
      <textarea class="metric-note" rows="2" placeholder="补充具体产品、活动、竞品或异常背景"></textarea>
    </div>
  `);
}

function recommendedIssues() {
  if (!state.current) return [];
  if (!state.previous) return [{ issueKey: "综合转化率||下降", product: "整体" }];
  const current = state.current;
  const previous = state.previous;
  const recommendations = [];
  const add = (issueKey, product = "整体") => {
    if (!recommendations.some((item) => item.issueKey === issueKey && item.product === product)) {
      recommendations.push({ issueKey, product });
    }
  };

  if (change(current.sessionsDaily, previous.sessionsDaily) <= -0.08) add("访客||下降");
  if (change(current.cvr, previous.cvr) <= -0.08) add("综合转化率||下降");
  if (change(current.ctr, previous.ctr) <= -0.08) add("广告点击率||下降");
  if (change(current.adCvr, previous.adCvr) <= -0.08) add("广告转化率||下降");
  if (pointChange(current.tacos, previous.tacos) >= 0.01 || current.tacos >= 0.12) add("广告TACOS||占比高");

  if (change(current.salesDaily, previous.salesDaily) <= -0.1 && !recommendations.length) {
    add("综合转化率||下降");
  }

  current.products
    .filter((product) => product.sales > 0 && product.tacos >= Math.max(current.tacos * 1.5, 0.15))
    .slice(0, 2)
    .forEach((product) => add("广告TACOS||占比高", product.name));

  return recommendations.slice(0, 5);
}

function renderRecommendedAnalysis() {
  $("analysisList").innerHTML = "";
  const recommendations = recommendedIssues();
  if (!recommendations.length) {
    addAnalysisItem();
    return;
  }
  recommendations.forEach((item) => addAnalysisItem(item));
}

function splitReasons(item) {
  const text = String(item.check_methods || "")
    .replace(/\r/g, "")
    .replace(/[ \t]{2,}(?=(?:[A-Z]|\d{1,2})[.、])/g, "\n")
    .replace(/(?<!^)(?=(?:9|10)[.、]\s*)/g, "\n");
  const reasons = text.split("\n").map((line) => line.trim()).filter(Boolean);
  if (item.thinking) reasons.unshift(`综合判断：${item.thinking}`);
  return [...new Set(reasons)];
}

function productOptions(selected) {
  return ["整体", ...state.products].map((product) =>
    `<option value="${escapeHtml(product)}" ${product === selected ? "selected" : ""}>${escapeHtml(product)}</option>`
  ).join("");
}

function issueOptions(selected) {
  return ISSUE_KEYS.map((item) =>
    `<option value="${escapeHtml(item.key)}" ${item.key === selected ? "selected" : ""}>${escapeHtml(item.label)}</option>`
  ).join("");
}

function addAnalysisItem(initial = {}) {
  if (!BANK.length) {
    $("analysisList").innerHTML = '<div class="empty">原因选项库加载失败，请重新打开工具。</div>';
    return;
  }
  if ($("analysisList").querySelector(".empty")) $("analysisList").innerHTML = "";

  state.analysisCounter += 1;
  const id = state.analysisCounter;
  const issueKey = initial.issueKey || ISSUE_KEYS[0].key;
  const product = initial.product || "整体";
  $("analysisList").insertAdjacentHTML("beforeend", `
    <div class="analysis-item" data-id="${id}">
      <div class="analysis-head">
        <strong>分析项 ${id}</strong>
        <button type="button" class="button danger remove-analysis">删除</button>
      </div>
      <div class="form-grid">
        <div class="field">
          <label>涉及产品</label>
          <select class="analysis-product">${productOptions(product)}</select>
        </div>
        <div class="field">
          <label>原因维度</label>
          <select class="analysis-issue">${issueOptions(issueKey)}</select>
        </div>
        <div class="field field-wide">
          <label>差距分析（从运营项目周报选择原因）</label>
          <select class="analysis-reason"></select>
        </div>
        <div class="field">
          <label>关键动作分类</label>
          <select class="action-category"></select>
        </div>
        <div class="field">
          <label>关键动作</label>
          <select class="action-option"></select>
        </div>
        <div class="field">
          <label>下周计划分类</label>
          <select class="plan-category"></select>
        </div>
        <div class="field">
          <label>下周计划</label>
          <select class="plan-option"></select>
        </div>
        <div class="field field-wide">
          <label>补充说明</label>
          <textarea class="analysis-note" rows="3" placeholder="补充具体产品、SKU、竞品价格、广告活动、时间节点、负责人或目标值"></textarea>
        </div>
      </div>
    </div>
  `);

  const element = $("analysisList").querySelector(`[data-id="${id}"]`);
  element.querySelector(".remove-analysis").addEventListener("click", () => element.remove());
  element.querySelector(".analysis-issue").addEventListener("change", () => refreshAnalysisItem(element));
  element.querySelector(".action-category").addEventListener("change", () => refreshActionSelect(element, "action"));
  element.querySelector(".plan-category").addEventListener("change", () => refreshActionSelect(element, "plan"));
  refreshAnalysisItem(element);
}

function bankItemFor(element) {
  const key = element.querySelector(".analysis-issue").value;
  return BANK.find((item) => `${item.dimension}||${item.issue}` === key) || BANK[0];
}

function refreshAnalysisItem(element) {
  const item = bankItemFor(element);
  const reasonSelect = element.querySelector(".analysis-reason");
  reasonSelect.innerHTML = splitReasons(item)
    .map((reason) => `<option value="${escapeHtml(reason)}">${escapeHtml(reason)}</option>`)
    .join("");

  ["action", "plan"].forEach((type) => {
    const categorySelect = element.querySelector(`.${type}-category`);
    const available = [...new Set((item.action_options || []).map((option) => option.category))];
    categorySelect.innerHTML = CATEGORIES.map((category) =>
      `<option value="${category}" ${available.includes(category) ? "" : "disabled"}>${category}</option>`
    ).join("");
    categorySelect.value = available[0] || "其他";
    refreshActionSelect(element, type);
  });
}

function refreshActionSelect(element, type) {
  const item = bankItemFor(element);
  const category = element.querySelector(`.${type}-category`).value;
  const select = element.querySelector(`.${type}-option`);
  const options = (item.action_options || []).filter((option) => option.category === category);
  select.innerHTML = options.length
    ? options.map((option) => `<option value="${escapeHtml(option.text)}">${escapeHtml(option.text)}</option>`).join("")
    : '<option value="">该分类暂无选项</option>';
}

function selectedMetrics() {
  return [...$("metrics").querySelectorAll(".metric-row")]
    .filter((row) => row.querySelector(".metric-check")?.checked)
    .map((row) => {
      const generated = row.dataset.generated === "true";
      const note = row.querySelector(".metric-note")?.value.trim() || "";
      if (generated) {
        return {
          title: row.querySelector(".metric-main strong")?.textContent.trim() || "",
          detail: row.querySelector(".metric-main span")?.textContent.trim() || "",
          note,
        };
      }
      return {
        title: `${row.querySelector(".manual-metric-name")?.value.trim() || "自定义指标"}：${row.querySelector(".manual-metric-value")?.value.trim() || "待补充"}`,
        detail: "",
        note,
      };
    });
}

function selectedAnalysis() {
  return [...$("analysisList").querySelectorAll(".analysis-item")].map((element) => {
    const item = bankItemFor(element);
    return {
      product: element.querySelector(".analysis-product").value,
      dimension: item.dimension,
      issue: item.issue,
      reason: element.querySelector(".analysis-reason").value,
      actionCategory: element.querySelector(".action-category").value,
      action: element.querySelector(".action-option").value,
      planCategory: element.querySelector(".plan-category").value,
      plan: element.querySelector(".plan-option").value,
      note: element.querySelector(".analysis-note").value.trim(),
    };
  });
}

function autoProgressText() {
  if (!state.current) return "尚未上传绩效复盘表。";
  const current = state.current;
  const month = state.month;
  return `${current.label}完成GMV ${formatMoney(current.sales)}、销量 ${formatNumber(current.units)}，二级利润率 ${formatPercent(current.margin)}，TACOS ${formatPercent(current.tacos)}。`
    + `${state.monthLabel}累计GMV ${formatMoney(month.sales)}、销量 ${formatNumber(month.units)}，二级利润率 ${formatPercent(month.margin)}。`;
}

function buildComparisonRows() {
  if (!state.current) return [];
  const current = state.current;
  const previous = state.previous;
  return [
    ["GMV", formatMoney(current.sales), previous ? formatMoney(previous.sales) : "-"],
    ["销量", formatNumber(current.units), previous ? formatNumber(previous.units) : "-"],
    ["平均客单价", formatMoney(current.aov), previous ? formatMoney(previous.aov) : "-"],
    ["二级利润率", formatPercent(current.margin), previous ? formatPercent(previous.margin) : "-"],
    ["TACOS", formatPercent(current.tacos), previous ? formatPercent(previous.tacos) : "-"],
    ["访客 / 会话", formatNumber(current.sessions), previous ? formatNumber(previous.sessions) : "-"],
    ["综合转化率", formatPercent(current.cvr), previous ? formatPercent(previous.cvr) : "-"],
    ["广告CTR", formatPercent(current.ctr), previous ? formatPercent(previous.ctr) : "-"],
    ["广告转化率", formatPercent(current.adCvr), previous ? formatPercent(previous.adCvr) : "-"],
  ];
}

function groupedList(items, categoryKey, textKey) {
  const groups = CATEGORIES.map((category) => ({
    category,
    items: items.filter((item) => item[categoryKey] === category && item[textKey]),
  })).filter((group) => group.items.length);
  if (!groups.length) return "<p>（请选择具体内容）</p>";
  return groups.map((group) => `
    <p><strong>${escapeHtml(group.category)}</strong></p>
    <ul>${group.items.map((item) => `
      <li>${escapeHtml(item.product)}：${escapeHtml(item[textKey])}${item.note ? `；补充：${escapeHtml(item.note)}` : ""}</li>
    `).join("")}</ul>
  `).join("");
}

function generateReport() {
  const metrics = selectedMetrics();
  const analysis = selectedAnalysis();
  const target = $("target").value.trim() || "（请填写本月目标）";
  const progress = $("progress").value.trim();
  const period = state.current
    ? `${state.monthLabel} ${state.current.label}（数据截至 ${formatDate(state.latestDate)}）`
    : "未上传数据";
  const comparisonRows = buildComparisonRows();

  $("report").innerHTML = `
    <h1>运营项目周报</h1>
    <div class="report-period">${escapeHtml(period)}</div>

    <h3>一、本月目标</h3>
    <p>${escapeHtml(target)}</p>

    <h3>二、本周完成进度</h3>
    <p>${escapeHtml(progress || autoProgressText())}</p>
    ${comparisonRows.length ? `
      <table class="data-table">
        <thead>
          <tr>
            <th>指标</th>
            <th>${escapeHtml(state.current.label)}</th>
            <th>${escapeHtml(state.previous?.label || "对比周")}</th>
          </tr>
        </thead>
        <tbody>
          ${comparisonRows.map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`).join("")}
        </tbody>
      </table>
    ` : ""}

    <h3>三、主要指标波动</h3>
    ${metrics.length ? `<ul>${metrics.map((metric) => `
      <li><strong>${escapeHtml(metric.title)}</strong>${metric.detail ? `，${escapeHtml(metric.detail)}` : ""}${metric.note ? `；补充：${escapeHtml(metric.note)}` : ""}</li>
    `).join("")}</ul>` : "<p>（请勾选主要指标波动）</p>"}

    <h3>四、差距分析（遇到的问题）</h3>
    ${analysis.length ? `<ul>${analysis.map((item) => `
      <li><strong>【${escapeHtml(item.product)}｜${escapeHtml(item.dimension)} / ${escapeHtml(item.issue)}】</strong>${escapeHtml(item.reason)}${item.note ? `；补充：${escapeHtml(item.note)}` : ""}</li>
    `).join("")}</ul>` : "<p>（请新增并选择分析项）</p>"}

    <h3>五、关键动作</h3>
    ${groupedList(analysis, "actionCategory", "action")}

    <h3>六、下周计划（接下来怎么做）</h3>
    ${groupedList(analysis, "planCategory", "plan")}

    <h3>七、说明</h3>
    <p>数据来源：${escapeHtml(state.fileName || "绩效复盘表")}；原因、关键动作及下周计划选项来源：《运营项目周报.xlsx》。</p>
  `;
}

function reportFileName(extension) {
  const month = state.monthLabel || "未命名月份";
  const week = state.current?.label || "周报";
  return `运营项目周报_${month}_${week}.${extension}`;
}

function ensureReport() {
  if (!$("report").querySelector("h1")) generateReport();
}

function downloadBlob(content, type, fileName) {
  const blob = new Blob([content], { type });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function downloadWord() {
  generateReport();
  const documentHtml = `
    <!doctype html>
    <html lang="zh-CN">
    <head>
      <meta charset="utf-8">
      <style>
        body{font-family:"Microsoft YaHei",Arial,sans-serif;color:#202939;line-height:1.65;padding:28px}
        h1{text-align:center}h3{color:#173d92;border-bottom:2px solid #dce7ff;padding-bottom:6px}
        table{width:100%;border-collapse:collapse}th,td{border:1px solid #ccd3df;padding:8px;text-align:left}
        th{background:#f1f5fb}li{margin:5px 0}
      </style>
    </head>
    <body>${$("report").innerHTML}</body>
    </html>
  `;
  downloadBlob(`\ufeff${documentHtml}`, "application/msword;charset=utf-8", reportFileName("doc"));
}

function downloadText() {
  generateReport();
  downloadBlob(`\ufeff${$("report").innerText}`, "text/plain;charset=utf-8", reportFileName("txt"));
}

async function copyReport() {
  generateReport();
  try {
    await navigator.clipboard.writeText($("report").innerText);
    setStatus("周报文本已复制。", "success");
  } catch {
    setStatus("浏览器未允许自动复制，请在报告区域手动选择文本。", "error");
  }
}

function loadSampleState() {
  const sample = {
    latestDate: new Date(2026, 3, 30),
    monthLabel: "2026年4月",
    previous: sampleAggregate("第三周", 2, 7, {
      units: 78,
      sales: 76254.99,
      profit: 3894.69,
      adFee: 4303.27,
      sessions: 19433,
      impressions: 356205,
      clicks: 2512,
      adOrders: 28,
      adSpend: 3586.17,
      adSales: 13965.39,
    }),
    current: sampleAggregate("第四周", 3, 9, {
      units: 82,
      sales: 63448.9,
      profit: 2941.43,
      adFee: 5608.6,
      sessions: 23718,
      impressions: 417794,
      clicks: 3266,
      adOrders: 49,
      adSpend: 4719.58,
      adSales: 19703.57,
    }),
    month: sampleAggregate("2026年4月", -1, 30, {
      units: 313,
      sales: 268020.64,
      profit: 13790.19,
      adFee: 18344.97,
      sessions: 78483,
      impressions: 1736073,
      clicks: 13666,
      adOrders: 159,
      adSpend: 15542.69,
      adSales: 66925.16,
    }),
  };
  const productNames = ["产品1", "产品4", "产品13", "产品18", "产品25", "产品26", "产品28"];
  sample.month.products = productNames.map((name) => ({ name }));
  sample.current.products = [
    { name: "产品1", sales: 23432.05, tacos: 0.0809 },
    { name: "产品4", sales: 19417.94, tacos: 0.0505 },
    { name: "产品18", sales: 4021.28, tacos: 0.2164 },
    { name: "产品26", sales: 3591.98, tacos: 0.217 },
    { name: "产品25", sales: 1076.12, tacos: 0.4438 },
  ];
  state.fileName = "绩效复盘表格优化-白板6.15.xlsx";
  applyAnalysisResult(sample);
  setStatus("已载入基于你提供表格的示例状态。正式使用时请上传绩效复盘表。", "success");
}

function sampleAggregate(label, index, dayCount, values) {
  const result = { label, index, dayCount, products: [], ...values };
  result.aov = result.units ? result.sales / result.units : 0;
  result.margin = result.sales ? result.profit / result.sales : 0;
  result.tacos = result.sales ? result.adFee / result.sales : 0;
  result.cvr = result.sessions ? result.units / result.sessions : 0;
  result.ctr = result.impressions ? result.clicks / result.impressions : 0;
  result.adCvr = result.clicks ? result.adOrders / result.clicks : 0;
  result.acos = result.adSales ? result.adSpend / result.adSales : 0;
  result.salesDaily = result.sales / dayCount;
  result.unitsDaily = result.units / dayCount;
  result.sessionsDaily = result.sessions / dayCount;
  result.impressionsDaily = result.impressions / dayCount;
  return result;
}
