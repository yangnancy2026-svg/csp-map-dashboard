const WAREHOUSE = {
  name: "墨尔本火车站",
  address: "Flinders Street Station, Melbourne VIC 3000",
  lat: -37.8183,
  lng: 144.9671,
};

const ROUTE_COLORS = {
  F203: "#1683e5",
  Z705: "#13b965",
  F204: "#ff9218",
  M266: "#984de8",
  M271: "#10aebb",
  L805: "#ff5067",
};
const FALLBACK_COLORS = ["#0b74d1", "#18b56b", "#f28b18", "#8f4de8", "#11a3a3", "#ff4d62", "#6472d9", "#c05a9f"];

const FIELD_ALIASES = {
  csp: ["派件网点简码", "派件网点", "预派件网点编码"],
  route: ["路由码"],
  driver: ["派件司机"],
  postcode: ["邮编"],
  address: ["收件人详细地址", "收件街道", "详细地址"],
  coords: ["地址经纬度", "经纬度", "latlng", "coordinates"],
  deliveryTime: ["派件时间"],
  signedTime: ["签收时间"],
  attempts: ["派件次数"],
  status: ["配送状态"],
  issueType: ["最新问题件类型"],
  issueReason: ["最新问题件原因"],
  weight: ["商家上传重量", "收件总重量"],
  size: ["大小"],
  pieces: ["件数", "件数(包裹)"],
};

let rawRows = [];
let parsedStops = [];
let excludedRows = [];
let map;
let warehouseMarker;
const routeLayer = L.layerGroup();
const lineLayer = L.layerGroup();
const markerLayer = L.layerGroup();

const els = {
  fileInput: document.getElementById("fileInput"),
  fileChip: document.getElementById("fileChip"),
  cspFilter: document.getElementById("cspFilter"),
  routeFilter: document.getElementById("routeFilter"),
  driverFilter: document.getElementById("driverFilter"),
  postcodeFilter: document.getElementById("postcodeFilter"),
  warehouseAddressInput: document.getElementById("warehouseAddressInput"),
  warehouseApplyBtn: document.getElementById("warehouseApplyBtn"),
  warehouseStatus: document.getElementById("warehouseStatus"),
  coverageToggle: document.getElementById("coverageToggle"),
  warehouseToggle: document.getElementById("warehouseToggle"),
  lineToggle: document.getElementById("lineToggle"),
  issueToggle: document.getElementById("issueToggle"),
  refreshBtn: document.getElementById("refreshBtn"),
  viewTitle: document.getElementById("viewTitle"),
  viewSubtitle: document.getElementById("viewSubtitle"),
  mapLegend: document.getElementById("mapLegend"),
  routeChecks: document.getElementById("routeChecks"),
  routeTable: document.getElementById("routeTable"),
  exportRouteBtn: document.getElementById("exportRouteBtn"),
  metricTitle: document.getElementById("metricTitle"),
  kpiGrid: document.getElementById("kpiGrid"),
  firstDistance: document.getElementById("firstDistance"),
  firstAddress: document.getElementById("firstAddress"),
  difficultyScore: document.getElementById("difficultyScore"),
  difficultyLabel: document.getElementById("difficultyLabel"),
  difficultyBadge: document.getElementById("difficultyBadge"),
  difficultyBar: document.getElementById("difficultyBar"),
  difficultyReason: document.getElementById("difficultyReason"),
};

init();

async function init() {
  initMap();
  bindEvents();
  await loadDefaultData();
}

function initMap() {
  map = L.map("map", { zoomControl: false }).setView([WAREHOUSE.lat, WAREHOUSE.lng], 11);
  L.control.zoom({ position: "topleft" }).addTo(map);
  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: "&copy; OpenStreetMap contributors",
  }).addTo(map);

  warehouseMarker = L.marker([WAREHOUSE.lat, WAREHOUSE.lng], {
    icon: L.divIcon({
      className: "warehouse-marker",
      html: `<div style="width:24px;height:24px;border-radius:50%;background:#e3343f;border:3px solid #fff;box-shadow:0 4px 14px rgba(0,0,0,.28)"></div>`,
      iconSize: [24, 24],
      iconAnchor: [12, 12],
    }),
  }).bindPopup(`<strong>站点取件位置</strong><br>${WAREHOUSE.name}<br>${WAREHOUSE.address}`);

  routeLayer.addTo(map);
  lineLayer.addTo(map);
  markerLayer.addTo(map);
  warehouseMarker.addTo(map);
}

function bindEvents() {
  els.fileInput.addEventListener("change", handleFile);
  els.refreshBtn.addEventListener("click", loadDefaultData);
  els.exportRouteBtn.addEventListener("click", exportRouteOverview);
  els.warehouseApplyBtn.addEventListener("click", applyWarehouseAddress);
  els.warehouseAddressInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") applyWarehouseAddress();
  });
  [els.cspFilter, els.routeFilter, els.driverFilter, els.postcodeFilter].forEach((el) => el.addEventListener("change", render));
  [els.coverageToggle, els.warehouseToggle, els.lineToggle, els.issueToggle].forEach((el) => el.addEventListener("change", render));
}

async function loadDefaultData() {
  rawRows = [];
  parsedStops = parseRows(rawRows);
  hydrateFilters();
  updateFileChip("未上传数据", "请上传 Excel / CSV 底表");
  render();
}

async function handleFile(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: "" });
  parsedStops = parseRows(rawRows);
  hydrateFilters();
  updateFileChip(file.name, `${rawRows.length} 行`);
  render();
}

function updateFileChip(name, meta) {
  els.fileChip.innerHTML = `<span>✓</span><div><strong>${escapeHtml(name)}</strong><small>${escapeHtml(meta)}</small></div>`;
}

async function applyWarehouseAddress() {
  const value = clean(els.warehouseAddressInput.value);
  if (!value) {
    setWarehouseStatus("请输入站点取件地址或经纬度。", "error");
    return;
  }

  els.warehouseApplyBtn.disabled = true;
  setWarehouseStatus("正在定位站点地址...", "loading");

  try {
    const directCoord = parseCoord(value);
    const location = directCoord
      ? { ...directCoord, displayName: value }
      : await geocodeWarehouseAddress(value);

    if (!location || !isInAustralia(location.lat, location.lng)) {
      setWarehouseStatus("定位失败或地址不在澳洲范围内。", "error");
      return;
    }

    WAREHOUSE.name = value;
    WAREHOUSE.address = location.displayName || value;
    WAREHOUSE.lat = location.lat;
    WAREHOUSE.lng = location.lng;
    refreshDistances();
    updateWarehouseMarker();
    setWarehouseStatus(`已更新：${WAREHOUSE.lat.toFixed(5)}, ${WAREHOUSE.lng.toFixed(5)}`, "success");
    render();
  } catch (error) {
    setWarehouseStatus("定位失败，请尝试更完整地址或直接输入经纬度。", "error");
  } finally {
    els.warehouseApplyBtn.disabled = false;
  }
}

async function geocodeWarehouseAddress(address) {
  const query = /australia|vic|nsw|qld|wa|sa|tas|act|nt/i.test(address) ? address : `${address}, Australia`;
  const url = `https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=au&q=${encodeURIComponent(query)}`;
  const response = await fetch(url);
  if (!response.ok) throw new Error("Geocoding request failed");
  const result = await response.json();
  if (!result.length) return null;
  return {
    lat: Number(result[0].lat),
    lng: Number(result[0].lon),
    displayName: result[0].display_name,
  };
}

function refreshDistances() {
  parsedStops = parsedStops.map((stop) => ({
    ...stop,
    distance: haversineKm(WAREHOUSE.lat, WAREHOUSE.lng, stop.lat, stop.lng),
  }));
}

function updateWarehouseMarker() {
  warehouseMarker.setLatLng([WAREHOUSE.lat, WAREHOUSE.lng]);
  warehouseMarker.bindPopup(`<strong>站点取件位置</strong><br>${escapeHtml(WAREHOUSE.name)}<br>${escapeHtml(WAREHOUSE.address)}`);
}

function setWarehouseStatus(message, type = "info") {
  els.warehouseStatus.textContent = message;
  const colors = {
    info: "#64748b",
    loading: "#0b3b7a",
    success: "#178f5a",
    error: "#d13d4b",
  };
  els.warehouseStatus.style.color = colors[type] || colors.info;
}

function parseRows(rows) {
  excludedRows = [];
  const stops = [];
  rows.forEach((row, index) => {
      const coord = getField(row, "coords");
      const parsed = parseCoord(coord);
      if (!parsed) {
        excludedRows.push({ index, reason: "经纬度为空或格式错误" });
        return;
      }
      if (!isInAustralia(parsed.lat, parsed.lng)) {
        excludedRows.push({ index, reason: "经纬度不在澳洲范围内", coord });
        return;
      }

      const signedAt = parseDate(getField(row, "signedTime"));
      const deliveredAt = parseDate(getField(row, "deliveryTime"));
      const issueType = clean(getField(row, "issueType"));
      const issueReason = clean(getField(row, "issueReason"));
      const address = clean(getField(row, "address")) || "未填地址";
      const distance = haversineKm(WAREHOUSE.lat, WAREHOUSE.lng, parsed.lat, parsed.lng);

      stops.push({
        id: row["运单号"] || `row-${index + 1}`,
        csp: clean(getField(row, "csp")) || "未识别",
        route: clean(getField(row, "route")) || "未分配",
        driver: clean(getField(row, "driver")) || "未分配",
        postcode: clean(getField(row, "postcode")) || "未知",
        address,
        status: clean(getField(row, "status")) || "未知",
        issue: issueType || issueReason,
        weight: Number(getField(row, "weight")) || 0,
        size: clean(getField(row, "size")),
        pieces: Number(getField(row, "pieces")) || 1,
        attempts: Number(getField(row, "attempts")) || 0,
        signedAt,
        deliveredAt,
        lat: parsed.lat,
        lng: parsed.lng,
        distance,
        isApartment: isApartmentAddress(address),
      });
    });
  return stops;
}

function getField(row, key) {
  for (const field of FIELD_ALIASES[key] || []) {
    if (row[field] !== undefined && row[field] !== null && row[field] !== "") return row[field];
  }
  return "";
}

function hydrateFilters() {
  fillSelect(els.cspFilter, unique(parsedStops.map((d) => d.csp)), "全部网点");
  fillSelect(els.routeFilter, unique(parsedStops.map((d) => d.route)), "全部路由");
  fillSelect(els.driverFilter, unique(parsedStops.map((d) => d.driver)), "全部司机");
  fillSelect(els.postcodeFilter, unique(parsedStops.map((d) => d.postcode)), "全部邮编");
  renderRouteChecks();
  renderLegend();
}

function fillSelect(select, values, allLabel) {
  select.innerHTML = "";
  select.append(new Option(allLabel, "__all__"));
  values.forEach((value) => select.append(new Option(value, value)));
}

function renderRouteChecks() {
  const grouped = groupBy(parsedStops, (d) => d.route);
  els.routeChecks.innerHTML = Object.entries(grouped)
    .sort((a, b) => b[1].length - a[1].length)
    .map(([route, items]) => {
      const color = routeColor(route);
      return `<div class="route-check"><span class="route-swatch" style="background:${color}"></span><span>${route}</span><small>${items.length}</small></div>`;
    })
    .join("");
}

function renderLegend() {
  const routes = unique(parsedStops.map((d) => d.route));
  const routeItems = routes.map((route) => `<span class="legend-item"><i class="legend-dot" style="background:${routeColor(route)}"></i>${route}</span>`);
  els.mapLegend.innerHTML = [
    ...routeItems,
    `<span class="legend-item"><i class="legend-dot" style="background:#e3343f"></i>站点取件位置</span>`,
    `<span class="legend-item"><i class="legend-dot" style="background:#ffd43b;border:1px solid #4b5563"></i>首单位置</span>`,
    `<span class="legend-item"><i class="legend-symbol">◆</i>问题件</span>`,
  ].join("");
}

function render() {
  const stops = filteredStops();
  markerLayer.clearLayers();
  routeLayer.clearLayers();
  lineLayer.clearLayers();

  if (els.warehouseToggle.checked) {
    if (!map.hasLayer(warehouseMarker)) warehouseMarker.addTo(map);
  } else if (map.hasLayer(warehouseMarker)) {
    map.removeLayer(warehouseMarker);
  }

  if (els.coverageToggle.checked) drawCoverage(stops);
  if (els.lineToggle.checked) drawFirstStopLine(stops);
  drawStops(stops);
  renderMetrics(stops);
  renderRouteTable(stops);

  const cspTitle = els.cspFilter.value === "__all__" ? "全部站点" : els.cspFilter.value;
  els.viewTitle.textContent = `站点派送范围分析`;
  els.viewSubtitle.textContent = `数据更新时间：${new Date().toLocaleString("zh-CN", { hour12: false })}`;
  els.metricTitle.textContent = `核心指标（${cspTitle}）`;

  const bounds = L.latLngBounds([[WAREHOUSE.lat, WAREHOUSE.lng], ...stops.map((d) => [d.lat, d.lng])]);
  if (bounds.isValid()) map.fitBounds(bounds.pad(0.12), { maxZoom: 12 });
}

function filteredStops() {
  return parsedStops.filter((stop) => {
    return matches(els.cspFilter, stop.csp) && matches(els.routeFilter, stop.route) && matches(els.driverFilter, stop.driver) && matches(els.postcodeFilter, stop.postcode);
  });
}

function matches(select, value) {
  return select.value === "__all__" || select.value === value;
}

function drawStops(stops) {
  const first = firstStop(stops);
  const radiusBase = pointRadius(stops.length);
  stops.forEach((stop) => {
    const isFirst = first && stop === first;
    const hasIssue = Boolean(stop.issue);
    const color = isFirst ? "#ffd43b" : routeColor(stop.route);
    const radius = isFirst ? Math.max(5, radiusBase + 2) : Math.min(radiusBase + 2, Math.max(radiusBase, radiusBase + Math.sqrt(stop.pieces || 1) * 0.18));
    if (hasIssue && els.issueToggle.checked && !isFirst) {
      drawIssueMarker(stop, color, radius);
      return;
    }
    L.circleMarker([stop.lat, stop.lng], {
      radius,
      color: "#ffffff",
      weight: stops.length > 5000 ? 0.4 : 1.2,
      fillColor: color,
      fillOpacity: stops.length > 5000 ? 0.62 : 0.88,
    })
      .bindPopup(
        `<strong>${escapeHtml(stop.route)} | ${escapeHtml(stop.postcode)}</strong><br>${escapeHtml(stop.address)}<br>司机：${escapeHtml(stop.driver)}<br>距站点：${formatKm(stop.distance)}<br>公寓/Unit：${stop.isApartment ? "是" : "否"}<br>状态：${escapeHtml(stop.status)}`,
      )
      .addTo(markerLayer);
  });
}

function drawIssueMarker(stop, color, radius) {
  const size = Math.max(10, radius * 2.7);
  const marker = L.marker([stop.lat, stop.lng], {
    icon: L.divIcon({
      className: "issue-marker",
      html: `<span style="--route-color:${color};width:${size}px;height:${size}px;">◆</span>`,
      iconSize: [size, size],
      iconAnchor: [size / 2, size / 2],
    }),
  });
  marker
    .bindPopup(
      `<strong>${escapeHtml(stop.route)} | ${escapeHtml(stop.postcode)}</strong><br>${escapeHtml(stop.address)}<br>司机：${escapeHtml(stop.driver)}<br>距站点：${formatKm(stop.distance)}<br>问题件：${escapeHtml(stop.issue)}<br>状态：${escapeHtml(stop.status)}`,
    )
    .addTo(markerLayer);
}

function drawCoverage(stops) {
  const grouped = groupBy(stops, (d) => d.route);
  Object.entries(grouped).forEach(([route, group]) => {
    const hull = convexHull(group.map((d) => [d.lng, d.lat]));
    if (hull.length < 3) return;
    L.polygon(hull.map(([lng, lat]) => [lat, lng]), {
      color: routeColor(route),
      weight: 2,
      opacity: 0.85,
      fillColor: routeColor(route),
      fillOpacity: 0.14,
    }).addTo(routeLayer);
  });
}

function drawFirstStopLine(stops) {
  const first = firstStop(stops);
  if (!first) return;
  L.polyline(
    [
      [WAREHOUSE.lat, WAREHOUSE.lng],
      [first.lat, first.lng],
    ],
    { color: "#334155", weight: 2, opacity: 0.75, dashArray: "8 7" },
  )
    .bindPopup(`站点到首单：${formatKm(first.distance)}`)
    .addTo(lineLayer);
}

function renderMetrics(stops) {
  const signed = stops.filter((d) => d.status.includes("成功") || d.signedAt);
  const issueRate = rate(stops.filter((d) => d.issue).length, stops.length);
  const repeatRate = rate(stops.filter((d) => d.attempts >= 2).length, stops.length);
  const apartmentRate = rate(stops.filter((d) => d.isApartment).length, stops.length);
  const first = firstStop(stops);
  const avgDistance = average(stops.map((d) => d.distance));
  const maxDistance = Math.max(0, ...stops.map((d) => d.distance));
  const completion = rate(signed.length, stops.length);
  const avgSignHours = average(
    stops
      .filter((d) => d.signedAt && d.deliveredAt)
      .map((d) => Math.max(0, (d.signedAt.getTime() - d.deliveredAt.getTime()) / 3600000)),
  );

  const stats = [
    ["总运单数", stops.length],
    ["剔除异常坐标", excludedRows.length],
    ["路由数量", unique(stops.map((d) => d.route)).length],
    ["派件司机", unique(stops.map((d) => d.driver)).length],
    ["覆盖邮编", unique(stops.map((d) => d.postcode)).length],
    ["平均站点距离", formatKm(avgDistance)],
    ["最远派送距离", formatKm(maxDistance)],
    ["签收率", `${completion.toFixed(1)}%`],
    ["公寓地址占比", `${apartmentRate.toFixed(1)}%`],
    ["问题件比例", `${issueRate.toFixed(1)}%`],
    ["二派及以上", `${repeatRate.toFixed(1)}%`],
    ["平均签收时长", `${avgSignHours.toFixed(1)}h`],
    ["大件重货占比", `${(heavyShare(stops) * 100).toFixed(1)}%`],
  ];
  els.kpiGrid.innerHTML = stats.map(([label, value]) => `<div class="kpi"><strong>${value}</strong><span>${label}</span></div>`).join("");

  els.firstDistance.textContent = first ? first.distance.toFixed(2) : "-";
  els.firstAddress.textContent = first ? `首单地址：${first.address}` : "首单地址：-";

  const difficulty = scoreDifficulty(stops, {
    firstDistance: first?.distance || 0,
    avgDistance,
    maxDistance,
    issueRate: issueRate / 100,
    repeatRate: repeatRate / 100,
    apartmentRate: apartmentRate / 100,
  });
  els.difficultyScore.textContent = difficulty.score;
  els.difficultyLabel.textContent = difficulty.label;
  els.difficultyBadge.textContent = difficulty.label;
  els.difficultyBadge.style.color = difficulty.color;
  els.difficultyBadge.style.background = difficulty.badgeBg;
  els.difficultyBar.style.left = `${difficulty.score}%`;
  els.difficultyReason.textContent = difficulty.reason;
}

function renderRouteTable(stops) {
  els.routeTable.innerHTML = buildRouteOverviewRows(stops)
    .map((row) => {
      return `<tr>
        <td><span class="route-name"><i class="route-swatch" style="background:${routeColor(row.route)}"></i>${escapeHtml(row.route)}</span></td>
        <td>${row.parcels}</td>
        <td>${row.postcodes}</td>
        <td>${row.drivers}</td>
        <td>${row.avgDistance.toFixed(2)}</td>
        <td>${row.maxDistance.toFixed(2)}</td>
        <td>${row.firstDistance == null ? "-" : row.firstDistance.toFixed(2)}</td>
        <td>${row.apartmentRate.toFixed(1)}%</td>
        <td>${row.signedRate.toFixed(1)}%</td>
        <td>${row.difficultyScore}</td>
        <td><span class="difficulty-pill" style="color:${row.difficultyColor};background:${row.difficultyBg}">${row.difficultyLabel}</span></td>
      </tr>`;
    })
    .join("");
}

function buildRouteOverviewRows(stops) {
  const grouped = groupBy(stops, (d) => d.route);
  return Object.entries(grouped)
    .sort((a, b) => b[1].length - a[1].length)
    .map(([route, items]) => {
      const signedRate = rate(items.filter((d) => d.status.includes("成功") || d.signedAt).length, items.length);
      const aptRate = rate(items.filter((d) => d.isApartment).length, items.length);
      const issueRate = rate(items.filter((d) => d.issue).length, items.length);
      const repeatRate = rate(items.filter((d) => d.attempts >= 2).length, items.length);
      const routeFirst = firstStop(items);
      const avgDistance = average(items.map((d) => d.distance));
      const maxDistance = Math.max(0, ...items.map((d) => d.distance));
      const difficulty = scoreDifficulty(items, {
        firstDistance: routeFirst?.distance || 0,
        avgDistance,
        maxDistance,
        issueRate: issueRate / 100,
        repeatRate: repeatRate / 100,
        apartmentRate: aptRate / 100,
      });
      return {
        route,
        parcels: items.length,
        postcodes: unique(items.map((d) => d.postcode)).length,
        drivers: unique(items.map((d) => d.driver)).length,
        avgDistance,
        maxDistance,
        firstDistance: routeFirst ? routeFirst.distance : null,
        apartmentRate: aptRate,
        issueRate,
        repeatRate,
        signedRate,
        difficultyScore: difficulty.score,
        difficultyLabel: difficulty.label,
        difficultyColor: difficulty.color,
        difficultyBg: difficulty.badgeBg,
      };
    });
}

function exportRouteOverview() {
  if (!window.XLSX) {
    alert("Excel 导出组件未加载，请刷新页面后重试。");
    return;
  }
  const rows = buildRouteOverviewRows(filteredStops()).map((row) => ({
    站点: els.cspFilter.value === "__all__" ? "全部站点" : els.cspFilter.value,
    取件地址: WAREHOUSE.name,
    路由码: row.route,
    运单数: row.parcels,
    覆盖邮编: row.postcodes,
    司机数: row.drivers,
    "平均距站点(km)": Number(row.avgDistance.toFixed(2)),
    "最远距站点(km)": Number(row.maxDistance.toFixed(2)),
    "首单距离(km)": row.firstDistance == null ? "" : Number(row.firstDistance.toFixed(2)),
    公寓地址占比: `${row.apartmentRate.toFixed(1)}%`,
    问题件比例: `${row.issueRate.toFixed(1)}%`,
    二派及以上比例: `${row.repeatRate.toFixed(1)}%`,
    签收率: `${row.signedRate.toFixed(1)}%`,
    难度评分: row.difficultyScore,
    难度等级: row.difficultyLabel,
  }));
  if (!rows.length) {
    alert("当前筛选没有可导出的路由数据。");
    return;
  }
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "路由概览");
  const date = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(workbook, `路由概览_${date}.xlsx`);
}

function scoreDifficulty(stops, values) {
  if (!stops.length) {
    return { score: 0, label: "待计算", color: "#64748b", badgeBg: "#eef2f7", reason: "当前筛选没有有效经纬度数据。" };
  }
  const spread = routeSpreadKm(stops);
  let score = 14;
  score += Math.min(18, values.firstDistance * 1.15);
  score += Math.min(20, values.avgDistance * 1.0);
  score += Math.min(16, spread * 1.55);
  score += Math.min(12, values.issueRate * 100);
  score += Math.min(10, values.repeatRate * 55);
  score += Math.min(14, values.apartmentRate * 60);
  score += Math.min(8, heavyShare(stops) * 28);
  score = Math.round(Math.min(100, score));

  const label = score >= 75 ? "极难" : score >= 60 ? "困难" : score >= 40 ? "中等" : "简单";
  const color = score >= 75 ? "#ff4d62" : score >= 60 ? "#ff6f3c" : score >= 40 ? "#d98b10" : "#18b56b";
  const badgeBg = score >= 60 ? "#fff0f1" : score >= 40 ? "#fff7e6" : "#edf9f1";
  const reason = `首单 ${formatKm(values.firstDistance)}，平均距站点 ${formatKm(values.avgDistance)}，离散度 ${formatKm(spread)}，公寓/Unit 地址占比 ${(values.apartmentRate * 100).toFixed(1)}%。`;
  return { score, label, color, badgeBg, reason };
}

function pointRadius(count) {
  if (count >= 20000) return 1.2;
  if (count >= 10000) return 1.6;
  if (count >= 5000) return 2.1;
  if (count >= 2000) return 2.8;
  if (count >= 800) return 3.4;
  if (count >= 300) return 4.2;
  return 5.2;
}

function isApartmentAddress(address) {
  const text = String(address || "").trim().toLowerCase();
  if (!text) return false;
  const keywordHit = /\b(unit|u|apt|apartment|flat|suite|ste|level|lvl|room|rm|lot|shop|floor|flr|studio)\b/.test(text);
  const hashHit = /#\s*[a-z0-9-]+/.test(text);
  const slashHit = /(^|\s)[a-z]?\d+[a-z]?\s*[\/\\]\s*\d+[a-z]?\b/.test(text);
  const rangeWithUnit = /\b\d+[a-z]?\s*-\s*\d+[a-z]?\s+[a-z]/.test(text) && /\b(unit|apt|flat|suite|level|shop)\b/.test(text);
  return keywordHit || hashHit || slashHit || rangeWithUnit;
}

function firstStop(stops) {
  return [...stops].sort((a, b) => {
    const ta = (a.signedAt || a.deliveredAt || new Date(8640000000000000)).getTime();
    const tb = (b.signedAt || b.deliveredAt || new Date(8640000000000000)).getTime();
    return ta - tb;
  })[0];
}

function routeColor(route) {
  if (ROUTE_COLORS[route]) return ROUTE_COLORS[route];
  const routes = unique(parsedStops.map((d) => d.route));
  const index = Math.max(0, routes.indexOf(route));
  return FALLBACK_COLORS[index % FALLBACK_COLORS.length];
}

function parseCoord(value) {
  if (!value) return null;
  const parts = String(value).match(/-?\d+(?:\.\d+)?/g);
  if (!parts || parts.length < 2) return null;
  let a = Number(parts[0]);
  let b = Number(parts[1]);
  if (!Number.isFinite(a) || !Number.isFinite(b)) return null;
  if (Math.abs(a) > 90 && Math.abs(b) <= 90) [a, b] = [b, a];
  if (Math.abs(a) > 90 || Math.abs(b) > 180) return null;
  return { lat: a, lng: b };
}

function isInAustralia(lat, lng) {
  return lat >= -44.5 && lat <= -9.0 && lng >= 112.0 && lng <= 154.5;
}

function parseDate(value) {
  if (!value) return null;
  const date = new Date(String(value).replace(/-/g, "/"));
  return Number.isNaN(date.getTime()) ? null : date;
}

function clean(value) {
  return String(value ?? "").trim();
}

function unique(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), "zh-CN"));
}

function groupBy(items, getKey) {
  return items.reduce((acc, item) => {
    const key = getKey(item);
    acc[key] ||= [];
    acc[key].push(item);
    return acc;
  }, {});
}

function average(values) {
  const nums = values.filter(Number.isFinite);
  return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
}

function rate(part, total) {
  return total ? (part / total) * 100 : 0;
}

function heavyShare(stops) {
  return stops.filter((d) => d.weight >= 5 || /BOX|BULK|LARGE/i.test(d.size)).length / Math.max(1, stops.length);
}

function routeSpreadKm(stops) {
  if (stops.length < 2) return 0;
  const center = {
    lat: average(stops.map((d) => d.lat)),
    lng: average(stops.map((d) => d.lng)),
  };
  return average(stops.map((d) => haversineKm(center.lat, center.lng, d.lat, d.lng)));
}

function formatKm(value) {
  return `${value.toFixed(value >= 10 ? 1 : 2)}km`;
}

function haversineKm(lat1, lng1, lat2, lng2) {
  const rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad;
  const dLng = (lng2 - lng1) * rad;
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLng / 2) ** 2;
  return 6371 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function convexHull(points) {
  const sorted = [...new Map(points.map((p) => [p.join(","), p])).values()].sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  if (sorted.length <= 1) return sorted;
  const cross = (o, a, b) => (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0]);
  const lower = [];
  for (const p of sorted) {
    while (lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], p) <= 0) lower.pop();
    lower.push(p);
  }
  const upper = [];
  for (let i = sorted.length - 1; i >= 0; i -= 1) {
    const p = sorted[i];
    while (upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], p) <= 0) upper.pop();
    upper.push(p);
  }
  upper.pop();
  lower.pop();
  return lower.concat(upper);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
