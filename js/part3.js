(() => {
// 全局状态与常量

let RAW_DATA = null;
let selectedCountries = [];
let selectedPrimary    = null;
let selectedSecondary  = null;
let undoStack = [], redoStack = [];

const HISTORY_LIMIT   = 50;
const AGGREGATE_NAME  = "整体";
const DEFAULT_PERIOD  = "2016-2020";
const SECONDARY_LIMIT = 12;

// 路径
const DATA_FILES = {
    subjectMeta:    "./data/领域.xlsx",
    aggregateFields:"./data/2-中东欧-时间段-领域.xlsx",
    countryFields:  "./data/2-中东欧前12国-时间段（仅20162020）-领域.xlsx",
    countryList:    "./data/中东欧16国.xlsx"    

};

// 配色（使用共享颜色文件 window.SHARED_*，若未加载则保留默认备选值）

let primaryPalette = null;   // 一级学科名 -> 颜色，init() 后固定不变
let countryPalette = null;   // 国家名     -> 颜色，init() 后固定不变
const part3Root = document.getElementById("part3");

const tooltip3      = d3.select("#tooltip3");
const formatNumber  = d3.format(",");
const formatPercent = d3.format(".1f");

// 初始化
document.addEventListener("DOMContentLoaded", init);

async function init() {
    showLoading("正在读取 Excel 数据...");
    try {
        RAW_DATA = await loadRawDataFromExcel();

        selectedCountries = [];
        selectedPrimary   = null;
        selectedSecondary = null;
        undoStack = []; redoStack = [];

        primaryPalette = d3.scaleOrdinal()
            .domain(RAW_DATA.primaryList.map(d => d.name))
            .range(window.SHARED_PRIMARY_COLORS || []);
        countryPalette = d3.scaleOrdinal()
            .domain(RAW_DATA.countries)
            .range(window.SHARED_COUNTRY_COLORS || []);

        renderSelectionOptions();
        setupEventListeners();
        updateAllCharts();
        showLoading("");
        window.addEventListener("resize", debounce(updateAllCharts, 150));
    } catch (error) {
        console.error(error);
        showLoading(
            `数据加载失败：${error.message}<br>
            请确认 Excel 文件已放入 <strong>data/</strong> 文件夹，且通过本地服务器打开页面，例如 <code>python -m http.server</code>。`,
            true
        );
    }
}

function showLoading(message, isError = false) {
    const el = part3Root ? part3Root.querySelector("#loading") : document.getElementById("loading");
    if (!el) return;
    if (!message) { el.style.display = "none"; return; }
    el.innerHTML = message;
    el.className = isError ? "loading error" : "loading";
    el.style.display = "block";
}

function debounce(fn, delay) {
    let t;
    return function(...args) { clearTimeout(t); t = setTimeout(() => fn.apply(this, args), delay); };
}

// excel读取
async function loadRawDataFromExcel() {
    if (typeof XLSX === "undefined")
        throw new Error("未检测到 SheetJS / XLSX，请检查 index.html 是否已引入 xlsx.full.min.js。");

    const subjectRows   = await loadFirstSheetRows(DATA_FILES.subjectMeta,    true);
    const aggregateRows = await loadFirstSheetRows(DATA_FILES.aggregateFields, true);
    const countryRows   = await loadFirstSheetRows(DATA_FILES.countryFields,   true);
    const countryListRows = await loadFirstSheetRows(DATA_FILES.countryList,  false);

    const subjectMeta = buildSubjectMeta(subjectRows);
    if (!subjectMeta.length)
        throw new Error("领域.xlsx 中未读取到有效的一级/二级学科映射。");

    const metaBySecondary = new Map(subjectMeta.map(d => [d.secondaryId, d]));
    const metaByPrimary   = new Map();
    subjectMeta.forEach(d => {
        if (!metaByPrimary.has(d.primaryId))
            metaByPrimary.set(d.primaryId, { primaryId: d.primaryId, primaryName: d.primaryName });
    });

    const records = [];
    aggregateRows.forEach(row => {
        const rec = normalizeRecord(row, AGGREGATE_NAME, metaBySecondary, metaByPrimary);
        if (rec) records.push(rec);
    });
    countryRows.forEach(row => {
        const country = cleanText(valueOf(row, ["国家中文名", "国家", "合作国家"]));
        const rec = normalizeRecord(row, country, metaBySecondary, metaByPrimary);
        if (rec) records.push(rec);
    });

    if (!records.length)
        throw new Error("未从领域数据表中读取到有效论文数量记录。");

    const periods = Array.from(new Set(records.map(d => d.period))).sort();
    const defaultPeriod = periods.includes(DEFAULT_PERIOD)
        ? DEFAULT_PERIOD
        : (periods[periods.length - 1] || DEFAULT_PERIOD);

    // 一级学科列表
    const aggDefault = records.filter(d => d.entity === AGGREGATE_NAME && d.period === defaultPeriod);
    const primaryTotals = d3.rollup(aggDefault, v => d3.sum(v, d => d.count), d => d.primaryId);
    const primaryList = Array.from(new Set(records.map(d => d.primaryId)))
        .map(id => {
            const meta = metaByPrimary.get(id);
            return { id, name: meta ? meta.primaryName : id, total: primaryTotals.get(id) || 0 };
        })
        .sort((a, b) => d3.descending(a.total, b.total) || d3.ascending(a.id, b.id));

    // 国家列表
    const countryTotals = d3.rollup(
        records.filter(d => d.entity !== AGGREGATE_NAME && d.period === defaultPeriod),
        v => d3.sum(v, d => d.count), d => d.entity
    );
    let countries = Array.from(countryTotals, ([c, t]) => ({ country: c, total: t }))
        .filter(d => d.country)
        .sort((a, b) => d3.descending(a.total, b.total))
        .map(d => d.country);

    if (countryListRows.length) {
        const ordered = countryListRows
            .map(row => cleanText(valueOf(row, ["国家中文名", "国家"])))
            .filter(Boolean);
        countries = countries.sort((a, b) => {
            const ia = ordered.indexOf(a), ib = ordered.indexOf(b);
            if (ia === -1 && ib === -1) return 0;
            return (ia === -1) ? 1 : (ib === -1) ? -1 : ia - ib;
        });
    }

    return { aggregateName: AGGREGATE_NAME, periods, defaultPeriod, countries, primaryList, subjectMeta, records };
}

// 依次尝试候选路径，返回第一个成功的 Sheet 的行数组
async function loadFirstSheetRows(candidatePath, required = true) {
    const errors = [];
    try {
        const res = await fetch(encodeURI(candidatePath), { cache: "no-store" });
        if (!res.ok) {
            errors.push(`${candidatePath} (${res.status})`);
        } else {
            const wb = XLSX.read(await res.arrayBuffer(), { type: "array" });
            const ws = wb.Sheets[wb.SheetNames[0]];
            // raw: false 将所有值转为字符串，避免数字 ID 被读成浮点数
            return XLSX.utils.sheet_to_json(ws, { defval: "", raw: false });
        }
    } catch (e) {
        errors.push(`${candidatePath} (${e.message})`);
    }
    if (required)
        throw new Error(`无法读取 Excel 文件：${candidatePath}。尝试结果：${errors.join("；")}`);
    return [];
}

// 将 领域.xlsx 行解析为学科元数据对象（仅保留图表实际使用的字段）
function buildSubjectMeta(rows) {
    const result = [];
    rows.forEach(row => {
        const primaryId   = normalizeId(valueOf(row, ["一级ID","一级id","一级学科ID"]), 2);
        const secondaryId = normalizeId(valueOf(row, ["二级ID","二级id","二级学科ID"]), 4);
        if (!primaryId || !secondaryId) return;
        result.push({
            primaryId,
            primaryName:   cleanText(valueOf(row, ["一级中文名称","一级学科","一级名称"])) || primaryId,
            secondaryId,
            secondaryName: cleanText(valueOf(row, ["二级中文名称","二级学科","二级名称"])) || secondaryId
        });
    });
    return result;
}

// 将一行 Excel 数据标准化为内部 record 对象；字段缺失或论文数 <= 0 时返回 null
function normalizeRecord(row, entity, metaBySecondary, metaByPrimary) {
    entity = cleanText(entity);
    if (!entity) return null;

    const secondaryId = normalizeId(valueOf(row, ["二级ID","二级id","二级学科ID"]), 4);
    let   primaryId   = normalizeId(valueOf(row, ["一级ID","一级id","一级学科ID"]), 2);
    // 容错：若无独立一级 ID 列，从二级 ID 前两位推断
    if (!primaryId && secondaryId) primaryId = secondaryId.slice(0, 2);

    const count = parseCount(valueOf(row, ["合作Wos论文数","合作WoS论文数","论文数","数量","count"]));
    if (!secondaryId || !primaryId || count <= 0) return null;

    const meta = metaBySecondary.get(secondaryId);
    const pmeta = metaByPrimary.get(primaryId);
    return {
        entity,
        period:        cleanText(valueOf(row, ["时间段","period"])) || DEFAULT_PERIOD,
        primaryId,
        primaryName:   meta ? meta.primaryName : (pmeta ? pmeta.primaryName : primaryId),
        secondaryId,
        secondaryName: meta ? meta.secondaryName : secondaryId,
        count
    };
}

// 数据解析
function valueOf(row, names) {
    for (const name of names) {
        if (Object.prototype.hasOwnProperty.call(row, name) &&
            row[name] != null && String(row[name]).trim() !== "")
            return row[name];
    }
    return "";
}

// 去除首尾及多余内部空白
function cleanText(value) {
    return String(value ?? "").replace(/\s+/g, " ").trim();
}

// 将 ID 规范化为固定位数字符串（处理 SheetJS 将整数读为 "1.0" 的问题）
function normalizeId(value, length) {
    const text = cleanText(value).replace(/\.0$/, "");
    if (!text) return "";
    return /^\d+$/.test(text) ? text.padStart(length, "0") : text;
}

// 解析论文数，容忍千分位逗号
function parseCount(value) {
    const n = Number(String(value ?? "").replace(/,/g, "").trim());
    return Number.isFinite(n) ? n : 0;
}

// 历史记录

// 生成当前三个状态的独立浅拷贝快照
function getStateSnapshot() {
    return {
        selectedCountries: [...selectedCountries],
        selectedPrimary:   selectedPrimary   ? { ...selectedPrimary }   : null,
        selectedSecondary: selectedSecondary ? { ...selectedSecondary } : null
    };
}

// 用户操作前调用：将当前状态压入撤销栈（相同状态时跳过）
function pushHistory() {
    const current = getStateSnapshot();
    const last = undoStack[undoStack.length - 1];
    if (last && JSON.stringify(last) === JSON.stringify(current)) return;
    undoStack.push(current);
    if (undoStack.length > HISTORY_LIMIT) undoStack.shift();
    redoStack = [];
    updateHistoryButtons();
}

// 将三个状态恢复为给定快照并重新渲染
function restoreState(state) {
    selectedCountries = [...state.selectedCountries];
    selectedPrimary   = state.selectedPrimary   ? { ...state.selectedPrimary }   : null;
    // 注意：必须拷贝 state.selectedSecondary 而非当前变量（原 bug 已修复）
    selectedSecondary = state.selectedSecondary ? { ...state.selectedSecondary } : null;
    renderSelectionOptions();
    updateAllCharts();
}

function undoState() {
    if (!undoStack.length) return;
    const cur = getStateSnapshot();
    redoStack.push(cur);
    restoreState(undoStack.pop());
}

function redoState() {
    if (!redoStack.length) return;
    const cur = getStateSnapshot();
    undoStack.push(cur);
    restoreState(redoStack.pop());
}

function updateHistoryButtons() {
    d3.select(part3Root ? part3Root.querySelector("#undoBtn") : null).property("disabled", !undoStack.length);
    d3.select(part3Root ? part3Root.querySelector("#redoBtn") : null).property("disabled", !redoStack.length);
}

// 计算函数

// 安全百分比
function safePercent(value, denominator) {
    if (!denominator || !Number.isFinite(value) || !Number.isFinite(denominator)) return 0;
    const p = value / denominator * 100;
    if (!Number.isFinite(p)) return 0;
    if (p > 100.01) console.warn(`[safePercent] ${formatPercent(p)}%，value=${value}，denom=${denominator}`);
    return Math.max(0, Math.min(100, p));
}

// 筛选指定实体
function getRecords(entity, period = RAW_DATA.defaultPeriod) {
    return RAW_DATA.records.filter(d => d.entity === entity && d.period === period && d.count > 0);
}

// 筛选所有中东欧国家
function getCountryRecords(period = RAW_DATA.defaultPeriod) {
    return RAW_DATA.records.filter(d =>
        d.entity !== RAW_DATA.aggregateName && d.period === period && d.count > 0);
}

function sumRecords(records) {
    return d3.sum(records, d => d.count);
}

// 通过 subjectMeta 查找一级学科中文名；未找到时返回 ID 本身
function getPrimaryLabel(primaryId) {
    const item = RAW_DATA.subjectMeta.find(d => d.primaryId === primaryId);
    return item ? item.primaryName : primaryId;
}

// 将超长标签截断并追加省略号（完整名称通过 SVG <title> 或 tooltip3 展示）
function truncateLabel(text, maxLen) {
    if (!text) return "";
    return text.length > maxLen ? text.slice(0, maxLen) + "…" : text;
}

// tooltip3 

function showtooltip3(event, html) {
    tooltip3.html(html)
        .style("left", `${event.pageX + 14}px`)
        .style("top",  `${event.pageY + 14}px`)
        .style("opacity", 1);
}
function movetooltip3(event) {
    tooltip3.style("left", `${event.pageX + 14}px`).style("top", `${event.pageY + 14}px`);
}
function hidetooltip3() { tooltip3.style("opacity", 0); }

// 事件监听与状态更新
function setupEventListeners() {
    d3.select(part3Root ? part3Root.querySelector("#undoBtn") : null).on("click", undoState);
    d3.select(part3Root ? part3Root.querySelector("#redoBtn") : null).on("click", redoState);

    d3.select(part3Root ? part3Root.querySelector("#resetBtn") : null).on("click", () => {
        pushHistory();
        selectedCountries = []; selectedPrimary = null; selectedSecondary = null;
        renderSelectionOptions(); updateAllCharts();
    });
    d3.select(part3Root ? part3Root.querySelector("#selectAllBtn") : null).on("click", () => {
        pushHistory(); selectedCountries = [...RAW_DATA.countries];
        renderSelectionOptions(); updateAllCharts();
    });
    d3.select(part3Root ? part3Root.querySelector("#clearAllBtn") : null).on("click", () => {
        pushHistory(); selectedCountries = [];
        renderSelectionOptions(); updateAllCharts();
    });
}

// 所有状态变更后的统一重绘入口
function updateAllCharts() {
    d3.select(part3Root ? part3Root.querySelector("#selectionStatus") : null).text(selectedPrimary ? selectedPrimary.name : "全部一级学科");
    drawSankeyChart();
    drawRadarChart();
    drawTrendChart();
    updateHistoryButtons();
}

// 用 D3 data-join 渲染国家复选框
function renderSelectionOptions() {
    const container = d3.select(part3Root ? part3Root.querySelector("#selectionOptions") : null);
    container.selectAll("label").remove();

    const labels = container.selectAll("label")
        .data(RAW_DATA.countries).enter()
        .append("label").attr("class", "compare-pill");

    labels.append("input")
        .attr("type", "checkbox").attr("value", d => d)
        .property("checked", d => selectedCountries.includes(d))
        .on("change", function(event, country) {
            pushHistory();
            if (this.checked) { if (!selectedCountries.includes(country)) selectedCountries.push(country); }
            else { selectedCountries = selectedCountries.filter(d => d !== country); }
            updateAllCharts();
        });
    labels.append("span").text(d => d);
}

// 七、桑基图
function makeSecondaryBuckets(records, primaryId, limit = SECONDARY_LIMIT, forcedSecondaryId = null) {
    const grouped = Array.from(
        d3.rollup(
            records.filter(d => d.primaryId === primaryId),
            v => ({ value: d3.sum(v, d => d.count), ...v[0] }),
            d => d.secondaryId
        ),
        ([secondaryId, info]) => ({
            id: secondaryId, secondaryId, secondaryName: info.secondaryName,
            primaryId: info.primaryId, primaryName: info.primaryName,
            value: info.value, members: [secondaryId], isOther: false,
            label: info.secondaryName
        })
    ).sort((a, b) => d3.descending(a.value, b.value));

    if (grouped.length <= limit) return grouped;

    let kept = grouped.slice(0, limit);

    if (forcedSecondaryId && !kept.some(d => d.secondaryId === forcedSecondaryId)) {
        const forced = grouped.find(d => d.secondaryId === forcedSecondaryId);
        if (forced) { kept = [...kept.slice(0, limit - 1), forced]; }
    }

    const keptIds = new Set(kept.map(d => d.secondaryId));
    const others  = grouped.filter(d => !keptIds.has(d.secondaryId));

    if (others.length) {
        kept.push({
            id: `other:${primaryId}`, secondaryId: `other:${primaryId}`,
            secondaryName: "其他", primaryId, primaryName: getPrimaryLabel(primaryId),
            value: d3.sum(others, d => d.value), members: others.flatMap(d => d.members),
            isOther: true, label: "其他"
        });
    }

    return kept.sort((a, b) => a.isOther ? 1 : b.isOther ? -1 : d3.descending(a.value, b.value));
}


function buildSankeyData() {
    const records = getCountryRecords(RAW_DATA.defaultPeriod);
    const total   = sumRecords(records);
    const nodes = [], links = [];
    const nodeIds = new Set();
    const addNode = n => { if (!nodeIds.has(n.id)) { nodeIds.add(n.id); nodes.push(n); } };

    RAW_DATA.countries.forEach(c =>
        addNode({ id: `country:${c}`, name: c, type: "country", country: c })
    );
    RAW_DATA.primaryList.forEach(p =>
        addNode({ id: `primary:${p.id}`, name: p.name, type: "primary", primaryId: p.id, primaryName: p.name })
    );

    // 国家 -> 一级学科连线
    d3.rollup(records, v => d3.sum(v, d => d.count), d => d.entity, d => d.primaryId)
        .forEach((primaryMap, country) => {
            primaryMap.forEach((value, primaryId) => {
                links.push({
                    source: `country:${country}`, target: `primary:${primaryId}`, value,
                    linkType: "country-primary", country, primaryId,
                    sourceName: country, targetName: getPrimaryLabel(primaryId), total
                });
            });
        });

    // 一级 -> 二级学科连线（仅展开时）
    if (selectedPrimary) {
        const { id: primaryId, name: primaryName } = selectedPrimary;
        const primaryTotal = sumRecords(records.filter(d => d.primaryId === primaryId));
        const buckets = makeSecondaryBuckets(records, primaryId, SECONDARY_LIMIT,
            selectedSecondary ? selectedSecondary.id : null);

        buckets.forEach(bucket => {
            const sid = bucket.isOther ? `secondary:other:${primaryId}` : `secondary:${bucket.secondaryId}`;
            addNode({ id: sid, name: bucket.secondaryName, type: "secondary",
                primaryId, primaryName, secondaryId: bucket.secondaryId,
                secondaryName: bucket.secondaryName, isOther: bucket.isOther });
            links.push({
                source: `primary:${primaryId}`, target: sid, value: bucket.value,
                linkType: "primary-secondary", primaryId, secondaryId: bucket.secondaryId,
                sourceName: primaryName, targetName: bucket.secondaryName, total: primaryTotal
            });
        });
    }

    return { nodes, links, total, mode: selectedPrimary ? "detail" : "overview" };
}

function drawSankeyChart() {
    const container = part3Root ? part3Root.querySelector("#sankeyChart") : document.getElementById("sankeyChart");
    container.innerHTML = "";

    const width  = container.clientWidth  || 880;
    const height = container.clientHeight || 660;

    const countryText = selectedCountries.length ? selectedCountries.join("、") : "整体";
    d3.select(part3Root ? part3Root.querySelector("#sankeySubtitle") : null).text(selectedPrimary
        ? `${countryText} · ${RAW_DATA.defaultPeriod} · ${selectedPrimary.name}二级学科展开`
        : `${countryText} · ${RAW_DATA.defaultPeriod} · 国家—一级学科总览`);

    if (!d3.sankey) {
        container.innerHTML = "<div class='empty-state'>d3-sankey 未加载，请确认网络可访问 CDN。</div>";
        return;
    }

    const margin = { top: 12, right: 16, bottom: 10, left: 16 };
    const cw = width - margin.left - margin.right;
    const ch = height - margin.top - margin.bottom;

    const svg = d3.select(container).append("svg")
        .attr("viewBox", `0 0 ${width} ${height}`)
        .attr("width", width).attr("height", height);
    const g = svg.append("g").attr("transform", `translate(${margin.left},${margin.top})`);

    const sankeyData = buildSankeyData();
    if (!sankeyData.links.length) {
        g.append("text").attr("class", "empty-state-text")
            .attr("x", cw / 2).attr("y", ch / 2).attr("text-anchor", "middle").text("暂无可展示数据");
        return;
    }

    const sankeyGen = d3.sankey()
        .nodeId(d => d.id).nodeWidth(18)
        .nodePadding(selectedPrimary ? 12 : 18)  // 展开时节点增多，减小间距
        .nodeAlign((node) => {
            if (node.type === "country")   return 0;
            if (node.type === "primary")   return 1;
            if (node.type === "secondary") return 2;
            return 1;
        })
        .extent([[0, 5], [cw, ch - 5]]);

    // 深拷贝：d3-sankey 直接修改传入对象，防止污染 sankeyData
    const graph = sankeyGen({
        nodes: sankeyData.nodes.map(d => ({ ...d })),
        links: sankeyData.links.map(d => ({ ...d }))
    });

    // 连线
    const link = g.append("g").attr("class", "sankey-links")
        .selectAll("path").data(graph.links).enter().append("path")
        .attr("class", d => getLinkClass(d))
        .attr("d", d3.sankeyLinkHorizontal())
        .attr("stroke", d => getLinkColor(d))
        .attr("stroke-width", d => Math.max(1, d.width))
        .on("mouseover", (event, d) => {
            showtooltip3(event,
                `<strong>${d.sourceName} → ${d.targetName}</strong><br>
                合作论文数：${formatNumber(d.value)} 篇<br>
                占当前层级总量：${formatPercent(safePercent(d.value, d.total))}%`);
        })
        .on("mousemove", movetooltip3).on("mouseout", hidetooltip3);

    // 节点
    const node = g.append("g").attr("class", "sankey-nodes")
        .selectAll("g").data(graph.nodes).enter().append("g")
        .attr("class", d => getNodeClass(d))
        .attr("transform", d => `translate(${d.x0},${d.y0})`)
        .on("click", (event, d) => handleSankeyNodeClick(event, d))
        .on("mouseover", (event, d) => {
            const refTotal = getNodeReferenceTotal(d, sankeyData);
            showtooltip3(event,
                `<strong>${getNodeTitle(d)}</strong><br>
                合作论文数：${formatNumber(d.value)} 篇<br>
                占比：${formatPercent(safePercent(d.value, refTotal))}%<br>
                <span class="tooltip3-hint">${getNodeHint(d)}</span>`);
        })
        .on("mousemove", movetooltip3).on("mouseout", hidetooltip3)
        .call(d3.drag().subject(d => d).on("drag", function(event, d) {
            const nw = d.x1 - d.x0, nh = d.y1 - d.y0;
            d.x0 = Math.max(0, Math.min(cw - nw, d.x0 + event.dx)); d.x1 = d.x0 + nw;
            d.y0 = Math.max(0, Math.min(ch - nh, d.y0 + event.dy)); d.y1 = d.y0 + nh;
            d3.select(this).attr("transform", `translate(${d.x0},${d.y0})`);
            sankeyGen.update(graph);
            link.attr("d", d3.sankeyLinkHorizontal());
            // 同步更新标签位置与对齐方向（节点跨越中线时需翻转）
            d3.select(this).select("text")
                .attr("x", d.x0 < cw / 2 ? nw + 7 : -7)
                .attr("text-anchor", d.x0 < cw / 2 ? "start" : "end");
        }));

    node.append("rect")
        .attr("height", d => Math.max(3, d.y1 - d.y0)).attr("width", d => d.x1 - d.x0)
        .attr("rx", 2).attr("fill", d => getNodeColor(d));
    node.append("text").attr("class", "sankey-label")
        .attr("x", d => d.x0 < cw / 2 ? (d.x1 - d.x0) + 7 : -7)
        .attr("y", d => (d.y1 - d.y0) / 2).attr("dy", "0.35em")
        .attr("text-anchor", d => d.x0 < cw / 2 ? "start" : "end")
        .text(d => truncateLabel(d.name, d.type === "secondary" ? 9 : 12));

    // 列标题
    svg.append("text").attr("class", "sankey-column-label").attr("x", margin.left).attr("y", 14).text("国家");
    svg.append("text").attr("class", "sankey-column-label")
        .attr("x", selectedPrimary ? width / 2 : width - margin.right)
        .attr("y", 14).attr("text-anchor", selectedPrimary ? "middle" : "end").text("一级学科");
    if (selectedPrimary) {
        svg.append("text").attr("class", "sankey-column-label")
            .attr("x", width - margin.right).attr("y", 14).attr("text-anchor", "end").text("二级学科");
    }
}

/* ─── 桑基图样式辅助 ─────────────────────────────────────────── */

function getLinkClass(d) {
    const cls = ["sankey-link"];
    const sel = new Set(selectedCountries);
    const hasSel = sel.size > 0;
    // 未选中国家的 country-primary 连线，或非展开学科的连线 -> 淡化为背景
    if (hasSel && d.linkType === "country-primary" && !sel.has(d.country))      cls.push("context-link");
    if (selectedPrimary && d.linkType === "country-primary" && d.primaryId !== selectedPrimary.id) cls.push("context-link");
    if (selectedPrimary && d.primaryId === selectedPrimary.id)                   cls.push("selected");
    if (selectedSecondary && d.secondaryId === selectedSecondary.id)             cls.push("secondary-selected");
    return cls.join(" ");
}

// country-primary 连线用国家色，primary-secondary 连线用学科色
function getLinkColor(d) {
    return d.linkType === "country-primary"
        ? countryPalette(d.country)
        : primaryPalette(getPrimaryLabel(d.primaryId));
}

function getNodeClass(d) {
    const cls = ["sankey-node", d.type];
    const sel = new Set(selectedCountries);
    if (d.type === "country") {
        if (sel.has(d.country)) cls.push("selected");
        else if (sel.size)      cls.push("context-node");
    }
    if (d.type === "primary") {
        if (selectedPrimary && selectedPrimary.id === d.primaryId) cls.push("selected");
        else if (selectedPrimary)                                  cls.push("context-node");
    }
    if (d.type === "secondary" && selectedSecondary && selectedSecondary.id === d.secondaryId)
        cls.push("selected");
    return cls.join(" ");
}

// country -> 国家色，primary/secondary -> 学科色，isOther -> 灰色
function getNodeColor(d) {
    if (d.type === "country")  return countryPalette(d.country);
    if (d.isOther)             return "#b7b7b7";
    return primaryPalette(d.primaryName || getPrimaryLabel(d.primaryId));
}

function getNodeTitle(d) {
    if (d.type === "country")   return d.country;
    if (d.type === "primary")   return d.primaryName;
    if (d.type === "secondary") return `${d.secondaryName}（${d.primaryName}）`;
    return d.name;
}

function getNodeHint(d) {
    if (d.type === "country")   return "点击选定或取消该国家";
    if (d.type === "primary")   return "点击展开该一级学科";
    if (d.type === "secondary") return d.isOther ? "其他二级学科合并项" : "点击高亮该二级学科";
    return "拖拽节点可调整布局";
}

// 节点占比的分母：country/primary 用全量总数，secondary 用该一级学科全量总数
function getNodeReferenceTotal(d, sankeyData) {
    const all = getCountryRecords();
    if (d.type === "country" || d.type === "primary") return sumRecords(all);
    if (d.type === "secondary" && selectedPrimary)
        return sumRecords(all.filter(r => r.primaryId === selectedPrimary.id));
    return sankeyData.total;
}

function handleSankeyNodeClick(event, d) {
    event.stopPropagation();
    if (d.type === "country") {
        pushHistory();
        if (selectedCountries.includes(d.country))
            selectedCountries = selectedCountries.filter(c => c !== d.country);
        else selectedCountries.push(d.country);
        renderSelectionOptions(); updateAllCharts();
    } else if (d.type === "primary") {
        pushHistory();
        if (selectedPrimary && selectedPrimary.id === d.primaryId) {
            selectedPrimary = null; selectedSecondary = null;
        } else {
            selectedPrimary = { id: d.primaryId, name: d.primaryName }; selectedSecondary = null;
        }
        updateAllCharts();
    } else if (d.type === "secondary" && !d.isOther) {
        pushHistory();
        selectedSecondary = (selectedSecondary && selectedSecondary.id === d.secondaryId)
            ? null
            : { id: d.secondaryId, name: d.secondaryName, primaryId: d.primaryId, primaryName: d.primaryName };
        updateAllCharts();
    }
}

// 雷达图
function getRadarEntities() {
    return selectedCountries.length ? selectedCountries : [RAW_DATA.aggregateName];
}

function getRadarConfig(entities) {
    if (!selectedPrimary) {
        return {
            mode: "primary",
            axes: RAW_DATA.primaryList.map(d => ({
                id: d.id, label: d.name, primaryId: d.id, members: [d.id], isOther: false
            })),
            title: "全部一级学科合作占比",
            denominatorLabel: "占全部学科"
        };
    }
    const combined = entities.flatMap(e => getRecords(e));
    return {
        mode: "secondary",
        axes: makeSecondaryBuckets(combined, selectedPrimary.id, SECONDARY_LIMIT,
                selectedSecondary ? selectedSecondary.id : null),
        primaryId: selectedPrimary.id, primaryName: selectedPrimary.name,
        title: `${selectedPrimary.name}：二级学科合作占比`,
        denominatorLabel: "占该一级学科"
    };
}

// 计算某实体在某雷达轴上的百分比
function computeRadarValue(entity, axis, config) {
    const records = getRecords(entity);
    if (config.mode === "primary") {
        return safePercent(
            d3.sum(records.filter(d => d.primaryId === axis.primaryId), d => d.count),
            sumRecords(records)
        );
    }
    const pr = records.filter(d => d.primaryId === config.primaryId);
    const memberSet = new Set(axis.members || [axis.secondaryId]);
    return safePercent(
        d3.sum(pr.filter(d => memberSet.has(d.secondaryId)), d => d.count),
        sumRecords(pr)
    );
}

function drawRadarChart() {
    const container = part3Root ? part3Root.querySelector("#radarChart") : document.getElementById("radarChart");
    container.innerHTML = "";

    const width    = container.clientWidth || 460;
    const entities = getRadarEntities();
    const config   = getRadarConfig(entities);

    d3.select(part3Root ? part3Root.querySelector("#radarSubtitle") : null).text(config.title);
    const entitiesLabel = entities.map(d => d === RAW_DATA.aggregateName ? "整体（全部国家合并）" : d).join("、");
    d3.select(part3Root ? part3Root.querySelector("#radarCaption") : null).html(
        `<p>当前口径：${config.denominatorLabel}；显示对象：${entitiesLabel}。${
            selectedPrimary
                ? `展开学科：${selectedPrimary.name}${selectedSecondary ? ` / ${selectedSecondary.name}` : ""}。`
                : "当前为一级学科总览。"
        }</p>`
    );

    if (!config.axes.length) {
        container.innerHTML = "<div class='empty-state'>暂无可展示数据</div>"; return;
    }

    const itemsPerRow = Math.max(1, Math.floor((width - 120) / 96));
    const legendRows  = Math.ceil(entities.length / itemsPerRow);
    const height = Math.max(container.clientHeight || 560, 540 + legendRows * 22);
    const margin = { top: 72, right: 72, bottom: 90 + legendRows * 12, left: 72 };
    const radius = Math.min(width - margin.left - margin.right, height - margin.top - margin.bottom) / 2;

    const series = entities.map(entity => ({
        entity,
        values: config.axes.map(axis => ({
            axis: axis.label, id: axis.id, secondaryId: axis.secondaryId,
            value: computeRadarValue(entity, axis, config), isOther: axis.isOther
        }))
    }));

    const domainMax = Math.max(10, Math.min(100,
        Math.ceil((d3.max(series, s => d3.max(s.values, d => d.value)) || 1) / 10) * 10));
    const rScale     = d3.scaleSqrt().domain([0, domainMax]).range([0, radius]);
    const angleSlice = Math.PI * 2 / config.axes.length;

    const svg = d3.select(container).append("svg")
        .attr("viewBox", `0 0 ${width} ${height}`)
        .attr("width", width).attr("height", height);
    const g = svg.append("g").attr("transform", `translate(${width / 2},${margin.top + radius})`);

    // 同心多边形网格
    for (let lv = 1; lv <= 5; lv++) {
        const lf = rScale(domainMax * lv / 5);
        g.append("polygon").attr("class", "radar-grid")
            .attr("points", d3.range(config.axes.length).map(i => {
                const a = angleSlice * i - Math.PI / 2;
                return `${Math.cos(a) * lf},${Math.sin(a) * lf}`;
            }).join(" "));
        g.append("text").attr("class", "radar-level").attr("x", 4).attr("y", -lf)
            .text(`${formatPercent(domainMax * lv / 5)}%`);
    }

    // 轴（辐射线 + 标签）
    const axis = g.selectAll(".radar-axis").data(config.axes).enter().append("g").attr("class", "radar-axis");
    axis.append("line")
        .attr("x1", 0).attr("y1", 0)
        .attr("x2", (d, i) => Math.cos(angleSlice * i - Math.PI / 2) * radius)
        .attr("y2", (d, i) => Math.sin(angleSlice * i - Math.PI / 2) * radius);

    const lbl = axis.append("g");
    lbl.append("title").text(d => d.label);  // 浏览器原生 tooltip3 显示完整学科名
    lbl.append("text")
        .attr("class", d => `radar-axis-label${selectedSecondary && d.secondaryId === selectedSecondary.id ? " selected" : ""}`)
        .attr("x", (d, i) => Math.cos(angleSlice * i - Math.PI / 2) * (radius + 20))
        .attr("y", (d, i) => Math.sin(angleSlice * i - Math.PI / 2) * (radius + 20))
        .attr("text-anchor", (d, i) => { const x = Math.cos(angleSlice * i - Math.PI / 2); return x > 0.25 ? "start" : x < -0.25 ? "end" : "middle"; })
        .attr("dy", "0.35em")
        .text(d => truncateLabel(d.label, config.mode === "primary" ? 10 : 7));

    // 系列（填充 + 轮廓 + 数据点）
    const radarLine = d3.lineRadial().curve(d3.curveLinearClosed)
        .radius(d => rScale(d.value)).angle((d, i) => i * angleSlice);
    const sg = g.selectAll(".radar-series").data(series).enter().append("g").attr("class", "radar-series");
    const manySeries = series.length > 5;

    sg.append("path").attr("class", "radar-area")
        .attr("d", d => radarLine(d.values))
        .attr("fill",   d => getRadarColor(d.entity))
        .attr("stroke", d => getRadarColor(d.entity))
        .style("fill-opacity", manySeries ? 0.04 : 0.10);

    sg.append("path").attr("class", "radar-line")
        .attr("d", d => radarLine(d.values))
        .attr("stroke", d => getRadarColor(d.entity))
        .style("stroke-opacity", manySeries ? 0.72 : 0.9);

    sg.selectAll("circle")
        .data(d => d.values.map(v => ({ ...v, entity: d.entity }))).enter()
        .append("circle").attr("class", "radar-point").attr("r", 3)
        .attr("cx", (d, i) => Math.cos(angleSlice * i - Math.PI / 2) * rScale(d.value))
        .attr("cy", (d, i) => Math.sin(angleSlice * i - Math.PI / 2) * rScale(d.value))
        .attr("fill", d => getRadarColor(d.entity))
        .on("mouseover", (event, d) => showtooltip3(event,
            `<strong>${d.entity} · ${d.axis}</strong><br>${config.denominatorLabel}：${formatPercent(d.value)}%`))
        .on("mousemove", movetooltip3).on("mouseout", hidetooltip3);

    // 图例
    const legend = svg.append("g").attr("class", "radar-legend")
        .attr("transform", `translate(${margin.left},${height - margin.bottom + 44})`);
    legend.selectAll("g").data(series).enter().append("g")
        .attr("transform", (d, i) => `translate(${(i % itemsPerRow) * 96},${Math.floor(i / itemsPerRow) * 20})`)
        .call(g2 => {
            g2.append("rect").attr("width", 12).attr("height", 12).attr("rx", 2)
                .attr("fill", d => getRadarColor(d.entity));
            g2.append("text").attr("x", 18).attr("y", 10)
                .text(d => d.entity === RAW_DATA.aggregateName ? "整体" : truncateLabel(d.entity, 8));
        });
}

function getRadarColor(entity) {
    return entity === RAW_DATA.aggregateName ? "#4b5563" : countryPalette(entity);
}

// 折线图
function drawTrendChart() {
    const container = part3Root ? part3Root.querySelector("#trendChart") : document.getElementById("trendChart");
    container.innerHTML = "";

    const width  = container.clientWidth  || 1100;
    const height = container.clientHeight || 430;
    const margin = { top: 24, right: 190, bottom: 55, left: 70 };
    const cw = width - margin.left - margin.right;
    const ch = height - margin.top - margin.bottom;

    const svg = d3.select(container).append("svg")
        .attr("viewBox", `0 0 ${width} ${height}`)
        .attr("width", width).attr("height", height);
    const g = svg.append("g").attr("transform", `translate(${margin.left},${margin.top})`);

    const periods = RAW_DATA.periods;
    const periodTotals = new Map(periods.map(p => [p, sumRecords(getRecords(RAW_DATA.aggregateName, p))]));

    const trendSeries = RAW_DATA.primaryList.map(field => ({
        field,
        values: periods.map(period => {
            const recs  = getRecords(RAW_DATA.aggregateName, period);
            const value = d3.sum(recs.filter(d => d.primaryId === field.id), d => d.count);
            return { period, fieldId: field.id, fieldName: field.name, value,
                     percent: safePercent(value, periodTotals.get(period) || 0) };
        })
    }));

    const xScale = d3.scalePoint().domain(periods).range([0, cw]).padding(0.45);
    const yMax   = Math.ceil((d3.max(trendSeries, s => d3.max(s.values, d => d.percent)) || 1) / 5) * 5;
    const yScale = d3.scaleLinear().domain([0, yMax]).nice().range([ch, 0]);

    g.append("g").attr("class", "grid")
        .call(d3.axisLeft(yScale).ticks(6).tickSize(-cw).tickFormat(""));
    g.append("g").attr("class", "axis x-axis").attr("transform", `translate(0,${ch})`).call(d3.axisBottom(xScale));
    g.append("g").attr("class", "axis y-axis").call(d3.axisLeft(yScale).ticks(6).tickFormat(d => `${d}%`));
    g.append("text").attr("class", "axis-label").attr("x", cw / 2).attr("y", ch + 42).attr("text-anchor", "middle").text("时间段");
    g.append("text").attr("class", "axis-label").attr("transform", "rotate(-90)").attr("x", -ch / 2).attr("y", -50).attr("text-anchor", "middle").text("合作占比");

    // 单调三次插值，视觉上比折线更平滑
    const line = d3.line().x(d => xScale(d.period)).y(d => yScale(d.percent)).curve(d3.curveMonotoneX);

    // 点击折线/图例展开对应一级学科
    const selectPrimary = (event, d) => {
        pushHistory();
        selectedPrimary = { id: d.field.id, name: d.field.name };
        selectedSecondary = null;
        updateAllCharts();
    };

    const fg = g.selectAll(".trend-series").data(trendSeries).enter().append("g")
        .attr("class", d => `trend-series${selectedPrimary && selectedPrimary.id === d.field.id ? " selected" : ""}`)
        .on("click", selectPrimary);

    fg.append("path").attr("class", "trend-line")
        .attr("d", d => line(d.values)).attr("stroke", d => primaryPalette(d.field.name));

    fg.selectAll("circle").data(d => d.values).enter().append("circle")
        .attr("class", "trend-point").attr("r", 4)
        .attr("cx", d => xScale(d.period)).attr("cy", d => yScale(d.percent))
        .attr("fill", d => primaryPalette(d.fieldName))
        .on("mouseover", (event, d) => showtooltip3(event,
            `<strong>${d.fieldName} · ${d.period}</strong><br>
            合作论文数：${formatNumber(d.value)} 篇<br>
            占比：${formatPercent(d.percent)}%`))
        .on("mousemove", movetooltip3).on("mouseout", hidetooltip3);

    // 右侧图例
    const li = svg.append("g").attr("class", "trend-legend")
        .attr("transform", `translate(${margin.left + cw + 28},${margin.top + 18})`)
        .selectAll("g").data(trendSeries).enter().append("g")
        .attr("class", "trend-legend-item").attr("transform", (d, i) => `translate(0,${i * 24})`)
        .on("click", selectPrimary);
    li.append("line").attr("x1", 0).attr("x2", 18).attr("y1", 6).attr("y2", 6)
        .attr("stroke", d => primaryPalette(d.field.name)).attr("stroke-width", 2);
    li.append("text").attr("x", 26).attr("y", 10).text(d => truncateLabel(d.field.name, 9));
}
})();