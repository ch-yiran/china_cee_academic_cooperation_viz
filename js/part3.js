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

// 配色：统一使用 ./js/colors.js 中暴露的 window.SHARED_*，避免各模块重复维护颜色。
// main.html 中请保持 <script src="./js/colors.js"></script> 位于 part3.js 之前。
let primaryPalette = null;   // 一级学科名 -> 通用一级学科色
let countryPalette = null;   // 国家名     -> 通用国家色
const part3Root = document.getElementById("part3");

function getSharedPrimaryColors() {
    if (Array.isArray(window.SHARED_PRIMARY_COLORS) && window.SHARED_PRIMARY_COLORS.length)
        return window.SHARED_PRIMARY_COLORS;
    console.warn("[part3] 未检测到 window.SHARED_PRIMARY_COLORS，请确认 colors.js 已在 part3.js 前加载。");
    return ["#9aa6b2"];
}

function getSharedCountryColors() {
    if (Array.isArray(window.SHARED_COUNTRY_COLORS) && window.SHARED_COUNTRY_COLORS.length)
        return window.SHARED_COUNTRY_COLORS;
    console.warn("[part3] 未检测到 window.SHARED_COUNTRY_COLORS，请确认 colors.js 已在 part3.js 前加载。");
    return getSharedPrimaryColors();
}

function getSharedPaletteColor(index = 0) {
    const colors = getSharedPrimaryColors();
    return colors[((index % colors.length) + colors.length) % colors.length];
}

function getCommonPrimaryColor(name) {
    return primaryPalette ? primaryPalette(name) : getSharedPaletteColor(0);
}

function getCommonCountryColor(name) {
    if (typeof window.getSharedCountryColor === "function") return window.getSharedCountryColor(name);
    return countryPalette ? countryPalette(name) : getSharedPaletteColor(0);
}

// 趋势图的上升/下降/持平不再单独硬编码红绿灰，改为从通用色板中取固定槽位。
// 当前 colors.js 中：1 偏绿，7 偏红，6 偏灰蓝，可与全站视觉保持一致。
function getCommonStatusColor(status) {
    if (status === "上升") return getSharedPaletteColor(1);
    if (status === "下降") return getSharedPaletteColor(7);
    return getSharedPaletteColor(6);
}

function colorWithAlpha(color, alpha = 0.14) {
    const text = String(color || "").trim();
    if (/^#[0-9a-fA-F]{6}$/.test(text)) {
        const r = parseInt(text.slice(1, 3), 16);
        const g = parseInt(text.slice(3, 5), 16);
        const b = parseInt(text.slice(5, 7), 16);
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }
    if (/^#[0-9a-fA-F]{3}$/.test(text)) {
        const r = parseInt(text[1] + text[1], 16);
        const g = parseInt(text[2] + text[2], 16);
        const b = parseInt(text[3] + text[3], 16);
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }
    return color;
}

function getTrendDeltaColor(status) {
    if (status === "上升") return "#1a9850";
    if (status === "下降") return "#d73027";
    return "#8c96a3";
}

const tooltip3      = d3.select("#tooltip3");
const formatNumber  = d3.format(",");
const formatPercent = d3.format(".1f");

// 初始化
document.addEventListener("DOMContentLoaded", init);

// 自动确保 d3-sankey 已加载：优先使用本地文件，其次尝试多个 CDN。
// 推荐在 js/ 目录放置 d3-sankey.min.js，这样可完全摆脱 CDN 网络限制。
async function ensureD3Sankey() {
    if (typeof d3 === "undefined") {
        throw new Error("未检测到 D3，请确认 d3.min.js 已在 part3.js 之前引入。");
    }
    if (d3.sankey && d3.sankeyLinkHorizontal) return;

    const sources = [
        "./js/d3-sankey.min.js",
        "https://cdnjs.cloudflare.com/ajax/libs/d3-sankey/0.12.3/d3-sankey.min.js",
        "https://unpkg.com/d3-sankey@0.12.3/dist/d3-sankey.min.js",
        "https://cdn.jsdelivr.net/npm/d3-sankey@0.12.3/dist/d3-sankey.min.js"
    ];

    const errors = [];
    for (const src of sources) {
        try {
            await loadScriptOnce(src);
            if (d3.sankey && d3.sankeyLinkHorizontal) return;
            errors.push(`${src}：脚本已加载，但没有注册 d3.sankey`);
        } catch (err) {
            errors.push(`${src}：${err.message || err}`);
        }
    }

    throw new Error(
        "d3-sankey 未加载成功。建议将 d3-sankey.min.js 下载到 js/ 目录，或确认浏览器能访问 cdnjs / unpkg / jsDelivr。"
    );
}

function loadScriptOnce(src) {
    return new Promise((resolve, reject) => {
        const absoluteSrc = new URL(src, window.location.href).href;
        const loaded = Array.from(document.scripts).find(s =>
            s.src === absoluteSrc && s.dataset.loaded === "true"
        );
        if (loaded) {
            resolve();
            return;
        }

        const script = document.createElement("script");
        script.src = src;
        script.async = true;
        script.crossOrigin = "anonymous";
        script.onload = () => {
            script.dataset.loaded = "true";
            resolve();
        };
        script.onerror = () => reject(new Error("加载失败"));
        document.head.appendChild(script);
    });
}

async function init() {
    showLoading("正在加载 d3-sankey 与 Excel 数据...");
    try {
        await ensureD3Sankey();
        RAW_DATA = await loadRawDataFromExcel();

        selectedCountries = [];
        selectedPrimary   = null;
        selectedSecondary = null;
        undoStack = []; redoStack = [];

        primaryPalette = d3.scaleOrdinal()
            .domain(RAW_DATA.primaryList.map(d => d.name))
            .range(getSharedPrimaryColors())
            .unknown(getSharedPaletteColor(6));
        countryPalette = d3.scaleOrdinal()
            .domain(RAW_DATA.countries)
            .range(getSharedCountryColors())
            .unknown(getSharedPaletteColor(6));

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
        .data(RAW_DATA.countries)
        .enter()
        .append("label")
        .attr("class", d => `compare-pill${selectedCountries.includes(d) ? " selected" : ""}`)
        .style("border-color", d => selectedCountries.includes(d) ? getCommonCountryColor(d) : null)
        .style("background-color", d => selectedCountries.includes(d) ? colorWithAlpha(getCommonCountryColor(d), 0.16) : null)
        .style("color", d => selectedCountries.includes(d) ? getCommonCountryColor(d) : null);

    labels.append("input")
        .attr("type", "checkbox")
        .attr("value", d => d)
        .property("checked", d => selectedCountries.includes(d))
        .style("accent-color", d => getCommonCountryColor(d))
        .on("change", function(event, country) {
            pushHistory();
            if (this.checked) {
                if (!selectedCountries.includes(country)) selectedCountries.push(country);
            } else {
                selectedCountries = selectedCountries.filter(d => d !== country);
            }
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
        container.innerHTML = "<div class='empty-state'>d3-sankey 仍未加载。请将 d3-sankey.min.js 放入 js/ 目录，或检查 CDN 网络。</div>";
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
        .text(d => truncateLabel(d.name, d.type === "secondary" ? 11 : 14));

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
        ? getCommonCountryColor(d.country)
        : getCommonPrimaryColor(getPrimaryLabel(d.primaryId));
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
    if (d.type === "country")  return getCommonCountryColor(d.country);
    if (d.isOther)             return getSharedPaletteColor(6);
    return getCommonPrimaryColor(d.primaryName || getPrimaryLabel(d.primaryId));
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

    const width  = container.clientWidth  || 520;
    const height = Math.max(container.clientHeight || 640, 640);
    const margin = { top: 22, right: 48, bottom: 210, left: 48 };
    const cx = width / 2;
    const cy = margin.top + (height - margin.top - margin.bottom) / 2 - 22;
    const radius = Math.max(104, Math.min(width - margin.left - margin.right, height - margin.top - margin.bottom) / 2 - 30);

    const svg = d3.select(container).append("svg")
        .attr("viewBox", `0 0 ${width} ${height}`)
        .attr("width", width)
        .attr("height", height);

    const entities = getRadarEntities();
    const recordsByEntity = new Map(entities.map(entity => [entity, getRecords(entity)]));

    function buildRadarAxes() {
        if (!selectedPrimary) {
            // 一级学科固定展示全部维度，正常情况下为 13 个方向。
            return RAW_DATA.primaryList.map(d => ({
                id: d.id,
                name: d.name,
                type: "primary"
            }));
        }

        const baseRecords = getRecords(RAW_DATA.aggregateName).filter(d => d.primaryId === selectedPrimary.id);
        const grouped = Array.from(
            d3.rollup(
                baseRecords,
                v => ({
                    value: d3.sum(v, d => d.count),
                    secondaryName: v[0].secondaryName,
                    secondaryId: v[0].secondaryId
                }),
                d => d.secondaryId
            ),
            ([secondaryId, info]) => ({
                id: secondaryId,
                name: info.secondaryName,
                type: "secondary",
                secondaryId,
                members: [secondaryId],
                value: info.value,
                isOther: false
            })
        ).sort((a, b) => d3.descending(a.value, b.value));

        if (grouped.length <= 8) return grouped;

        const kept = grouped.slice(0, 7);
        const keptIds = new Set(kept.map(d => d.secondaryId));
        const others = grouped.filter(d => !keptIds.has(d.secondaryId));

        kept.push({
            id: `other:${selectedPrimary.id}`,
            name: "其他",
            type: "secondary",
            secondaryId: `other:${selectedPrimary.id}`,
            members: others.map(d => d.secondaryId),
            value: d3.sum(others, d => d.value),
            isOther: true
        });

        return kept;
    }

    const axes = buildRadarAxes();

    if (!axes.length) {
        svg.append("text")
            .attr("class", "empty-state-text")
            .attr("x", width / 2)
            .attr("y", height / 2)
            .attr("text-anchor", "middle")
            .text("暂无可展示数据");
        return;
    }

    const radarData = entities.map(entity => {
        const records = recordsByEntity.get(entity) || [];
        const total = sumRecords(records);
        const values = axes.map(axis => {
            let value = 0;
            if (axis.type === "primary") {
                value = sumRecords(records.filter(d => d.primaryId === axis.id));
            } else if (axis.isOther) {
                const memberSet = new Set(axis.members || []);
                value = sumRecords(records.filter(d =>
                    d.primaryId === selectedPrimary.id && memberSet.has(d.secondaryId)
                ));
            } else {
                value = sumRecords(records.filter(d => d.secondaryId === axis.secondaryId));
            }
            return {
                axis,
                value,
                percent: safePercent(value, total)
            };
        });
        return { entity, total, values };
    });

    const maxValue = Math.max(1, d3.max(radarData.flatMap(d => d.values.map(v => v.percent))) || 1);
    const rMax = Math.max(5, Math.ceil(maxValue / 5) * 5);

    // 一级学科差别很大时，用更强的非线性放大；二级学科稍弱一些。
    const rScale = d3.scalePow()
        .exponent(selectedPrimary ? 0.46 : 0.24)
        .domain([0, rMax])
        .range([0, radius]);

    // 关键修正：不用 scalePoint 的 [0, 2π] 端点，否则第一个和最后一个维度会重叠。
    const angleMap = new Map(axes.map((d, i) => [d.id, (Math.PI * 2 * i) / axes.length]));
    function angleOf(axisId) {
        return angleMap.get(axisId) ?? 0;
    }

    function point(axisId, percent, extra = 0) {
        const a = angleOf(axisId) - Math.PI / 2;
        const r = rScale(percent) + extra;
        return [cx + Math.cos(a) * r, cy + Math.sin(a) * r];
    }

    function wrapText(selection, maxChars = 6) {
        selection.each(function(textValue) {
            const text = d3.select(this);
            const raw = String(textValue || "");
            text.text(null);

            if (raw.length <= maxChars) {
                text.append("tspan").attr("x", 0).attr("dy", 0).text(raw);
                return;
            }

            const parts = [];
            for (let i = 0; i < raw.length; i += maxChars) parts.push(raw.slice(i, i + maxChars));
            parts.slice(0, 2).forEach((part, i) => {
                text.append("tspan")
                    .attr("x", 0)
                    .attr("dy", i === 0 ? 0 : "1.16em")
                    .text(part);
            });
            if (parts.length > 2) {
                const last = text.select("tspan:last-child");
                last.text(last.text().replace(/.$/, "…"));
            }
        });
    }

    const g = svg.append("g");

    const levels = d3.range(1, 6).map(i => rMax * i / 5);
    g.selectAll(".radar-grid")
        .data(levels)
        .enter().append("circle")
        .attr("class", "radar-grid")
        .attr("cx", cx)
        .attr("cy", cy)
        .attr("r", d => rScale(d));

    g.selectAll(".radar-level")
        .data(levels)
        .enter().append("text")
        .attr("class", "radar-level")
        .attr("x", cx + 8)
        .attr("y", d => cy - rScale(d) + 4)
        .text(d => `${formatPercent(d)}%`);

    const axisG = g.selectAll(".radar-axis")
        .data(axes)
        .enter().append("g")
        .attr("class", "radar-axis");

    axisG.append("line")
        .attr("x1", cx)
        .attr("y1", cy)
        .attr("x2", d => point(d.id, rMax)[0])
        .attr("y2", d => point(d.id, rMax)[1]);

    // 标签与图形保持小距离，避免显示不全。
    const labelDistance = selectedPrimary ? 22 : 20;
    const labels = axisG.append("text")
        .attr("class", d => {
            const selected = selectedSecondary && selectedSecondary.id === d.id;
            return `radar-axis-label${selected ? " selected" : ""}`;
        })
        .attr("transform", d => {
            const a = angleOf(d.id) - Math.PI / 2;
            const x = cx + Math.cos(a) * (radius + labelDistance);
            const y = cy + Math.sin(a) * (radius + labelDistance);
            return `translate(${x},${y})`;
        })
        .attr("text-anchor", d => {
            const a = angleOf(d.id) - Math.PI / 2;
            const x = Math.cos(a);
            if (Math.abs(x) < 0.16) return "middle";
            return x > 0 ? "start" : "end";
        })
        .attr("dominant-baseline", d => {
            const y = Math.sin(angleOf(d.id) - Math.PI / 2);
            if (y < -0.55) return "auto";
            if (y > 0.55) return "hanging";
            return "middle";
        })
        .datum(d => truncateLabel(d.name, selectedPrimary ? 14 : 10));

    labels.call(wrapText, selectedPrimary ? 7 : 5);

    const line = d3.lineRadial()
        .angle((d, i) => (Math.PI * 2 * i) / axes.length)
        .radius(d => rScale(d.percent))
        .curve(d3.curveLinearClosed);

    const series = g.selectAll(".radar-series")
        .data(radarData)
        .enter().append("g")
        .attr("class", "radar-series");

    series.append("path")
        .attr("class", "radar-area")
        .attr("transform", `translate(${cx},${cy})`)
        .attr("d", d => line(d.values))
        .attr("fill", d => getRadarColor(d.entity))
        .attr("stroke", d => getRadarColor(d.entity));

    series.append("path")
        .attr("class", "radar-line")
        .attr("transform", `translate(${cx},${cy})`)
        .attr("d", d => line(d.values))
        .attr("stroke", d => getRadarColor(d.entity));

    series.selectAll(".radar-point")
        .data(d => d.values.map(v => ({ ...v, entity: d.entity })))
        .enter().append("circle")
        .attr("class", "radar-point")
        .attr("cx", d => point(d.axis.id, d.percent)[0])
        .attr("cy", d => point(d.axis.id, d.percent)[1])
        .attr("r", 5.2)
        .attr("fill", d => getRadarColor(d.entity))
        .on("mouseover", (event, d) => showtooltip3(event,
            `<strong>${d.entity} · ${d.axis.name}</strong><br>
            合作论文数：${formatNumber(d.value)} 篇<br>
            占比：${formatPercent(d.percent)}%`))
        .on("mousemove", movetooltip3)
        .on("mouseout", hidetooltip3);

    const legend = svg.append("g")
        .attr("class", "radar-legend")
        .attr("transform", `translate(${Math.max(20, width / 2 - 170)}, ${height - 158})`);

    const legendItem = legend.selectAll("g")
        .data(entities)
        .enter().append("g")
        .attr("transform", (d, i) => {
            const col = i % 2;
            const row = Math.floor(i / 2);
            return `translate(${col * 165},${row * 26})`;
        });

    legendItem.append("rect")
        .attr("width", 18)
        .attr("height", 18)
        .attr("rx", 3)
        .attr("fill", d => getRadarColor(d));

    legendItem.append("text")
        .attr("x", 26)
        .attr("y", 15)
        .text(d => d);

    const radarSubtitle = part3Root ? part3Root.querySelector("#radarSubtitle") : document.getElementById("radarSubtitle");
    if (radarSubtitle) {
        radarSubtitle.textContent = selectedPrimary
            ? `${selectedPrimary.name}：二级学科结构`
            : `一级学科结构（${axes.length} 个维度）`;
    }

    const caption = part3Root ? part3Root.querySelector("#radarCaption") : document.getElementById("radarCaption");
    if (caption) {
        caption.innerHTML = "";
        caption.style.display = "none";
    }
}

function getRadarColor(entity) {
    return entity === RAW_DATA.aggregateName ? getSharedPaletteColor(6) : getCommonCountryColor(entity);
}

function getTrendStatusColor(status) {
    return getCommonStatusColor(status);
}

// 折线图
// 折线图：改为更接近 Part4 风格的双时段斜率图，减少拥挤并提升可读性





function drawTrendChart() {
    const container = part3Root ? part3Root.querySelector("#trendChart") : document.getElementById("trendChart");
    container.innerHTML = "";

    const width  = container.clientWidth  || 1080;
    const height = Math.max(container.clientHeight || 720, 720);
    const margin = { top: 78, right: 210, bottom: 56, left: 210 };
    const cw = Math.max(390, width - margin.left - margin.right);
    const ch = Math.max(540, height - margin.top - margin.bottom);

    const svg = d3.select(container).append("svg")
        .attr("viewBox", `0 0 ${width} ${height}`)
        .attr("width", width)
        .attr("height", height);

    const g = svg.append("g").attr("transform", `translate(${margin.left},${margin.top})`);

    const periods = ["2011-2015", "2016-2020"];
    const periodTotals = new Map(periods.map(period => [
        period,
        sumRecords(getRecords(RAW_DATA.aggregateName, period))
    ]));

    const data = RAW_DATA.primaryList.map(field => {
        const leftRecords  = getRecords(RAW_DATA.aggregateName, periods[0]).filter(d => d.primaryId === field.id);
        const rightRecords = getRecords(RAW_DATA.aggregateName, periods[1]).filter(d => d.primaryId === field.id);
        const leftValue  = sumRecords(leftRecords);
        const rightValue = sumRecords(rightRecords);
        const leftPercent  = safePercent(leftValue,  periodTotals.get(periods[0]));
        const rightPercent = safePercent(rightValue, periodTotals.get(periods[1]));
        const delta = rightPercent - leftPercent;
        const status = Math.abs(delta) < 0.05 ? "持平" : (delta > 0 ? "上升" : "下降");
        return { field, leftValue, rightValue, leftPercent, rightPercent, delta, status };
    }).sort((a, b) =>
        d3.descending(Math.max(a.leftPercent, a.rightPercent), Math.max(b.leftPercent, b.rightPercent)) ||
        d3.descending(a.rightPercent, b.rightPercent)
    );

    if (!data.length) {
        g.append("text")
            .attr("class", "empty-state-text")
            .attr("x", cw / 2)
            .attr("y", ch / 2)
            .attr("text-anchor", "middle")
            .text("暂无可展示数据");
        return;
    }

    // 分别计算两个时间段的学科排名，用于左右标签显示 No.x。
    const leftRankMap = new Map(
        [...data]
            .sort((a, b) => d3.descending(a.leftPercent, b.leftPercent) || d3.ascending(a.field.name, b.field.name))
            .map((d, i) => [d.field.id, i + 1])
    );
    const rightRankMap = new Map(
        [...data]
            .sort((a, b) => d3.descending(a.rightPercent, b.rightPercent) || d3.ascending(a.field.name, b.field.name))
            .map((d, i) => [d.field.id, i + 1])
    );

    // 进一步放大小学科区间：指数越小，0–5% 区间拉伸越明显。
    const maxPercent = Math.max(1, d3.max(data, d => Math.max(d.leftPercent, d.rightPercent)) || 1);
    const yMax = Math.ceil(maxPercent / 5) * 5;
    const yScale = d3.scalePow()
        .exponent(0.28)
        .domain([0, yMax])
        .range([ch, 0]);

    const lineGapRatio = 0.56;
    const lineInset = cw * (1 - lineGapRatio) / 2;
    const xScale = d3.scalePoint()
        .domain(periods)
        .range([lineInset, cw - lineInset]);

    function adjustedPositions(values, minGap = 32) {
        const sorted = values
            .map(d => ({ ...d, originalY: yScale(d.percent) }))
            .sort((a, b) => a.originalY - b.originalY);

        sorted.forEach((d, i) => {
            d.labelY = i === 0 ? d.originalY : Math.max(d.originalY, sorted[i - 1].labelY + minGap);
        });

        const overflow = sorted.length ? sorted[sorted.length - 1].labelY - ch : 0;
        if (overflow > 0) sorted.forEach(d => { d.labelY -= overflow; });

        sorted.forEach((d, i) => {
            if (i === 0) d.labelY = Math.max(0, d.labelY);
            else d.labelY = Math.max(d.labelY, sorted[i - 1].labelY + minGap);
            d.labelY = Math.min(ch, d.labelY);
        });

        for (let i = sorted.length - 2; i >= 0; i--) {
            sorted[i].labelY = Math.min(sorted[i].labelY, sorted[i + 1].labelY - minGap);
            sorted[i].labelY = Math.max(0, sorted[i].labelY);
        }

        return new Map(sorted.map(d => [d.id, d.labelY]));
    }

    const leftLabelY = adjustedPositions(data.map(d => ({
        id: d.field.id,
        percent: d.leftPercent
    })), 32);

    const rightLabelY = adjustedPositions(data.map(d => ({
        id: d.field.id,
        percent: d.rightPercent
    })), 32);

    const tickValues = [0, 0.2, 0.5, 1, 2, 5, 10, 20, 30, 40, 50, yMax]
        .filter((d, i, arr) => d <= yMax && arr.indexOf(d) === i);

    const grid = g.append("g").attr("class", "grid trend-grid");

    grid.selectAll("line")
        .data(tickValues)
        .enter().append("line")
        .attr("x1", xScale(periods[0]))
        .attr("x2", xScale(periods[1]))
        .attr("y1", d => yScale(d))
        .attr("y2", d => yScale(d))
        .attr("stroke", "#edf1f5")
        .attr("stroke-width", 1);

    grid.selectAll("text")
        .data(tickValues)
        .enter().append("text")
        .attr("class", "trend-grid-percent")
        .attr("x", xScale(periods[0]) + 12)
        .attr("y", d => yScale(d) - 7)
        .attr("text-anchor", "start")
        .attr("fill", "#8c96a3")
        .text(d => `${formatPercent(d)}%`);

    g.append("line")
        .attr("class", "trend-period-axis")
        .attr("x1", xScale(periods[0]))
        .attr("x2", xScale(periods[0]))
        .attr("y1", -10)
        .attr("y2", ch + 10)
        .attr("stroke", "#26384a")
        .attr("stroke-width", 2.8);

    g.append("line")
        .attr("class", "trend-period-axis")
        .attr("x1", xScale(periods[1]))
        .attr("x2", xScale(periods[1]))
        .attr("y1", -10)
        .attr("y2", ch + 10)
        .attr("stroke", "#26384a")
        .attr("stroke-width", 2.8);

    svg.append("text")
        .attr("class", "trend-period-label")
        .attr("x", margin.left + xScale(periods[0]))
        .attr("y", margin.top - 30)
        .attr("text-anchor", "middle")
        .attr("font-weight", 700)
        .attr("fill", "#333333")
        .text("2011–2015年");

    svg.append("text")
        .attr("class", "trend-period-label")
        .attr("x", margin.left + xScale(periods[1]))
        .attr("y", margin.top - 30)
        .attr("text-anchor", "middle")
        .attr("font-weight", 700)
        .attr("fill", "#333333")
        .text("2016–2020年");

    const lineGen = d3.line()
        .x(d => xScale(d.period))
        .y(d => yScale(d.percent))
        .curve(d3.curveBumpX);

    const series = g.selectAll(".trend-series")
        .data(data)
        .enter().append("g")
        .attr("class", d => {
            const statusClass = d.status === "上升" ? "rising" : (d.status === "下降" ? "falling" : "flat");
            return `trend-series ${statusClass}${selectedPrimary && selectedPrimary.id === d.field.id ? " selected" : ""}`;
        })
        .on("click", (event, d) => {
            pushHistory();
            selectedPrimary = selectedPrimary && selectedPrimary.id === d.field.id
                ? null
                : { id: d.field.id, name: d.field.name };
            selectedSecondary = null;
            updateAllCharts();
        });

    series.append("path")
        .attr("class", "trend-line")
        .attr("d", d => lineGen([
            { period: periods[0], percent: d.leftPercent },
            { period: periods[1], percent: d.rightPercent }
        ]))
        .style("stroke", d => getCommonPrimaryColor(d.field.name))
        .attr("stroke-width", d => selectedPrimary && selectedPrimary.id === d.field.id ? 5.8 : (Math.max(d.leftPercent, d.rightPercent) >= 5 ? 4.8 : 3.7))
        .attr("opacity", d => selectedPrimary && selectedPrimary.id && selectedPrimary.id !== d.field.id ? 0.22 : 0.9)
        .attr("fill", "none")
        .on("mouseover", (event, d) => showtooltip3(event,
            `<strong>${d.field.name}</strong><br>
            ${periods[0]}：No.${leftRankMap.get(d.field.id)}，${formatPercent(d.leftPercent)}%，${formatNumber(d.leftValue)} 篇<br>
            ${periods[1]}：No.${rightRankMap.get(d.field.id)}，${formatPercent(d.rightPercent)}%，${formatNumber(d.rightValue)} 篇<br>
            变化：${d.delta >= 0 ? "+" : ""}${formatPercent(d.delta)} pct`))
        .on("mousemove", movetooltip3)
        .on("mouseout", hidetooltip3);

    const pointData = data.flatMap(d => [
        {
            series: d,
            period: periods[0],
            value: d.leftValue,
            percent: d.leftPercent,
            fieldName: d.field.name
        },
        {
            series: d,
            period: periods[1],
            value: d.rightValue,
            percent: d.rightPercent,
            fieldName: d.field.name
        }
    ]);

    g.selectAll(".trend-point")
        .data(pointData)
        .enter().append("circle")
        .attr("class", "trend-point")
        .attr("cx", d => xScale(d.period))
        .attr("cy", d => yScale(d.percent))
        .attr("r", d => Math.max(d.series.leftPercent, d.series.rightPercent) >= 5 ? 6.5 : 5.4)
        .style("fill", d => getCommonPrimaryColor(d.series.field.name))
        .attr("stroke", "#ffffff")
        .attr("stroke-width", 2)
        .on("mouseover", (event, d) => showtooltip3(event,
            `<strong>${d.fieldName} · ${d.period}</strong><br>
            排名：No.${d.period === periods[0] ? leftRankMap.get(d.series.field.id) : rightRankMap.get(d.series.field.id)}<br>
            合作论文数：${formatNumber(d.value)} 篇<br>
            占比：${formatPercent(d.percent)}%<br>
            变化方向：${d.series.status}`))
        .on("mousemove", movetooltip3)
        .on("mouseout", hidetooltip3);

    const labelConnector = g.append("g").attr("class", "trend-label-connectors");

    labelConnector.selectAll(".left-connector")
        .data(data)
        .enter().append("path")
        .attr("d", d => {
            const x0 = xScale(periods[0]);
            const y0 = yScale(d.leftPercent);
            const x1 = lineInset - 34;
            const y1 = leftLabelY.get(d.field.id);
            return `M${x0},${y0} C${x0 - 22},${y0} ${x1 + 18},${y1} ${x1},${y1}`;
        })
        .attr("fill", "none")
        .attr("stroke", "#d7dde5")
        .attr("stroke-width", 1);

    labelConnector.selectAll(".right-connector")
        .data(data)
        .enter().append("path")
        .attr("d", d => {
            const x0 = xScale(periods[1]);
            const y0 = yScale(d.rightPercent);
            const x1 = cw - lineInset + 34;
            const y1 = rightLabelY.get(d.field.id);
            return `M${x0},${y0} C${x0 + 22},${y0} ${x1 - 18},${y1} ${x1},${y1}`;
        })
        .attr("fill", "none")
        .attr("stroke", "#d7dde5")
        .attr("stroke-width", 1);

    const leftLabels = g.append("g").attr("class", "trend-side-labels left")
        .selectAll("text")
        .data(data)
        .enter().append("text")
        .attr("class", "trend-label-left")
        .attr("x", lineInset - 42)
        .attr("y", d => leftLabelY.get(d.field.id) + 5)
        .attr("text-anchor", "end")
        .attr("fill", "#333333");

    leftLabels.append("tspan")
        .attr("font-weight", 700)
        .text(d => d.field.name);

    leftLabels.append("tspan")
        .attr("class", "trend-label-rank")
        .attr("dx", 8)
        .attr("font-weight", 700)
        .style("fill", "#a66f6f")
        .text(d => `No.${leftRankMap.get(d.field.id)}`);

    leftLabels.append("tspan")
        .attr("class", "trend-label-muted")
        .attr("dx", 8)
        .attr("font-weight", 400)
        .text(d => `${formatPercent(d.leftPercent)}%`);

    const rightLabels = g.append("g").attr("class", "trend-side-labels right")
        .selectAll("text")
        .data(data)
        .enter().append("text")
        .attr("class", "trend-label-right")
        .attr("x", cw - lineInset + 42)
        .attr("y", d => rightLabelY.get(d.field.id) + 5)
        .attr("text-anchor", "start")
        .attr("fill", "#333333");

    rightLabels.append("tspan")
        .attr("font-weight", 700)
        .style("fill", d => {
            const leftRank = leftRankMap.get(d.field.id);
            const rightRank = rightRankMap.get(d.field.id);
            if (rightRank < leftRank) return getTrendDeltaColor("上升");
            if (rightRank > leftRank) return getTrendDeltaColor("下降");
            return "#333333";
        })
        .text(d => d.field.name);

    rightLabels.append("tspan")
        .attr("class", "trend-label-rank")
        .attr("dx", 8)
        .attr("font-weight", 700)
        .style("fill", "#a66f6f")
        .text(d => `No.${rightRankMap.get(d.field.id)}`);

    rightLabels.append("tspan")
        .attr("class", "trend-label-muted")
        .attr("dx", 8)
        .attr("font-weight", 400)
        .text(d => `${formatPercent(d.rightPercent)}%`);

    rightLabels.append("tspan")
        .attr("class", "trend-label-delta")
        .attr("dx", 8)
        .attr("font-weight", 700)
        .style("fill", d => getTrendDeltaColor(d.status))
        .text(d => Math.abs(d.delta) < 0.05 ? "—" : `${d.delta >= 0 ? "+" : ""}${formatPercent(d.delta)}pct`);
}
})();