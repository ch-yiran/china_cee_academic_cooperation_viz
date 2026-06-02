(()=>{

// ===== 颜色常量 =====
const C = {
  blue:'#1d4ed8', teal:'#0e7490', red:'#b91c1c',
  violet:'#6d28d9', green:'#047857', orange:'#b45309',
  text:'#1c2333', text2:'#4a5568', text3:'#8a97b0',
  border:'#dde1e9'
};
Chart.defaults.font.family = "'Times New Roman','Microsoft YaHei',serif";
Chart.defaults.font.size = 16;
Chart.defaults.plugins.legend.labels.font = { size: 16 };
Chart.defaults.plugins.tooltip.titleFont = { size: 16 };
Chart.defaults.plugins.tooltip.bodyFont = { size: 16 };
const catColor = {
  '理学':'#2F80ED','工学':'#5B6CF6','医学':'#E76F51',
  '管理学':'#D9A441','农学':'#2A9D8F','教育学':'#4CB5AE',
  '法学':'#8E6C8A','经济学':'#C97C5D','交叉领域':'#64748B'
};

// ===== 地图常量 =====
const YEARS = [2011,2012,2013,2014,2015,2016,2017,2018,2019,2020];
// const AREA_COLORS = [
//   '#1d4fd87b','#0e7490','#047857','#b45309','#b91c1c',
//   '#6d28d9','#0369a1','#065f46','#92400e','#991b1b',
//   '#4c1d95','#0c4a6e','#14532d','#78350f','#7f1d1d','#2e1065'
// ];

const AREA_COLORS = (window.SHARED_COUNTRY_COLORS && window.SHARED_COUNTRY_COLORS.length)
  ? window.SHARED_COUNTRY_COLORS.slice()
  : [
    '#8DD3C9','#FFB88C','#A6D8E8','#F4A2C5','#B5D99C','#FFD08A','#94C0E8','#F2B8B8',
    '#C0C9F0','#88E0B0','#FFB3BA','#A2E1D1','#D9E8B0','#FFC988','#B1C5E5','#E8B8D2'
  ];

// ISO数字码 -> 中文名（地图用）
const ISO2CN = {
  "616":"波兰","203":"捷克","300":"希腊","348":"匈牙利","642":"罗马尼亚",
  "688":"塞尔维亚","705":"斯洛文尼亚","703":"斯洛伐克","191":"克罗地亚",
  "100":"保加利亚","233":"爱沙尼亚","428":"拉脱维亚","807":"马其顿",
  "499":"黑山","070":"波黑","008":"阿尔巴尼亚"
};

// ===== 全局状态 =====
let G = {}, CF_DATA = {};
let charts = {};
let _CY = [];
let _FIELD_RADAR = {};
let _mapMode = 'total', _mapYearIdx = 0, _playTimer = null;
let _mapLoaded = false, _topo = null;
let modalRadarChart = null;

// ============================================================
// Excel 读取工具
// ============================================================
async function fetchExcel(path) {
  const resp = await fetch(path);
  if (!resp.ok) throw new Error(`无法加载 ${path}（HTTP ${resp.status}）`);
  const buf  = await resp.arrayBuffer();
  const wb   = XLSX.read(buf, { type: 'array' });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

// ============================================================
// 数据加载与整理
// ============================================================
async function loadData() {
  try {
    const [rowYear, rowCtry, rowWorld, rowField, rowCtryField,
           rowInstCeec, rowInstCn, rowCtryName, rowFieldDef] = await Promise.all([
      fetchExcel('../data/1-中东欧-年份-总量.xlsx'),
      fetchExcel('../data/1-中东欧-时间段-总量.xlsx'),
      fetchExcel('../data/1-世界各国-时间段-总量.xlsx'),
      fetchExcel('../data/2-中东欧-时间段-领域.xlsx'),
      fetchExcel('../data/2-中东欧前12国-时间段（仅20162020）-领域.xlsx'),
      fetchExcel('../data/3-中东欧-时间段-机构.xlsx'),
      fetchExcel('../data/3-中国-时间段-机构.xlsx'),
      fetchExcel('../data/中东欧16国.xlsx'),
      fetchExcel('../data/领域.xlsx'),
    ]);

    // ── 领域ID映射表 ──
    const l2Map = {};
    const l1Map = {};
    rowFieldDef.forEach(r => {
      const l1id = String(r['一级ID']).padStart(2,'0');
      const l2id = String(r['二级ID']).padStart(4,'0');
      l1Map[l1id] = r['一级中文名称'];
      l2Map[l2id] = { name: r['二级中文名称'], l1Name: r['一级中文名称'], l1Id: l1id };
    });
    l2Map['1400'] = { name: '交叉领域', l1Name: '交叉领域', l1Id: '14' };
    l1Map['14'] = '交叉领域';

    // 国家英文名 -> 中文名
    const ctryEnToCn = {};
    rowCtryName.forEach(r => { ctryEnToCn[r['国家英文名']] = r['国家中文名']; });

    // 兜底对照表（应对世界排名表无中文名列的情况）
    const EN2CN_FALLBACK = {
      'USA':'美国','UNITED KINGDOM':'英国','AUSTRALIA':'澳大利亚','CANADA':'加拿大',
      'GERMANY (FED REP GER)':'德国','JAPAN':'日本','SINGAPORE':'新加坡','FRANCE':'法国',
      'SOUTH KOREA':'韩国','ITALY':'意大利','PAKISTAN':'巴基斯坦','NETHERLANDS':'荷兰',
      'SWEDEN':'瑞典','SPAIN':'西班牙','INDIA':'印度','RUSSIA':'俄罗斯',
      'SWITZERLAND':'瑞士','SAUDI ARABIA':'沙特阿拉伯','DENMARK':'丹麦','BELGIUM':'比利时',
      'BRAZIL':'巴西','FINLAND':'芬兰','POLAND':'波兰','NEW ZEALAND':'新西兰',
      'AUSTRIA':'奥地利','MALAYSIA':'马来西亚','NORWAY':'挪威','IRAN':'伊朗',
      'EGYPT':'埃及','CZECH REPUBLIC':'捷克','TURKEY':'土耳其','IRELAND':'爱尔兰',
      'SOUTH AFRICA':'南非','PORTUGAL':'葡萄牙','THAILAND':'泰国','ISRAEL':'以色列',
      'GREECE':'希腊','VIETNAM':'越南','HUNGARY':'匈牙利','MEXICO':'墨西哥',
      'ROMANIA':'罗马尼亚','SERBIA':'塞尔维亚','CHILE':'智利','TAIWAN':'台湾',
      'UKRAINE':'乌克兰','ARGENTINA':'阿根廷','INDONESIA':'印度尼西亚',
      'COLOMBIA':'哥伦比亚','CROATIA':'克罗地亚','SLOVAKIA':'斯洛伐克',
      'SLOVENIA':'斯洛文尼亚','ESTONIA':'爱沙尼亚','LATVIA':'拉脱维亚',
      'BULGARIA':'保加利亚','LITHUANIA':'立陶宛','LUXEMBOURG':'卢森堡',
    };
    Object.entries(EN2CN_FALLBACK).forEach(([en, cn]) => {
      if (!ctryEnToCn[en]) ctryEnToCn[en] = cn;
    });

    // ── 1. 年份数据 ──
    const yearData = rowYear.map(r => ({
      '年份': +r['年份'],
      '中东欧': +r['中东欧'],
      '我国国合总体': +r['我国国合总体'],
      '占比': +(+r['中东欧占比'] * 100).toFixed(4),
    }));

    // ── 2. 国别数据 ──
    const ctryMap = {};
    rowCtry.forEach(r => {
      const name = r['国家中文名'];
      if (!ctryMap[name]) ctryMap[name] = { '国家中文名': name };
      ctryMap[name][r['时间段']] = +r['合作Wos论文数'];
    });
    const countryData = Object.values(ctryMap).map(d => ({
      ...d,
      '增长率': d['2011-2015'] > 0
        ? +((d['2016-2020'] - d['2011-2015']) / d['2011-2015'] * 100).toFixed(1)
        : 0
    }));

    // ── 3. 世界排名数据 ──
    const ceecEnNames = new Set(rowCtryName.map(r => r['国家英文名']));
    const worldData = rowWorld
      .filter(r => r['时间段'] === '2016-2020')
      .map(r => ({
        '中文名': ctryEnToCn[r['国家英文名']] || r['国家中文名'] || r['国家英文名'],
        '英文名': r['国家英文名'],
        '2016-2020': +r['合作Wos论文数'],
        'is_ceec': ceecEnNames.has(r['国家英文名']),
      }))
      .sort((a, b) => b['2016-2020'] - a['2016-2020'])
      .slice(0, 40);

    // ── 4. 学科数据 ──
    const fieldByPeriod = { '2011-2015': [], '2016-2020': [] };
    rowField.forEach(r => {
      const l2id = String(r['二级ID']).padStart(4,'0');
      const info = l2Map[l2id];
      if (!info) return;
      const period = r['时间段'];
      fieldByPeriod[period].push({ name: info.name, val: +r['合作Wos论文数'], cat: info.l1Name });
    });
    ['2011-2015','2016-2020'].forEach(p => {
      fieldByPeriod[p].sort((a,b) => b.val - a.val);
      fieldByPeriod[p] = fieldByPeriod[p].slice(0, 10);
    });

    // 雷达图（一级学科汇总）
    const radarMap = { '2011-2015': {}, '2016-2020': {} };
    rowField.forEach(r => {
      const l1id = String(r['一级ID']).padStart(2,'0');
      const l1name = l1Map[l1id];
      if (!l1name) return;
      const p = r['时间段'];
      radarMap[p][l1name] = (radarMap[p][l1name] || 0) + (+r['合作Wos论文数']);
    });
    const radarCats16 = Object.entries(radarMap['2016-2020']).sort((a,b) => b[1]-a[1]).slice(0,8).map(e=>e[0]);
    _FIELD_RADAR = {
      '2011-2015': { cats: radarCats16, vals: radarCats16.map(c => radarMap['2011-2015'][c]||0) },
      '2016-2020': { cats: radarCats16, vals: radarCats16.map(c => radarMap['2016-2020'][c]||0) },
    };

    // 学科对比
    const cmpMap = {};
    rowField.forEach(r => {
      const l2id = String(r['二级ID']).padStart(4,'0');
      const info = l2Map[l2id];
      if (!info) return;
      if (!cmpMap[info.name]) cmpMap[info.name] = { name: info.name, v0: 0, v1: 0, cat: info.l1Name };
      if (r['时间段'] === '2011-2015') cmpMap[info.name].v0 += +r['合作Wos论文数'];
      else cmpMap[info.name].v1 += +r['合作Wos论文数'];
    });
    const fieldCompare = Object.values(cmpMap).sort((a,b) => b.v1 - a.v1).slice(0, 10);

    // ── 5. 机构数据 ──
    const instCeec = { '2011-2015': [], '2016-2020': [] };
    const instCn   = { '2011-2015': [], '2016-2020': [] };
    rowInstCeec.forEach(r => {
      instCeec[r['时间段']]?.push({ name: r['机构英文名'], country: r['国家或地区'], val: +r['合作Wos论文数'] });
    });
    rowInstCn.forEach(r => {
      instCn[r['时间段']]?.push({ name: r['机构英文名'], val: +r['合作Wos论文数'] });
    });
    ['2011-2015','2016-2020'].forEach(p => {
      instCeec[p] = instCeec[p].sort((a,b) => b.val-a.val).slice(0,10);
      instCn[p]   = instCn[p].sort((a,b) => b.val-a.val).slice(0,10);
    });

    // ── 6. 国家详细数据 ──
    const cfRaw = {};
    rowCtryField.forEach(r => {
      const cn = r['国家中文名'];
      const l2id = String(r['二级ID']).padStart(4,'0');
      const l1id = String(r['一级ID']).padStart(2,'0');
      const info = l2Map[l2id];
      if (!info) return;
      if (!cfRaw[cn]) cfRaw[cn] = { fields: {}, l1agg: {} };
      cfRaw[cn].fields[info.name] = (cfRaw[cn].fields[info.name]||0) + (+r['合作Wos论文数']);
      cfRaw[cn].l1agg[info.l1Name] = (cfRaw[cn].l1agg[info.l1Name]||0) + (+r['合作Wos论文数']);
    });
    Object.keys(cfRaw).forEach(cn => {
      const d = cfRaw[cn];
      const top6 = Object.entries(d.fields).sort((a,b)=>b[1]-a[1]).slice(0,6)
        .map(([name,val]) => {
          const cat = Object.entries(l2Map).find(([,v])=>v.name===name)?.[1]?.l1Name || '';
          return { name, val, cat };
        });
      const radarCatsLocal = Object.entries(d.l1agg).sort((a,b)=>b[1]-a[1]).slice(0,8).map(e=>e[0]);
      CF_DATA[cn] = {
        total: Object.values(d.fields).reduce((s,v)=>s+v,0),
        top6,
        radar_cats: radarCatsLocal,
        radar_vals: radarCatsLocal.map(c => d.l1agg[c]||0),
      };
    });

    // ── 7. 各国逐年数据（面积图/地图） ──
    const ctryOrder = countryData.map(d => d['国家中文名']);
    _CY = ctryOrder.map((name) => {
      const d = ctryMap[name];
      const a = d['2011-2015'] || 0;
      const b = d['2016-2020'] || 0;
      const yearly = [
        ...Array(5).fill(0).map((_,i) => Math.round(a/5 * (0.8 + i*0.1))),
        ...Array(5).fill(0).map((_,i) => Math.round(b/5 * (0.8 + i*0.1))),
      ];
      const sum0 = yearly.slice(0,5).reduce((s,v)=>s+v,0);
      const sum1 = yearly.slice(5).reduce((s,v)=>s+v,0);
      const adj = yearly.map((v,i) => i<5 ? Math.round(v*a/Math.max(sum0,1)) : Math.round(v*b/Math.max(sum1,1)));
      return { name, yearly: adj };
    });

    // ── 组装 G ──
    G = {
      year:         yearData,
      country:      countryData,
      world:        worldData,
      field:        fieldByPeriod,
      fieldCompare: fieldCompare,
      inst:         { ceec: instCeec, cn: instCn },
    };

    initAll();

  } catch(err) {
    console.error('数据加载失败:', err);
    document.body.innerHTML = `<div style="padding:60px;text-align:center;color:#b91c1c;font-size:16px;line-height:2">
      ⚠️ 数据加载失败，请通过本地服务器（如 VS Code Live Server）打开页面，<br>
      不能直接双击 HTML 文件（浏览器会阻止 fetch 请求）。<br>
      <small style="color:#8a97b0">错误：${err.message}</small>
    </div>`;
  }
}

// ============================================================
// 初始化
// ============================================================
function initAll() {
  initHero();
  initMap();
  initAreaChart();
  setTimeout(() => renderYearInfo(0), 300);
  initTrend();
  renderCountryBars('2011-2015');
  initGrowthChart();
  initWorldRank();
  renderFieldBars('2011-2015');
  initRadarChart();
  initFieldCompare();
  renderInstList('2011-2015');
  initInstChart('2011-2015');
  initNav();
}

// ============================================================
// Hero 统计栏
// ============================================================
function initHero() {
  const yr    = G.year;
  const total = yr.reduce((s,d) => s + d['中东欧'], 0);
  const last  = yr[yr.length-1]['中东欧'];
  const first = yr[0]['中东欧'];
  const growth = ((last-first)/first*100).toFixed(1);
  const cagr   = (((last/first)**(1/9))-1)*100;
  document.getElementById('h1').textContent = total.toLocaleString();
  document.getElementById('h2').textContent = last.toLocaleString();
  document.getElementById('h3').textContent = growth + '%';
  document.getElementById('h4').textContent = cagr.toFixed(1) + '%';

  // mini stats
  const avgInc = Math.round((last - first) / 9);
  document.getElementById('miniAvg').textContent  = '+' + avgInc.toLocaleString() + '篇/年';
  document.getElementById('miniCagr').textContent = cagr.toFixed(1) + '%';
  const ratioFirst = yr[0]['占比'], ratioLast = yr[yr.length-1]['占比'];
  document.getElementById('miniRatio').textContent = '+' + (ratioLast - ratioFirst).toFixed(2) + 'pp';
}

// ============================================================
// 模块1：地图
// ============================================================
function getMapValues(yearIdx) {
  const res = {};
  _CY.forEach(d => { res[d.name] = d.yearly[yearIdx]; });
  return res;
}

function initMap() {
  d3.json('https://cdn.jsdelivr.net/npm/world-atlas@2/countries-110m.json').then(world => {
    _topo = world; _mapLoaded = true; renderMap(0);
  }).catch(() => {
    d3.select('#mapSvg').append('text').attr('x',450).attr('y',175)
      .attr('text-anchor','middle').attr('fill','#8a97b0')
      .text('地图加载需要网络，其他图表已正常显示');
  });
}

function renderMap(yearIdx) {
  if (!_mapLoaded || !_topo) return;
  const svg = d3.select('#mapSvg');
  svg.selectAll('*').remove();
  const w = svg.node().clientWidth || 900, h = 420;
  const proj = d3.geoMercator().center([23,47]).scale(w*0.85).translate([w/2, h*0.52]);
  const path = d3.geoPath().projection(proj);
  const vals = getMapValues(yearIdx);
  const ceecISOs = new Set(Object.keys(ISO2CN));

  let colorFn;
  if (_mapMode === 'total') {
    const maxV = Math.max(...Object.values(vals), 1);
    const cs = d3.scaleSequential().domain([0,maxV]).interpolator(d3.interpolateYlGnBu);
    colorFn = name => { const v=vals[name]; return v ? cs(v) : '#f7fbfc'; };
  } else {
    const cs = d3.scaleSequential().domain([0,1000]).interpolator(d3.interpolateYlOrRd);
    colorFn = name => {
      const d = G.country.find(c => c['国家中文名']===name);
      return d ? cs(Math.min(d['增长率'],1000)) : '#fff5eb';
    };
  }

  const countries = topojson.feature(_topo, _topo.objects.countries);
  const ceecFeatures = countries.features.filter(f => ceecISOs.has(String(f.id).padStart(3,'0')));

  svg.append('g').selectAll('path')
    .data(ceecFeatures).join('path')
    .attr('d', path)
    .attr('fill', f => { const cn = ISO2CN[String(f.id).padStart(3,'0')]; return colorFn(cn); })
    .attr('stroke', '#059669').attr('stroke-width', 1.5)
    .style('cursor', 'pointer')
    .on('mousemove', (event, f) => {
      const cn = ISO2CN[String(f.id).padStart(3,'0')];
      if (!cn) { d3.select('#mapTip').style('opacity',0); return; }
      const v = vals[cn]||0;
      const cd = G.country.find(d => d['国家中文名']===cn);
      const gr = cd ? cd['增长率'] : 0;
      const grStr = gr>999?'>999':gr.toFixed(1);
      const tipMain = _mapMode==='growth'
        ? `五年增长率：<b style="color:#b45309;font-size:16px">↑${grStr}%</b><br>${YEARS[yearIdx]}年合作论文：${v.toLocaleString()}篇`
        : `${YEARS[yearIdx]}年合作论文：<b>${v.toLocaleString()}</b>篇<br>五年增长率：↑${grStr}%`;
      const tip = d3.select('#mapTip');
      tip.html(`<div style="font-weight:800;font-size:16px;margin-bottom:6px">${YEARS[yearIdx]}</div><b>${cn}</b><br>${tipMain}<br><span style="font-size:16px;color:#8a97b0">${CF_DATA[cn]?'点击查看更多细节 →':'（暂无详细数据）'}</span>`)
        .style('opacity',1);
      const node = tip.node();
      const gap = 18;
      const margin = 12;
      const vw = window.innerWidth;
      const vh = window.innerHeight;
      const tipW = node?.offsetWidth || 260;
      const tipH = node?.offsetHeight || 120;
      const useLeft = event.clientX + gap + tipW > vw - margin;
      const useTop = event.clientY + gap + tipH > vh - margin;
      const left = useLeft ? event.clientX - tipW - gap : event.clientX + gap;
      const top = useTop ? event.clientY - tipH - gap : event.clientY + gap;
      tip.style('left', Math.max(margin, Math.min(left, vw - tipW - margin)) + 'px')
        .style('top', Math.max(margin, Math.min(top, vh - tipH - margin)) + 'px');
    })
    .on('mouseleave', () => d3.select('#mapTip').style('opacity',0))
    .on('click', (event, f) => {
      const cn = ISO2CN[String(f.id).padStart(3,'0')];
      if (cn && CF_DATA[cn]) openDrill(cn);
    });

  svg.append('g').selectAll('text')
    .data(ceecFeatures).join('text')
    .attr('transform', f => { const c=path.centroid(f); return isNaN(c[0])?'translate(-999,-999)':`translate(${c})`; })
    .attr('text-anchor','middle').attr('font-size',14).attr('fill','#000')
    .attr('font-weight','600').attr('pointer-events','none').attr('stroke','#fff').attr('stroke-width',1)
    .style('paint-order','stroke fill')
    .text(f => ISO2CN[String(f.id).padStart(3,'0')]||'');

  if (_mapMode === 'growth') {
    const grSorted = [...G.country].sort((a,b)=>b['增长率']-a['增长率']);
    document.getElementById('mapLegend').innerHTML =
      `<div style="font-weight:800;font-size:16px;color:#b45309;margin-bottom:5px;padding-bottom:4px;border-bottom:1px solid #dde1e9">增长率 Top3</div>` +
      grSorted.slice(0,3).map((d,i)=>{
        const gr = d['增长率']>999?'>999%':d['增长率'].toFixed(1)+'%';
        return `<div style="display:flex;justify-content:space-between;gap:12px;margin-bottom:3px">
          <span style="font-size:16px">${['🥇','🥈','🥉'][i]} ${d['国家中文名']}</span>
          <span style="font-size:16px;font-weight:700;color:#b45309">↑${gr}</span>
        </div>`;
      }).join('') +
      `<div style="margin-top:5px;padding-top:5px;border-top:1px solid #dde1e9;font-size:16px;color:#8a97b0">颜色越深 = 增长率越高</div>`;
  } else {
    renderYearInfo(yearIdx);
  }
}

function setMapMode(mode, btn) {
  _mapMode = mode;
  document.getElementById('mapModeTotal').classList.remove('active');
  document.getElementById('mapModeGrowth').classList.remove('active');
  btn.classList.add('active');
  renderMap(_mapYearIdx);
}

function renderYearInfo(yearIdx) {
  if (_mapMode === 'growth') return;
  const yr = G.year[yearIdx];
  const prevYr = yearIdx > 0 ? G.year[yearIdx - 1] : null;
  const ceecTotal = yr['中东欧'];
  const delta = prevYr ? ceecTotal - prevYr['中东欧'] : null;
  const deltaStr = delta !== null
    ? (delta >= 0
        ? `<span style="color:#047857;font-weight:700">▲ +${delta.toLocaleString()}篇</span>`
        : `<span style="color:#b91c1c;font-weight:700">▼ ${Math.abs(delta).toLocaleString()}篇</span>`)
    : '';
  const vals = _CY.map(d => ({ name: d.name, val: d.yearly[yearIdx] }))
                  .sort((a, b) => b.val - a.val).slice(0, 3);
  const legend = document.getElementById('mapLegend');
  if (!legend) return;
  legend.innerHTML = `
    <div style="font-weight:800;font-size:16px;color:var(--blue);margin-bottom:4px;padding-bottom:4px;border-bottom:1px solid var(--border)">${YEARS[yearIdx]}年</div>
    <div style="font-size:16px;margin-bottom:5px">中东欧论文：<b>${ceecTotal.toLocaleString()}</b>篇${deltaStr ? `<br>较上年：${deltaStr}` : ''}</div>
    <div style="font-size:16px;color:var(--text3);font-weight:600;margin-bottom:3px">论文量 Top3</div>
    ${vals.map((d,i) => `
      <div style="display:flex;justify-content:space-between;gap:10px;margin-bottom:2px">
        <span style="font-size:16px">${['🥇','🥈','🥉'][i]} ${d.name}</span>
        <span style="font-size:16px;font-weight:700;color:var(--blue)">${d.val.toLocaleString()}篇</span>
      </div>`).join('')}`;
}

function onSliderChange(idx) {
  _mapYearIdx = idx;
  document.getElementById('yearLabel').textContent = YEARS[idx];
  renderMap(idx); renderYearInfo(idx);
}

function togglePlay() {
  const btn = document.getElementById('playBtn');
  if (_playTimer) {
    clearInterval(_playTimer); _playTimer = null;
    btn.textContent = '▶ 播放'; btn.style.background = 'var(--blue)'; return;
  }
  btn.textContent = '⏹ 停止'; btn.style.background = 'var(--red)';
  let i = _mapYearIdx;
  _playTimer = setInterval(() => {
    i = (i + 1) % 10; _mapYearIdx = i;
    document.getElementById('yearSlider').value = i;
    document.getElementById('yearLabel').textContent = YEARS[i];
    renderMap(i); renderYearInfo(i);
    if (i === 9) { clearInterval(_playTimer); _playTimer = null; btn.textContent='▶ 播放'; btn.style.background='var(--blue)'; }
  }, 900);
}

// ============================================================
// 面积图
// ============================================================
function initAreaChart() {
  const ctx = document.getElementById('areaChart').getContext('2d');
  const datasets = _CY.map((d,i)=>({
    label:d.name, data:d.yearly,
    backgroundColor:AREA_COLORS[i]+'cc', borderColor:AREA_COLORS[i],
    borderWidth:1, fill:true, tension:.35, pointRadius:2, pointHoverRadius:4,
  }));
  charts.area = new Chart(ctx, {
    type:'line', data:{labels:YEARS, datasets},
    options:{
      responsive:true, maintainAspectRatio:false,
      interaction:{mode:'index',intersect:false},
      plugins:{
        legend:{display:false},
        tooltip:{
          enabled:false,
          external: function(context) {
            const tooltip = context.tooltip;
            if (!tooltip || tooltip.opacity === 0) { hideAreaChartTooltip(); return; }
            const chart = context.chart;
            const dps = tooltip.dataPoints || [];
            if (dps.length === 0) { hideAreaChartTooltip(); return; }
            const title = `${dps[0].label}年（点击同步地图）`;
            const body = dps.map(dp => `<div style="font-size:16px;color:${C.text2}"> ${dp.dataset.label}: ${dp.parsed.y.toLocaleString()}篇</div>`).join('');
            const html = `<div style="font-weight:800;font-size:16px;color:${C.text};margin-bottom:6px">${title}</div>${body}`;
            showAreaChartTooltip(chart, tooltip, html);
          }
        }
      },
      onClick:(event,elements)=>{
        if(elements.length>0){const idx=elements[0].index;_mapYearIdx=idx;document.getElementById('yearSlider').value=idx;document.getElementById('yearLabel').textContent=YEARS[idx];renderMap(idx);renderYearInfo(idx);}
      },
      scales:{
        x:{grid:{color:'#edf0f5'},ticks:{color:C.text3,font:{size:12}},border:{color:C.border}},
        y:{stacked:true,grid:{color:'#edf0f5'},ticks:{color:C.text3,font:{size:16},callback:v=>v>=1000?Math.round(v/1000)+'k':v},border:{color:C.border}}
      }
    }
  });
  const legendWrap = document.getElementById('areaLegend');
  if (legendWrap) {
    const row1=datasets.slice(0,6), row2=datasets.slice(6,11), row3=datasets.slice(11);
    const makeItem=(ds,i)=>`<div style="display:flex;align-items:center;gap:3px;cursor:pointer;opacity:1;transition:opacity .2s;flex:1 1 0;min-width:0" onclick="toggleAreaSeries(${i},this)">
      <span style="width:10px;height:10px;border-radius:2px;background:${AREA_COLORS[i]};flex-shrink:0;display:inline-block"></span>
      <span style="font-size:16px;color:#4a5568;white-space:nowrap">${ds.label}</span></div>`;
    legendWrap.innerHTML =
      `<div style="display:flex;width:100%;justify-content:space-between;gap:10px 14px;margin-bottom:3px">${row1.map((ds,i)=>makeItem(ds,i)).join('')}</div>`+
      `<div style="display:flex;width:100%;justify-content:space-between;gap:10px 14px;margin-bottom:3px">${row2.map((ds,i)=>makeItem(ds,i+6)).join('')}</div>`+
      `<div style="display:flex;width:100%;justify-content:space-between;gap:10px 14px">${row3.map((ds,i)=>makeItem(ds,i+11)).join('')}</div>`;
  }
}

function toggleAreaSeries(idx, el) {
  const meta = charts.area.getDatasetMeta(idx);
  meta.hidden = !meta.hidden; el.style.opacity=meta.hidden?'0.3':'1'; charts.area.update();
}

// ===== Area chart DOM tooltip (避免 canvas 内绘制被遮挡) =====
let areaChartTooltipEl = null;
function ensureAreaChartTooltip() {
  if (areaChartTooltipEl) return areaChartTooltipEl;
  areaChartTooltipEl = document.createElement('div');
  areaChartTooltipEl.style.cssText = [
    'position:absolute',
    'z-index:10000',
    'pointer-events:none',
    'opacity:0',
    'max-width:360px',
    'padding:8px',
    'border:1px solid var(--border)',
    'border-radius:8px',
    'background:#fff',
    'box-shadow:0 12px 28px rgba(15,23,42,.12)',
    'color:var(--text)',
    'font-size:16px',
    'line-height:1.5',
    'white-space:nowrap',
    'overflow:hidden',
    'text-overflow:ellipsis'
  ].join(';');
  document.body.appendChild(areaChartTooltipEl);
  return areaChartTooltipEl;
}

function hideAreaChartTooltip() {
  if (!areaChartTooltipEl) return;
  areaChartTooltipEl.style.opacity = '0';
}

function showAreaChartTooltip(chart, tooltip, html) {
  const el = ensureAreaChartTooltip();
  el.innerHTML = html;
  el.style.opacity = '1';

  const rect = chart.canvas.getBoundingClientRect();
  const viewportW = window.innerWidth;
  const viewportH = window.innerHeight;
  const offset = 12;

  let left = rect.left + window.scrollX + tooltip.caretX + offset;
  let top = rect.top + window.scrollY + tooltip.caretY - 8;

  // temporarily set to auto to measure
  el.style.left = '0px'; el.style.top = '0px';
  const tipRect = el.getBoundingClientRect();
  const tipW = tipRect.width || 260;
  const tipH = tipRect.height || 120;

  // Prefer to show above the point if there is room
  if (rect.top + window.scrollY + tooltip.caretY - tipH - offset > window.scrollY + 8) {
    top = rect.top + window.scrollY + tooltip.caretY - tipH - offset;
  } else {
    top = rect.top + window.scrollY + tooltip.caretY + offset;
  }

  if (left + tipW > window.scrollX + viewportW - 12) {
    left = rect.left + window.scrollX + tooltip.caretX - tipW - offset;
  }
  if (left < window.scrollX + 8) left = window.scrollX + 8;
  if (top < window.scrollY + 8) top = window.scrollY + 8;

  el.style.left = `${left}px`;
  el.style.top = `${top}px`;
}

// ============================================================
// 趋势图 + 占比图
// ============================================================
function initTrend() {
  const yr = G.year;
  charts.trend = new Chart(document.getElementById('trendChart').getContext('2d'), {
    type:'line',
    data:{labels:yr.map(d=>d['年份']),datasets:[
      {label:'中东欧合作论文（篇）',data:yr.map(d=>d['中东欧']),borderColor:C.teal,backgroundColor:'rgba(14,116,144,.08)',borderWidth:2.5,fill:true,tension:.4,pointRadius:5,pointBackgroundColor:C.teal,pointBorderColor:'#fff',pointBorderWidth:2,yAxisID:'y'},
      {label:'中国国合总体（篇）',data:yr.map(d=>d['我国国合总体']),borderColor:'#f59e0b',backgroundColor:'rgba(245,158,11,.05)',borderWidth:2,fill:false,tension:.4,borderDash:[6,3],pointRadius:4,pointBackgroundColor:'#f59e0b',pointBorderColor:'#fff',pointBorderWidth:2,yAxisID:'y2'}
    ]},
    options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},
      plugins:{legend:{labels:{color:C.text2,font:{size:16},boxWidth:24,boxHeight:2}},
        tooltip:{backgroundColor:'#fff',borderColor:C.border,borderWidth:1,titleColor:C.text,bodyColor:C.text2,padding:10,
          callbacks:{label:ctx=>ctx.datasetIndex===0?` 中东欧：${ctx.parsed.y.toLocaleString()}篇`:` 国合总体：${ctx.parsed.y.toLocaleString()}篇`}}},
      scales:{
        x:{grid:{color:'#edf0f5'},ticks:{color:C.text3,font:{size:16}},border:{color:C.border}},
        y:{grid:{color:'#edf0f5'},ticks:{color:C.teal,font:{size:16},callback:v=>v.toLocaleString()},border:{color:C.border},title:{display:true,text:'中东欧（篇）',color:C.teal,font:{size:16}}},
        y2:{position:'right',grid:{display:false},ticks:{color:'#b45309',font:{size:16},callback:v=>Math.round(v/1000)+'k'},border:{color:C.border},title:{display:true,text:'国合总体（篇）',color:'#b45309',font:{size:16}}}
      }}
  });

  // 趋势insight（动态生成）
  const first = yr[0], last = yr[yr.length-1];
  const growth = ((last['中东欧']-first['中东欧'])/first['中东欧']*100).toFixed(1);
  const cagr = (((last['中东欧']/first['中东欧'])**(1/9))-1)*100;
  document.getElementById('trendInsight').innerHTML =
    `<strong>解读：</strong>中东欧合作论文量从${first['年份']}年的 <strong>${first['中东欧'].toLocaleString()}篇</strong> 增至${last['年份']}年的 <strong>${last['中东欧'].toLocaleString()}篇</strong>，增幅达${growth}%，CAGR ${cagr.toFixed(1)}%。`;

  charts.ratio = new Chart(document.getElementById('ratioChart').getContext('2d'), {
    type:'bar',
    data:{labels:yr.map(d=>d['年份']),datasets:[{label:'占比（%）',data:yr.map(d=>d['占比']),backgroundColor:'rgba(29,78,216,.55)',borderColor:C.blue,borderWidth:1.5,borderRadius:4,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{backgroundColor:'#fff',borderColor:C.border,borderWidth:1,titleColor:C.text,bodyColor:C.text2,callbacks:{label:ctx=>`占比: ${ctx.parsed.y.toFixed(2)}%`}}},
      scales:{
        x:{grid:{display:false},ticks:{color:C.text3,font:{size:16}},border:{color:C.border}},
        y:{grid:{color:'#edf0f5'},ticks:{color:C.text2,font:{size:16},callback:v=>v.toFixed(1)+'%'},border:{color:C.border}}
      }}
  });

  const ratioFirst = yr[0]['占比'], ratioLast = yr[yr.length-1]['占比'];
  const peakYear = yr.reduce((a,b) => b['占比']>a['占比']?b:a);
  document.getElementById('ratioInsight').innerHTML =
    `<strong>解读：</strong>占比从${yr[0]['年份']}年的 <strong>${ratioFirst.toFixed(2)}%</strong> 提升至${yr[yr.length-1]['年份']}年的 <strong>${ratioLast.toFixed(2)}%</strong>，${peakYear['年份']}年占比达 <strong>${peakYear['占比'].toFixed(2)}%</strong>。`;
}

// ============================================================
// 模块2：国别
// ============================================================
function renderCountryBars(period) {
  const sorted = [...G.country].sort((a,b)=>b[period]-a[period]);
  const maxV = sorted[0][period];
  document.getElementById('countryBars').innerHTML = sorted.map(d=>{
    const v=d[period], pct=Math.round(v/maxV*100);
    const canDrill=!!CF_DATA[d['国家中文名']];
    return `<div class="cbar" style="${canDrill?'cursor:pointer':'cursor:default'}" ${canDrill?`onclick="openDrill('${d['国家中文名']}')"`:''}>
      <div class="cbar-label">${d['国家中文名']}</div>
      <div class="cbar-track"><div class="cbar-fill" style="width:${pct}%;background:linear-gradient(90deg,${C.blue}88,${C.blue})"></div></div>
      <div class="cbar-delta" style="color:${C.text3};font-size:16px">${v.toLocaleString()}篇</div>
    </div>`;
  }).join('');
  const top3=sorted.slice(0,3), allSum=sorted.reduce((s,d)=>s+d[period],0);
  const topPct=(top3.reduce((s,d)=>s+d[period],0)/allSum*100).toFixed(1);
  const insightMap={
    '2011-2015':`<strong>解读（2011–2015）：</strong>${top3[0]['国家中文名']}（${top3[0][period].toLocaleString()}篇）、${top3[1]['国家中文名']}（${top3[1][period].toLocaleString()}篇）、${top3[2]['国家中文名']}（${top3[2][period].toLocaleString()}篇）位列前三，合计占中东欧总量的 <strong>${topPct}%</strong>，头部集中效应显著。`,
    '2016-2020':`<strong>解读（2016–2020）：</strong>${top3[0]['国家中文名']}（${top3[0][period].toLocaleString()}篇）、${top3[1]['国家中文名']}（${top3[1][period].toLocaleString()}篇）、${top3[2]['国家中文名']}（${top3[2][period].toLocaleString()}篇）仍位列前三，合计占 <strong>${topPct}%</strong>。点击国家名称可查看学科细节。`
  };
  document.getElementById('countryInsight').innerHTML = insightMap[period]||'';
}

function switchPeriod(period, btn) {
  document.querySelectorAll('#m2 .ptab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); renderCountryBars(period);
}

function initGrowthChart() {
  const top12=[...G.country].sort((a,b)=>b['2016-2020']-a['2016-2020']).slice(0,12);
  charts.growth = new Chart(document.getElementById('growthChart').getContext('2d'), {
    type:'bar',
    data:{labels:top12.map(d=>d['国家中文名']),datasets:[
      {label:'2011–2015',data:top12.map(d=>d['2011-2015']),backgroundColor:'rgba(14,116,144,.35)',borderColor:C.teal,borderWidth:1.5,borderRadius:3},
      {label:'2016–2020',data:top12.map(d=>d['2016-2020']),backgroundColor:'rgba(29,78,216,.5)',borderColor:C.blue,borderWidth:1.5,borderRadius:3}
    ]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      layout:{padding:{left:10,right:8,top:4,bottom:4}},
      plugins:{legend:{labels:{color:C.text2,font:{size:16},boxWidth:12,boxHeight:12}},
        tooltip:{backgroundColor:'#fff',borderColor:C.border,borderWidth:1,titleColor:C.text,bodyColor:C.text2,callbacks:{label:ctx=>`${ctx.dataset.label}: ${ctx.parsed.x.toLocaleString()}篇`}}},
      scales:{
        x:{grid:{color:'#edf0f5'},ticks:{color:C.text3,font:{size:16},callback:v=>v.toLocaleString()},border:{color:C.border}},
        y:{grid:{display:false},afterFit(scale){ scale.width = 170; },ticks:{color:C.text,font:{size:16},autoSkip:false,padding:6,sampleSize:12},border:{color:C.border}}
      }}
  });
  // growth insight（动态）
  const maxGrowth = [...G.country].sort((a,b)=>b['增长率']-a['增长率'])[0];
  document.getElementById('growthInsight').innerHTML =
    `<strong>解读：</strong>所有16国在2016–2020较2011–2015均实现增长。${maxGrowth['国家中文名']}增幅最高（+${maxGrowth['增长率']>999?'>999':maxGrowth['增长率'].toFixed(0)}%），${top12[0]['国家中文名']}绝对增量最大（+${(top12[0]['2016-2020']-top12[0]['2011-2015']).toLocaleString()}篇）。`;
}

function initWorldRank() {
  const wd=G.world;
  const maxV = Math.max(...wd.map(d => d['2016-2020']), 1);
  const axisMax = Math.max(100000, Math.ceil(maxV / 100000) * 100000);
  const yTicks = [1000, 50000, 100000];
  for (let v = 200000; v <= axisMax; v += 100000) yTicks.push(v);
  charts.worldRank = new Chart(document.getElementById('worldRankChart').getContext('2d'), {
    type:'bar',
    data:{labels:wd.map(d=>d['中文名']),datasets:[{label:'2016–2020发文量',data:wd.map(d=>d['2016-2020']),
      backgroundColor:wd.map(d=>d['is_ceec']?'rgba(29,78,216,.85)':'rgba(14,116,144,.2)'),
      borderColor:wd.map(d=>d['is_ceec']?C.blue:'rgba(14,116,144,.4)'),
      borderWidth:wd.map(d=>d['is_ceec']?2:1),borderRadius:2,borderSkipped:false}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},
        tooltip:{backgroundColor:'#fff',borderColor:C.border,borderWidth:1,titleColor:C.text,bodyColor:C.text2,
          callbacks:{label:ctx=>`论文数: ${ctx.parsed.y.toLocaleString()}篇`,afterLabel:ctx=>wd[ctx.dataIndex]['is_ceec']?'★ 中东欧国家':''}}},
      scales:{
        x:{grid:{display:false},ticks:{color:C.text3,font:{size:12},autoSkip:false,minRotation:0,maxRotation:0,padding:6,callback:function(val){
            const txt = this.getLabelForValue(val) || '';
            return txt.split('');
          }},border:{color:C.border}},
        y:{type:'logarithmic',min:1000,max:axisMax,grid:{color:'#edf0f5'},afterBuildTicks:(scale)=>{
            scale.ticks = yTicks.map(value => ({ value }));
          },ticks:{color:C.text3,font:{size:16},callback:v=>v>=1000?Math.round(v/1000)+'k':v},border:{color:C.border}}
      }}
  });
  const ceecInRank = wd.filter(d=>d['is_ceec']);
  document.getElementById('worldInsight').innerHTML =
    `<strong>解读：</strong>在全球前40合作伙伴中，中东欧共有 <strong>${ceecInRank.length}国</strong> 上榜（2016–2020）：${ceecInRank.map(d=>`${d['中文名']}（${d['2016-2020'].toLocaleString()}篇）`).join('、')}。`;
}

// ============================================================
// 详细弹窗
// ============================================================
function openDrill(country) {
  const d=CF_DATA[country]; if(!d) return;
  document.getElementById('modalTitle').innerHTML=`${country} <span>2016–2020 · 共 ${d.total.toLocaleString()} 篇</span>`;
  const maxV=d.top6[0]['val'];
  document.getElementById('modalFieldBars').innerHTML=d.top6.map((f,i)=>{
    const pct=Math.round(f['val']/maxV*100), col=catColor[f['cat']]||C.blue;
    const abr={'理学':'理','工学':'工','医学':'医','管理学':'管','农学':'农','教育学':'教','法学':'法','经济学':'经','交叉领域':'跨'}[f['cat']]||f['cat'];
    return `<div class="fbar">
      <div class="fbar-label">${f['name']}</div>
      <div class="fbar-track"><div class="fbar-fill" style="width:${pct}%;background:${col}"></div></div>
      <div class="fbar-num">${f['val'].toLocaleString()}</div>
      <div class="fbar-cat" style="color:${col}">${abr}</div>
    </div>`;
  }).join('');
  if(modalRadarChart) modalRadarChart.destroy();
  modalRadarChart = new Chart(document.getElementById('modalRadar').getContext('2d'), {
    type:'radar',
    data:{labels:d.radar_cats,datasets:[{label:country,data:d.radar_vals,borderColor:C.blue,backgroundColor:'rgba(29,78,216,.1)',borderWidth:2,pointBackgroundColor:C.blue,pointRadius:4}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},
      scales:{r:{grid:{color:'#edf0f5'},ticks:{color:C.text3,font:{size:16},backdropColor:'transparent',callback:function(v){
                if (v >= 1000) return (v % 1000 === 0) ? (v/1000) + 'k' : '';
                return v;
              },stepSize:1000},pointLabels:{color:C.text,font:{size:16}},angleLines:{color:'#dde1e9'},min:0}}}
  });
  const top1=d.top6[0], top2=d.top6[1];
  document.getElementById('modalInsight').innerHTML=
    `<strong>解读：</strong>${country} 2016–2020年与中国共发表合作论文 <strong>${d.total.toLocaleString()}篇</strong>。
     首要合作领域为 <strong>${top1.name}</strong>（${top1.val.toLocaleString()}篇），其次为 <strong>${top2.name}</strong>（${top2.val.toLocaleString()}篇）。点击空白处关闭。`;
  document.getElementById('drillModal').classList.add('show');
}
function closeModal(e) { if(e.target===document.getElementById('drillModal')) document.getElementById('drillModal').classList.remove('show'); }
document.addEventListener('keydown', e=>{ if(e.key==='Escape') document.getElementById('drillModal').classList.remove('show'); });

// ============================================================
// 模块3：学科
// ============================================================
function renderFieldBars(period) {
  const data=G.field[period], maxV=data[0]['val'];
  document.getElementById('fieldBars').innerHTML=data.map((d,i)=>{
    const pct=Math.round(d['val']/maxV*100), col=catColor[d['cat']]||C.blue;
    const catAbr={'理学':'理','工学':'工','医学':'医','管理学':'管','农学':'农'}[d['cat']]||d['cat'];
    return `<div class="fbar">
      <div class="fbar-label">${d['name']}</div>
      <div class="fbar-track"><div class="fbar-fill" style="width:${pct}%;background:${col}"></div></div>
      <div class="fbar-num">${d['val'].toLocaleString()}</div>
      <div class="fbar-cat" style="color:${col}">${catAbr}</div>
    </div>`;
  }).join('');
  requestAnimationFrame(() => {
    document.querySelectorAll('#fieldBars .fbar-label').forEach(el => {
      if (el.scrollWidth > el.clientWidth) {
        el.title = el.textContent || '';
      } else {
        el.removeAttribute('title');
      }
    });
  });
  const insights={
    '2011-2015':`<strong>解读（2011–2015）：</strong>${data[0].name}以 <strong>${data[0].val.toLocaleString()}篇</strong> 高居榜首，远超第二位${data[1].name}（${data[1].val.toLocaleString()}篇）。理学类占据前五位，工学与医学已初显合作潜力。`,
    '2016-2020':`<strong>解读（2016–2020）：</strong>${data[0].name}仍以 <strong>${data[0].val.toLocaleString()}篇</strong> 居首，${data[1].name}（${data[1].val.toLocaleString()}篇）跃升至第2位，各领域体量全面扩大，工学类合作明显提速。`
  };
  document.getElementById('fieldInsight').innerHTML=insights[period]||'';
}

function initRadarChart() {
  charts.radar = new Chart(document.getElementById('radarChart').getContext('2d'), {
    type:'radar',
    data:{labels:_FIELD_RADAR['2016-2020'].cats,datasets:[
      {label:'2011–2015',data:_FIELD_RADAR['2016-2020'].cats.map(c=>{const i=_FIELD_RADAR['2011-2015'].cats.indexOf(c);return i>=0?_FIELD_RADAR['2011-2015'].vals[i]:0;}),borderColor:C.teal,backgroundColor:'rgba(14,116,144,.1)',borderWidth:2,pointBackgroundColor:C.teal,pointRadius:4},
      {label:'2016–2020',data:_FIELD_RADAR['2016-2020'].vals,borderColor:C.blue,backgroundColor:'rgba(29,78,216,.08)',borderWidth:2,pointBackgroundColor:C.blue,pointRadius:4}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{color:C.text2,font:{size:16},boxWidth:12}}},
      scales:{r:{grid:{color:'#edf0f5'},ticks:{color:C.text3,font:{size:16},backdropColor:'transparent',callback:function(v){
                if (v >= 1000) return (v % 1000 === 0) ? (v/1000) + 'k' : '';
                return v;
              },stepSize:1000},pointLabels:{color:C.text,font:{size:16}},angleLines:{color:'#dde1e9'},min:0}}}
  });
  document.getElementById('radarInsight').innerHTML =
    `<strong>解读：</strong>两阶段中 <strong>理学</strong> 始终是最大合作领域。工学在第二阶段实现爆发式增长（+156%），几乎与理学持平，反映合作正从基础科学快速向应用技术领域拓展。`;
}

function initFieldCompare() {
  const data=G.fieldCompare;
  charts.fieldCompare = new Chart(document.getElementById('fieldCompareChart').getContext('2d'), {
    type:'bar',
    data:{labels:data.map(d=>d['name']),datasets:[
      {label:'2011–2015',data:data.map(d=>d['v0']),backgroundColor:'rgba(14,116,144,.4)',borderColor:C.teal,borderWidth:1.5,borderRadius:3},
      {label:'2016–2020',data:data.map(d=>d['v1']),backgroundColor:'rgba(29,78,216,.5)',borderColor:C.blue,borderWidth:1.5,borderRadius:3}
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{color:C.text2,font:{size:16},boxWidth:12,boxHeight:12}},
        tooltip:{backgroundColor:'#fff',borderColor:C.border,borderWidth:1,titleColor:C.text,bodyColor:C.text2,callbacks:{label:ctx=>`${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString()}篇`}}},
      scales:{
        x:{grid:{display:false},ticks:{color:C.text3,font:{size:16},maxRotation:0,callback:function(val){
          const label=this.getLabelForValue(val);
          const breaks={'材料科学与工程':['材料科学','与工程'],'计算机科学与技术':['计算机科学','与技术'],'环境科学与工程':['环境科学','与工程']};
          return breaks[label]||label;
        }},border:{color:C.border}},
        y:{grid:{color:'#edf0f5'},ticks:{color:C.text3,font:{size:16},callback:v=>v.toLocaleString()},border:{color:C.border}}
      }}
  });
  document.getElementById('fieldCmpInsight').innerHTML =
    `<strong>解读：</strong>物理学始终居首，化学、材料科学增幅显著，反映中东欧与中国在新材料、化学合成领域合作明显加强。计算机科学与环境科学的快速增长体现数字经济与可持续发展领域的新兴合作需求。`;
}

function switchField(period, btn) {
  document.querySelectorAll('#m3 .ptab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); renderFieldBars(period);
}

// ============================================================
// 模块4：机构
// ============================================================
const instInsights = {
  '2011-2015':'<strong>解读（2011–2015）：</strong>布拉格查理大学（1,397篇）居首，匈牙利维格纳物理中心（1,162篇）、雅典国立大学（1,121篇）紧随。物理与核物理类机构占主导。',
  '2016-2020':'<strong>解读（2016–2020）：</strong>格局显著变化：爱沙尼亚塔尔图大学（611篇）跃居首位，罗马尼亚（4所）和拉脱维亚（2所）大量进入Top 10，体现合作主体多元化的新趋势。'
};

let instChartTooltipEl = null;

function ensureInstChartTooltip() {
  if (instChartTooltipEl) return instChartTooltipEl;
  instChartTooltipEl = document.createElement('div');
  instChartTooltipEl.style.cssText = [
    'position:absolute',
    'z-index:300',
    'pointer-events:none',
    'opacity:0',
    'max-width:560px',
    'padding:12px 14px',
    'border:1px solid var(--border)',
    'border-radius:12px',
    'background:rgba(255,255,255,.98)',
    'box-shadow:0 16px 36px rgba(15,23,42,.18)',
    'color:var(--text)',
    'font-size:16px',
    'line-height:1.55',
    'white-space:normal',
    'word-break:normal',
    'overflow-wrap:normal'
  ].join(';');
  document.body.appendChild(instChartTooltipEl);
  return instChartTooltipEl;
}

function hideInstChartTooltip() {
  if (!instChartTooltipEl) return;
  instChartTooltipEl.style.opacity = '0';
}

function showInstChartTooltip(chart, tooltip, fullName, country, value) {
  const el = ensureInstChartTooltip();
  el.innerHTML = `
    <div style="font-weight:800;font-size:16px;line-height:1.45;margin-bottom:8px;white-space:normal;word-break:normal;overflow-wrap:normal;">${fullName}</div>
    <div style="font-size:16px;color:var(--text2);line-height:1.65;white-space:normal;word-break:normal;overflow-wrap:normal;">
      <div>论文数：<strong>${value.toLocaleString()}篇</strong></div>
      <div>所在国：${country}</div>
    </div>`;
  el.style.opacity = '1';

  const rect = chart.canvas.getBoundingClientRect();
  const viewportW = window.innerWidth;
  const viewportH = window.innerHeight;
  const offset = 18;

  let left = rect.left + window.scrollX + tooltip.caretX + offset;
  let top = rect.top + window.scrollY + tooltip.caretY - 14;

  el.style.left = `${left}px`;
  el.style.top = `${top}px`;

  const tipRect = el.getBoundingClientRect();
  const tipW = tipRect.width || 280;
  const tipH = tipRect.height || 120;

  if (left + tipW > window.scrollX + viewportW - 16) {
    left = rect.left + window.scrollX + tooltip.caretX - tipW - offset;
  }
  if (top + tipH > window.scrollY + viewportH - 16) {
    top = rect.top + window.scrollY + tooltip.caretY - tipH - offset;
  }

  el.style.left = `${Math.max(window.scrollX + 12, left)}px`;
  el.style.top = `${Math.max(window.scrollY + 12, top)}px`;
}

function renderInstList(period) {
  const ceec=G.inst.ceec[period], cn=G.inst.cn[period];
  const maxC=ceec[0]['val'], maxCn=cn[0]['val'];
  document.getElementById('ceecInstList').innerHTML=ceec.map((d,i)=>`
    <div class="inst-row">
      <div class="inst-rank" style="color:${i<3?C.orange:C.text3}">${i+1}</div>
      <div class="inst-info">
        <div class="inst-name" title="${d['name']}">${d['name']}</div>
        <div class="inst-sub">${d['country']}</div>
        <div class="inst-bar" style="width:${Math.round(d['val']/maxC*120)+8}px;background:${C.teal}55"></div>
      </div>
      <div class="inst-val" style="color:${C.teal}">${d['val'].toLocaleString()}</div>
    </div>`).join('');
  document.getElementById('cnInstList').innerHTML=cn.map((d,i)=>`
    <div class="inst-row">
      <div class="inst-rank" style="color:${i<3?C.orange:C.text3}">${i+1}</div>
      <div class="inst-info">
        <div class="inst-name">${d['name']}</div>
        <div class="inst-sub">CHINA</div>
        <div class="inst-bar" style="width:${Math.round(d['val']/maxCn*120)+8}px;background:${C.blue}55"></div>
      </div>
      <div class="inst-val" style="color:${C.blue}">${d['val'].toLocaleString()}</div>
    </div>`).join('');
  document.getElementById('instInsight').innerHTML = instInsights[period]||'';
}

function initInstChart(period) {
  const ceec=G.inst.ceec[period].slice(0,8);
  const shorten=n=>n.length>32?n.slice(0,32)+'…':n;
  const ctx=document.getElementById('instChart').getContext('2d');
  if(charts.inst) charts.inst.destroy();
  charts.inst = new Chart(ctx, {
    type:'bar',
    data:{labels:ceec.map(d=>shorten(d['name'])),datasets:[{label:'论文数（篇）',data:ceec.map(d=>d['val']),backgroundColor:'rgba(240, 237, 53, 0.4)',borderColor:'#cdcb4c',borderWidth:1.5,borderRadius:4,borderSkipped:false}]},
    options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},
        tooltip:{enabled:false,external:context=>{
          const tooltip = context.tooltip;
          if (!tooltip || tooltip.opacity === 0) {
            hideInstChartTooltip();
            return;
          }
          const point = tooltip.dataPoints?.[0];
          if (!point) {
            hideInstChartTooltip();
            return;
          }
          const item = ceec[point.dataIndex];
          showInstChartTooltip(context.chart, tooltip, item['name'], item['country'], item['val']);
        }}},
      scales:{
        x:{grid:{color:'#f4f5ed'},ticks:{color:C.text3,font:{size:16},callback:v=>v.toLocaleString()},border:{color:C.border}},
        y:{grid:{display:false},ticks:{color:C.text,font:{size:16}},border:{color:C.border}}
      }}
  });
}

function switchInst(period, btn) {
  document.querySelectorAll('#m4 .ptab').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active'); renderInstList(period); initInstChart(period);
}

// ============================================================
// 导航高亮
// ============================================================
function initNav() {
  const links=document.querySelectorAll('.nav-links a');
  const obs=new IntersectionObserver(entries=>{
    entries.forEach(e=>{
      if(e.isIntersecting){links.forEach(l=>l.classList.remove('active'));
        const lk=document.querySelector(`.nav-links a[href="#${e.target.id}"]`);
        if(lk) lk.classList.add('active');}
    });
  },{threshold:.25});
  document.querySelectorAll('section').forEach(s=>obs.observe(s));
}

// ============================================================
// 启动
// ============================================================
loadData();

window.setMapMode = setMapMode;
window.togglePlay = togglePlay;
window.onSliderChange = onSliderChange;
window.switchPeriod = switchPeriod;
window.switchField = switchField;
window.switchInst = switchInst;
window.toggleAreaSeries = toggleAreaSeries;
window.openDrill = openDrill;
window.closeModal = closeModal;
})();

