(() => {

// 机构英文名 → 中文名映射表（基于原始数据中的机构名称）
const INSTITUTION_NAME_MAP = {
  'Chinese Academy of Sciences': '中国科学院',
  'Institute of High Energy Physics, CAS': '中科院高能所',
  'Peking University': '北京大学',
  'University of Science & Technology of China, CAS': '中国科学技术大学',
  'Tsinghua University': '清华大学',
  'Shandong University': '山东大学',
  'Sun Yat Sen University': '中山大学',
  'Nanjing University': '南京大学',
  'Shanghai Jiao Tong University': '上海交通大学',
  'Central China Normal University': '华中师范大学',
  'Fudan University': '复旦大学',
  'Zhejiang University': '浙江大学',
  'University of Chinese Academy of Sciences, CAS': '中国科学院大学',
  'Beihang University': '北京航空航天大学',
  'Institute of Modern Physics, CAS': '中科院近代物理所',
  'University of Electronic Science & Technology of China': '电子科技大学',
  'Wuhan University': '武汉大学'
};

// 工具函数：用 fetch 获取 Excel 文件并解析为 JSON
async function fetchExcel(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`无法读取文件: ${url}`);
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  // header: 1 表示第一行是表头，返回对象数组
  return XLSX.utils.sheet_to_json(worksheet);
}

// ============================================================
// 图1 数据处理：从 3-中国-时间段-机构.xlsx 筛选前15名
// ============================================================
function processChartA(institutionRaw) {
  // 步骤1：按时间段分组，获取全部数据
  const periods = ['2011-2015', '2016-2020'];
  const allPeriodData = {};

  periods.forEach(period => {
    // 步骤2：筛选该时段的所有机构，按合作论文数降序排列
    const rows = institutionRaw
      .filter(d => d['时间段'] === period)
      .map(d => ({
        name_en: d['机构英文名'],
        name_cn: INSTITUTION_NAME_MAP[d['机构英文名']] || d['机构英文名'],
        papers: +d['合作Wos论文数'] || 0
      }))
      .sort((a, b) => b.papers - a.papers);

    // 步骤3：为每个机构赋予全局排名（1, 2, 3...）
    rows.forEach((d, i) => { d.globalRank = i + 1; });
    allPeriodData[period] = rows;
  });

  // 步骤4：取两个时段各自的前15名，求并集（即要显示的机构列表）
  const top15_2011 = allPeriodData['2011-2015'].slice(0, 15);
  const top15_2016 = allPeriodData['2016-2020'].slice(0, 15);

  const namesInTop15 = new Set([
    ...top15_2011.map(d => d.name_en),
    ...top15_2016.map(d => d.name_en)
  ]);

  // 步骤5：为并集中的每个机构，查找其在两个时段的真实排名和论文数
  const chartData = Array.from(namesInTop15).map(name_en => {
    const d2011 = allPeriodData['2011-2015'].find(d => d.name_en === name_en);
    const d2016 = allPeriodData['2016-2020'].find(d => d.name_en === name_en);

    const papers_2011 = d2011 ? d2011.papers : 0;
    const papers_2016 = d2016 ? d2016.papers : 0;
    const rank_2011 = d2011 ? d2011.globalRank : null;  // null 表示未进前15（用于绘图时放底部）
    const rank_2016 = d2016 ? d2016.globalRank : null;

    // 步骤6：判断状态
    let status;
    const in2011 = d2011 && d2011.globalRank <= 15;
    const in2016 = d2016 && d2016.globalRank <= 15;

    if (!in2011 && in2016) status = '新进入';
    else if (in2011 && !in2016) status = '退出';
    else if (rank_2016 < rank_2011) status = '上升';
    else if (rank_2016 > rank_2011) status = '下降';
    else status = '持平';

    return {
      name: INSTITUTION_NAME_MAP[name_en] || name_en,
      name_en,
      papers_2011,
      papers_2016,
      rank_2011,
      rank_2016,
      status
    };
  });

  // 步骤7：排序——按 2016-2020 排名升序，新进入的放最后
  chartData.sort((a, b) => {
    const ra = a.rank_2016 === null ? 999 : a.rank_2016;
    const rb = b.rank_2016 === null ? 999 : b.rank_2016;
    return ra - rb;
  });

  return chartData;
}

// ============================================================
// 图2 数据处理：从 2-中东欧-时间段-领域.xlsx + 领域.xlsx 汇总
// ============================================================
function processChartC(fieldPeriodRaw, fieldMapRaw) {
  // 步骤1：从 领域.xlsx 构建一级ID → 一级中文名称 映射
  // 注意：Excel中一级ID可能是数字，需要统一为两位字符串
  const fieldNameMap = {};
  fieldMapRaw.forEach(d => {
    const id = String(d['一级ID']).padStart(2, '0');
    if (!fieldNameMap[id]) {
      fieldNameMap[id] = d['一级中文名称'];
    }
  });

  // 步骤2：从 2-中东欧-时间段-领域.xlsx 按"时间段"和"一级ID"汇总论文数
  const summary = {};
  fieldPeriodRaw.forEach(d => {
    const period = d['时间段'];
    const fieldId = String(d['一级ID']).padStart(2, '0');
    const papers = +d['合作Wos论文数'] || 0;
    const key = `${period}|${fieldId}`;
    if (!summary[key]) summary[key] = { period, fieldId, papers: 0 };
    summary[key].papers += papers;
  });

  // 步骤3：重组为按领域分组的数据，计算增长率
  const fieldIds = [...new Set(Object.values(summary).map(d => d.fieldId))].sort();

  const chartData = fieldIds.map(fid => {
    const p1 = summary[`2011-2015|${fid}`];
    const p2 = summary[`2016-2020|${fid}`];
    const papers_2011 = p1 ? p1.papers : 0;
    const papers_2016 = p2 ? p2.papers : 0;
    const growth = papers_2011 > 0 ? (papers_2016 - papers_2011) / papers_2011 : 0;

    return {
      field_id: fid,
      field_name: fieldNameMap[fid] || fid,
      papers_2011,
      papers_2016,
      growth,
      visual_value: Math.round(Math.sqrt(papers_2016) * 10) / 10  // 开方，用于气泡/扇区大小
    };
  });

  // 步骤4：按 2016-2020 论文数降序排列（玫瑰图扇区按大小排列更美观）
  chartData.sort((a, b) => b.papers_2016 - a.papers_2016);

  return chartData;
}

// ============================================================
// 可视化绘制模块
// ============================================================

const tooltip = d3.select('body').append('div')
  .attr('class', 'tooltip').attr('id', 'tooltip')
  .style('opacity', 0);

function showtooltip(event, html) {
  tooltip.html(html)
    .style('left', (event.pageX + 15) + 'px')
    .style('top', (event.pageY - 10) + 'px')
    .style('opacity', 1);
}
function hidetooltip() { tooltip.style('opacity', 0); }

// ===== 图1：Bump Chart（排名变化轨迹图） =====
function drawChartA(data) {
  const container = d3.select('#chart-a');
  container.selectAll('*').remove();

  const width = 800, height = 680;
  const margin = {top: 50, right: 280, bottom: 50, left: 170};
  const innerW = width - margin.left - margin.right;
  const innerH = height - margin.top - margin.bottom;

  const svg = container.append('svg').attr('width', width).attr('height', height);
  const g = svg.append('g').attr('transform', `translate(${margin.left},${margin.top})`);

  // 绘图用排名：未进入前15的放在第16位（底部虚线位置）
  const plotData = data.map(d => ({
    ...d,
    plotRank_2011: d.rank_2011 === null ? 16 : d.rank_2011,
    plotRank_2016: d.rank_2016 === null ? 16 : d.rank_2016,
    exit: d.status === '退出',
    enter: d.status === '新进入'
  }));

  const xScale = d3.scalePoint()
    .domain(['2011-2015', '2016-2020'])
    .range([0, innerW]);

  const yScale = d3.scaleLinear()
    .domain([1, 16])
    .range([0, innerH]);

  const colorMap = {
    '上升': '#1a9850',
    '下降': '#d73027',
    '持平': '#969696',
    '新进入': '#4575b4',
    '退出': '#fc8d59'
  };

  // 水平网格线（1-15名 + 第16名虚线参考）
  g.selectAll('.grid-line')
    .data(d3.range(1, 17))
    .enter().append('line')
    .attr('x1', -30).attr('x2', innerW + 30)
    .attr('y1', d => yScale(d)).attr('y2', d => yScale(d))
    .attr('stroke', d => d === 16 ? '#ffcccc' : '#f0f0f0')
    .attr('stroke-width', d => d === 16 ? 1.5 : 1)
    .attr('stroke-dasharray', d => d === 16 ? '4,4' : 'none');

  // 第16名标注"未进入前15"
  g.append('text')
    .attr('x', -35).attr('y', yScale(16) + 4)
    .attr('text-anchor', 'end').attr('font-size', '16px').attr('fill', '#d73027')
    .text('未进入前15');
  g.append('text')
    .attr('x', innerW + 35).attr('y', yScale(16) + 4)
    .attr('text-anchor', 'start').attr('font-size', '16px').attr('fill', '#d73027')
    .text('未进入前15');

  // 中轴线
  g.append('line').attr('x1', innerW/2).attr('x2', innerW/2).attr('y1', 0).attr('y2', innerH)
    .attr('stroke', '#e0e0e0').attr('stroke-width', 1).attr('stroke-dasharray', '5,5');

  // 左右纵轴线
  g.append('line').attr('x1', 0).attr('x2', 0).attr('y1', -10).attr('y2', innerH+10)
    .attr('stroke', '#333').attr('stroke-width', 2);
  g.append('line').attr('x1', innerW).attr('x2', innerW).attr('y1', -10).attr('y2', innerH+10)
    .attr('stroke', '#333').attr('stroke-width', 2);

  // 轴标题
  g.append('text').attr('x', 0).attr('y', -25)
    .attr('text-anchor', 'middle').attr('font-size', '16px').attr('font-weight', 'bold').attr('fill', '#333')
    .text('2011-2015年');
  g.append('text').attr('x', innerW).attr('y', -25)
    .attr('text-anchor', 'middle').attr('font-size', '16px').attr('font-weight', 'bold').attr('fill', '#333')
    .text('2016-2020年');

  // 左侧排名标签 1-15
  g.selectAll('.rank-label-left')
    .data(d3.range(1, 16))
    .enter().append('text')
    .attr('x', -10).attr('y', d => yScale(d) + 4)
    .attr('text-anchor', 'end').attr('font-size', '16px').attr('fill', '#999')
    .text(d => 'No.' + d);

  // 右侧排名标签 1-15
  g.selectAll('.rank-label-right')
    .data(d3.range(1, 16))
    .enter().append('text')
    .attr('x', innerW + 10).attr('y', d => yScale(d) + 4)
    .attr('text-anchor', 'start').attr('font-size', '16px').attr('fill', '#999')
    .text(d => 'No.' + d);

  // Bump 曲线
  const lineGen = d3.line()
    .x(d => xScale(d.period))
    .y(d => yScale(d.rank))
    .curve(d3.curveBumpX);

  // 绘制连线
  plotData.forEach(d => {
    const points = [
      {period: '2011-2015', rank: d.plotRank_2011},
      {period: '2016-2020', rank: d.plotRank_2016}
    ];
    const isSpecial = d.exit || d.enter;
    const strokeColor = colorMap[d.status] || '#666';

    g.append('path')
      .datum(points)
      .attr('d', lineGen)
      .attr('stroke', strokeColor)
      .attr('stroke-width', isSpecial ? 2.5 : 4)
      .attr('stroke-dasharray', isSpecial ? '8,4' : 'none')
      .attr('fill', 'none')
      .attr('opacity', 0.85)
      .on('mouseover', function(event) {
        d3.select(this).attr('opacity', 1).attr('stroke-width', isSpecial ? 4 : 6);
        let html = `<strong>${d.name}</strong><br/>
                    2011-2015: ${d.rank_2011 !== null ? '第' + d.rank_2011 + '名 (' + d.papers_2011 + '篇)' : '未进入前15 (' + d.papers_2011 + '篇)'}<br/>
                    2016-2020: ${d.rank_2016 !== null ? '第' + d.rank_2016 + '名 (' + d.papers_2016 + '篇)' : '未进入前15 (' + d.papers_2016 + '篇)'}<br/>`;
        if (d.status === '新进入') html += `<span style="color:#4575b4;font-weight:bold;">● 新进入前15</span>`;
        else if (d.status === '退出') html += `<span style="color:#fc8d59;font-weight:bold;">● 掉出前15</span>`;
        else html += `排名变化: ${d.rank_2016 !== null && d.rank_2011 !== null ? (d.rank_2016 - d.rank_2011 > 0 ? '+' : '') + (d.rank_2016 - d.rank_2011) + ' 位' : '-'}`;
        showtooltip(event, html);
      })
      .on('mouseout', function() {
        d3.select(this).attr('opacity', 0.85).attr('stroke-width', isSpecial ? 2.5 : 4);
        hidetooltip();
      });
  });

  // 绘制节点
  plotData.forEach(d => {
    const col = colorMap[d.status] || '#666';

    // 左侧节点（2011-2015）
    g.append('circle')
      .attr('cx', xScale('2011-2015'))
      .attr('cy', yScale(d.plotRank_2011))
      .attr('r', d.enter ? 5 : 7)
      .attr('fill', d.enter ? 'white' : col)
      .attr('stroke', col).attr('stroke-width', 2);

    // 右侧节点（2016-2020）
    g.append('circle')
      .attr('cx', xScale('2016-2020'))
      .attr('cy', yScale(d.plotRank_2016))
      .attr('r', d.exit ? 5 : 7)
      .attr('fill', d.exit ? 'white' : col)
      .attr('stroke', col).attr('stroke-width', 2);
  });

  // 左侧机构名标签
  plotData.forEach(d => {
    const isEnter = d.enter;
    if (d.plotRank_2011 !== 16) {
      g.append('text')
        .attr('x', xScale('2011-2015') - 55)
        .attr('y', yScale(d.plotRank_2011) + 4)
        .attr('text-anchor', 'end')
        .attr('font-size', isEnter ? '14px' : '16px')
        .attr('fill', isEnter ? '#bbb' : '#333')
        .attr('font-weight', isEnter ? 'normal' : 'bold')
        .text(d.name);
    }
  });

  // 右侧机构名标签
  plotData.forEach(d => {
    const isExit = d.exit;
    if (d.plotRank_2016 !== 16) {
      g.append('text')
        .attr('x', xScale('2016-2020') + 55)
        .attr('y', yScale(d.plotRank_2016) + 4)
        .attr('text-anchor', 'start')
        .attr('font-size', isExit ? '14px' : '16px')
        .attr('fill', isExit ? '#bbb' : '#333')
        .attr('font-weight', isExit ? 'normal' : 'bold')
        .text(d.name);
    }
  });

  // 图例
  const legendW = 90, legendH = 130;
  const legend = svg.append('g').attr('transform', `translate(${width - 110}, ${height - 155})`);

  const legendData = [
    {c: '#1a9850', t: '排名上升'},
    {c: '#d73027', t: '排名下降'},
    {c: '#4575b4', t: '新进入前15'},
    {c: '#fc8d59', t: '掉出前15'},
    {c: '#969696', t: '排名持平'}
  ];
  legendData.forEach((d, i) => {
    legend.append('circle').attr('cx', 14).attr('cy', i * 22 + 18).attr('r', 5).attr('fill', d.c);
    legend.append('text').attr('x', 28).attr('y', i * 22 + 22).attr('font-size', '14px').attr('fill', '#555').text(d.t);
  });
}

// ===== 图2：Nightingale Rose Chart（南丁格尔玫瑰图） =====
function drawChartC(data) {
  const container = d3.select('#chart-c');
  container.selectAll('*').remove();

  const width = 950, height = 950;
  const margin = 80;
  const innerRadius = 70;
  const radius = Math.min(width, height) / 2 - margin - 80;

  const svg = container.append('svg').attr('width', width).attr('height', height);
  const g = svg.append('g').attr('transform', `translate(${width/2},${height/2})`);

  const N = data.length;
  const angleStep = (2 * Math.PI) / N;
  const pad = 0.03;

  const sqrtValue = d => Math.sqrt(Math.max(d.papers_2016, 0));
  const maxVal = d3.max(data, d => sqrtValue(d));
  const rScale = d3.scaleLinear().domain([0, maxVal]).range([innerRadius, innerRadius + radius]);

  const colorScale = d3.scaleThreshold()
    .domain([1.0, 1.5, 2.0])
    .range(['#4575b4', '#74add1', '#f46d43', '#d73027']);

  // 同心圆参考线
  const gridRs = d3.range(innerRadius, innerRadius + radius + 1, radius / 4);
  g.selectAll('.grid-circle')
    .data(gridRs)
    .enter().append('circle')
    .attr('r', d => d)
    .attr('fill', 'none').attr('stroke', '#f0f0f0').attr('stroke-width', 1);

  // 扇区
  const arcGen = d3.arc()
    .innerRadius(innerRadius)
    .outerRadius(d => rScale(sqrtValue(d)))
    .startAngle((d, i) => i * angleStep + pad)
    .endAngle((d, i) => (i + 1) * angleStep - pad)
    .padAngle(0.01);

  g.selectAll('.rose-petal')
    .data(data)
    .enter().append('path')
    .attr('class', 'rose-petal')
    .attr('d', arcGen)
    .attr('fill', d => colorScale(d.growth))
    .attr('stroke', '#fff').attr('stroke-width', 1.5)
    .attr('opacity', 0.9)
    .on('mouseover', function(event, d) {
      d3.select(this).attr('opacity', 1).attr('stroke', '#333').attr('stroke-width', 2);
      const growthPct = (d.growth * 100).toFixed(0);
      showtooltip(event, `<strong>${d.field_name}</strong><br/>
                          2016-2020: ${d.papers_2016}篇<br/>
                          2011-2015: ${d.papers_2011}篇<br/>
                          增长率: +${growthPct}%`);
    })
    .on('mouseout', function() {
      d3.select(this).attr('opacity', 0.9).attr('stroke', '#fff').attr('stroke-width', 1.5);
      hidetooltip();
    });

  // 复合标签：学科名称 + 数值，合并标注
  g.selectAll('.label-combo')
    .data(data)
    .enter().append('text')
    .attr('transform', (d, i) => {
      const angle = i * angleStep + angleStep / 2;
      const r = rScale(sqrtValue(d)) + 48;
      const x = r * Math.sin(angle);
      const y = -r * Math.cos(angle);
      return `translate(${x},${y})`;
    })
    .attr('text-anchor', 'middle')
    .attr('font-size', '14px')
    .attr('fill', '#555')
    .style('pointer-events', 'none')
    .each(function(d) {
      const text = d3.select(this);
      text.append('tspan')
        .attr('x', 0)
        .attr('dy', 0)
        .attr('font-weight', 'bold')
        .attr('fill', '#333')
        .text(d.field_name);
      text.append('tspan')
        .attr('x', 0)
        .attr('dy', '1.25em')
        .attr('font-size', '13px')
        .attr('fill', '#666')
        .text(`${d.papers_2016}篇`);
    });

  // 中心标题
  g.append('text')
    .attr('text-anchor', 'middle')
    .attr('font-size', '16px').attr('font-weight', 'bold').attr('fill', '#333')
    .text('学科合作规模');
  g.append('text')
    .attr('text-anchor', 'middle').attr('y', 16)
    .attr('font-size', '14px').attr('fill', '#888')
    .text('(2016-2020)');

  // 图例
  const legendG = svg.append('g').attr('transform', `translate(${width - 150}, ${height - 110})`);
  const legendData = [
    {c: '#4575b4', t: '增长 < 100%'},
    {c: '#74add1', t: '增长 100-150%'},
    {c: '#f46d43', t: '增长 150-200%'},
    {c: '#d73027', t: '增长 > 200%'}
  ];
  legendData.forEach((d, i) => {
    legendG.append('rect').attr('x', 0).attr('y', i * 24).attr('width', 14).attr('height', 14).attr('fill', d.c).attr('rx', 2);
    legendG.append('text').attr('x', 20).attr('y', i * 24 + 11).attr('font-size', '14px').attr('fill', '#555').text(d.t);
  });
}

// ============================================================
// 主程序：加载数据 → 处理 → 绘制
// ============================================================
async function main() {
  try {
    // 并行读取三个 Excel 源文件
    const [fieldPeriodRaw, institutionRaw, fieldMapRaw] = await Promise.all([
      fetchExcel('../data/2-中东欧-时间段-领域.xlsx'),
      fetchExcel('../data/3-中国-时间段-机构.xlsx'),
      fetchExcel('../data/领域.xlsx')
    ]);

    console.log('【数据读取成功】');
    console.log('../data/2-中东欧-时间段-领域.xlsx:', fieldPeriodRaw.length, '条记录');
    console.log('../data/3-中国-时间段-机构.xlsx:', institutionRaw.length, '条记录');
    console.log('../data/领域.xlsx:', fieldMapRaw.length, '条记录');

    // 图1：处理机构排名数据
    const chartAData = processChartA(institutionRaw);
    console.log('【图1数据处理】共', chartAData.length, '个机构进入并集（两个时段前15名的并集）');
    console.log(chartAData);

    // 图2：处理学科汇总数据
    const chartCData = processChartC(fieldPeriodRaw, fieldMapRaw);
    console.log('【图2数据处理】共', chartCData.length, '个一级学科');
    console.log(chartCData);

    // 绘制图表
    drawChartA(chartAData);
    drawChartC(chartCData);

    // 隐藏加载提示，显示图表
    d3.select('#loading-status').style('display', 'none');
    d3.select('#section-a').style('display', 'block');
    d3.select('#section-c').style('display', 'block');

  } catch (err) {
    console.error('数据加载或处理失败:', err);
    d3.select('#loading-status').html(
      `<p style="color:#d73027;">数据加载失败：${err.message}</p>
      <p style="color:#666;font-size:16px;">请确保三个 .xlsx 文件与 index.html 位于同一目录，且通过本地服务器（如 Live Server）访问。</p>`
    );
  }
}

// 页面加载完成后执行
document.addEventListener('DOMContentLoaded', main);

})();

