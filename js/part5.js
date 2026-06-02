(()=>{
let allCountriesData = [];
let filteredCountries = [];
let currentSelectedCountry = null;
let currentClusterFilter = "all";
let radarCharts = [];
let clusterChart = null;

// 维度定义 
const dimensions = [
    { key: "scale", name: "合作规模", label: "1" },
    { key: "growth", name: "增长速度", label: "2" },
    { key: "breadth", name: "领域广度", label: "3" },
    { key: "institution", name: "机构参与度", label: "4" },
    { key: "stability", name: "合作稳定性", label: "5" }
];

// 地区映射
const regionMapping = {
    "波兰": "中欧", "捷克": "中欧", "匈牙利": "中欧", "斯洛文尼亚": "中欧", "斯洛伐克": "中欧",
    "罗马尼亚": "东欧", "保加利亚": "东欧", "爱沙尼亚": "东欧", "拉脱维亚": "东欧",
    "希腊": "巴尔干", "塞尔维亚": "巴尔干", "克罗地亚": "巴尔干", "马其顿": "巴尔干", 
    "波黑": "巴尔干", "阿尔巴尼亚": "巴尔干", "黑山": "巴尔干"
};

// 一级学科总数
const TOTAL_DISCIPLINES = 14;

// 等待页面加载完成
document.addEventListener('DOMContentLoaded', () => {
    console.log('页面加载完成，开始加载数据...');
    loadExcelData();
});

// 加载Excel数据
async function loadExcelData() {
    showLoading(true);
    
    try {
        await loadSheetJSLibrary();
        
        const files = [
            '../data/1-中东欧-时间段-总量.xlsx',
            '../data/2-中东欧前12国-时间段（仅20162020）-领域.xlsx',
            '../data/3-中东欧-时间段-机构.xlsx',
            '../data/中东欧16国.xlsx'
        ];
        
        const data = {};
        for (const file of files) {
            data[file] = await loadExcelFile(file);
            console.log(`成功加载: ${file}`);
        }
        
        allCountriesData = parseCountriesData(data);
        filteredCountries = [...allCountriesData];
        updateStatsDisplay();
        renderRadarMatrix();
        renderClusterTree();
        bindEvents();
        
    } catch (error) {
        console.error('加载数据失败:', error);
        const container = document.getElementById('radarMatrix');
        container.innerHTML = `<div class="loading" style="color: #e74c3c;">数据加载失败: ${error.message}<br><br>请确保 data 文件夹中存在所需的Excel文件</div>`;
        showLoading(false);
    }
}

// 加载SheetJS库
function loadSheetJSLibrary() {
    return new Promise((resolve, reject) => {
        if (typeof XLSX !== 'undefined') {
            console.log('SheetJS库已加载');
            resolve();
            return;
        }
        const script = document.createElement('script');
        script.src = 'https://cdn.sheetjs.com/xlsx-0.20.2/package/dist/xlsx.full.min.js';
        script.onload = () => {
            console.log('SheetJS库加载完成');
            resolve();
        };
        script.onerror = () => reject(new Error('SheetJS库加载失败'));
        document.head.appendChild(script);
    });
}

// 读取Excel文件
async function loadExcelFile(filename) {
    const response = await fetch(filename);
    if (!response.ok) {
        throw new Error(`无法读取文件: ${filename}`);
    }
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet);
}

// 解析所有数据
function parseCountriesData(data) {
    const countries = [];
    
    const countryList = data['../data/中东欧16国.xlsx'] || [];
    const periodData = data['../data/1-中东欧-时间段-总量.xlsx'] || [];
    const fieldData = data['../data/2-中东欧前12国-时间段（仅20162020）-领域.xlsx'] || [];
    const institutionData = data['../data/3-中东欧-时间段-机构.xlsx'] || [];
    
    console.log(`国家列表: ${countryList.length}个`);
    console.log(`时段数据: ${periodData.length}条`);
    console.log(`领域数据: ${fieldData.length}条`);
    console.log(`机构数据: ${institutionData.length}条`);
    
    const period1 = periodData.filter(d => d['时间段'] === '2011-2015');
    const period2 = periodData.filter(d => d['时间段'] === '2016-2020');
    
    const pubMap1 = {};
    const pubMap2 = {};
    period1.forEach(d => { pubMap1[d['国家中文名']] = d['合作Wos论文数']; });
    period2.forEach(d => { pubMap2[d['国家中文名']] = d['合作Wos论文数']; });
    
    // 计算每个国家的领域广度
    const fieldMap = {};
    fieldData.forEach(d => {
        const country = d['国家中文名'];
        const level1 = String(d['一级ID']).substring(0, 2);
        if (country && level1 && level1 !== 'undefined') {
            if (!fieldMap[country]) fieldMap[country] = new Set();
            fieldMap[country].add(level1);
        }
    });
    
    // 国家英文名转中文名的映射
    const nameMap = {
        'POLAND': '波兰', 'CZECH REPUBLIC': '捷克', 'GREECE': '希腊',
        'HUNGARY': '匈牙利', 'ROMANIA': '罗马尼亚', 'SERBIA': '塞尔维亚',
        'SLOVENIA': '斯洛文尼亚', 'SLOVAKIA': '斯洛伐克', 'CROATIA': '克罗地亚',
        'BULGARIA': '保加利亚', 'ESTONIA': '爱沙尼亚', 'LATVIA': '拉脱维亚',
        'MACEDONIA': '马其顿', 'BOSNIA & HERZEGOVINA': '波黑', 'ALBANIA': '阿尔巴尼亚',
        'MONTENEGRO': '黑山'
    };
    
    // 第一步：统计所有中东欧国家参与合作的总机构数量（2011-2020）
    let totalCEECInstitutions = new Set();
    institutionData.forEach(d => {
        const countryEn = d['国家或地区'];
        let chName = nameMap[countryEn] || countryEn;
        if (chName && chName !== countryEn) {
            totalCEECInstitutions.add(d['机构英文名']);
        }
    });
    const totalCEECCount = totalCEECInstitutions.size;
    console.log(`中东欧合作机构总数: ${totalCEECCount}`);
    
    // 第二步：统计每个国家2016-2020年参与合作的机构
    const instMap = {};
    institutionData.forEach(d => {
        const countryEn = d['国家或地区'];
        let chName = nameMap[countryEn] || countryEn;
        if (chName && chName !== countryEn) {
            if (d['时间段'] === '2016-2020') {
                if (!instMap[chName]) instMap[chName] = new Set();
                instMap[chName].add(d['机构英文名']);
            }
        }
    });
    
    // 计算最大总论文数（2011-2015 + 2016-2020）
    let maxTotalPub = 0;
    const allCountryNames = new Set([...Object.keys(pubMap1), ...Object.keys(pubMap2)]);
    allCountryNames.forEach(name => {
        const total = (pubMap1[name] || 0) + (pubMap2[name] || 0);
        if (total > maxTotalPub) maxTotalPub = total;
    });
    const maxPub = maxTotalPub;
    console.log(`最大总论文数: ${maxPub}`);
    
    const maxGrowth = 15;
    
    for (const country of countryList) {
        const name = country['国家中文名'];
        const pub2015 = pubMap1[name] || 1;
        const pub2020 = pubMap2[name] || 1;
        
        // 总论文数 = 两个时期相加
        const total = pub2015 + pub2020;
        
        // 1. 合作规模（归一化）- 使用总论文数
        const scale = total / maxPub;
        
        // 2. 增长速度
        const growthRaw = pub2020 / pub2015;
        const growth = Math.min(1, growthRaw / maxGrowth);
        
        // 3. 领域广度：直接使用实际数据，没有就是0
        const breadth = (fieldMap[name]?.size || 0) / TOTAL_DISCIPLINES;
        
        // 4. 机构参与度：该国机构数 / 中东欧总机构数
        const instCount = instMap[name]?.size || 0;
        const institution = instCount / totalCEECCount;
        
        // 5. 合作稳定性
        const mean = (pub2015 + pub2020) / 2;
        const variance = Math.pow(pub2015 - mean, 2) + Math.pow(pub2020 - mean, 2);
        const cv = Math.sqrt(variance / 2) / mean;
        const stability = cv > 0 ? Math.min(1, 1 / (cv * 3)) : 1;
        
        countries.push({
            name: name,
            nameEn: country['国家英文名'],
            region: regionMapping[name] || "其他",
            color: getCountryColor(name),
            scale: Math.min(1, scale),
            growth: growth,
            breadth: Math.min(1, breadth),
            institution: institution,
            stability: stability,
            total: total,
            pub2015: pub2015,
            pub2020: pub2020,
            disciplines: fieldMap[name]?.size || 0,
            institutions: instMap[name]?.size || 0
        });
    }
    
    console.log(`解析完成，共 ${countries.length} 个国家`);
    return countries.sort((a, b) => b.total - a.total);
}

// 获取国家颜色 - 沉稳学术风格
function getCountryColor(name) {
    if (typeof getSharedCountryColor === 'function') return getSharedCountryColor(name);
    const colors = {
        '波兰': '#B2182B','捷克': '#2166AC','希腊': '#1B9E77','匈牙利': '#7570B3',
        '罗马尼亚': '#D95F02','塞尔维亚': '#E7298A','斯洛文尼亚': '#66A61E','斯洛伐克': '#E6AB02',
        '克罗地亚': '#A6761D','保加利亚': '#666666','爱沙尼亚': '#1B9E77','拉脱维亚': '#D95F02',
        '马其顿': '#7570B3','波黑': '#E7298A','阿尔巴尼亚': '#66A61E','黑山': '#E6AB02'
    };
    return colors[name] || '#888888';
}

// 计算聚类类型
function getClusterType(country) {
    const avgScale = 0.25, avgGrowth = 0.35, avgBreadth = 0.25, avgInstitution = 0.22, avgStability = 0.35;
    
    const highScale = country.scale > avgScale;
    const highGrowth = country.growth > avgGrowth;
    const highBreadth = country.breadth > avgBreadth;
    const highInstitution = country.institution > avgInstitution;
    const highStability = country.stability > avgStability;
    
    const highCount = [highScale, highGrowth, highBreadth, highInstitution, highStability].filter(v => v).length;
    
    if (highCount >= 4) return "全面深度合作型";
    if (country.growth > avgGrowth * 1.3 && country.scale < avgScale * 1.5) return "快速增长型";
    if (country.breadth < avgBreadth && country.scale > avgScale * 0.6) return "领域聚焦型";
    return "起步探索型";
}

// 更新统计显示
function updateStatsDisplay() {
    document.getElementById('visibleCount').textContent = filteredCountries.length;
    document.getElementById('totalCount').textContent = allCountriesData.length;
}

// 显示/隐藏加载状态
function showLoading(show) {
    const container = document.getElementById('radarMatrix');
    if (show) {
        container.innerHTML = '<div class="loading">正在加载数据，请稍候...</div>';
    }
}

// 渲染雷达图矩阵
function renderRadarMatrix() {
    const container = document.getElementById('radarMatrix');
    container.innerHTML = '';
    radarCharts = [];
    
    let countriesToShow = [...filteredCountries];
    const sortBy = document.getElementById('sortBy').value;
    if (sortBy !== 'none') {
        countriesToShow.sort((a, b) => b[sortBy] - a[sortBy]);
    }
    
    if (currentClusterFilter !== 'all') {
        countriesToShow = countriesToShow.filter(c => getClusterType(c) === currentClusterFilter);
    }
    
    if (countriesToShow.length === 0) {
        container.innerHTML = '<div class="loading">暂无数据</div>';
        return;
    }
    
    for (const country of countriesToShow) {
        const card = createRadarCard(country);
        container.appendChild(card);
    }
}

// 创建雷达图卡片
function createRadarCard(country) {
    const card = document.createElement('div');
    card.className = 'radar-card';
    if (currentSelectedCountry === country.name) card.classList.add('selected');
    
    const clusterType = getClusterType(country);
    
    card.innerHTML = `
        <div class="country-name" style="background: ${country.color}20; color: ${country.color}; border-left: 3px solid ${country.color}; padding:4px 8px;">
            <div>${country.name}</div>
            <div style="font-size: 14px; margin-top:2px;">(${clusterType})</div>
        </div>
        <canvas width="160" height="160" style="width:100%; height:auto; max-width:160px; margin:0 auto; display:block;"></canvas>
    `;
    
    const canvas = card.querySelector('canvas');
    drawRadarChart(canvas, country);
    
    card.addEventListener('click', (e) => {
        e.stopPropagation();
        if (currentSelectedCountry === country.name) {
            currentSelectedCountry = null;
            document.querySelectorAll('.radar-card').forEach(c => c.classList.remove('selected'));
        } else {
            currentSelectedCountry = country.name;
            document.querySelectorAll('.radar-card').forEach(c => c.classList.remove('selected'));
            card.classList.add('selected');
            updateOtherCharts(country);
        }
    });
    
    card.addEventListener('mouseenter', () => {
        showDetailTooltip(country, card);
    });
    
    card.addEventListener('mouseleave', () => {
        hideDetailTooltip();
    });
    
    return card;
}

// 绘制雷达图
function drawRadarChart(canvas, country) {
    const ctx = canvas.getContext('2d');
    const width = canvas.width, height = canvas.height;
    const centerX = width / 2, centerY = height / 2;
    const radius = Math.min(width, height) * 0.38;
    const angles = dimensions.map((_, i) => (i * 2 * Math.PI / dimensions.length) - Math.PI / 2);
    
    ctx.clearRect(0, 0, width, height);
    
    // 绘制背景网格
    const levels = [0.2, 0.4, 0.6, 0.8, 1.0];
    for (const level of levels) {
        ctx.beginPath();
        angles.forEach((angle, i) => {
            const x = centerX + radius * level * Math.cos(angle);
            const y = centerY + radius * level * Math.sin(angle);
            if (i === 0) ctx.moveTo(x, y);
            else ctx.lineTo(x, y);
        });
        ctx.closePath();
        ctx.strokeStyle = '#e0e0e0';
        ctx.fillStyle = level === 0.2 ? '#fafafa' : 'transparent';
        ctx.fill();
        ctx.stroke();
    }
    
    // 绘制轴线
    angles.forEach(angle => {
        const x = centerX + radius * Math.cos(angle);
        const y = centerY + radius * Math.sin(angle);
        ctx.beginPath();
        ctx.moveTo(centerX, centerY);
        ctx.lineTo(x, y);
        ctx.strokeStyle = '#e0e0e0';
        ctx.stroke();
    });
    
    // 绘制数据区域
    ctx.beginPath();
    const values = dimensions.map(d => country[d.key]);
    angles.forEach((angle, i) => {
        const value = Math.min(1, Math.max(0, values[i]));
        const x = centerX + radius * value * Math.cos(angle);
        const y = centerY + radius * value * Math.sin(angle);
        if (i === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
    });
    ctx.closePath();
    ctx.fillStyle = country.color + '90'; 
    ctx.fill();
    ctx.strokeStyle = country.color;
    ctx.lineWidth = 1.5;
    ctx.stroke();
    
    // 绘制数据点
    angles.forEach((angle, i) => {
        const value = Math.min(1, Math.max(0, values[i]));
        const x = centerX + radius * value * Math.cos(angle);
        const y = centerY + radius * value * Math.sin(angle);
        ctx.beginPath();
        ctx.arc(x, y, 3, 0, 2 * Math.PI);
        ctx.fillStyle = '#666';
        ctx.fill();
        ctx.strokeStyle = 'white';
        ctx.lineWidth = 1;
        ctx.stroke();
    });
    
    // 绘制维度标签
    ctx.font = '11px "Helvetica Neue", Arial, sans-serif';
    ctx.fillStyle = '#999';
    angles.forEach((angle, i) => {
        const labelX = centerX + (radius + 10) * Math.cos(angle);
        const labelY = centerY + (radius + 10) * Math.sin(angle);
        ctx.fillText((i + 1).toString(), labelX - 4, labelY + 4);
    });
}

// 显示详情提示框
function showDetailTooltip(country, element) {
    const tooltip = document.getElementById('detailTooltip');
    const clusterType = getClusterType(country);
    
    let clusterColor = '#666';
    if (clusterType === '全面深度合作型') clusterColor = '#2A9D8F';
    else if (clusterType === '快速增长型') clusterColor = '#F4A261';
    else if (clusterType === '领域聚焦型') clusterColor = '#9B5DE5';
    else clusterColor = '#6C757D';
    
    tooltip.innerHTML = `
        <div style="font-size:18px; font-weight:600; margin-bottom:8px; color:${country.color}">${country.name}</div>
        <div style="font-size:16px; margin-bottom:6px;">类型: <span style="color:${clusterColor}">${clusterType}</span></div>
        <div style="font-size:16px;">总论文数: ${country.total} 篇</div>
        <div style="font-size:16px;">2011-2015: ${country.pub2015} | 2016-2020: ${country.pub2020}</div>
        <div style="font-size:16px;">合作学科: ${country.disciplines} 个一级学科</div>
        <div style="font-size:16px;">合作机构: ${country.institutions} 个</div>
        <hr style="margin:6px 0; border-color:#444;">
        ${dimensions.map(d => `${d.label} ${d.name}: ${(country[d.key] * 100).toFixed(1)}%`).join('<br>')}
    `;
    
    const rect = element.getBoundingClientRect();
    tooltip.style.left = rect.right + 10 + 'px';
    tooltip.style.top = rect.top + 'px';
    tooltip.style.opacity = '1';
}

// 隐藏提示框
function hideDetailTooltip() {
    const tooltip = document.getElementById('detailTooltip');
    tooltip.style.opacity = '0';
}

// 更新其他图表
function updateOtherCharts(country) {
    if (clusterChart) {
        // 清除聚类树中所有叶子的高亮
        clusterChart.selectAll('.leaf-polygon')
            .attr('stroke', 'none')
            .attr('stroke-width', 0);
        
        // 高亮对应的叶子节点
        clusterChart.selectAll('.leaf-node')
            .filter(function() {
                return d3.select(this).attr('data-country') === country.name;
            })
            .select('.leaf-polygon')
            .attr('stroke', '#ff6b35')
            .attr('stroke-width', 2);
    }
}

// 高亮雷达图中对应的国家卡片
function highlightRadarCard(countryName) {
    const cards = document.querySelectorAll('.radar-card');
    cards.forEach(card => {
        const nameDiv = card.querySelector('.country-name');
        if (nameDiv && nameDiv.textContent.includes(countryName)) {
            card.classList.add('selected');
            card.scrollIntoView({ behavior: 'smooth', block: 'center' });
        } else {
            card.classList.remove('selected');
        }
    });
}

// 渲染径向树状图
function renderClusterTree() {
    const container = document.getElementById('clusterTree');
    container.innerHTML = '';
    
    const clusteredData = buildClusterData();
    
    const width = 500;
    const height = 500;
    const radius = Math.min(width, height) * 0.25;
    
    const svg = d3.select('#clusterTree')
        .append('svg')
        .attr('width', width)
        .attr('height', height)
        .append('g')
        .attr('transform', `translate(${width / 2}, ${height / 2})`);
    
    clusterChart = svg;
    
    const treeLayout = d3.tree()
        .size([2 * Math.PI, radius])
        .separation((a, b) => {
            if (a.parent === b.parent) {
                return 0.6;
            }
            return 1;
        });
    
    const root = d3.hierarchy(clusteredData);
    const treeData = treeLayout(root);
    
    svg.selectAll('.cluster-link')
        .data(treeData.links())
        .enter()
        .append('path')
        .attr('class', 'cluster-link')
        .attr('d', d3.linkRadial()
            .angle(d => d.x)
            .radius(d => d.y)
        )
        .attr('fill', 'none')
        .attr('stroke', '#ccc')
        .attr('stroke-width', 1);
    
    const rootNode = treeData.descendants().find(d => d.depth === 0);
    const clusterNodes = treeData.descendants().filter(d => d.data.type && d.depth > 0);
    const leafNodes = treeData.descendants().filter(d => d.data.name);
    
    // 根节点
    if (rootNode) {
        const x = Math.sin(rootNode.x) * rootNode.y;
        const y = -Math.cos(rootNode.x) * rootNode.y;
        
        const rootGroup = svg.append('g')
            .attr('transform', `translate(${x+3}, ${y+6})`)
            .attr('cursor', 'pointer')
            .on('click', function() {
                currentClusterFilter = 'all';
                currentSelectedCountry = null;
                document.getElementById('clusterFilter').value = 'all';
                renderRadarMatrix();
                renderClusterTree();
            });
        
        rootGroup.append('circle')
            .attr('r', 5)
            .attr('fill', '#1a3a6b')
            .attr('stroke', 'white')
            .attr('stroke-width', 2);
        
        rootGroup.append('text')
            .attr('x', 0)
            .attr('y', 22)
            .attr('text-anchor', 'middle')
            .style('font-size', '15px')
            .style('font-weight', 'bold')
            .style('fill', '#1a3a6b')
            .style('cursor', 'pointer')
            .text('root');
    }
    
    // 聚类节点
    clusterNodes.forEach(d => {
        const x = Math.sin(d.x) * d.y;
        const y = -Math.cos(d.x) * d.y;
        const angleDeg = d.x * 180 / Math.PI;
        
        const group = svg.append('g')
            .attr('transform', `translate(${x}, ${y})`)
            .attr('cursor', 'pointer')
            .on('click', function() {
                currentClusterFilter = d.data.type;
                currentSelectedCountry = null;
                document.getElementById('clusterFilter').value = currentClusterFilter;
                renderRadarMatrix();
                renderClusterTree();
            });
        
        group.append('circle')
            .attr('r', 5)
            .attr('fill', d.data.color)
            .attr('stroke', 'white')
            .attr('stroke-width', 1.5);
        
        let labelX = 0, labelY = 0, textAnchor = 'middle';
        
        if (angleDeg >= -45 && angleDeg <= 45) {
            labelX = 12; labelY = 3; textAnchor = 'start';
        } else if (angleDeg > 45 && angleDeg < 135) {
            labelX = 0; labelY = 16; textAnchor = 'middle';
        } else if (angleDeg >= 135 && angleDeg <= 225) {
            labelX = -12; labelY = 3; textAnchor = 'end';
        } else {
            labelX = 0; labelY = -10; textAnchor = 'middle';
        }
        
        let typeName = d.data.type;
        if (typeName === '全面深度合作型') typeName = '全面型';
        else if (typeName === '快速增长型') typeName = '增长型';
        else if (typeName === '领域聚焦型') typeName = '聚焦型';
        else if (typeName === '起步探索型') typeName = '起步型';
        
        group.append('text')
            .attr('x', labelX)
            .attr('y', labelY)
            .attr('text-anchor', textAnchor)
            .style('font-size', '14px')
            .style('font-weight', 'bold')
            .style('fill', d.data.color)
            .style('cursor', 'pointer')
            .text(typeName);
    });
    
    // 叶子节点
    leafNodes.forEach(d => {
        const angle = d.x;
        const rad = d.y;
        const x = Math.sin(angle) * rad;
        const y = -Math.cos(angle) * rad;
        const angleDeg = angle * 180 / Math.PI;
        
        const size = Math.sqrt(d.data.size || 100) * 1.2 + 3;
        const dirX = Math.sin(angle);
        const dirY = -Math.cos(angle);
        const perpX = Math.cos(angle);
        const perpY = Math.sin(angle);
        const width = 4;
        
        const tipX = x + dirX * size;
        const tipY = y + dirY * size;
        const leftX = x + dirX * (size/2) - perpX * width;
        const leftY = y + dirY * (size/2) - perpY * width;
        const rightX = x + dirX * (size/2) + perpX * width;
        const rightY = y + dirY * (size/2) + perpY * width;
        const baseX = x - dirX * 1;
        const baseY = y - dirY * 1;
        
        const countryName = d.data.name;
        
        const group = svg.append('g')
            .attr('class', 'leaf-node')
            .attr('data-country', countryName)
            .attr('cursor', 'pointer')
            .on('click', function(event) {
                event.stopPropagation();
                
                currentSelectedCountry = countryName;
                renderRadarMatrix();
                
                svg.selectAll('.leaf-polygon')
                    .attr('stroke', 'none')
                    .attr('stroke-width', 0);
                
                d3.select(this).select('.leaf-polygon')
                    .attr('stroke', '#ff6b35')
                    .attr('stroke-width', 2);
            });
        
        const polygon = group.append('polygon')
            .attr('class', 'leaf-polygon')
            .attr('points', `${baseX},${baseY} ${leftX},${leftY} ${tipX},${tipY} ${rightX},${rightY}`)
            .attr('fill', d.data.color)
            .attr('stroke', currentSelectedCountry === countryName ? '#ff6b35' : 'none')
            .attr('stroke-width', currentSelectedCountry === countryName ? 2 : 0)
            .attr('opacity', 0.9);
        
        polygon.on('mouseenter', function() {
            d3.select(this).attr('opacity', 1).attr('stroke', '#333').attr('stroke-width', 0.5);
        }).on('mouseleave', function() {
            const isSelected = (currentSelectedCountry === countryName);
            d3.select(this)
                .attr('opacity', 0.9)
                .attr('stroke', isSelected ? '#ff6b35' : 'none')
                .attr('stroke-width', isSelected ? 2 : 0);
        });
        
        let labelX = tipX, labelY = tipY, textAnchor = 'middle';
        if (angleDeg >= -60 && angleDeg <= 60) {
            labelX = tipX + 5; textAnchor = 'start';
        } else if (angleDeg > 60 && angleDeg < 120) {
            labelY = tipY + 5; textAnchor = 'middle';
        } else if (angleDeg >= 120 && angleDeg <= 240) {
            labelX = tipX - 5; textAnchor = 'end';
        } else {
            labelY = tipY - 5; textAnchor = 'middle';
        }
        
        let shortName = d.data.name;
        const shortMap = {
            '斯洛文尼亚': '斯文', '斯洛伐克': '斯伐', '保加利亚': '保加',
            '罗马尼亚': '罗马', '塞尔维亚': '塞尔', '克罗地亚': '克罗',
            '阿尔巴尼亚': '阿尔', '马其顿': '马其', '拉脱维亚': '拉脱',
            '爱沙尼亚': '爱沙', '黑山': '黑山', '波黑': '波黑',
            '波兰': '波兰', '捷克': '捷克', '希腊': '希腊', '匈牙利': '匈牙'
        };
        shortName = shortMap[shortName] || shortName.substring(0, 2);
        
        group.append('text')
            .attr('x', labelX)
            .attr('y', labelY)
            .attr('text-anchor', textAnchor)
            .style('font-size', '14px')
            .style('fill', '#555')
            .style('font-weight', '500')
            .style('pointer-events', 'none')
            .text(shortName);
        
        group.append('title')
            .text(`${d.data.name}\n论文数: ${d.data.size}篇`);
    });
}

// 构建聚类树数据
function buildClusterData() {
    const clusters = {
        "全面深度合作型": [],
        "快速增长型": [],
        "领域聚焦型": [],
        "起步探索型": []
    };
    
    filteredCountries.forEach(country => {
        const type = getClusterType(country);
        clusters[type].push(country);
    });
    
    Object.keys(clusters).forEach(key => {
        clusters[key].sort((a, b) => b.total - a.total);
    });
    
    return {
        children: Object.entries(clusters)
            .filter(([_, countries]) => countries.length > 0)
            .map(([type, countries]) => ({
                type: type,
                color: type === "全面深度合作型" ? "#2A9D8F" : 
                       type === "快速增长型" ? "#F4A261" :
                       type === "领域聚焦型" ? "#9B5DE5" : "#6C757D",
                children: countries.map(c => ({
                    name: c.name,
                    size: c.total,
                    color: c.color,
                    ...c
                }))
            }))
    };
}

// 绑定事件监听
function bindEvents() {
    document.getElementById('regionSelect').addEventListener('change', (e) => {
        const region = e.target.value;
        if (region === 'all') {
            filteredCountries = [...allCountriesData];
        } else {
            filteredCountries = allCountriesData.filter(c => c.region === region);
        }
        updateStatsDisplay();
        renderRadarMatrix();
        renderClusterTree();
    });
    
    document.getElementById('sortBy').addEventListener('change', () => {
        renderRadarMatrix();
    });
    
    document.getElementById('clusterFilter').addEventListener('change', (e) => {
        currentClusterFilter = e.target.value;
        renderRadarMatrix();
        renderClusterTree();
    });
}
})();