(() => {
  // ========== 所有原有代码放入此 IIFE ==========

  let allData = [];
  let filteredData = [];
  let selectedCountry = null;
  let currentRankType = 'total';

  // 中东欧16国地区划分
  const regionMapping = {
      "波兰": "中欧",
      "捷克": "中欧",
      "匈牙利": "中欧",
      "斯洛伐克": "中欧",
      "斯洛文尼亚": "中欧",
      "爱沙尼亚": "东欧",
      "拉脱维亚": "东欧",
      "立陶宛": "东欧",
      "罗马尼亚": "东欧",
      "保加利亚": "东欧",
      "塞尔维亚": "巴尔干",
      "克罗地亚": "巴尔干",
      "希腊": "巴尔干",
      "马其顿": "巴尔干",
      "黑山": "巴尔干",
      "波黑": "巴尔干",
      "阿尔巴尼亚": "巴尔干"
  };

  // 颜色映射
  const colorScale = d3.scaleOrdinal()
      .domain(["中欧", "东欧", "巴尔干"])
      .range(["#5a7ca8", "#8898aa", "#7bc99c"]);

  // ★ 创建一个专属于本模块的 tooltip（解决冲突）
  const tooltip = d3.select('body').append('div')
      .attr('class', 'tooltip')
      .style('opacity', 0)
      .style('position', 'absolute')
      .style('pointer-events', 'none')
      .style('z-index', 9999);

  // 初始化
  document.addEventListener('DOMContentLoaded', () => {
      loadExcelData();
      setupEventListeners();
  });

  // 加载Excel数据
  function loadExcelData() {
      fetch('../data/1-中东欧-时间段-总量.xlsx')
          .then(response => response.arrayBuffer())
          .then(buffer => {
              const workbook = XLSX.read(buffer, { type: 'array' });
              const sheetName = workbook.SheetNames[0];
              const worksheet = workbook.Sheets[sheetName];
              const jsonData = XLSX.utils.sheet_to_json(worksheet);
              
              processData(jsonData);
              document.getElementById('loading').style.display = 'none';
              updateCharts();
          })
          .catch(error => {
              console.error('数据加载失败:', error);
              document.getElementById('loading').textContent = '数据加载失败，请检查Excel文件是否存在';
          });
  }

  // 数据处理
  function processData(rawData) {
      const countryGroups = d3.group(rawData, d => d['国家中文名']);
      
      allData = Array.from(countryGroups, ([country, values]) => {
          const data2011_2015 = values.find(v => v['时间段'] === '2011-2015');
          const data2016_2020 = values.find(v => v['时间段'] === '2016-2020');
          
          const count2011_2015 = data2011_2015 ? data2011_2015['合作Wos论文数'] : 0;
          const count2016_2020 = data2016_2020 ? data2016_2020['合作Wos论文数'] : 0;
          const total = count2011_2015 + count2016_2020;
          const growthRate = count2011_2015 > 0 
              ? ((count2016_2020 - count2011_2015) / count2011_2015 * 100).toFixed(1)
              : '—';
          
          return {
              country,
              region: regionMapping[country] || '其他',
              count2011_2015,
              count2016_2020,
              total,
              growthRate: parseFloat(growthRate) || 0
          };
      });
      
      allData = allData.filter(d => d.count2011_2015 > 0 && d.count2016_2020 > 0);
  }

  // 事件监听器
  function setupEventListeners() {
      document.getElementById('rankType').addEventListener('change', (e) => {
          currentRankType = e.target.value;
          updateCharts();
      });
      
      document.querySelectorAll('.control-group input[type="checkbox"]').forEach(checkbox => {
          checkbox.addEventListener('change', () => updateCharts());
      });
      
      document.getElementById('resetBtn').addEventListener('click', () => {
          selectedCountry = null;
          document.querySelectorAll('.control-group input[type="checkbox"]').forEach(cb => cb.checked = true);
          document.getElementById('rankType').value = 'total';
          currentRankType = 'total';
          updateCharts();
      });
  }

  function getFilteredData() {
      const selectedRegions = Array.from(
          document.querySelectorAll('.control-group input[type="checkbox"]:checked')
      ).map(cb => cb.value);
      return allData.filter(d => selectedRegions.includes(d.region));
  }

  function updateCharts() {
      filteredData = getFilteredData();
      drawBubbleChart();
      drawBarChart();
  }

  function drawBubbleChart() {
      const container = document.getElementById('bubbleChart');
      container.innerHTML = '';
      
      const margin = { top: 20, right: 30, bottom: 60, left: 70 };
      const width = container.clientWidth - margin.left - margin.right;
      const height = container.clientHeight - margin.top - margin.bottom;
      
      const svg = d3.select('#bubbleChart')
          .append('svg')
          .attr('width', width + margin.left + margin.right)
          .attr('height', height + margin.top + margin.bottom)
          .append('g')
          .attr('transform', `translate(${margin.left},${margin.top})`);
      
      const maxX = d3.max(filteredData, d => d.count2011_2015) * 1.15;
      const maxY = d3.max(filteredData, d => d.count2016_2020) * 1.15;
      const maxTotal = d3.max(filteredData, d => d.total);
      
      const xScale = d3.scaleLinear().domain([0, maxX]).range([0, width]);
      const yScale = d3.scaleLinear().domain([0, maxY]).range([height, 0]);
      const sizeScale = d3.scaleSqrt().domain([0, maxTotal]).range([5, 26]);
      
      // 网格、坐标轴、等增长线（略，与原代码完全相同）
      svg.append('g').attr('class', 'grid')
          .call(d3.axisLeft(yScale).ticks(8).tickSize(-width).tickFormat(''));
      svg.append('g').attr('class', 'axis x-axis')
          .attr('transform', `translate(0,${height})`)
          .call(d3.axisBottom(xScale).ticks(8));
      svg.append('g').attr('class', 'axis y-axis')
          .call(d3.axisLeft(yScale).ticks(8));
      svg.append('text').attr('class', 'axis-label')
          .attr('x', width / 2).attr('y', height + 40)
          .style('text-anchor', 'middle').text('2011-2015年发文量');
      svg.append('text').attr('class', 'axis-label')
          .attr('transform', 'rotate(-90)')
          .attr('x', -height / 2).attr('y', -50)
          .style('text-anchor', 'middle').text('2016-2020年发文量');
      
      const lineEnd = Math.min(maxX, maxY);
      svg.append('line').attr('class', 'equal-growth-line')
          .attr('x1', xScale(0)).attr('y1', yScale(0))
          .attr('x2', xScale(lineEnd)).attr('y2', yScale(lineEnd));
      svg.append('text').attr('x', xScale(lineEnd * 0.85))
          .attr('y', yScale(lineEnd * 0.85) - 10)
          .attr('fill', '#999').attr('font-size', '16px').text('等增长线');
      
      // 气泡（使用新的 tooltip 函数）
      svg.selectAll('.bubble').data(filteredData).enter()
          .append('circle').attr('class', 'bubble')
          .attr('cx', d => xScale(d.count2011_2015))
          .attr('cy', d => yScale(d.count2016_2020))
          .attr('r', d => sizeScale(d.total))
          .attr('fill', d => colorScale(d.region))
          .classed('selected', d => d.country === selectedCountry)
          .classed('dimmed', d => selectedCountry && d.country !== selectedCountry)
          .on('mouseover', showTooltip)
          .on('mousemove', moveTooltip)
          .on('mouseout', hideTooltip)
          .on('click', handleBubbleClick);
      
      const topCountries = [...filteredData].sort((a, b) => b.total - a.total).slice(0, 8);
      svg.selectAll('.country-label').data(topCountries).enter()
          .append('text').attr('class', 'country-label')
          .attr('x', d => xScale(d.count2011_2015))
          .attr('y', d => yScale(d.count2016_2020) - sizeScale(d.total) - 6)
          .attr('text-anchor', 'middle').text(d => d.country);
  }

  function drawBarChart() {
      // 与原代码完全相同（略）
      const container = document.getElementById('barChart');
      container.innerHTML = '';
      const margin = { top: 20, right: 30, bottom: 30, left: 100 };
      const width = container.clientWidth - margin.left - margin.right;
      const barHeight = 26;
      const height = filteredData.length * barHeight + 20;
      const svg = d3.select('#barChart').append('svg')
          .attr('width', width + margin.left + margin.right)
          .attr('height', height + margin.top + margin.bottom)
          .append('g').attr('transform', `translate(${margin.left},${margin.top})`);
      
      const sortedData = [...filteredData].sort((a, b) => {
          return currentRankType === 'total' ? b.total - a.total : b.growthRate - a.growthRate;
      });
      if (selectedCountry) {
          const idx = sortedData.findIndex(d => d.country === selectedCountry);
          if (idx > 0) {
              const [item] = sortedData.splice(idx, 1);
              sortedData.unshift(item);
          }
      }
      const maxValue = currentRankType === 'total'
          ? d3.max(sortedData, d => d.total)
          : d3.max(sortedData, d => d.growthRate);
      const xScale = d3.scaleLinear().domain([0, maxValue * 1.15]).range([0, width]);
      const yScale = d3.scaleBand()
          .domain(sortedData.map(d => d.country))
          .range([0, sortedData.length * barHeight])
          .padding(0.25);
      
      svg.append('g').attr('class', 'axis x-axis')
          .attr('transform', `translate(0,${sortedData.length * barHeight})`)
          .call(d3.axisBottom(xScale).ticks(5));
      svg.append('g').attr('class', 'axis y-axis').call(d3.axisLeft(yScale));
      
      svg.selectAll('.bar').data(sortedData).enter()
          .append('rect').attr('class', 'bar')
          .attr('x', 0).attr('y', d => yScale(d.country))
          .attr('width', d => xScale(currentRankType === 'total' ? d.total : d.growthRate))
          .attr('height', yScale.bandwidth())
          .attr('fill', d => colorScale(d.region))
          .classed('selected', d => d.country === selectedCountry)
          .classed('dimmed', d => selectedCountry && d.country !== selectedCountry)
          .on('click', handleBarClick);
      
      svg.selectAll('.value-label').data(sortedData).enter()
          .append('text').attr('class', 'value-label')
          .attr('x', d => xScale(currentRankType === 'total' ? d.total : d.growthRate) + 6)
          .attr('y', d => yScale(d.country) + yScale.bandwidth() / 2)
          .attr('dy', '0.35em')
          .text(d => currentRankType === 'total' ? `${d.total}篇` : `${d.growthRate}%`);
  }

  function handleBubbleClick(event, d) {
      event.stopPropagation();
      selectedCountry = selectedCountry === d.country ? null : d.country;
      updateCharts();
  }

  function handleBarClick(event, d) {
      event.stopPropagation();
      selectedCountry = selectedCountry === d.country ? null : d.country;
      updateCharts();
  }

  // ★ 重写的 tooltip 函数（使用本模块私有的 tooltip 元素）
  function showTooltip(event, d) {
      tooltip.html(`
          <strong>${d.country}</strong><br>
          所属地区：${d.region}<br>
          2011-2015发文量：${d.count2011_2015}篇<br>
          2016-2020发文量：${d.count2016_2020}篇<br>
          总发文量：${d.total}篇<br>
          增长率：${d.growthRate}%
      `)
      .style('opacity', 1)
      .style('left', (event.pageX + 15) + 'px')
      .style('top', (event.pageY - 15) + 'px');
  }

  function moveTooltip(event) {
      tooltip.style('left', (event.pageX + 15) + 'px')
            .style('top', (event.pageY - 15) + 'px');
  }

  function hideTooltip() {
      tooltip.style('opacity', 0);
  }

  // 点击空白处取消选择
  document.addEventListener('click', (event) => {
      if (!event.target.closest('.bubble') && !event.target.closest('.bar')) {
          if (selectedCountry) {
              selectedCountry = null;
              updateCharts();
          }
      }
  });

  // ★ 注意：本模块不需要 window.xxx = xxx，因为无 onclick 需求
})();