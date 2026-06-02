// 共享配色：供各模块复用，避免多处重复定义
const SHARED_PRIMARY_COLORS = [
    "#5a7ca8","#7bc99c","#c68642","#8d7ab8",
    "#66a6a1","#b8a15a","#9aa6b2","#d08b8b",
    "#7f9c6e","#9a88a4","#6d8fb3","#c0a27a",
    "#8bb8c8","#a66f6f","#6e8b5f","#8f7c5d"
];

const SHARED_COUNTRY_COLORS = [
    "#5a7ca8","#c68642","#7bc99c","#8d7ab8",
    "#66a6a1","#b8a15a","#9aa6b2","#d08b8b",
    "#7f9c6e","#9a88a4","#6d8fb3","#c0a27a",
    "#8bb8c8","#a66f6f","#6e8b5f","#8f7c5d",
    "#4f7f82","#a28c4f","#7a8fa8","#b0747a",
    "#5d7a9a","#ad8d72","#6e9e8e","#9b7c92"
];

// 若需按名字查找颜色，可在此维护默认顺序（常见中东欧16国顺序）
const SHARED_COUNTRY_NAMES = [
  '波兰','捷克','希腊','匈牙利','罗马尼亚','塞尔维亚','斯洛文尼亚','斯洛伐克',
  '克罗地亚','保加利亚','爱沙尼亚','拉脱维亚','马其顿','波黑','阿尔巴尼亚','黑山'
];

function getSharedCountryColor(name) {
  const idx = SHARED_COUNTRY_NAMES.indexOf(name);
  if (idx >= 0) return SHARED_COUNTRY_COLORS[idx % SHARED_COUNTRY_COLORS.length];
  // fallback: hash name to pick a color deterministically
  let h = 0;
  for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) | 0;
  return SHARED_COUNTRY_COLORS[Math.abs(h) % SHARED_COUNTRY_COLORS.length];
}

// 导出变量到全局作用域（简单方式，页面以 <script> 引入）
window.SHARED_PRIMARY_COLORS = SHARED_PRIMARY_COLORS;
window.SHARED_COUNTRY_COLORS = SHARED_COUNTRY_COLORS;
window.SHARED_COUNTRY_NAMES = SHARED_COUNTRY_NAMES;
window.getSharedCountryColor = getSharedCountryColor;
