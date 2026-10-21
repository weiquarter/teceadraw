// 產生壓測用的假人員名單 CSV（欄位順序同「人員」分頁），貼進 Sheet 用，也給 stress-test.mjs 讀工號
// 用法：node tools/make-test-people.mjs [人數=600] > test-people.csv

const n = Number(process.argv[2] || 600);
const units = ['業務部', '行政部', '人資處', '財務部', '資訊處系統開發部第二課', '教保一組', '教保二組',
  '總務組', '企劃部', '研發中心', '客服部', '法務室', '稽核室', '公關部', '品保部'];
const surnames = '陳林黃張李王吳劉蔡楊許鄭謝郭洪曾邱廖賴周徐蘇葉莊呂江何蕭羅高'.split('');
const given = ['怡君', '雅婷', '家豪', '志明', '淑芬', '美玲', '建宏', '俊傑', '佳穎', '宗翰', '冠宇', '欣怡',
  '承恩', '子晴', '柏翰', '詩涵', '宜蓁', '彥廷', '品妤', '信宏', '惠如', '明哲', '雅雯', '思妤'];

const rows = [['縣市', '部門', '單位', '工號', '姓名', '可否摸彩', '登錄狀態']];
for (let i = 0; i < n; i++) {
  const name = surnames[(i * 7) % surnames.length] + given[(i * 11) % given.length];
  rows.push(['台北市', '測試部門', units[i % units.length], 'T' + String(i + 1).padStart(5, '0'), name, i % 20 === 19 ? '否' : '是', '既有']);
}
process.stdout.write('﻿' + rows.map(r => r.join(',')).join('\n') + '\n');
