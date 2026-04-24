// 🔍 廣告自動化診斷腳本 - 在超人特工倉Console執行
// =====================================================

console.log('=== 🤖 廣告自動化診斷開始 ===');

// 1. 檢查最近24小時執行記錄
const checkRecentExecutions = () => {
  const pageText = document.body.textContent;
  const now = new Date();
  const last24h = new Date(now - 24*60*60*1000);
  
  // 檢查執行時間記錄
  const timePattern = /(\d{4}\/\d{1,2}\/\d{1,2})\s+(\d{1,2}:\d{2}:\d{2})/g;
  const allTimes = [...pageText.matchAll(timePattern)];
  
  const recentExecutions = allTimes.filter(match => {
    const execTime = new Date(match[0]);
    return execTime > last24h;
  });
  
  console.log(`📊 最近24小時執行次數: ${recentExecutions.length}`);
  if (recentExecutions.length === 0) {
    console.log('❌ 警告：最近24小時沒有自動執行記錄！');
  }
  
  return recentExecutions;
};

// 2. 檢查ROAS異常
const checkROASAlerts = () => {
  const pageText = document.body.textContent;
  
  // 找低ROAS記錄
  const lowROAS = pageText.match(/ROAS[:\s]*[0-2]\.\d+/g);
  console.log(`🔴 低ROAS記錄 (≤2.0): ${lowROAS?.length || 0}`);
  
  if (lowROAS && lowROAS.length > 0) {
    console.log('發現低ROAS:', lowROAS.slice(0, 3));
    console.log('⚠️ 這些應該會觸發自動調整，但可能沒有執行');
  }
};

// 3. 檢查執行間隔異常
const checkExecutionInterval = () => {
  const pageText = document.body.textContent;
  const timePattern = /(\d{4}\/\d{1,2}\/\d{1,2})\s+(\d{1,2}:\d{2}:\d{2})/g;
  const allTimes = [...pageText.matchAll(timePattern)];
  
  if (allTimes.length >= 2) {
    const intervals = [];
    for (let i = 0; i < Math.min(5, allTimes.length - 1); i++) {
      const time1 = new Date(allTimes[i][0]);
      const time2 = new Date(allTimes[i + 1][0]);
      const interval = Math.abs(time1 - time2) / (1000 * 60); // 分鐘
      intervals.push(interval);
    }
    
    console.log('⏱️ 執行間隔分析 (分鐘):', intervals.map(i => Math.round(i)));
    
    const avgInterval = intervals.reduce((a, b) => a + b, 0) / intervals.length;
    console.log(`平均執行間隔: ${Math.round(avgInterval)} 分鐘`);
    
    if (avgInterval < 5) {
      console.log('⚠️ 警告：執行間隔太短，可能是排程器問題');
    } else if (avgInterval > 120) {
      console.log('⚠️ 警告：執行間隔太長，可能自動化已停止');
    }
  }
};

// 4. 檢查錯誤訊息
const checkErrors = () => {
  const pageText = document.body.textContent;
  const errors = pageText.match(/(錯誤|異常|失敗|Error|timeout|expired)/gi);
  console.log(`🚨 錯誤/異常記錄: ${errors?.length || 0}`);
  
  if (errors && errors.length > 0) {
    console.log('錯誤關鍵字:', [...new Set(errors)]);
  }
};

// 執行診斷
checkRecentExecutions();
checkROASAlerts(); 
checkExecutionInterval();
checkErrors();

console.log('\n💡 建議檢查項目:');
console.log('1. 檢查 Cookie 是否過期 (BigSeller登入狀態)');
console.log('2. 檢查排程器設定 (應該是定時執行，不是每分鐘)');
console.log('3. 檢查 ROAS + 預算利用率條件是否太嚴格');
console.log('4. 檢查 Railway 服務是否正常運行');

console.log('\n✅ 廣告自動化診斷完成');
