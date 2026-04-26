#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
超人眼鏡 - 終極自動化觸發器
確保廣告自動化永不失效的備援方案
"""

import requests
import schedule
import time
import json
import logging
from datetime import datetime, timezone, timedelta

# 設置日誌
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('automation_trigger.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

class AutomationTrigger:
    def __init__(self):
        self.railway_base = 'https://yindan-system-production.up.railway.app'
        self.test_endpoint = f'{self.railway_base}/api/superman-glasses/test-automation'
        self.execute_endpoint = f'{self.railway_base}/api/superman-glasses/check-and-execute'
        self.manual_endpoint = f'{self.railway_base}/api/superman-glasses/auto-execute'
        
        self.last_success = None
        self.failure_count = 0
        self.max_failures = 3
        
        logging.info("🚀 超人眼鏡自動化觸發器已啟動")
    
    def test_system_health(self):
        """檢查系統健康度"""
        try:
            response = requests.get(self.test_endpoint, timeout=30)
            
            if response.status_code == 200:
                result = response.json()
                if result.get('ok') and result.get('ready_for_automation'):
                    logging.info("✅ 系統健康檢查通過 - 自動化就緒")
                    return True, result.get('results', {})
                else:
                    logging.warning("⚠️ 系統未就緒 - 需要修復")
                    return False, result.get('results', {})
            else:
                logging.error(f"❌ 健康檢查失敗 - HTTP {response.status_code}")
                return False, {}
                
        except Exception as e:
            logging.error(f"❌ 健康檢查異常: {e}")
            return False, {}
    
    def trigger_automation(self):
        """觸發自動化執行"""
        try:
            # 先檢查系統健康度
            is_healthy, health_data = self.test_system_health()
            
            if not is_healthy:
                logging.error("❌ 系統不健康，跳過執行")
                self.failure_count += 1
                return False
            
            # 執行自動化
            response = requests.get(self.execute_endpoint, timeout=60)
            
            if response.status_code == 200:
                result = response.json()
                
                if result.get('ok'):
                    execution_result = result.get('result', {})
                    
                    # 記錄執行結果
                    roas_adjustments = execution_result.get('roas_adjustments', 0)
                    budget_increases = execution_result.get('budget_increases', 0)
                    paused_ads = execution_result.get('paused_ads', 0)
                    restarted_ads = execution_result.get('restarted_ads', 0)
                    errors = execution_result.get('errors', [])
                    
                    logging.info(f"✅ 自動化執行成功:")
                    logging.info(f"   📊 ROAS調整: {roas_adjustments}筆")
                    logging.info(f"   💰 預算增加: {budget_increases}筆")
                    logging.info(f"   ⏸️ 廣告暫停: {paused_ads}筆")
                    logging.info(f"   ▶️ 廣告重啟: {restarted_ads}筆")
                    
                    if errors:
                        logging.warning(f"   ⚠️ 錯誤: {len(errors)}個")
                        for error in errors[:3]:  # 只顯示前3個錯誤
                            logging.warning(f"     - {error}")
                    
                    self.last_success = datetime.now()
                    self.failure_count = 0
                    
                    return True
                else:
                    logging.error(f"❌ 自動化執行失敗: {result.get('error', 'Unknown')}")
                    self.failure_count += 1
                    return False
            else:
                logging.error(f"❌ 自動化觸發失敗 - HTTP {response.status_code}")
                self.failure_count += 1
                return False
                
        except Exception as e:
            logging.error(f"❌ 自動化執行異常: {e}")
            self.failure_count += 1
            return False
    
    def emergency_trigger(self):
        """緊急觸發 - 當正常觸發失敗時使用"""
        try:
            logging.info("🚨 執行緊急觸發...")
            
            # 模擬手動觸發
            trigger_data = {
                "cookies": "",  # 從Railway獲取
                "force": True,
                "source": "emergency_trigger"
            }
            
            response = requests.post(
                self.manual_endpoint,
                json=trigger_data,
                timeout=60
            )
            
            if response.status_code == 200:
                result = response.json()
                if result.get('ok'):
                    logging.info("✅ 緊急觸發成功")
                    self.failure_count = 0
                    return True
            
            logging.error("❌ 緊急觸發失敗")
            return False
            
        except Exception as e:
            logging.error(f"❌ 緊急觸發異常: {e}")
            return False
    
    def hourly_check(self):
        """每小時檢查"""
        tw_tz = timezone(timedelta(hours=8))
        now_tw = datetime.now(tw_tz)
        current_hour = now_tw.hour
        
        # 只在工作時間執行 (9:00-18:00)
        if not (9 <= current_hour <= 18):
            logging.info(f"⏰ 非工作時間 ({current_hour}:00)，跳過執行")
            return
        
        logging.info(f"⏰ 開始執行定時任務 ({current_hour}:00)")
        
        # 嘗試正常觸發
        success = self.trigger_automation()
        
        # 如果失敗且失敗次數超過限制，使用緊急觸發
        if not success and self.failure_count >= self.max_failures:
            logging.warning(f"⚠️ 連續失敗 {self.failure_count} 次，啟動緊急觸發")
            self.emergency_trigger()
    
    def daily_report(self):
        """每日報告"""
        tw_tz = timezone(timedelta(hours=8))
        now_tw = datetime.now(tw_tz)
        
        logging.info("📊 每日自動化報告:")
        logging.info(f"   最後成功時間: {self.last_success or '無'}")
        logging.info(f"   連續失敗次數: {self.failure_count}")
        logging.info(f"   系統運行時間: {now_tw.strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 執行健康檢查
        is_healthy, health_data = self.test_system_health()
        if is_healthy:
            logging.info("✅ 系統健康狀態良好")
        else:
            logging.error("❌ 系統需要注意")
    
    def start_scheduler(self):
        """啟動排程器"""
        # 每小時執行
        schedule.every().hour.do(self.hourly_check)
        
        # 每天8點執行報告
        schedule.every().day.at("08:00").do(self.daily_report)
        
        # 立即執行一次健康檢查
        self.test_system_health()
        
        logging.info("⚙️ 排程器已啟動 - 每小時觸發自動化")
        logging.info("📅 工作時間: 09:00-18:00")
        logging.info("📊 每日報告: 08:00")
        
        # 主循環
        while True:
            try:
                schedule.run_pending()
                time.sleep(60)  # 每分鐘檢查一次
            except KeyboardInterrupt:
                logging.info("👋 觸發器已停止")
                break
            except Exception as e:
                logging.error(f"⚠️ 排程器錯誤: {e}")
                time.sleep(300)  # 錯誤時等5分鐘

if __name__ == "__main__":
    print("🚀 超人眼鏡 - 終極自動化觸發器")
    print("=" * 50)
    print("這個程式會確保廣告自動化永不失效")
    print("工作時間: 每小時觸發一次 (09:00-18:00)")
    print("緊急備援: 連續失敗時自動切換")
    print("日誌檔案: automation_trigger.log")
    print("=" * 50)
    
    try:
        trigger = AutomationTrigger()
        trigger.start_scheduler()
    except Exception as e:
        logging.error(f"❌ 啟動失敗: {e}")
