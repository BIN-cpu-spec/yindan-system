#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
超人眼鏡自動化 - 部署後立即驗證
確保部署後系統立即可用
"""

import requests
import json
import time
from datetime import datetime

class DeploymentValidator:
    def __init__(self):
        self.railway_base = 'https://yindan-system-production.up.railway.app'
        self.tests = []
        
    def log(self, message, status="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        symbols = {"INFO": "ℹ️", "OK": "✅", "ERROR": "❌", "WARN": "⚠️"}
        print(f"[{timestamp}] {symbols.get(status, 'ℹ️')} {message}")
    
    def test_basic_connection(self):
        """測試基本連接"""
        self.log("測試Railway基本連接...")
        try:
            response = requests.get(f'{self.railway_base}/', timeout=10)
            if response.status_code == 200:
                self.log("Railway主頁連接正常", "OK")
                return True
            else:
                self.log(f"Railway連接異常: {response.status_code}", "ERROR")
                return False
        except Exception as e:
            self.log(f"Railway連接失敗: {e}", "ERROR")
            return False
    
    def test_automation_api(self):
        """測試自動化API端點"""
        self.log("測試自動化API端點...")
        
        endpoints = [
            ('/api/superman-glasses/test-automation', 'GET', '系統檢測'),
            ('/api/superman-glasses/scheduler-status', 'GET', '排程狀態'),
            ('/api/superman-glasses/save-cookie', 'POST', 'Cookie保存'),
        ]
        
        all_passed = True
        
        for endpoint, method, name in endpoints:
            try:
                url = f'{self.railway_base}{endpoint}'
                
                if method == 'GET':
                    response = requests.get(url, timeout=10)
                else:
                    # POST測試用空數據
                    response = requests.post(url, 
                        json={"test": True}, 
                        headers={"Content-Type": "application/json"},
                        timeout=10
                    )
                
                if response.status_code in [200, 400]:  # 400也算正常（參數錯誤但端點存在）
                    self.log(f"{name} API - 端點存在", "OK")
                else:
                    self.log(f"{name} API - 端點異常: {response.status_code}", "WARN")
                    all_passed = False
                    
            except Exception as e:
                self.log(f"{name} API - 測試失敗: {e}", "ERROR")
                all_passed = False
        
        return all_passed
    
    def test_system_readiness(self):
        """測試系統就緒狀態"""
        self.log("檢查系統自動化就緒狀態...")
        try:
            response = requests.get(f'{self.railway_base}/api/superman-glasses/test-automation', timeout=15)
            
            if response.status_code == 200:
                result = response.json()
                
                if result.get('ok'):
                    results = result.get('results', {})
                    
                    self.log("系統檢測API回應正常", "OK")
                    self.log(f"Cookie狀態: {results.get('cookie_status', 'unknown')}", 
                            "OK" if results.get('cookie_status') == 'available' else "WARN")
                    self.log(f"BigSeller API: {results.get('bigseller_api_status', 'unknown')}", 
                            "OK" if results.get('bigseller_api_status') == 'working' else "WARN")
                    self.log(f"自動化就緒: {results.get('automation_ready', False)}", 
                            "OK" if results.get('automation_ready') else "WARN")
                    
                    if results.get('error_details'):
                        for error in results.get('error_details', []):
                            self.log(f"問題: {error}", "WARN")
                    
                    return results.get('automation_ready', False)
                else:
                    self.log("系統檢測回應異常", "ERROR")
                    return False
            else:
                self.log(f"系統檢測API錯誤: {response.status_code}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"系統檢測失敗: {e}", "ERROR")
            return False
    
    def simulate_cookie_sync(self):
        """模擬Cookie同步（如果需要）"""
        self.log("檢查是否需要Cookie同步...")
        
        # 這裡你可以手動提供Cookie進行測試
        test_cookie = input("請提供BigSeller Cookie進行測試（可留空跳過）: ").strip()
        
        if not test_cookie:
            self.log("跳過Cookie同步測試", "WARN")
            return True
        
        try:
            sync_data = {
                "cookie": test_cookie,
                "source": "DEPLOYMENT-TEST",
                "version": "test"
            }
            
            response = requests.post(
                f'{self.railway_base}/api/superman-glasses/save-cookie',
                json=sync_data,
                headers={"Content-Type": "application/json"},
                timeout=15
            )
            
            if response.status_code == 200:
                result = response.json()
                if result.get('ok'):
                    self.log("Cookie同步測試成功", "OK")
                    return True
                else:
                    self.log(f"Cookie同步失敗: {result.get('msg', 'Unknown')}", "ERROR")
                    return False
            else:
                self.log(f"Cookie同步HTTP錯誤: {response.status_code}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"Cookie同步測試異常: {e}", "ERROR")
            return False
    
    def test_manual_trigger(self):
        """測試手動觸發"""
        self.log("測試手動觸發自動化...")
        
        try:
            response = requests.post(
                f'{self.railway_base}/api/superman-glasses/auto-execute',
                json={"cookies": "", "test_mode": True},
                headers={"Content-Type": "application/json"},
                timeout=30
            )
            
            if response.status_code == 200:
                result = response.json()
                if result.get('ok'):
                    self.log("手動觸發測試成功", "OK")
                    return True
                else:
                    self.log(f"手動觸發失敗: {result.get('error', 'Unknown')}", "ERROR")
                    return False
            else:
                self.log(f"手動觸發HTTP錯誤: {response.status_code}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"手動觸發測試異常: {e}", "ERROR")
            return False
    
    def check_scheduler_status(self):
        """檢查排程器狀態"""
        self.log("檢查定時任務排程狀態...")
        
        try:
            response = requests.get(f'{self.railway_base}/api/superman-glasses/scheduler-status', timeout=10)
            
            if response.status_code == 200:
                result = response.json()
                if result.get('ok'):
                    status = result.get('status', {})
                    self.log("排程器狀態檢查成功", "OK")
                    
                    if status.get('healthy', False):
                        self.log("排程器健康狀態良好", "OK")
                    else:
                        issues = status.get('issues', [])
                        self.log("排程器發現問題:", "WARN")
                        for issue in issues:
                            self.log(f"  - {issue}", "WARN")
                    
                    return status.get('healthy', False)
                else:
                    self.log("排程器狀態回應異常", "ERROR")
                    return False
            else:
                self.log(f"排程器狀態API錯誤: {response.status_code}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"排程器狀態檢查異常: {e}", "ERROR")
            return False
    
    def run_full_validation(self):
        """執行完整驗證"""
        print("🚀 超人眼鏡自動化 - 部署後驗證")
        print("=" * 60)
        
        results = {
            "basic_connection": self.test_basic_connection(),
            "automation_api": self.test_automation_api(),
            "system_readiness": self.test_system_readiness(),
            "cookie_sync": self.simulate_cookie_sync(),
            "manual_trigger": self.test_manual_trigger(),
            "scheduler_status": self.check_scheduler_status()
        }
        
        print("\n" + "=" * 60)
        print("📊 驗證結果摘要:")
        print("=" * 60)
        
        passed = 0
        total = len(results)
        
        for test_name, result in results.items():
            status = "✅ 通過" if result else "❌ 失敗"
            print(f"  {test_name:20} : {status}")
            if result:
                passed += 1
        
        print("=" * 60)
        print(f"總體結果: {passed}/{total} 通過")
        
        if passed == total:
            print("🎉 恭喜！自動化系統完全就緒")
            print("\n下一步:")
            print("1. 系統已開始每小時自動執行")
            print("2. 可以不再依賴Extension")
            print("3. 建議運行backup_trigger.py作為備援")
            
        elif passed >= 4:
            print("⚠️ 系統基本可用，但有小問題需要修復")
            print("\n建議:")
            print("1. 檢查失敗的測試項目")
            print("2. 手動同步Cookie")
            print("3. 監控第一次自動執行")
            
        else:
            print("🚨 系統需要修復後才能使用")
            print("\n急需修復:")
            print("1. 檢查Railway部署是否成功")
            print("2. 檢查app.py是否正確上傳")
            print("3. 檢查API端點是否可用")
        
        return passed >= 4

if __name__ == "__main__":
    validator = DeploymentValidator()
    validator.run_full_validation()
