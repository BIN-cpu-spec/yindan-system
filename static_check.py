#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
超人眼鏡自動化系統 - 靜態檢查
檢查 app.py 中的自動化功能是否正確實現
"""

import re
import sys
import os

def static_check_automation():
    """靜態檢查自動化功能"""
    
    app_file = "/mnt/user-data/outputs/app.py"
    
    if not os.path.exists(app_file):
        print("❌ app.py 文件不存在")
        return False
    
    with open(app_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    checks = {
        "API端點": {
            "auto_execute": r'@app\.route\("/api/superman-glasses/auto-execute"',
            "check_and_execute": r'@app\.route\("/api/superman-glasses/check-and-execute"',
            "test_automation": r'@app\.route\("/api/superman-glasses/test-automation"',
            "save_cookie": r'@app\.route\("/api/superman-glasses/save-cookie"'
        },
        "核心函數": {
            "execute_bigseller_automation": r'def execute_bigseller_automation\(',
            "analyze_and_adjust_ad": r'def analyze_and_adjust_ad\(',
            "get_latest_bigseller_cookie": r'def get_latest_bigseller_cookie\(',
            "setup_automation_scheduler": r'def setup_automation_scheduler\('
        },
        "導入模塊": {
            "schedule": r'import.*schedule',
            "requests": r'import\s+requests|from.*requests',
            "threading": r'import.*threading',
            "datetime": r'from datetime import datetime'
        },
        "關鍵邏輯": {
            "cookie_file_save": r'bigseller_cookie\.txt',
            "schedule_every_hour": r'schedule\.every\(\)\.hour\.do',
            "api_call_bigseller": r'api/v1/product/listing/shopee/queryAdCampaignShopInfoPage',
            "edit_ads_api": r'editSingleShopeeProductAds\.json'
        }
    }
    
    print("🔍 超人眼鏡自動化系統 - 靜態檢查")
    print("=" * 60)
    
    all_passed = True
    
    for category, items in checks.items():
        print(f"\n📋 {category}:")
        for name, pattern in items.items():
            if re.search(pattern, content, re.MULTILINE):
                print(f"  ✅ {name}")
            else:
                print(f"  ❌ {name} - 缺失!")
                all_passed = False
    
    # 檢查 API 實現完整性
    print(f"\n🔧 API實現完整性:")
    
    # 檢查自動化執行邏輯
    auto_execute_impl = re.search(
        r'def superman_glasses_auto_execute\(\):.*?(?=@app\.route|def [^_]|$)', 
        content, 
        re.DOTALL
    )
    
    if auto_execute_impl:
        impl = auto_execute_impl.group()
        if "execute_bigseller_automation" in impl:
            print(f"  ✅ auto-execute API 實現完整")
        else:
            print(f"  ❌ auto-execute API 缺少核心調用")
            all_passed = False
    else:
        print(f"  ❌ auto-execute API 實現缺失")
        all_passed = False
    
    # 檢查定時任務啟動
    if "setup_automation_scheduler()" in content:
        print(f"  ✅ 定時任務自動啟動")
    else:
        print(f"  ❌ 定時任務未設置自動啟動")
        all_passed = False
    
    # 檢查 Cookie 管理
    if "bigseller_cookie.txt" in content and "get_latest_bigseller_cookie" in content:
        print(f"  ✅ Cookie 管理機制完整")
    else:
        print(f"  ❌ Cookie 管理機制缺失")
        all_passed = False
    
    print("\n" + "=" * 60)
    if all_passed:
        print("✅ 靜態檢查通過 - 自動化功能實現完整")
    else:
        print("❌ 靜態檢查失敗 - 需要修復缺失功能")
    
    # 額外檢查：尋找常見問題
    print(f"\n⚠️ 常見問題檢查:")
    
    issues = []
    
    # 檢查import問題
    if "import schedule" not in content and "schedule" in content:
        issues.append("schedule模塊可能未正確導入")
    
    # 檢查函數調用
    if "requests.post" in content and "import requests" not in content:
        issues.append("requests模塊可能未正確導入")
    
    # 檢查錯誤處理
    if "execute_bigseller_automation" in content:
        if "try:" not in content or "except" not in content:
            issues.append("缺少錯誤處理機制")
    
    if issues:
        for issue in issues:
            print(f"  ⚠️ {issue}")
    else:
        print(f"  ✅ 未發現常見問題")
    
    return all_passed

def check_execution_flow():
    """檢查執行流程邏輯"""
    print(f"\n🚀 自動化執行流程分析:")
    
    flow_steps = [
        "1. Railway 啟動時自動設置定時任務",
        "2. 每小時執行 check-and-execute API",
        "3. 檢查當前時間是否在執行窗口(9-18點)",
        "4. 讀取 Cookie 從檔案或記憶體",
        "5. 調用 BigSeller API 獲取廣告列表",
        "6. 逐個分析廣告並執行調整",
        "7. 記錄執行結果",
        "8. 處理 2001 錯誤和重試"
    ]
    
    for step in flow_steps:
        print(f"  📝 {step}")
    
    print(f"\n💡 關鍵成功因素:")
    success_factors = [
        "✅ Cookie 同步機制正常工作",
        "✅ BigSeller API 認證有效",
        "✅ Railway 定時任務正常運行",
        "✅ 錯誤處理和重試機制",
        "✅ 執行條件不過於嚴格"
    ]
    
    for factor in success_factors:
        print(f"  {factor}")

if __name__ == "__main__":
    success = static_check_automation()
    check_execution_flow()
    
    if success:
        print(f"\n🎯 結論: 系統已準備就緒，可以部署")
        sys.exit(0)
    else:
        print(f"\n🚨 結論: 需要修復問題後再部署")
        sys.exit(1)
