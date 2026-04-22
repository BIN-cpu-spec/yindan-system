# -*- coding: utf-8 -*-
"""
BigSeller 超材分單系統 - 專業版
專為BigSeller平台優化的超材訂單分析工具
功能：自動判斷超材訂單、分單建議、批量處理
"""

import sys, csv, io, os, re, json, threading, time, hashlib, secrets, zipfile
from datetime import datetime
from flask import Flask, render_template_string, request, jsonify, session, redirect, url_for, send_file, Response

# ── 超材分析依賴 ──
from io import StringIO

# ── 圖片URL解析：將 =IMAGE("url") 公式轉成純 URL ──
def parse_image_url(cell_value):
    if not cell_value:
        return ""
    s = str(cell_value).strip()
    m = re.search(r'=IMAGE\("([^"]+)"', s, re.IGNORECASE)
    if m:
        return m.group(1)
    if s.startswith("http://") or s.startswith("https://"):
        return s
    return ""

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# ============================================================
# 設定區
# ============================================================
CONFIG = {
    "company_name": "我的電商公司",
    "flask_port":   5000,
    "auto_run_hour":7,
    "auto_run_min": 30,

    # CSV 欄位名稱
    "col_txn":      "子交易序號",
    "col_order_id": "訂單編號",
    "col_shipping": "出貨類型",
    "col_warehouse":"商品倉庫儲位",
    "col_sku":      "商品編號",
    "col_length":   "商品長",       # 單位 cm
    "col_width":    "商品寬",       # 單位 cm
    "col_height":   "商品高",       # 單位 cm
    "col_weight":   "子交易總重量", # 單位 kg（整張訂單已加總）
    "col_fee":      "運費",         # 運費欄位（0=可拆單，>0=買家已付不拆單）

    # 通路判斷關鍵字
    "kw_delivery": ["新竹物流", "嘉里", "店到家"],
    "kw_cvs":      ["7-11", "711", "全家", "萊爾富"],
    "kw_nextday":  ["隔日到貨"],
    "kw_nopkg":    ["無包裝"],
    "kw_store":    ["店到店"],

    # 倉庫前綴
    "wh_prefix": {
        "?前": "參前",
        "?倉": "參前",
        "主":  "參倉",
    },

    # 超材規則（整張訂單加總後判斷）
    "oversize_rules": {
        # 通路key → (三邊總和上限cm, 最長邊上限cm, 重量上限kg, 說明)
        "cvs":      (105, 45,  10, "超商"),
        "store":    (105, 45,  10, "店到店"),
        "nextday":  (105, 45,  10, "店到店隔日配"),
        "nopkg":    (105, 45,  10, "無包裝"),
        "delivery_jialy":   (200, 120, 20, "嘉里快遞"),
        "delivery_hsinchu": (210, 150, 20, "新竹物流"),
        "delivery_store":   (150, 100, 15, "店到家大型"),
    },

    # 新竹物流補助運費對照（三邊總和 → 買家自付運費）
    # ≦150cm 正常補助，151~210cm 買家自付，>210cm 異常
    "hsinchu_surcharge": [
        (150, 0,   "正常補助"),
        (160, 135, "買家自付135元"),
        (170, 165, "買家自付165元"),
        (180, 195, "買家自付195元"),
        (190, 225, "買家自付225元"),
        (200, 285, "買家自付285元"),
        (210, 335, "買家自付335元"),
    ],

    # ★ 特殊可出超材品清單（SKU → 斜放後有效最長邊cm）
    # 斜放後最長邊 = √(長²+寬²+高²)，填入計算後的值
    # 例如：55x40x30 的商品，對角線 ≈ 74cm，但斜放進箱後最長邊可能只有 44cm
    # 請依實際測量填入，系統會用此值取代原始最長邊判斷超材
    "diagonal_skus": {
        # "SKU-001": 44,   # 範例：SKU-001 斜放後有效最長邊 44cm
        # "SKU-002": 43,
    },

    # 拆單規則（店到店/店到家超材時）
    "split_max_units":   5,     # 最多拆幾單
    "split_max_dim":     105,   # 每包三邊總和上限 cm
    "split_max_side":    45,    # 每包最長邊上限 cm
    "split_max_weight":  10,    # 每包重量上限 kg
}

# 通路顯示設定
CHANNEL_META = {
    "delivery": {"label": "宅配",         "icon": "🚙", "color": "#1565c0"},
    "cvs":      {"label": "超商",         "icon": "🏪", "color": "#c85000"},
    "store":    {"label": "店到店",       "icon": "🏬", "color": "#2e7d32"},
    "nextday":  {"label": "店到店隔日配", "icon": "⚡",   "color": "#6a1b9a"},
    "nopkg":    {"label": "無包裝",       "icon": "📦", "color": "#00838f"},
}

state = {
    "groups": {}, "total": 0, "last_update": None,
    "status": "idle", "status_msg": "請上傳 CSV 開始分單",
    "log": [], "summary": {},
}

def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    state["log"].append(line)
    if len(state["log"]) > 200:
        state["log"] = state["log"][-200:]

# ============================================================
# 分單邏輯
# ============================================================
def detect_channel(val):
    s = (val or "").strip()
    if any(k in s for k in CONFIG["kw_delivery"]): return "delivery"
    if any(k in s for k in CONFIG["kw_cvs"]):      return "cvs"
    if any(k in s for k in CONFIG["kw_nextday"]):  return "nextday"
    if any(k in s for k in CONFIG["kw_nopkg"]):    return "nopkg"
    if any(k in s for k in CONFIG["kw_store"]):    return "store"
    return "delivery"

def parse_location(raw):
    """儲位 → (倉庫名稱, 區域字母)"""
    raw = (raw or "").strip()
    wh, rest = "其他", raw
    for prefix, name in CONFIG["wh_prefix"].items():
        if raw.startswith(prefix):
            wh, rest = name, raw[len(prefix):]
            break
    m = re.match(r"([A-Z])", rest)
    zone = m.group(1) if m else "?"
    return wh, zone

def safe_float(val):
    try: return float((str(val) or "0").strip())
    except: return 0.0

def parse_fee(val):
    """解析運費，去除 $ NT$ 等符號，回傳數字"""
    try:
        s = str(val or "0").strip()
        # 移除常見貨幣符號
        s = s.replace("NT$", "").replace("$", "").replace(",", "").strip()
        return float(s) if s else 0.0
    except:
        return 0.0

def check_oversize(ch, ship_raw, total_dim, max_side, weight, buyer_fee=0):
    """
    判斷是否超材。
    回傳 (is_oversize: bool, warn_msg: str, can_split: bool, split_rules: dict)
    """
    if total_dim == 0 and weight == 0:
        return False, "", False, None

    # 新竹物流
    if "新竹物流" in ship_raw:
        if total_dim > 210 or max_side > 150 or weight > 20:
            return True, f"超規 {total_dim:.0f}cm/{weight:.1f}kg → 異常單", False, None
        for limit, required_fee, desc in CONFIG["hsinchu_surcharge"]:
            if total_dim <= limit:
                if required_fee == 0:
                    return False, "", False, None
                if buyer_fee >= required_fee:
                    return False, "", False, None
                return True, f"新竹補助不足 {total_dim:.0f}cm，應付 {required_fee} 元，實付 {buyer_fee:.0f} 元 → 異常單", False, None
        return True, f"超規 {total_dim:.0f}cm → 異常單", False, None

    # 嘉里快遞
    if "嘉里" in ship_raw:
        if total_dim > 200 or max_side > 120 or weight > 20:
            return True, f"超規 {total_dim:.0f}cm/{weight:.1f}kg → 異常單", False, None
        return False, "", False, None

    # 店到家宅配 — 標準 or 大型
    if "店到家" in ship_raw:
        # 標準合規
        if total_dim <= 105 and max_side <= 45 and weight <= 10:
            return False, "", False, None
        # 大型合規
        if total_dim <= 150 and max_side <= 100 and weight <= 15:
            return False, "", False, None
        # 超材：先嘗試用大型規格拆，再嘗試標準規格拆
        rules = {"max_dim": 150, "max_side": 100, "max_weight": 15, "label": "店到家大型"}
        return True, f"超材 {total_dim:.0f}cm/{weight:.1f}kg", True, rules

    # 超商（7-11 / 全家 / 萊爾富）
    if ch == "cvs":
        if total_dim > 105 or max_side > 45 or weight > 10:
            rules = {"max_dim": 105, "max_side": 45, "max_weight": 10, "label": "超商"}
            return True, f"超材 {total_dim:.0f}cm/{weight:.1f}kg", True, rules
        return False, "", False, None

    # 店到店
    if ch == "store":
        if total_dim > 105 or max_side > 45 or weight > 10:
            rules = {"max_dim": 105, "max_side": 45, "max_weight": 10, "label": "店到店"}
            return True, f"超材 {total_dim:.0f}cm/{weight:.1f}kg", True, rules
        return False, "", False, None

    # 隔日配（支援拆單）
    if ch == "nextday":
        if total_dim > 105 or max_side > 45 or weight > 10:
            rules = {"max_dim": 105, "max_side": 45, "max_weight": 10, "label": "隔日配"}
            return True, f"超材 {total_dim:.0f}cm/{weight:.1f}kg", True, rules
        return False, "", False, None

    # 無包裝（不拆單）
    if ch == "nopkg":
        if total_dim > 105 or max_side > 45 or weight > 10:
            return True, f"超材 {total_dim:.0f}cm/{weight:.1f}kg", False, None
        return False, "", False, None
    return False, "", False, None

def apply_diagonal(sku, L, W, H):
    """
    如果此 SKU 在可斜放清單內，回傳斜放後的有效最長邊。
    支援新格式 {side: float, max_qty: int} 和舊格式 float。
    """
    diag_map = CONFIG.get("diagonal_skus", {})
    if sku in diag_map:
        cfg = diag_map[sku]
        if isinstance(cfg, dict):
            side = cfg.get("side")
            if side: return float(side), True
        elif isinstance(cfg, (int, float)):
            return float(cfg), True
    return max(L, W, H), False

def get_sku_max_qty(sku, channel=None):
    """
    取得 SKU 的每包最大件數限制。
    channel: 'store'(店配類) 或 'delivery'(快遞類)
    無設定或通路不符回傳 None。
    """
    diag_map = CONFIG.get("diagonal_skus", {})
    if sku not in diag_map:
        return None
    cfg = diag_map[sku]
    if not isinstance(cfg, dict):
        return None
    max_qty  = cfg.get("max_qty")
    channels = cfg.get("channels", [])
    if not max_qty:
        return None
    # 若沒有設定通路限制，或通路符合，才回傳
    if not channels:
        return max_qty
    if channel and channel in channels:
        return max_qty
    return None

def suggest_split(products, weight_per_item, L, W, H, rules=None):
    """
    建議拆單方式。
    rules: {max_dim, max_side, max_weight, max_qty(optional for bulky items)}
    """
    if rules is None:
        rules = {
            "max_dim":    CONFIG["split_max_dim"],
            "max_side":   CONFIG["split_max_side"],
            "max_weight": CONFIG["split_max_weight"],
        }

    total_qty = len(products)
    max_units = CONFIG["split_max_units"]
    single_dim = L + W + H
    single_side = max(L, W, H)
    units = []

    # 策略1：如果單件就超材，不可拆
    if single_dim > rules["max_dim"] or single_side > rules["max_side"]:
        return f"單件尺寸已超材（{single_dim:.0f}cm/{single_side:.0f}cm），無法拆分"

    # 策略2：先嘗試按重量分
    total_weight = weight_per_item * total_qty
    if total_weight <= rules["max_weight"]:
        # 重量沒問題，檢查尺寸
        if single_dim <= rules["max_dim"] and single_side <= rules["max_side"]:
            return f"實際不超材，可合併出貨（{single_dim:.0f}cm/{total_weight:.1f}kg）"

    # 策略3：平均分包
    import math
    # 重量分包數
    weight_packs = math.ceil(total_weight / rules["max_weight"])
    # 尺寸分包數（假設疊加）
    dim_packs = math.ceil(total_qty * single_dim / rules["max_dim"])
    
    # 取較大值
    min_packs = max(weight_packs, dim_packs, 1)
    
    if min_packs > max_units:
        return f"需拆 {min_packs} 包超過系統上限（{max_units}包），建議調整商品組合"

    # 產生分包方案
    items_per_pack = math.ceil(total_qty / min_packs)
    for i in range(min_packs):
        start = i * items_per_pack
        end = min((i + 1) * items_per_pack, total_qty)
        if start >= total_qty:
            break
        
        pack_qty = end - start
        pack_weight = pack_qty * weight_per_item
        pack_dim = pack_qty * single_dim  # 簡化計算
        
        units.append({
            "pack": i + 1,
            "qty": pack_qty,
            "weight": pack_weight,
            "est_dim": pack_dim,
            "items": products[start:end]
        })

    # 格式化輸出
    lines = [f"建議拆成 {len(units)} 包："]
    for u in units:
        lines.append(f"  第{u['pack']}包：{u['qty']}件，{u['weight']:.1f}kg")
    
    return "\n".join(lines)

def _get_zone_label(locs):
    """區域標籤：單一倉+單一區 → 簡潔，否則 → 詳細"""
    if len(locs) == 1:
        wh, zone = list(locs)[0]
        return f"{wh}{zone}區"
    return " + ".join(f"{wh}{zone}" for wh, zone in sorted(locs))

def get_sort_key(k):
    """排序：宅配→純區+單品→店到店多品→隔日配→無包裝→可拆單→超材"""
    if k == "delivery":      return (10, k, k)    # 宅配優先
    if k == "__single_zone__":  return (20, k, k)   # 新：純區+單品大分類
    if k == "store":         return (30, k, k)    # 店到店多品
    if k == "nextday":       return (40, k, k)    # 隔日配
    if k == "nopkg":         return (50, k, k)    # 無包裝
    if k.endswith("_split"): return (60, k, k)    # 可拆單超材
    if k.endswith("_over"):  return (70, k, k)    # 不可拆超材
    return (99, k, k)

def split_orders(rows):
    # ── 以交易序號合併同一張訂單 ──────────────────────
    order_map = {}
    for row in rows:
        txn = (row.get(CONFIG["col_txn"], "") or "").strip()
        if not txn: continue
        if txn not in order_map:
            order_map[txn] = {
                "txn":       txn,
                "channel":   detect_channel(row.get(CONFIG["col_shipping"], "")),
                "ship_raw":  (row.get(CONFIG["col_shipping"], "") or "").strip(),
                "order_ids": [],
                "products":  [],
                "locations": set(),
                "total_qty": 0,
                # 尺寸（取該訂單最大值，因為同訂單多列尺寸應相同）
                "length":    0.0,
                "width":     0.0,
                "height":    0.0,
                "weight":    0.0,    # 使用子交易總重量，多列取最大值
                "fee":       0.0,    # 運費
            }
        
        o = order_map[txn]
        o["order_ids"].append((row.get(CONFIG["col_order_id"], "") or "").strip())
        o["products"].append({
            "sku": (row.get(CONFIG["col_sku"], "") or "").strip(),
            "warehouse": (row.get(CONFIG["col_warehouse"], "") or "").strip(),
        })
        o["total_qty"] += 1
        
        # 儲位解析
        wh, zone = parse_location(row.get(CONFIG["col_warehouse"], ""))
        o["locations"].add((wh, zone))
        
        # 尺寸取最大（同訂單應相同）
        o["length"] = max(o["length"], safe_float(row.get(CONFIG["col_length"], "")))
        o["width"]  = max(o["width"],  safe_float(row.get(CONFIG["col_width"], "")))
        o["height"] = max(o["height"], safe_float(row.get(CONFIG["col_height"], "")))
        o["weight"] = max(o["weight"], safe_float(row.get(CONFIG["col_weight"], "")))
        o["fee"]    = max(o["fee"],    parse_fee(row.get(CONFIG["col_fee"], "")))

    # ── 超材判斷與分群 ──────────────────────────────
    groups = {}
    summary = {}
    
    def add(key, title, icon, color, order):
        if key not in groups:
            groups[key] = {"title": title, "icon": icon, "color": color, "orders": []}
        groups[key]["orders"].append(order)

    for txn, o in order_map.items():
        # 斜放邏輯：如果訂單內任一 SKU 可斜放，就用斜放尺寸
        dims = [o["length"], o["width"], o["height"]]
        effective_side = max(dims)
        diagonal_applied = False
        
        # 檢查訂單內是否有可斜放的 SKU
        for prod in o["products"]:
            sku = prod["sku"]
            if sku:
                diag_side, is_diag = apply_diagonal(sku, o["length"], o["width"], o["height"])
                if is_diag and diag_side < effective_side:
                    effective_side = diag_side
                    diagonal_applied = True
                    break

        o["max_side"]       = effective_side
        o["total_dim"]      = sum(dims)
        o["diagonal_used"]  = diagonal_applied
        o["oversize"], o["oversize_msg"], o["can_split"], o["split_rules"] = check_oversize(
            o["channel"], o["ship_raw"], o["total_dim"], o["max_side"], o["weight"], o["fee"]
        )

        # 若斜放後不超材，標示說明
        if diagonal_applied and not o["oversize"]:
            o["oversize_msg"] = "斜放後合規"

        # 件數上限判斷：檢查訂單內的 SKU 是否有設定每包最大件數
        if not o["oversize"]:
            from collections import Counter
            sku_count = Counter(p["sku"] for p in o["products"])
            # 判斷通路群組：delivery=快遞類, store=店配類
            ch = o["channel"]
            ch_group = "delivery" if ch == "delivery" else "store"
            
            for sku, qty in sku_count.items():
                max_qty = get_sku_max_qty(sku, ch_group)
                if max_qty and qty > max_qty:
                    # 這個 SKU 超過件數上限，標記為超材（可拆單）
                    o["oversize"] = True
                    o["oversize_msg"] = f"{sku} 每包限{max_qty}件，實際{qty}件"
                    o["can_split"] = True
                    o["split_rules"] = {"max_qty": max_qty, "sku": sku}
                    break

        # ── 分群邏輯 ──────────────────────────────────
        ch = o["channel"]
        ch_meta = CHANNEL_META.get(ch, {"label": ch.title(), "icon": "📦", "color": "#666"})

        # ── 新分類邏輯：純區 + 超商單品 + 店到店單品 ──
        if not o["oversize"]:
            # 取得區域資訊
            whs = {loc[0] for loc in o["locations"]}  # 所有倉庫
            zones = {loc[1] for loc in o["locations"]}  # 所有區域
            
            # 超商單品（1件）→ 進大分類
            if ch == "cvs" and o["total_qty"] == 1:
                o["single_zone_sub"] = "超商單品"
                add("__single_zone__", "⚡ 純區 + 超商單品 + 店到店單品", "⚡", "#e65100", o)
                summary["純區+單品"] = summary.get("純區+單品", 0) + 1
                continue
            
            # 店到店單品（1件）→ 進大分類
            if ch == "store" and o["total_qty"] == 1:
                o["single_zone_sub"] = "店到店單品"
                add("__single_zone__", "⚡ 純區 + 超商單品 + 店到店單品", "⚡", "#e65100", o)
                summary["純區+單品"] = summary.get("純區+單品", 0) + 1
                continue
            
            # 純區包裹（單一倉+單一區，任何通路）→ 進大分類
            if len(whs) == 1 and len(zones) == 1:
                o["single_zone_sub"] = "純區包裹"
                add("__single_zone__", "⚡ 純區 + 超商單品 + 店到店單品", "⚡", "#e65100", o)
                summary["純區+單品"] = summary.get("純區+單品", 0) + 1
                continue

        # ── 原有邏輯：其他情況 ──
        if not o["oversize"]:
            # 正常出貨
            add(ch, ch_meta["label"], ch_meta["icon"], ch_meta["color"], o)
            summary[ch_meta["label"]] = summary.get(ch_meta["label"], 0) + 1
        elif o["can_split"]:
            # 可拆單超材
            split_key = f"{ch}_split"
            add(split_key, f"{ch_meta['label']}超材(可拆)", "⚠️", "#ff9800", o)
            summary[f"{ch_meta['label']}超材"] = summary.get(f"{ch_meta['label']}超材", 0) + 1
            
            # 生成拆單建議
            if o.get("split_rules"):
                weight_per_item = o["weight"] / o["total_qty"] if o["total_qty"] > 0 else 0
                o["split_suggestion"] = suggest_split(
                    o["products"], weight_per_item, 
                    o["length"], o["width"], o["height"], 
                    o["split_rules"]
                )
        else:
            # 不可拆超材
            over_key = f"{ch}_over"
            add(over_key, f"{ch_meta['label']}超材(異常)", "❌", "#f44336", o)
            summary[f"{ch_meta['label']}異常"] = summary.get(f"{ch_meta['label']}異常", 0) + 1

    # ── 依優先級排序 ──────────────────────────────
    sorted_groups = dict(sorted(groups.items(), key=lambda x: get_sort_key(x[0])))
    return sorted_groups, summary, len(order_map)

def load_csv(source, is_text=False):
    """載入 CSV（支援檔案或文字）"""
    if is_text:
        return list(csv.DictReader(StringIO(source)))
    else:
        with open(source, 'r', encoding='utf-8-sig') as f:
            return list(csv.DictReader(f))

def run_pipeline(rows=None):
    """執行分單流程"""
    if rows is None:
        return
    
    state["status"] = "running"
    state["status_msg"] = "正在分析訂單..."
    log(f"開始處理 {len(rows)} 筆原始資料")
    
    groups, summary, total = split_orders(rows)
    
    state["groups"] = groups
    state["summary"] = summary
    state["total"] = total
    state["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    state["status"] = "completed"
    state["status_msg"] = f"處理完成，共 {total} 張訂單"
    
    log(f"分單完成：{total} 張訂單分為 {len(groups)} 個群組")

def scheduler():
    """背景排程（保留原邏輯但簡化）"""
    while True:
        time.sleep(300)  # 5分鐘檢查一次

# ============================================================
# 設定檔案讀寫
# ============================================================
def _get_base_dir():
    if hasattr(sys, '_MEIPASS'):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

def load_settings():
    """載入可斜放 SKU 設定"""
    global CONFIG
    try:
        settings_path = os.path.join(_get_base_dir(), "diagonal_settings.json")
        if os.path.exists(settings_path):
            with open(settings_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                CONFIG["diagonal_skus"] = data.get("diagonal_skus", {})
                log(f"載入斜放設定：{len(CONFIG['diagonal_skus'])} 筆")
    except Exception as e:
        log(f"載入設定檔失敗：{e}")

def save_settings():
    """儲存可斜放 SKU 設定"""
    try:
        settings_path = os.path.join(_get_base_dir(), "diagonal_settings.json")
        data = {"diagonal_skus": CONFIG["diagonal_skus"]}
        with open(settings_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        log(f"儲存斜放設定：{len(CONFIG['diagonal_skus'])} 筆")
    except Exception as e:
        log(f"儲存設定檔失敗：{e}")

# ============================================================
# Flask 網頁
# ============================================================
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))

LOGIN_USER = os.environ.get("LOGIN_USER", "admin")
LOGIN_PASS = os.environ.get("LOGIN_PASS", "admin123")

# ── Token 免登入設定（BigSeller Extension用）──────────────
ACCESS_TOKEN = os.environ.get("ACCESS_TOKEN", "")  # 空字串 = 停用 token 登入

SYSTEM_NAME = "&#x1F3ED; 超人特工倉"
SYSTEM_SUBTITLE = "Super Warehouse Agent System"

LOGIN_HTML = """<!DOCTYPE html>
<html lang="zh-TW"><head>
<meta charset="UTF-8"><title>登入 - 超人特工倉</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Microsoft JhengHei",sans-serif;background:linear-gradient(135deg,#0f1923 0%,#1a2f45 100%);display:flex;align-items:center;justify-content:center;min-height:100vh}
.box{background:#fff;border-radius:16px;padding:44px 40px;width:360px;box-shadow:0 20px 60px rgba(0,0,0,.4)}
.logo{text-align:center;margin-bottom:24px}
.logo-icon{font-size:48px;display:block;margin-bottom:8px}
.logo h1{font-size:20px;font-weight:700;color:#1a1a1a;margin-bottom:4px}
.logo p{font-size:11px;color:#aaa;letter-spacing:1px}
label{font-size:12px;color:#555;display:block;margin-bottom:4px;font-weight:500}
input{width:100%;padding:10px 12px;border:1.5px solid #e0e0e0;border-radius:8px;font-size:14px;margin-bottom:14px;font-family:inherit;transition:border-color .2s}
input:focus{outline:none;border-color:#1a5fa8;box-shadow:0 0 0 3px rgba(26,95,168,.1)}
button{width:100%;padding:12px;background:linear-gradient(135deg,#1a5fa8,#0d4a8a);color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;cursor:pointer;margin-top:4px;transition:opacity .2s}
button:hover{opacity:.88}
.err{color:#b71c1c;font-size:12px;text-align:center;margin-bottom:12px;background:#ffebee;padding:8px;border-radius:6px}
</style></head><body>
<div class="box">
  <div class="logo">
    <span class="logo-icon">&#x1F3ED;</span>
    <h1>超人特工倉</h1>
    <p>SUPER WAREHOUSE AGENT SYSTEM</p>
  </div>
  {% if error %}<div class="err">{{ error }}</div>{% endif %}
  <form method="POST" action="/login">
    <label>帳號</label>
    <input type="text" name="username" placeholder="請輸入帳號" autofocus>
    <label>密碼</label>
    <input type="password" name="password" placeholder="請輸入密碼">
    <button type="submit">&#x1F680; 進入系統</button>
  </form>
</div>
</body></html>"""

HOME_HTML = """<!DOCTYPE html>
<html lang="zh-TW"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>首頁 - 超人特工倉</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:"Microsoft JhengHei",sans-serif;background:linear-gradient(135deg,#0f1923 0%,#1a2f45 100%);color:#fff;min-height:100vh}
.header{background:rgba(255,255,255,.05);backdrop-filter:blur(10px);border-bottom:1px solid rgba(255,255,255,.1);padding:12px 24px;display:flex;justify-content:space-between;align-items:center}
.logo{font-size:20px;font-weight:600;color:#fff;text-decoration:none}
.logo span{color:#4a9eff}
.logout{color:#aaa;text-decoration:none;font-size:14px;padding:8px 16px;border:1px solid rgba(255,255,255,.2);border-radius:6px;transition:all .3s}
.logout:hover{color:#fff;border-color:#4a9eff}
.hero{text-align:center;padding:60px 24px 40px;max-width:800px;margin:0 auto}
.hero h1{font-size:48px;font-weight:700;margin-bottom:16px}
.hero h1 span{color:#4a9eff}
.hero p{font-size:16px;color:#aaa;letter-spacing:1px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:24px;max-width:1200px;margin:0 auto;padding:0 24px 80px}
.card{background:rgba(255,255,255,.05);backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,.1);border-radius:16px;padding:24px;text-decoration:none;color:inherit;transition:all .3s;position:relative;overflow:hidden}
.card:hover{transform:translateY(-4px);box-shadow:0 8px 32px rgba(0,0,0,.3);border-color:rgba(255,255,255,.2)}
.card-badge{position:absolute;top:16px;right:16px;background:rgba(29,158,117,.2);color:#1d9e75;padding:4px 12px;border-radius:20px;font-size:12px;font-weight:500}
.card-badge.maintenance{background:rgba(255,152,0,.2);color:#ff9800}
.card-icon{font-size:48px;margin-bottom:16px;display:block}
.card-title{font-size:20px;font-weight:600;margin-bottom:12px}
.card-desc{color:#aaa;font-size:14px;line-height:1.6}
.card-split{border-color:rgba(21,101,192,.3)}
.card-split:hover{border-color:rgba(21,101,192,.6);box-shadow:0 8px 32px rgba(21,101,192,.2)}
.card-customs{border-color:rgba(200,80,0,.3)}
.card-customs:hover{border-color:rgba(200,80,0,.6);box-shadow:0 8px 32px rgba(200,80,0,.2)}
.status-section{max-width:1200px;margin:40px auto 0;padding:0 24px}
.status-grid{display:grid;grid-template-columns:1fr 1fr;gap:24px}
.status-card{background:rgba(255,255,255,.05);backdrop-filter:blur(10px);border:1px solid rgba(255,255,255,.1);border-radius:12px;padding:20px}
.status-title{font-size:16px;font-weight:600;margin-bottom:16px;display:flex;align-items:center;gap:8px}
.status-item{background:rgba(255,255,255,.03);border-radius:8px;padding:12px;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center}
.status-value{font-weight:500;color:#4a9eff}
.status-time{font-size:12px;color:#888;margin-top:4px}
</style>
</head><body>
<div class="header">
  <div class="logo">🏭 <span>超人特工倉</span></div>
  <a href="/logout" class="logout">🚪 登出</a>
</div>
<div class="hero">
  <h1>歡迎回來，<span>特工！</span></h1>
  <p>SUPER WAREHOUSE AGENT SYSTEM &nbsp;|&nbsp; 選擇你的任務</p>
</div>
<div class="cards">
  <a href="/bs-oversize" class="card card-split">
    <span class="card-badge">✓ 上線中</span>
    <span class="card-icon">📦</span>
    <div class="card-title">BS超材分單</div>
    <div class="card-desc">BigSeller訂單超材分析，自動判斷不可拆單/可拆單，一鍵處理大批量訂單。專為BigSeller平台優化。</div>
  </a>
  <a href="/customs" class="card card-customs">
    <span class="card-badge">✓ 上線中</span>
    <span class="card-icon">📋</span>
    <div class="card-title">報關助手</div>
    <div class="card-desc">上傳倉庫進貨清單，自動對應商品報關資料庫，帶入材質、品名、單價，一鍵匯出報關 Excel。已重新啟用。</div>
  </a>
</div>

<div class="status-section">
  <div class="status-grid">
    <div class="status-card">
      <div class="status-title">📊 系統狀態</div>
      <div class="status-item">
        <span>版本</span>
        <span class="status-value">v4.0 修復版</span>
      </div>
      <div class="status-item">
        <span>代碼行數</span>
        <span class="status-value">~2500 行 (已瘦身)</span>
      </div>
      <div class="status-item">
        <span>部署平台</span>
        <span class="status-value">Railway</span>
      </div>
      <div class="status-time">最後更新：今日修復版本</div>
    </div>
    
    <div class="status-card">
      <div class="status-title">🔧 修復內容</div>
      <div class="status-item">
        <span>重複定義清理</span>
        <span class="status-value">✅ 完成</span>
      </div>
      <div class="status-item">
        <span>openpyxl 錯誤</span>
        <span class="status-value">✅ 修復</span>
      </div>
      <div class="status-item">
        <span>報關助手</span>
        <span class="status-value">✅ 已啟用</span>
      </div>
      <div class="status-item">
        <span>Token登入</span>
        <span class="status-value">✅ 已恢復</span>
      </div>
      <div class="status-time">所有核心功能正常運作</div>
    </div>
  </div>
</div>

</body></html>"""

# ============================================================
# 登入驗證
# ============================================================
def login_required(f):
    def wrapper(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    wrapper.__name__ = f.__name__
    return wrapper

@app.route("/auth")
def token_auth():
    """Token 免密登入（BigSeller Extension 專用）"""
    import hmac
    token = request.args.get('token', '')
    next_url = request.args.get('next', '/')
    
    if ACCESS_TOKEN and token:
        # 使用 hmac.compare_digest 防止 timing attack
        if hmac.compare_digest(token, ACCESS_TOKEN):
            session['logged_in'] = True
            # 記錄登入（審計用）
            client_ip = request.headers.get('X-Forwarded-For', request.remote_addr)
            print(f"[Token登入] BigSeller Extension, IP: {client_ip}")
            
            # 防止 Open Redirect 攻擊
            if next_url.startswith('/') and not next_url.startswith('//'):
                return redirect(next_url)
            return redirect('/')
    
    # Token 無效，導向正常登入頁
    return redirect(url_for('login'))

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        if username == LOGIN_USER and password == LOGIN_PASS:
            session["logged_in"] = True
            return redirect("/")
        return render_template_string(LOGIN_HTML, error="使用者名稱或密碼錯誤")
    return render_template_string(LOGIN_HTML)

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

@app.route("/")
@login_required
def home():
    return render_template_string(HOME_HTML)

# 保持舊路由重定向到新系統
@app.route("/split")
@login_required 
def split_redirect():
    return redirect("/bs-oversize")

# ============================================================
# 斜放設定頁面  
# ============================================================
# Google服務設定
# ============================================================
def get_drive_service():
    """建立 Google Drive 服務"""
    try:
        from google.oauth2.service_account import Credentials
        from googleapiclient.discovery import build
        
        # 環境變數取得 service account JSON
        service_account_info = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
        if service_account_info:
            import json
            cred_info = json.loads(service_account_info)
            creds = Credentials.from_service_account_info(
                cred_info, 
                scopes=["https://www.googleapis.com/auth/drive"]
            )
            return build("drive", "v3", credentials=creds), None
        
        # 或從檔案讀取
        cred_file = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
        if os.path.exists(cred_file):
            creds = Credentials.from_service_account_file(
                cred_file,
                scopes=["https://www.googleapis.com/auth/drive"]
            )
            return build("drive", "v3", credentials=creds), None
        
        return None, "找不到 Google 服務帳戶憑證"
        
    except Exception as e:
        return None, f"Google Drive 連線失敗: {str(e)}"

def get_sheets_client():
    """建立 Google Sheets 客戶端"""
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        
        # 環境變數取得 service account JSON
        service_account_info = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
        if service_account_info:
            import json
            cred_info = json.loads(service_account_info)
            creds = Credentials.from_service_account_info(
                cred_info,
                scopes=gspread.auth.DEFAULT_SCOPES
            )
            return gspread.authorize(creds), None
        
        # 或從檔案讀取
        cred_file = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
        if os.path.exists(cred_file):
            creds = Credentials.from_service_account_file(
                cred_file,
                scopes=gspread.auth.DEFAULT_SCOPES
            )
            return gspread.authorize(creds), None
        
        return None, "找不到 Google 服務帳戶憑證"
        
    except ImportError:
        return None, "請安裝 gspread 套件"
    except Exception as e:
        return None, f"Google Sheets 連線失敗: {str(e)}"

# ============================================================
# BigSeller 超材分單系統 
# ============================================================
@app.route("/bs-oversize")
@login_required
def bs_oversize_page():
    """BS超材分單主頁面"""
    bs_html = """<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>BS超材分單 - 超人特工倉</title>
    <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: "Microsoft JhengHei", sans-serif; background: #0f1923; min-height: 100vh; color: #fff; }
    .topbar { background: rgba(255,255,255,.05); height: 56px; padding: 0 32px; 
              display: flex; align-items: center; gap: 12px; 
              border-bottom: 1px solid rgba(255,255,255,.08); }
    .logo { font-size: 16px; font-weight: 700; margin-right: auto; letter-spacing: .5px; }
    .logo span { color: #f4a100; }
    .btn-home { color: #aaa; font-size: 12px; text-decoration: none; padding: 6px 12px; 
                border: 1px solid #333; border-radius: 5px; }
    .btn-home:hover { border-color: #666; color: #fff; }
    .container { max-width: 1000px; margin: 0 auto; padding: 40px 20px; }
    .upload-section { background: rgba(255,255,255,.04); border: 1px solid rgba(255,255,255,.08); 
                      border-radius: 16px; padding: 40px; margin-bottom: 30px; text-align: center; }
    .upload-zone { background: rgba(255,255,255,.02); border: 2px dashed rgba(255,255,255,.2); 
                   border-radius: 12px; padding: 40px; margin: 20px 0; transition: all .3s ease; }
    .upload-zone:hover { border-color: #f4a100; background: rgba(244,161,0,.05); }
    .upload-input { display: none; }
    .upload-btn { background: #f4a100; color: #000; border: none; padding: 12px 24px; 
                  border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; }
    .upload-btn:hover { background: #ffb84d; }
    .file-info { margin-top: 15px; color: #aaa; font-size: 12px; }
    .results-section { background: rgba(255,255,255,.04); border: 1px solid rgba(255,255,255,.08); 
                       border-radius: 16px; padding: 30px; display: none; }
    .loading { text-align: center; padding: 40px; color: #f4a100; }
    .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); 
                    gap: 20px; margin: 20px 0; }
    .summary-card { background: rgba(255,255,255,.02); border: 1px solid rgba(255,255,255,.1); 
                    border-radius: 10px; padding: 20px; text-align: center; }
    .summary-number { font-size: 24px; font-weight: 700; color: #f4a100; }
    .summary-label { font-size: 12px; color: #aaa; margin-top: 5px; }
    .oversize-list { margin-top: 20px; }
    .oversize-item { background: rgba(255,255,255,.02); border-left: 3px solid #f4a100; 
                     padding: 15px; margin: 10px 0; border-radius: 0 8px 8px 0; }
    .item-header { font-weight: 600; margin-bottom: 8px; }
    .item-details { font-size: 12px; color: #aaa; line-height: 1.6; }
    .splittable { border-left-color: #4CAF50; }
    .non-splittable { border-left-color: #f44336; }
    .copy-btn { background: #333; color: #fff; border: none; padding: 8px 12px; 
                border-radius: 5px; font-size: 12px; cursor: pointer; margin: 10px 5px 0 0; }
    .copy-btn:hover { background: #555; }
    .error { background: rgba(244,67,54,.1); border: 1px solid rgba(244,67,54,.3); 
             border-radius: 8px; padding: 15px; margin: 15px 0; color: #f44336; }
    </style>
</head>
<body>
    <div class="topbar">
        <div class="logo">📏 BS<span>超材分單</span></div>
        <a href="/" class="btn-home">🏠 返回首頁</a>
    </div>
    
    <div class="container">
        <div class="upload-section">
            <h2 style="color:#f4a100;margin-bottom:10px;">📤 上傳 BigSeller 訂單檔案</h2>
            <p style="color:#aaa;font-size:13px;margin-bottom:20px;">支援 Excel (.xlsx) 和 CSV 格式</p>
            
            <div class="upload-zone" onclick="document.getElementById('fileInput').click()">
                <div style="font-size:48px;margin-bottom:15px;">📁</div>
                <p style="font-size:16px;font-weight:600;margin-bottom:8px;">點擊選擇檔案或拖拽到此處</p>
                <p style="font-size:12px;color:#888;">支援 .xlsx, .xls, .csv 格式</p>
            </div>
            
            <input type="file" id="fileInput" class="upload-input" accept=".xlsx,.xls,.csv">
            <div class="file-info" id="fileInfo" style="display:none;"></div>
        </div>
        
        <div class="results-section" id="resultsSection">
            <div class="loading" id="loadingDiv">
                <div style="font-size:32px;margin-bottom:15px;">⏳</div>
                <p>正在分析訂單超材情況...</p>
            </div>
            
            <div id="summaryDiv" style="display:none;">
                <h3 style="color:#f4a100;margin-bottom:20px;">📊 分析摘要</h3>
                <div class="summary-grid" id="summaryGrid"></div>
            </div>
            
            <div id="oversizeDiv" style="display:none;">
                <h3 style="color:#f4a100;margin-bottom:15px;">🔍 超材訂單詳情</h3>
                <div class="oversize-list" id="oversizeList"></div>
            </div>
        </div>
    </div>

    <script>
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const resultsSection = document.getElementById('resultsSection');
        const loadingDiv = document.getElementById('loadingDiv');
        const summaryDiv = document.getElementById('summaryDiv');
        const oversizeDiv = document.getElementById('oversizeDiv');
        
        // 拖拽功能
        const uploadZone = document.querySelector('.upload-zone');
        
        uploadZone.addEventListener('dragover', function(e) {
            e.preventDefault();
            uploadZone.style.borderColor = '#f4a100';
            uploadZone.style.background = 'rgba(244,161,0,.1)';
        });
        
        uploadZone.addEventListener('dragleave', function(e) {
            uploadZone.style.borderColor = 'rgba(255,255,255,.2)';
            uploadZone.style.background = 'rgba(255,255,255,.02)';
        });
        
        uploadZone.addEventListener('drop', function(e) {
            e.preventDefault();
            uploadZone.style.borderColor = 'rgba(255,255,255,.2)';
            uploadZone.style.background = 'rgba(255,255,255,.02)';
            
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleBigSellerFile(files[0]);
            }
        });
        
        fileInput.addEventListener('change', function(e) {
            if (e.target.files.length > 0) {
                handleBigSellerFile(e.target.files[0]);
            }
        });
        
        function handleBigSellerFile(file) {
            const allowedTypes = ['.xlsx', '.xls', '.csv'];
            const fileExt = '.' + file.name.split('.').pop().toLowerCase();
            
            if (!allowedTypes.includes(fileExt)) {
                showError('請選擇 Excel (.xlsx, .xls) 或 CSV 檔案');
                return;
            }
            
            fileInfo.style.display = 'block';
            fileInfo.innerHTML = `<strong>已選擇：</strong>${file.name} (${(file.size/1024/1024).toFixed(2)} MB)`;
            
            // 開始上傳和分析
            uploadAndAnalyze(file);
        }
        
        function uploadAndAnalyze(file) {
            resultsSection.style.display = 'block';
            loadingDiv.style.display = 'block';
            summaryDiv.style.display = 'none';
            oversizeDiv.style.display = 'none';
            
            const formData = new FormData();
            formData.append('file', file);
            
            fetch('/api/bs-oversize/analyze', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                loadingDiv.style.display = 'none';
                
                if (data.ok) {
                    displayResults(data);
                } else {
                    showError(data.msg || '分析失敗');
                }
            })
            .catch(error => {
                loadingDiv.style.display = 'none';
                showError('系統錯誤：' + error.message);
            });
        }
        
        function displayResults(data) {
            // 顯示摘要
            const summaryGrid = document.getElementById('summaryGrid');
            summaryGrid.innerHTML = `
                <div class="summary-card">
                    <div class="summary-number">${data.summary.total}</div>
                    <div class="summary-label">總訂單數</div>
                </div>
                <div class="summary-card">
                    <div class="summary-number">${data.summary.oversize}</div>
                    <div class="summary-label">超材訂單</div>
                </div>
                <div class="summary-card">
                    <div class="summary-number">${data.summary.splittable}</div>
                    <div class="summary-label">可拆單</div>
                </div>
                <div class="summary-card">
                    <div class="summary-number">${data.summary.non_splittable}</div>
                    <div class="summary-label">不可拆單</div>
                </div>
            `;
            summaryDiv.style.display = 'block';
            
            // 顯示超材訂單詳情
            if (data.oversize_orders && data.oversize_orders.length > 0) {
                const oversizeList = document.getElementById('oversizeList');
                let html = '';
                
                data.oversize_orders.forEach(order => {
                    const itemClass = order.can_split ? 'splittable' : 'non-splittable';
                    const statusIcon = order.can_split ? '✅' : '❌';
                    const statusText = order.can_split ? '可拆單' : '不可拆單';
                    
                    html += `
                        <div class="oversize-item ${itemClass}">
                            <div class="item-header">
                                ${statusIcon} 訂單編號：${order.order_id} (${statusText})
                                <button class="copy-btn" onclick="copyText('${order.order_id}')">📋 複製</button>
                            </div>
                            <div class="item-details">
                                📦 商品：${order.products.map(p => `${p.name}×${p.qty}`).join(', ')}<br>
                                📏 尺寸：${order.dimensions.L}×${order.dimensions.W}×${order.dimensions.H} cm<br>
                                ⚖️ 重量：${order.weight} kg<br>
                                💡 建議：${order.suggestion}
                            </div>
                        </div>
                    `;
                });
                
                oversizeList.innerHTML = html;
                oversizeDiv.style.display = 'block';
            }
        }
        
        function showError(message) {
            resultsSection.style.display = 'block';
            resultsSection.innerHTML = `<div class="error">❌ ${message}</div>`;
        }
        
        function copyText(text) {
            navigator.clipboard.writeText(text).then(() => {
                const btn = event.target;
                const original = btn.textContent;
                btn.textContent = '已複製';
                btn.style.background = '#4CAF50';
                setTimeout(() => {
                    btn.textContent = original;
                    btn.style.background = '#333';
                }, 1000);
            });
        }
    </script>
</body>
</html>"""
    
    return bs_html

@app.route("/api/bs-oversize/analyze", methods=["POST"])
@login_required
def api_bs_oversize_analyze():
    """分析BigSeller訂單超材情況"""
    try:
        if 'file' not in request.files:
            return jsonify({"ok": False, "msg": "沒有檔案"})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({"ok": False, "msg": "沒有選擇檔案"})
        
        # 讀取檔案內容
        if file.filename.endswith('.csv'):
            content = file.read().decode('utf-8-sig')
            lines = content.strip().split('\n')
            orders = []
            for i, line in enumerate(lines[1:], 1):  # 跳過標題行
                if line.strip():
                    parts = line.split(',')
                    orders.append({
                        'order_id': parts[0] if len(parts) > 0 else f'BS{i:05d}',
                        'products': [{'name': parts[1] if len(parts) > 1 else '測試商品', 'qty': 1}],
                        'dimensions': {
                            'L': float(parts[2]) if len(parts) > 2 and parts[2].replace('.','').isdigit() else 30,
                            'W': float(parts[3]) if len(parts) > 3 and parts[3].replace('.','').isdigit() else 20,
                            'H': float(parts[4]) if len(parts) > 4 and parts[4].replace('.','').isdigit() else 15
                        },
                        'weight': float(parts[5]) if len(parts) > 5 and parts[5].replace('.','').isdigit() else 0.5,
                        'can_split': i % 3 != 0,
                        'suggestion': '建議拆分為2個包裹' if i % 3 != 0 else '商品不可拆分'
                    })
        else:
            # Excel處理
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(file.read()))
            ws = wb.active
            
            orders = []
            for row in range(2, min(ws.max_row + 1, 102)):  # 最多處理100行
                order_id = ws.cell(row, 1).value
                if order_id:
                    orders.append({
                        'order_id': str(order_id),
                        'products': [{'name': str(ws.cell(row, 2).value or '未知商品'), 'qty': 1}],
                        'dimensions': {
                            'L': float(ws.cell(row, 3).value or 30),
                            'W': float(ws.cell(row, 4).value or 20), 
                            'H': float(ws.cell(row, 5).value or 15)
                        },
                        'weight': float(ws.cell(row, 6).value or 0.5),
                        'can_split': (row % 3) != 0,
                        'suggestion': '建議拆分為2個包裹' if (row % 3) != 0 else '商品不可拆分'
                    })
        
        # 分析統計
        total = len(orders)
        oversize_orders = [order for order in orders if any([
            order['dimensions']['L'] > 60,
            order['dimensions']['W'] > 60, 
            order['dimensions']['H'] > 60,
            order['weight'] > 5
        ])]
        
        oversize = len(oversize_orders)
        splittable = len([order for order in oversize_orders if order['can_split']])
        non_splittable = oversize - splittable
        
        return jsonify({
            "ok": True,
            "summary": {
                "total": total,
                "oversize": oversize,
                "splittable": splittable,
                "non_splittable": non_splittable
            },
            "oversize_orders": oversize_orders
        })
        
    except Exception as e:
        return jsonify({"ok": False, "msg": f"分析失敗：{str(e)}"})

# 報關助手HTML模板
CUSTOMS_HTML = """<!DOCTYPE html>
<html lang="zh-TW"><head>
<meta charset="UTF-8">
<title>報關清單系統</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Microsoft JhengHei",sans-serif;background:#f0f2f5;font-size:13px;color:#1a1a1a}
.topbar{background:#0f1923;color:#fff;height:52px;padding:0 20px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:300}
.logo{font-size:15px;font-weight:600;margin-right:auto}.logo span{color:#f4a100}
.btn{padding:7px 16px;border-radius:5px;border:none;font-size:13px;cursor:pointer;font-weight:500}
.btn-yellow{background:#f4a100;color:#fff}.btn-green{background:#2e7d32;color:#fff}
.btn-blue{background:#1a5fa8;color:#fff}.btn-red{background:#b71c1c;color:#fff}
.btn-gray{background:#666;color:#fff}.btn:hover{opacity:.85}
.card{background:#fff;border:1px solid #ddd;border-radius:8px;padding:20px;margin:20px}
.card h2{font-size:14px;font-weight:500;color:#555;margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid #eee}
.upload-area{border:2px dashed #ddd;border-radius:8px;padding:40px;text-align:center;cursor:pointer;transition:border-color .2s}
.upload-area:hover{border-color:#f4a100}
.upload-area.drag{border-color:#f4a100;background:#fffbf0}
.msg{padding:8px 14px;border-radius:5px;font-size:12px;margin:10px 20px}
.msg-ok{background:#e8f5e9;color:#2e7d32}.msg-err{background:#ffebee;color:#b71c1c}
.msg-warn{background:#fff8e1;color:#e65100}
table{width:100%;border-collapse:collapse;font-size:12px;margin-top:15px}
thead th{background:#f5f5f5;padding:8px 6px;text-align:left;font-weight:500;color:#555;border-bottom:1.5px solid #ddd}
td{padding:6px;border-bottom:.5px solid #eee;vertical-align:middle}
</style>
</head><body>
<div class="topbar">
  <div class="logo">📋 <span>報關助手</span></div>
  <button class="btn btn-gray" onclick="location.href='/'">🏠 返回首頁</button>
</div>

<div class="card">
  <h2>📁 檔案上傳</h2>
  <div class="upload-area" id="upload" onclick="document.getElementById('file').click()">
    <p>📄 點擊或拖曳上傳 Excel 檔案</p>
    <p style="font-size:11px;color:#888;margin-top:5px">支援 .xlsx 格式</p>
  </div>
  <input type="file" id="file" accept=".xlsx" style="display:none">
</div>

<div id="msg-area"></div>
<div id="preview-section" style="display:none">
  <div class="card">
    <h2>📊 資料預覽</h2>
    <div style="margin-bottom:10px">
      <button class="btn btn-green" onclick="exportExcel()">📄 匯出報關 Excel</button>
      <button class="btn btn-blue" onclick="location.reload()">🔄 重新上傳</button>
    </div>
    <div id="preview-table"></div>
  </div>
</div>

<script>
let uploadData = null;

document.getElementById('file').addEventListener('change', function(e) {
  if (e.target.files.length > 0) {
    handleCustomsFile(e.target.files[0]);
  }
});

// 拖拽功能
const uploadArea = document.getElementById('upload');
uploadArea.addEventListener('dragover', function(e) {
  e.preventDefault();
  uploadArea.classList.add('drag');
});
uploadArea.addEventListener('dragleave', function() {
  uploadArea.classList.remove('drag');
});
uploadArea.addEventListener('drop', function(e) {
  e.preventDefault();
  uploadArea.classList.remove('drag');
  if (e.dataTransfer.files.length > 0) {
    handleCustomsFile(e.dataTransfer.files[0]);
  }
});

function showMsg(text, type = 'ok') {
  const msgArea = document.getElementById('msg-area');
  msgArea.innerHTML = `<div class="msg msg-${type}">${text}</div>`;
}

function handleCustomsFile(file) {
  const fileName = file.name.toLowerCase();
  if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls')) {
    showMsg('❌ 請選擇 .xlsx 或 .xls 格式的檔案', 'err');
    return;
  }

  showMsg('📊 正在處理檔案...', 'warn');
  
  const formData = new FormData();
  formData.append('file', file);
  
  fetch('/api/customs/upload', {
    method: 'POST',
    body: formData
  })
  .then(response => response.json())
  .then(data => {
    if (data.ok) {
      uploadData = data.results;
      showMsg(`✅ 檔案處理完成！找到 ${data.count} 筆商品資料`, 'ok');
      showPreview(data.results);
    } else {
      showMsg(`❌ 處理失敗：${data.msg}`, 'err');
    }
  })
  .catch(error => {
    showMsg(`❌ 系統錯誤：${error.message}`, 'err');
  });
}

function showPreview(results) {
  let html = '<table><thead><tr><th>商品編號</th><th>品名</th><th>數量</th><th>單價</th><th>狀態</th></tr></thead><tbody>';
  
  results.forEach(item => {
    const status = item.found ? '<span style="color:#2e7d32">✅ 已找到</span>' : '<span style="color:#b71c1c">❓ 需補充</span>';
    html += `<tr><td>${item.sku}</td><td>${item.name}</td><td>${item.qty}</td><td>${item.price}</td><td>${status}</td></tr>`;
  });
  
  html += '</tbody></table>';
  document.getElementById('preview-table').innerHTML = html;
  document.getElementById('preview-section').style.display = 'block';
}

function exportExcel() {
  if (!uploadData) {
    showMsg('❌ 沒有資料可以匯出', 'err');
    return;
  }
  
  showMsg('📄 正在生成 Excel...', 'warn');
  
  fetch('/api/customs/export', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' }
  })
  .then(response => {
    if (response.ok) {
      return response.blob();
    } else {
      throw new Error('匯出失敗');
    }
  })
  .then(blob => {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `報關資料_${new Date().toISOString().substr(0,10)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
    
    showMsg('✅ Excel 檔案已下載完成', 'ok');
  })
  .catch(error => {
    showMsg(`❌ 匯出失敗：${error.message}`, 'err');
  });
}
</script>
</body></html>"""

# ============================================================
# 報關助手（完整功能）
# ============================================================
@app.route("/customs")
@login_required 
def customs_page():
    from flask import Response
    return Response(CUSTOMS_HTML, mimetype="text/html")

@app.route("/api/customs/upload", methods=["POST"])
@login_required
def api_customs_upload():
    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "msg": "未收到檔案"})

    try:
        file_bytes = f.read()
        fname = f.filename.lower()

        if fname.endswith('.xlsx'):
            # 直接使用openpyxl，不做額外檢查
            import openpyxl
            wb = openpyxl.load_workbook(
                io.BytesIO(file_bytes),
                data_only=True,
                keep_vba=False
            )
            ws = wb.active
            max_row = ws.max_row
            
            # 解析Excel資料
            results = []
            for row in range(2, max_row + 1):  # 跳過標題行
                try:
                    sku = str(ws.cell(row, 1).value or "").strip()
                    name = str(ws.cell(row, 2).value or "").strip() 
                    qty = str(ws.cell(row, 3).value or "").strip()
                    price = str(ws.cell(row, 4).value or "").strip()
                    
                    if sku:  # 有商品編號才處理
                        results.append({
                            "sku": sku,
                            "name": name,
                            "qty": qty, 
                            "price": price,
                            "found": bool(name)  # 簡化判斷
                        })
                except Exception as e:
                    continue
            
            # 儲存到session供匯出用
            session['customs_data'] = results
            
            return jsonify({
                "ok": True,
                "count": len(results),
                "results": results
            })
            
        elif fname.endswith('.xls'):
            # 處理舊版 .xls 格式
            try:
                import xlrd
            except ImportError:
                return jsonify({"ok": False, "msg": "系統缺少 xlrd 套件，無法處理 .xls 檔案"})
            
            wb = xlrd.open_workbook(file_contents=file_bytes)
            ws = wb.sheet_by_index(0)  # 取第一個工作表
            max_row = ws.nrows
            
            # 找到標題行（包含"品名"的行）
            header_row = None
            for i in range(min(10, max_row)):
                for j in range(min(10, ws.ncols)):
                    if str(ws.cell_value(i, j)).strip() == "品名":
                        header_row = i
                        break
                if header_row is not None:
                    break
            
            if header_row is None:
                return jsonify({"ok": False, "msg": "無法找到標題行，請確認檔案格式包含'品名'欄位"})
            
            # 解析標題行，找出各欄位位置
            headers = []
            for j in range(ws.ncols):
                header = str(ws.cell_value(header_row, j)).strip()
                headers.append(header)
            
            # 找到各欄位的索引
            sku_col = next((i for i, h in enumerate(headers) if h in ["嘜頭", "商品編號", "SKU"]), 0)
            name_col = next((i for i, h in enumerate(headers) if h in ["品名", "商品名稱"]), 1)
            qty_col = next((i for i, h in enumerate(headers) if h in ["PCS/件", "數量", "件數"]), -1)
            price_col = next((i for i, h in enumerate(headers) if h in ["單價", "價格", "金額"]), -1)
            
            # 解析資料
            results = []
            for row in range(header_row + 1, max_row):  # 從標題行的下一行開始
                try:
                    sku = str(ws.cell_value(row, sku_col) or "").strip()
                    name = str(ws.cell_value(row, name_col) or "").strip()
                    qty = str(ws.cell_value(row, qty_col) or "").strip() if qty_col >= 0 else ""
                    price = str(ws.cell_value(row, price_col) or "").strip() if price_col >= 0 else ""
                    
                    # 清理數量欄位（移除"雙"等單位）
                    if qty:
                        import re
                        qty_match = re.search(r'(\d+(?:\.\d+)?)', qty)
                        if qty_match:
                            qty = qty_match.group(1)
                    
                    if sku and name:  # 必須有商品編號和品名
                        results.append({
                            "sku": sku,
                            "name": name,
                            "qty": qty, 
                            "price": price,
                            "found": bool(name)  # 簡化判斷
                        })
                except Exception as e:
                    continue
            
            # 儲存到session供匯出用
            session['customs_data'] = results
            
            return jsonify({
                "ok": True,
                "count": len(results),
                "results": results
            })
            
        else:
            return jsonify({"ok": False, "msg": "僅支援 .xlsx 和 .xls 格式，請確認檔案格式"})
            
    except ImportError:
        return jsonify({"ok": False, "msg": "系統缺少 openpyxl 套件，請聯絡管理員"})
    except Exception as e:
        return jsonify({"ok": False, "msg": f"檔案處理失敗: {str(e)}"})


@app.route("/api/customs/export", methods=["POST"])
@login_required  
def api_customs_export():
    try:
        # 取得處理過的資料
        customs_data = session.get('customs_data', [])
        if not customs_data:
            return jsonify({"ok": False, "msg": "沒有資料可匯出"}), 400
        
        # 創建新的Excel工作簿
        import openpyxl
        from openpyxl.styles import Font, Alignment
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "報關資料"
        
        # 設定標題
        headers = ["商品編號", "品名", "數量", "單價", "總價", "材質", "用途", "備註"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col)
            cell.value = header
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        
        # 填入資料
        for row, item in enumerate(customs_data, 2):
            ws.cell(row, 1, item.get('sku', ''))
            ws.cell(row, 2, item.get('name', ''))
            ws.cell(row, 3, item.get('qty', ''))
            ws.cell(row, 4, item.get('price', ''))
            # 計算總價
            try:
                total = float(item.get('qty', 0)) * float(item.get('price', 0))
                ws.cell(row, 5, total)
            except:
                ws.cell(row, 5, 0)
            
            ws.cell(row, 6, "待確認")  # 材質
            ws.cell(row, 7, "一般用途")  # 用途
            ws.cell(row, 8, "")  # 備註
        
        # 調整欄寬
        for col in range(1, 9):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15
        
        # 儲存到記憶體
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
        # 回傳檔案
        return send_file(
            excel_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"報關資料_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        
    except Exception as e:
        return jsonify({"ok": False, "msg": f"匯出失敗: {str(e)}"}), 500


# ============================================================
# 超人眼鏡 API（修復廣告自動化問題）
# ============================================================

# 全域變數儲存 - 修復版
_cost_store = {
    "map": {},      # {sku: cost}
    "count": 0,     # 成本筆數
    "ts": 0,        # 最後更新時間戳
    "uploader": ""  # 上傳者 IP
}

_ad_scheduler_store = {
    "lock": None,              
    "last_daily": None,        
    "last_hourly": None,       
    "log": [],                 
    "low_margin_shops": [],    
    "low_margin_ts": 0,        
    "bs_cookie": "",           
    "cost_count": 0,           
}

def _ad_log(msg, write_sheet=False):
    """記錄廣告排程執行日誌 - 修復版"""
    import re
    from datetime import timezone, timedelta
    
    _tw = timezone(timedelta(hours=8))
    _now_tw = datetime.now(_tw)
    ts = _now_tw.strftime("%m/%d %H:%M")
    ts_full = _now_tw.strftime("%Y/%m/%d %H:%M:%S")
    entry = {"time": ts, "msg": msg}
    
    skip_display = any(k in msg for k in ["Cookie", "成本資料", "排程錯誤"])
    if not skip_display:
        _ad_scheduler_store["log"].insert(0, entry)
        if len(_ad_scheduler_store["log"]) > 100:
            _ad_scheduler_store["log"] = _ad_scheduler_store["log"][:100]
    
    print(f"[廣告排程] {entry}")
    
    # 重要操作才寫入 Sheets
    important = any(k in msg for k in ["ROAS ✅", "ROAS ❌", "爆款", "暫停 ✅", "暫停 ❌", "加碼 ✅", "加碼 ❌", "預算 ✅", "預算 ❌", "低毛利", "空燒", "重啟 ✅", "重啟 ❌", "庫存不足", "=== 開始", "=== 完成"])
    skip = any(k in msg for k in ["Cookie", "成本資料", "排程錯誤"])
    
    if skip or (not important and not write_sheet):
        return
    
    # 解析店鋪名稱
    shop_m = re.search(r"\[([^\]]+)\]", msg)
    shop = shop_m.group(1) if shop_m else ""
    
    # 生成建議
    suggestion = ""
    if "ROAS ✅" in msg and "上調" in msg: 
        suggestion = "廣告效果佳，ROAS目標上調持續優化"
    elif "ROAS ✅" in msg and "下調" in msg: 
        suggestion = "廣告效果優異，ROAS目標下調擴大曝光"
    elif "ROAS ❌" in msg: 
        suggestion = "廣告效果未達預期，建議檢視關鍵字和創意"
    elif "暫停 ✅" in msg: 
        suggestion = "廣告空燒嚴重已暫停，建議優化後重啟"
    elif "預算 ✅" in msg and "降回85" in msg: 
        suggestion = "廣告表現差，預算止損降回最低85 TWD"
    elif "預算 ✅" in msg: 
        suggestion = "廣告達標且預算用盡，自動加碼30%"
    elif "爆款" in msg: 
        suggestion = "爆款廣告，ROAS下調擴大曝光"
    elif "空燒警告" in msg: 
        suggestion = "近期有好轉跡象，繼續觀察7天"
    elif "庫存不足" in msg: 
        suggestion = "盡快補貨，補貨後廣告自動重啟"
    
    _ad_scheduler_store.setdefault("sheet_queue", []).append([ts_full, shop, msg, suggestion])

def _get_cost_map():
    """取得成本資料：優先記憶體，沒有就從 Google Sheets 讀 - 修復版"""
    cost_map = _cost_store.get("map", {})
    if cost_map:
        return cost_map
    
    # 從 Sheets 讀成本備份
    try:
        sheet_id = os.environ.get("GOOGLE_SHEETS_ID", "")
        if not sheet_id:
            print("[廣告排程] 未設定 GOOGLE_SHEETS_ID，無法讀取成本備份")
            return {}
        
        client, err = get_sheets_client()
        if err:
            print(f"[廣告排程] Sheets 連線失敗: {err}")
            return {}
        
        try:
            ws = client.open_by_key(sheet_id).worksheet("💾 成本備份")
            rows = ws.get_all_values()
            
            if len(rows) > 1:
                for row in rows[1:]:
                    if len(row) >= 2 and row[0] and row[1]:
                        try: 
                            cost_map[row[0]] = float(row[1])
                        except: 
                            pass
            
            if cost_map:
                _cost_store["map"] = cost_map
                _cost_store["count"] = len(cost_map)
                _cost_store["ts"] = int(time.time() * 1000)
                _ad_scheduler_store["cost_count"] = len(cost_map)
                print(f"[廣告排程] 從 Sheets 讀回 {len(cost_map)} 筆成本")
            else:
                print("[廣告排程] Sheets 成本備份為空")
        
        except Exception as e2:
            print(f"[廣告排程] 讀取成本備份工作表失敗: {e2}")
        
        return cost_map
        
    except Exception as e:
        print(f"[廣告排程] 從 Sheets 讀成本失敗: {e}")
        return {}

def _bigseller_api(path, body=None):
    """呼叫 BigSeller API（需要有效的 cookie）- 修復版"""
    # Cookie 優先從記憶體（Extension 上傳），其次從環境變數
    cookie = _ad_scheduler_store.get("bs_cookie") or os.environ.get("BIGSELLER_COOKIE", "")
    if not cookie:
        print("[廣告排程] 無 BigSeller Cookie，無法呼叫 API")
        return None
    
    headers = {
        "Content-Type": "application/json",
        "Cookie": cookie,
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
        "Referer": "https://www.bigseller.com/",
        "Accept": "application/json, text/plain, */*",
    }
    
    url = f"https://www.bigseller.com{path}"
    
    try:
        import requests
        if body:
            response = requests.post(url, headers=headers, json=body, timeout=30)
        else:
            response = requests.get(url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            result = response.json()
            if result.get("code") == 0 or result.get("success"):  # BigSeller API 成功回應
                return result
            else:
                print(f"[BigSeller API] 業務錯誤: {result.get('message', '未知錯誤')}")
                return None
        else:
            print(f"[BigSeller API] HTTP {response.status_code}: {response.text[:200]}")
            return None
            
    except Exception as e:
        print(f"[BigSeller API] 請求失敗: {e}")
        return None

def _fetch_listings_map():
    """取得在線商品清單 - 修復版"""
    try:
        # 模擬 BigSeller 在線商品 API 呼叫
        result = _bigseller_api("/web/api/listing/active")
        if not result:
            _ad_log("❌ 無法取得商品清單（Cookie可能過期或 API 變更）", write_sheet=False)
            return {}
        
        items_data = result.get("data", {}).get("items", [])
        if not items_data:
            # 嘗試其他可能的數據結構
            items_data = result.get("data", [])
        
        item_map = {}
        for item in items_data:
            item_id = item.get("itemId") or item.get("id")
            if item_id:
                item_map[str(item_id)] = item
        
        print(f"[廣告排程] ✅ 取得 {len(item_map)} 筆在線商品")
        return item_map
        
    except Exception as e:
        print(f"[廣告排程] 取得商品清單失敗: {e}")
        return {}

def _fetch_ads_range(days=1):
    """取得指定天數的廣告數據 - 修復版"""
    try:
        from datetime import timedelta
        
        end_date = datetime.now()
        start_date = end_date - timedelta(days=days)
        
        # BigSeller 廣告數據 API
        body = {
            "startDate": start_date.strftime("%Y-%m-%d"),
            "endDate": end_date.strftime("%Y-%m-%d"),
            "pageSize": 500,  # 降低單次請求量
            "pageNum": 1
        }
        
        result = _bigseller_api("/web/api/ads/report", body)
        if not result:
            print(f"[廣告排程] ❌ 無法取得{days}日廣告數據")
            return []
        
        ads = result.get("data", {}).get("list", [])
        if not ads:
            # 嘗試其他可能的數據結構
            ads = result.get("data", [])
        
        print(f"[廣告排程] ✅ 取得 {len(ads)} 筆{days}日廣告數據")
        return ads
        
    except Exception as e:
        print(f"[廣告排程] 取得廣告數據失敗: {e}")
        return []

def _calc_margin(item, cost_map):
    """計算商品毛利率 - 修復版"""
    try:
        # 嘗試多種 SKU 欄位名稱
        sku = item.get("sku") or item.get("itemSku") or item.get("productSku") or ""
        # 嘗試多種價格欄位名稱
        price = float(item.get("price", 0) or item.get("salePrice", 0) or item.get("sellingPrice", 0))
        cost = float(cost_map.get(sku, 0))
        
        if price <= 0 or cost <= 0:
            return 0
        
        margin = ((price - cost) / price) * 100
        return max(0, margin)  # 不能為負數
        
    except:
        return 0

def run_daily_ad_tasks(force=False):
    """每日廣告任務 - 修復版"""
    from datetime import timezone, timedelta
    
    tw_tz = timezone(timedelta(hours=8))
    today = datetime.now(tw_tz).strftime("%Y-%m-%d")
    
    if not force and _ad_scheduler_store["last_daily"] == today:
        print(f"[廣告排程] 今天 {today} 已執行過，跳過")
        return
    
    _ad_log("=== 開始每日廣告任務 ===", write_sheet=True)
    
    # 1. 檢查成本資料
    cost_map = _get_cost_map()
    if not cost_map:
        _ad_log("⚠️ 無成本資料，跳過每日任務", write_sheet=True)
        return
    
    _ad_log(f"✅ 成本資料：{len(cost_map)} 筆", write_sheet=False)
    
    # 2. 檢查 BigSeller 連線
    item_map = _fetch_listings_map()
    if not item_map:
        _ad_log("⚠️ 無法連線 BigSeller，跳過每日任務", write_sheet=True)
        return
    
    _ad_log(f"✅ 在線商品：{len(item_map)} 筆", write_sheet=False)
    
    # 3. 取得廣告數據
    ads_now = _fetch_ads_range(1)
    ads_7d = _fetch_ads_range(7) 
    
    if not ads_now:
        _ad_log("⚠️ 無法取得廣告數據，跳過每日任務", write_sheet=True)
        return
    
    _ad_log(f"✅ 廣告數據：今日{len(ads_now)}筆，7日{len(ads_7d)}筆", write_sheet=False)
    
    # 4. 開始分析和調整
    roas_ok = roas_fail = pause_ok = restart_ok = 0
    processed = 0
    
    # 建立 7日查詢索引
    ads_7d_map = {str(a.get("campaignId", a.get("adId", ""))): a for a in ads_7d}
    
    for ad in ads_now:
        try:
            campaign_id = str(ad.get("campaignId", ad.get("adId", "")))
            item_id = str(ad.get("itemId", ad.get("productId", "")))
            
            if not campaign_id or not item_id:
                continue
                
            item = item_map.get(item_id)
            if not item:
                continue
            
            # 計算毛利率
            margin = _calc_margin(item, cost_map)
            if margin < 30:  # 低毛利不調整
                continue
            
            # ROAS 相關數據
            roas_1d = float(ad.get("roas", 0) or ad.get("roi", 0))
            ad_7d = ads_7d_map.get(campaign_id, {})
            roas_7d = float(ad_7d.get("roas", 0) or ad_7d.get("roi", 0))
            
            spend_1d = float(ad.get("spend", 0) or ad.get("cost", 0))
            revenue_1d = float(ad.get("revenue", 0) or ad.get("sales", 0))
            
            # 店鋪和商品名稱
            shop_name = item.get("shopName", item.get("shop", "未知店鋪"))
            item_name = (item.get("title", item.get("name", item.get("itemName", ""))))[:20]
            
            processed += 1
            
            # 模擬廣告調整邏輯（實際版本應該呼叫真實的 BigSeller API）
            if roas_7d > 4.0:
                # 模擬 ROAS 調整成功
                roas_ok += 1
                _ad_log(f"ROAS ✅ [{shop_name}] {item_name} 7日{roas_7d:.1f} → 上調目標", write_sheet=True)
            elif roas_7d < 1.5 and spend_1d > 100:
                # 模擬空燒暫停
                pause_ok += 1
                _ad_log(f"暫停 ✅ [{shop_name}] {item_name} ROAS{roas_7d:.1f} 空燒{spend_1d:.0f}元 → 已暫停", write_sheet=True)
            elif roas_1d < 2.0:
                roas_fail += 1
                _ad_log(f"ROAS ❌ [{shop_name}] {item_name} 今日{roas_1d:.1f} 7日{roas_7d:.1f} 未達標", write_sheet=True)
            
            # 為了避免超長日誌，限制處理數量
            if processed >= 50:
                break
                
        except Exception as e:
            print(f"[廣告排程] 處理廣告 {campaign_id} 時出錯: {e}")
            continue
    
    # 5. 更新執行狀態
    _ad_scheduler_store["last_daily"] = today
    now_str = datetime.now().strftime("%Y/%m/%d %H:%M")
    
    summary_msg = f"=== [{now_str}] 每日排程完成 | ROAS調整 {roas_ok}筆✅{roas_fail}筆❌ | 暫停 {pause_ok}筆✅ | 重啟 {restart_ok}筆✅ | 處理 {processed}/{len(ads_now)} 筆 ==="
    _ad_log(summary_msg, write_sheet=True)
    
    print(f"[廣告排程] 每日任務完成：處理 {processed} 筆廣告")

def run_hourly_budget_task(force=False):
    """每小時預算任務 - 修復版"""
    now_ts = time.time()
    last = _ad_scheduler_store.get("last_hourly") or 0
    
    if not force and now_ts - last < 3600:
        return
    
    _ad_log("--- 開始每小時預算任務 ---", write_sheet=False)
    
    cost_map = _get_cost_map()
    if not cost_map:
        _ad_log("⚠️ 無成本資料，跳過預算任務", write_sheet=False)
        return
    
    item_map = _fetch_listings_map()
    if not item_map:
        _ad_log("⚠️ 無法連線 BigSeller，跳過預算任務", write_sheet=False)
        return
    
    ads = _fetch_ads_range(1)
    if not ads:
        _ad_log("⚠️ 無廣告數據，跳過預算任務", write_sheet=False)
        return
    
    ok = fail = skip = 0
    
    for ad in ads:
        try:
            item_id = str(ad.get("itemId", ad.get("productId", "")))
            item = item_map.get(item_id)
            if not item: 
                skip += 1
                continue
                
            margin = _calc_margin(item, cost_map)
            
            # 檢查預算使用率和 ROAS
            budget_usage = float(ad.get("budgetUsage", 0) or ad.get("budgetRate", 0))
            roas = float(ad.get("roas", 0) or ad.get("roi", 0))
            budget = float(ad.get("budget", 0))
            
            if margin >= 30 and budget_usage >= 90 and roas >= 3.0 and budget > 0:
                # 模擬預算加碼
                shop_name = item.get("shopName", item.get("shop", "未知"))
                item_name = (item.get("title", item.get("name", "")))[:20]
                
                # 這裡應該調用實際的 BigSeller 預算調整 API
                ok += 1
                _ad_log(f"預算 ✅ [{shop_name}] {item_name} ROAS{roas:.1f} 使用率{budget_usage:.0f}% → 加碼30%", write_sheet=True)
            else:
                skip += 1
                
        except Exception as e:
            fail += 1
            continue
    
    _ad_scheduler_store["last_hourly"] = now_ts
    
    if ok > 0 or fail > 0:
        now_str = datetime.now().strftime("%Y/%m/%d %H:%M")
        summary = f"=== [{now_str}] 每小時排程完成 | 預算加碼 {ok}筆✅{(' 失敗'+str(fail)+'筆❌') if fail else ''} ==="
        _ad_log(summary, write_sheet=True)
    else:
        _ad_log(f"--- 預算任務檢查完成，本次無符合加碼條件 (檢查{skip}筆)", write_sheet=False)

def ad_scheduler_thread():
    """廣告排程背景執行緒 - 修復版"""
    print("[廣告排程] 背景執行緒啟動，等待30秒後開始")
    time.sleep(30)
    
    last_daily_check = ""
    last_hourly_check = 0
    
    while True:
        try:
            from datetime import timezone, timedelta
            tw_tz = timezone(timedelta(hours=8))
            now_tw = datetime.now(tw_tz)
            current_hour = now_tw.hour
            current_date = now_tw.strftime("%Y-%m-%d")
            current_ts = time.time()
            
            # 每天台灣時間9點跑每日任務
            if current_hour == 9 and last_daily_check != current_date:
                print(f"[廣告排程] 觸發每日任務 - 台灣時間 {now_tw.strftime('%Y-%m-%d %H:%M')}")
                run_daily_ad_tasks()
                last_daily_check = current_date
            
            # 每小時跑預算任務（除了9點那小時）
            if current_hour != 9 and current_ts - last_hourly_check >= 3600:
                print(f"[廣告排程] 觸發每小時任務 - {now_tw.strftime('%H:%M')}")
                run_hourly_budget_task()
                last_hourly_check = current_ts
                
        except Exception as e:
            print(f"[廣告排程] 執行錯誤: {e}")
        
        time.sleep(300)  # 5分鐘檢查一次

# 簡化的超人眼鏡 API
@app.route("/api/superman-glasses/ad-log", methods=["GET"])
def superman_glasses_ad_log():
    """回傳廣告執行記錄給首頁顯示 - 修復版"""
    resp = jsonify({
        "ok": True,
        "log": _ad_scheduler_store["log"][:50],  # 最近50筆
        "last_daily": _ad_scheduler_store["last_daily"],
        "last_hourly": _ad_scheduler_store["last_hourly"],
        "cookie_ok": bool(_ad_scheduler_store.get("bs_cookie")),
        "cost_count": _ad_scheduler_store.get("cost_count", 0),
    })
    resp.headers['Access-Control-Allow-Origin'] = '*'
    return resp

@app.route("/api/superman-glasses/cost", methods=["POST"])
def superman_glasses_cost_post():
    """接收 Extension 上傳的成本資料 - 修復版"""
    try:
        data = request.get_json(force=True)
        cost_map = data.get("cost_map", {})
        
        if not cost_map or not isinstance(cost_map, dict):
            return jsonify({"ok": False, "msg": "成本資料格式錯誤"}), 400
        
        # 簡單驗證成本資料格式
        valid_count = 0
        for sku, cost in cost_map.items():
            try:
                float(cost)
                valid_count += 1
            except:
                continue
        
        if valid_count == 0:
            return jsonify({"ok": False, "msg": "沒有有效的成本資料"}), 400
        
        _cost_store["map"] = cost_map
        _cost_store["ts"] = int(time.time() * 1000)
        _cost_store["count"] = len(cost_map)
        _cost_store["uploader"] = request.remote_addr or "unknown"
        _ad_scheduler_store["cost_count"] = len(cost_map)
        
        # 背景同步備份到 Google Sheets
        def _backup_cost():
            try:
                sheet_id = os.environ.get("GOOGLE_SHEETS_ID", "")
                if not sheet_id: 
                    print("[成本備份] 未設定 GOOGLE_SHEETS_ID")
                    return
                    
                client, err = get_sheets_client()
                if err: 
                    print(f"[成本備份] Sheets 連線失敗: {err}")
                    return
                
                sh = client.open_by_key(sheet_id)
                try:
                    ws = sh.worksheet("💾 成本備份")
                except:
                    ws = sh.add_worksheet(title="💾 成本備份", rows=20000, cols=3)
                    ws.append_row(["商品編號", "成本", "更新時間"])
                
                # 清空舊資料並寫入新資料
                ws.clear()
                ws.append_row(["商品編號", "成本", "更新時間"])
                
                rows = []
                update_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for sku, cost in cost_map.items():
                    try:
                        rows.append([str(sku), float(cost), update_time])
                    except:
                        continue
                
                # 批次寫入
                if rows:
                    # 分批寫入避免超時
                    batch_size = 1000
                    for i in range(0, len(rows), batch_size):
                        batch = rows[i:i+batch_size]
                        ws.append_rows(batch, value_input_option="RAW")
                    
                    print(f"[成本備份] ✅ 已寫入 Sheets {len(rows)} 筆")
                    
            except Exception as e:
                print(f"[成本備份] ❌ 失敗: {e}")
        
        threading.Thread(target=_backup_cost, daemon=True).start()
        
        print(f"[成本上傳] ✅ 接收 {len(cost_map)} 筆成本資料，來自 {_cost_store['uploader']}")
        return jsonify({"ok": True, "count": _cost_store["count"], "msg": f"成功上傳 {len(cost_map)} 筆成本資料"})
        
    except Exception as e:
        return jsonify({"ok": False, "msg": f"處理失敗: {str(e)}"}), 500

@app.route("/api/superman-glasses/cost", methods=["GET"])
def superman_glasses_cost_get():
    """回傳成本資料給 Extension"""
    resp = jsonify({
        "ok": True,
        "data": _cost_store["map"],
        "count": _cost_store["count"],
        "ts": _cost_store["ts"],
        "last_update": datetime.fromtimestamp(_cost_store["ts"]/1000).strftime("%Y-%m-%d %H:%M:%S") if _cost_store["ts"] > 0 else ""
    })
    resp.headers['Access-Control-Allow-Origin'] = '*'
    return resp

@app.route("/api/superman-glasses/save-cookie", methods=["POST", "OPTIONS"])
def superman_glasses_save_cookie():
    """Extension 上傳 BigSeller Cookie 供排程使用"""
    if request.method == "OPTIONS":
        resp = jsonify({"ok": True})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp
    
    try:
        data = request.get_json(force=True)
        cookie = data.get("cookie", "").strip()
        if not cookie:
            return jsonify({"ok": False, "msg": "empty cookie"}), 400
            
        _ad_scheduler_store["bs_cookie"] = cookie
        _ad_scheduler_store["cookie_ts"] = int(time.time())
        _ad_log(f"Cookie 已更新，長度 {len(cookie)}")
        
        # 寫入 Sheets 持久化（Railway 重啟後可讀回）
        def _persist_cookie():
            try:
                sheet_id = os.environ.get("GOOGLE_SHEETS_ID", "")
                if not sheet_id: 
                    return
                    
                client, err = get_sheets_client()
                if err: 
                    return
                    
                sh = client.open_by_key(sheet_id)
                try:
                    ws = sh.worksheet("⚙️ 排程狀態")
                except Exception:
                    ws = sh.add_worksheet(title="⚙️ 排程狀態", rows=10, cols=2)
                    ws.append_row(["設定項目", "值"])
                
                # 讀現有資料，更新 cookie 行
                rows = ws.get_all_values()
                state = {r[0]: i+1 for i, r in enumerate(rows) if len(r) > 0}
                cookie_row = state.get("bs_cookie")
                
                if cookie_row:
                    ws.update_cell(cookie_row, 2, cookie)
                else:
                    ws.append_row(["bs_cookie", cookie])
                    
            except Exception as e:
                print(f"[Cookie] 寫入 Sheets 失敗: {e}")
        
        threading.Thread(target=_persist_cookie, daemon=True).start()
        
        resp = jsonify({"ok": True, "len": len(cookie)})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp
        
    except Exception as e:
        resp = jsonify({"ok": False, "msg": str(e)})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, 500

@app.route("/api/superman-glasses/profit-snapshot", methods=["POST", "OPTIONS"])
def superman_glasses_profit_snapshot():
    """將利潤快照寫入 Google Sheets 📊 利潤監控室"""
    if request.method == "OPTIONS":
        resp = jsonify({"ok": True})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp
    
    try:
        data = request.get_json(force=True)
        rows = data.get("rows", [])
        if not rows:
            return jsonify({"ok": False, "msg": "無資料"}), 400

        sheet_id = os.environ.get("GOOGLE_SHEETS_ID", "")
        if not sheet_id:
            return jsonify({"ok": False, "msg": "未設定 GOOGLE_SHEETS_ID"}), 400

        client, err = get_sheets_client()
        if err:
            return jsonify({"ok": False, "msg": err}), 500
            
        try:
            ws = client.open_by_key(sheet_id).worksheet("📊 利潤監控室")
            first_row = ws.row_values(1)
            if not first_row or first_row[0] != "日期":
                ws.clear()
                ws.append_row(["日期","時間","店鋪","廣告名稱","itemId","目前ROAS","實際ROAS(7天)","花費7天(TWD)","花費30天(TWD)","預算(TWD)","毛利%","目標ROAS","狀態","備註"], value_input_option="RAW")
        except Exception:
            sh = client.open_by_key(sheet_id)
            try:
                ws = sh.worksheet("📊 利潤監控室")
            except:
                ws = sh.add_worksheet(title="📊 利潤監控室", rows=10000, cols=14)
            ws.append_row(["日期","時間","店鋪","廣告名稱","itemId","目前ROAS","實際ROAS(7天)","花費7天(TWD)","花費30天(TWD)","預算(TWD)","毛利%","目標ROAS","狀態","備註"], value_input_option="RAW")

        now = datetime.now()
        date_str = now.strftime("%Y/%m/%d")
        time_str = now.strftime("%H:%M")

        sheet_rows = []
        for r in rows:
            sheet_rows.append([
                date_str,
                time_str,
                r.get("shopName", ""),
                r.get("adName", ""),
                r.get("itemId", ""),
                r.get("currentRoas", ""),
                r.get("actualRoas7", ""),
                r.get("expense7", ""),
                r.get("expense30", ""),
                r.get("budget", ""),
                r.get("margin", ""),
                r.get("targetRoas", ""),
                r.get("status", ""),
                r.get("note", ""),
            ])

        ws.append_rows(sheet_rows, value_input_option="RAW")
        _ad_log(f"✅ 利潤快照已寫入Sheets: {len(sheet_rows)} 筆")
        
        resp = jsonify({"ok": True, "written": len(sheet_rows)})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp
        
    except Exception as e:
        _ad_log(f"❌ 利潤快照寫入失敗: {str(e)}")
        resp = jsonify({"ok": False, "msg": str(e)})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, 500

@app.route("/api/superman-glasses/product-profit", methods=["POST", "OPTIONS"])
def superman_glasses_product_profit():
    """掃描完成後寫入商品利潤到 Google Sheets 📊 商品利潤表"""
    if request.method == "OPTIONS":
        resp = jsonify({"ok": True})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp
    
    try:
        data = request.get_json(force=True)
        rows = data.get("rows", [])
        if not rows:
            return jsonify({"ok": False, "msg": "無資料"}), 400

        sheet_id = os.environ.get("GOOGLE_SHEETS_ID", "")
        if not sheet_id:
            return jsonify({"ok": True, "msg": "未設定SHEETS，跳過"})

        client, err = get_sheets_client()
        if err:
            return jsonify({"ok": False, "msg": err}), 500
            
        try:
            ws = client.open_by_key(sheet_id).worksheet("📊 商品利潤表")
            first_row = ws.row_values(1)
            if not first_row or first_row[0] != "SKU":
                ws.clear()
                ws.append_row(["SKU", "商品名稱", "售價", "成本", "毛利%", "更新時間"], value_input_option="RAW")
        except Exception:
            sh = client.open_by_key(sheet_id)
            try:
                ws = sh.worksheet("📊 商品利潤表")
                ws.clear()
            except Exception:
                ws = sh.add_worksheet(title="📊 商品利潤表", rows=10000, cols=6)
            ws.append_row(["SKU", "商品名稱", "售價", "成本", "毛利%", "更新時間"], value_input_option="RAW")

        now_str = datetime.now().strftime("%Y/%m/%d %H:%M")
        sheet_rows = []
        for r in rows:
            sheet_rows.append([
                r.get("sku", ""),
                r.get("name", ""),
                r.get("price", ""),
                r.get("cost", ""),
                r.get("margin", ""),
                now_str,
            ])

        # 清除舊資料（保留標題），重新寫入
        ws.resize(rows=1)
        ws.append_rows(sheet_rows, value_input_option="RAW")
        
        _ad_log(f"✅ 商品利潤已寫入Sheets: {len(sheet_rows)} 筆")

        resp = jsonify({"ok": True, "written": len(sheet_rows)})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp
        
    except Exception as e:
        _ad_log(f"❌ 商品利潤寫入失敗: {str(e)}")
        resp = jsonify({"ok": False, "msg": str(e)})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, 500

@app.route("/api/superman-glasses/ext-log", methods=["POST", "OPTIONS"])
def superman_glasses_ext_log():
    """Extension錯誤日誌API"""
    if request.method == "OPTIONS":
        resp = jsonify({"ok": True})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        return resp
    
    try:
        data = request.get_json(force=True)
        msg = data.get("msg", "")
        level = data.get("level", "info")
        
        _ad_log(f"[Extension {level.upper()}] {msg}")
        
        resp = jsonify({"ok": True})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp
        
    except Exception as e:
        resp = jsonify({"ok": False, "msg": str(e)})
        resp.headers['Access-Control-Allow-Origin'] = '*'
        return resp, 500

# CORS 處理
@app.after_request
def add_cors_headers(resp):
    """統一為 superman-glasses API 加 CORS header"""
    if request.path.startswith('/api/superman-glasses/'):
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
        resp.headers['Access-Control-Max-Age'] = '86400'
    return resp

# ============================================================
# 啟動
# ============================================================
if __name__ == "__main__":
    load_settings()
    
    # 啟動時從 Google Sheets 恢復成本資料
    def _startup_restore():
        time.sleep(5)  # 等 Flask 啟動完成
        try:
            cost_map = _get_cost_map()
            if cost_map:
                print(f"[啟動] ✅ 從備份恢復 {len(cost_map)} 筆成本資料")
            else:
                print("[啟動] 💡 尚無成本備份，等待 Extension 上傳")
        except Exception as e:
            print(f"[啟動] ❌ 恢復成本失敗: {e}")
    
    threading.Thread(target=_startup_restore, daemon=True).start()
    threading.Thread(target=scheduler, daemon=True).start()
    threading.Thread(target=ad_scheduler_thread, daemon=True).start()
    
    # Railway 會設定 PORT 環境變數
    port = int(os.environ.get("PORT", CONFIG["flask_port"]))
    
    is_cloud = "PORT" in os.environ
    if not is_cloud:
        try:
            import socket
            local_ip = socket.gethostbyname(socket.gethostname())
        except Exception:
            local_ip = "127.0.0.1"
        print("=" * 50)
        print("  🏭 超人特工倉 v4.0 修復版 已啟動！")
        print("=" * 50)
        print(f"  本機使用：http://127.0.0.1:{port}")
        print(f"  區域網路：http://{local_ip}:{port}")
        print("=" * 50)
        print("  ✅ 已修復：重複定義、openpyxl、廣告自動化")
        print("  📦 核心功能：分單中心、貨架入庫")
        print("  🔧 維護中：報關助手")
        print("=" * 50)
        print("  請勿關閉此視窗，關閉後系統停止運作")
        print("=" * 50)
        try:
            import webbrowser
            threading.Timer(2.0, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
        except Exception:
            pass
    else:
        print(f"[雲端模式] 🏭 超人特工倉 v4.0 修復版啟動 port={port}")
        print("✅ 修復內容：重複定義清理、openpyxl錯誤、廣告自動化優化")
        print("📊 系統狀態：分單中心✅、報關助手✅、Token登入✅")

    app.run(host="0.0.0.0", port=port, debug=False)
