# -*- coding: utf-8 -*-
"""
4Sale 自動分單印單系統 v4
五個通路：宅配 / 超商 / 店到店 / 店到店隔日配 / 無包裝
區域細分：倉庫(前倉/主倉/備用倉) + 區域字母(A/B/C...)
"""

import sys, csv, io, os, re, json, threading, time, hashlib, secrets
from datetime import datetime
from flask import Flask, render_template_string, request, jsonify, session, redirect, url_for, send_file

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
    "col_txn":      "交易序號",
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

def check_oversize(ch, ship_raw, total_dim, max_side, weight):
    """
    判斷是否超材。
    回傳 (is_oversize: bool, warn_msg: str, can_split: bool, split_rules: dict)
    split_rules: {"max_dim": int, "max_side": int, "max_weight": int} 或 None
    """
    if total_dim == 0 and weight == 0:
        return False, "", False, None

    # 新竹物流
    if "新竹物流" in ship_raw:
        if total_dim > 210 or max_side > 150 or weight > 20:
            return True, f"超規 {total_dim:.0f}cm/{weight:.1f}kg → 異常單", False, None
        for limit, fee, desc in CONFIG["hsinchu_surcharge"]:
            if total_dim <= limit:
                if fee > 0:
                    return True, f"新竹補助不足 {total_dim:.0f}cm → {desc}", False, None
                return False, "", False, None
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

    # 隔日配 / 無包裝
    if ch in ("nextday", "nopkg"):
        if total_dim > 105 or max_side > 45 or weight > 10:
            return True, f"超材 {total_dim:.0f}cm/{weight:.1f}kg", False, None
        return False, "", False, None

    return False, "", False, None

def apply_diagonal(sku, L, W, H):
    """
    如果此 SKU 在可斜放清單內，回傳斜放後的有效最長邊。
    否則回傳原始最長邊。
    """
    import math
    diag_map = CONFIG.get("diagonal_skus", {})
    if sku in diag_map:
        effective = diag_map[sku]
        return effective, True   # (有效最長邊, 是否斜放)
    return max(L, W, H), False

def suggest_split(products, weight_per_item, L, W, H, rules=None):
    """
    建議拆單方式。
    rules: {"max_dim": int, "max_side": int, "max_weight": int}
    """
    if rules is None:
        rules = {
            "max_dim":    CONFIG["split_max_dim"],
            "max_side":   CONFIG["split_max_side"],
            "max_weight": CONFIG["split_max_weight"],
        }
    max_units  = CONFIG["split_max_units"]
    max_dim    = rules["max_dim"]
    max_side   = rules["max_side"]
    max_weight = rules["max_weight"]

    # 展開所有件數為一個清單
    all_items = []
    for p in products:
        for _ in range(max(1, p.get("qty", 1))):
            all_items.append(p["sku"])

    total_items = len(all_items)

    # 單件尺寸判斷
    sides = sorted([L, W, H], reverse=True)
    single_dim  = sum(sides)
    single_side = sides[0]
    single_w    = weight_per_item

    # 如果單件就超規，無法拆單
    if single_dim > max_dim or single_side > max_side or single_w > max_weight:
        return None, "單件商品已超規，無法拆單"

    # 每包最多放幾件
    max_per_pkg_dim    = int(max_dim    // single_dim)  if single_dim  > 0 else 999
    max_per_pkg_side   = int(max_side   // single_side) if single_side > 0 else 999
    max_per_pkg_weight = int(max_weight // single_w)    if single_w    > 0 else 999
    max_per_pkg = min(max_per_pkg_dim, max_per_pkg_side, max_per_pkg_weight)
    max_per_pkg = max(1, max_per_pkg)

    # 需要幾包
    import math
    needed = math.ceil(total_items / max_per_pkg)

    if needed > max_units:
        return None, f"需要 {needed} 包，超過最多 {max_units} 單上限，無法自動拆單"

    # 建立拆單建議（同SKU盡量放同一包）
    packages = []
    current_pkg = []
    current_count = 0

    # 按SKU分組，同SKU連續放
    from collections import Counter
    sku_counts = Counter(all_items)
    sorted_skus = sorted(sku_counts.items(), key=lambda x: -x[1])

    remaining = []
    for sku, qty in sorted_skus:
        remaining.extend([sku] * qty)

    for item in remaining:
        if current_count >= max_per_pkg:
            packages.append(current_pkg[:])
            current_pkg = []
            current_count = 0
        current_pkg.append(item)
        current_count += 1

    if current_pkg:
        packages.append(current_pkg)

    result = []
    for i, pkg in enumerate(packages):
        pkg_counter = Counter(pkg)
        items_str = ", ".join(f"{sku}×{cnt}" for sku, cnt in sorted(pkg_counter.items()))
        pkg_dim    = single_dim * len(pkg)
        pkg_weight = single_w   * len(pkg)
        result.append({
            "pkg_no":  i + 1,
            "count":   len(pkg),
            "items":   items_str,
            "dim":     round(pkg_dim, 1),
            "weight":  round(pkg_weight, 2),
        })

    return result, ""

def _get_zone_label(locs):
    """取得區域描述字串，供宅配分群用"""
    whs   = set(wh for wh, z in locs)
    zones = set(z  for wh, z in locs if z != "?")
    if not zones: zones = {"?"}
    if len(whs) == 1 and len(zones) == 1:
        return f"{next(iter(whs))}{next(iter(zones))}區"
    elif len(whs) == 1:
        return f"{next(iter(whs))}混單"
    return "混單"

def get_sort_key(k):
    """排序：宅配→超商→店到店單項品→店到店→隔日配→無包裝→可拆單→超材"""
    if k == "__delivery__":     return (0,  k, k)
    if k == "__store_single__": return (25, k, k)
    if k == "__nopkg__":        return (88, k, k)
    if k == "__splittable__":   return (92, k, k)
    if k == "__oversize__":     return (95, k, k)
    order = {"超商":1, "店到店":2, "店到店隔日配":3}
    for label, pri in order.items():
        if k.startswith(label):
            if "混單" in k: return (pri*10+9, "zzz", k)
            return (pri*10, k, k)
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
                "weight":    0.0,
                "fee":       0.0,  # 運費
            }
        oid    = (row.get(CONFIG["col_order_id"], "") or "").strip()
        raw_wh = (row.get(CONFIG["col_warehouse"], "") or "").strip()
        sku    = (row.get(CONFIG["col_sku"], "") or "").strip()
        wh, zone = parse_location(raw_wh)

        # 尺寸取最大值（避免空值覆蓋有效值）
        L = safe_float(row.get(CONFIG["col_length"], 0))
        W = safe_float(row.get(CONFIG["col_width"],  0))
        H = safe_float(row.get(CONFIG["col_height"], 0))
        Wt= safe_float(row.get(CONFIG["col_weight"], 0))
        Fe= parse_fee(row.get(CONFIG["col_fee"], 0))
        if L > 0: order_map[txn]["length"]  = max(order_map[txn]["length"],  L)
        if W > 0: order_map[txn]["width"]   = max(order_map[txn]["width"],   W)
        if H > 0: order_map[txn]["height"]  = max(order_map[txn]["height"],  H)
        if Wt> 0: order_map[txn]["weight"]  = max(order_map[txn]["weight"],  Wt)
        if Fe> 0: order_map[txn]["fee"]     = max(order_map[txn]["fee"],     Fe)

        order_map[txn]["order_ids"].append(oid)
        order_map[txn]["products"].append({
            "oid": oid, "sku": sku,
            "zone_raw": raw_wh, "wh": wh, "zone": zone,
        })
        order_map[txn]["locations"].add((wh, zone))
        order_map[txn]["total_qty"] += 1

    # ── 計算三邊總和、最長邊，斜放判斷，超材檢查 ────────
    for txn, o in order_map.items():
        L, W, H = o["length"], o["width"], o["height"]

        # 斜放判斷：若訂單內有可斜放 SKU，用對角線取代最長邊
        skus_in_order = [p["sku"] for p in o["products"]]
        diagonal_applied = False
        effective_side = max(L, W, H)
        for sku in skus_in_order:
            eff, is_diag = apply_diagonal(sku, L, W, H)
            if is_diag:
                effective_side = eff
                diagonal_applied = True
                break

        dims = sorted([L, W, H], reverse=True)
        o["max_side"]       = effective_side
        o["total_dim"]      = sum(dims)
        o["diagonal_used"]  = diagonal_applied
        o["oversize"], o["oversize_msg"], o["can_split"], o["split_rules"] = check_oversize(
            o["channel"], o["ship_raw"], o["total_dim"], o["max_side"], o["weight"]
        )

        # 若斜放後不超材，標示說明
        if diagonal_applied and not o["oversize"]:
            o["oversize_msg"] = "斜放後合規"

        # 拆單建議（運費=0才可拆單）
        o["split_suggestion"] = None
        o["split_error"]      = ""
        o["fee_paid"]         = o["fee"] > 0

        # 運費>0 → 標示買家已付，不拆單
        if o["oversize"] and o["fee"] > 0:
            o["oversize_msg"] += f"｜買家已付運費 {o['fee']:.0f} 元"

        can_split = o["oversize"] and o["can_split"] and o["fee"] == 0
        if can_split and o["total_qty"] > 0:
            weight_per = o["weight"] / o["total_qty"] if o["total_qty"] > 0 else 0
            from collections import Counter
            sku_count = Counter(p["sku"] for p in o["products"])
            prods = [{"sku": sku, "qty": qty} for sku, qty in sku_count.items()]
            suggestion, err = suggest_split(prods, weight_per, L, W, H, o["split_rules"])
            o["split_suggestion"] = suggestion
            o["split_error"]      = err

    # ── 套用分單規則 ──────────────────────────────────
    groups = {}

    def add(key, title, icon, color, o):
        if key not in groups:
            groups[key] = {"title": title, "icon": icon, "color": color, "orders": []}
        groups[key]["orders"].append(o)

    summary = {}

    for txn, o in order_map.items():
        ch    = o["channel"]
        locs  = o["locations"]
        meta  = CHANNEL_META.get(ch, CHANNEL_META["delivery"])
        label = meta["label"]
        icon  = meta["icon"]
        color = meta["color"]

        # 超材處理
        if o["oversize"]:
            o["oversize_channel"] = label
            if o["can_split"] and o["split_suggestion"]:
                # 可拆單 → 獨立可拆單分類
                add("__splittable__", "✂ 可拆單", "✂", "#e65100", o)
                summary["✂ 可拆單"] = summary.get("✂ 可拆單", 0) + 1
            else:
                # 不可拆單 → 超材異常分類
                add("__oversize__", "⚠ 超材", "⚠", "#b71c1c", o)
                summary["⚠ 超材"] = summary.get("⚠ 超材", 0) + 1
            continue

        # 無包裝 → 全部合併成一組
        if ch == "nopkg":
            add("__nopkg__", "📦 無包裝", "📦", "#00838f", o)
            summary["無包裝"] = summary.get("無包裝", 0) + 1
            continue

        # 宅配（新竹/嘉里/店到家）→ 外層一個大分類，內部依區域分群
        if ch == "delivery":
            o["delivery_zone"] = _get_zone_label(locs)  # 記錄區域供分群用
            add("__delivery__", "🚚 宅配", "🚚", "#1565c0", o)
            summary["宅配"] = summary.get("宅配", 0) + 1
            continue

        # 店到店單項品 → 獨立一組（單一SKU不管幾件）
        skus = set(p["sku"] for p in o["products"] if p["sku"])
        if ch == "store" and len(skus) == 1:
            add("__store_single__", "🏬 店到店 ｜ 單項品", "🏬", "#1b5e20", o)
            summary["店到店"] = summary.get("店到店", 0) + 1
            continue

        # 其餘通路（超商、店到店多品、隔日配）→ 依倉庫區域細分
        whs   = set(wh for wh, z in locs)
        zones = set(z  for wh, z in locs if z != "?")
        if not zones:
            zones = {"?"}

        if len(whs) == 1 and len(zones) == 1:
            wh    = next(iter(whs))
            zone  = next(iter(zones))
            key   = f"{label} - {wh}{zone}區"
            title = f"{icon} {label} ｜ {wh} {zone} 區"
        elif len(whs) == 1:
            key   = f"{label} - {next(iter(whs))}混單"
            title = f"{icon} {label} ｜ {next(iter(whs))} 混單"
        else:
            key   = f"{label} - 混單"
            title = f"{icon} {label} ｜ 混單（跨倉）"

        if "混單" in key:
            color = color + "bb"

        add(key, title, icon, color, o)
        summary[label] = summary.get(label, 0) + 1

    state["summary"] = summary
    return dict(sorted(groups.items(), key=lambda x: get_sort_key(x[0])))

def load_csv(source, is_text=False):
    for enc in ["utf-8-sig", "big5", "cp950", "utf-8"]:
        try:
            if is_text:
                reader = csv.DictReader(io.StringIO(source))
            else:
                f = open(source, encoding=enc, newline="")
                reader = csv.DictReader(f)
            rows = list(reader)
            if not is_text: f.close()
            if rows: return rows, enc
        except Exception:
            continue
    return [], None

def run_pipeline(rows=None):
    state["status"] = "fetching"
    if not rows:
        state["status"] = "error"
        state["status_msg"] = "無資料"
        return
    groups = split_orders(rows)
    state["groups"]      = groups
    state["total"]       = len(set((r.get(CONFIG["col_txn"]) or "") for r in rows))
    state["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    state["status"]      = "ready"
    state["status_msg"]  = f"共 {state['total']} 張訂單，分成 {len(groups)} 組 | {state['last_update']}"
    for k, g in groups.items():
        log(f"  {g['title']}：{len(g['orders'])} 張")
    log("分單完成")

def scheduler():
    while True:
        now = datetime.now()
        if now.hour == CONFIG["auto_run_hour"] and now.minute == CONFIG["auto_run_min"]:
            log("定時觸發")
            time.sleep(61)
        time.sleep(30)

# ============================================================
# 設定檔持久化（斜放商品清單）
# ============================================================
def _get_base_dir():
    """取得執行檔所在目錄（支援 PyInstaller 打包）"""
    if getattr(sys, 'frozen', False):
        # 打包成 exe 後，exe 所在目錄
        return os.path.dirname(sys.executable)
    else:
        # 一般 Python 執行
        return os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(_get_base_dir(), "settings.json")

def load_settings():
    """從 settings.json 讀取設定，啟動時呼叫"""
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if "diagonal_skus" in data:
                CONFIG["diagonal_skus"] = data["diagonal_skus"]
                log(f"已載入特殊可出超材品設定：{len(CONFIG['diagonal_skus'])} 筆")
    except Exception as e:
        log(f"讀取設定檔失敗：{e}")

def save_settings():
    """將設定寫入 settings.json"""
    try:
        data = {"diagonal_skus": CONFIG["diagonal_skus"]}
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"儲存設定檔失敗：{e}")

# ============================================================
# Flask 網頁
# ============================================================
app = Flask(__name__)

# -*- coding: utf-8 -*-
"""
4Sale 自動分單印單系統 v4
五個通路：宅配 / 超商 / 店到店 / 店到店隔日配 / 無包裝
區域細分：倉庫(前倉/主倉/備用倉) + 區域字母(A/B/C...)
"""

import sys, csv, io, os, re, json, threading, time
from datetime import datetime
from flask import Flask, render_template_string, request, jsonify

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
    "col_txn":      "交易序號",
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

def check_oversize(ch, ship_raw, total_dim, max_side, weight):
    """
    判斷是否超材。
    回傳 (is_oversize: bool, warn_msg: str, can_split: bool, split_rules: dict)
    split_rules: {"max_dim": int, "max_side": int, "max_weight": int} 或 None
    """
    if total_dim == 0 and weight == 0:
        return False, "", False, None

    # 新竹物流
    if "新竹物流" in ship_raw:
        if total_dim > 210 or max_side > 150 or weight > 20:
            return True, f"超規 {total_dim:.0f}cm/{weight:.1f}kg → 異常單", False, None
        for limit, fee, desc in CONFIG["hsinchu_surcharge"]:
            if total_dim <= limit:
                if fee > 0:
                    return True, f"新竹補助不足 {total_dim:.0f}cm → {desc}", False, None
                return False, "", False, None
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

    # 隔日配 / 無包裝
    if ch in ("nextday", "nopkg"):
        if total_dim > 105 or max_side > 45 or weight > 10:
            return True, f"超材 {total_dim:.0f}cm/{weight:.1f}kg", False, None
        return False, "", False, None

    return False, "", False, None

def apply_diagonal(sku, L, W, H):
    """
    如果此 SKU 在可斜放清單內，回傳斜放後的有效最長邊。
    否則回傳原始最長邊。
    """
    import math
    diag_map = CONFIG.get("diagonal_skus", {})
    if sku in diag_map:
        effective = diag_map[sku]
        return effective, True   # (有效最長邊, 是否斜放)
    return max(L, W, H), False

def suggest_split(products, weight_per_item, L, W, H, rules=None):
    """
    建議拆單方式。
    rules: {"max_dim": int, "max_side": int, "max_weight": int}
    """
    if rules is None:
        rules = {
            "max_dim":    CONFIG["split_max_dim"],
            "max_side":   CONFIG["split_max_side"],
            "max_weight": CONFIG["split_max_weight"],
        }
    max_units  = CONFIG["split_max_units"]
    max_dim    = rules["max_dim"]
    max_side   = rules["max_side"]
    max_weight = rules["max_weight"]

    # 展開所有件數為一個清單
    all_items = []
    for p in products:
        for _ in range(max(1, p.get("qty", 1))):
            all_items.append(p["sku"])

    total_items = len(all_items)

    # 單件尺寸判斷
    sides = sorted([L, W, H], reverse=True)
    single_dim  = sum(sides)
    single_side = sides[0]
    single_w    = weight_per_item

    # 如果單件就超規，無法拆單
    if single_dim > max_dim or single_side > max_side or single_w > max_weight:
        return None, "單件商品已超規，無法拆單"

    # 每包最多放幾件
    max_per_pkg_dim    = int(max_dim    // single_dim)  if single_dim  > 0 else 999
    max_per_pkg_side   = int(max_side   // single_side) if single_side > 0 else 999
    max_per_pkg_weight = int(max_weight // single_w)    if single_w    > 0 else 999
    max_per_pkg = min(max_per_pkg_dim, max_per_pkg_side, max_per_pkg_weight)
    max_per_pkg = max(1, max_per_pkg)

    # 需要幾包
    import math
    needed = math.ceil(total_items / max_per_pkg)

    if needed > max_units:
        return None, f"需要 {needed} 包，超過最多 {max_units} 單上限，無法自動拆單"

    # 建立拆單建議（同SKU盡量放同一包）
    packages = []
    current_pkg = []
    current_count = 0

    # 按SKU分組，同SKU連續放
    from collections import Counter
    sku_counts = Counter(all_items)
    sorted_skus = sorted(sku_counts.items(), key=lambda x: -x[1])

    remaining = []
    for sku, qty in sorted_skus:
        remaining.extend([sku] * qty)

    for item in remaining:
        if current_count >= max_per_pkg:
            packages.append(current_pkg[:])
            current_pkg = []
            current_count = 0
        current_pkg.append(item)
        current_count += 1

    if current_pkg:
        packages.append(current_pkg)

    result = []
    for i, pkg in enumerate(packages):
        pkg_counter = Counter(pkg)
        items_str = ", ".join(f"{sku}×{cnt}" for sku, cnt in sorted(pkg_counter.items()))
        pkg_dim    = single_dim * len(pkg)
        pkg_weight = single_w   * len(pkg)
        result.append({
            "pkg_no":  i + 1,
            "count":   len(pkg),
            "items":   items_str,
            "dim":     round(pkg_dim, 1),
            "weight":  round(pkg_weight, 2),
        })

    return result, ""

def _get_zone_label(locs):
    """取得區域描述字串，供宅配分群用"""
    whs   = set(wh for wh, z in locs)
    zones = set(z  for wh, z in locs if z != "?")
    if not zones: zones = {"?"}
    if len(whs) == 1 and len(zones) == 1:
        return f"{next(iter(whs))}{next(iter(zones))}區"
    elif len(whs) == 1:
        return f"{next(iter(whs))}混單"
    return "混單"

def get_sort_key(k):
    """排序：宅配→超商→店到店單項品→店到店→隔日配→無包裝→可拆單→超材"""
    if k == "__delivery__":     return (0,  k, k)
    if k == "__store_single__": return (25, k, k)
    if k == "__nopkg__":        return (88, k, k)
    if k == "__splittable__":   return (92, k, k)
    if k == "__oversize__":     return (95, k, k)
    order = {"超商":1, "店到店":2, "店到店隔日配":3}
    for label, pri in order.items():
        if k.startswith(label):
            if "混單" in k: return (pri*10+9, "zzz", k)
            return (pri*10, k, k)
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
                "weight":    0.0,
                "fee":       0.0,  # 運費
            }
        oid    = (row.get(CONFIG["col_order_id"], "") or "").strip()
        raw_wh = (row.get(CONFIG["col_warehouse"], "") or "").strip()
        sku    = (row.get(CONFIG["col_sku"], "") or "").strip()
        wh, zone = parse_location(raw_wh)

        # 尺寸取最大值（避免空值覆蓋有效值）
        L = safe_float(row.get(CONFIG["col_length"], 0))
        W = safe_float(row.get(CONFIG["col_width"],  0))
        H = safe_float(row.get(CONFIG["col_height"], 0))
        Wt= safe_float(row.get(CONFIG["col_weight"], 0))
        Fe= parse_fee(row.get(CONFIG["col_fee"], 0))
        if L > 0: order_map[txn]["length"]  = max(order_map[txn]["length"],  L)
        if W > 0: order_map[txn]["width"]   = max(order_map[txn]["width"],   W)
        if H > 0: order_map[txn]["height"]  = max(order_map[txn]["height"],  H)
        if Wt> 0: order_map[txn]["weight"]  = max(order_map[txn]["weight"],  Wt)
        if Fe> 0: order_map[txn]["fee"]     = max(order_map[txn]["fee"],     Fe)

        order_map[txn]["order_ids"].append(oid)
        order_map[txn]["products"].append({
            "oid": oid, "sku": sku,
            "zone_raw": raw_wh, "wh": wh, "zone": zone,
        })
        order_map[txn]["locations"].add((wh, zone))
        order_map[txn]["total_qty"] += 1

    # ── 計算三邊總和、最長邊，斜放判斷，超材檢查 ────────
    for txn, o in order_map.items():
        L, W, H = o["length"], o["width"], o["height"]

        # 斜放判斷：若訂單內有可斜放 SKU，用對角線取代最長邊
        skus_in_order = [p["sku"] for p in o["products"]]
        diagonal_applied = False
        effective_side = max(L, W, H)
        for sku in skus_in_order:
            eff, is_diag = apply_diagonal(sku, L, W, H)
            if is_diag:
                effective_side = eff
                diagonal_applied = True
                break

        dims = sorted([L, W, H], reverse=True)
        o["max_side"]       = effective_side
        o["total_dim"]      = sum(dims)
        o["diagonal_used"]  = diagonal_applied
        o["oversize"], o["oversize_msg"], o["can_split"], o["split_rules"] = check_oversize(
            o["channel"], o["ship_raw"], o["total_dim"], o["max_side"], o["weight"]
        )

        # 若斜放後不超材，標示說明
        if diagonal_applied and not o["oversize"]:
            o["oversize_msg"] = "斜放後合規"

        # 拆單建議（運費=0才可拆單）
        o["split_suggestion"] = None
        o["split_error"]      = ""
        o["fee_paid"]         = o["fee"] > 0

        # 運費>0 → 標示買家已付，不拆單
        if o["oversize"] and o["fee"] > 0:
            o["oversize_msg"] += f"｜買家已付運費 {o['fee']:.0f} 元"

        can_split = o["oversize"] and o["can_split"] and o["fee"] == 0
        if can_split and o["total_qty"] > 0:
            weight_per = o["weight"] / o["total_qty"] if o["total_qty"] > 0 else 0
            from collections import Counter
            sku_count = Counter(p["sku"] for p in o["products"])
            prods = [{"sku": sku, "qty": qty} for sku, qty in sku_count.items()]
            suggestion, err = suggest_split(prods, weight_per, L, W, H, o["split_rules"])
            o["split_suggestion"] = suggestion
            o["split_error"]      = err

    # ── 套用分單規則 ──────────────────────────────────
    groups = {}

    def add(key, title, icon, color, o):
        if key not in groups:
            groups[key] = {"title": title, "icon": icon, "color": color, "orders": []}
        groups[key]["orders"].append(o)

    summary = {}

    for txn, o in order_map.items():
        ch    = o["channel"]
        locs  = o["locations"]
        meta  = CHANNEL_META.get(ch, CHANNEL_META["delivery"])
        label = meta["label"]
        icon  = meta["icon"]
        color = meta["color"]

        # 超材處理
        if o["oversize"]:
            o["oversize_channel"] = label
            if o["can_split"] and o["split_suggestion"]:
                # 可拆單 → 獨立可拆單分類
                add("__splittable__", "✂ 可拆單", "✂", "#e65100", o)
                summary["✂ 可拆單"] = summary.get("✂ 可拆單", 0) + 1
            else:
                # 不可拆單 → 超材異常分類
                add("__oversize__", "⚠ 超材", "⚠", "#b71c1c", o)
                summary["⚠ 超材"] = summary.get("⚠ 超材", 0) + 1
            continue

        # 無包裝 → 全部合併成一組
        if ch == "nopkg":
            add("__nopkg__", "📦 無包裝", "📦", "#00838f", o)
            summary["無包裝"] = summary.get("無包裝", 0) + 1
            continue

        # 宅配（新竹/嘉里/店到家）→ 外層一個大分類，內部依區域分群
        if ch == "delivery":
            o["delivery_zone"] = _get_zone_label(locs)  # 記錄區域供分群用
            add("__delivery__", "🚚 宅配", "🚚", "#1565c0", o)
            summary["宅配"] = summary.get("宅配", 0) + 1
            continue

        # 店到店單項品 → 獨立一組（單一SKU不管幾件）
        skus = set(p["sku"] for p in o["products"] if p["sku"])
        if ch == "store" and len(skus) == 1:
            add("__store_single__", "🏬 店到店 ｜ 單項品", "🏬", "#1b5e20", o)
            summary["店到店"] = summary.get("店到店", 0) + 1
            continue

        # 其餘通路（超商、店到店多品、隔日配）→ 依倉庫區域細分
        whs   = set(wh for wh, z in locs)
        zones = set(z  for wh, z in locs if z != "?")
        if not zones:
            zones = {"?"}

        if len(whs) == 1 and len(zones) == 1:
            wh    = next(iter(whs))
            zone  = next(iter(zones))
            key   = f"{label} - {wh}{zone}區"
            title = f"{icon} {label} ｜ {wh} {zone} 區"
        elif len(whs) == 1:
            key   = f"{label} - {next(iter(whs))}混單"
            title = f"{icon} {label} ｜ {next(iter(whs))} 混單"
        else:
            key   = f"{label} - 混單"
            title = f"{icon} {label} ｜ 混單（跨倉）"

        if "混單" in key:
            color = color + "bb"

        add(key, title, icon, color, o)
        summary[label] = summary.get(label, 0) + 1

    state["summary"] = summary
    return dict(sorted(groups.items(), key=lambda x: get_sort_key(x[0])))

def load_csv(source, is_text=False):
    for enc in ["utf-8-sig", "big5", "cp950", "utf-8"]:
        try:
            if is_text:
                reader = csv.DictReader(io.StringIO(source))
            else:
                f = open(source, encoding=enc, newline="")
                reader = csv.DictReader(f)
            rows = list(reader)
            if not is_text: f.close()
            if rows: return rows, enc
        except Exception:
            continue
    return [], None

def run_pipeline(rows=None):
    state["status"] = "fetching"
    if not rows:
        state["status"] = "error"
        state["status_msg"] = "無資料"
        return
    groups = split_orders(rows)
    state["groups"]      = groups
    state["total"]       = len(set((r.get(CONFIG["col_txn"]) or "") for r in rows))
    state["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    state["status"]      = "ready"
    state["status_msg"]  = f"共 {state['total']} 張訂單，分成 {len(groups)} 組 | {state['last_update']}"
    for k, g in groups.items():
        log(f"  {g['title']}：{len(g['orders'])} 張")
    log("分單完成")

def scheduler():
    while True:
        now = datetime.now()
        if now.hour == CONFIG["auto_run_hour"] and now.minute == CONFIG["auto_run_min"]:
            log("定時觸發")
            time.sleep(61)
        time.sleep(30)

# ============================================================
# 設定檔持久化（斜放商品清單）
# ============================================================
def _get_base_dir():
    """取得執行檔所在目錄（支援 PyInstaller 打包）"""
    if getattr(sys, 'frozen', False):
        # 打包成 exe 後，exe 所在目錄
        return os.path.dirname(sys.executable)
    else:
        # 一般 Python 執行
        return os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(_get_base_dir(), "settings.json")

def load_settings():
    """從 settings.json 讀取設定，啟動時呼叫"""
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if "diagonal_skus" in data:
                CONFIG["diagonal_skus"] = data["diagonal_skus"]
                log(f"已載入特殊可出超材品設定：{len(CONFIG['diagonal_skus'])} 筆")
    except Exception as e:
        log(f"讀取設定檔失敗：{e}")

def save_settings():
    """將設定寫入 settings.json"""
    try:
        data = {"diagonal_skus": CONFIG["diagonal_skus"]}
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"儲存設定檔失敗：{e}")

# ============================================================
# Flask 網頁
# ============================================================
app = Flask(__name__)

# -*- coding: utf-8 -*-
"""
4Sale 自動分單印單系統 v4
五個通路：宅配 / 超商 / 店到店 / 店到店隔日配 / 無包裝
區域細分：倉庫(前倉/主倉/備用倉) + 區域字母(A/B/C...)
"""

import sys, csv, io, os, re, json, threading, time
from datetime import datetime
from flask import Flask, render_template_string, request, jsonify

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
    "col_txn":      "交易序號",
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

def check_oversize(ch, ship_raw, total_dim, max_side, weight):
    """
    判斷是否超材。
    回傳 (is_oversize: bool, warn_msg: str, can_split: bool, split_rules: dict)
    split_rules: {"max_dim": int, "max_side": int, "max_weight": int} 或 None
    """
    if total_dim == 0 and weight == 0:
        return False, "", False, None

    # 新竹物流
    if "新竹物流" in ship_raw:
        if total_dim > 210 or max_side > 150 or weight > 20:
            return True, f"超規 {total_dim:.0f}cm/{weight:.1f}kg → 異常單", False, None
        for limit, fee, desc in CONFIG["hsinchu_surcharge"]:
            if total_dim <= limit:
                if fee > 0:
                    return True, f"新竹補助不足 {total_dim:.0f}cm → {desc}", False, None
                return False, "", False, None
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

    # 隔日配 / 無包裝
    if ch in ("nextday", "nopkg"):
        if total_dim > 105 or max_side > 45 or weight > 10:
            return True, f"超材 {total_dim:.0f}cm/{weight:.1f}kg", False, None
        return False, "", False, None

    return False, "", False, None

def apply_diagonal(sku, L, W, H):
    """
    如果此 SKU 在可斜放清單內，回傳斜放後的有效最長邊。
    否則回傳原始最長邊。
    """
    import math
    diag_map = CONFIG.get("diagonal_skus", {})
    if sku in diag_map:
        effective = diag_map[sku]
        return effective, True   # (有效最長邊, 是否斜放)
    return max(L, W, H), False

def suggest_split(products, weight_per_item, L, W, H, rules=None):
    """
    建議拆單方式。
    rules: {"max_dim": int, "max_side": int, "max_weight": int}
    """
    if rules is None:
        rules = {
            "max_dim":    CONFIG["split_max_dim"],
            "max_side":   CONFIG["split_max_side"],
            "max_weight": CONFIG["split_max_weight"],
        }
    max_units  = CONFIG["split_max_units"]
    max_dim    = rules["max_dim"]
    max_side   = rules["max_side"]
    max_weight = rules["max_weight"]

    # 展開所有件數為一個清單
    all_items = []
    for p in products:
        for _ in range(max(1, p.get("qty", 1))):
            all_items.append(p["sku"])

    total_items = len(all_items)

    # 單件尺寸判斷
    sides = sorted([L, W, H], reverse=True)
    single_dim  = sum(sides)
    single_side = sides[0]
    single_w    = weight_per_item

    # 如果單件就超規，無法拆單
    if single_dim > max_dim or single_side > max_side or single_w > max_weight:
        return None, "單件商品已超規，無法拆單"

    # 每包最多放幾件
    max_per_pkg_dim    = int(max_dim    // single_dim)  if single_dim  > 0 else 999
    max_per_pkg_side   = int(max_side   // single_side) if single_side > 0 else 999
    max_per_pkg_weight = int(max_weight // single_w)    if single_w    > 0 else 999
    max_per_pkg = min(max_per_pkg_dim, max_per_pkg_side, max_per_pkg_weight)
    max_per_pkg = max(1, max_per_pkg)

    # 需要幾包
    import math
    needed = math.ceil(total_items / max_per_pkg)

    if needed > max_units:
        return None, f"需要 {needed} 包，超過最多 {max_units} 單上限，無法自動拆單"

    # 建立拆單建議（同SKU盡量放同一包）
    packages = []
    current_pkg = []
    current_count = 0

    # 按SKU分組，同SKU連續放
    from collections import Counter
    sku_counts = Counter(all_items)
    sorted_skus = sorted(sku_counts.items(), key=lambda x: -x[1])

    remaining = []
    for sku, qty in sorted_skus:
        remaining.extend([sku] * qty)

    for item in remaining:
        if current_count >= max_per_pkg:
            packages.append(current_pkg[:])
            current_pkg = []
            current_count = 0
        current_pkg.append(item)
        current_count += 1

    if current_pkg:
        packages.append(current_pkg)

    result = []
    for i, pkg in enumerate(packages):
        pkg_counter = Counter(pkg)
        items_str = ", ".join(f"{sku}×{cnt}" for sku, cnt in sorted(pkg_counter.items()))
        pkg_dim    = single_dim * len(pkg)
        pkg_weight = single_w   * len(pkg)
        result.append({
            "pkg_no":  i + 1,
            "count":   len(pkg),
            "items":   items_str,
            "dim":     round(pkg_dim, 1),
            "weight":  round(pkg_weight, 2),
        })

    return result, ""

def _get_zone_label(locs):
    """取得區域描述字串，供宅配分群用"""
    whs   = set(wh for wh, z in locs)
    zones = set(z  for wh, z in locs if z != "?")
    if not zones: zones = {"?"}
    if len(whs) == 1 and len(zones) == 1:
        return f"{next(iter(whs))}{next(iter(zones))}區"
    elif len(whs) == 1:
        return f"{next(iter(whs))}混單"
    return "混單"

def get_sort_key(k):
    """排序：宅配→超商→店到店單項品→店到店→隔日配→無包裝→可拆單→超材"""
    if k == "__delivery__":     return (0,  k, k)
    if k == "__store_single__": return (25, k, k)
    if k == "__nopkg__":        return (88, k, k)
    if k == "__splittable__":   return (92, k, k)
    if k == "__oversize__":     return (95, k, k)
    order = {"超商":1, "店到店":2, "店到店隔日配":3}
    for label, pri in order.items():
        if k.startswith(label):
            if "混單" in k: return (pri*10+9, "zzz", k)
            return (pri*10, k, k)
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
                "weight":    0.0,
                "fee":       0.0,  # 運費
            }
        oid    = (row.get(CONFIG["col_order_id"], "") or "").strip()
        raw_wh = (row.get(CONFIG["col_warehouse"], "") or "").strip()
        sku    = (row.get(CONFIG["col_sku"], "") or "").strip()
        wh, zone = parse_location(raw_wh)

        # 尺寸取最大值（避免空值覆蓋有效值）
        L = safe_float(row.get(CONFIG["col_length"], 0))
        W = safe_float(row.get(CONFIG["col_width"],  0))
        H = safe_float(row.get(CONFIG["col_height"], 0))
        Wt= safe_float(row.get(CONFIG["col_weight"], 0))
        Fe= parse_fee(row.get(CONFIG["col_fee"], 0))
        if L > 0: order_map[txn]["length"]  = max(order_map[txn]["length"],  L)
        if W > 0: order_map[txn]["width"]   = max(order_map[txn]["width"],   W)
        if H > 0: order_map[txn]["height"]  = max(order_map[txn]["height"],  H)
        if Wt> 0: order_map[txn]["weight"]  = max(order_map[txn]["weight"],  Wt)
        if Fe> 0: order_map[txn]["fee"]     = max(order_map[txn]["fee"],     Fe)

        order_map[txn]["order_ids"].append(oid)
        order_map[txn]["products"].append({
            "oid": oid, "sku": sku,
            "zone_raw": raw_wh, "wh": wh, "zone": zone,
        })
        order_map[txn]["locations"].add((wh, zone))
        order_map[txn]["total_qty"] += 1

    # ── 計算三邊總和、最長邊，斜放判斷，超材檢查 ────────
    for txn, o in order_map.items():
        L, W, H = o["length"], o["width"], o["height"]

        # 斜放判斷：若訂單內有可斜放 SKU，用對角線取代最長邊
        skus_in_order = [p["sku"] for p in o["products"]]
        diagonal_applied = False
        effective_side = max(L, W, H)
        for sku in skus_in_order:
            eff, is_diag = apply_diagonal(sku, L, W, H)
            if is_diag:
                effective_side = eff
                diagonal_applied = True
                break

        dims = sorted([L, W, H], reverse=True)
        o["max_side"]       = effective_side
        o["total_dim"]      = sum(dims)
        o["diagonal_used"]  = diagonal_applied
        o["oversize"], o["oversize_msg"], o["can_split"], o["split_rules"] = check_oversize(
            o["channel"], o["ship_raw"], o["total_dim"], o["max_side"], o["weight"]
        )

        # 若斜放後不超材，標示說明
        if diagonal_applied and not o["oversize"]:
            o["oversize_msg"] = "斜放後合規"

        # 拆單建議（運費=0才可拆單）
        o["split_suggestion"] = None
        o["split_error"]      = ""
        o["fee_paid"]         = o["fee"] > 0

        # 運費>0 → 標示買家已付，不拆單
        if o["oversize"] and o["fee"] > 0:
            o["oversize_msg"] += f"｜買家已付運費 {o['fee']:.0f} 元"

        can_split = o["oversize"] and o["can_split"] and o["fee"] == 0
        if can_split and o["total_qty"] > 0:
            weight_per = o["weight"] / o["total_qty"] if o["total_qty"] > 0 else 0
            from collections import Counter
            sku_count = Counter(p["sku"] for p in o["products"])
            prods = [{"sku": sku, "qty": qty} for sku, qty in sku_count.items()]
            suggestion, err = suggest_split(prods, weight_per, L, W, H, o["split_rules"])
            o["split_suggestion"] = suggestion
            o["split_error"]      = err

    # ── 套用分單規則 ──────────────────────────────────
    groups = {}

    def add(key, title, icon, color, o):
        if key not in groups:
            groups[key] = {"title": title, "icon": icon, "color": color, "orders": []}
        groups[key]["orders"].append(o)

    summary = {}

    for txn, o in order_map.items():
        ch    = o["channel"]
        locs  = o["locations"]
        meta  = CHANNEL_META.get(ch, CHANNEL_META["delivery"])
        label = meta["label"]
        icon  = meta["icon"]
        color = meta["color"]

        # 超材處理
        if o["oversize"]:
            o["oversize_channel"] = label
            if o["can_split"] and o["split_suggestion"]:
                # 可拆單 → 獨立可拆單分類
                add("__splittable__", "✂ 可拆單", "✂", "#e65100", o)
                summary["✂ 可拆單"] = summary.get("✂ 可拆單", 0) + 1
            else:
                # 不可拆單 → 超材異常分類
                add("__oversize__", "⚠ 超材", "⚠", "#b71c1c", o)
                summary["⚠ 超材"] = summary.get("⚠ 超材", 0) + 1
            continue

        # 無包裝 → 全部合併成一組
        if ch == "nopkg":
            add("__nopkg__", "📦 無包裝", "📦", "#00838f", o)
            summary["無包裝"] = summary.get("無包裝", 0) + 1
            continue

        # 宅配（新竹/嘉里/店到家）→ 外層一個大分類，內部依區域分群
        if ch == "delivery":
            o["delivery_zone"] = _get_zone_label(locs)  # 記錄區域供分群用
            add("__delivery__", "🚚 宅配", "🚚", "#1565c0", o)
            summary["宅配"] = summary.get("宅配", 0) + 1
            continue

        # 店到店單項品 → 獨立一組（單一SKU不管幾件）
        skus = set(p["sku"] for p in o["products"] if p["sku"])
        if ch == "store" and len(skus) == 1:
            add("__store_single__", "🏬 店到店 ｜ 單項品", "🏬", "#1b5e20", o)
            summary["店到店"] = summary.get("店到店", 0) + 1
            continue

        # 其餘通路（超商、店到店多品、隔日配）→ 依倉庫區域細分
        whs   = set(wh for wh, z in locs)
        zones = set(z  for wh, z in locs if z != "?")
        if not zones:
            zones = {"?"}

        if len(whs) == 1 and len(zones) == 1:
            wh    = next(iter(whs))
            zone  = next(iter(zones))
            key   = f"{label} - {wh}{zone}區"
            title = f"{icon} {label} ｜ {wh} {zone} 區"
        elif len(whs) == 1:
            key   = f"{label} - {next(iter(whs))}混單"
            title = f"{icon} {label} ｜ {next(iter(whs))} 混單"
        else:
            key   = f"{label} - 混單"
            title = f"{icon} {label} ｜ 混單（跨倉）"

        if "混單" in key:
            color = color + "bb"

        add(key, title, icon, color, o)
        summary[label] = summary.get(label, 0) + 1

    state["summary"] = summary
    return dict(sorted(groups.items(), key=lambda x: get_sort_key(x[0])))

def load_csv(source, is_text=False):
    for enc in ["utf-8-sig", "big5", "cp950", "utf-8"]:
        try:
            if is_text:
                reader = csv.DictReader(io.StringIO(source))
            else:
                f = open(source, encoding=enc, newline="")
                reader = csv.DictReader(f)
            rows = list(reader)
            if not is_text: f.close()
            if rows: return rows, enc
        except Exception:
            continue
    return [], None

def run_pipeline(rows=None):
    state["status"] = "fetching"
    if not rows:
        state["status"] = "error"
        state["status_msg"] = "無資料"
        return
    groups = split_orders(rows)
    state["groups"]      = groups
    state["total"]       = len(set((r.get(CONFIG["col_txn"]) or "") for r in rows))
    state["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    state["status"]      = "ready"
    state["status_msg"]  = f"共 {state['total']} 張訂單，分成 {len(groups)} 組 | {state['last_update']}"
    for k, g in groups.items():
        log(f"  {g['title']}：{len(g['orders'])} 張")
    log("分單完成")

def scheduler():
    while True:
        now = datetime.now()
        if now.hour == CONFIG["auto_run_hour"] and now.minute == CONFIG["auto_run_min"]:
            log("定時觸發")
            time.sleep(61)
        time.sleep(30)

# ============================================================
# 設定檔持久化（斜放商品清單）
# ============================================================
def _get_base_dir():
    """取得執行檔所在目錄（支援 PyInstaller 打包）"""
    if getattr(sys, 'frozen', False):
        # 打包成 exe 後，exe 所在目錄
        return os.path.dirname(sys.executable)
    else:
        # 一般 Python 執行
        return os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(_get_base_dir(), "settings.json")

def load_settings():
    """從 settings.json 讀取設定，啟動時呼叫"""
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if "diagonal_skus" in data:
                CONFIG["diagonal_skus"] = data["diagonal_skus"]
                log(f"已載入特殊可出超材品設定：{len(CONFIG['diagonal_skus'])} 筆")
    except Exception as e:
        log(f"讀取設定檔失敗：{e}")

def save_settings():
    """將設定寫入 settings.json"""
    try:
        data = {"diagonal_skus": CONFIG["diagonal_skus"]}
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"儲存設定檔失敗：{e}")

# ============================================================
# Flask 網頁
# ============================================================
app = Flask(__name__)

# -*- coding: utf-8 -*-
"""
4Sale 自動分單印單系統 v4
五個通路：宅配 / 超商 / 店到店 / 店到店隔日配 / 無包裝
區域細分：倉庫(前倉/主倉/備用倉) + 區域字母(A/B/C...)
"""

import sys, csv, io, os, re, json, threading, time
from datetime import datetime
from flask import Flask, render_template_string, request, jsonify

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
    "col_txn":      "交易序號",
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
    max_units     = CONFIG["split_max_units"]
    max_dim       = rules["max_dim"]
    max_side      = rules["max_side"]
    max_weight    = rules["max_weight"]
    max_qty_limit = rules.get("max_qty")

    all_items = []
    for p in products:
        for _ in range(max(1, p.get("qty", 1))):
            all_items.append(p["sku"])

    total_items = len(all_items)
    sides       = sorted([L, W, H], reverse=True)
    single_dim  = sum(sides)
    single_side = sides[0]
    single_w    = weight_per_item

    if max_qty_limit:
        max_per_pkg = max_qty_limit
    else:
        if single_dim > max_dim or single_side > max_side or single_w > max_weight:
            return None, "單件商品已超規，無法拆單"
        max_per_pkg = min(
            int(max_dim    // single_dim)  if single_dim  > 0 else 999,
            int(max_side   // single_side) if single_side > 0 else 999,
            int(max_weight // single_w)    if single_w    > 0 else 999,
        )
        max_per_pkg = max(1, max_per_pkg)

    # 需要幾包
    import math
    needed = math.ceil(total_items / max_per_pkg)

    if needed > max_units:
        return None, f"需要 {needed} 包，超過最多 {max_units} 單上限，無法自動拆單"

    # 建立拆單建議（同SKU盡量放同一包）
    packages = []
    current_pkg = []
    current_count = 0

    # 按SKU分組，同SKU連續放
    from collections import Counter
    sku_counts = Counter(all_items)
    sorted_skus = sorted(sku_counts.items(), key=lambda x: -x[1])

    remaining = []
    for sku, qty in sorted_skus:
        remaining.extend([sku] * qty)

    for item in remaining:
        if current_count >= max_per_pkg:
            packages.append(current_pkg[:])
            current_pkg = []
            current_count = 0
        current_pkg.append(item)
        current_count += 1

    if current_pkg:
        packages.append(current_pkg)

    result = []
    for i, pkg in enumerate(packages):
        pkg_counter = Counter(pkg)
        items_str = ", ".join(f"{sku}×{cnt}" for sku, cnt in sorted(pkg_counter.items()))
        pkg_dim    = single_dim * len(pkg)
        pkg_weight = single_w   * len(pkg)
        result.append({
            "pkg_no":  i + 1,
            "count":   len(pkg),
            "items":   items_str,
            "dim":     round(pkg_dim, 1),
            "weight":  round(pkg_weight, 2),
        })

    return result, ""

def _get_zone_label(locs):
    """取得區域描述字串，供宅配分群用"""
    whs   = set(wh for wh, z in locs)
    zones = set(z  for wh, z in locs if z != "?")
    if not zones: zones = {"?"}
    if len(whs) == 1 and len(zones) == 1:
        return f"{next(iter(whs))}{next(iter(zones))}區"
    elif len(whs) == 1:
        return f"{next(iter(whs))}混單"
    return "混單"

def get_sort_key(k):
    """排序：宅配→超商→店到店單項品→店到店→隔日配→無包裝→巨無霸→可拆單→超材"""
    if k == "__delivery__":     return (0,  k, k)
    if k == "__store_single__": return (25, k, k)
    if k == "__nopkg__":        return (88, k, k)
    if k == "__giant__":        return (91, k, k)
    if k == "__splittable__":   return (92, k, k)
    if k == "__oversize__":     return (95, k, k)
    order = {"超商":1, "店到店":2, "店到店隔日配":3}
    for label, pri in order.items():
        if k.startswith(label):
            if "混單" in k: return (pri*10+9, "zzz", k)
            return (pri*10, k, k)
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
                "weight":    0.0,
                "fee":       0.0,  # 運費
            }
        oid    = (row.get(CONFIG["col_order_id"], "") or "").strip()
        raw_wh = (row.get(CONFIG["col_warehouse"], "") or "").strip()
        sku    = (row.get(CONFIG["col_sku"], "") or "").strip()
        wh, zone = parse_location(raw_wh)

        # 尺寸取最大值（避免空值覆蓋有效值）
        L = safe_float(row.get(CONFIG["col_length"], 0))
        W = safe_float(row.get(CONFIG["col_width"],  0))
        H = safe_float(row.get(CONFIG["col_height"], 0))
        Wt= safe_float(row.get(CONFIG["col_weight"], 0))
        Fe= parse_fee(row.get(CONFIG["col_fee"], 0))
        if L > 0: order_map[txn]["length"]  = max(order_map[txn]["length"],  L)
        if W > 0: order_map[txn]["width"]   = max(order_map[txn]["width"],   W)
        if H > 0: order_map[txn]["height"]  = max(order_map[txn]["height"],  H)
        if Wt> 0: order_map[txn]["weight"]  = max(order_map[txn]["weight"],  Wt)
        if Fe> 0: order_map[txn]["fee"]     = max(order_map[txn]["fee"],     Fe)

        order_map[txn]["order_ids"].append(oid)
        order_map[txn]["products"].append({
            "oid": oid, "sku": sku,
            "zone_raw": raw_wh, "wh": wh, "zone": zone,
        })
        order_map[txn]["locations"].add((wh, zone))
        order_map[txn]["total_qty"] += 1

    # ── 計算三邊總和、最長邊，斜放判斷，超材檢查 ────────
    for txn, o in order_map.items():
        L, W, H = o["length"], o["width"], o["height"]

        # 高度為 0 → ERP 無法輸入小數，自動補 0.2cm
        if H == 0 and (L > 0 or W > 0):
            H = 0.2
            o["height"] = 0.2
            o["height_auto"] = True
        else:
            o["height_auto"] = False

        # 斜放判斷：若訂單內有可斜放 SKU，用對角線取代最長邊
        dims = sorted([L, W, H], reverse=True)
        skus_in_order = [p["sku"] for p in o["products"]]
        diagonal_applied = False
        effective_side = dims[0] if dims else 0
        for sku in skus_in_order:
            eff, is_diag = apply_diagonal(sku, L, W, H)
            if is_diag:
                effective_side = eff
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
                max_q = get_sku_max_qty(sku, ch_group)
                if max_q and qty > max_q:
                    o["oversize"]     = True
                    o["can_split"]    = True
                    o["split_rules"]  = {
                        "max_dim":    CONFIG["split_max_dim"],
                        "max_side":   CONFIG["split_max_side"],
                        "max_weight": CONFIG["split_max_weight"],
                        "label":      f"每包最多{max_q}件",
                        "max_qty":    max_q,
                    }
                    o["oversize_msg"] = f"件數超限（{sku} 共{qty}件，每包上限{max_q}件）"
                    break

        # 拆單建議（運費=0才可拆單，店到店/店到家/隔日配）
        o["split_suggestion"] = None
        o["split_error"]      = ""
        o["fee_paid"]         = o["fee"] > 0

        # 運費>0 → 標示買家已付，不拆單
        if o["oversize"] and o["fee"] > 0:
            o["oversize_msg"] += f"｜買家已付運費 {o['fee']:.0f} 元"

        can_split = (
            o["oversize"] and
            o["can_split"] and
            o["fee"] == 0 and
            any(k in o["ship_raw"] for k in ["店到店", "店到家", "隔日到貨"])
        )
        if can_split and o["total_qty"] > 0:
            weight_per = o["weight"] / o["total_qty"] if o["total_qty"] > 0 else 0
            from collections import Counter
            sku_count = Counter(p["sku"] for p in o["products"])
            prods = [{"sku": sku, "qty": qty} for sku, qty in sku_count.items()]
            suggestion, err = suggest_split(prods, weight_per, L, W, H, o["split_rules"])
            o["split_suggestion"] = suggestion
            o["split_error"]      = err

    # ── 套用分單規則 ──────────────────────────────────
    groups = {}

    def add(key, title, icon, color, o):
        if key not in groups:
            groups[key] = {"title": title, "icon": icon, "color": color, "orders": []}
        groups[key]["orders"].append(o)

    summary = {}

    for txn, o in order_map.items():
        ch    = o["channel"]
        locs  = o["locations"]
        meta  = CHANNEL_META.get(ch, CHANNEL_META["delivery"])
        label = meta["label"]
        icon  = meta["icon"]
        color = meta["color"]

        # 超材處理
        if o["oversize"]:
            o["oversize_channel"] = label
            if o["can_split"] and o["split_suggestion"]:
                # 可拆單 → 獨立可拆單分類
                add("__splittable__", "✂ 可拆單", "✂", "#e65100", o)
                summary["✂ 可拆單"] = summary.get("✂ 可拆單", 0) + 1
            else:
                # 不可拆單 → 超材異常分類
                add("__oversize__", "⚠ 超材", "⚠", "#b71c1c", o)
                summary["⚠ 超材"] = summary.get("⚠ 超材", 0) + 1
            continue

        # 無包裝 → 全部合併成一組
        if ch == "nopkg":
            add("__nopkg__", "📦 無包裝", "📦", "#00838f", o)
            summary["無包裝"] = summary.get("無包裝", 0) + 1
            continue

        # 宅配（新竹/嘉里/店到家）→ 先判斷巨無霸，再進宅配大分類
        if ch == "delivery":
            is_giant = (
                (o["total_dim"] >= 170 or o["weight"] >= 15) and
                any(k in o["ship_raw"] for k in ["新竹物流", "嘉里"])
            )
            if is_giant:
                reasons = []
                if o["total_dim"] >= 170: reasons.append(f"三邊{o['total_dim']:.0f}cm")
                if o["weight"] >= 15:     reasons.append(f"重量{o['weight']:.1f}kg")
                o["oversize_channel"] = label
                o["oversize_msg"] = "巨無霸 " + "、".join(reasons) + " → 人工確認"
                add("__giant__", "&#129427; 巨無霸", "&#129427;", "#6a0dad", o)
                summary["&#129427; 巨無霸"] = summary.get("&#129427; 巨無霸", 0) + 1
                continue
            o["delivery_zone"] = _get_zone_label(locs)
            add("__delivery__", "&#128665; 宅配", "&#128665;", "#1565c0", o)
            summary["宅配"] = summary.get("宅配", 0) + 1
            continue

        # 店到店單項品 → 獨立一組（單一SKU不管幾件）
        skus = set(p["sku"] for p in o["products"] if p["sku"])
        if ch == "store" and len(skus) == 1:
            add("__store_single__", "🏬 店到店 ｜ 單項品", "🏬", "#1b5e20", o)
            summary["店到店"] = summary.get("店到店", 0) + 1
            continue

        # 其餘通路（超商、店到店多品、隔日配）→ 依倉庫區域細分
        whs   = set(wh for wh, z in locs)
        zones = set(z  for wh, z in locs if z != "?")
        if not zones:
            zones = {"?"}

        if len(whs) == 1 and len(zones) == 1:
            wh    = next(iter(whs))
            zone  = next(iter(zones))
            key   = f"{label} - {wh}{zone}區"
            title = f"{icon} {label} ｜ {wh} {zone} 區"
        elif len(whs) == 1:
            key   = f"{label} - {next(iter(whs))}混單"
            title = f"{icon} {label} ｜ {next(iter(whs))} 混單"
        else:
            key   = f"{label} - 混單"
            title = f"{icon} {label} ｜ 混單（跨倉）"

        if "混單" in key:
            color = color + "bb"

        add(key, title, icon, color, o)
        summary[label] = summary.get(label, 0) + 1

    state["summary"] = summary
    return dict(sorted(groups.items(), key=lambda x: get_sort_key(x[0])))

def load_csv(source, is_text=False):
    for enc in ["utf-8-sig", "big5", "cp950", "utf-8"]:
        try:
            if is_text:
                reader = csv.DictReader(io.StringIO(source))
            else:
                f = open(source, encoding=enc, newline="")
                reader = csv.DictReader(f)
            rows = list(reader)
            if not is_text: f.close()
            if rows: return rows, enc
        except Exception:
            continue
    return [], None

def run_pipeline(rows=None):
    state["status"] = "fetching"
    if not rows:
        state["status"] = "error"
        state["status_msg"] = "無資料"
        return
    groups = split_orders(rows)
    state["groups"]      = groups
    state["total"]       = len(set((r.get(CONFIG["col_txn"]) or "") for r in rows))
    state["last_update"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    state["status"]      = "ready"
    state["status_msg"]  = f"共 {state['total']} 張訂單，分成 {len(groups)} 組 | {state['last_update']}"
    for k, g in groups.items():
        log(f"  {g['title']}：{len(g['orders'])} 張")
    log("分單完成")

def scheduler():
    while True:
        now = datetime.now()
        if now.hour == CONFIG["auto_run_hour"] and now.minute == CONFIG["auto_run_min"]:
            log("定時觸發")
            time.sleep(61)
        time.sleep(30)

# ============================================================
# 設定檔持久化（斜放商品清單）
# ============================================================
def _get_base_dir():
    """取得執行檔所在目錄（支援 PyInstaller 打包）"""
    if getattr(sys, 'frozen', False):
        # 打包成 exe 後，exe 所在目錄
        return os.path.dirname(sys.executable)
    else:
        # 一般 Python 執行
        return os.path.dirname(os.path.abspath(__file__))

SETTINGS_FILE = os.path.join(_get_base_dir(), "settings.json")

def load_settings():
    """從 settings.json 讀取設定，啟動時呼叫"""
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if "diagonal_skus" in data:
                CONFIG["diagonal_skus"] = data["diagonal_skus"]
                log(f"已載入特殊可出超材品設定：{len(CONFIG['diagonal_skus'])} 筆")
    except Exception as e:
        log(f"讀取設定檔失敗：{e}")

def save_settings():
    """將設定寫入 settings.json"""
    try:
        data = {"diagonal_skus": CONFIG["diagonal_skus"]}
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"儲存設定檔失敗：{e}")

# ============================================================
# Flask 網頁
# ============================================================
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))

LOGIN_USER = os.environ.get("LOGIN_USER", "admin")
LOGIN_PASS = os.environ.get("LOGIN_PASS", "admin123")

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
<meta charset="UTF-8"><title>超人特工倉 - 首頁</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Microsoft JhengHei",sans-serif;background:#0f1923;min-height:100vh;color:#fff}
.topbar{background:rgba(255,255,255,.05);backdrop-filter:blur(10px);height:56px;padding:0 32px;display:flex;align-items:center;gap:12px;border-bottom:1px solid rgba(255,255,255,.08)}
.logo{font-size:16px;font-weight:700;margin-right:auto;letter-spacing:.5px}
.logo span{color:#f4a100}
.logout{color:#aaa;font-size:12px;text-decoration:none;padding:6px 12px;border:1px solid #333;border-radius:5px}
.logout:hover{border-color:#666;color:#fff}
.hero{text-align:center;padding:60px 20px 40px}
.hero h1{font-size:32px;font-weight:700;margin-bottom:8px}
.hero h1 span{color:#f4a100}
.hero p{font-size:14px;color:#888;letter-spacing:1px}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:20px;max-width:960px;margin:0 auto;padding:0 24px 60px}
.card{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.08);border-radius:16px;padding:32px 28px;text-decoration:none;color:#fff;transition:all .25s;cursor:pointer;position:relative;overflow:hidden}
.card:hover{background:rgba(255,255,255,.09);border-color:rgba(255,255,255,.2);transform:translateY(-4px);box-shadow:0 16px 40px rgba(0,0,0,.3)}
.card-icon{font-size:44px;margin-bottom:16px;display:block}
.card-title{font-size:18px;font-weight:700;margin-bottom:6px}
.card-desc{font-size:13px;color:#888;line-height:1.7}
.card-badge{position:absolute;top:16px;right:16px;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:500}
.badge-ready{background:rgba(46,125,50,.3);color:#81c784;border:1px solid rgba(46,125,50,.4)}
.badge-soon{background:rgba(100,100,100,.2);color:#888;border:1px solid rgba(100,100,100,.3)}
.card-split{border-color:rgba(26,95,168,.4)}
.card-split:hover{border-color:rgba(26,95,168,.8);box-shadow:0 16px 40px rgba(26,95,168,.15)}
.card-customs{border-color:rgba(244,161,0,.3)}
.card-customs:hover{border-color:rgba(244,161,0,.7);box-shadow:0 16px 40px rgba(244,161,0,.12)}
.card-tools{border-color:rgba(0,150,136,.3)}
.card-tools:hover{border-color:rgba(0,150,136,.7);box-shadow:0 16px 40px rgba(0,150,136,.12)}
</style></head><body>
<div class="topbar">
  <div class="logo">&#x1F3ED; <span>超人特工倉</span></div>
  <a href="/logout" class="logout">&#x1F6AA; 登出</a>
</div>
<div class="hero">
  <h1>歡迎回來，<span>特工！</span></h1>
  <p>SUPER WAREHOUSE AGENT SYSTEM &nbsp;|&nbsp; 選擇你的任務</p>
</div>
<div class="cards">
  <a href="/split" class="card card-split">
    <span class="card-badge badge-ready">&#x2713; 上線中</span>
    <span class="card-icon">&#x1F4E6;</span>
    <div class="card-title">分單中心</div>
    <div class="card-desc">上傳 4Sale 訂單 CSV，自動依通路和倉庫區域分單，一鍵複製交易序號到 4Sale 暫存區。</div>
  </a>
  <a href="/customs" class="card card-customs">
    <span class="card-badge badge-ready">&#x2713; 上線中</span>
    <span class="card-icon">&#x1F4CB;</span>
    <div class="card-title">報關助手</div>
    <div class="card-desc">上傳倉庫進貨清單，自動對應商品報關資料庫，帶入材質、品名、單價，一鍵匯出報關 Excel。</div>
  </a>
  <a href="/rack" class="card card-tools">
    <span class="card-badge badge-ready">&#x2713; 上線中</span>
    <span class="card-icon">&#x1F4E6;</span>
    <div class="card-title">貨架入庫</div>
    <div class="card-desc">掃描貨號入庫到重型貨架，記錄每個儲位的商品，一秒查詢貨號在哪個儲位。</div>
  </a>
</div>
</body></html>"""

def login_required(f):
    from functools import wraps
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect("/login")
        return f(*args, **kwargs)
    return decorated

@app.route("/login", methods=["GET", "POST"])
def login():
    error = ""
    if request.method == "POST":
        u = request.form.get("username", "")
        p = request.form.get("password", "")
        if u == LOGIN_USER and p == LOGIN_PASS:
            session["logged_in"] = True
            return redirect("/")
        error = "帳號或密碼錯誤，特工請再試一次！"
    return render_template_string(LOGIN_HTML, error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect("/login")

@app.route("/")
@login_required
def home():
    return render_template_string(HOME_HTML)

@app.route("/split")
@login_required
def split_page():
    return redirect("/split_app")



HTML = """<!DOCTYPE html>
<html lang="zh-TW"><head>
<meta charset="UTF-8">
<title>{{ company }} 分單系統</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"微軟正黑體","Microsoft JhengHei",sans-serif;background:#f0f2f5;font-size:13px;color:#1a1a1a}
.topbar{background:#0f1923;color:#fff;height:52px;padding:0 20px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:300}
.logo{font-size:15px;font-weight:600;margin-right:auto}
.logo span{color:#4fa3e0}
.btn{padding:7px 16px;border-radius:5px;border:none;font-size:13px;cursor:pointer;font-weight:500}
.btn-white{background:#fff;color:#0f1923}
.statusbar{background:#fff;border-bottom:1px solid #e0e0e0;padding:7px 20px;display:flex;align-items:center;gap:8px;font-size:12px}
.dot{width:8px;height:8px;border-radius:50%}
.dot-idle{background:#aaa}.dot-fetching{background:#f4a100}.dot-ready{background:#22c55e}.dot-error{background:#ef4444}

/* 通路統計卡片 */
.summary-bar{display:flex;gap:8px;padding:10px 20px;flex-wrap:wrap;background:#f8f8f8;border-bottom:1px solid #eee}
.sum-card{background:#fff;border:1px solid #e0e0e0;border-radius:6px;padding:6px 14px;font-size:12px;display:flex;align-items:center;gap:6px}
.sum-card .num{font-size:16px;font-weight:600;color:#1a1a1a}

/* 上傳區 */
.upload-center{padding:40px 20px;display:flex;justify-content:center}
.upload-box{background:#fff;border:2px dashed #c0cfe0;border-radius:12px;padding:40px 48px;text-align:center;max-width:460px;width:100%}
.upload-lbl{display:inline-block;background:#1a5fa8;color:#fff;padding:12px 32px;border-radius:6px;font-size:15px;font-weight:500;cursor:pointer;margin-bottom:8px}

.rebar{padding:7px 20px;background:#fff;border-bottom:1px solid #eee;display:flex;align-items:center;gap:8px}
.re-lbl{background:#f0f0f0;color:#444;padding:5px 12px;border-radius:5px;font-size:12px;cursor:pointer;font-weight:500}

/* Tabs */
.tabs{display:flex;gap:3px;padding:12px 20px 0;flex-wrap:wrap;background:#f0f2f5}
.tab{padding:5px 12px;border-radius:16px 16px 0 0;border:1px solid #d0d0d0;border-bottom:none;background:#e8e8e8;cursor:pointer;font-size:11px;color:#555;white-space:nowrap}
.tab.active{background:#fff;color:#0f1923;font-weight:500}
.tab:hover:not(.active){background:#f0f0f0}
.tab-sep{align-self:flex-end;color:#ccc;padding-bottom:3px;font-size:14px}

/* 主內容 */
.main{padding:0 20px 24px;background:#fff;border:1px solid #ddd;border-top:none;margin:0 20px;border-radius:0 0 8px 8px}
.grp{padding-top:16px}
.grp-hd{display:flex;align-items:center;gap:10px;padding:10px 16px;border-radius:6px 6px 0 0;color:#fff;font-size:13px;font-weight:500}
.gcnt{background:rgba(255,255,255,.25);padding:2px 10px;border-radius:12px;font-size:12px}
.copy-btn{margin-left:auto;display:flex;align-items:center;gap:5px;background:rgba(255,255,255,.2);border:1px solid rgba(255,255,255,.4);color:#fff;padding:4px 12px;border-radius:5px;font-size:12px;cursor:pointer;font-family:inherit;white-space:nowrap}
.copy-btn:hover{background:rgba(255,255,255,.35)}
.copy-btn.copied{background:rgba(255,255,255,.5);color:#1a1a1a}
table{width:100%;border-collapse:collapse;background:#fff;font-size:12px}
thead th{background:#f5f5f5;padding:7px 10px;text-align:left;font-weight:500;color:#555;border-bottom:1.5px solid #ddd;white-space:nowrap}
td{padding:6px 10px;border-bottom:.5px solid #efefef;vertical-align:top}
tr:last-child td{border-bottom:none}
tr.dr:hover td{background:#fafcff}
.mono{font-size:11px;color:#888;font-family:monospace}
.wh-tag{display:inline-block;font-size:10px;padding:1px 6px;border-radius:3px;background:#e8f5e9;color:#2e7d32;border:1px solid #c8e6c9;margin:1px}
.mix-tag{background:#fff3e0;color:#e65100;border-color:#ffe0b2}
.sum-row td{background:#f9f9f9;font-size:11px;color:#888;font-weight:500}
.oversize-row td{background:#fff8f8;}
.oversize-row:hover td{background:#ffefef!important}
.log-tog{margin:8px 20px;font-size:11px;color:#888;cursor:pointer}
.log-box{margin:0 20px 16px;background:#111;color:#7ec87e;font-family:monospace;font-size:11px;padding:10px;border-radius:6px;max-height:150px;overflow-y:auto;display:none}
.log-box.show{display:block}
.hidden{display:none}
@media print{
  .topbar,.statusbar,.summary-bar,.rebar,.tabs,.log-tog,.log-box{display:none!important}
  body,html{background:#fff!important}
  .main{border:none!important;margin:0!important;padding:0!important}
  .grp.hidden{display:block!important}
  .grp-hd{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  td,th{padding:4px 7px}
  .grp{page-break-inside:avoid;margin-bottom:10px}
}
</style></head><body>

<div class="topbar">
  <a href="/" style="color:#aaa;font-size:12px;text-decoration:none;margin-right:8px">&#x1F3ED;</a>
  <div class="logo"><span>分單中心</span></div>
  <span style="font-size:11px;color:#778">{{ last_update }}</span>
  <a href="/settings/diagonal" style="color:#aaa;font-size:12px;text-decoration:none;padding:5px 10px;border:1px solid #444;border-radius:5px">特殊可出超材品設定</a>
  <button class="btn btn-white" onclick="window.print()">列印全部</button>
  <a href="/" style="color:#aaa;font-size:12px;text-decoration:none;padding:5px 10px;border:1px solid #444;border-radius:5px">&#x2302; 首頁</a>
</div>

<div class="statusbar">
  <span class="dot dot-{{ status }}"></span>
  <span>{{ status_msg }}</span>
</div>

{% if summary %}
<div class="summary-bar">
  {% for label in summary_keys %}
  <div class="sum-card">
    <span>{{ label }}</span>
    <span class="num">{{ summary[label] }}</span>
    <span style="color:#aaa;font-size:11px">張</span>
  </div>
  {% endfor %}
  <div class="sum-card" style="margin-left:auto;background:#f5f5f5">
    <span>合計</span>
    <span class="num">{{ total }}</span>
    <span style="color:#aaa;font-size:11px">張</span>
  </div>
</div>
{% endif %}

{% if not groups %}
<div class="upload-center">
  <div class="upload-box">
    <div style="font-size:48px;margin-bottom:12px">&#128194;</div>
    <div style="font-size:17px;font-weight:600;margin-bottom:6px">上傳 4Sale 訂單 CSV</div>
    <div style="font-size:13px;color:#888;margin-bottom:24px;line-height:1.7">
      4Sale 後台 → 訂單管理 → 篩選可出貨 → 匯出 CSV<br>
      上傳後自動依通路及倉庫區域分單
    </div>
    <label class="upload-lbl" for="csv-in">選擇 CSV 檔案</label>
    <input type="file" id="csv-in" accept=".csv" style="display:none" onchange="doUpload()">
    <div id="up-name" style="font-size:12px;color:#aaa;margin-top:6px"></div>
    <div style="margin-top:16px;padding-top:14px;border-top:1px solid #eee;font-size:12px;color:#bbb">
      需包含欄位：交易序號、出貨類型、商品倉庫儲位、商品編號
    </div>
  </div>
</div>
{% else %}

<div class="rebar">
  <span style="font-size:12px;color:#888">重新上傳：</span>
  <label class="re-lbl" for="csv-in">換一份 CSV</label>
  <input type="file" id="csv-in" accept=".csv" style="display:none" onchange="doUpload()">
  <span id="up-name" style="font-size:12px;color:#aaa;margin-left:6px"></span>
</div>

{# 批次勾選複製工具列 #}
<div style="padding:8px 20px;background:#f0f4ff;border-bottom:1px solid #c5cae9;display:flex;align-items:center;gap:8px;flex-wrap:wrap">
  <span style="font-size:12px;font-weight:500;color:#3949ab;white-space:nowrap">批次複製：</span>
  {% for key in group_keys %}
  {% set g = groups[key] %}
  <label style="display:inline-flex;align-items:center;gap:3px;font-size:11px;cursor:pointer;white-space:nowrap;padding:2px 6px;border:1px solid #c5cae9;border-radius:12px;background:#fff">
    <input type="checkbox" class="grp-check" data-grpkey="{{ key }}" style="cursor:pointer">
    {% if key == '__oversize__' %}&#9888; 超材（{{ g.orders|length }}）
    {% elif key == '__delivery__' %}&#128665; 宅配（{{ g.orders|length }}）
    {% elif key == '__store_single__' %}店到店-單項品（{{ g.orders|length }}）
    {% elif key == '__nopkg__' %}&#128230; 無包裝（{{ g.orders|length }}）
    {% else %}{{ key }}（{{ g.orders|length }}）{% endif %}
  </label>
  {# 隱藏的序號資料 #}
  <span style="display:none" class="grp-txn-data" data-grpkey="{{ key }}">{% for o in g.orders %}{{ o.txn }}&#10;{% endfor %}</span>
  {% endfor %}
  <button onclick="copySelectedGrps()" style="margin-left:auto;background:#3949ab;color:#fff;border:none;padding:5px 14px;border-radius:5px;font-size:12px;cursor:pointer;font-weight:500;white-space:nowrap">&#128203; 複製已勾選</button>
  <span id="grp-copy-msg" style="font-size:11px;color:#3949ab"></span>
</div>

<div class="tabs">
  <div class="tab active" onclick="showGrp('all')" data-key="all">全部（{{ total }}）</div>
  {% set ns = namespace(prev_ch='') %}
  {% for key in group_keys %}
  {% set ch = key.split(' - ')[0] if ' - ' in key else key %}
  {% if ns.prev_ch != '' and ch != ns.prev_ch %}
  <span class="tab-sep">|</span>
  {% endif %}
  {% set ns.prev_ch = ch %}
  <div class="tab" onclick="showGrp('{{ key }}')" data-key="{{ key }}">
    {% if key == '__oversize__' %}&#9888; 超材（{{ groups[key].orders|length }}）
    {% elif key == '__giant__' %}&#129427; 巨無霸（{{ groups[key].orders|length }}）
    {% elif key == '__delivery__' %}&#128665; 宅配（{{ groups[key].orders|length }}）
    {% elif key == '__store_single__' %}&#128364; 店到店-單項品（{{ groups[key].orders|length }}）
    {% elif key == '__nopkg__' %}&#128230; 無包裝（{{ groups[key].orders|length }}）
    {% else %}{{ groups[key].icon | safe }} {{ key }}（{{ groups[key].orders|length }}）{% endif %}
  </div>
  {% endfor %}
</div>

<div class="main">
{% for key in group_keys %}
{% set g = groups[key] %}

{% if key == '__delivery__' %}
{# 宅配大分類：依區域分群顯示 #}
<div class="grp" data-key="{{ key }}">
  <div class="grp-hd" style="background:#1565c0">
    &#x1F69A; 宅配訂單
    <span class="gcnt">{{ g.orders|length }} 張</span>
    <button class="copy-btn" onclick="copyTxns(this)">&#128203; 複製交易序號</button>
    <span style="display:none" class="txn-data">{% for o in g.orders %}{{ o.txn }}&#10;{% endfor %}</span>
  </div>
  {% set zones_list = [] %}
  {% for o in g.orders %}{% if o.delivery_zone not in zones_list %}{% set _ = zones_list.append(o.delivery_zone) %}{% endif %}{% endfor %}

  {# 勾選複製工具列 #}
  <div style="padding:8px 16px;background:#e8eaf6;border-bottom:1px solid #c5cae9;display:flex;align-items:center;gap:10px;flex-wrap:wrap">
    <span style="font-size:12px;color:#3949ab;font-weight:500">勾選區域批次複製：</span>
    {% for zone_name in zones_list | sort %}
    <label style="display:inline-flex;align-items:center;gap:4px;font-size:12px;cursor:pointer;color:#333">
      <input type="checkbox" class="zone-check" data-zone="{{ zone_name }}" style="cursor:pointer">
      {{ zone_name }}
    </label>
    {% endfor %}
    <button onclick="copySelectedZones()" style="margin-left:auto;background:#3949ab;color:#fff;border:none;padding:5px 14px;border-radius:5px;font-size:12px;cursor:pointer;font-weight:500">&#128203; 複製已勾選</button>
    <span id="zone-copy-msg" style="font-size:11px;color:#3949ab"></span>
  </div>

  {% for zone_name in zones_list | sort %}
  {% set ns_zc = namespace(n=0) %}
  {% for o in g.orders %}{% if o.delivery_zone==zone_name %}{% set ns_zc.n=ns_zc.n+1 %}{% endif %}{% endfor %}
  <div style="padding:8px 16px;background:#e3f2fd;border-bottom:1px solid #bbdefb;display:flex;align-items:center;gap:10px">
    <span style="font-size:12px;font-weight:500;color:#1565c0">{{ zone_name }}（{{ ns_zc.n }} 張）</span>
    <button class="copy-btn" style="margin-left:auto;font-size:11px;padding:3px 10px"
      onclick="copyZoneTxns('{{ zone_name }}', this)">&#128203; 複製此區序號</button>
    <span style="display:none" class="zone-txn-data" data-zone="{{ zone_name }}">{% for o in g.orders %}{% if o.delivery_zone==zone_name %}{{ o.txn }}&#10;{% endif %}{% endfor %}</span>
  </div>
  <table>
    <thead><tr><th>#</th><th>交易序號</th><th>出貨類型</th><th>商品 / 儲位</th><th>尺寸/重量</th><th>件數</th></tr></thead>
    <tbody>
    {% set ns_di = namespace(i=0) %}
    {% for o in g.orders %}{% if o.delivery_zone == zone_name %}
    {% set ns_di.i = ns_di.i + 1 %}
    <tr class="dr">
      <td class="mono">{{ ns_di.i }}</td>
      <td class="mono" style="font-weight:500">{{ o.txn }}</td>
      <td style="font-size:11px;color:#666">{{ o.ship_raw }}{% if o.diagonal_used %}<div style="color:#1565c0;font-size:10px;margin-top:2px">特殊可出超材品</div>{% endif %}</td>
      <td>{% for p in o.products %}<div style="font-size:11px;padding:1px 0">{{ p.sku }}<span class="wh-tag">{{ p.zone_raw }}</span></div>{% endfor %}</td>
      <td style="font-size:11px;color:#555;white-space:nowrap">{% if o.total_dim > 0 %}<div>{{ o.length|int }}×{{ o.width|int }}×{{ o.height|int }}cm</div><div>三邊 {{ o.total_dim|int }}cm</div>{% endif %}{% if o.weight > 0 %}<div>{{ o.weight }}kg</div>{% endif %}</td>
      <td style="font-weight:600;color:#1a5fa8;text-align:center">{{ o.total_qty }}</td>
    </tr>
    {% endif %}{% endfor %}
    </tbody>
  </table>
  {% endfor %}
</div>

{% elif key == '__store_single__' %}
{# 店到店單項品 #}
<div class="grp" data-key="{{ key }}">
  <div class="grp-hd" style="background:#1b5e20">
    &#x1F3EC; 店到店 ｜ 單項品
    <span class="gcnt">{{ g.orders|length }} 張</span>
    <button class="copy-btn" onclick="copyTxns(this)">&#128203; 複製交易序號</button>
    <span style="display:none" class="txn-data">{% for o in g.orders %}{{ o.txn }}&#10;{% endfor %}</span>
  </div>
  <table>
    <thead><tr><th>#</th><th>交易序號</th><th>出貨類型</th><th>商品 / 儲位</th><th>尺寸/重量</th><th>件數</th></tr></thead>
    <tbody>
    {% for o in g.orders %}
    <tr class="dr">
      <td class="mono">{{ loop.index }}</td>
      <td class="mono" style="font-weight:500">{{ o.txn }}</td>
      <td style="font-size:11px;color:#666">{{ o.ship_raw }}</td>
      <td>{% for p in o.products %}<div style="font-size:11px;padding:1px 0">{{ p.sku }}<span class="wh-tag">{{ p.zone_raw }}</span></div>{% endfor %}</td>
      <td style="font-size:11px;color:#555;white-space:nowrap">{% if o.total_dim > 0 %}<div>{{ o.length|int }}×{{ o.width|int }}×{{ o.height|int }}cm</div><div>三邊 {{ o.total_dim|int }}cm</div>{% endif %}{% if o.weight > 0 %}<div>{{ o.weight }}kg</div>{% endif %}</td>
      <td style="font-weight:600;color:#1a5fa8;text-align:center">{{ o.total_qty }}</td>
    </tr>
    {% endfor %}
    <tr class="sum-row"><td colspan="2">小計 {{ g.orders|length }} 張</td><td colspan="4">{% set ns2=namespace(t=0) %}{% for o in g.orders %}{% set ns2.t=ns2.t+o.total_qty %}{% endfor %}共 {{ ns2.t }} 件</td></tr>
    </tbody>
  </table>
</div>

{% elif key == "__giant__" %}
{# 巨無霸分類：人工確認 #}
<div class="grp" data-key="{{ key }}">
  <div class="grp-hd" style="background:#6a0dad">
    &#129427; 巨無霸｜人工確認
    <span class="gcnt">{{ g.orders|length }} 張</span>
    <button class="copy-btn" onclick="copyTxns(this)">&#128203; 複製交易序號</button>
    <span style="display:none" class="txn-data">{% for o in g.orders %}{{ o.txn }}&#10;{% endfor %}</span>
  </div>
  <div style="padding:8px 16px;background:#f3e5f5;border-bottom:1px solid #ce93d8;font-size:12px;color:#6a0dad">
    &#9888; 三邊總和 &gt; 170cm 且重量 &gt; 15kg，請人工確認是否可出貨
  </div>
  <table>
    <thead><tr><th>#</th><th>交易序號</th><th>出貨類型</th><th>商品 / 儲位</th><th>尺寸/重量</th><th>件數</th></tr></thead>
    <tbody>
    {% for o in g.orders %}
    <tr class="dr" style="background:#fdf6ff">
      <td class="mono">{{ loop.index }}</td>
      <td class="mono" style="font-weight:500">{{ o.txn }}</td>
      <td style="font-size:11px;color:#666">
        {{ o.ship_raw }}
        <div style="color:#6a0dad;font-size:11px;font-weight:500;margin-top:2px">&#129427; {{ o.oversize_msg }}</div>
      </td>
      <td>{% for p in o.products %}<div style="font-size:11px;padding:1px 0">{{ p.sku }}<span class="wh-tag">{{ p.zone_raw }}</span></div>{% endfor %}</td>
      <td style="font-size:11px;color:#555;white-space:nowrap">
        {% if o.total_dim > 0 %}<div>{{ o.length|int }}x{{ o.width|int }}x{{ o.height|int }}cm</div><div>三邊 {{ o.total_dim|int }}cm</div>{% endif %}
        {% if o.weight > 0 %}<div style="color:#6a0dad;font-weight:500">{{ o.weight }}kg</div>{% endif %}
      </td>
      <td style="font-weight:600;color:#6a0dad;text-align:center">{{ o.total_qty }}</td>
    </tr>
    {% endfor %}
    <tr class="sum-row"><td colspan="2">小計 {{ g.orders|length }} 張</td><td colspan="4">{% set ns2=namespace(t=0) %}{% for o in g.orders %}{% set ns2.t=ns2.t+o.total_qty %}{% endfor %}共 {{ ns2.t }} 件</td></tr>
    </tbody>
  </table>
</div>

{% elif key == '__nopkg__' %}
{# 無包裝 #}
<div class="grp" data-key="{{ key }}">
  <div class="grp-hd" style="background:#00838f">
    &#x1F4E6; 無包裝
    <span class="gcnt">{{ g.orders|length }} 張</span>
    <button class="copy-btn" onclick="copyTxns(this)">&#128203; 複製交易序號</button>
    <span style="display:none" class="txn-data">{% for o in g.orders %}{{ o.txn }}&#10;{% endfor %}</span>
  </div>
  <table>
    <thead><tr><th>#</th><th>交易序號</th><th>出貨類型</th><th>商品 / 儲位</th><th>尺寸/重量</th><th>件數</th></tr></thead>
    <tbody>
    {% for o in g.orders %}
    <tr class="dr">
      <td class="mono">{{ loop.index }}</td>
      <td class="mono" style="font-weight:500">{{ o.txn }}</td>
      <td style="font-size:11px;color:#666">{{ o.ship_raw }}</td>
      <td>{% for p in o.products %}<div style="font-size:11px;padding:1px 0">{{ p.sku }}<span class="wh-tag">{{ p.zone_raw }}</span></div>{% endfor %}</td>
      <td style="font-size:11px;color:#555;white-space:nowrap">{% if o.total_dim > 0 %}<div>{{ o.length|int }}×{{ o.width|int }}×{{ o.height|int }}cm</div>{% endif %}{% if o.weight > 0 %}<div>{{ o.weight }}kg</div>{% endif %}</td>
      <td style="font-weight:600;color:#1a5fa8;text-align:center">{{ o.total_qty }}</td>
    </tr>
    {% endfor %}
    <tr class="sum-row"><td colspan="2">小計 {{ g.orders|length }} 張</td><td colspan="4">{% set ns2=namespace(t=0) %}{% for o in g.orders %}{% set ns2.t=ns2.t+o.total_qty %}{% endfor %}共 {{ ns2.t }} 件</td></tr>
    </tbody>
  </table>
</div>

{% elif key == '__splittable__' %}
{# 可拆單分類：依通路分群顯示 #}
<div class="grp" data-key="{{ key }}">
  <div class="grp-hd" style="background:#e65100">
    &#x2702; 可拆單訂單
    <span class="gcnt">{{ g.orders|length }} 張</span>
    <button class="copy-btn" onclick="copyTxns(this)">&#128203; 複製交易序號</button>
    <span style="display:none" class="txn-data">{% for o in g.orders %}{{ o.txn }}&#10;{% endfor %}</span>
  </div>
  {% set ch_list = [] %}
  {% for o in g.orders %}{% if o.oversize_channel not in ch_list %}{% set _ = ch_list.append(o.oversize_channel) %}{% endif %}{% endfor %}
  {% for ch_name in ch_list %}
  <div style="padding:8px 16px 0;background:#fff8f5;border-bottom:1px solid #ffccbc">
    <span style="font-size:12px;font-weight:500;color:#e65100">
      {{ ch_name }}（{% set ns_s=namespace(n=0) %}{% for o in g.orders %}{% if o.oversize_channel==ch_name %}{% set ns_s.n=ns_s.n+1 %}{% endif %}{% endfor %}{{ ns_s.n }} 張）
    </span>
  </div>
  <table>
    <thead><tr><th>#</th><th>交易序號</th><th>出貨類型</th><th>商品 / 儲位</th><th>尺寸/重量</th><th>件數</th></tr></thead>
    <tbody>
    {% set ns_si = namespace(i=0) %}
    {% for o in g.orders %}{% if o.oversize_channel == ch_name %}
    {% set ns_si.i = ns_si.i + 1 %}
    <tr class="dr oversize-row">
      <td class="mono">{{ ns_si.i }}</td>
      <td class="mono" style="font-weight:500">{{ o.txn }}</td>
      <td style="font-size:11px;color:#666">
        {{ o.ship_raw }}
        <div style="color:#e65100;font-size:11px;font-weight:500;margin-top:2px">&#x26A0; {{ o.oversize_msg }}</div>
        {% if o.split_suggestion %}
        <div style="margin-top:6px;padding:6px 8px;background:#fff3e0;border-radius:4px;font-size:11px">
          <div style="font-weight:500;color:#e65100;margin-bottom:3px">
            建議拆成 {{ o.split_suggestion|length }} 單
            {% if o.split_rules %}（依{{ o.split_rules.label }}規格）{% endif %}：
          </div>
          {% for pkg in o.split_suggestion %}
          <div style="color:#555;padding:1px 0">第{{ pkg.pkg_no }}包：{{ pkg.items }}（{{ pkg.count }}件 / {{ pkg.dim }}cm / {{ pkg.weight }}kg）</div>
          {% endfor %}
        </div>
        {% endif %}
        {% if o.split_error %}<div style="color:#888;font-size:10px;margin-top:3px">{{ o.split_error }}</div>{% endif %}
      </td>
      <td>{% for p in o.products %}<div style="font-size:11px;padding:1px 0">{{ p.sku }}<span class="wh-tag">{{ p.zone_raw }}</span></div>{% endfor %}</td>
      <td style="font-size:11px;color:#555;white-space:nowrap">
        {% if o.total_dim > 0 %}<div>{{ o.length|int }}×{{ o.width|int }}×{{ o.height|int }}cm</div><div>三邊 {{ o.total_dim|int }}cm</div>{% endif %}
        {% if o.weight > 0 %}<div>{{ o.weight }}kg</div>{% endif %}
      </td>
      <td style="font-weight:600;color:#e65100;text-align:center">{{ o.total_qty }}</td>
    </tr>
    {% endif %}{% endfor %}
    </tbody>
  </table>
  {% endfor %}
</div>

{% elif key == '__oversize__' %}
{# 超材大分類：依通路分群顯示 #}
<div class="grp" data-key="{{ key }}">
  <div class="grp-hd" style="background:#b71c1c">
    &#x26A0; 超材訂單
    <span class="gcnt">{{ g.orders|length }} 張</span>
    <button class="copy-btn" onclick="copyTxns(this)">&#128203; 複製交易序號</button>
    <span style="display:none" class="txn-data">{% for o in g.orders %}{{ o.txn }}&#10;{% endfor %}</span>
  </div>
  {# 依通路分群 #}
  {% set channels = [] %}
  {% for o in g.orders %}
    {% if o.oversize_channel not in channels %}{% set _ = channels.append(o.oversize_channel) %}{% endif %}
  {% endfor %}
  {% for ch_name in channels %}
  <div style="padding:10px 16px 0;background:#fff3f3;border-bottom:1px solid #ffcdd2">
    <span style="font-size:12px;font-weight:500;color:#b71c1c">{{ ch_name }}（{% set ns3=namespace(n=0) %}{% for o in g.orders %}{% if o.oversize_channel==ch_name %}{% set ns3.n=ns3.n+1 %}{% endif %}{% endfor %}{{ ns3.n }} 張）</span>
  </div>
  <table>
    <thead><tr>
      <th>#</th><th>交易序號</th><th>出貨類型</th>
      <th>商品 / 儲位</th><th>尺寸/重量</th><th>件數</th>
    </tr></thead>
    <tbody>
    {% set ns_idx = namespace(i=0) %}
    {% for o in g.orders %}{% if o.oversize_channel == ch_name %}
    {% set ns_idx.i = ns_idx.i + 1 %}
    <tr class="dr oversize-row">
      <td class="mono">{{ ns_idx.i }}</td>
      <td class="mono" style="font-weight:500">{{ o.txn }}</td>
      <td style="font-size:11px;color:#666">
        {{ o.ship_raw }}
        <div style="color:#b71c1c;font-size:11px;font-weight:500;margin-top:2px">&#x26A0; {{ o.oversize_msg }}</div>
        {% if o.fee_paid %}
        <div style="color:#e65100;font-size:11px;margin-top:2px">&#x1F4B0; 買家已付運費 {{ o.fee|int }} 元，不拆單</div>
        {% endif %}
        {% if o.split_suggestion %}
        <div style="margin-top:6px;padding:6px 8px;background:#fff3e0;border-radius:4px;font-size:11px">
          <div style="font-weight:500;color:#e65100;margin-bottom:3px">建議拆成 {{ o.split_suggestion|length }} 單：</div>
          {% for pkg in o.split_suggestion %}
          <div style="color:#555;padding:1px 0">第{{ pkg.pkg_no }}包：{{ pkg.items }}（{{ pkg.count }}件 / {{ pkg.dim }}cm / {{ pkg.weight }}kg）</div>
          {% endfor %}
        </div>
        {% endif %}
        {% if o.split_error %}<div style="color:#888;font-size:10px;margin-top:3px">{{ o.split_error }}</div>{% endif %}
      </td>
      <td>{% for p in o.products %}<div style="font-size:11px;padding:1px 0">{{ p.sku }}<span class="wh-tag">{{ p.zone_raw }}</span></div>{% endfor %}</td>
      <td style="font-size:11px;color:#555;white-space:nowrap">
        {% if o.total_dim > 0 %}<div>{{ o.length|int }}×{{ o.width|int }}×{{ o.height|int }}cm</div><div>三邊 {{ o.total_dim|int }}cm</div>{% endif %}
        {% if o.weight > 0 %}<div>{{ o.weight }}kg</div>{% endif %}
      </td>
      <td style="font-weight:600;color:#b71c1c;text-align:center">{{ o.total_qty }}</td>
    </tr>
    {% endif %}{% endfor %}
    </tbody>
  </table>
  {% endfor %}
</div>

{% else %}
{# 一般分組 #}
<div class="grp" data-key="{{ key }}">
  <div class="grp-hd" style="background:{{ g.color }}">
    {{ g.title | safe }}
    <span class="gcnt">{{ g.orders|length }} 張</span>
    <button class="copy-btn" onclick="copyTxns(this)">&#128203; 複製交易序號</button>
    <span style="display:none" class="txn-data">{% for o in g.orders %}{{ o.txn }}&#10;{% endfor %}</span>
  </div>
  <table>
    <thead><tr>
      <th>#</th><th>交易序號</th><th>出貨類型</th>
      <th>商品 / 儲位</th><th>尺寸/重量</th><th>件數</th>
    </tr></thead>
    <tbody>
    {% for o in g.orders %}
    <tr class="dr">
      <td class="mono">{{ loop.index }}</td>
      <td class="mono" style="font-weight:500">{{ o.txn }}</td>
      <td style="font-size:11px;color:#666">
        {{ o.ship_raw }}
        {% if o.diagonal_used and not o.oversize %}
        <div style="color:#1565c0;font-size:10px;margin-top:2px">斜放後合規</div>
        {% endif %}
      </td>
      <td>{% for p in o.products %}<div style="font-size:11px;padding:1px 0">{{ p.sku }}<span class="wh-tag">{{ p.zone_raw }}</span></div>{% endfor %}</td>
      <td style="font-size:11px;color:#555;white-space:nowrap">
        {% if o.total_dim > 0 %}<div>{{ o.length|int }}×{{ o.width|int }}×{{ o.height|int }}cm</div><div>三邊 {{ o.total_dim|int }}cm</div>{% endif %}
        {% if o.weight > 0 %}<div>{{ o.weight }}kg</div>{% endif %}
      </td>
      <td style="font-weight:600;color:#1a5fa8;text-align:center">{{ o.total_qty }}</td>
    </tr>
    {% endfor %}
    <tr class="sum-row">
      <td colspan="2">小計 {{ g.orders|length }} 張</td>
      <td colspan="4">
        {% set ns2 = namespace(t=0) %}
        {% for o in g.orders %}{% set ns2.t = ns2.t + o.total_qty %}{% endfor %}
        共 {{ ns2.t }} 件
      </td>
    </tr>
    </tbody>
  </table>
</div>
{% endif %}

{% endfor %}
</div>
{% endif %}

<div class="log-tog" onclick="document.getElementById('lg').classList.toggle('show')">顯示/隱藏記錄</div>
<div class="log-box" id="lg">{% for l in log_lines %}<div>{{ l }}</div>{% endfor %}</div>

<script>
function showGrp(key){
  document.querySelectorAll('.tab').forEach(t=>t.classList.toggle('active',t.dataset.key===key));
  document.querySelectorAll('.grp').forEach(s=>s.classList.toggle('hidden',key!=='all'&&s.dataset.key!==key));
}
function copySelectedGrps() {
  var NL = String.fromCharCode(10);
  var checked = document.querySelectorAll('.grp-check:checked');
  if (checked.length === 0) { alert('請先勾選要複製的分組'); return; }
  var all = [];
  checked.forEach(function(cb) {
    var el = document.querySelector('.grp-txn-data[data-grpkey="' + cb.dataset.grpkey + '"]');
    if (el) {
      var parts = el.textContent.trim().split(NL);
      parts.forEach(function(l){ if(l.trim()) all.push(l.trim()); });
    }
  });
  var text = all.join(NL);
  navigator.clipboard.writeText(text).then(function() {
    var msg = document.getElementById('grp-copy-msg');
    if (msg) { msg.textContent = '已複製 ' + all.length + ' 筆（' + checked.length + ' 個分組）'; setTimeout(function(){ msg.textContent=''; }, 3000); }
  }).catch(function() {
    var ta = document.createElement('textarea');
    ta.value = text; document.body.appendChild(ta); ta.select();
    document.execCommand('copy'); document.body.removeChild(ta);
    var msg = document.getElementById('grp-copy-msg');
    if (msg) { msg.textContent = '已複製 ' + all.length + ' 筆'; setTimeout(function(){ msg.textContent=''; }, 3000); }
  });
}

function copyZoneTxns(zone, btn) {
  const el = document.querySelector('.zone-txn-data[data-zone="' + zone + '"]');
  if (!el) return;
  const txns = el.textContent.trim();
  navigator.clipboard.writeText(txns).then(() => {
    const orig = btn.innerHTML; btn.textContent = '已複製！'; btn.classList.add('copied');
    setTimeout(() => { btn.innerHTML = orig; btn.classList.remove('copied'); }, 2000);
  }).catch(() => {
    const ta = document.createElement('textarea');
    ta.value = txns; document.body.appendChild(ta); ta.select();
    document.execCommand('copy'); document.body.removeChild(ta);
    btn.textContent = '已複製！';
    setTimeout(() => { btn.innerHTML = '&#128203; 複製此區序號'; }, 2000);
  });
}

function copySelectedZones() {
  var NL = String.fromCharCode(10);
  var checked = document.querySelectorAll('.zone-check:checked');
  if (checked.length === 0) { alert('請先勾選要複製的區域'); return; }
  var all = [];
  checked.forEach(function(cb) {
    var el = document.querySelector('.zone-txn-data[data-zone="' + cb.dataset.zone + '"]');
    if (el) {
      var parts = el.textContent.trim().split(NL);
      parts.forEach(function(l){ if(l.trim()) all.push(l.trim()); });
    }
  });
  var text = all.join(NL);
  navigator.clipboard.writeText(text).then(function() {
    var msg = document.getElementById('zone-copy-msg');
    if (msg) { msg.textContent = '已複製 ' + all.length + ' 筆（' + checked.length + ' 個區域）'; setTimeout(function(){ msg.textContent=''; }, 3000); }
  }).catch(function() {
    var ta = document.createElement('textarea');
    ta.value = text; document.body.appendChild(ta); ta.select();
    document.execCommand('copy'); document.body.removeChild(ta);
  });
}

function copyTxns(btn){
  const txns = btn.parentElement.querySelector('.txn-data').textContent.trim();
  navigator.clipboard.writeText(txns).then(()=>{
    btn.textContent = '已複製！';
    btn.classList.add('copied');
    setTimeout(()=>{
      btn.innerHTML = '&#128203; 複製交易序號';
      btn.classList.remove('copied');
    }, 2000);
  }).catch(()=>{
    // 備用方案：舊瀏覽器
    const ta = document.createElement('textarea');
    ta.value = txns;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    btn.textContent = '已複製！';
    setTimeout(()=>{ btn.innerHTML = '&#128203; 複製交易序號'; }, 2000);
  });
}
function doUpload(){
  const f=document.getElementById('csv-in').files[0];
  if(!f)return;
  document.getElementById('up-name').textContent='上傳中：'+f.name;
  const fd=new FormData();fd.append('file',f);
  fetch('/api/upload',{method:'POST',body:fd})
    .then(r=>r.json())
    .then(d=>{if(d.ok)location.reload();else alert('失敗：'+d.msg)});
}
const lg=document.getElementById('lg');
if(lg)lg.scrollTop=lg.scrollHeight;
</script>
</body></html>"""

@app.route("/split_app")
@login_required
def index():
    try:
        return render_template_string(HTML,
            company     = CONFIG["company_name"],
            groups      = state["groups"],
            group_keys  = list(state["groups"].keys()),
            total       = state["total"],
            last_update = state["last_update"] or "—",
            status      = state["status"],
            status_msg  = state["status_msg"],
            summary     = state["summary"],
            summary_keys= list(state["summary"].keys()),
            log_lines   = state["log"][-30:],
        )
    except Exception as e:
        log(f"頁面渲染錯誤：{e}")
        return f"<html><body><h2>系統啟動中，請重新整理</h2><p>{e}</p></body></html>", 500

@app.route("/api/upload", methods=["POST"])
@login_required
def api_upload():
    f = request.files.get("file")
    if not f: return jsonify({"ok": False, "msg": "無檔案"})
    content = None
    for enc in ["utf-8-sig", "big5", "cp950", "utf-8"]:
        try:
            f.seek(0); content = f.read().decode(enc); break
        except Exception: continue
    if not content: return jsonify({"ok": False, "msg": "無法解析編碼"})
    rows, enc = load_csv(content, is_text=True)
    if not rows: return jsonify({"ok": False, "msg": "CSV 解析失敗"})
    run_pipeline(rows=rows)
    return jsonify({"ok": True})

@app.route("/api/status")
@login_required
def api_status():
    return jsonify({k: state[k] for k in ("status","status_msg","last_update","total")})

# ── 特殊可出超材品設定頁面 ──────────────────────────────────────────
DIAGONAL_HTML = """<!DOCTYPE html>
<html lang="zh-TW"><head>
<meta charset="UTF-8">
<title>特殊可出超材品設定</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Microsoft JhengHei",sans-serif;background:#f0f2f5;font-size:13px;color:#1a1a1a}
.topbar{background:#0f1923;color:#fff;height:52px;padding:0 20px;display:flex;align-items:center;gap:12px;position:sticky;top:0;z-index:300}
.logo{font-size:15px;font-weight:600;margin-right:auto}.logo span{color:#4fa3e0}
.btn{padding:7px 16px;border-radius:5px;border:none;font-size:13px;cursor:pointer;font-weight:500}
.btn-blue{background:#1a5fa8;color:#fff}.btn-red{background:#b71c1c;color:#fff}.btn:hover{opacity:.85}
.card{background:#fff;border:1px solid #ddd;border-radius:8px;padding:20px;margin:20px}
.card h2{font-size:14px;font-weight:500;color:#555;margin-bottom:14px;padding-bottom:8px;border-bottom:1px solid #eee}
.desc{font-size:12px;color:#888;margin-bottom:14px;line-height:1.8}
.form-grid{display:grid;grid-template-columns:160px 140px 1fr auto;gap:10px;align-items:end;margin-bottom:8px}
label{font-size:12px;color:#666;display:block;margin-bottom:4px}
input[type=text],input[type=number]{width:100%;padding:7px 10px;border:1px solid #ddd;border-radius:5px;font-size:13px;font-family:inherit}
input:focus{outline:none;border-color:#1a5fa8}
.ch-group{padding:8px 12px;background:#f8f8f8;border:1px solid #ddd;border-radius:5px}
.ch-group label{display:inline-flex;align-items:center;gap:5px;margin-right:16px;font-size:12px;color:#333;cursor:pointer}
.ch-group label input{width:auto}
table{width:100%;border-collapse:collapse;font-size:12px}
thead th{background:#f5f5f5;padding:8px 10px;text-align:left;font-weight:500;color:#555;border-bottom:1.5px solid #ddd}
td{padding:8px 10px;border-bottom:.5px solid #eee}
.tag{display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;background:#e3f2fd;color:#1565c0;border:1px solid #bbdefb}
.tag-g{background:#e8f5e9;color:#2e7d32;border:1px solid #a5d6a7}
.tag-o{background:#fff3e0;color:#e65100;border:1px solid #ffcc80}
.empty{text-align:center;padding:30px;color:#aaa;font-size:12px}
.msg{padding:8px 14px;border-radius:5px;font-size:12px;margin:0 20px 10px}
.msg-ok{background:#e8f5e9;color:#2e7d32}.msg-err{background:#ffebee;color:#b71c1c}
</style></head><body>
<div class="topbar">
  <div class="logo"><span>特殊可出超材品</span> 設定</div>
  <a href="/" style="color:#aaa;font-size:12px;text-decoration:none">&#x2302; 返回首頁</a>
</div>
<div id="msg-area"></div>
<div class="card">
  <h2>新增設定</h2>
  <div class="desc">
    <b>斜放後有效最長邊：</b>商品斜放後最長維度（cm），讓系統用此值判斷超材。<br>
    <b>每包最大件數：</b>超過此件數視為超材，可針對不同通路群組生效。<br>
    兩個欄位至少填一個。
  </div>
  <div class="form-grid">
    <div><label>商品 SKU</label><input type="text" id="new-sku" placeholder="例：APB002"></div>
    <div><label>斜放後最長邊（cm）</label><input type="number" id="new-side" placeholder="留空不設定" min="1" max="200" step="0.5"></div>
    <div>
      <label>每包最大件數 &amp; 適用通路</label>
      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">
        <input type="number" id="new-qty" placeholder="件數上限（留空不限）" min="1" step="1" style="width:180px">
        <div class="ch-group">
          <label><input type="checkbox" id="ch-store" checked> 店配類<small style="color:#888">（店到店/隔日/超商/店到家）</small></label>
          <label><input type="checkbox" id="ch-delivery" checked> 快遞類<small style="color:#888">（新竹/嘉里）</small></label>
        </div>
      </div>
    </div>
    <div style="padding-bottom:2px"><button class="btn btn-blue" onclick="addSku()">新增</button></div>
  </div>
</div>
<div class="card">
  <h2>目前設定清單</h2>
  <table>
    <thead><tr>
      <th>商品 SKU</th><th>斜放後最長邊</th><th>每包最大件數</th><th>適用通路</th><th style="width:60px"></th>
    </tr></thead>
    <tbody id="sku-tbody"><tr><td colspan="5" class="empty">尚無設定</td></tr></tbody>
  </table>
</div>
<script>
var CH_LABELS = {store:'店配類', delivery:'快遞類'};
function loadSkus() {
  fetch('/api/diagonal').then(function(r){return r.json();}).then(function(data) {
    var tbody = document.getElementById('sku-tbody');
    var keys = data ? Object.keys(data) : [];
    if (keys.length === 0) { tbody.innerHTML = '<tr><td colspan="5" class="empty">尚無設定</td></tr>'; return; }
    tbody.innerHTML = keys.map(function(sku) {
      var cfg = data[sku];
      var side = cfg.side ? '<strong>' + cfg.side + ' cm</strong>' : '<span style="color:#bbb">—</span>';
      var qty  = cfg.max_qty ? '<span class="tag-o">' + cfg.max_qty + ' 件</span>' : '<span style="color:#bbb">不限</span>';
      var chs  = (cfg.channels || []).map(function(c){ return '<span class="tag-g">' + (CH_LABELS[c]||c) + '</span>'; }).join(' ');
      if (!chs) chs = '<span style="color:#bbb">—</span>';
      return '<tr><td><span class="tag">' + sku + '</span></td><td>' + side + '</td><td>' + qty + '</td><td>' + chs + '</td>' +
             '<td><button class="btn btn-red" style="padding:4px 10px;font-size:11px" onclick="deleteSku(\'' + sku + '\')">刪除</button></td></tr>';
    }).join('');
  });
}
function showMsg(msg, ok) {
  var area = document.getElementById('msg-area');
  area.innerHTML = '<div class="msg ' + (ok?'msg-ok':'msg-err') + '">' + msg + '</div>';
  setTimeout(function(){ area.innerHTML = ''; }, 3000);
}
function addSku() {
  var sku  = document.getElementById('new-sku').value.trim();
  var side = document.getElementById('new-side').value.trim();
  var qty  = document.getElementById('new-qty').value.trim();
  var chs  = [];
  if (document.getElementById('ch-store').checked)    chs.push('store');
  if (document.getElementById('ch-delivery').checked) chs.push('delivery');
  if (!sku) { showMsg('請輸入 SKU', false); return; }
  if (!side && !qty) { showMsg('請至少填入斜放最長邊或件數上限', false); return; }
  if (qty && chs.length === 0) { showMsg('請至少勾選一個通路群組', false); return; }
  fetch('/api/diagonal', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({sku:sku, side:side?parseFloat(side):null, max_qty:qty?parseInt(qty):null, channels:chs})
  }).then(function(r){return r.json();}).then(function(d) {
    if (d.ok) {
      showMsg('已新增 ' + sku, true);
      document.getElementById('new-sku').value = '';
      document.getElementById('new-side').value = '';
      document.getElementById('new-qty').value = '';
      loadSkus();
    } else { showMsg('新增失敗：' + d.msg, false); }
  });
}
function deleteSku(sku) {
  if (!confirm('確定刪除 ' + sku + '？')) return;
  fetch('/api/diagonal/' + sku, {method:'DELETE'}).then(function(r){return r.json();}).then(function(d) {
    if (d.ok) { showMsg('已刪除 ' + sku, true); loadSkus(); }
  });
}
loadSkus();
</script>
</body></html>"""

@app.route("/settings/diagonal")
@login_required
def settings_diagonal():
    return render_template_string(DIAGONAL_HTML)

@app.route("/api/diagonal", methods=["GET"])
@login_required
def api_diagonal_get():
    return jsonify(CONFIG["diagonal_skus"])

@app.route("/api/diagonal", methods=["POST"])
@login_required
def api_diagonal_post():
    data     = request.get_json()
    sku      = (data.get("sku") or "").strip()
    side     = data.get("side")
    max_qty  = data.get("max_qty")
    channels = data.get("channels", [])
    if not sku:
        return jsonify({"ok": False, "msg": "SKU 不可為空"})
    if not side and not max_qty:
        return jsonify({"ok": False, "msg": "請至少填入一個設定"})
    try:
        CONFIG["diagonal_skus"][sku] = {
            "side":     float(side) if side else None,
            "max_qty":  int(max_qty) if max_qty else None,
            "channels": channels,  # ["store", "delivery"]
        }
        save_settings()
        log(f"新增特殊可出超材品：{sku} side={side} max_qty={max_qty} channels={channels}")
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)})

@app.route("/api/diagonal/<sku>", methods=["DELETE"])
@login_required
def api_diagonal_delete(sku):
    if sku in CONFIG["diagonal_skus"]:
        del CONFIG["diagonal_skus"][sku]
        save_settings()
        log(f"刪除特殊可出超材品設定：{sku}")
    return jsonify({"ok": True})

# ============================================================
# 報關模組
# ============================================================

def get_sheets_client():
    """建立 Google Sheets 連線"""
    try:
        import gspread, base64, re
        from google.oauth2.service_account import Credentials

        cred_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT", "")
        if not cred_json:
            return None, "未設定 GOOGLE_SERVICE_ACCOUNT 環境變數"

        cred_dict = None

        # 方法1: Base64 解碼
        try:
            cred_dict = json.loads(base64.b64decode(cred_json + "==").decode("utf-8"))
        except Exception:
            pass

        # 方法2: 強制修復換行後直接解析
        if not cred_dict:
            try:
                fixed = cred_json.replace(chr(13)+chr(10), chr(92)+"n").replace(chr(13), chr(92)+"n").replace(chr(10), chr(92)+"n")
                cred_dict = json.loads(fixed)
            except Exception:
                pass

        # 方法3: 直接解析
        if not cred_dict:
            try:
                cred_dict = json.loads(cred_json)
            except Exception as e:
                return None, f"JSON 解析失敗: {str(e)[:100]}"

        if not cred_dict:
            return None, "無法解析 GOOGLE_SERVICE_ACCOUNT"

        scopes = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_info(cred_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client, None
    except Exception as e:
        return None, str(e)


def load_customs_db():
    """從 Google Sheets 載入商品報關資料庫，回傳 {sku: {...}} dict"""
    client, err = get_sheets_client()
    if err:
        return {}, err
    try:
        sheet_id = os.environ.get("GOOGLE_SHEETS_ID")
        sh = client.open_by_key(sheet_id)
        ws = sh.worksheet("商品報關資料庫")

        # 用 FORMULA 模式讀取，才能拿到 =IMAGE("url") 公式字串
        # get_all_records() 只讀顯示值，無法讀到公式
        header_row = ws.row_values(1)
        all_values = ws.get_all_values(value_render_option='FORMULA')
        col_idx = {name: i for i, name in enumerate(header_row)}

        def get_cell(row_data, col_name, default=""):
            idx = col_idx.get(col_name)
            if idx is None or idx >= len(row_data):
                return default
            return str(row_data[idx]).strip()

        db = {}
        for row_data in all_values[1:]:  # 跳過標題列
            sku = get_cell(row_data, "SKU編碼")
            if sku:
                db[sku] = {
                    "sku":          sku,
                    "name":         get_cell(row_data, "系統名稱"),
                    "style":        get_cell(row_data, "樣式"),
                    "size":         get_cell(row_data, "尺寸"),
                    "material":     get_cell(row_data, "材質"),
                    "customs_name": get_cell(row_data, "報關品名"),
                    "price":        get_cell(row_data, "單價"),
                    "unit":         get_cell(row_data, "單位"),
                    "product_size": get_cell(row_data, "商品尺寸"),
                    "image":        parse_image_url(get_cell(row_data, "圖片")),
                }
        return db, None
    except Exception as e:
        return {}, str(e)

def append_to_customs_db(new_item):
    """新增一筆資料到 Google Sheets 商品報關資料庫"""
    client, err = get_sheets_client()
    if err:
        return err
    try:
        sheet_id = os.environ.get("GOOGLE_SHEETS_ID")
        sh = client.open_by_key(sheet_id)
        ws = sh.worksheet("商品報關資料庫")
        ws.append_row([
            new_item.get("sku", ""),
            new_item.get("name", ""),
            new_item.get("style", ""),
            new_item.get("size", ""),
            new_item.get("material", ""),
            new_item.get("customs_name", ""),
            new_item.get("price", ""),
            new_item.get("unit", ""),
            new_item.get("product_size", ""),
            new_item.get("image", ""),
        ])
        return None
    except Exception as e:
        return str(e)

# ============================================================
# 貨架入庫模組
# ============================================================

def load_rack_sheet():
    """取得貨架庫位紀錄 worksheet，不存在則自動建立"""
    client, err = get_sheets_client()
    if err:
        return None, err
    try:
        sheet_id = os.environ.get("GOOGLE_SHEETS_ID")
        sh = client.open_by_key(sheet_id)
        try:
            ws = sh.worksheet("貨架庫位紀錄")
        except Exception:
            ws = sh.add_worksheet(title="貨架庫位紀錄", rows=1000, cols=6)
            ws.append_row(["貨號", "儲位", "數量", "入庫時間", "備註"])
        return ws, None
    except Exception as e:
        return None, str(e)

def rack_save_records(items, rack_code):
    """items = [{"sku": "BCE001-001", "qty": 2}], rack_code = "RACK-A-1" """
    ws, err = load_rack_sheet()
    if err:
        return err
    try:
        now = datetime.now().strftime("%Y/%m/%d %H:%M")
        rows = [[item.get("sku",""), rack_code, item.get("qty",1), now, ""] for item in items]
        if rows:
            ws.append_rows(rows)
        return None
    except Exception as e:
        return str(e)

def rack_query_sku(sku):
    """查詢貨號在哪些儲位"""
    ws, err = load_rack_sheet()
    if err:
        return [], err
    try:
        all_values = ws.get_all_values()
        results = []
        for row in all_values[1:]:
            if len(row) < 2:
                continue
            if sku.upper() in str(row[0]).strip().upper():
                results.append({
                    "sku":  row[0] if len(row)>0 else "",
                    "rack": row[1] if len(row)>1 else "",
                    "qty":  row[2] if len(row)>2 else "",
                    "time": row[3] if len(row)>3 else "",
                    "note": row[4] if len(row)>4 else "",
                })
        return results, None
    except Exception as e:
        return [], str(e)

def rack_query_rack(rack_code):
    """查詢某儲位裡有哪些貨號"""
    ws, err = load_rack_sheet()
    if err:
        return [], err
    try:
        all_values = ws.get_all_values()
        results = []
        for row in all_values[1:]:
            if len(row) < 2:
                continue
            if str(row[1]).strip().upper() == rack_code.strip().upper():
                results.append({
                    "sku":  row[0] if len(row)>0 else "",
                    "rack": row[1] if len(row)>1 else "",
                    "qty":  row[2] if len(row)>2 else "",
                    "time": row[3] if len(row)>3 else "",
                    "note": row[4] if len(row)>4 else "",
                })
        return results, None
    except Exception as e:
        return [], str(e)

RACK_HTML = """<!DOCTYPE html>
<html lang="zh-TW"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>貨架入庫系統</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Microsoft JhengHei",sans-serif;background:#0f1923;min-height:100vh;color:#fff;font-size:14px}
.topbar{background:rgba(255,255,255,.05);backdrop-filter:blur(10px);height:52px;padding:0 20px;display:flex;align-items:center;gap:12px;border-bottom:1px solid rgba(255,255,255,.08);position:sticky;top:0;z-index:300}
.logo{font-size:15px;font-weight:700;margin-right:auto}.logo span{color:#f4a100}
a.back{color:#aaa;font-size:12px;text-decoration:none;padding:5px 10px;border:1px solid #333;border-radius:5px}
a.back:hover{border-color:#666;color:#fff}
.tabs{display:flex;border-bottom:2px solid rgba(255,255,255,.08);padding:0 20px}
.tab{padding:14px 24px;cursor:pointer;font-size:13px;color:#888;border-bottom:2px solid transparent;margin-bottom:-2px;transition:all .2s}
.tab.active{color:#f4a100;border-bottom-color:#f4a100;font-weight:600}
.tab:hover{color:#fff}
.panel{display:none;padding:24px 20px;max-width:800px;margin:0 auto}
.panel.active{display:block}
.card{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:20px;margin-bottom:16px}
.card h3{font-size:14px;font-weight:600;margin-bottom:14px;color:#f4a100}
input[type=text],input[type=number]{width:100%;padding:10px 14px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);border-radius:8px;color:#fff;font-size:14px;font-family:inherit;outline:none}
input:focus{border-color:#f4a100;background:rgba(244,161,0,.08)}
input::placeholder{color:#555}
.btn{padding:10px 20px;border-radius:8px;border:none;font-size:13px;cursor:pointer;font-weight:600;transition:all .2s;white-space:nowrap}
.btn-yellow{background:#f4a100;color:#0f1923}.btn-yellow:hover{background:#ffb300}
.btn-green{background:#2e7d32;color:#fff}.btn-green:hover{background:#388e3c}
.btn-red{background:#b71c1c;color:#fff}.btn-red:hover{background:#c62828}
.scan-list{list-style:none;margin:12px 0}
.scan-item{display:flex;align-items:center;gap:10px;padding:10px 14px;background:rgba(255,255,255,.06);border-radius:8px;margin-bottom:6px;border:1px solid rgba(255,255,255,.08)}
.scan-sku{flex:1;font-weight:600;font-size:15px;letter-spacing:.5px}
.scan-qty{width:70px;padding:5px 8px;text-align:center}
.scan-del{cursor:pointer;color:#e57373;font-size:18px;background:none;border:none;padding:0 4px}
.confirm-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.88);z-index:999;align-items:center;justify-content:center;flex-direction:column;gap:20px;text-align:center}
.confirm-overlay.show{display:flex}
.confirm-rack{font-size:80px;font-weight:900;color:#f4a100;letter-spacing:6px;text-shadow:0 0 60px rgba(244,161,0,.6);animation:pulse 1s ease-in-out infinite alternate}
@keyframes pulse{from{transform:scale(1)}to{transform:scale(1.06)}}
.confirm-sub{font-size:24px;color:#fff;opacity:.7;letter-spacing:2px}
.confirm-items{background:rgba(255,255,255,.08);border-radius:12px;padding:16px 28px;min-width:340px;max-width:520px;text-align:left}
.confirm-items li{padding:7px 0;border-bottom:1px solid rgba(255,255,255,.08);font-size:15px}
.confirm-items li:last-child{border:none}
.confirm-btns{display:flex;gap:14px;margin-top:4px}
.result-table{width:100%;border-collapse:collapse;margin-top:12px}
.result-table th{background:rgba(244,161,0,.15);padding:10px 12px;text-align:left;font-size:12px;color:#f4a100;border-bottom:1px solid rgba(244,161,0,.2)}
.result-table td{padding:10px 12px;border-bottom:1px solid rgba(255,255,255,.06);font-size:13px}
.result-table tr:hover td{background:rgba(255,255,255,.04)}
.rack-badge{display:inline-block;padding:3px 10px;background:rgba(244,161,0,.15);color:#f4a100;border-radius:20px;font-weight:700;font-size:13px;border:1px solid rgba(244,161,0,.3)}
.msg{padding:10px 14px;border-radius:8px;font-size:13px;margin-bottom:12px}
.msg-ok{background:rgba(46,125,50,.2);color:#81c784;border:1px solid rgba(46,125,50,.3)}
.msg-err{background:rgba(183,28,28,.2);color:#ef9a9a;border:1px solid rgba(183,28,28,.3)}
.empty{text-align:center;padding:32px;color:#555;font-size:13px}
.row{display:flex;gap:10px}
</style></head><body>

<div class="topbar">
  <div class="logo">&#x1F4E6; <span>貨架入庫系統</span></div>
  <a href="/" class="back">&#x2302; 返回首頁</a>
  <a href="/logout" class="back">登出</a>
</div>

<div class="tabs">
  <div class="tab active" onclick="switchTab('inbound',this)">&#x1F4E5; 入庫作業</div>
  <div class="tab" onclick="switchTab('search',this)">&#x1F50D; 查找貨位</div>
</div>

<!-- 入庫作業 -->
<div class="panel active" id="panel-inbound">
  <div id="msg-inbound"></div>
  <div class="card">
    <h3>&#x1F4F7; 步驟一：掃描 / 輸入貨號</h3>
    <input type="text" id="sku-input" placeholder="掃描或輸入貨號，按 Enter 加入清單" autocomplete="off" autofocus>
    <ul class="scan-list" id="scan-list">
      <li class="empty" id="scan-empty">尚未掃描任何貨號</li>
    </ul>
  </div>
  <div class="card">
    <h3>&#x1F4CD; 步驟二：掃描 / 輸入儲位條碼</h3>
    <div class="row">
      <input type="text" id="rack-input" placeholder="掃描或輸入儲位，例：RACK-A-1" autocomplete="off">
      <button class="btn btn-yellow" onclick="confirmRack()">&#x2705; 確認入庫</button>
    </div>
    <div style="font-size:12px;color:#555;margin-top:8px">掃描儲位後會出現大字二次確認，避免刷錯儲位</div>
  </div>
</div>

<!-- 查找貨位 -->
<div class="panel" id="panel-search">
  <div id="msg-search"></div>
  <div class="card">
    <h3>&#x1F50D; 查詢貨號所在儲位</h3>
    <div class="row">
      <input type="text" id="search-sku" placeholder="輸入或掃描貨號" autocomplete="off">
      <button class="btn btn-yellow" onclick="searchSku()">查詢</button>
    </div>
    <div id="result-sku"></div>
  </div>
  <div class="card">
    <h3>&#x1F4E6; 查詢儲位內有哪些貨號</h3>
    <div class="row">
      <input type="text" id="search-rack" placeholder="輸入或掃描儲位條碼，例：RACK-A-1" autocomplete="off">
      <button class="btn btn-yellow" onclick="searchRack()">查詢</button>
    </div>
    <div id="result-rack"></div>
  </div>
</div>

<!-- 大字二次確認 -->
<div class="confirm-overlay" id="confirm-overlay">
  <div class="confirm-rack" id="confirm-rack-text"></div>
  <div class="confirm-sub" id="confirm-rack-sub"></div>
  <div class="confirm-items">
    <div style="font-size:13px;color:#aaa;margin-bottom:10px">即將入庫以下貨號：</div>
    <ul id="confirm-item-list"></ul>
  </div>
  <div class="confirm-btns">
    <button class="btn btn-green" onclick="doSave()">&#x2705; 確認正確，儲存</button>
    <button class="btn btn-red" onclick="cancelConfirm()">&#x274C; 取消重掃</button>
  </div>
</div>

<script>
var scannedItems = {};
var pendingRack = '';

function switchTab(name, el){
  document.querySelectorAll('.tab').forEach(function(t){t.classList.remove('active');});
  document.querySelectorAll('.panel').forEach(function(p){p.classList.remove('active');});
  el.classList.add('active');
  document.getElementById('panel-'+name).classList.add('active');
  if(name==='inbound') document.getElementById('sku-input').focus();
  else document.getElementById('search-sku').focus();
}

document.getElementById('sku-input').addEventListener('keydown',function(e){
  if(e.key!=='Enter') return;
  var sku = this.value.trim().toUpperCase();
  if(!sku) return;
  scannedItems[sku] = (scannedItems[sku]||0)+1;
  renderScanList();
  this.value=''; this.focus();
});

document.getElementById('rack-input').addEventListener('keydown',function(e){
  if(e.key==='Enter') confirmRack();
});
document.getElementById('search-sku').addEventListener('keydown',function(e){
  if(e.key==='Enter') searchSku();
});
document.getElementById('search-rack').addEventListener('keydown',function(e){
  if(e.key==='Enter') searchRack();
});

function renderScanList(){
  var ul=document.getElementById('scan-list');
  var keys=Object.keys(scannedItems);
  document.getElementById('scan-empty').style.display=keys.length?'none':'block';
  Array.from(ul.querySelectorAll('.scan-item')).forEach(function(el){el.remove();});
  keys.forEach(function(sku){
    var li=document.createElement('li'); li.className='scan-item';
    li.innerHTML='<span class="scan-sku">'+sku+'</span>'+
      '<input type="number" class="scan-qty" value="'+scannedItems[sku]+'" min="1" onchange="updateQty(\''+sku+'\',this.value)">'+
      '<span style="font-size:12px;color:#666">件</span>'+
      '<button class="scan-del" onclick="removeSku(\''+sku+'\')">&#x2715;</button>';
    ul.appendChild(li);
  });
}

function updateQty(sku,val){var n=parseInt(val);if(n>0)scannedItems[sku]=n;else removeSku(sku);}
function removeSku(sku){delete scannedItems[sku];renderScanList();}

function confirmRack(){
  if(!Object.keys(scannedItems).length){showMsg('inbound','&#x26A0; 請先掃描至少一個貨號！',false);return;}
  var rack=document.getElementById('rack-input').value.trim().toUpperCase();
  if(!rack){showMsg('inbound','&#x26A0; 請掃描或輸入儲位條碼！',false);return;}
  pendingRack=rack;
  document.getElementById('confirm-rack-text').textContent=rack;
  var parts=rack.split('-');
  var sub=parts.length>=3?parts[1]+' 排　第 '+parts[2]+' 層':rack;
  document.getElementById('confirm-rack-sub').textContent=sub;
  var ul=document.getElementById('confirm-item-list'); ul.innerHTML='';
  Object.keys(scannedItems).forEach(function(sku){
    var li=document.createElement('li');
    li.textContent=sku+'　×　'+scannedItems[sku]+' 件';
    ul.appendChild(li);
  });
  document.getElementById('confirm-overlay').classList.add('show');
}

function cancelConfirm(){
  document.getElementById('confirm-overlay').classList.remove('show');
  document.getElementById('rack-input').focus();
}

function doSave(){
  var items=Object.keys(scannedItems).map(function(sku){return{sku:sku,qty:scannedItems[sku]};});
  fetch('/api/rack/save',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({items:items,rack:pendingRack})
  }).then(function(r){return r.json();}).then(function(d){
    document.getElementById('confirm-overlay').classList.remove('show');
    if(d.ok){
      showMsg('inbound','&#x2705; 入庫成功！共 '+items.length+' 種貨號 → '+pendingRack,true);
      scannedItems={};renderScanList();
      document.getElementById('rack-input').value='';
      document.getElementById('sku-input').focus();
    } else {
      showMsg('inbound','&#x274C; 儲存失敗：'+d.msg,false);
    }
  });
}

function searchSku(){
  var sku=document.getElementById('search-sku').value.trim();
  if(!sku){showMsg('search','請輸入貨號',false);return;}
  document.getElementById('result-sku').innerHTML='<div style="color:#888;padding:10px">查詢中...</div>';
  fetch('/api/rack/query-sku?sku='+encodeURIComponent(sku))
    .then(function(r){return r.json();}).then(function(d){
      if(!d.ok){document.getElementById('result-sku').innerHTML='<div class="msg msg-err">'+d.msg+'</div>';return;}
      if(!d.results.length){document.getElementById('result-sku').innerHTML='<div class="empty">&#x1F4ED; 找不到此貨號的入庫紀錄</div>';return;}
      var h='<table class="result-table"><tr><th>貨號</th><th>儲位</th><th>數量</th><th>入庫時間</th></tr>';
      d.results.forEach(function(r){h+='<tr><td>'+r.sku+'</td><td><span class="rack-badge">'+r.rack+'</span></td><td>'+r.qty+'</td><td>'+r.time+'</td></tr>';});
      document.getElementById('result-sku').innerHTML=h+'</table>';
    });
}

function searchRack(){
  var rack=document.getElementById('search-rack').value.trim();
  if(!rack){showMsg('search','請輸入儲位條碼',false);return;}
  document.getElementById('result-rack').innerHTML='<div style="color:#888;padding:10px">查詢中...</div>';
  fetch('/api/rack/query-rack?rack='+encodeURIComponent(rack))
    .then(function(r){return r.json();}).then(function(d){
      if(!d.ok){document.getElementById('result-rack').innerHTML='<div class="msg msg-err">'+d.msg+'</div>';return;}
      if(!d.results.length){document.getElementById('result-rack').innerHTML='<div class="empty">&#x1F4ED; 此儲位目前沒有入庫紀錄</div>';return;}
      var h='<table class="result-table"><tr><th>貨號</th><th>儲位</th><th>數量</th><th>入庫時間</th></tr>';
      d.results.forEach(function(r){h+='<tr><td>'+r.sku+'</td><td><span class="rack-badge">'+r.rack+'</span></td><td>'+r.qty+'</td><td>'+r.time+'</td></tr>';});
      document.getElementById('result-rack').innerHTML=h+'</table>';
    });
}

function showMsg(panel,msg,ok){
  var el=document.getElementById('msg-'+panel);
  el.innerHTML='<div class="msg '+(ok?'msg-ok':'msg-err')+'">'+msg+'</div>';
  setTimeout(function(){el.innerHTML='';},4000);
}
</script>
</body></html>"""

@app.route("/rack")
@login_required
def rack_page():
    return render_template_string(RACK_HTML)

@app.route("/api/rack/save", methods=["POST"])
@login_required
def api_rack_save():
    data = request.get_json()
    items = data.get("items", [])
    rack  = data.get("rack", "").strip().upper()
    if not items or not rack:
        return jsonify({"ok": False, "msg": "缺少貨號或儲位"})
    err = rack_save_records(items, rack)
    if err:
        return jsonify({"ok": False, "msg": err})
    log(f"貨架入庫：{len(items)} 種貨號 → {rack}")
    return jsonify({"ok": True})

@app.route("/api/rack/query-sku")
@login_required
def api_rack_query_sku():
    sku = request.args.get("sku", "").strip()
    if not sku:
        return jsonify({"ok": False, "msg": "請提供貨號"})
    results, err = rack_query_sku(sku)
    if err:
        return jsonify({"ok": False, "msg": err})
    return jsonify({"ok": True, "results": results})

@app.route("/api/rack/query-rack")
@login_required
def api_rack_query_rack():
    rack = request.args.get("rack", "").strip()
    if not rack:
        return jsonify({"ok": False, "msg": "請提供儲位條碼"})
    results, err = rack_query_rack(rack)
    if err:
        return jsonify({"ok": False, "msg": err})
    return jsonify({"ok": True, "results": results})


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
label{font-size:12px;color:#666;display:block;margin-bottom:4px}
input[type=text],input[type=number],select{width:100%;padding:7px 10px;border:1px solid #ddd;border-radius:5px;font-size:13px;font-family:inherit}
input:focus,select:focus{outline:none;border-color:#f4a100}
table{width:100%;border-collapse:collapse;font-size:12px}
thead th{background:#f5f5f5;padding:8px 6px;text-align:left;font-weight:500;color:#555;border-bottom:1.5px solid #ddd;white-space:nowrap}
td{padding:6px;border-bottom:.5px solid #eee;vertical-align:middle}
tr.ok{background:#f1f8e9}
tr.missing{background:#ffebee}
tr.missing td:first-child::after{content:" &#128560;";color:#b71c1c}
.tag-ok{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;background:#e8f5e9;color:#2e7d32}
.tag-miss{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;background:#ffebee;color:#b71c1c}
.msg{padding:8px 14px;border-radius:5px;font-size:12px;margin:10px 20px}
.msg-ok{background:#e8f5e9;color:#2e7d32}.msg-err{background:#ffebee;color:#b71c1c}
.msg-warn{background:#fff8e1;color:#e65100}
.modal{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);z-index:999;align-items:center;justify-content:center}
.modal.show{display:flex}
.modal-box{background:#fff;border-radius:10px;padding:28px;width:520px;max-height:90vh;overflow-y:auto}
.modal-box h3{font-size:15px;font-weight:600;margin-bottom:16px;color:#e65100}
.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px}
.form-group{display:flex;flex-direction:column;gap:4px}
img.thumb{width:50px;height:50px;object-fit:cover;border-radius:4px;border:1px solid #eee}
#preview-section{display:none}
</style></head><body>

<div class="topbar">
  <div class="logo"><span>&#128230;</span> 報關清單系統</div>
  <a href="/" style="color:#aaa;font-size:12px;text-decoration:none">&#x2302; 返回首頁</a>
  <a href="/logout" style="color:#aaa;font-size:12px;text-decoration:none">登出</a>
</div>

<div id="msg-area"></div>

<div class="card">
  <h2>&#128228; 上傳倉庫進貨清單</h2>
  <p style="font-size:12px;color:#888;margin-bottom:16px">
    請先在 Excel 的 <strong>R 欄</strong>填入 SKU 編碼，再上傳檔案。<br>
    系統會自動對應商品報關資料庫，帶入材質、品名、單價等資料。
  </p>
  <div class="upload-area" id="drop-zone" onclick="document.getElementById('xlsx-in').click()">
    <div style="font-size:36px;margin-bottom:8px">&#128196;</div>
    <div style="font-weight:500;margin-bottom:4px">點擊或拖曳上傳 Excel 檔案</div>
    <div style="font-size:12px;color:#aaa">支援 .xlsx 及 .xls 格式，圖片會自動忽略不影響上傳</div>
  </div>
  <input type="file" id="xlsx-in" accept=".xlsx,.xls" style="display:none" onchange="uploadFile(this)">
</div>

<div id="preview-section">
  <!-- 櫃號/封號/日期填寫區 -->
  <div class="card">
    <h2>&#128230; 填寫報關資訊</h2>
    <p style="font-size:12px;color:#888;margin-bottom:14px">此資訊會顯示在匯出的報關單標題區</p>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
      <div>
        <label style="font-size:12px;color:#666;display:block;margin-bottom:4px">出口商（A1）</label>
        <input type="text" id="exporter" placeholder="例：ORAL TRADING CO., LTD." style="width:100%;padding:8px 10px;border:1.5px solid #ddd;border-radius:6px;font-size:13px">
      </div>
      <div>
        <label style="font-size:12px;color:#666;display:block;margin-bottom:4px">進口商（A2）</label>
        <input type="text" id="importer" placeholder="例：台灣進口商名稱" style="width:100%;padding:8px 10px;border:1.5px solid #ddd;border-radius:6px;font-size:13px">
      </div>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px">
      <div>
        <label style="font-size:12px;color:#666;display:block;margin-bottom:4px">櫃號</label>
        <input type="text" id="cabinet-no" placeholder="例：TCKU3456789" style="width:100%;padding:8px 10px;border:1.5px solid #ddd;border-radius:6px;font-size:13px">
      </div>
      <div>
        <label style="font-size:12px;color:#666;display:block;margin-bottom:4px">封號</label>
        <input type="text" id="seal-no" placeholder="例：CN12345678" style="width:100%;padding:8px 10px;border:1.5px solid #ddd;border-radius:6px;font-size:13px">
      </div>
      <div>
        <label style="font-size:12px;color:#666;display:block;margin-bottom:4px">出貨日期</label>
        <input type="date" id="ship-date" style="width:100%;padding:8px 10px;border:1.5px solid #ddd;border-radius:6px;font-size:13px">
      </div>
    </div>
  </div>
  <div class="card">
    <h2>&#128270; 預覽結果</h2>
    <div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">
      <span id="stat-total" style="font-size:12px;color:#555"></span>
      <span id="stat-ok" style="font-size:12px;color:#2e7d32"></span>
      <span id="stat-miss" style="font-size:12px;color:#b71c1c"></span>
    </div>
    <div style="overflow-x:auto">
      <table id="preview-table">
        <thead><tr>
          <th>SKU</th><th>類型</th><th>產品尺寸</th><th>材質</th><th>報關品名</th>
          <th>箱號</th><th>PCS/件</th><th>件數</th><th>總PCS</th><th>單位</th>
          <th>單價</th><th>總金額RMB</th><th>毛重</th><th>長</th><th>寬</th><th>高</th>
          <th>材積</th><th>總重量</th><th>圖片</th><th>狀態</th>
        </tr></thead>
        <tbody id="preview-body"></tbody>
      </table>
    </div>
    <div style="margin-top:16px;display:flex;gap:10px">
      <button class="btn btn-green" onclick="exportExcel()">&#128229; 匯出報關 Excel</button>
      <button class="btn btn-gray" onclick="resetPage()">&#128260; 重新上傳</button>
    </div>
  </div>
</div>


<!-- 新品建檔 Modal -->
<div class="modal" id="new-item-modal">
  <div class="modal-box">
    <h3>&#127381; 發現新商品！請建立報關資料</h3>
    <p style="font-size:12px;color:#888;margin-bottom:16px">
      SKU「<strong id="modal-sku"></strong>」在資料庫中查無資料，<br>
      請填入報關資料後儲存，系統會自動更新商品報關資料庫。
    </p>
    <div class="form-grid">
      <div class="form-group"><label>SKU編碼</label><input type="text" id="f-sku" readonly style="background:#f5f5f5"></div>
      <div class="form-group"><label>系統名稱</label><input type="text" id="f-name" placeholder="例：針織手提袋"></div>
      <div class="form-group"><label>樣式</label><input type="text" id="f-style" placeholder="例：格子款"></div>
      <div class="form-group"><label>尺寸</label><input type="text" id="f-size" placeholder="例：F"></div>
      <div class="form-group"><label>材質 *</label><input type="text" id="f-material" placeholder="例：滌綸"></div>
      <div class="form-group"><label>報關品名 *</label><input type="text" id="f-customs-name" placeholder="例：手提袋"></div>
      <div class="form-group"><label>單價 *</label><input type="number" id="f-price" placeholder="例：3.4" step="0.01"></div>
      <div class="form-group"><label>單位</label><input type="text" id="f-unit" placeholder="例：個"></div>
      <div class="form-group"><label>商品尺寸</label><input type="text" id="f-product-size" placeholder="例：20*35*10"></div>
      <div class="form-group"><label>圖片網址</label><input type="text" id="f-image" placeholder="貼上圖片URL（選填）"></div>
    </div>
    <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:8px">
      <button class="btn btn-gray" onclick="skipNewItem()">跳過此筆</button>
      <button class="btn btn-yellow" onclick="saveNewItem()">&#128190; 儲存並繼續</button>
    </div>
  </div>
</div>

<script>
var allRows = [];
var missingSkus = [];
var currentMissingIdx = 0;
var customsDb = {};
var processedRows = [];

// 拖曳上傳
var dz = document.getElementById('drop-zone');
dz.addEventListener('dragover', function(e){ e.preventDefault(); dz.classList.add('drag'); });
dz.addEventListener('dragleave', function(){ dz.classList.remove('drag'); });
dz.addEventListener('drop', function(e){
  e.preventDefault(); dz.classList.remove('drag');
  var f = e.dataTransfer.files[0];
  if(f) processFile(f);
});

function uploadFile(input) {
  var f = input.files[0];
  if(f) processFile(f);
}

function processFile(file) {
  if(!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
    showMsg('&#128561; 哎呀！這不是 Excel 檔案，請選擇 .xlsx 或 .xls 格式！', false);
    return;
  }
  // 檔案大小警告（超過 2MB）
  if(file.size > 2 * 1024 * 1024) {
    var sizeMB = (file.size / 1024 / 1024).toFixed(1);
    var NL = String.fromCharCode(10);
    if(!confirm('&#9888; 您的檔案有 ' + sizeMB + ' MB，可能包含圖片！' + NL + NL + '建議先在 Excel 刪除圖片再上傳，可節省流量費用。' + NL + NL + '要繼續上傳嗎？')) {
      return;
    }
  }
  showMsg('&#8987; 正在讀取檔案並對應資料庫...', true);
  var fd = new FormData();
  fd.append('file', file);
  fetch('/api/customs/upload', {method:'POST', body:fd})
    .then(function(r){return r.json();})
    .then(function(d){
      if(!d.ok) { showMsg('&#128561; ' + d.msg, false); return; }
      if(d.db_err) {
        showMsg('&#9888; 無法連接商品報關資料庫：' + d.db_err + '，所有商品將顯示為查無資料', false);
      } else {
        showMsg('&#10003; 資料庫載入成功（共 ' + d.db_count + ' 筆商品）', true);
      }
      allRows = d.rows;
      customsDb = d.db;
      checkMissing();
    });
}

function checkMissing() {
  missingSkus = [];
  allRows.forEach(function(row, i){
    if(row.status === 'missing') missingSkus.push(i);
  });
  if(missingSkus.length > 0) {
    currentMissingIdx = 0;
    showMsg('&#128561; 發現 ' + missingSkus.length + ' 個新商品需要建檔，請逐一填入資料！', false);
    showNewItemModal(missingSkus[0]);
  } else {
    renderPreview();
  }
}

function showNewItemModal(rowIdx) {
  var row = allRows[rowIdx];
  document.getElementById('modal-sku').textContent = row.sku;
  document.getElementById('f-sku').value = row.sku;
  document.getElementById('f-name').value = '';
  document.getElementById('f-style').value = '';
  document.getElementById('f-size').value = '';
  document.getElementById('f-material').value = '';
  document.getElementById('f-customs-name').value = '';
  document.getElementById('f-price').value = '';
  document.getElementById('f-unit').value = '';
  document.getElementById('f-product-size').value = '';
  document.getElementById('f-image').value = '';
  document.getElementById('new-item-modal').classList.add('show');
}

function saveNewItem() {
  var sku = document.getElementById('f-sku').value.trim();
  var material = document.getElementById('f-material').value.trim();
  var customsName = document.getElementById('f-customs-name').value.trim();
  var price = document.getElementById('f-price').value.trim();
  if(!material || !customsName || !price) {
    showMsg('&#9888; 材質、報關品名、單價為必填！', false);
    return;
  }
  var newItem = {
    sku: sku,
    name: document.getElementById('f-name').value.trim(),
    style: document.getElementById('f-style').value.trim(),
    size: document.getElementById('f-size').value.trim(),
    material: material,
    customs_name: customsName,
    price: parseFloat(price),
    unit: document.getElementById('f-unit').value.trim(),
    product_size: document.getElementById('f-product-size').value.trim(),
    image: document.getElementById('f-image').value.trim(),
  };
  fetch('/api/customs/new-item', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify(newItem)
  }).then(function(r){return r.json();}).then(function(d){
    if(!d.ok) { showMsg('&#128561; 儲存失敗：' + d.msg, false); return; }
    // 更新本地資料
    customsDb[sku] = newItem;
    var rowIdx = missingSkus[currentMissingIdx];
    allRows[rowIdx].status = 'ok';
    allRows[rowIdx].material = newItem.material;
    allRows[rowIdx].customs_name = newItem.customs_name;
    allRows[rowIdx].price = newItem.price;
    allRows[rowIdx].unit = newItem.unit;
    allRows[rowIdx].image = newItem.image;
    allRows[rowIdx].total_rmb = (parseFloat(allRows[rowIdx].total_pcs) * newItem.price).toFixed(2);
    document.getElementById('new-item-modal').classList.remove('show');
    currentMissingIdx++;
    if(currentMissingIdx < missingSkus.length) {
      showNewItemModal(missingSkus[currentMissingIdx]);
    } else {
      showMsg('&#127881; 所有新品建檔完成！', true);
      renderPreview();
    }
  });
}

function skipNewItem() {
  document.getElementById('new-item-modal').classList.remove('show');
  currentMissingIdx++;
  if(currentMissingIdx < missingSkus.length) {
    showNewItemModal(missingSkus[currentMissingIdx]);
  } else {
    renderPreview();
  }
}

function renderPreview() {
  var total = allRows.length;
  var ok = allRows.filter(function(r){return r.status==='ok';}).length;
  var miss = total - ok;
  document.getElementById('stat-total').textContent = '共 ' + total + ' 筆';
  document.getElementById('stat-ok').textContent = '&#10003; 對應成功 ' + ok + ' 筆';
  document.getElementById('stat-miss').textContent = miss > 0 ? '&#9888; 未建檔 ' + miss + ' 筆' : '';
  var tbody = document.getElementById('preview-body');
  tbody.innerHTML = allRows.map(function(r){
    var cls = r.status === 'ok' ? 'ok' : 'missing';
    var statusTag = r.status === 'ok'
      ? '<span class="tag-ok">&#10003; 已對應</span>'
      : '<span class="tag-miss">&#128560; 查無資料</span>';
    var img = r.image ? '<img class="thumb" src="' + r.image + '">' : '—';
    return '<tr class="' + cls + '">' +
      '<td>' + (r.sku||'—') + '</td>' +
      '<td>' + (r.type||'') + '</td>' +
      '<td>' + (r.product_size_orig||'') + '</td>' +
      '<td>' + (r.material||'') + '</td>' +
      '<td>' + (r.customs_name||'') + '</td>' +
      '<td>' + (r.box_no||'') + '</td>' +
      '<td>' + (r.pcs_per||'') + '</td>' +
      '<td>' + (r.qty||'') + '</td>' +
      '<td>' + (r.total_pcs||'') + '</td>' +
      '<td>' + (r.unit||'') + '</td>' +
      '<td>' + (r.price||'') + '</td>' +
      '<td>' + (r.total_rmb||'') + '</td>' +
      '<td>' + (r.gross_weight||'') + '</td>' +
      '<td>' + (r.len||'') + '</td>' +
      '<td>' + (r.wid||'') + '</td>' +
      '<td>' + (r.hei||'') + '</td>' +
      '<td>' + (r.volume||'') + '</td>' +
      '<td>' + (r.total_weight||'') + '</td>' +
      '<td>' + img + '</td>' +
      '<td>' + statusTag + '</td>' +
      '</tr>';
  }).join('');
  document.getElementById('preview-section').style.display = 'block';
  window.scrollTo(0, document.getElementById('preview-section').offsetTop);
}

function exportExcel() {
  var cabinetNo = document.getElementById('cabinet-no').value.trim();
  var sealNo    = document.getElementById('seal-no').value.trim();
  var shipDate  = document.getElementById('ship-date').value;
  var exporter  = document.getElementById('exporter').value.trim();
  var importer  = document.getElementById('importer').value.trim();
  // 日期格式轉換 yyyy-mm-dd → yyyy年mm月dd日
  var shipDateStr = '';
  if(shipDate) {
    var parts = shipDate.split('-');
    shipDateStr = parts[0] + '年' + parseInt(parts[1]) + '月' + parseInt(parts[2]) + '日';
  }
  showMsg('&#8987; 正在產生 Excel...', true);
  fetch('/api/customs/export', {
    method:'POST', headers:{'Content-Type':'application/json'},
    body: JSON.stringify({
      rows: allRows,
      cabinet_no: cabinetNo,
      seal_no:    sealNo,
      ship_date:  shipDateStr,
      exporter:   exporter,
      importer:   importer
    })
  }).then(function(r){return r.blob();}).then(function(blob){
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    var fname = '報關清單' + (cabinetNo ? '_' + cabinetNo : '') + '_' + new Date().toISOString().slice(0,10) + '.xlsx';
    a.download = fname;
    a.click();
    showMsg('&#127881; 匯出成功！', true);
  });
}

function resetPage() {
  allRows = []; missingSkus = []; currentMissingIdx = 0;
  document.getElementById('preview-section').style.display = 'none';
  document.getElementById('xlsx-in').value = '';
  showMsg('', true);
}

function showMsg(msg, ok) {
  var area = document.getElementById('msg-area');
  if(!msg) { area.innerHTML = ''; return; }
  area.innerHTML = '<div class="msg ' + (ok?'msg-ok':'msg-err') + '">' + msg + '</div>';
}
</script>
</body></html>"""

@app.route("/customs")
@login_required
def customs_page():
    return render_template_string(CUSTOMS_HTML)

@app.route("/api/customs/upload", methods=["POST"])
@login_required
def api_customs_upload():
    try:
        import openpyxl
    except ImportError:
        return jsonify({"ok": False, "msg": "請安裝 openpyxl"})

    f = request.files.get("file")
    if not f:
        return jsonify({"ok": False, "msg": "未收到檔案"})

    try:
        file_bytes = f.read()
        fname = f.filename.lower()

        if fname.endswith('.xls'):
            # 舊版 xls → 用 xlrd 讀取
            try:
                import xlrd
                wb_xls = xlrd.open_workbook(file_contents=file_bytes)
                ws_xls = wb_xls.sheet_by_index(0)
                # 轉成統一格式處理
                def cv_xls(r, c):
                    try:
                        v = ws_xls.cell_value(r, c)
                        return str(v).strip() if v != '' else ""
                    except:
                        return ""
                max_row = ws_xls.nrows
                use_xls = True
            except ImportError:
                return jsonify({"ok": False, "msg": "&#128561; 哎呀！.xls 格式需要安裝 xlrd，請另存為 .xlsx 再上傳！"})
        elif fname.endswith('.xlsx'):
            import openpyxl
            # keep_vba=False, 不載入圖片避免記憶體爆炸
            wb = openpyxl.load_workbook(
                io.BytesIO(file_bytes),
                data_only=True,
                keep_vba=False
            )
            ws = wb.active
            use_xls = False
            max_row = ws.max_row
        else:
            return jsonify({"ok": False, "msg": "&#128561; 哎呀！只支援 .xlsx 或 .xls 格式，請確認檔案類型！"})

        # 統一的讀取函數
        def cv(r, c):
            if use_xls:
                try:
                    v = ws_xls.cell_value(r-1, c-1)  # xls 是 0-indexed
                    return str(v).strip() if v != '' else ""
                except:
                    return ""
            else:
                v = ws.cell(r, c).value
                return str(v).strip() if v is not None else ""

        def has_data_in_row(r):
            return any(cv(r, c) for c in range(1, 5))

        # 找標題列並建立欄位對應
        HEADER_KEYWORDS = ["箱號", "箱号", "PCS", "品名", "材質", "材质"]
        HEADER_MAP = {
            "類型": "type", "类型": "type", "嘜頭": "type",
            "產品尺寸": "product_size_orig", "产品尺寸": "product_size_orig",
            "材質": "material", "材质": "material",
            "品名": "customs_name_raw",
            "箱號": "box_no", "箱号": "box_no",
            "PCS/件": "pcs_per", "PCS": "pcs_per",
            "件數": "qty", "件数": "qty",
            "總PCS": "total_pcs", "总PCS": "total_pcs",
            "單位": "unit", "单位": "unit",
            "單價": "price", "单价": "price",
            "總金額RMB": "total_rmb", "总金额RMB": "total_rmb",
            "毛重": "gross_weight",
            "長": "len", "长": "len",
            "寬": "wid", "宽": "wid",
            "高": "hei",
            "材積": "volume", "材积": "volume",
            "總重量": "total_weight", "总重量": "total_weight",
            "SKU": "sku", "SKU編碼": "sku",
        }

        header_row = None
        col_map = {}  # field_name -> col_index (1-based)
        data_start = 2

        for r in range(1, min(15, max_row+1)):
            row_vals = [cv(r, c) for c in range(1, min(25, (ws_xls.ncols if use_xls else ws.max_column)+1))]
            row_text = " ".join(row_vals)
            if any(kw in row_text for kw in HEADER_KEYWORDS):
                header_row = r
                data_start = r + 1
                # 建立欄位對應
                for c, val in enumerate(row_vals, 1):
                    val = val.strip()
                    if val in HEADER_MAP:
                        col_map[HEADER_MAP[val]] = c
                log(f"找到標題列第{r}列，欄位對應：{col_map}")
                break

        def get_col(row_r, field, default_col=None):
            """用欄位名稱取值，找不到時用預設欄位號"""
            c = col_map.get(field, default_col)
            if c:
                return cv(row_r, c)
            return ""

        # SKU 欄：優先用偵測到的，否則預設第18欄
        sku_col = col_map.get("sku", 18)

        # 載入資料庫
        db, db_err = load_customs_db()
        if db_err:
            log(f"載入報關資料庫失敗：{db_err}")

        log(f"報關資料庫載入：{len(db)} 筆，SKU清單：{list(db.keys())[:10]}")

        rows = []
        errors = []
        for r in range(data_start, max_row+1):
            sku = cv(r, sku_col)
            if not sku:
                if not has_data_in_row(r):
                    continue
                errors.append(f"第 {r} 列缺少 SKU（第{sku_col}欄）")

            total_pcs = get_col(r, "total_pcs", 8)
            price_db  = str(db.get(sku, {}).get("price", "")) if sku in db else ""
            try:
                total_rmb = round(float(total_pcs) * float(price_db), 2) if total_pcs and price_db else ""
            except:
                total_rmb = ""

            row = {
                "sku":               sku,
                "type":              get_col(r, "type", 1),
                "product_size_orig": get_col(r, "product_size_orig", 2),
                "material":          db.get(sku, {}).get("material", get_col(r, "material", 3)) if sku in db else get_col(r, "material", 3),
                "customs_name":      db.get(sku, {}).get("customs_name", "") if sku in db else "",
                "box_no":            get_col(r, "box_no", 5),
                "pcs_per":           get_col(r, "pcs_per", 6),
                "qty":               get_col(r, "qty", 7),
                "total_pcs":         total_pcs,
                "unit":              db.get(sku, {}).get("unit", get_col(r, "unit", 9)) if sku in db else get_col(r, "unit", 9),
                "price":             price_db if sku in db else get_col(r, "price", 10),
                "total_rmb":         str(total_rmb) if total_rmb != "" else get_col(r, "total_rmb", 11),
                "gross_weight":      get_col(r, "gross_weight", 12),
                "len":               get_col(r, "len", 13),
                "wid":               get_col(r, "wid", 14),
                "hei":               get_col(r, "hei", 15),
                "volume":            get_col(r, "volume", 16),
                "total_weight":      get_col(r, "total_weight", 17),
                "image":             db.get(sku, {}).get("image", "") if sku in db else "",
                "status":            "ok" if sku in db else ("missing" if sku else "no_sku"),
            }
            rows.append(row)

        if errors:
            log(f"欄位警告：{errors}")

        return jsonify({"ok": True, "rows": rows, "db": db, "warnings": errors, "db_count": len(db), "db_err": db_err})
    except Exception as e:
        return jsonify({"ok": False, "msg": f"讀取失敗：{e}"})

@app.route("/api/customs/new-item", methods=["POST"])
@login_required
def api_customs_new_item():
    data = request.get_json()
    err = append_to_customs_db(data)
    if err:
        return jsonify({"ok": False, "msg": err})
    log(f"新品建檔：{data.get('sku')} {data.get('customs_name')}")
    return jsonify({"ok": True})

@app.route("/api/customs/export", methods=["POST"])
@login_required
def api_customs_export():
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment
        data = request.get_json()
        rows       = data.get("rows", [])
        cabinet_no = data.get("cabinet_no", "")
        seal_no    = data.get("seal_no", "")
        ship_date  = data.get("ship_date", "")
        exporter   = data.get("exporter", "")
        importer   = data.get("importer", "")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "本次報關清單"

        # ── A1：出口商（可填寫）──
        ws.cell(1, 1, exporter)
        ws.cell(1, 1).font = Font(bold=True, size=11)

        # ── A2：進口商（可填寫）──
        ws.cell(2, 1, importer)
        ws.cell(2, 1).font = Font(size=11)

        # ── A3：固定 CFR ──
        ws.cell(3, 1, "CFR")
        ws.cell(3, 1).font = Font(bold=True)

        ws.cell(1, 14, "義烏:")
        ws.cell(1, 14).font = Font(bold=True)
        ws.cell(1, 15, "地址:浙江省義烏市江東街道青口東洲路1017號（東山路與東洲路交叉口）")
        ws.cell(2, 15, "電話:0579-85202818     傳真:0579-85202819")

        # ── 第4列空白 ──
        ws.cell(4, 1, "")

        # ── 第5列：櫃號/封號/出貨日期 ──────────────────────
        cabinet_text = f"櫃號: {cabinet_no}      封號：{seal_no}"
        ws.cell(5, 1, cabinet_text)
        ws.cell(5, 1).font = Font(bold=True)
        if ship_date:
            ws.cell(5, 16, f"出貨日期:{ship_date}")
            ws.cell(5, 16).font = Font(bold=True)

        # ── 第6列：欄位標題 ──────────────────────────────────
        HEADER_ROW = 6
        headers = ["類型","產品尺寸","材質","品名","箱號","PCS/件","件數","總PCS",
                   "單位","單價","總金額RMB","毛重","長","寬","高","材積","總重量","圖片"]
        header_fill = PatternFill("solid", fgColor="FF1565C0")
        for c, h in enumerate(headers, 1):
            cell = ws.cell(HEADER_ROW, c, h)
            cell.fill = header_fill
            cell.font = Font(bold=True, color="FFFFFFFF")
            cell.alignment = Alignment(horizontal="center")

        # 數字格式
        fmt_int  = '0'     # F/G/H 欄：整數無小數
        fmt_2dec = '0.00'  # P/Q 欄：兩位小數

        def to_num(val):
            try:
                return float(str(val).replace(',', '')) if val not in ('', None) else ''
            except:
                return val

        # ── 第7列起：商品資料 ──────────────────────────────
        miss_fill = PatternFill("solid", fgColor="FFFFEBEE")
        for r, row in enumerate(rows, HEADER_ROW + 1):
            ws.cell(r, 1,  row.get("type",""))
            ws.cell(r, 2,  row.get("product_size_orig",""))
            ws.cell(r, 3,  row.get("material",""))
            ws.cell(r, 4,  row.get("customs_name",""))
            ws.cell(r, 5,  row.get("box_no",""))
            # F(6) G(7) H(8)：整數無小數點
            for col, key in [(6,"pcs_per"),(7,"qty"),(8,"total_pcs")]:
                v = to_num(row.get(key,""))
                c = ws.cell(r, col, int(v) if isinstance(v, float) else v)
                if isinstance(v, float):
                    c.number_format = fmt_int
            ws.cell(r, 9,  row.get("unit",""))
            ws.cell(r, 10, to_num(row.get("price","")))
            ws.cell(r, 11, to_num(row.get("total_rmb","")))
            ws.cell(r, 12, to_num(row.get("gross_weight","")))
            ws.cell(r, 13, to_num(row.get("len","")))
            ws.cell(r, 14, to_num(row.get("wid","")))
            ws.cell(r, 15, to_num(row.get("hei","")))
            # P(16) Q(17)：兩位小數
            for col, key in [(16,"volume"),(17,"total_weight")]:
                v = to_num(row.get(key,""))
                c = ws.cell(r, col, v)
                if isinstance(v, float):
                    c.number_format = fmt_2dec
            ws.cell(r, 18, row.get("image",""))
            if row.get("status") != "ok":
                for c in range(1, 19):
                    ws.cell(r, c).fill = miss_fill

        # 自動調整欄寬
        for col in ws.columns:
            max_len = max((len(str(cell.value or "")) for cell in col), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

        fname = f"報關清單{'_'+cabinet_no if cabinet_no else ''}_{datetime.now().strftime('%Y%m%d')}.xlsx"
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        return send_file(buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=fname)
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


# ============================================================
# 貨架入庫工具
# ============================================================

WAREHOUSE_HTML = """<!DOCTYPE html>
<html lang="zh-TW"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>貨架入庫系統</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Microsoft JhengHei",sans-serif;background:#0f1923;color:#fff;min-height:100vh}
.topbar{background:#0a1219;height:52px;padding:0 20px;display:flex;align-items:center;gap:12px;border-bottom:1px solid #1e2d3d}
.logo{font-size:15px;font-weight:600;margin-right:auto;color:#fff}.logo span{color:#f4a100}
.topbar a{color:#aaa;font-size:12px;text-decoration:none}
.topbar a:hover{color:#f4a100}
.tabs{display:flex;gap:0;border-bottom:2px solid #1e2d3d;margin:0 20px}
.tab-btn{padding:12px 28px;background:none;border:none;color:#888;font-size:14px;cursor:pointer;font-family:inherit;border-bottom:2px solid transparent;margin-bottom:-2px;transition:.2s}
.tab-btn.active{color:#f4a100;border-bottom-color:#f4a100;font-weight:600}
.tab-content{display:none;padding:20px}
.tab-content.active{display:block}
.scan-card{background:#1a2535;border:1px solid #1e2d3d;border-radius:12px;padding:24px;margin-bottom:16px}
.scan-card h2{font-size:14px;color:#f4a100;margin-bottom:16px;font-weight:600;letter-spacing:1px}
.scan-input-wrap{position:relative}
.scan-input{width:100%;padding:14px 48px 14px 16px;background:#0f1923;border:2px solid #1e2d3d;border-radius:8px;color:#fff;font-size:16px;font-family:inherit;transition:.2s}
.scan-input:focus{outline:none;border-color:#f4a100}
.scan-input::placeholder{color:#444}
.scan-icon{position:absolute;right:14px;top:50%;transform:translateY(-50%);font-size:20px;pointer-events:none}
.scanned-list{display:flex;flex-direction:column;gap:6px;margin-top:12px;max-height:280px;overflow-y:auto}
.scanned-item{display:flex;align-items:center;justify-content:space-between;background:#0f1923;border:1px solid #1e2d3d;border-radius:6px;padding:8px 12px}
.scanned-item .sku{font-size:14px;font-weight:600;color:#fff}
.scanned-item .qty-ctrl{display:flex;align-items:center;gap:8px}
.qty-btn{width:28px;height:28px;border-radius:50%;border:1px solid #f4a100;background:none;color:#f4a100;font-size:18px;cursor:pointer;display:flex;align-items:center;justify-content:center}
.qty-btn:hover{background:#f4a100;color:#000}
.qty-num{font-size:16px;font-weight:700;color:#f4a100;min-width:24px;text-align:center}
.del-btn{background:none;border:none;color:#555;font-size:18px;cursor:pointer;padding:0 4px}
.del-btn:hover{color:#f44}
.empty-hint{color:#444;font-size:13px;text-align:center;padding:20px}
.confirm-overlay{display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.85);z-index:999;align-items:center;justify-content:center;flex-direction:column}
.confirm-overlay.show{display:flex}
.confirm-box{background:#1a2535;border:3px solid #f4a100;border-radius:20px;padding:40px 60px;text-align:center;max-width:500px;width:90%}
.confirm-rack{font-size:72px;font-weight:900;color:#f4a100;letter-spacing:4px;line-height:1;margin:16px 0;word-break:break-all}
.confirm-label{font-size:16px;color:#aaa;margin-bottom:8px}
.confirm-items{font-size:13px;color:#ccc;margin:16px 0;text-align:left;background:#0f1923;border-radius:8px;padding:12px;max-height:200px;overflow-y:auto}
.confirm-items .ci{padding:4px 0;border-bottom:1px solid #1e2d3d;display:flex;justify-content:space-between}
.confirm-items .ci:last-child{border:none}
.confirm-btns{display:flex;gap:12px;margin-top:24px;justify-content:center}
.btn{padding:12px 28px;border-radius:8px;border:none;font-size:14px;cursor:pointer;font-weight:600;font-family:inherit;transition:.2s}
.btn-confirm{background:#2e7d32;color:#fff}.btn-confirm:hover{background:#1b5e20}
.btn-cancel{background:#555;color:#fff}.btn-cancel:hover{background:#333}
.btn-yellow{background:#f4a100;color:#000}.btn-yellow:hover{background:#e69500}
.result-card{background:#1a2535;border:1px solid #1e2d3d;border-radius:10px;padding:16px;margin-bottom:10px}
.result-sku{font-size:16px;font-weight:700;color:#f4a100;margin-bottom:8px}
.result-loc{display:flex;flex-wrap:wrap;gap:8px}
.loc-tag{background:#0f1923;border:1px solid #f4a100;border-radius:20px;padding:6px 14px;font-size:13px;color:#fff}
.loc-tag .qty{color:#f4a100;font-weight:700}
.loc-tag .time{font-size:11px;color:#666;display:block;margin-top:2px}
.no-result{color:#555;text-align:center;padding:40px;font-size:14px}
.rec-table{width:100%;border-collapse:collapse;font-size:13px}
.rec-table th{background:#0f1923;padding:8px 10px;text-align:left;color:#888;font-weight:500;border-bottom:1px solid #1e2d3d}
.rec-table td{padding:8px 10px;border-bottom:1px solid #1a2535;vertical-align:middle}
.rec-table tr:hover td{background:#1a2535}
.badge-rack{background:#1e3a5f;color:#64b5f6;padding:2px 10px;border-radius:10px;font-size:12px;font-weight:600}
.msg-bar{padding:10px 16px;border-radius:6px;font-size:13px;margin-bottom:12px;display:none}
.msg-ok{background:#1b5e20;color:#a5d6a7;display:block}
.msg-err{background:#4a0000;color:#ef9a9a;display:block}
.search-wrap{display:flex;gap:8px;margin-bottom:16px}
.search-input{flex:1;padding:10px 14px;background:#0f1923;border:2px solid #1e2d3d;border-radius:8px;color:#fff;font-size:14px;font-family:inherit}
.search-input:focus{outline:none;border-color:#f4a100}
.btn-sm{padding:10px 20px;border-radius:8px;border:none;font-size:13px;cursor:pointer;font-weight:600;font-family:inherit}
.spinner{display:inline-block;width:16px;height:16px;border:2px solid #333;border-top-color:#f4a100;border-radius:50%;animation:spin .7s linear infinite;vertical-align:middle;margin-right:6px}
@keyframes spin{to{transform:rotate(360deg)}}
</style></head><body>
<div class="topbar">
  <div class="logo">&#x1F4E6; <span>貨架入庫系統</span></div>
  <a href="/">&#x2302; 返回首頁</a>
  <a href="/logout" style="margin-left:8px">登出</a>
</div>
<div style="margin:16px 20px 0">
  <div class="tabs">
    <button class="tab-btn active" onclick="switchTab('inbound',this)">&#x1F4E5; 入庫作業</button>
    <button class="tab-btn" onclick="switchTab('search',this)">&#x1F50D; 查找儲位</button>
    <button class="tab-btn" onclick="switchTab('records',this)">&#x1F4CB; 入庫紀錄</button>
  </div>
</div>

<!-- 入庫作業 -->
<div id="tab-inbound" class="tab-content active">
  <div id="msg-inbound" class="msg-bar"></div>
  <div class="scan-card">
    <h2>&#x25CF; STEP 1 &nbsp;掃描 / 輸入貨號</h2>
    <p style="font-size:12px;color:#666;margin-bottom:12px">掃描商品條碼或手動輸入貨號，同一貨號掃兩次 = 2件</p>
    <div class="scan-input-wrap">
      <input type="text" id="sku-input" class="scan-input" placeholder="掃描條碼或輸入貨號..." onkeydown="if(event.key==='Enter')addSku()">
      <span class="scan-icon">&#x1F4F7;</span>
    </div>
    <div style="display:flex;gap:8px;margin-top:10px">
      <button class="btn btn-yellow" onclick="addSku()" style="flex:1">+ 加入清單</button>
      <button class="btn btn-cancel" onclick="clearSkus()">清空</button>
    </div>
    <div id="scanned-list" class="scanned-list"><div class="empty-hint">尚未掃描任何貨號</div></div>
  </div>
  <div class="scan-card">
    <h2>&#x25CF; STEP 2 &nbsp;掃描 / 輸入儲位條碼</h2>
    <p style="font-size:12px;color:#666;margin-bottom:12px">格式：RACK-A-1，掃描後會出現大字確認畫面</p>
    <div class="scan-input-wrap">
      <input type="text" id="rack-input" class="scan-input" placeholder="掃描貨架條碼，例：RACK-A-1" onkeydown="if(event.key==='Enter')confirmRack()">
      <span class="scan-icon">&#x1F4CD;</span>
    </div>
    <button class="btn btn-yellow" onclick="confirmRack()" style="width:100%;margin-top:10px">&#x1F50D; 確認儲位</button>
  </div>
</div>

<!-- 查找儲位 -->
<div id="tab-search" class="tab-content">
  <div class="scan-card">
    <h2>&#x25CF; 查找商品在哪個儲位</h2>
    <div class="search-wrap">
      <input type="text" id="search-input" class="search-input" placeholder="輸入貨號或關鍵字..." onkeydown="if(event.key==='Enter')doSearch()">
      <button class="btn btn-sm btn-yellow" onclick="doSearch()">查找</button>
    </div>
    <div id="search-result"></div>
  </div>
  <div class="scan-card">
    <h2>&#x25CF; 查找貨架裡有什麼</h2>
    <div class="search-wrap">
      <input type="text" id="rack-search-input" class="search-input" placeholder="輸入儲位條碼，例：RACK-A-1" onkeydown="if(event.key==='Enter')doRackSearch()">
      <button class="btn btn-sm btn-yellow" onclick="doRackSearch()">查找</button>
    </div>
    <div id="rack-search-result"></div>
  </div>
</div>

<!-- 入庫紀錄 -->
<div id="tab-records" class="tab-content">
  <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
    <span style="font-size:13px;color:#888">最近 100 筆入庫紀錄</span>
    <button class="btn btn-sm btn-yellow" onclick="loadRecords()">&#x21BB; 重新整理</button>
  </div>
  <div style="overflow-x:auto">
    <table class="rec-table">
      <thead><tr><th>入庫時間</th><th>貨號</th><th>儲位</th><th>數量</th></tr></thead>
      <tbody id="records-body"><tr><td colspan="4" style="text-align:center;color:#555;padding:30px">點擊重新整理載入紀錄</td></tr></tbody>
    </table>
  </div>
</div>

<!-- 儲位確認大畫面 -->
<div class="confirm-overlay" id="confirm-overlay">
  <div class="confirm-box">
    <div class="confirm-label">&#x1F4CD; 確認入庫到此儲位</div>
    <div class="confirm-rack" id="confirm-rack-text">RACK-A-1</div>
    <div style="font-size:13px;color:#888;margin-bottom:8px">以下商品將入庫：</div>
    <div class="confirm-items" id="confirm-items-list"></div>
    <div class="confirm-btns">
      <button class="btn btn-cancel" onclick="hideConfirm()">&#x274C; 取消</button>
      <button class="btn btn-confirm" onclick="doInbound()">&#x2705; 確認入庫</button>
    </div>
  </div>
</div>

<script>
var skuList = [];

function switchTab(name, btn) {
  document.querySelectorAll('.tab-content').forEach(function(el){el.classList.remove('active');});
  document.querySelectorAll('.tab-btn').forEach(function(el){el.classList.remove('active');});
  document.getElementById('tab-'+name).classList.add('active');
  btn.classList.add('active');
  if(name==='records') loadRecords();
}

function addSku() {
  var val = document.getElementById('sku-input').value.trim().toUpperCase();
  if(!val) return;
  var existing = skuList.find(function(x){return x.sku===val;});
  if(existing){existing.qty+=1;}else{skuList.push({sku:val,qty:1});}
  document.getElementById('sku-input').value='';
  document.getElementById('sku-input').focus();
  renderSkuList();
}

function renderSkuList() {
  var el = document.getElementById('scanned-list');
  if(skuList.length===0){el.innerHTML='<div class="empty-hint">尚未掃描任何貨號</div>';return;}
  el.innerHTML = skuList.map(function(item,i){
    return '<div class="scanned-item">'+
      '<span class="sku">'+item.sku+'</span>'+
      '<div class="qty-ctrl">'+
        '<button class="qty-btn" onclick="changeQty('+i+',-1)">&#x2212;</button>'+
        '<span class="qty-num">'+item.qty+'</span>'+
        '<button class="qty-btn" onclick="changeQty('+i+',1)">+</button>'+
        '<button class="del-btn" onclick="removeSku('+i+')">&#x2715;</button>'+
      '</div></div>';
  }).join('');
}

function changeQty(i,delta){skuList[i].qty=Math.max(1,skuList[i].qty+delta);renderSkuList();}
function removeSku(i){skuList.splice(i,1);renderSkuList();}
function clearSkus(){skuList=[];renderSkuList();}

function confirmRack() {
  if(skuList.length===0){showMsg('inbound','&#x26A0; 請先掃描至少一個貨號！',false);return;}
  var rack=document.getElementById('rack-input').value.trim().toUpperCase();
  if(!rack){showMsg('inbound','&#x26A0; 請輸入或掃描儲位條碼！',false);return;}
  document.getElementById('confirm-rack-text').textContent=rack;
  document.getElementById('confirm-items-list').innerHTML=skuList.map(function(item){
    return '<div class="ci"><span>'+item.sku+'</span><span style="color:#f4a100">'+item.qty+' 件</span></div>';
  }).join('');
  document.getElementById('confirm-overlay').classList.add('show');
}

function hideConfirm(){document.getElementById('confirm-overlay').classList.remove('show');}

function doInbound() {
  var rack=document.getElementById('rack-input').value.trim().toUpperCase();
  hideConfirm();
  showMsg('inbound','<span class="spinner"></span>正在寫入紀錄...',true);
  fetch('/api/warehouse/inbound',{
    method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({rack:rack,items:skuList})
  }).then(function(r){return r.json();}).then(function(d){
    if(d.ok){
      showMsg('inbound','&#x2705; 入庫成功！'+d.count+' 筆紀錄已儲存',true);
      skuList=[];renderSkuList();
      document.getElementById('rack-input').value='';
    }else{showMsg('inbound','&#x274C; 入庫失敗：'+d.msg,false);}
  });
}

function doSearch() {
  var q=document.getElementById('search-input').value.trim().toUpperCase();
  if(!q) return;
  var el=document.getElementById('search-result');
  el.innerHTML='<div style="color:#888;padding:12px"><span class="spinner"></span>查找中...</div>';
  fetch('/api/warehouse/search?q='+encodeURIComponent(q))
    .then(function(r){return r.json();}).then(function(d){
      if(!d.ok){el.innerHTML='<div class="no-result">查詢失敗：'+d.msg+'</div>';return;}
      if(d.results.length===0){el.innerHTML='<div class="no-result">&#x1F50D; 找不到「'+q+'」的入庫紀錄</div>';return;}
      var grouped={};
      d.results.forEach(function(r){if(!grouped[r.sku])grouped[r.sku]=[];grouped[r.sku].push(r);});
      el.innerHTML=Object.keys(grouped).map(function(sku){
        var locs=grouped[sku].map(function(r){
          return '<div class="loc-tag"><span class="badge-rack">'+r.rack+'</span>'+
            (r.qty?' <span class="qty">x'+r.qty+'</span>':'')+
            '<span class="time">'+r.time+'</span></div>';
        }).join('');
        return '<div class="result-card"><div class="result-sku">&#x1F4E6; '+sku+'</div>'+
          '<div class="result-loc">'+locs+'</div></div>';
      }).join('');
    });
}

function doRackSearch() {
  var q=document.getElementById('rack-search-input').value.trim().toUpperCase();
  if(!q) return;
  var el=document.getElementById('rack-search-result');
  el.innerHTML='<div style="color:#888;padding:12px"><span class="spinner"></span>查找中...</div>';
  fetch('/api/warehouse/search-rack?rack='+encodeURIComponent(q))
    .then(function(r){return r.json();}).then(function(d){
      if(!d.ok){el.innerHTML='<div class="no-result">查詢失敗：'+d.msg+'</div>';return;}
      if(d.results.length===0){el.innerHTML='<div class="no-result">&#x1F50D; 儲位「'+q+'」目前沒有入庫紀錄</div>';return;}
      el.innerHTML='<div class="result-card">'+
        '<div class="result-sku">&#x1F4CD; 儲位 '+q+'（共 '+d.results.length+' 筆）</div>'+
        '<table class="rec-table" style="margin-top:8px">'+
        '<thead><tr><th>貨號</th><th>數量</th><th>入庫時間</th></tr></thead><tbody>'+
        d.results.map(function(r){
          return '<tr><td style="font-weight:600;color:#f4a100">'+r.sku+'</td>'+
            '<td>'+(r.qty||'—')+'</td><td style="color:#888;font-size:12px">'+r.time+'</td></tr>';
        }).join('')+'</tbody></table></div>';
    });
}

function loadRecords() {
  var tbody=document.getElementById('records-body');
  tbody.innerHTML='<tr><td colspan="4" style="text-align:center;color:#888;padding:20px"><span class="spinner"></span>載入中...</td></tr>';
  fetch('/api/warehouse/records').then(function(r){return r.json();}).then(function(d){
    if(!d.ok||d.records.length===0){
      tbody.innerHTML='<tr><td colspan="4" style="text-align:center;color:#555;padding:30px">尚無入庫紀錄</td></tr>';return;
    }
    tbody.innerHTML=d.records.map(function(r){
      return '<tr><td style="color:#888;font-size:12px">'+r.time+'</td>'+
        '<td style="font-weight:600;color:#f4a100">'+r.sku+'</td>'+
        '<td><span class="badge-rack">'+r.rack+'</span></td>'+
        '<td>'+(r.qty||'—')+'</td></tr>';
    }).join('');
  });
}

function showMsg(zone,msg,ok) {
  var el=document.getElementById('msg-'+zone);
  el.className='msg-bar '+(ok?'msg-ok':'msg-err');
  el.innerHTML=msg;
  if(ok) setTimeout(function(){el.className='msg-bar';},4000);
}

document.getElementById('sku-input').focus();
</script>
</body></html>"""


def get_warehouse_sheet():
    """取得貨架庫位紀錄 worksheet，不存在就自動建立"""
    client, err = get_sheets_client()
    if err:
        return None, err
    try:
        sheet_id = os.environ.get("GOOGLE_SHEETS_ID")
        sh = client.open_by_key(sheet_id)
        try:
            ws = sh.worksheet("貨架庫位紀錄")
        except Exception:
            ws = sh.add_worksheet(title="貨架庫位紀錄", rows=1000, cols=6)
            ws.append_row(["貨號", "儲位", "數量", "入庫時間", "備註"])
        return ws, None
    except Exception as e:
        return None, str(e)


@app.route("/warehouse")
@login_required
def warehouse_page():
    return render_template_string(WAREHOUSE_HTML)


@app.route("/api/warehouse/inbound", methods=["POST"])
@login_required
def api_warehouse_inbound():
    data  = request.get_json()
    rack  = data.get("rack", "").strip().upper()
    items = data.get("items", [])
    if not rack:
        return jsonify({"ok": False, "msg": "儲位條碼不能為空"})
    if not items:
        return jsonify({"ok": False, "msg": "請先掃描至少一個貨號"})
    ws, err = get_warehouse_sheet()
    if err:
        return jsonify({"ok": False, "msg": err})
    try:
        now = datetime.now().strftime("%Y/%m/%d %H:%M")
        rows_to_add = []
        for item in items:
            sku = str(item.get("sku", "")).strip().upper()
            qty = item.get("qty", 1)
            if sku:
                rows_to_add.append([sku, rack, qty, now, ""])
        if rows_to_add:
            ws.append_rows(rows_to_add)
        return jsonify({"ok": True, "count": len(rows_to_add)})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)})


@app.route("/api/warehouse/search")
@login_required
def api_warehouse_search():
    q = request.args.get("q", "").strip().upper()
    if not q:
        return jsonify({"ok": False, "msg": "請輸入查詢關鍵字"})
    ws, err = get_warehouse_sheet()
    if err:
        return jsonify({"ok": False, "msg": err})
    try:
        rows = ws.get_all_records()
        results = []
        for row in rows:
            sku = str(row.get("貨號", "")).strip().upper()
            if q in sku:
                results.append({
                    "sku":  sku,
                    "rack": str(row.get("儲位", "")),
                    "qty":  str(row.get("數量", "")),
                    "time": str(row.get("入庫時間", "")),
                })
        results.reverse()
        return jsonify({"ok": True, "results": results})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)})


@app.route("/api/warehouse/search-rack")
@login_required
def api_warehouse_search_rack():
    rack = request.args.get("rack", "").strip().upper()
    if not rack:
        return jsonify({"ok": False, "msg": "請輸入儲位條碼"})
    ws, err = get_warehouse_sheet()
    if err:
        return jsonify({"ok": False, "msg": err})
    try:
        rows = ws.get_all_records()
        results = []
        for row in rows:
            r = str(row.get("儲位", "")).strip().upper()
            if rack in r:
                results.append({
                    "sku":  str(row.get("貨號", "")),
                    "rack": r,
                    "qty":  str(row.get("數量", "")),
                    "time": str(row.get("入庫時間", "")),
                })
        results.reverse()
        return jsonify({"ok": True, "results": results})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)})


@app.route("/api/warehouse/records")
@login_required
def api_warehouse_records():
    ws, err = get_warehouse_sheet()
    if err:
        return jsonify({"ok": False, "msg": err})
    try:
        rows = ws.get_all_records()
        records = [{"sku": str(r.get("貨號","")), "rack": str(r.get("儲位","")),
                    "qty": str(r.get("數量","")), "time": str(r.get("入庫時間",""))}
                   for r in rows]
        records.reverse()
        return jsonify({"ok": True, "records": records[:100]})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)})


if __name__ == "__main__":
    load_settings()
    threading.Thread(target=scheduler, daemon=True).start()
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
        print("  印單系統已啟動！")
        print("=" * 50)
        print(f"  本機使用：http://127.0.0.1:{port}")
        print(f"  區域網路：http://{local_ip}:{port}")
        print("=" * 50)
        print("  請勿關閉此視窗，關閉後系統停止運作")
        print("=" * 50)
        try:
            import webbrowser
            threading.Timer(2.0, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
        except Exception:
            pass
    else:
        print(f"[雲端模式] 系統啟動 port={port}")

    app.run(host="0.0.0.0", port=port, debug=False)
