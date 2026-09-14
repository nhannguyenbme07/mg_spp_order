"""
MGV Spare-Parts Order Suggestion — engine (no UI).

Logic:
  - HN  = tổng Quantity các kho có mã đuôi '23'
  - HCM = tổng Quantity các kho có mã đuôi '21'
  - Đuôi khác (10, 30, 31, ...) = hàng đi đường/trung chuyển -> KHÔNG tính vào tồn HN/HCM
  - Khớp mã 2 tầng: Part Number 2026 (ưu tiên) -> Former P/N (fallback)
  - Trigger đặt: Onhand(khu vực) <= Min(khu vực) -> đặt = Max - Onhand ; làm tròn LÊN ; giữ >0
  - Onhand HN & HCM luôn điền ở mọi dòng
  - Cost = Quantity * Unit Price (USD) từ tiêu chuẩn
  - Tra tên linh kiện từ PartList (SPP master) cho item nhập tay không có trong tiêu chuẩn
"""
import re
import math
import openpyxl

REGION_SUFFIX = {"23": "HN", "21": "HCM"}
BRAND_ORDER = {"STAGO": 0, "SEBIA": 1}
ITEM_GROUP_CHOICES = ["SPP", "CON", "REA", "TOOL"]
DEFAULT_REASON = "Replenish safety stock"
DEFAULT_PURPOSE = "For Service"

# 14 cột output (No + 13) — 'Cost' nằm ngay sau 'Quantity'; đã sửa typo 'Onhand HCM'
OUTPUT_COLUMNS = [
    "No", "Brand", "Item Code", "Item Name", "Item Group", "Quantity", "Cost", "PIC",
    "Analyzer", "Reason of Order", "Purpose", "Pending NW",
    "Onhand HN", "Onhand HCM", "Note",
]


# ----------------------------------------------------------------------------
# Chuẩn hoá
# ----------------------------------------------------------------------------
def norm_code(x):
    if x is None:
        return ""
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    s = str(x).strip().upper()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def display_code(x):
    if x is None:
        return ""
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x).strip()


def _clean_text(x):
    """Bỏ ký tự rác (_x000D_, \\r, \\n) và khoảng trắng thừa trong tên."""
    if x is None:
        return None
    s = str(x).replace("_x000D_", "").replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip() or None


def split_former(x):
    if x is None:
        return []
    return [norm_code(p) for p in re.split(r"[;/,]", str(x)) if norm_code(p)]


def _num(x):
    try:
        return float(x) if x is not None else 0.0
    except (TypeError, ValueError):
        return 0.0


def _as_int_if_whole(v):
    try:
        return int(v) if float(v).is_integer() else v
    except (TypeError, ValueError):
        return v


def compute_cost(qty, unit_price):
    """Cost = Quantity * Unit Price. None nếu thiếu đơn giá."""
    if unit_price in (None, ""):
        return None
    try:
        return round(_num(qty) * float(unit_price), 2)
    except (TypeError, ValueError):
        return None


# ----------------------------------------------------------------------------
# Đọc cột tiêu chuẩn theo TÊN header (bền khi thêm/bớt cột)
# ----------------------------------------------------------------------------
def _norm_header(h):
    return re.sub(r"\s+", " ", str(h).strip().upper().replace("\n", " ").replace("\r", " ")) if h is not None else ""


def _resolve_std_columns(ws):
    """Trả về dict {field: col_index} dò theo header ở dòng 1."""
    idx = {}
    for c in range(1, ws.max_column + 1):
        h = _norm_header(ws.cell(1, c).value)
        if not h:
            continue
        if "PART NUM" in h and "pn" not in idx:
            idx["pn"] = c
        elif h == "PART NAME":
            idx["name"] = c
        elif "FORMER" in h:
            idx["former"] = c
        elif "ANALYZER" in h:
            idx["analyzer"] = c
        elif h == "BRAND":
            idx["brand"] = c
        elif "UNIT PRICE" in h or "UNIT COST" in h or h == "PRICE":
            idx["unit_price"] = c
        elif h == "HAN MIN":
            idx["han_min"] = c
        elif h == "HAN MAX":
            idx["han_max"] = c
        elif h == "HCM MIN":
            idx["hcm_min"] = c
        elif h == "HCM MAX":
            idx["hcm_max"] = c
    return idx


def load_standard(path):
    """Đọc tiêu chuẩn -> (items, lookup). Đọc cột theo header nên chịu được thay đổi layout."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Standard Inventory v1.0"] if "Standard Inventory v1.0" in wb.sheetnames else wb.worksheets[0]
    col = _resolve_std_columns(ws)

    def cv(r, field):
        c = col.get(field)
        return ws.cell(r, c).value if c else None

    items, lookup = [], {}
    for r in range(2, ws.max_row + 1):
        pn = cv(r, "pn")
        if pn is None:
            continue
        item = {
            "pn": norm_code(pn),
            "pn_display": display_code(pn),
            "name": _clean_text(cv(r, "name")),
            "former": split_former(cv(r, "former")),
            "analyzer": cv(r, "analyzer"),
            "brand": cv(r, "brand"),
            "unit_price": (float(cv(r, "unit_price")) if cv(r, "unit_price") not in (None, "") else None),
            "han_min": _num(cv(r, "han_min")),
            "han_max": _num(cv(r, "han_max")),
            "hcm_min": _num(cv(r, "hcm_min")),
            "hcm_max": _num(cv(r, "hcm_max")),
        }
        idx = len(items)
        items.append(item)
        lookup.setdefault(item["pn"], idx)
    for idx, item in enumerate(items):
        for f in item["former"]:
            lookup.setdefault(f, idx)
    wb.close()
    return items, lookup


# ----------------------------------------------------------------------------
# PartList (SPP master) -> tra tên/brand/thiết bị từ mã. Header bị đảo nên tự dò cột.
# ----------------------------------------------------------------------------
def load_partlist(path):
    """Trả về dict { norm_code : {'name','brand','equipment'} }."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    # Dò 2 cột đầu: cột 'mã' = ít khoảng trắng hơn, cột 'tên' = nhiều khoảng trắng
    a_space = b_space = 0
    for r in range(2, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if isinstance(a, str) and " " in a:
            a_space += 1
        if isinstance(b, str) and " " in b:
            b_space += 1
    code_col, name_col = (1, 2) if a_space <= b_space else (2, 1)
    # Brand / equipment theo header (nếu có)
    hdr = {_norm_header(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1)}
    brand_col = hdr.get("BRAND")
    equip_col = next((c for h, c in hdr.items() if "EQUIP" in h or "COMPAT" in h), None)
    price_col = next((c for h, c in hdr.items()
                      if "UNIT PRICE" in h or "UNIT COST" in h or h == "PRICE"), None)

    out = {}
    for r in range(2, ws.max_row + 1):
        code = norm_code(ws.cell(r, code_col).value)
        if not code:
            continue
        price = None
        if price_col:
            pv = ws.cell(r, price_col).value
            price = float(pv) if pv not in (None, "") else None
        out.setdefault(code, {
            "name": _clean_text(ws.cell(r, name_col).value),
            "brand": (ws.cell(r, brand_col).value if brand_col else None),
            "equipment": (_clean_text(ws.cell(r, equip_col).value) if equip_col else None),
            "unit_price": price,
        })
    wb.close()
    return out


def price_for_codes(codes, partlist):
    """Tra đơn giá từ PartList theo danh sách mã (mã đầu tiên có giá sẽ được dùng)."""
    if not partlist:
        return None
    for c in codes:
        e = partlist.get(norm_code(c))
        if e and e.get("unit_price") not in (None, ""):
            return e["unit_price"]
    return None


def resolve_name(code, items, lookup, partlist=None):
    """Tra tên linh kiện: tiêu chuẩn trước, rồi PartList. None nếu không thấy."""
    code_n = norm_code(code)
    idx = lookup.get(code_n)
    if idx is not None and items[idx].get("name"):
        return items[idx]["name"]
    if partlist and code_n in partlist:
        return partlist[code_n].get("name")
    return None


# ----------------------------------------------------------------------------
# Nạp tồn kho — nhận diện được NHIỀU định dạng file
#   (1) PBI_Export: header dòng 1, cột 'Warehouse','Item Code','Item Name','Brand','Quantity'
#   (2) MGInventoryOnhand report: sheet 'Sheet1', header kiểu '...View[INVENTLOCATIONID]',
#       '[ITEMID]','[Item Name]','[RP_BRAND]','[CBQTY]', có 1 dòng trống dưới header
# ----------------------------------------------------------------------------
INV_ALIASES = {
    "wh":    {"WAREHOUSE", "INVENTLOCATIONID", "INVENT LOCATION ID", "LOCATION", "LOCATIONID"},
    "code":  {"ITEM CODE", "ITEMCODE", "ITEMID", "ITEM ID"},
    "name":  {"ITEM NAME", "ITEMNAME"},
    "brand": {"BRAND", "RP_BRAND"},
    "qty":   {"QUANTITY", "QTY", "CBQTY"},
}


def _extract_header(h):
    """Chuẩn hoá header; nếu dạng '...View[TOKEN]' thì lấy phần trong [] cuối."""
    if h is None:
        return ""
    s = str(h)
    m = re.findall(r"\[([^\]]+)\]", s)
    if m:
        s = m[-1]
    return re.sub(r"\s+", " ", s.strip().upper())


def _match_inv_columns(ws, header_row):
    col = {}
    for c in range(1, ws.max_column + 1):
        h = _extract_header(ws.cell(header_row, c).value)
        if not h:
            continue
        for field, names in INV_ALIASES.items():
            if field in col:
                continue
            if h in names or (field == "brand" and h.endswith("BRAND")):
                col[field] = c
    return col


def _find_inv_header(ws):
    """Tìm dòng header trong 6 dòng đầu (cần tối thiểu Warehouse + Item Code + Quantity)."""
    for r in range(1, min(6, ws.max_row) + 1):
        col = _match_inv_columns(ws, r)
        if {"wh", "code", "qty"} <= set(col):
            return r, col
    return None, None


def load_inventory(path, lookup, n_items):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = header_row = col = None
    for sheet in wb.worksheets:
        hr, cmap = _find_inv_header(sheet)
        if hr:
            ws, header_row, col = sheet, hr, cmap
            break
    if ws is None:
        wb.close()
        raise ValueError("Không nhận diện được định dạng file tồn kho "
                         "(cần có cột Warehouse/Location, Item Code/ItemID, Quantity/CBQTY).")

    c_wh, c_code, c_qty = col["wh"], col["code"], col["qty"]
    c_name, c_brand = col.get("name"), col.get("brand")

    std_onhand = [{"HN": 0.0, "HCM": 0.0} for _ in range(n_items)]
    raw_onhand, raw_info, seen_suffix = {}, {}, {}
    for r in range(header_row + 1, ws.max_row + 1):
        code = ws.cell(r, c_code).value
        if code is None:                        # bỏ dòng trống (vd dòng 2 của report)
            continue
        wh = ws.cell(r, c_wh).value
        qty = _num(ws.cell(r, c_qty).value)
        code_n = norm_code(code)
        if code_n not in raw_info:
            raw_info[code_n] = {
                "name": _clean_text(ws.cell(r, c_name).value) if c_name else None,
                "brand": ws.cell(r, c_brand).value if c_brand else None,
            }
        suffix = str(wh)[-2:] if wh is not None else ""
        seen_suffix[suffix] = seen_suffix.get(suffix, 0) + 1
        region = REGION_SUFFIX.get(suffix)
        if region is None:                      # đuôi 10/30/31... = đang về/trung chuyển
            continue
        raw_onhand.setdefault(code_n, {"HN": 0.0, "HCM": 0.0})[region] += qty
        idx = lookup.get(code_n)
        if idx is not None:
            std_onhand[idx][region] += qty
    wb.close()
    return std_onhand, raw_onhand, raw_info, {
        "suffix_counts": seen_suffix, "sheet": ws.title, "header_row": header_row}


# ----------------------------------------------------------------------------
# Danh sách cần đặt (tự động)
# ----------------------------------------------------------------------------
def _base_row(item, onhand):
    return {
        "Brand": (item["brand"] or "").strip(),
        "Item Code": item["pn_display"],
        "Item Name": item["name"],
        "Item Group": "SPP",
        "Quantity": None,
        "Cost": None,
        "PIC": None,
        "Analyzer": item["analyzer"],
        "Reason of Order": DEFAULT_REASON,
        "Purpose": DEFAULT_PURPOSE,
        "Pending NW": None,
        "Onhand HN": _as_int_if_whole(onhand["HN"]),
        "Onhand HCM": _as_int_if_whole(onhand["HCM"]),
        "Note": None,
        "Unit Price": item.get("unit_price"),
        "Std Min": None,
        "Std Max": None,
        "_source": "auto",
    }


def compute_orders(items, std_onhand, region, partlist=None):
    rows = []
    for idx, item in enumerate(items):
        onhand = std_onhand[idx]
        if region == "HN":
            mn, mx, cur = item["han_min"], item["han_max"], onhand["HN"]
        else:
            mn, mx, cur = item["hcm_min"], item["hcm_max"], onhand["HCM"]
        if cur <= mn:
            order_raw = mx - cur
            if order_raw > 0:
                order = math.ceil(order_raw)
                row = _base_row(item, onhand)
                # Giá: ưu tiên PartList (theo mã 2026 + Former P/N), fallback về giá trong standard
                price = price_for_codes([item["pn"]] + item["former"], partlist)
                if price is None:
                    price = item.get("unit_price")
                row["Unit Price"] = price
                row["Quantity"] = int(order)
                row["Cost"] = compute_cost(order, price)
                row["Std Min"] = _as_int_if_whole(mn)
                row["Std Max"] = _as_int_if_whole(mx)
                if order != order_raw:
                    row["Note"] = f"Qty rounded up from {_as_int_if_whole(order_raw)} (fractional on-hand)"
                rows.append(row)
    return sort_rows(rows)


def sort_rows(rows):
    return sorted(
        rows,
        key=lambda x: (BRAND_ORDER.get(str(x["Brand"]).upper(), 99), str(x["Item Name"] or "").upper()),
    )


# ----------------------------------------------------------------------------
# Item nhập tay
# ----------------------------------------------------------------------------
def enrich_manual(code, qty, reason, purpose, item_group,
                  items, lookup, std_onhand, raw_onhand, raw_info,
                  region="HN", typed_name=None, partlist=None):
    """Enrich 1 item nhập tay: tồn HN/HCM, Min/Max, đơn giá/cost, và tra tên (typed > tiêu chuẩn > PartList)."""
    code_n = norm_code(code)
    idx = lookup.get(code_n)
    std_min = std_max = unit_price = None

    if idx is not None:                                  # có trong tiêu chuẩn
        item = items[idx]
        onhand = std_onhand[idx]
        resolved_name = item["name"]
        brand, analyzer, pn_disp = item["brand"], item["analyzer"], item["pn_display"]
        if region == "HN":
            std_min, std_max = _as_int_if_whole(item["han_min"]), _as_int_if_whole(item["han_max"])
        else:
            std_min, std_max = _as_int_if_whole(item["hcm_min"]), _as_int_if_whole(item["hcm_max"])
    else:                                                # ngoài tiêu chuẩn -> PartList / tồn kho thô
        onhand = raw_onhand.get(code_n, {"HN": 0.0, "HCM": 0.0})
        pl = (partlist or {}).get(code_n, {})
        info = raw_info.get(code_n, {})
        resolved_name = pl.get("name") or info.get("name")
        brand = (pl.get("brand") or info.get("brand") or "")
        brand = str(brand).title() if brand else ""
        analyzer = pl.get("equipment")
        pn_disp = display_code(code)

    typed = _clean_text(typed_name)
    final_name = typed if typed else resolved_name       # typed thắng; trống -> tra tự động

    # Giá: ưu tiên PartList (mã + Former P/N nếu có trong tiêu chuẩn); fallback giá standard
    codes = [code_n] + (items[idx]["former"] if idx is not None else [])
    unit_price = price_for_codes(codes, partlist)
    if unit_price is None and idx is not None:
        unit_price = items[idx].get("unit_price")

    return {
        "Brand": (brand or "").strip(),
        "Item Code": pn_disp,
        "Item Name": final_name,
        "Item Group": item_group or "SPP",
        "Quantity": _as_int_if_whole(_num(qty)),
        "Cost": compute_cost(qty, unit_price),
        "PIC": None,
        "Analyzer": analyzer,
        "Reason of Order": reason or "",
        "Purpose": purpose or "",
        "Pending NW": None,
        "Onhand HN": _as_int_if_whole(onhand["HN"]),
        "Onhand HCM": _as_int_if_whole(onhand["HCM"]),
        "Note": None,
        "Unit Price": unit_price,
        "Std Min": std_min,
        "Std Max": std_max,
        "_source": "manual",
    }
