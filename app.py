"""
MGV Spare-Parts Order Suggestion — Streamlit UI.

Chạy:  streamlit run app.py
"""
import os
import tempfile
import pandas as pd
import streamlit as st

import engine as E
import writer as W
import auth as A

st.set_page_config(page_title="MGV - Đề xuất đặt phụ tùng", layout="wide")

# --- Cổng đăng nhập (mật khẩu lưu trong Streamlit secrets) ---
if not A.login():
    st.stop()

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_STANDARD = os.path.join(HERE, "standard.xlsx")
PARTLIST_PATH = os.path.join(HERE, "partlist.xlsx")

# Cột hiển thị trong bảng review: gồm 'Cost' (sau Quantity) + 'Unit Price' + Std Min/Max (chỉ trên màn hình)
_BASE = [c for c in E.OUTPUT_COLUMNS if c != "No"]
_i = _BASE.index("Cost") + 1
DISPLAY_COLS = _BASE[:_i] + ["Unit Price"] + _BASE[_i:] + ["Std Min", "Std Max"]
# Cột do hệ thống tính -> khoá không cho sửa
LOCKED_COLS = ["Brand", "Item Code", "Item Name", "Cost", "Unit Price",
               "Analyzer", "Onhand HN", "Onhand HCM", "Std Min", "Std Max"]


# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _load_standard_cached(path, mtime):
    return E.load_standard(path)


@st.cache_data(show_spinner=False)
def _load_partlist_cached(path, mtime):
    return E.load_partlist(path)


def get_standard(uploaded):
    if uploaded is not None:
        tmp = os.path.join(tempfile.gettempdir(), "mgv_std_override.xlsx")
        with open(tmp, "wb") as f:
            f.write(uploaded.getbuffer())
        return E.load_standard(tmp), "file bạn tải lên"
    if os.path.exists(DEFAULT_STANDARD):
        return _load_standard_cached(DEFAULT_STANDARD, os.path.getmtime(DEFAULT_STANDARD)), "standard.xlsx (mặc định)"
    return None, None


def get_partlist():
    if os.path.exists(PARTLIST_PATH):
        return _load_partlist_cached(PARTLIST_PATH, os.path.getmtime(PARTLIST_PATH))
    return {}


def rows_to_df(rows):
    df = pd.DataFrame(rows)
    for c in DISPLAY_COLS:
        if c not in df.columns:
            df[c] = None
    return df[DISPLAY_COLS]


def df_to_rows(df):
    out = []
    for _, r in df.iterrows():
        code = str(r.get("Item Code") or "").strip()
        if not code or code.lower() == "nan":
            continue
        out.append({c: (None if pd.isna(r.get(c)) else r.get(c)) for c in DISPLAY_COLS})
    return out


# ===========================================================================
# SIDEBAR
# ===========================================================================
A.logout_button(st.sidebar)
st.sidebar.header("⚙️ Tiêu chuẩn tồn kho")
std_upload = st.sidebar.file_uploader(
    "Cập nhật tiêu chuẩn (tuỳ chọn)", type=["xlsx"],
    help="Để trống sẽ dùng standard.xlsx đi kèm. Muốn đổi tiêu chuẩn: sửa/thay file Excel này.",
)
(std_pack, std_src) = get_standard(std_upload)
if std_pack is None:
    st.sidebar.error("Không tìm thấy standard.xlsx. Hãy tải file tiêu chuẩn lên.")
    st.stop()
items, lookup = std_pack
partlist = get_partlist()
st.sidebar.success(f"Đã nạp {len(items)} linh kiện tiêu chuẩn\n\nNguồn: {std_src}")
_priced = sum(1 for v in partlist.values() if v.get("unit_price") not in (None, "")) if partlist else 0
st.sidebar.caption(f"PartList: {len(partlist)} mã · {_priced} có giá" if partlist else "PartList: (không có partlist.xlsx)")


# ===========================================================================
# MAIN
# ===========================================================================
st.title("📦 Đề xuất đặt phụ tùng dịch vụ — MediGroup Vietnam")
st.caption("Onhand ≤ Min → đặt (Max − Onhand). HN = kho đuôi 23, HCM = kho đuôi 21. "
           "Hàng đi đường/trung chuyển (đuôi 10, 30, 31…) không tính vào tồn. "
           "Cost = Quantity × Unit Price (USD).")

# --- ① & ② đầu vào ---
c1, c2 = st.columns([2, 1])
inv_upload = c1.file_uploader("① Import file tồn kho (PBI_Export hoặc MG Inventory Report)", type=["xlsx"])
region_choice = c2.radio("② Khu vực", ["HN", "HCM", "Cả nước"], horizontal=True)
regions = ["HN", "HCM"] if region_choice == "Cả nước" else [region_choice]

# --- ③ Item nhập tay (TRƯỚC khi phân tích) ---
st.subheader("③ Thêm item đặt thủ công (tuỳ chọn — sẽ gộp chung để kiểm tra 1 lần)")
st.caption("Nhập mã + số lượng. Gõ xong mã: nếu mã có trên hệ thống (tiêu chuẩn/PartList), "
           "ô Item Name sẽ TỰ ĐIỀN tên; nếu không có thì để trống cho anh tự nhập. "
           "Khi bấm Phân tích sẽ điền tồn HN/HCM, Min/Max, đơn giá/cost và gộp vào bảng đề xuất.")

MANUAL_COLS = ["Region", "Item Code", "Item Name", "Quantity",
               "Reason of Order", "Purpose", "Item Group"]


def _new_manual_seed():
    return pd.DataFrame([{
        "Region": regions[0], "Item Code": "", "Item Name": "", "Quantity": 1,
        "Reason of Order": "", "Purpose": E.DEFAULT_PURPOSE, "Item Group": "SPP",
    }])[MANUAL_COLS]


# Reset bảng nhập tay khi đổi khu vực (tránh Region cũ không hợp lệ)
if st.session_state.get("manual_region") != region_choice:
    st.session_state["manual_seed"] = _new_manual_seed()
    st.session_state["manual_region"] = region_choice
    st.session_state["manual_rev"] = st.session_state.get("manual_rev", 0) + 1

edited = st.data_editor(
    st.session_state["manual_seed"], key=f"manual_{st.session_state['manual_rev']}",
    width="stretch", num_rows="dynamic", column_order=MANUAL_COLS,
    column_config={
        "Region": st.column_config.SelectboxColumn(options=regions, required=True),
        "Item Name": st.column_config.TextColumn(help="Tự điền nếu mã có trên hệ thống"),
        "Item Group": st.column_config.SelectboxColumn(options=E.ITEM_GROUP_CHOICES),
        "Quantity": st.column_config.NumberColumn(min_value=0, step=1),
    },
)


def _cell(df, i, col):
    v = df.at[i, col] if col in df.columns else None
    return "" if v is None else str(v).strip()


def _autofill_names(df):
    """Điền Item Name khi mã tra được tên và ô name đang trống. Trả (df, có_thay_đổi)."""
    df = df.copy()
    changed = False
    for i in df.index:
        code = _cell(df, i, "Item Code")
        name = _cell(df, i, "Item Name")
        if code and code.lower() != "nan" and (not name or name.lower() == "nan"):
            nm = E.resolve_name(code, items, lookup, partlist)
            if nm:
                df.at[i, "Item Name"] = nm
                changed = True
    return df, changed


filled, changed = _autofill_names(edited)
if changed:                                   # re-seed editor với tên đã điền -> hiển thị ngay
    st.session_state["manual_seed"] = filled
    st.session_state["manual_rev"] += 1
    st.rerun()

manual_df = edited

# Mã đã nhập nhưng chưa tra được tên -> nhắc tự nhập
_missing = [_cell(manual_df, i, "Item Code") for i in manual_df.index
            if _cell(manual_df, i, "Item Code") and _cell(manual_df, i, "Item Code").lower() != "nan"
            and not _cell(manual_df, i, "Item Name")]
if _missing:
    st.caption("⚠️ Chưa tra được tên (anh tự nhập giúp): " + ", ".join(_missing))

# --- ④ Phân tích ---
if st.button("④ Phân tích", type="primary", width="stretch"):
    if inv_upload is None:
        st.warning("Hãy import file tồn kho trước.")
    else:
        tmp = os.path.join(tempfile.gettempdir(), "mgv_inv.xlsx")
        with open(tmp, "wb") as f:
            f.write(inv_upload.getbuffer())
        std_onhand, raw_onhand, raw_info, audit = E.load_inventory(tmp, lookup, len(items))

        result = {}
        for r in regions:
            auto = E.compute_orders(items, std_onhand, r, partlist=partlist)
            manual_rows = []
            for _, m in manual_df.iterrows():
                code = str(m.get("Item Code") or "").strip()
                if not code or code.lower() == "nan" or str(m.get("Region")) != r:
                    continue
                manual_rows.append(E.enrich_manual(
                    code, m.get("Quantity"), m.get("Reason of Order"),
                    m.get("Purpose"), m.get("Item Group"),
                    items, lookup, std_onhand, raw_onhand, raw_info,
                    region=r, typed_name=m.get("Item Name"), partlist=partlist))
            result[r] = rows_to_df(E.sort_rows(auto + manual_rows))

        st.session_state["result"] = result
        st.session_state["regions"] = regions
        st.session_state["audit"] = audit

# --- ⑤ Review & ⑥ Xuất ---
if "result" in st.session_state:
    result = st.session_state["result"]
    regions = st.session_state["regions"]
    st.info(f"Định dạng nhận diện: sheet '{st.session_state['audit'].get('sheet','?')}' · "
            "Kho ghi nhận theo đuôi mã: " + ", ".join(
        f"{k}×{v}" for k, v in sorted(st.session_state["audit"]["suffix_counts"].items())))

    st.subheader("⑤ Bảng đề xuất (đã gộp tự động + nhập tay — có thể chỉnh / xoá dòng)")
    edited = {}
    tabs = st.tabs([f"🏢 {r}" for r in regions])
    for tab, r in zip(tabs, regions):
        with tab:
            edited[r] = st.data_editor(
                result[r], key=f"review_{r}", width="stretch", num_rows="dynamic",
                column_config={
                    "Item Group": st.column_config.SelectboxColumn(options=E.ITEM_GROUP_CHOICES),
                    "Quantity": st.column_config.NumberColumn(min_value=0, step=1),
                    "Cost": st.column_config.NumberColumn(format="%.2f"),
                    "Unit Price": st.column_config.NumberColumn(format="%.2f"),
                },
                disabled=LOCKED_COLS,
            )
            ev = edited[r]
            qty = pd.to_numeric(ev["Quantity"], errors="coerce").fillna(0)
            up = pd.to_numeric(ev["Unit Price"], errors="coerce").fillna(0)
            st.caption(f"{len(ev)} item · tổng SL = {int(qty.sum())} · "
                       f"tổng Cost = {(qty * up).sum():,.2f} USD")

    st.subheader("⑥ Xuất file Excel")
    if st.button("Tạo file đặt hàng", type="primary"):
        sheets = {}
        for r in regions:
            rows = df_to_rows(edited[r])
            for row in rows:                    # recompute Cost theo qty đã chỉnh
                row["Cost"] = E.compute_cost(row.get("Quantity"), row.get("Unit Price"))
            sheets[r] = E.sort_rows(rows)
        out = os.path.join(tempfile.gettempdir(), "MGV_Order_List.xlsx")
        W.write_workbook(sheets, out)
        with open(out, "rb") as f:
            data = f.read()
        total = sum(len(v) for v in sheets.values())
        st.success(f"Đã tạo file: {total} dòng · sheet: {', '.join(sheets)}")
        st.download_button(
            "⬇️ Tải Order_List.xlsx", data=data,
            file_name="MGV_Order_List.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Import file tồn kho, (tuỳ chọn) nhập item thủ công, chọn khu vực, rồi bấm **Phân tích**.")
