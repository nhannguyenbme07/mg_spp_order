"""
Màn đăng nhập bằng mật khẩu, lưu trong Streamlit secrets.

Hỗ trợ 2 kiểu cấu hình trong .streamlit/secrets.toml (hoặc Secrets trên Cloud):

  # (a) Một mật khẩu chung
  app_password = "matkhau_cua_ban"

  # (b) Nhiều tài khoản (lọc từng người)
  [passwords]
  nhan = "matkhau_nhan"
  tuan = "matkhau_tuan"

Cách dùng trong app.py:
  import auth as A
  if not A.login():
      st.stop()
  ...
  A.logout_button()   # trong phần sidebar
"""
import hmac
import streamlit as st

FLAG = "auth_ok"
USER = "auth_user"


def _get_config():
    """Trả về ('single', password) | ('multi', {user: pass}) | (None, None)."""
    try:
        s = st.secrets
        if "passwords" in s:
            return "multi", dict(s["passwords"])
        if "users" in s:
            return "multi", dict(s["users"])
        for k in ("app_password", "password", "APP_PASSWORD"):
            if k in s:
                return "single", str(s[k])
    except Exception:
        return None, None
    return None, None


def login(title="🔒 Đăng nhập — MediGroup Vietnam"):
    """Trả True nếu đã đăng nhập; nếu chưa thì hiện form và trả False (caller nên st.stop())."""
    if st.session_state.get(FLAG):
        return True

    mode, cfg = _get_config()
    if mode is None:
        st.title(title)
        st.error(
            "Chưa cấu hình mật khẩu. Hãy thêm vào Streamlit **secrets**:\n\n"
            "```toml\napp_password = \"mat_khau_cua_ban\"\n```\n\n"
            "Trên Streamlit Cloud: mở app → **Settings → Secrets** và dán dòng trên."
        )
        st.stop()

    st.title(title)
    st.caption("Ứng dụng nội bộ — vui lòng đăng nhập để tiếp tục.")

    with st.form("login_form"):
        user = st.text_input("Tài khoản") if mode == "multi" else None
        pw = st.text_input("Mật khẩu", type="password")
        submitted = st.form_submit_button("Đăng nhập")

    if submitted:
        ok = False
        if mode == "single":
            ok = hmac.compare_digest(pw, cfg)
        else:
            expected = cfg.get((user or "").strip())
            if expected is not None:
                ok = hmac.compare_digest(pw, str(expected))
        if ok:
            st.session_state[FLAG] = True
            if mode == "multi":
                st.session_state[USER] = (user or "").strip()
            st.rerun()
        else:
            st.error("Sai tài khoản hoặc mật khẩu.")
    return False


def logout_button(location=None):
    """Hiện thông tin đăng nhập + nút Đăng xuất (mặc định trong sidebar)."""
    if not st.session_state.get(FLAG):
        return
    loc = location or st.sidebar
    who = st.session_state.get(USER)
    loc.caption(f"👤 {who}" if who else "👤 Đã đăng nhập")
    if loc.button("Đăng xuất"):
        for k in (FLAG, USER):
            st.session_state.pop(k, None)
        st.rerun()
