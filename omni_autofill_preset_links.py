# -*- coding: utf-8 -*-
"""
OmniLogin + Google Sheet Autofill (Presets — Add & Delete)
- Điền Google Sheet (publish to web) vào form trên các tab/profile OmniLogin
- Presets: tick để mở nhiều link cùng lúc; có thể LƯU link tuỳ ý và XÓA preset đã tick.
"""

import csv
import json
import random
import time
from io import StringIO
from pathlib import Path
from urllib.parse import urlparse
import tkinter as tk
from tkinter import ttk, messagebox

import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from urllib.parse import urlsplit, urlunsplit, urlencode

def _build_like_current(self, base_url_or_current: str, path: str, query: dict | None = None) -> str:
    """
    Trả về URL mới có cùng scheme + domain với base_url_or_current,
    nhưng đổi path/query theo yêu cầu. Tự nhận biết site dạng hash-router (#/...)
    """
    u = urlsplit(base_url_or_current)
    # Phát hiện site SPA dùng hash (#/Register ...)
    is_hash_router = False
    if "/#/" in base_url_or_current:
        is_hash_router = True
    elif u.fragment and u.fragment.startswith("/"):
        is_hash_router = True

    # Chuẩn hoá đầu vào
    path = "/" + path.lstrip("/")
    q = urlencode(query or {})

    if is_hash_router:
        # Dạng https://host/#/Account/ChangeMoneyPassword
        new_frag = path  # fragment sẽ chứa path
        if q:
            new_frag = f"{new_frag}?{q}"
        return f"{u.scheme}://{u.netloc}/#{new_frag.lstrip('/')}"
    else:
        # Dạng https://host/Account/ChangeMoneyPassword
        return urlunsplit((u.scheme, u.netloc, path, q, ""))
def _url_change_money_pwd(self, url_like_current: str) -> str:
    return self._build_like_current(url_like_current, "/Account/ChangeMoneyPassword")

def _url_withdraw(self, url_like_current: str) -> str:
    # /Financial?type=withdraw
    return self._build_like_current(url_like_current, "/Financial", {"type": "withdraw"})


API_BASE = "http://localhost:35353"   # OmniLogin API
CONFIG_PATH = Path.home() / ".omni_autofill_multi.json"

# ===== PRESET MẶC ĐỊNH (không bị xoá bằng nút "xoá preset") =====
DEFAULT_PRESET_LINKS = [

    ("m.8eea2.buzz/Register", "https://m.8eea2.buzz/Register"),
    ("m.88clb1ax.buzz/Register", "https://m.88clb1ax.buzz/Register"),
]

# -------------------- Helpers (sheet/parsing) --------------------
def build_publish_url(publish_id: str, gid: str, fmt: str = "csv") -> str:
    publish_id = (publish_id or "").strip()
    gid = (gid or "").strip()
    fmt = (fmt or "csv").lower()
    if fmt not in ("csv", "tsv"):
        fmt = "csv"
    return f"https://docs.google.com/spreadsheets/d/e/{publish_id}/pub?gid={gid}&single=true&output={fmt}"

def parse_sheet_text(text: str, fmt: str = "csv"):
    delim = "," if (fmt or "csv").lower() == "csv" else "\t"
    reader = csv.reader(StringIO(text), delimiter=delim)
    return [row for row in reader]

def norm_header(s: str) -> str:
    try:
        import unicodedata
        s = unicodedata.normalize("NFD", s)
        s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    except Exception:
        pass
    return (s or "").strip().lower()

def find_columns(header_row):
    hmap = {norm_header(h): i for i, h in enumerate(header_row)}
    def pick(*keys):
        for k in keys:
            kk = norm_header(k)
            if kk in hmap:
                return hmap[kk]
        return None
    return {
        "username": pick("username", "ten tai khoan", "tai khoan", "tên tài khoản", "tk"),
        "password": pick("password", "mat khau", "mật khẩu", "pass", "mk"),
        "fullname": pick("fullname", "ho ten", "họ tên", "ten that", "họ và tên"),
        "phone":    pick("phone", "so dien thoai", "số điện thoại", "sdt", "dien thoai"),
        "email":    pick("email", "mail", "e-mail"),
        "birthday": pick("birthday", "ngay sinh", "ngày sinh", "dob", "date of birth"),
    }

def random_phone_11() -> str:
    return "0" + "".join(random.choice("0123456789") for _ in range(10))

def random_birth_year(y0=1980, y1=2000) -> str:
    return str(random.randint(y0, y1))

# -------------------- App --------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("OmniLogin + Google Sheet (Presets — Add & Delete)")
        self.geometry("635x470")
        self.resizable(True, True)

        # session map: profile_id -> {"driver":..., "debug_addr":..., "driver_path":...}
        self.sessions = {}
        self.current_profile_id = None

        # config vars
        self.profile_id_var = tk.StringVar()

        self.pub_id_var = tk.StringVar()
        self.gid_var = tk.StringVar()
        self.format_var = tk.StringVar(value="csv")
        self.header_row_var = tk.IntVar(value=1)
        self.row_index_var = tk.IntVar(value=1)

        # profile picker
        self.profile_select_var = tk.StringVar()

        # preset checkboxes (động)
        self.preset_vars = {}      # url -> BooleanVar
        self.preset_meta = {}      # url -> label
        self.user_presets = []     # [(label,url), ...] do người dùng lưu

        # custom links (CSV hoặc xuống dòng)
        self.open_custom_var = tk.StringVar(value="")

        # UI
        self._build_ui()
        self._load_config()
        self._rebuild_preset_ui()
        self._setup_autosave()

    # ---------- UI ----------
    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        f1 = ttk.LabelFrame(self, text="OmniLogin")
        f1.pack(fill="x", padx=10, pady=8)

        ttk.Label(f1, text="Profile ID:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(f1, textvariable=self.profile_id_var, width=11).grid(row=0, column=1, sticky="w", **pad)
        ttk.Button(f1, text="Mở Profile", command=self.ui_open_profile).grid(row=0, column=2, **pad)
        ttk.Button(f1, text="Đóng Profile", command=self.ui_close_profile).grid(row=0, column=3, **pad)

        # Preset checkboxes container
        ttk.Label(f1, text="Mở cùng lúc các link:").grid(row=1, column=0, sticky="nw", **pad)
        self.preset_frame = ttk.Frame(f1)
        self.preset_frame.grid(row=1, column=1, columnspan=2, sticky="w", **pad)

        # Nút xoá preset đã tick (thay cho nút Lưu TAB)
        ttk.Button(f1, text="− Xóa preset đã tick", command=self.delete_selected_presets)\
            .grid(row=1, column=3, sticky="ne", **pad)

        # Custom links
        ttk.Label(f1, text="Link khác:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(f1, textvariable=self.open_custom_var, width=44).grid(row=2, column=1, columnspan=2, sticky="we", **pad)
        ttk.Button(f1, text="+ Lưu vào Presets", command=self.save_custom_links_to_presets).grid(row=2, column=3, sticky="w", **pad)

        # Profile picker row
        ttk.Label(f1, text="Chọn profile đang mở:").grid(row=3, column=0, sticky="w", **pad)
        self.combo_profiles = ttk.Combobox(f1, textvariable=self.profile_select_var, width=11, state="readonly", values=[])
        self.combo_profiles.grid(row=3, column=1, sticky="w", **pad)
        ttk.Button(f1, text="Làm mới DS", command=self.ui_refresh_profile_list).grid(row=3, column=2, **pad)
        ttk.Button(f1, text="Đặt làm hiện tại", command=self.ui_pick_profile_current).grid(row=3, column=3, **pad)

        f2 = ttk.LabelFrame(self, text="Google Sheet (Publish to web)")
        f2.pack(fill="x", padx=10, pady=8)

        ttk.Label(f2, text="Publish ID:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(f2, textvariable=self.pub_id_var, width=44).grid(row=0, column=1, columnspan=3, sticky="we", **pad)

        ttk.Label(f2, text="gid:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(f2, textvariable=self.gid_var, width=12).grid(row=1, column=1, sticky="w", **pad)
        ttk.Radiobutton(f2, text="CSV", value="csv", variable=self.format_var).grid(row=1, column=2, sticky="w", **pad)
        ttk.Radiobutton(f2, text="TSV", value="tsv", variable=self.format_var).grid(row=1, column=3, sticky="w", **pad)

        ttk.Label(f2, text="Header row:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Spinbox(f2, from_=1, to=999, textvariable=self.header_row_var, width=6).grid(row=2, column=1, sticky="w", **pad)
        ttk.Label(f2, text="Row bắt đầu:").grid(row=2, column=2, sticky="e", **pad)
        ttk.Spinbox(f2, from_=1, to=999999, textvariable=self.row_index_var, width=10).grid(row=2, column=3, sticky="w", **pad)

        # Hàng nút điền
        ttk.Button(f2, text="Xem dòng…", command=self.ui_preview_row).grid(row=3, column=1, **pad)
        ttk.Button(f2, text="Điền → Profile hiện tại", command=self.ui_fill_current).grid(row=3, column=2, **pad)
        ttk.Button(f2, text="Điền → TẤT CẢ profile (tăng dòng)", command=self.ui_fill_all_inc).grid(row=3, column=3, **pad)

        # Multi-tab (no increment)
        ttk.Button(f2, text="Điền TAB hiện tại", command=self.on_fill_current_tab_no_inc).grid(row=4, column=2, **pad)
        ttk.Button(f2, text="Điền MỌI TAB (profile hiện tại)", command=self.on_fill_all_tabs_current_profile).grid(row=4, column=3, **pad)
        ttk.Button(
            f2,
            text="Đổi pass rút tiền (theo domain hiện tại) → Rút tiền",
            command=self.ui_set_money_pwd_123456_and_go_withdraw_auto
        ).grid(row=7, column=3, **pad)

        ttk.Button(self, text="Thoát", command=self.destroy).pack(pady=8)

    # ---------- Preset helpers ----------
    def _hostname_label(self, url: str) -> str:
        try:
            u = urlparse(url)
            host = (u.netloc or u.path or url).strip("/ ")
            if u.path and "reg" in u.path.lower():
                host += u.path
            return host[:40]
        except Exception:
            return url[:40]

    def _all_presets(self):
        seen = set(); merged = []
        for lbl, u in DEFAULT_PRESET_LINKS + self.user_presets:
            if u not in seen:
                merged.append((lbl, u)); seen.add(u)
        return merged

    def _rebuild_preset_ui(self):
        for w in self.preset_frame.winfo_children():
            w.destroy()
        self.preset_vars.clear(); self.preset_meta.clear()

        presets = self._all_presets()
        cols = 3
        for i, (label, url) in enumerate(presets):
            var = tk.BooleanVar(value=False)
            self.preset_vars[url] = var
            self.preset_meta[url] = label
            r, c = divmod(i, cols)
            ttk.Checkbutton(self.preset_frame, text=label, variable=var).grid(row=r, column=c, sticky="w", padx=6, pady=2)

        self._attach_autosave_to_preset_vars()
        # set lại trạng thái tick từ config (nếu có)
        sel = getattr(self, "_selected_urls_from_cfg", set())
        for url, var in self.preset_vars.items():
            if url in sel:
                var.set(True)

    def _persist_presets(self):
        data = {}
        if CONFIG_PATH.exists():
            try:
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            except Exception:
                data = {}
        data["user_presets"] = self.user_presets
        data["open_presets"] = [u for u, v in self.preset_vars.items() if v.get()]
        data["open_custom"]  = self.open_custom_var.get().strip()
        try:
            CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không lưu preset: {e}")

    def save_custom_links_to_presets(self):
        raw = (self.open_custom_var.get() or "")
        links = [u.strip() for u in raw.replace("\n", ",").split(",") if u.strip()]
        if not links:
            messagebox.showwarning("Nhắc", "Chưa nhập link nào ở ô Link khác.")
            return
        added = 0
        for u in links:
            if not (u.startswith("http://") or u.startswith("https://")):
                continue
            label = self._hostname_label(u)
            if all(u != uu for _, uu in self.user_presets) and all(u != uu for _, uu in DEFAULT_PRESET_LINKS):
                self.user_presets.append((label, u))
                added += 1
        if added:
            self._rebuild_preset_ui()
            for lbl, u in self.user_presets[-added:]:
                if u in self.preset_vars:
                    self.preset_vars[u].set(True)
            self._persist_presets()
            messagebox.showinfo("OK", f"Đã lưu {added} link vào presets.")
        else:
            messagebox.showinfo("Thông báo", "Không có link hợp lệ mới để lưu.")

    def delete_selected_presets(self):
        selected = [u for u, v in self.preset_vars.items() if v.get()]
        if not selected:
            messagebox.showwarning("Nhắc", "Hãy tick vào các preset muốn xóa.")
            return
        user_urls = {u for _, u in self.user_presets}
        keep = []
        removed = 0
        for lbl, u in self.user_presets:
            if u in selected:
                removed += 1
            else:
                keep.append((lbl, u))
        self.user_presets = keep
        self._rebuild_preset_ui()
        self._persist_presets()
        skipped = len([u for u in selected if u not in user_urls])
        if removed and not skipped:
            messagebox.showinfo("OK", f"Đã xóa {removed} preset.")
        elif removed and skipped:
            messagebox.showinfo("OK", f"Đã xóa {removed} preset. ({skipped} preset mặc định không thể xoá)")
        else:
            messagebox.showinfo("Thông báo", "Không có preset do bạn tạo trong các mục đã tick (preset mặc định không xoá).")

    # ---------- Config ----------
    def _load_config(self):
        self._selected_urls_from_cfg = set()
        try:
            if CONFIG_PATH.exists():
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                self.pub_id_var.set(data.get("publish_id", ""))
                self.gid_var.set(data.get("gid", ""))
                self.format_var.set(data.get("format", "csv"))
                self.header_row_var.set(int(data.get("header_row", 1)))
                self.row_index_var.set(int(data.get("row_index", 1)))
                self.profile_id_var.set(data.get("last_profile_id", ""))

                ups = data.get("user_presets", [])
                if isinstance(ups, list):
                    self.user_presets = [(str(a), str(b)) for a, b in ups if isinstance(a, str) and isinstance(b, str)]

                self._selected_urls_from_cfg = set(data.get("open_presets", []))
                if not self._selected_urls_from_cfg:  # fallback khoá cũ
                    if data.get("open_google"):   self._selected_urls_from_cfg.add("https://www.google.com")
                    if data.get("open_register"): self._selected_urls_from_cfg.add("https://m.8eea2.buzz/Register")

                self.open_custom_var.set(data.get("open_custom", ""))
        except Exception as e:
            messagebox.showwarning("Cảnh báo", f"Không đọc được config: {e}")

    def _save_config(self):
        data = {
            "publish_id": self.pub_id_var.get().strip(),
            "gid": self.gid_var.get().strip(),
            "format": self.format_var.get(),
            "header_row": int(self.header_row_var.get() or 1),
            "row_index": int(self.row_index_var.get() or 1),
            "last_profile_id": self.profile_id_var.get().strip(),
            "user_presets": self.user_presets,
            "open_presets": [url for url, var in self.preset_vars.items() if var.get()],
            "open_custom": self.open_custom_var.get().strip(),
        }
        try:
            CONFIG_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không lưu config: {e}")

    def _setup_autosave(self):
        def trigger(*_):
            self.after(150, self._save_config)
        for var in (
            self.profile_id_var,
            self.pub_id_var, self.gid_var, self.format_var,
            self.header_row_var, self.row_index_var,
            self.open_custom_var
        ):
            var.trace_add("write", trigger)
        self._attach_autosave_to_preset_vars()

    def _attach_autosave_to_preset_vars(self):
        def trigger(*_):
            self.after(150, self._save_config)
        for var in self.preset_vars.values():
            try:
                var.trace_add("write", trigger)
            except Exception:
                pass

    # ---------- OmniLogin / Sessions ----------
    def _attach_driver(self, debug_addr, driver_path):
        opts = webdriver.ChromeOptions()
        opts.add_experimental_option("debuggerAddress", debug_addr)
        drv = webdriver.Chrome(service=Service(driver_path), options=opts)
        drv.implicitly_wait(10)
        return drv

    def _wait_dom_ready(self, drv, timeout=25):
        WebDriverWait(drv, timeout).until(lambda d: d.execute_script("return document.readyState") == "complete")
        time.sleep(0.4)

    def open_profile_and_attach(self, profile_id: str):
        params = {"profile_id": profile_id, "headless": "false"}
        resp = requests.get(f"{API_BASE}/open", params=params, timeout=30)
        data = resp.json()
        if data.get("error"):
            try:
                requests.get(f"{API_BASE}/stop/{profile_id}", timeout=15)
            except Exception:
                pass
            resp = requests.get(f"{API_BASE}/open", params=params, timeout=30)
            data = resp.json()
            if data.get("error"):
                raise RuntimeError(f"Không mở được profile {profile_id}: {data}")
        debug_addr = data.get("remote_debug_address") or data.get("remote_debugger_address")
        driver_path = data.get("drive_location") or data.get("driver_location") or data.get("drive")
        if not debug_addr or not driver_path:
            raise RuntimeError("Thiếu remote_debug_address / driver_path.")
        driver = self._attach_driver(debug_addr, driver_path)
        self.sessions[profile_id] = {"driver": driver, "debug_addr": debug_addr, "driver_path": driver_path}
        self.current_profile_id = profile_id
        self.driver = driver
        self.ui_refresh_profile_list()
        self.profile_select_var.set(profile_id)
        self._open_selected_links_in_tabs(driver)
        return driver

    def _selected_urls(self):
        urls = [u for u, v in self.preset_vars.items() if v.get()]
        extra_raw = (self.open_custom_var.get() or "")
        urls += [u.strip() for u in extra_raw.replace("\n", ",").split(",") if u.strip()]
        return [u for u in urls if u.startswith("http://") or u.startswith("https://")]

    def _open_selected_links_in_tabs(self, drv):
        urls = self._selected_urls()
        if not urls:
            return
        first = True
        for url in urls:
            try:
                if first:
                    drv.get(url); self._wait_dom_ready(drv); first = False
                else:
                    try:
                        drv.switch_to.new_window('tab')
                    except Exception:
                        drv.execute_script("window.open('about:blank','_blank');")
                    drv.switch_to.window(drv.window_handles[-1])
                    drv.get(url); self._wait_dom_ready(drv)
            except Exception:
                pass

    def get_driver(self, profile_id=None):
        pid = profile_id or self.current_profile_id
        ses = self.sessions.get(pid)
        if not ses:
            raise RuntimeError(f"Chưa mở profile {pid}")
        return ses["driver"]

    def stop_profile(self, profile_id: str):
        ses = self.sessions.pop(profile_id, None)
        if ses:
            try: ses["driver"].quit()
            except Exception: pass
        try: requests.get(f"{API_BASE}/stop/{profile_id}", timeout=10)
        except Exception: pass
        if self.current_profile_id == profile_id:
            self.current_profile_id = None
        self.ui_refresh_profile_list()

    # ---------- UI Handlers ----------
    def ui_open_profile(self):
        pid = (self.profile_id_var.get() or "").strip()
        if not pid:
            messagebox.showerror("Lỗi", "Chưa nhập Profile ID.")
            return
        try:
            self.open_profile_and_attach(pid)
            messagebox.showinfo("OK", f"Đã mở profile {pid}.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không mở được profile:\n{e}")

    def ui_close_profile(self):
        pid = (self.profile_id_var.get() or "").strip()
        if not pid:
            messagebox.showerror("Lỗi", "Chưa nhập Profile ID.")
            return
        self.stop_profile(pid)
        messagebox.showinfo("OK", f"Đã đóng {pid}.")

    def ui_refresh_profile_list(self):
        vals = list(self.sessions.keys())
        self.combo_profiles["values"] = vals
        if self.current_profile_id in vals:
            self.profile_select_var.set(self.current_profile_id)
        elif vals:
            self.profile_select_var.set(vals[0])
        else:
            self.profile_select_var.set("")

    def ui_pick_profile_current(self):
        sel = (self.profile_select_var.get() or "").strip()
        if not sel:
            messagebox.showerror("Lỗi", "Chưa chọn profile trong combobox.")
            return
        if sel not in self.sessions:
            messagebox.showerror("Lỗi", "Profile này chưa mở.")
            return
        self.current_profile_id = sel
        self.profile_id_var.set(sel)
        messagebox.showinfo("OK", f"Đã đặt profile hiện tại = {sel}")

    def _fetch_sheet_rows(self):
        pub_id = (self.pub_id_var.get() or "").strip()
        gid = (self.gid_var.get() or "").strip()
        fmt = self.format_var.get()
        if not pub_id or not gid:
            raise RuntimeError("Thiếu Publish ID hoặc gid.")
        url = build_publish_url(pub_id, gid, fmt)
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        rows = parse_sheet_text(r.text, fmt)
        if not rows:
            raise RuntimeError("Sheet rỗng.")
        return rows

    def ui_preview_row(self):
        try:
            rows = self._fetch_sheet_rows()
            header_idx = int(self.header_row_var.get() or 1) - 1
            header = rows[header_idx]
            idxmap = find_columns(header)
            data_rows = rows[header_idx + 1:]
            ridx = int(self.row_index_var.get() or 1) - 1
            if ridx < 0 or ridx >= len(data_rows):
                raise RuntimeError("Row vượt dữ liệu.")
            row = data_rows[ridx]
            def get(col):
                i = idxmap.get(col)
                return row[i] if (i is not None and i < len(row)) else ""
            msg = (
                f"Row #{ridx+1}\n"
                f"- username: {get('username')}\n"
                f"- password: {get('password')} (nếu trống sẽ = username+20)\n"
                f"- fullname: {get('fullname')}\n"
                f"- phone   : {get('phone')} (nếu trống sẽ random 11 số)\n"
                f"- email   : {get('email')}\n"
                f"- birthday: {get('birthday')}\n"
            )
            messagebox.showinfo("Preview", msg)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không đọc sheet:\n{e}")

    # ---------- Fill logic ----------
    def _js_fill(self):
        return r"""
    (function(user, pass, fullname, phone, email, birth){
      function setNative(el, v){
        if(!el) return;
        const proto  = Object.getPrototypeOf(el) || HTMLInputElement.prototype;
        const desc   = Object.getOwnPropertyDescriptor(proto, 'value');
        const setter = desc && desc.set;
        if (setter) setter.call(el, ''); else el.value = '';
        if (setter) setter.call(el, String(v)); else el.value = String(v);
        el.dispatchEvent(new Event('input',  {bubbles:true}));
        el.dispatchEvent(new Event('change', {bubbles:true}));
        el.dispatchEvent(new Event('blur',   {bubbles:true}));
      }
      function setVal(el, v){
        if(!el || v==null || v==='') return false;
        try{ el.scrollIntoView({block:'center'});}catch(e){}
        try{ el.focus(); }catch(e){}
        setNative(el, v);
        return true;
      }
      function norm(s){
        return (s||'').toLowerCase()
          .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
          .replace(/\s+/g,' ').trim();
      }
      function textHints(el){
        let t = (el.getAttribute('placeholder') || el.getAttribute('aria-label') || '');
        const id = el.id;
        if (id) {
          try{
            const lab = document.querySelector('label[for="'+CSS.escape(id)+'"]');
            if (lab) t += ' ' + lab.textContent;
          }catch(e){}
        }
        let p = el.parentElement, hop = 0;
        while (p && hop < 3){
          const l = p.querySelector('label');
          if (l){ t += ' ' + (l.textContent||''); break; }
          p = p.parentElement; hop++;
        }
        return norm(t);
      }
      function hasAny(text, keys){
        text = norm(text);
        return keys.some(k => text.includes(norm(k)));
      }
      const K_ACC  = ['tai khoan','tài khoản','account','username','user name'];
      const K_NAME = ['ho va ten','họ tên','họ và tên','ten that','tên that','ten day du','full name','fullname','real name','realname','ten cua ban'];
      const K_PHONE= ['so dien thoai','số điện thoại','mobile','phone','sdt'];
      const K_EMAIL= ['email','e-mail','mail'];
      const K_BIRTH= ['ngay sinh','ngày sinh','dob','date of birth'];
      const root = document.querySelector('.el-dialog__body, .el-dialog, .modal, .popup, .dialog, body') || document;
      const pick = sel => root.querySelector(sel);
      const all  = sel => Array.from(root.querySelectorAll(sel));
      let userEl =
          pick('input[ng-model*="account"]') ||
          pick('input[placeholder*="2-15"]') ||
          pick('input[placeholder*="ten tai khoan"],input[placeholder*="tên tài khoản"],input[placeholder*="tai khoan"]') ||
          pick('input[name*="account"],input[id*="account"],input[name*="user"],input[id*="user"]');
      let passList = all('input[type="password"]');
      let passEl   = passList[0] || null;
      let cpassEl  = passList[1] || null;
      if (!cpassEl) {
        cpassEl = pick('input[ng-model*="confirm"], input[placeholder*="xac nhan"][type="password"], input[placeholder*="xác nhận"][type="password"]');
      }
      let nameEl =
          pick('input[ng-model*="real"],input[ng-model*="true"],input[ng-model*="full"],input[ng-model*="name"]:not([name*="account"]):not([id*="account"])') ||
          pick('input[name*="fullname"],input[id*="fullname"],input[name*="real"],input[id*="real"]') ||
          pick('input[placeholder*="ho va ten"],input[placeholder*="họ và tên"],input[placeholder*="ho ten"],input[placeholder*="ten that"],input[placeholder*="ten day du"]');
      if (!nameEl){
        const cand = all('input[type="text"], input:not([type]), input[type="search"]').filter(el=>{
          const t = textHints(el);
          return hasAny(t, K_NAME) && !hasAny(t, K_ACC);
        });
        if (cand.length) nameEl = cand[0];
      }
      if (!nameEl && userEl){
        try{
          const inputs = all('input[type="text"], input:not([type]), input[type="search"]');
          let idx = inputs.indexOf(userEl);
          if (idx >= 0 && inputs[idx+1]) nameEl = inputs[idx+1];
        }catch(e){}
      }
      let phoneEl =
          pick('input[type="tel"]') ||
          pick('input[ng-model*="phone"],input[name*="phone"],input[id*="phone"]') ||
          pick('input[placeholder*="so dien thoai"],input[placeholder*="số điện thoại"],input[placeholder*="dien thoai"]');
      let emailEl =
          pick('input[type="email"],input[ng-model*="email"]') ||
          pick('input[placeholder*="email"],input[name*="mail"],input[id*="mail"]');
      let birthEl =
          pick('input[type="date"]') ||
          pick('input[ng-model*="birth"],input[ng-model*="dob"]') ||
          pick('input[placeholder*="ngay sinh"],input[placeholder*="ngày sinh"]');
      const ok = {};
      ok.user  = setVal(userEl,  user);
      ok.pass  = setVal(passEl,  pass);
      ok.cpass = setVal(cpassEl, pass);
      ok.name  = setVal(nameEl,  fullname);
      ok.phone = setVal(phoneEl, phone);
      ok.email = setVal(emailEl, email);
      ok.birth = setVal(birthEl, birth);
      return ok;
    })(arguments[0], arguments[1], arguments[2], arguments[3], arguments[4], arguments[5]);
    """
    def ui_set_money_pwd_123456_and_go_withdraw_auto(self):
        """Dùng domain hiện tại để vào ChangeMoneyPassword -> đặt 123456 -> chuyển qua Withdraw."""
        if not self.current_profile_id:
            messagebox.showerror("Lỗi", "Chưa có profile hiện tại.")
            return
        try:
            drv = self.get_driver(self.current_profile_id)
            try: drv.switch_to.default_content()
            except Exception: pass
            self._wait_dom_ready(drv)

            # domain nguồn (URL hiện tại)
            current_url = drv.current_url

            # 1) tới trang đổi mật khẩu rút tiền (cùng domain)
            change_url = self._url_change_money_pwd(current_url)
            drv.get(change_url); self._wait_dom_ready(drv)

            # 2) điền 123456 hai ô và gửi
            res = drv.execute_script(self._js_set_money_pwd_and_submit(), "123456")
            time.sleep(1.0)

            # 3) tới trang rút tiền (cùng domain)
            withdraw_url = self._url_withdraw(change_url)  # hoặc dùng current_url cũng OK
            drv.get(withdraw_url); self._wait_dom_ready(drv)

            messagebox.showinfo("OK", f"Đã đặt mật khẩu và mở Rút tiền.\n{res}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Thao tác thất bại:\n{e}")

    def _fill_one_profile(self, profile_id: str, row_index_1_based: int):
        rows = self._fetch_sheet_rows()
        header_idx = int(self.header_row_var.get() or 1) - 1
        header = rows[header_idx]
        idxmap = find_columns(header)
        data_rows = rows[header_idx + 1:]

        ridx = int(row_index_1_based or 1) - 1
        if ridx < 0 or ridx >= len(data_rows):
            raise RuntimeError(f"Row #{row_index_1_based} vượt dữ liệu.")

        row = data_rows[ridx]
        def get(col):
            i = idxmap.get(col)
            return row[i] if (i is not None and i < len(row)) else ""

        username = (get("username") or "").strip()
        fullname = (get("fullname") or "").strip()
        phone = (get("phone") or "").strip()
        email = (get("email") or "").strip()
        birth = (get("birthday") or "").strip()
        if not birth:
            birth = random_birth_year(1980, 2000)

        pwd_sheet = (get("password") or "").strip()
        password = pwd_sheet or (username + "20")
        if not phone:
            phone = random_phone_11()

        drv = self.get_driver(profile_id)
        try:
            drv.switch_to.default_content()
        except Exception:
            pass

        self._wait_dom_ready(drv)
        res = drv.execute_script(self._js_fill(), username, password, fullname, phone, email, birth)
        return res

    # ---------- Buttons ----------
    def ui_fill_current(self):
        if not self.current_profile_id:
            messagebox.showerror("Lỗi", "Chưa có profile hiện tại.")
            return
        try:
            res = self._fill_one_profile(self.current_profile_id, int(self.row_index_var.get() or 1))
            messagebox.showinfo("OK", f"Đã điền cho {self.current_profile_id}.\nFields: {res}")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Điền form thất bại:\n{e}")

    def on_fill_current_tab_no_inc(self):
        if not self.current_profile_id:
            messagebox.showerror("Lỗi", "Chưa có profile hiện tại.")
            return
        try:
            self._fill_one_profile(self.current_profile_id, int(self.row_index_var.get() or 1))
            messagebox.showinfo("OK", "Đã điền TAB hiện tại (không tăng dòng).")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không điền được TAB hiện tại:\n{e}")

    def ui_fill_all_inc(self):
        if not self.sessions:
            messagebox.showerror("Lỗi", "Chưa mở profile nào.")
            return
        start_row = int(self.row_index_var.get() or 1)
        errs, filled = [], 0
        for i, pid in enumerate(list(self.sessions.keys())):
            try:
                self._fill_one_profile(pid, start_row + i)
                filled += 1
            except Exception as e:
                errs.append(f"{pid}: {e}")
        try:
            self.row_index_var.set(start_row + filled)
        except Exception:
            pass
        if errs:
            messagebox.showwarning("Xong (có lỗi)", "Một số profile lỗi:\n" + "\n".join(errs))
        else:
            messagebox.showinfo("OK", "Đã điền xong cho tất cả profile.")

    def on_fill_all_tabs_current_profile(self):
        if not self.current_profile_id:
            messagebox.showerror("Lỗi", "Chưa có profile hiện tại.")
            return
        pid = self.current_profile_id
        try:
            drv = self.get_driver(pid)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không lấy được driver: {e}")
            return

        handles_snapshot = drv.window_handles[:]
        if not handles_snapshot:
            messagebox.showerror("Lỗi", "Profile hiện tại chưa có tab nào đang mở.")
            return

        start_row = int(self.row_index_var.get() or 1)
        errs, processed = [], 0
        for j, h in enumerate(handles_snapshot):
            try:
                if h not in drv.window_handles:
                    continue
                drv.switch_to.window(h)
                try:
                    drv.switch_to.default_content()
                except Exception:
                    pass
                self._wait_dom_ready(drv)
                self._fill_one_profile(pid, start_row)  # dùng CHUNG 1 dòng
                processed += 1
            except Exception as e:
                errs.append(f"Tab #{j+1}: {e}")

        if errs:
            messagebox.showwarning("Xong (có lỗi)", "Một số tab lỗi:\n" + "\n".join(errs))
        else:
            messagebox.showinfo("OK", f"Đã điền xong {processed} tab (chung dòng {start_row}).")

# --------------- main ---------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
