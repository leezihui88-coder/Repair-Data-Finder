#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MOXA Repair Data Finder  v1.0
===============================
自動從 PDM (Windchill) 查找成品料號下的 1199 PCB 料號，
並在 GeDCC (DMP) 取得對應線路圖連結。

v1.0 功能：
- 繼承 Schematic Finder v1.4 所有功能
- 自動記住 PDM / GeDCC 帳號密碼（使用 Windows Credential Manager）
- 自動填入登入表單，免除每次手動輸入
- 密碼變更時詢問是否更新儲存的密碼
"""

# ── 標準 import ─────────────────────────────────────────────────────
import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.scrolledtext as scrolledtext
import threading
import queue
import time
import webbrowser
import csv
import os
import re
import json
from datetime import datetime

# 密碼管理
try:
    import keyring
    _KEYRING_OK = True
except ImportError:
    _KEYRING_OK = False

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException,
    StaleElementReferenceException, WebDriverException
)
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.chrome.options import Options as ChromeOptions


# ── 色彩配置（專業深色主題）─────────────────────────────────────────
C = {
    'bg_dark':    '#12161F',
    'bg_panel':   '#1C2133',
    'bg_card':    '#232B3E',
    'bg_input':   '#1A2030',
    'accent':     '#00B4B4',
    'accent_dim': '#007A7A',
    'text':       '#DDE3F0',
    'text_dim':   '#7E8BAA',
    'success':    '#3FB950',
    'warning':    '#E3A740',
    'error':      '#F85149',
    'border':     '#2E3857',
    'row_sel':    '#1A3A5C',
    'log_bg':     '#0D1117',
    'log_fg':     '#7EE787',
    'ts':         '#484F58',
    'btn_save':   '#1A4A2E',
}

PDM_URL    = "https://pap.moxa.com/Windchill/app/"
GEDCC_URL  = "http://global-gedcc.moxa.com/DMP/public/login/ShowLogin.jsp"
GEDCC_HOME = "http://global-gedcc.moxa.com/DMP/private/user_index.jsp"


# ══════════════════════════════════════════════════════════════════════
#  密碼管理器（Windows Credential Manager via keyring）
# ══════════════════════════════════════════════════════════════════════
class CredentialManager:
    """
    以 Windows Credential Manager 安全儲存帳密。
    keyring 不可用時自動 fallback 到加密 JSON（Base64 混淆，非強加密）。
    """
    _SVC_PDM   = "MOXA-RepairDataFinder-PDM"
    _SVC_GEDCC = "MOXA-RepairDataFinder-GeDCC"
    _UN_KEY    = "__rdf__"
    # fallback 檔路徑
    _FB_PATH   = os.path.join(os.path.expanduser("~"), ".rdf_creds.json")

    # ── 取得 ──────────────────────────────────────────────────────────
    @classmethod
    def load(cls, system: str) -> tuple:
        """回傳 (username, password)，無資料則 ('', '')"""
        key = cls._SVC_PDM if system == 'pdm' else cls._SVC_GEDCC
        if _KEYRING_OK:
            try:
                raw = keyring.get_password(key, cls._UN_KEY)
                if raw:
                    d = json.loads(raw)
                    return d.get("u", ""), d.get("p", "")
            except Exception:
                pass
        # fallback: JSON 檔
        try:
            with open(cls._FB_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            enc = data.get(key, "")
            if enc:
                import base64
                d = json.loads(base64.b64decode(enc).decode())
                return d.get("u", ""), d.get("p", "")
        except Exception:
            pass
        return "", ""

    # ── 儲存 ──────────────────────────────────────────────────────────
    @classmethod
    def save(cls, system: str, user: str, pwd: str):
        key = cls._SVC_PDM if system == 'pdm' else cls._SVC_GEDCC
        payload = json.dumps({"u": user, "p": pwd})
        if _KEYRING_OK:
            try:
                keyring.set_password(key, cls._UN_KEY, payload)
                return
            except Exception:
                pass
        # fallback
        try:
            import base64
            enc = base64.b64encode(payload.encode()).decode()
            try:
                with open(cls._FB_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}
            data[key] = enc
            with open(cls._FB_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f)
        except Exception:
            pass

    # ── 刪除 ──────────────────────────────────────────────────────────
    @classmethod
    def delete(cls, system: str):
        key = cls._SVC_PDM if system == 'pdm' else cls._SVC_GEDCC
        if _KEYRING_OK:
            try:
                keyring.delete_password(key, cls._UN_KEY)
            except Exception:
                pass
        try:
            with open(cls._FB_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            data.pop(key, None)
            with open(cls._FB_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f)
        except Exception:
            pass


# ══════════════════════════════════════════════════════════════════════
#  帳密輸入對話框（深色主題）
# ══════════════════════════════════════════════════════════════════════
class CredentialDialog(tk.Toplevel):
    """
    深色主題帳號密碼對話框。
    result = (username, password, remember:bool)  or  None（取消）
    """
    def __init__(self, parent, system_label: str,
                 prefill_user: str = "", message: str = ""):
        super().__init__(parent)
        self.result = None
        self.title(f"登入 — {system_label}")
        self.configure(bg=C['bg_panel'])
        self.resizable(False, False)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        pad = dict(padx=20, pady=6)

        # ── 標題列 ────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=C['bg_dark'], pady=12)
        hdr.pack(fill='x')
        tk.Label(hdr, text='🔐  帳號密碼',
                 font=('Segoe UI', 13, 'bold'),
                 fg=C['accent'], bg=C['bg_dark']).pack(side='left', padx=18)
        tk.Label(hdr, text=system_label,
                 font=('Segoe UI', 10),
                 fg=C['text_dim'], bg=C['bg_dark']).pack(side='left')

        # ── 提示訊息（有時才顯示）────────────────────────────────────
        if message:
            tk.Label(self, text=message,
                     font=('Segoe UI', 10),
                     fg=C['warning'], bg=C['bg_panel'],
                     wraplength=320, justify='left').pack(
                         anchor='w', **pad)

        body = tk.Frame(self, bg=C['bg_panel'])
        body.pack(fill='x', padx=20, pady=10)

        # ── 帳號 ──────────────────────────────────────────────────────
        tk.Label(body, text='帳號', font=('Segoe UI', 10),
                 fg=C['text_dim'], bg=C['bg_panel']).grid(
                     row=0, column=0, sticky='w', pady=(0, 4))
        self._user_var = tk.StringVar(value=prefill_user)
        uf = tk.Frame(body, bg=C['border'], bd=1, relief='solid')
        uf.grid(row=1, column=0, sticky='ew', pady=(0, 10))
        self._user_ent = tk.Entry(uf,
            textvariable=self._user_var,
            font=('Consolas', 12), width=28,
            bg=C['bg_input'], fg=C['text'],
            insertbackground=C['accent'], relief='flat', bd=6)
        self._user_ent.pack(fill='x')

        # ── 密碼 ──────────────────────────────────────────────────────
        tk.Label(body, text='密碼', font=('Segoe UI', 10),
                 fg=C['text_dim'], bg=C['bg_panel']).grid(
                     row=2, column=0, sticky='w', pady=(0, 4))
        self._pwd_var = tk.StringVar()
        pf = tk.Frame(body, bg=C['border'], bd=1, relief='solid')
        pf.grid(row=3, column=0, sticky='ew', pady=(0, 6))
        self._pwd_ent = tk.Entry(pf,
            textvariable=self._pwd_var,
            font=('Consolas', 12), width=28, show='●',
            bg=C['bg_input'], fg=C['text'],
            insertbackground=C['accent'], relief='flat', bd=6)
        self._pwd_ent.pack(fill='x')

        body.columnconfigure(0, weight=1)

        # ── 顯示/隱藏密碼 ────────────────────────────────────────────
        self._show_pwd = tk.BooleanVar(value=False)
        tk.Checkbutton(self, text='顯示密碼',
                       variable=self._show_pwd,
                       command=self._toggle_show,
                       font=('Segoe UI', 9),
                       fg=C['text_dim'], bg=C['bg_panel'],
                       selectcolor=C['bg_card'],
                       activeforeground=C['text'],
                       activebackground=C['bg_panel'],
                       relief='flat', bd=0).pack(anchor='w', padx=20)

        # ── 記住密碼 ──────────────────────────────────────────────────
        self._remember = tk.BooleanVar(value=True)
        tk.Checkbutton(self, text='記住密碼（儲存到 Windows Credential Manager）',
                       variable=self._remember,
                       font=('Segoe UI', 9),
                       fg=C['text_dim'], bg=C['bg_panel'],
                       selectcolor=C['bg_card'],
                       activeforeground=C['text'],
                       activebackground=C['bg_panel'],
                       relief='flat', bd=0).pack(anchor='w', padx=20, pady=(2, 12))

        # ── 按鈕列 ────────────────────────────────────────────────────
        btn_row = tk.Frame(self, bg=C['bg_panel'])
        btn_row.pack(fill='x', padx=20, pady=(0, 18))
        tk.Button(btn_row, text='  確認登入  ',
                  font=('Segoe UI', 11, 'bold'),
                  bg=C['accent'], fg='white',
                  activebackground=C['accent_dim'],
                  activeforeground='white',
                  relief='flat', bd=0, pady=8,
                  cursor='hand2',
                  command=self._ok).pack(side='right')
        tk.Button(btn_row, text='取消',
                  font=('Segoe UI', 10),
                  bg=C['bg_card'], fg=C['text_dim'],
                  activebackground=C['border'],
                  relief='flat', bd=0, pady=8, padx=14,
                  cursor='hand2',
                  command=self._cancel).pack(side='right', padx=(0, 8))

        # ── 按下 Enter 確認 ───────────────────────────────────────────
        self.bind('<Return>', lambda _: self._ok())
        self.bind('<Escape>', lambda _: self._cancel())

        # ── 置中於父視窗 ──────────────────────────────────────────────
        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")

        # 預設游標到密碼欄（若帳號已填）
        if prefill_user:
            self._pwd_ent.focus_set()
        else:
            self._user_ent.focus_set()

        self.wait_window()

    def _toggle_show(self):
        self._pwd_ent.config(show='' if self._show_pwd.get() else '●')

    def _ok(self):
        u = self._user_var.get().strip()
        p = self._pwd_var.get()
        if not u:
            messagebox.showwarning('欄位不可為空', '請輸入帳號。', parent=self)
            self._user_ent.focus_set()
            return
        if not p:
            messagebox.showwarning('欄位不可為空', '請輸入密碼。', parent=self)
            self._pwd_ent.focus_set()
            return
        self.result = (u, p, self._remember.get())
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ══════════════════════════════════════════════════════════════════════
#  詢問是否儲存新密碼對話框
# ══════════════════════════════════════════════════════════════════════
class AskSaveDialog(tk.Toplevel):
    """result = True（儲存）/ False（不儲存）"""
    def __init__(self, parent, system_label: str):
        super().__init__(parent)
        self.result = False
        self.title("偵測到密碼變更")
        self.configure(bg=C['bg_panel'])
        self.resizable(False, False)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._no)

        tk.Label(self, text='🔔  密碼已更新',
                 font=('Segoe UI', 13, 'bold'),
                 fg=C['warning'], bg=C['bg_panel']).pack(padx=24, pady=(20, 6))
        tk.Label(self,
                 text=f'偵測到 {system_label} 使用了與上次不同的密碼。\n\n是否將新密碼儲存到 Windows Credential Manager？',
                 font=('Segoe UI', 10),
                 fg=C['text'], bg=C['bg_panel'],
                 wraplength=300, justify='center').pack(padx=24, pady=(0, 16))

        btn_row = tk.Frame(self, bg=C['bg_panel'])
        btn_row.pack(pady=(0, 20))
        tk.Button(btn_row, text='  儲存新密碼  ',
                  font=('Segoe UI', 11, 'bold'),
                  bg=C['btn_save'], fg='#56D364',
                  activebackground='#0D3A1E',
                  relief='flat', bd=0, pady=8,
                  cursor='hand2',
                  command=self._yes).pack(side='left', padx=(0, 8))
        tk.Button(btn_row, text='不儲存',
                  font=('Segoe UI', 10),
                  bg=C['bg_card'], fg=C['text_dim'],
                  activebackground=C['border'],
                  relief='flat', bd=0, pady=8, padx=14,
                  cursor='hand2',
                  command=self._no).pack(side='left')

        self.bind('<Return>', lambda _: self._yes())
        self.bind('<Escape>', lambda _: self._no())

        self.update_idletasks()
        px = parent.winfo_x() + (parent.winfo_width()  - self.winfo_width())  // 2
        py = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{px}+{py}")
        self.wait_window()

    def _yes(self): self.result = True;  self.destroy()
    def _no(self):  self.result = False; self.destroy()


# ══════════════════════════════════════════════════════════════════════
#  主應用程式
# ══════════════════════════════════════════════════════════════════════
class App:

    def __init__(self, root: tk.Tk):
        self.root    = root
        self.root.title("MOXA Repair Data Finder  v1.0")
        self.root.geometry("1280x840")
        self.root.configure(bg=C['bg_dark'])
        self.root.minsize(960, 640)

        self.q:       queue.Queue = queue.Queue()
        self.running: bool        = False
        self.pdm_drv  = None
        self.dmp_drv  = None
        self._iid_map: dict       = {}

        self._setup_styles()
        self._build_ui()
        self._poll()

    # ── Styles ────────────────────────────────────────────────────────
    def _setup_styles(self):
        s = ttk.Style()
        s.theme_use('clam')
        s.configure('Bar.Horizontal.TProgressbar',
                    troughcolor=C['bg_card'], background=C['accent'],
                    lightcolor=C['accent'], darkcolor=C['accent_dim'],
                    thickness=10)
        s.configure('T.Treeview',
                    background=C['bg_panel'], foreground=C['text'],
                    fieldbackground=C['bg_panel'], rowheight=36,
                    borderwidth=0, font=('Consolas', 11))
        s.configure('T.Treeview.Heading',
                    background=C['bg_card'], foreground=C['accent'],
                    relief='flat', font=('Segoe UI', 11, 'bold'))
        s.map('T.Treeview',
              background=[('selected', C['row_sel'])],
              foreground=[('selected', 'white')])

    # ══════════════════════════════════════════════════════════════════
    #  UI 建構
    # ══════════════════════════════════════════════════════════════════
    def _build_ui(self):
        self._topbar()
        body = tk.Frame(self.root, bg=C['bg_dark'])
        body.pack(fill='both', expand=True)
        self._left(body)
        self._right(body)

    # ── 頂部 ──────────────────────────────────────────────────────────
    def _topbar(self):
        bar = tk.Frame(self.root, bg='#0A0D13', height=50)
        bar.pack(fill='x')
        bar.pack_propagate(False)

        tk.Label(bar, text='MOXA', font=('Segoe UI', 14, 'bold'),
                 fg=C['accent'], bg='#0A0D13').pack(side='left', padx=(18, 4), pady=8)
        tk.Label(bar, text='Repair Data Finder v1.0',
                 font=('Segoe UI', 11), fg=C['text'],
                 bg='#0A0D13').pack(side='left')

        right = tk.Frame(bar, bg='#0A0D13')
        right.pack(side='right', padx=20)
        self._pdm_dot = self._dot(right, 'PDM (Windchill)')
        tk.Frame(right, bg='#0A0D13', width=20).pack(side='left')
        self._dmp_dot = self._dot(right, 'GeDCC (DMP)')

    def _dot(self, p, label):
        d = tk.Label(p, text='●', font=('Segoe UI', 12),
                     fg=C['error'], bg='#0A0D13')
        d.pack(side='left')
        tk.Label(p, text=f'  {label}', font=('Segoe UI', 8),
                 fg=C['text_dim'], bg='#0A0D13').pack(side='left')
        return d

    # ── 左側控制區 ────────────────────────────────────────────────────
    def _left(self, parent):
        pnl = tk.Frame(parent, bg=C['bg_panel'], width=260)
        pnl.pack(side='left', fill='y', padx=(10, 6), pady=10)
        pnl.pack_propagate(False)

        self._sec(pnl, '料號輸入', top=18)
        ew = tk.Frame(pnl, bg=C['border'], bd=1, relief='solid')
        ew.pack(fill='x', padx=14, pady=(3, 14))
        self._ent = tk.Entry(ew, font=('Consolas', 12),
                             bg=C['bg_input'], fg=C['text_dim'],
                             insertbackground=C['accent'],
                             relief='flat', bd=8)
        self._ent.pack(fill='x')
        self._ent.insert(0, '例：9020240602061')
        self._ent.bind('<FocusIn>',  lambda _: self._ent_in())
        self._ent.bind('<FocusOut>', lambda _: self._ent_out())
        self._ent.bind('<Return>', lambda _: self._start())

        defs = [
            ('▶   開始檢索', C['accent'],  'normal',   self._start, '_bs'),
            ('⏹   停  止',  '#3A4155',    'disabled', self._stop,  '_bx'),
            ('🗑   清除結果', C['bg_card'], 'normal',   self._clear, '_bc'),
        ]
        for text, bg, st, cmd, attr in defs:
            b = tk.Button(pnl, text=text,
                          font=('Segoe UI', 11,
                                'bold' if attr == '_bs' else 'normal'),
                          bg=bg, fg='white',
                          activebackground=C['accent_dim'],
                          activeforeground='white',
                          relief='flat', bd=0, pady=10,
                          cursor='hand2', state=st, command=cmd)
            b.pack(fill='x', padx=14, pady=(0, 6))
            setattr(self, attr, b)

        self._div(pnl)
        self._sec(pnl, '統計摘要')
        self._sv_total   = self._stat(pnl, '找到 PCB 料號', '0')
        self._sv_done    = self._stat(pnl, '✅ 已完成',      '0')
        self._sv_pending = self._stat(pnl, '⏳ 處理中',      '0')
        self._sv_err     = self._stat(pnl, '⚠ 無結果',      '0')

        self._div(pnl)
        self._sec(pnl, '帳密管理')
        tk.Button(pnl, text='🔑  管理 PDM 密碼',
                  font=('Segoe UI', 9), bg=C['bg_card'], fg=C['text_dim'],
                  activebackground=C['border'],
                  relief='flat', bd=0, pady=7,
                  cursor='hand2',
                  command=lambda: self._manage_cred('pdm')).pack(
                      fill='x', padx=14, pady=(0, 4))

        self._div(pnl)
        self._sec(pnl, '資料匯出')
        tk.Button(pnl, text='📤  匯出 CSV',
                  font=('Segoe UI', 10), bg=C['bg_card'], fg=C['text'],
                  activebackground=C['border'],
                  relief='flat', bd=0, pady=9,
                  cursor='hand2', command=self._export).pack(
                      fill='x', padx=14, pady=(0, 6))

    def _sec(self, p, t, top=10):
        tk.Label(p, text=t, font=('Segoe UI', 11, 'bold'),
                 fg=C['accent'], bg=C['bg_panel']).pack(
                     anchor='w', padx=14, pady=(top, 4))

    def _div(self, p):
        tk.Frame(p, bg=C['border'], height=1).pack(fill='x', padx=14, pady=12)

    def _stat(self, p, label, val):
        c = tk.Frame(p, bg=C['bg_card'], pady=7, padx=12)
        c.pack(fill='x', padx=14, pady=2)
        tk.Label(c, text=label, font=('Segoe UI', 9),
                 fg=C['text_dim'], bg=C['bg_card']).pack(anchor='w')
        lv = tk.Label(c, text=val, font=('Consolas', 20, 'bold'),
                      fg=C['accent'], bg=C['bg_card'])
        lv.pack(anchor='w')
        return lv

    # ── 右側工作區 ────────────────────────────────────────────────────
    def _right(self, parent):
        r = tk.Frame(parent, bg=C['bg_dark'])
        r.pack(side='left', fill='both', expand=True, padx=(0, 10), pady=10)
        self._progress_card(r)
        self._log_card(r)
        self._table_card(r)

    def _progress_card(self, p):
        c = tk.Frame(p, bg=C['bg_panel'], pady=12, padx=16)
        c.pack(fill='x', pady=(0, 8))
        top = tk.Frame(c, bg=C['bg_panel'])
        top.pack(fill='x')
        tk.Label(top, text='執行進度', font=('Segoe UI', 11, 'bold'),
                 fg=C['accent'], bg=C['bg_panel']).pack(side='left')
        self._pct = tk.Label(top, text='0 %',
                              font=('Consolas', 11, 'bold'),
                              fg=C['accent'], bg=C['bg_panel'])
        self._pct.pack(side='right')
        self._bar = ttk.Progressbar(c, style='Bar.Horizontal.TProgressbar',
                                     mode='determinate')
        self._bar.pack(fill='x', pady=(6, 5))
        self._stlbl = tk.Label(c,
                                text='就緒 — 請輸入成品料號後點擊「開始檢索」',
                                font=('Segoe UI', 9), fg=C['text_dim'],
                                bg=C['bg_panel'], anchor='w')
        self._stlbl.pack(fill='x')

    def _log_card(self, p):
        tk.Label(p, text='執行日誌 (Real-time Log)',
                 font=('Segoe UI', 11, 'bold'),
                 fg=C['accent'], bg=C['bg_dark']).pack(anchor='w', pady=(0, 3))
        self._log = scrolledtext.ScrolledText(
            p, font=('Consolas', 10),
            bg=C['log_bg'], fg=C['log_fg'],
            insertbackground='lime',
            relief='flat', bd=4, wrap='word', height=9,
            state='disabled')
        self._log.pack(fill='both')
        for tag, col in [
            ('INFO', '#7EE787'), ('STEP', '#79C0FF'), ('WARN', '#E3A740'),
            ('ERROR', '#F85149'), ('OK', '#56D364'), ('TS', '#484F58'),
        ]:
            self._log.tag_configure(tag, foreground=col)

    def _table_card(self, p):
        hdr = tk.Frame(p, bg=C['bg_dark'])
        hdr.pack(fill='x', pady=(8, 3))
        tk.Label(hdr, text='檢索結果',
                 font=('Segoe UI', 11, 'bold'),
                 fg=C['accent'], bg=C['bg_dark']).pack(side='left')
        tk.Label(hdr, text='  雙擊列可開啟 GeDCC 查詢',
                 font=('Segoe UI', 10), fg=C['text_dim'],
                 bg=C['bg_dark']).pack(side='left')

        self._fg_banner = tk.Label(p, text='',
            font=('Segoe UI', 11, 'bold'),
            fg=C['warning'], bg=C['bg_dark'], anchor='w')
        self._fg_banner.pack(fill='x', pady=(0, 3))

        wrap = tk.Frame(p, bg=C['bg_panel'])
        wrap.pack(fill='both', expand=True)

        cols = ('status', 'pcb')
        self._tv = ttk.Treeview(wrap, columns=cols,
                                 show='headings', style='T.Treeview')
        for col, hd, w, anc in [
            ('status', '狀態',                    90, 'center'),
            ('pcb',    '1199 PCB 料號/名稱（雙擊開啟 GeDCC 查詢）', 1040, 'w'),
        ]:
            self._tv.heading(col, text=hd)
            self._tv.column(col, width=w, anchor=anc, minwidth=50)

        sb = ttk.Scrollbar(wrap, orient='vertical', command=self._tv.yview)
        self._tv.configure(yscrollcommand=sb.set)
        self._tv.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        self._tv.tag_configure('ok',   background='#122412', foreground='#56D364')
        self._tv.tag_configure('wait', background='#1A1A0A', foreground='#E3A740')
        self._tv.tag_configure('na',   background='#1E1010', foreground='#F85149')
        self._tv.bind('<Double-Button-1>', self._dbl)
        self._tv.bind('<Motion>', self._on_hover)

    # ══════════════════════════════════════════════════════════════════
    #  帳密管理（主執行緒直接呼叫）
    # ══════════════════════════════════════════════════════════════════
    def _manage_cred(self, system: str):
        """手動管理帳密按鈕 — 讓使用者更新或刪除儲存的密碼"""
        label = 'PDM (Windchill)' if system == 'pdm' else 'GeDCC (DMP)'
        saved_user, _ = CredentialManager.load(system)
        dlg = CredentialDialog(self.root, label, prefill_user=saved_user,
                               message='輸入新帳密後儲存，或按取消保留現有設定。')
        if dlg.result:
            user, pwd, remember = dlg.result
            if remember:
                CredentialManager.save(system, user, pwd)
                messagebox.showinfo('儲存成功',
                    f'{label} 帳密已更新並儲存到 Windows Credential Manager。')
            else:
                messagebox.showinfo('未儲存', '帳密未儲存（可再次點擊「管理帳密」更新）。')

    # ══════════════════════════════════════════════════════════════════
    #  Thread → 主執行緒對話框：askpass / asksave
    # ══════════════════════════════════════════════════════════════════
    def _ask_credentials(self, system: str,
                         prefill_user: str = "", message: str = "") -> tuple | None:
        """
        從背景執行緒安全呼叫 CredentialDialog。
        阻塞直到使用者確認或取消（最長等候 5 分鐘）。
        回傳 (user, pwd, remember) 或 None。
        """
        evt = threading.Event()
        holder = []
        self._q(type='askpass', system=system, prefill_user=prefill_user,
                message=message, event=evt, holder=holder)
        evt.wait(timeout=300)
        return holder[0] if holder else None

    def _ask_save_password(self, system: str) -> bool:
        """從背景執行緒詢問是否儲存新密碼。阻塞直到使用者回答。"""
        evt = threading.Event()
        holder = []
        self._q(type='asksave', system=system, event=evt, holder=holder)
        evt.wait(timeout=60)
        return bool(holder[0]) if holder else False

    # ══════════════════════════════════════════════════════════════════
    #  UI 事件
    # ══════════════════════════════════════════════════════════════════
    def _ent_in(self):
        if self._ent.get() == '例：9020240602061':
            self._ent.delete(0, 'end')
            self._ent.config(fg=C['text'])

    def _ent_out(self):
        if not self._ent.get().strip():
            self._ent.insert(0, '例：9020240602061')
            self._ent.config(fg=C['text_dim'])

    def _on_hover(self, event):
        row = self._tv.identify_row(event.y)
        self._tv.config(cursor='hand2' if row else '')

    def _dbl(self, event):
        sel = self._tv.selection()
        if not sel: return
        vals = self._tv.item(sel[0])['values']
        if not vals or len(vals) < 2: return
        m = re.search(r'(1199\d{6,})', str(vals[1]))
        if not m: return
        threading.Thread(target=self._open_gedcc,
                         args=(m.group(1),), daemon=True).start()

    def _start(self):
        pn = self._ent.get().strip()
        if not pn or pn == '例：9020240602061':
            messagebox.showwarning('輸入錯誤', '請輸入有效的成品料號')
            return
        self.running = True
        self._bs.config(state='disabled')
        self._bx.config(state='normal')
        self._ent.config(state='disabled')
        threading.Thread(target=self._run, args=(pn,), daemon=True).start()

    def _stop(self):
        self.running = False
        self._log_w('⏹ 使用者中止', 'WARN')

    def _clear(self):
        for i in self._tv.get_children(): self._tv.delete(i)
        self._iid_map.clear()
        self._fg_banner.config(text='')
        self._log.config(state='normal')
        self._log.delete('1.0', 'end')
        self._log.config(state='disabled')
        self._bar['value'] = 0
        self._pct.config(text='0 %')
        self._stlbl.config(text='就緒 — 請輸入成品料號後點擊「開始檢索」')
        self._upd_stats()

    def _export(self):
        items = self._tv.get_children()
        if not items:
            messagebox.showinfo('無資料', '目前沒有可匯出的結果')
            return
        fname = f"repair_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        path  = os.path.join(os.path.expanduser('~'), 'Desktop', fname)
        with open(path, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(['狀態', '1199 PCB 料號/名稱'])
            for i in items:
                v = self._tv.item(i)['values']
                w.writerow([v[0], v[1]])
        messagebox.showinfo('匯出成功', f'已存至桌面：\n{fname}')

    def _on_done(self):
        self.running = False
        self._bs.config(state='normal')
        self._bx.config(state='disabled')
        self._ent.config(state='normal')

    def _upd_stats(self):
        items = self._tv.get_children()
        total = len(items)
        done  = sum(1 for i in items if 'ok'   in self._tv.item(i)['tags'])
        wait  = sum(1 for i in items if 'wait' in self._tv.item(i)['tags'])
        err   = sum(1 for i in items if 'na'   in self._tv.item(i)['tags'])
        self._sv_total.config(text=str(total))
        self._sv_done.config(text=str(done))
        self._sv_pending.config(text=str(wait))
        self._sv_err.config(text=str(err))

    # ── 訊息佇列 ─────────────────────────────────────────────────────
    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                t = m.get('type')

                if t == 'log':
                    self._log_w(m['text'], m.get('lv', 'INFO'))

                elif t == 'prog':
                    v = int(m['v'])
                    self._bar['value'] = v
                    self._pct.config(text=f'{v} %')
                    if m.get('s'): self._stlbl.config(text=m['s'])

                elif t == 'dot':
                    d = self._pdm_dot if m['sys'] == 'pdm' else self._dmp_dot
                    d.config(fg=C['success'] if m['ok'] else C['error'])

                elif t == 'add':
                    iid = self._tv.insert('', 'end',
                        values=(m['st'], m.get('pcbtxt', m['pcb'])),
                        tags=(m.get('tag', 'ok'),))
                    self._iid_map[m['pcb']] = iid
                    self._upd_stats()

                elif t == 'banner':
                    self._fg_banner.config(text=m.get('text', ''))

                elif t == 'popup':
                    messagebox.showinfo('需要手動登入',
                        f'請在彈出的 {m["sys"]} 瀏覽器視窗中完成帳密輸入。\n\n登入後程式將自動繼續。')

                elif t == 'done':
                    self._on_done()

                # ── 新：背景執行緒請求帳密對話框 ──────────────────────
                elif t == 'askpass':
                    label = 'PDM (Windchill)' if m['system'] == 'pdm' else 'GeDCC (DMP)'
                    dlg = CredentialDialog(self.root, label,
                                          prefill_user=m.get('prefill_user', ''),
                                          message=m.get('message', ''))
                    m['holder'].append(dlg.result)
                    m['event'].set()

                # ── 新：背景執行緒請求詢問是否儲存新密碼 ─────────────
                elif t == 'asksave':
                    label = 'PDM (Windchill)' if m['system'] == 'pdm' else 'GeDCC (DMP)'
                    dlg = AskSaveDialog(self.root, label)
                    m['holder'].append(dlg.result)
                    m['event'].set()

        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    def _q(self, **kw):       self.q.put(kw)
    def _log_w(self, t, lv='INFO'):
        self._log.config(state='normal')
        ts = datetime.now().strftime('%H:%M:%S')
        self._log.insert('end', f'[{ts}] ', 'TS')
        self._log.insert('end', f'{t}\n', lv)
        self._log.see('end')
        self._log.config(state='disabled')
    def _ql(self, t, lv='INFO'): self._q(type='log', text=t, lv=lv)
    def _qp(self, v, s=''):      self._q(type='prog', v=v, s=s)

    # ══════════════════════════════════════════════════════════════════
    #  自動化主流程（背景執行緒）
    # ══════════════════════════════════════════════════════════════════
    def _run(self, fg_pn: str):
        try:
            self._ql(f'🚀 開始檢索：{fg_pn}', 'STEP')
            self._qp(5, '正在初始化瀏覽器...')
            self.pdm_drv = self._init_browser('pdm')
            if not self.pdm_drv:
                self._ql('❌ 無法啟動 Edge 或 Chrome 瀏覽器', 'ERROR')
                self._q(type='done'); return

            self._qp(10, '正在開啟 PDM 登入頁面...')
            self.pdm_drv.get(PDM_URL)

            # ── PDM 自動登入 ──────────────────────────────────────────
            if not self._pdm_login():
                self._ql('❌ PDM 登入失敗', 'ERROR')
                self._q(type='done'); return

            self._q(type='dot', sys='pdm', ok=True)
            self._ql('✅ PDM 登入成功', 'OK')
            self._qp(20, f'正在 PDM 搜尋 {fg_pn}...')

            pcb_list = self._pdm_find_1199(fg_pn)
            if not pcb_list:
                self._ql('❌ BOM 中未找到任何 1199 PCB 料號', 'ERROR')
                self._q(type='done'); return

            self._ql(f'✅ 共找到 {len(pcb_list)} 個 1199 PCB 料號', 'OK')
            self._qp(55, f'共 {len(pcb_list)} 個 PCB 料號')

            fg_name = self._get_fg_name(fg_pn)
            self._q(type='banner',
                    text=f'🔎  {fg_pn}  {fg_name}'.strip())

            if not self.running:
                self._q(type='done'); return

            for p in pcb_list:
                pcb_txt = p.get('text') or p['num']
                pcb_txt = re.sub(r'(版本|名稱)\s*[:：]?\s*', ' ', pcb_txt).split('狀態')[0].strip()
                self._q(type='add', st='✅ 找到', pcb=p['num'],
                        pcbtxt=pcb_txt, tag='ok')

            self._qp(100, '✅ 檢索完成！')
            self._ql('🎉 掃描完畢！雙擊任一列可開啟 GeDCC 查詢。', 'OK')

        except Exception as exc:
            self._ql(f'❌ 執行例外：{exc}', 'ERROR')
        finally:
            self._q(type='done')

    # ══════════════════════════════════════════════════════════════════
    #  PDM 登入
    # ══════════════════════════════════════════════════════════════════
    def _pdm_login(self) -> bool:
        """
        PDM 登入主流程（HTTP Basic Auth 版本）：

        PDM 使用瀏覽器原生 HTTP Basic Auth 彈窗（非 HTML 表單），
        解法：將帳密嵌入 URL（https://user:pwd@host/path）直接繞過彈窗。

        情境 A — 第一次：對話框輸入 → 儲存 → URL 嵌入帳密自動登入
        情境 B — 之後每次：直接 URL 嵌入帳密，靜默登入
        情境 C — 帳密更換：登入失敗 → 對話框重新輸入 → 儲存新帳密
        """
        from urllib.parse import quote, urlparse

        time.sleep(1.5)

        # ── 已有 session，直接通過 ───────────────────────────────────
        if self._pdm_is_home():
            self._ql('✅ PDM 已登入（Session 尚有效）', 'OK')
            return True

        # ── 取得帳密 ─────────────────────────────────────────────────
        saved_user, saved_pwd = CredentialManager.load('pdm')

        if not saved_user:
            # 情境 A：第一次，詢問帳密
            self._ql('ℹ️ 第一次使用，請在對話框輸入 PDM 帳號密碼', 'INFO')
            cred = self._ask_credentials(
                'pdm',
                message='請輸入 PDM 帳號密碼。\n儲存後，下次將自動登入，無須再輸入。')
            if not cred:
                self._ql('❌ 未輸入帳密，無法自動登入', 'ERROR')
                return False
            saved_user, saved_pwd, remember = cred
            if remember:
                CredentialManager.save('pdm', saved_user, saved_pwd)
                self._ql('✅ PDM 帳密已儲存', 'OK')

        # ── 情境 B：URL 嵌入帳密，繞過 HTTP Basic Auth 彈窗 ────────
        self._ql('🔐 正在自動登入 PDM...', 'INFO')
        login_ok = self._pdm_auth_via_url(saved_user, saved_pwd)

        if login_ok:
            self._ql('✅ PDM 自動登入成功', 'OK')
            return True

        # ── 情境 C：帳密錯誤，重新輸入 ─────────────────────────────
        self._ql('⚠ PDM 自動登入失敗，帳密可能已更換', 'WARN')
        self._ql('👉 請在對話框輸入新帳號密碼', 'WARN')
        cred = self._ask_credentials(
            'pdm',
            prefill_user=saved_user,
            message='⚠ 登入失敗，帳密可能已更換。\n請重新輸入正確的帳號密碼。')
        if not cred:
            return False

        new_user, new_pwd, _ = cred
        login_ok = self._pdm_auth_via_url(new_user, new_pwd)
        if not login_ok:
            return False

        # 新帳密與舊的不同 → 詢問是否儲存
        if new_user != saved_user or new_pwd != saved_pwd:
            self._ql('🔔 偵測到 PDM 帳密與上次不同', 'WARN')
            if self._ask_save_password('pdm'):
                CredentialManager.save('pdm', new_user, new_pwd)
                self._ql('✅ PDM 新帳密已儲存', 'OK')
        return True

    def _pdm_auth_via_url(self, user: str, pwd: str) -> bool:
        """
        將帳密嵌入 URL，讓瀏覽器直接通過 HTTP Basic Auth，
        不會出現「登錄以存取此網站」彈窗。
        例：https://dana_lee:mypass@pap.moxa.com/Windchill/app/
        """
        from urllib.parse import quote, urlparse
        drv = self.pdm_drv
        try:
            parsed   = urlparse(PDM_URL)
            enc_user = quote(user, safe='')
            enc_pwd  = quote(pwd,  safe='')
            auth_url = f"{parsed.scheme}://{enc_user}:{enc_pwd}@{parsed.netloc}{parsed.path}"
            drv.get(auth_url)
            # 等待主頁面載入（最多 30 秒）
            return self._wait_pdm_login(timeout=30)
        except Exception as e:
            self._ql(f'URL 登入例外：{str(e)[:60]}', 'WARN')
            return False

    def _pdm_is_home(self) -> bool:
        try:
            WebDriverWait(self.pdm_drv, 4).until(
                EC.presence_of_element_located((By.ID, 'gloabalSearchField')))
            return True
        except Exception:
            return False

    def _pdm_wait_for_form(self, timeout: int = 15) -> bool:
        """
        等待 PDM 登入表單出現（密碼欄位）。
        涵蓋 Windchill 原生、ADFS、Windows Auth 等多種情境。
        回傳 True 表示找到密碼欄位。
        """
        drv = self.pdm_drv
        self._ql('⏳ 等待 PDM 登入頁面...', 'INFO')
        deadline = time.time() + timeout
        while time.time() < deadline:
            # 若已到主頁就不需要再找表單
            if self._pdm_is_home():
                return False  # 已登入，不需填入
            try:
                drv.find_element(By.XPATH, '//input[@type="password"]')
                self._ql('✅ 偵測到 PDM 登入表單', 'INFO')
                return True
            except Exception:
                pass
            time.sleep(0.8)
        return False

    def _pdm_fill_form(self, user: str, pwd: str) -> bool:
        """
        填入 PDM 登入表單帳號密碼並送出。
        支援 Windchill (j_username)、ADFS (UserName)、一般 text 欄位。
        """
        drv = self.pdm_drv
        # 帳號欄位（優先順序：Windchill → ADFS → 通用）
        user_xpaths = [
            '//input[@name="j_username"]',
            '//input[@id="j_username"]',
            '//input[@name="UserName"]',       # ADFS
            '//input[@id="userNameInput"]',    # ADFS
            '//input[@autocomplete="username"]',
            '//input[@type="text" and contains(@name,"user")]',
            '//input[@type="text" and contains(@id,"user")]',
            '//input[@type="text"][1]',
        ]
        # 密碼欄位
        pwd_xpaths = [
            '//input[@name="j_password"]',
            '//input[@id="j_password"]',
            '//input[@name="Password"]',       # ADFS
            '//input[@id="passwordInput"]',    # ADFS
            '//input[@type="password"]',
        ]

        user_el, pwd_el = None, None
        for xp in user_xpaths:
            try:
                user_el = drv.find_element(By.XPATH, xp)
                break
            except Exception:
                continue
        for xp in pwd_xpaths:
            try:
                pwd_el = drv.find_element(By.XPATH, xp)
                break
            except Exception:
                continue

        if not user_el or not pwd_el:
            self._ql('⚠ 找不到帳號或密碼欄位', 'WARN')
            return False

        try:
            self._ql('🔐 正在自動填入 PDM 帳密...', 'INFO')
            drv.execute_script("arguments[0].value='';", user_el)
            user_el.click(); user_el.send_keys(user)
            drv.execute_script("arguments[0].value='';", pwd_el)
            pwd_el.click(); pwd_el.send_keys(pwd)
            pwd_el.send_keys(Keys.RETURN)
            time.sleep(2)
            return True
        except Exception as e:
            self._ql(f'填入表單時發生錯誤：{str(e)[:60]}', 'WARN')
            return False

    # ══════════════════════════════════════════════════════════════════
    #  WNJPHandler 對話框：自動點擊「開啟」
    # ══════════════════════════════════════════════════════════════════
    def _auto_click_wnjp_dialog(self, drv=None):
        """
        自動點擊瀏覽器的「開啟 WNJPHandler」對話框。
        Chrome/Edge 的此對話框屬於瀏覽器內部 UI，無法用 UI Automation 存取，
        改用 pyautogui 依瀏覽器視窗位置計算按鈕座標並點擊。

        對話框出現在瀏覽器內容區域頂部中央，「開啟」按鈕在右側。
        最多等候 8 秒。
        """
        try:
            import pyautogui
        except ImportError:
            self._ql('⚠ 未安裝 pyautogui，無法自動點擊「開啟」', 'WARN')
            return

        target_drv = drv or self.dmp_drv
        try:
            # 取得瀏覽器視窗的位置與大小
            rect = target_drv.get_window_rect()
            win_x = rect['x']
            win_y = rect['y']
            win_w = rect['width']
            win_h = rect['height']
        except Exception:
            # 若無法取得視窗資訊，使用螢幕中央估算
            sw, sh = pyautogui.size()
            win_x, win_y, win_w, win_h = sw // 4, sh // 4, sw // 2, sh // 2

        # 對話框出現在視窗上半部，「開啟」按鈕在對話框右側
        # 依截圖比例估算座標（視窗寬度中央偏右，垂直約 35%）
        click_x = win_x + int(win_w * 0.62)
        click_y = win_y + int(win_h * 0.35)

        self._ql('⏳ 等待 WNJPHandler 對話框...', 'INFO')
        for attempt in range(16):   # 最多等 8 秒（每 0.5 秒試一次）
            time.sleep(0.5)
            try:
                # 截圖確認對話框是否出現（在估算區域附近找「開啟」文字的按鈕色塊）
                screenshot = pyautogui.screenshot(
                    region=(win_x + int(win_w * 0.3),
                            win_y + int(win_h * 0.2),
                            int(win_w * 0.5),
                            int(win_h * 0.3)))
                # 檢查該區域是否有對話框特徵色（白色背景區塊）
                pixels = screenshot.getdata()
                white_count = sum(1 for r, g, b in pixels if r > 240 and g > 240 and b > 240)
                if white_count > len(pixels) * 0.3:  # 超過 30% 為白色 → 對話框存在
                    pyautogui.click(click_x, click_y)
                    self._ql('✅ 已自動點擊「開啟」按鈕', 'OK')
                    self._q(type='dot', sys='dmp', ok=True)
                    return
            except Exception:
                pass

        self._ql('⚠ 未偵測到對話框，請手動點擊「開啟」', 'WARN')

    # ══════════════════════════════════════════════════════════════════
    #  GeDCC 開啟（雙擊觸發）
    # ══════════════════════════════════════════════════════════════════
    def _open_gedcc(self, pcb_num: str):
        try:
            # ── 確認瀏覽器存活 ───────────────────────────────────────
            need_new = False
            if self.dmp_drv is None:
                need_new = True
            else:
                try:
                    _ = self.dmp_drv.current_url
                except Exception:
                    need_new = True

            if need_new:
                self.dmp_drv = self._init_browser('dmp')
                if not self.dmp_drv:
                    return
                self._q(type='dot', sys='dmp', ok=False)

            drv = self.dmp_drv

            # ── 開啟 GeDCC，並自動處理 WNJPHandler 對話框 ──────────
            if 'DMP/private' not in drv.current_url:
                drv.get(GEDCC_URL)
                self._ql('🌐 正在開啟 GeDCC...', 'INFO')
                # 在背景同步等待並自動點擊「開啟」對話框
                threading.Thread(
                    target=self._auto_click_wnjp_dialog, daemon=True).start()
                try:
                    WebDriverWait(drv, 30).until(
                        lambda d: 'DMP/private' in d.current_url)
                    self._ql('✅ GeDCC 已就緒', 'OK')
                    self._q(type='dot', sys='dmp', ok=True)
                except TimeoutException:
                    # WNJPHandler 模式：不跳轉到 DMP/private，直接繼續
                    self._ql('ℹ️ GeDCC 以 WNJPHandler 模式開啟', 'INFO')

            # ── 嘗試在網頁搜尋框填入料號 ────────────────────────────
            if 'DMP/private' in drv.current_url:
                if 'user_index' not in drv.current_url:
                    drv.get(GEDCC_HOME)
                    time.sleep(1.5)
                try:
                    sel_el = WebDriverWait(drv, 5).until(
                        EC.presence_of_element_located((By.XPATH,
                        '//select[option[normalize-space()="料號"]]'
                        '|//select[contains(@name,"type") or contains(@id,"type")]')))
                    Select(sel_el).select_by_visible_text('料號')
                    time.sleep(0.3)
                except Exception:
                    pass
                for xp in ['//input[@name="bqKeyword"]', '//input[@id="bqKeyword"]',
                            '//form//input[@type="text"][1]']:
                    try:
                        inp = WebDriverWait(drv, 3).until(
                            EC.element_to_be_clickable((By.XPATH, xp)))
                        drv.execute_script("arguments[0].value = '';", inp)
                        inp.click()
                        inp.send_keys(pcb_num)
                        inp.send_keys(Keys.RETURN)
                        self._ql(f'✅ GeDCC 查詢已送出：{pcb_num}', 'OK')
                        return
                    except Exception:
                        continue
                self._ql('❌ 找不到 GeDCC 搜尋框', 'ERROR')

        except Exception as e:
            self._ql(f'GeDCC 開啟錯誤：{str(e)[:60]}', 'ERROR')

    # ══════════════════════════════════════════════════════════════════
    #  瀏覽器初始化
    # ══════════════════════════════════════════════════════════════════
    def _init_browser(self, label='pdm'):
        import subprocess as _sp

        # GeDCC 瀏覽器：使用 persistent profile 記住 WNJPHandler 許可
        # 第一次手動點「開啟」後，之後永遠不再詢問
        prefs = {
            'protocol_handler.excluded_schemes': {'wnjp': False},
        }
        extra_args = []
        if label == 'dmp':
            profile_dir = os.path.join(
                os.path.expanduser('~'), '.rdf_profiles', 'gedcc')
            os.makedirs(profile_dir, exist_ok=True)
            extra_args.append(f'--user-data-dir={profile_dir}')
            prefs['protocol_handler.allowed_origin_protocol_pairs'] = {
                'http://global-gedcc.moxa.com': {'wnjp': True}
            }

        try:
            o = EdgeOptions()
            for a in ['--ignore-certificate-errors', '--ignore-ssl-errors',
                      '--log-level=3', '--silent'] + extra_args:
                o.add_argument(a)
            o.add_experimental_option('excludeSwitches',
                                       ['enable-logging', 'enable-automation'])
            o.add_experimental_option('prefs', prefs)
            drv = webdriver.Edge(
                service=EdgeService(log_output=_sp.DEVNULL), options=o)
            self._ql(f'✅ 已成功連結 Edge 瀏覽器 ({label})', 'OK')
            return drv
        except Exception as ee:
            self._ql(f'ℹ️ Edge 啟動未成 ({ee})，嘗試 Chrome...', 'INFO')
        try:
            o = ChromeOptions()
            for a in ['--ignore-certificate-errors', '--ignore-ssl-errors',
                      '--log-level=3', '--silent'] + extra_args:
                o.add_argument(a)
            o.add_experimental_option('excludeSwitches',
                                       ['enable-logging', 'enable-automation'])
            o.add_experimental_option('prefs', prefs)
            drv = webdriver.Chrome(
                service=ChromeService(log_output=_sp.DEVNULL), options=o)
            self._ql(f'✅ 已成功連結 Chrome 瀏覽器 ({label})', 'OK')
            return drv
        except Exception as ce:
            self._ql(f'❌ Chrome 啟動也失敗 ({ce})', 'ERROR')
            return None

    def _wait_pdm_login(self, timeout=180) -> bool:
        try:
            WebDriverWait(self.pdm_drv, timeout).until(
                EC.presence_of_element_located((By.ID, 'gloabalSearchField')))
            self._ql('✅ PDM 登入完成（偵測到搜尋框）', 'OK')
            time.sleep(1)
            return True
        except TimeoutException:
            return False

    # ══════════════════════════════════════════════════════════════════
    #  PDM 資料擷取（保留 v1.4 邏輯）
    # ══════════════════════════════════════════════════════════════════
    def _get_fg_name(self, fg_pn: str) -> str:
        try:
            return (self.pdm_drv.execute_script("""
                var el = document.querySelector('td[attrid="name"]');
                if (el) return (el.innerText || el.textContent || '').trim();
                var frames = document.querySelectorAll('iframe, frame');
                for (var i = 0; i < frames.length; i++) {
                    try {
                        var fdoc = frames[i].contentDocument || frames[i].contentWindow.document;
                        var fel = fdoc.querySelector('td[attrid="name"]');
                        if (fel) return (fel.innerText || fel.textContent || '').trim();
                    } catch(e) {}
                }
                return '';
            """) or '').strip()
        except Exception:
            return ''

    def _pdm_find_1199(self, fg_pn: str) -> list:
        drv = self.pdm_drv; parts = []
        try:
            time.sleep(3)
            try:
                type_drop = WebDriverWait(drv, 8).until(
                    EC.presence_of_element_located((By.XPATH,
                    '//select[option[contains(.,"全部類型")]]'
                    '|//select[option[contains(.,"All Types")]]')))
                Select(type_drop).select_by_index(0)
            except Exception:
                pass

            sb = WebDriverWait(drv, 15).until(
                EC.element_to_be_clickable((By.ID, 'gloabalSearchField')))
            sb.click(); sb.send_keys(Keys.CONTROL + 'a'); sb.send_keys(Keys.DELETE)
            sb.send_keys(fg_pn); time.sleep(0.5); sb.send_keys(Keys.RETURN)
            time.sleep(8)

            result_link = None
            xps = [
                f'//a[normalize-space(text())="{fg_pn}"]',
                f'//a[contains(.,"{fg_pn}")]',
                f'//td[contains(.,"{fg_pn}")]//a[1]',
            ]
            for _ in range(2):
                for xp in xps:
                    try:
                        result_link = WebDriverWait(drv, 5).until(
                            EC.element_to_be_clickable((By.XPATH, xp)))
                        break
                    except Exception:
                        continue
                if result_link: break
                time.sleep(4)

            if not result_link: return parts
            result_link.click(); time.sleep(6)

            struct_tab = None
            for xp in [
                '//span[normalize-space()="結構"]',
                '//a[normalize-space()="結構"]',
                '//*[contains(.,"結構") and @role="tab"]',
            ]:
                try:
                    struct_tab = WebDriverWait(drv, 10).until(
                        EC.element_to_be_clickable((By.XPATH, xp)))
                    break
                except Exception:
                    continue
            if not struct_tab: return parts
            struct_tab.click(); time.sleep(10)

            for xp in ['//a[normalize-space()="全部"]', '//button[normalize-space()="全部"]']:
                try:
                    drv.find_element(By.XPATH, xp).click(); time.sleep(5); break
                except Exception:
                    pass

            parts = self._scan_1199_via_structure_search()
        except Exception as e:
            self._ql(f'PDM 錯誤：{e}', 'ERROR')
        return parts

    def _scan_1199_via_structure_search(self) -> list:
        drv = self.pdm_drv; seen = set(); parts = []
        _js_scan = r"""
            var results = [];
            var seen1199 = {};

            function isRealPart(text) {
                if (!text) return false;
                var t = text.toUpperCase();
                var bad = ['改剖槽', '備註', '只有', 'REF', '參考', 'SPEC', '舊料'];
                for (var b of bad) { if (t.indexOf(b) >= 0) return false; }
                return true;
            }

            function getIdentText(el, num) {
                var cur = el; var best = num;
                for (var k = 0; k < 10; k++) {
                    if (!cur || cur === document.body) break;
                    var t = (cur.innerText || cur.textContent || '').replace(/[\r\n\t]+/g, ' ').replace(/  +/g, ' ').trim();
                    var idx = t.indexOf(num);
                    if (idx >= 0) {
                        var after = t.substring(idx);
                        var cut = after.search(/\s\d{10,}/);
                        var frag = (cut > 0 ? after.substring(0, cut) : after).trim();
                        frag = frag.replace(/(版本|名稱)\s*[:：]?\s*/g, ' ').replace(/  +/g, ' ').trim();
                        if (frag.length > best.length) best = frag;
                        if (frag.indexOf(',') > 0 && frag.length > num.length+2) break;
                        if (/[\u4e00-\u9fff]/.test(frag) && frag.length > num.length+10) break;
                    }
                    cur = cur.parentElement;
                }
                return best;
            }

            function extract3xParent(text) {
                var m = text.match(/(3\d{9,})/); if (!m) return null;
                var u = text.toUpperCase();
                if (u.indexOf('SRAW')<0 && u.indexOf('BO ')<0 && u.indexOf(', BO')<0 && u.indexOf('RAW')<0) return null;
                var m3 = text.match(/(3\d{9,}[^\x00-\x1F]*)/); if (!m3) return m[1];
                var res = m3[1].trim(); var idx = res.indexOf('狀態');
                res = idx > 0 ? res.substring(0, idx) : res.substring(0, 80);
                return res.replace(/(版本|名稱)\s*[:：]?\s*/g, ' ').replace(/  +/g, ' ').trim();
            }

            try {
                if (window.Ext && window.Ext.ComponentQuery) {
                    var cmpList = Ext.ComponentQuery.query('*');
                    for (var mi=0; mi<cmpList.length; mi++){
                        var cmp = cmpList[mi];
                        if (cmp && typeof cmp.getStore === 'function') {
                            var s = cmp.getStore(); if (!s) continue;
                            var recs = [];
                            try { if(s.each) s.each(function(r){recs.push(r)}); else recs=s.getRange(); } catch(e){continue;}
                            for (var ri=0; ri<recs.length; ri++){
                                var r = recs[ri]; var d = r.data || {};
                                var plain = (d.number||'') + ' ' + (d.name||'');
                                var mNum = plain.match(/(1199\d{6,})/); if (!mNum) continue;
                                var num = mNum[1]; if (seen1199[num]) continue;
                                if (!isRealPart(plain)) continue;
                                seen1199[num] = true;
                                var txt = plain.replace(/(版本|名稱)\s*[:：]?\s*/g, ' ').replace(/  +/g, ' ').trim();
                                var p3x = ''; var pN = r.parentNode;
                                while(pN){
                                    var pd = pN.data || {};
                                    var pTxt = (pd.number||'') + ' ' + (pd.name||'');
                                    var f3x = extract3xParent(pTxt); if(f3x){p3x=f3x; break;}
                                    pN = pN.parentNode;
                                }
                                results.push({ num:num, text:txt, parentText:p3x });
                            }
                        }
                    }
                }
            } catch(e){}

            var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
            var tNode;
            while (tNode = walker.nextNode()){
                var m1199 = tNode.textContent.match(/(1199\d{6,})/); if (!m1199) continue;
                var num = m1199[1]; if (seen1199[num]) continue;
                var el = tNode.parentElement; if (!el) continue;
                var rect = el.getBoundingClientRect(); if (rect.height===0) continue;
                var fullTxt = (el.innerText || el.textContent || '');
                if (!isRealPart(fullTxt)) continue;
                seen1199[num] = true;
                var txt = getIdentText(el, num);
                var p3x = ''; var cl = el;
                for (var d=0; d<25 && cl; d++){
                    cl = cl.parentElement; if(!cl || cl===document.body) break;
                    var tag = cl.tagName||'';
                    if (tag!=='LI'&&tag!=='TR'&&!cl.className.includes('node')&&!cl.className.includes('row')) continue;
                    var f3x = extract3xParent(cl.innerText||''); if(f3x){p3x=f3x; break;}
                }
                results.push({ num:num, text:txt, parentText:p3x });
            }
            return results;
        """
        try:
            drv.switch_to.default_content()
            try:
                psb = WebDriverWait(drv, 10).until(
                    EC.presence_of_element_located((By.ID, 'psbIFrame')))
                drv.switch_to.frame(psb)
            except Exception:
                return []

            sb = None
            for xp in [
                '//*[@id="StructureBrowserFindToolbar"]//input',
                '//div[contains(@id,"StructureBrowserFind")]//input',
            ]:
                try:
                    sb = WebDriverWait(drv, 5).until(
                        EC.presence_of_element_located((By.XPATH, xp)))
                    break
                except Exception:
                    continue
            if not sb: return []

            drv.execute_script("arguments[0].value = '';", sb)
            sb.click(); sb.send_keys('1199'); time.sleep(0.5)
            sb.send_keys(Keys.RETURN); time.sleep(2)

            def _collect(items):
                found = False
                for item in (items or []):
                    n = item.get('num', '')
                    if not n or n in seen: continue
                    seen.add(n); found = True
                    parts.append({'num': n,
                                  'text': item.get('text', n),
                                  'parentText': item.get('parentText', '')})
                    self._ql(f'  [{len(parts)}] {item.get("text", n)[:50]}', 'OK')
                return found

            _collect(drv.execute_script(_js_scan))
            for _ in range(15):
                if not self.running: break
                sb.send_keys(Keys.RETURN); time.sleep(1.2)
                _collect(drv.execute_script(_js_scan))
        except Exception:
            pass
        finally:
            drv.switch_to.default_content()
        return parts


# ══════════════════════════════════════════════════════════════════════
#  程式入口
# ══════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    root = tk.Tk()
    app  = App(root)
    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    ww, wh = root.winfo_width(), root.winfo_height()
    root.geometry(f'{ww}x{wh}+{(sw-ww)//2}+{(sh-wh)//2}')
    root.mainloop()
