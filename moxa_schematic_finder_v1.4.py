#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MOXA Schematic Link Finder  v1.4
=================================
自動從 PDM (Windchill) 查找成品料號下的 1199 PCB 料號，
並在 GeDCC (DMP) 取得對應線路圖連結。

v1.4 更新提示：
- 加強 1199 料號過濾邏輯，排除「只有改剖槽」等非正式零件備註。
- 增加版本特徵 (如 A.1, B.22) 辨識。
"""

# ── 正式 import ─────────────────────────────────────────────────────
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
from datetime import datetime

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
}

PDM_URL    = "https://pap.moxa.com/Windchill/app/"
GEDCC_URL  = "http://global-gedcc.moxa.com/DMP/public/login/ShowLogin.jsp"
GEDCC_HOME = "http://global-gedcc.moxa.com/DMP/private/user_index.jsp"


# ══════════════════════════════════════════════════════════════════════
#  主應用程式
# ══════════════════════════════════════════════════════════════════════
class App:

    def __init__(self, root: tk.Tk):
        self.root    = root
        self.root.title("MOXA Schematic Link Finder  v1.4")
        self.root.geometry("1280x840")
        self.root.configure(bg=C['bg_dark'])
        self.root.minsize(960, 640)

        self.q:       queue.Queue = queue.Queue()
        self.running: bool        = False
        self.pdm_drv  = None
        self.dmp_drv  = None
        self._iid_map: dict       = {}   # pcb_num → treeview iid

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
                 fg=C['accent'], bg='#0A0D13').pack(side='left', padx=(18,4), pady=8)
        tk.Label(bar, text='Schematic Link Finder v1.4',
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
        ew.pack(fill='x', padx=14, pady=(3,14))
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
            ('▶   開始檢索', C['accent'],  'normal',   self._start,  '_bs'),
            ('⏹   停  止',  '#3A4155',    'disabled', self._stop,   '_bx'),
            ('🗑   清除結果', C['bg_card'], 'normal',   self._clear,  '_bc'),
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
        self._sec(pnl, '資料匯出')
        tk.Button(pnl, text='📤  匯出 CSV',
                  font=('Segoe UI', 10), bg=C['bg_card'], fg=C['text'],
                  activebackground=C['border'],
                  relief='flat', bd=0, pady=9,
                  cursor='hand2', command=self._export).pack(
                      fill='x', padx=14, pady=(0,6))

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
        r.pack(side='left', fill='both', expand=True, padx=(0,10), pady=10)
        self._progress_card(r)
        self._log_card(r)
        self._table_card(r)

    def _progress_card(self, p):
        c = tk.Frame(p, bg=C['bg_panel'], pady=12, padx=16)
        c.pack(fill='x', pady=(0,8))
        top = tk.Frame(c, bg=C['bg_panel'])
        top.pack(fill='x')
        tk.Label(top, text='執行進度', font=('Segoe UI', 11, 'bold'),
                 fg=C['accent'], bg=C['bg_panel']).pack(side='left')
        self._pct = tk.Label(top, text='0 %',
                              font=('Consolas', 11, 'bold'),
                              fg=C['accent'], bg=C['bg_panel'])
        self._pct.pack(side='right')
        self._bar = ttk.Progressbar(c,
                                     style='Bar.Horizontal.TProgressbar',
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
                 fg=C['accent'], bg=C['bg_dark']).pack(anchor='w', pady=(0,3))
        self._log = scrolledtext.ScrolledText(
            p, font=('Consolas', 10),
            bg=C['log_bg'], fg=C['log_fg'],
            insertbackground='lime',
            relief='flat', bd=4, wrap='word', height=9,
            state='disabled')
        self._log.pack(fill='both')
        for tag, col in [
            ('INFO','#7EE787'),('STEP','#79C0FF'),('WARN','#E3A740'),
            ('ERROR','#F85149'),('OK','#56D364'),('TS','#484F58'),
        ]:
            self._log.tag_configure(tag, foreground=col)

    def _table_card(self, p):
        hdr = tk.Frame(p, bg=C['bg_dark'])
        hdr.pack(fill='x', pady=(8,3))
        tk.Label(hdr, text='檢索結果',
                 font=('Segoe UI', 11, 'bold'),
                 fg=C['accent'], bg=C['bg_dark']).pack(side='left')
        tk.Label(hdr, text='  雙擊列可開啟 GeDCC 查詢',
                 font=('Segoe UI', 10), fg=C['text_dim'],
                 bg=C['bg_dark']).pack(side='left')

        # 產品識別 banner（搜尋後更新）
        self._fg_banner = tk.Label(p,
            text='',
            font=('Segoe UI', 11, 'bold'),
            fg=C['warning'], bg=C['bg_dark'], anchor='w')
        self._fg_banner.pack(fill='x', pady=(0,3))

        wrap = tk.Frame(p, bg=C['bg_panel'])
        wrap.pack(fill='both', expand=True)

        cols = ('status','pcb')
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
    #  UI 事件
    # ══════════════════════════════════════════════════════════════════
    def _ent_in(self):
        if self._ent.get() == '例：9020240602061':
            self._ent.delete(0,'end')
            self._ent.config(fg=C['text'])

    def _ent_out(self):
        if not self._ent.get().strip():
            self._ent.insert(0,'例：9020240602061')
            self._ent.config(fg=C['text_dim'])

    def _on_hover(self, event):
        row = self._tv.identify_row(event.y)
        if row: self._tv.config(cursor='hand2')
        else: self._tv.config(cursor='')

    def _dbl(self, event):
        sel = self._tv.selection()
        if not sel: return
        vals = self._tv.item(sel[0])['values']
        if not vals or len(vals) < 2: return
        pcb_cell = str(vals[1])
        m = re.search(r'(1199\d{6,})', pcb_cell)
        if not m: return
        pcb_num = m.group(1)
        threading.Thread(target=self._open_gedcc, args=(pcb_num,), daemon=True).start()

    def _open_gedcc(self, pcb_num: str):
        """雙擊觸發：開啟 GeDCC，登入後自動填入料號並按 Enter 查詢"""
        try:
            need_new = False
            if self.dmp_drv is None: need_new = True
            else:
                try: _ = self.dmp_drv.current_url
                except Exception: need_new = True

            if need_new:
                self.dmp_drv = self._init_browser('dmp')
                if not self.dmp_drv: return
                self._q(type='dot', sys='dmp', ok=False)

            drv = self.dmp_drv
            if 'DMP/private' not in drv.current_url:
                drv.get(GEDCC_URL)
                time.sleep(1)
                self._ql(f'⚠ 請在 GeDCC 視窗登入', 'WARN')
                self._q(type='popup', sys='GeDCC (DMP)')
                WebDriverWait(drv, 180).until(lambda d: 'DMP/private' in d.current_url)
                self._ql('✅ GeDCC 登入成功', 'OK')
                self._q(type='dot', sys='dmp', ok=True)

            if 'user_index' not in drv.current_url:
                drv.get(GEDCC_HOME); time.sleep(1.5)

            try:
                sel_el = WebDriverWait(drv, 5).until(EC.presence_of_element_located((
                    By.XPATH, '//select[option[normalize-space()="料號"]]'
                    '|//select[contains(@name,"type") or contains(@id,"type")]'
                )))
                Select(sel_el).select_by_visible_text('料號')
                time.sleep(0.3)
            except Exception: pass

            for xp in [
                '//input[@name="bqKeyword"]','//input[@id="bqKeyword"]',
                '//form//input[@type="text"][1]',
            ]:
                try:
                    inp = WebDriverWait(drv, 3).until(EC.element_to_be_clickable((By.XPATH, xp)))
                    drv.execute_script("arguments[0].value = '';", inp)
                    inp.click(); inp.send_keys(pcb_num); inp.send_keys(Keys.RETURN)
                    self._ql(f'✅ GeDCC 查詢已送出：{pcb_num}', 'OK')
                    return
                except Exception: continue
            self._ql(f'❌ 找不到 GeDCC 搜尋框', 'ERROR')
        except Exception as e:
            self._ql(f'GeDCC 開啟錯誤：{str(e)[:60]}', 'ERROR')

    def _start(self):
        pn = self._ent.get().strip()
        if not pn or pn == '例：9020240602061':
            messagebox.showwarning('輸入錯誤','請輸入有效的成品料號'); return
        self.running = True
        self._bs.config(state='disabled')
        self._bx.config(state='normal')
        self._ent.config(state='disabled')
        threading.Thread(target=self._run, args=(pn,), daemon=True).start()

    def _stop(self):
        self.running = False
        self._log_w('⏹ 使用者中止','WARN')

    def _clear(self):
        for i in self._tv.get_children(): self._tv.delete(i)
        self._iid_map.clear()
        self._fg_banner.config(text='')
        self._log.config(state='normal')
        self._log.delete('1.0','end')
        self._log.config(state='disabled')
        self._bar['value'] = 0
        self._pct.config(text='0 %')
        self._stlbl.config(text='就緒 — 請輸入成品料號後點擊「開始檢索」')
        self._upd_stats()

    def _export(self):
        items = self._tv.get_children()
        if not items:
            messagebox.showinfo('無資料','目前沒有可匯出的結果'); return
        fname = f"schematic_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        path  = os.path.join(os.path.expanduser('~'), 'Desktop', fname)
        with open(path,'w',newline='',encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(['狀態','1199 PCB 料號/名稱'])
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

    # ── 訊息佇列 ────────────────────────────────────────────────────────
    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                t = m.get('type')
                if t == 'log': self._log_w(m['text'], m.get('lv','INFO'))
                elif t == 'prog':
                    v = int(m['v'])
                    self._bar['value'] = v
                    self._pct.config(text=f'{v} %')
                    if m.get('s'): self._stlbl.config(text=m['s'])
                elif t == 'dot':
                    d = self._pdm_dot if m['sys']=='pdm' else self._dmp_dot
                    d.config(fg=C['success'] if m['ok'] else C['error'])
                elif t == 'add':
                    iid = self._tv.insert('','end',
                        values=(m['st'], m.get('pcbtxt', m['pcb'])),
                        tags=(m.get('tag','ok'),))
                    self._iid_map[m['pcb']] = iid
                    self._upd_stats()
                elif t == 'banner':
                    self._fg_banner.config(text=m.get('text',''))
                elif t == 'popup':
                    messagebox.showinfo('需要手動登入',
                        f'請在彈出的 {m["sys"]} 瀏覽器視窗中完成帳密輸入。\n\n登入後程式將自動繼續。')
                elif t == 'done': self._on_done()
        except queue.Empty: pass
        self.root.after(80, self._poll)

    def _q(self, **kw): self.q.put(kw)
    def _log_w(self, t, lv='INFO'):
        self._log.config(state='normal')
        ts = datetime.now().strftime('%H:%M:%S')
        self._log.insert('end', f'[{ts}] ', 'TS')
        self._log.insert('end', f'{t}\n', lv)
        self._log.see('end'); self._log.config(state='disabled')
    def _ql(self, t, lv='INFO'):  self._q(type='log', text=t, lv=lv)
    def _qp(self, v, s=''):       self._q(type='prog', v=v, s=s)

    # ══════════════════════════════════════════════════════════════════
    #  自動化主流程（背景執行緒）
    # ══════════════════════════════════════════════════════════════════
    def _run(self, fg_pn: str):
        try:
            self._ql(f'🚀 開始檢索：{fg_pn}', 'STEP')
            self._qp(5, '正在初始化瀏覽器...')
            self.pdm_drv = self._init_browser('pdm')
            if not self.pdm_drv:
                self._ql('❌ 無法啟動 Edge 或 Chrome 瀏覽器', 'ERROR'); self._q(type='done'); return

            self._qp(10, '正在開啟 PDM 登入頁面...')
            self.pdm_drv.get(PDM_URL)
            self._ql('⚠ 請在彈出的 PDM 瀏覽器視窗輸入帳號密碼', 'WARN')
            self._q(type='popup', sys='PDM (Windchill)')
            if not self._wait_pdm_login():
                self._ql('❌ PDM 登入等候逾時', 'ERROR'); self._q(type='done'); return

            self._q(type='dot', sys='pdm', ok=True)
            self._ql('✅ PDM 登入成功', 'OK')
            self._qp(20, f'正在 PDM 搜尋 {fg_pn}...')

            pcb_list = self._pdm_find_1199(fg_pn)
            if not pcb_list:
                self._ql('❌ BOM 中未找到任何 1199 PCB 料號', 'ERROR'); self._q(type='done'); return

            self._ql(f'✅ 共找到 {len(pcb_list)} 個 1199 PCB 料號', 'OK')
            self._qp(55, f'共 {len(pcb_list)} 個 PCB 料號')

            fg_name = self._get_fg_name(fg_pn)
            info_txt = f'{fg_pn}\n{fg_name}' if fg_name else fg_pn
            self._q(type='banner', text=f'🔎  {fg_pn}  {fg_name}'.strip(), info=info_txt)

            if not self.running: self._q(type='done'); return

            for p in pcb_list:
                pcb_txt = p.get('text') or p['num']
                fg_txt = p.get('parentText') or fg_pn
                pcb_txt = re.sub(r'(版本|名稱)\s*[:：]?\s*', ' ', pcb_txt).split('狀態')[0].strip()
                fg_txt = re.sub(r'(版本|名稱)\s*[:：]?\s*', ' ', fg_txt).split('狀態')[0].strip()

                self._q(type='add', st='✅ 找到', pcb=p['num'], pcbtxt=pcb_txt, fgtxt=fg_txt, tag='ok')

            self._qp(100, '✅ 檢索完成！')
            self._ql('🎉 掃描完畢！雙擊任一列可開啟 GeDCC 查詢。', 'OK')
        except Exception as exc:
            self._ql(f'❌ 執行例外：{exc}', 'ERROR')
        finally: self._q(type='done')

    def _init_browser(self, label='pdm'):
        import subprocess as _sp
        # 嘗試啟動 Edge
        try:
            o = EdgeOptions()
            for a in ['--ignore-certificate-errors', '--ignore-ssl-errors', '--log-level=3', '--silent']: o.add_argument(a)
            o.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
            o.add_experimental_option('prefs', {'protocol_handler.excluded_schemes': {'wnjp': False}})
            drv = webdriver.Edge(service=EdgeService(log_output=_sp.DEVNULL), options=o)
            self._ql(f'✅ 已成功連結 Edge 瀏覽器 ({label})', 'OK')
            return drv
        except Exception as ee:
            self._ql(f'ℹ️ Edge 啟動未成 ({ee})，嘗試啟動 Chrome...', 'INFO')
        
        # 嘗試啟動 Chrome
        try:
            o = ChromeOptions()
            for a in ['--ignore-certificate-errors', '--ignore-ssl-errors', '--log-level=3', '--silent']: o.add_argument(a)
            o.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
            drv = webdriver.Chrome(service=ChromeService(log_output=_sp.DEVNULL), options=o)
            self._ql(f'✅ 已成功連結 Chrome 瀏覽器 ({label})', 'OK')
            return drv
        except Exception as ce:
            self._ql(f'❌ Chrome 啟動也失敗 ({ce})', 'ERROR')
            return None

    def _wait_pdm_login(self, timeout=180) -> bool:
        drv = self.pdm_drv
        try:
            WebDriverWait(drv, timeout).until(EC.presence_of_element_located((By.ID, 'gloabalSearchField')))
            self._ql('✅ 偵測到搜尋框，PDM 登入完成', 'OK'); time.sleep(1); return True
        except TimeoutException: return False

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
        except Exception: return ''

    def _pdm_find_1199(self, fg_pn: str) -> list:
        drv = self.pdm_drv; wait = WebDriverWait(drv, 40); parts = []
        try:
            time.sleep(3)
            try:
                type_drop = WebDriverWait(drv, 8).until(EC.presence_of_element_located((By.XPATH,
                    '//select[option[contains(.,"全部類型")]]|//select[option[contains(.,"All Types")]]')))
                Select(type_drop).select_by_index(0)
            except: pass

            sb = WebDriverWait(drv, 15).until(EC.element_to_be_clickable((By.ID, 'gloabalSearchField')))
            sb.click(); sb.send_keys(Keys.CONTROL + 'a'); sb.send_keys(Keys.DELETE)
            sb.send_keys(fg_pn); time.sleep(0.5); sb.send_keys(Keys.RETURN)
            time.sleep(8)

            result_link = None
            xps = [f'//a[normalize-space(text())="{fg_pn}"]', f'//a[contains(.,"{fg_pn}")]', f'//td[contains(.,"{fg_pn}")]//a[1]']
            for _a in range(2):
                for xp in xps:
                    try: result_link = WebDriverWait(drv, 5).until(EC.element_to_be_clickable((By.XPATH, xp))); break
                    except: continue
                if result_link: break
                time.sleep(4)

            if not result_link: return parts
            result_link.click(); time.sleep(6)

            struct_tab = None
            for xp in ['//span[normalize-space()="結構"]','//a[normalize-space()="結構"]','//*[contains(.,"結構") and @role="tab"]']:
                try: struct_tab = WebDriverWait(drv, 10).until(EC.element_to_be_clickable((By.XPATH, xp))); break
                except: continue
            if not struct_tab: return parts
            struct_tab.click(); time.sleep(10)

            for xp in ['//a[normalize-space()="全部"]','//button[normalize-space()="全部"]']:
                try: drv.find_element(By.XPATH, xp).click(); time.sleep(5); break
                except: pass

            parts = self._scan_1199_via_structure_search()
        except Exception as e: self._ql(f'PDM 錯誤：{e}', 'ERROR')
        return parts

    def _scan_1199_via_structure_search(self) -> list:
        drv = self.pdm_drv; seen = set(); parts = []
        _js_scan = r"""
            var results = [];
            var seen1199 = {};

            // v1.4 新增：判斷是否為真正的零件（排除備註如「只有改剖槽」）
            function isRealPart(text) {
                if (!text) return false;
                var t = text.toUpperCase();
                // 1. 排除明確的雜訊關鍵字
                var bad = ['改剖槽', '備註', '只有', 'REF', '參考', 'SPEC', '舊料'];
                for (var b of bad) { if (t.indexOf(b) >= 0) return false; }
                // 2. 調鬆條件：不再強制檢查版本格式，只要沒有上述雜訊詞彙即視為有效
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

            // 方法 1: ExtJS Store
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
                                
                                // v1.4 新增：檢查零件有效性
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

            // 方法 2: DOM Scan
            var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, null, false);
            var tNode;
            while (tNode = walker.nextNode()){
                var m1199 = tNode.textContent.match(/(1199\d{6,})/); if (!m1199) continue;
                var num = m1199[1]; if (seen1199[num]) continue;
                var el = tNode.parentElement; if (!el) continue;
                var rect = el.getBoundingClientRect(); if (rect.height===0) continue;
                
                var fullTxt = (el.innerText || el.textContent || '');
                // v1.4 新增：檢查零件有效性
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
                psb = WebDriverWait(drv,10).until(EC.presence_of_element_located((By.ID,'psbIFrame')))
                drv.switch_to.frame(psb)
            except: return []

            sb = None
            for xp in ['//*[@id="StructureBrowserFindToolbar"]//input','//div[contains(@id,"StructureBrowserFind")]//input']:
                try: sb = WebDriverWait(drv,5).until(EC.presence_of_element_located((By.XPATH,xp))); break
                except: continue
            if not sb: return []

            drv.execute_script("arguments[0].value = '';", sb)
            sb.click(); sb.send_keys('1199'); time.sleep(0.5); sb.send_keys(Keys.RETURN); time.sleep(2)

            def _collect(items):
                found = False
                for item in (items or []):
                    n = item.get('num',''); 
                    if not n or n in seen: continue
                    seen.add(n); found = True
                    parts.append({'num':n, 'text':item.get('text',n), 'parentText':item.get('parentText','')})
                    self._ql(f'  [{len(parts)}] {item.get("text",n)[:50]}', 'OK')
                return found

            _collect(drv.execute_script(_js_scan))
            for _ in range(15):
                if not self.running: break
                sb.send_keys(Keys.RETURN); time.sleep(1.2)
                _collect(drv.execute_script(_js_scan))
        except: pass
        finally: drv.switch_to.default_content()
        return parts

if __name__ == '__main__':
    root = tk.Tk(); app = App(root)
    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    ww, wh = root.winfo_width(), root.winfo_height()
    root.geometry(f'{ww}x{wh}+{(sw-ww)//2}+{(sh-wh)//2}')
    root.mainloop()
