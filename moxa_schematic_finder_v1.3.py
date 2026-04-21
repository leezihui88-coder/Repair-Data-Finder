#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MOXA Schematic Link Finder  v1.3
=================================
自動從 PDM (Windchill) 查找成品料號下的 1199 PCB 料號，
並在 GeDCC (DMP) 取得對應線路圖連結。

執行方式：雙擊 run.bat
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
from selenium.webdriver.edge.options import Options as EdgeOptions


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
        self.root.title("MOXA Schematic Link Finder  v1.3")
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
                    fieldbackground=C['bg_panel'], rowheight=28,
                    borderwidth=0, font=('Consolas', 9))
        s.configure('T.Treeview.Heading',
                    background=C['bg_card'], foreground=C['accent'],
                    relief='flat', font=('Segoe UI', 9, 'bold'))
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
        tk.Label(bar, text='Schematic Link Finder',
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
        pnl = tk.Frame(parent, bg=C['bg_panel'], width=220)
        pnl.pack(side='left', fill='y', padx=(10, 6), pady=10)
        pnl.pack_propagate(False)

        self._sec(pnl, '料號輸入', top=18)

        ew = tk.Frame(pnl, bg=C['border'], bd=1, relief='solid')
        ew.pack(fill='x', padx=14, pady=(3,14))
        self._ent = tk.Entry(ew, font=('Consolas', 11),
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
                          font=('Segoe UI', 10,
                                'bold' if attr == '_bs' else 'normal'),
                          bg=bg, fg='white',
                          activebackground=C['accent_dim'],
                          activeforeground='white',
                          relief='flat', bd=0, pady=10,
                          cursor='hand2', state=st, command=cmd)
            b.pack(fill='x', padx=14, pady=(0, 6))
            setattr(self, attr, b)

        self._div(pnl)
        self._sec(pnl, '產品資訊')

        fg_card = tk.Frame(pnl, bg=C['bg_card'], pady=8, padx=12)
        fg_card.pack(fill='x', padx=14, pady=(0,4))
        tk.Label(fg_card, text='料號 / 名稱',
                 font=('Segoe UI', 8), fg=C['text_dim'],
                 bg=C['bg_card']).pack(anchor='w')
        self._fg_info = tk.Label(fg_card,
            text='—',
            font=('Segoe UI', 9, 'bold'),
            fg=C['warning'], bg=C['bg_card'],
            wraplength=220, justify='left', anchor='w')
        self._fg_info.pack(fill='x', anchor='w')

        self._div(pnl)
        self._sec(pnl, '統計摘要')
        self._sv_total   = self._stat(pnl, '找到 PCB 料號', '0')
        self._sv_done    = self._stat(pnl, '✅ 已完成',      '0')
        self._sv_pending = self._stat(pnl, '⏳ 處理中',      '0')
        self._sv_err     = self._stat(pnl, '⚠ 無結果',      '0')

        self._div(pnl)
        self._sec(pnl, '資料匯出')
        tk.Button(pnl, text='📤  匯出 CSV',
                  font=('Segoe UI', 9), bg=C['bg_card'], fg=C['text'],
                  activebackground=C['border'],
                  relief='flat', bd=0, pady=9,
                  cursor='hand2', command=self._export).pack(
                      fill='x', padx=14, pady=(0,6))

    def _sec(self, p, t, top=10):
        tk.Label(p, text=t, font=('Segoe UI', 10, 'bold'),
                 fg=C['accent'], bg=C['bg_panel']).pack(
                     anchor='w', padx=14, pady=(top, 4))

    def _div(self, p):
        tk.Frame(p, bg=C['border'], height=1).pack(fill='x', padx=14, pady=12)

    def _stat(self, p, label, val):
        c = tk.Frame(p, bg=C['bg_card'], pady=7, padx=12)
        c.pack(fill='x', padx=14, pady=2)
        tk.Label(c, text=label, font=('Segoe UI', 8),
                 fg=C['text_dim'], bg=C['bg_card']).pack(anchor='w')
        lv = tk.Label(c, text=val, font=('Consolas', 18, 'bold'),
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
        tk.Label(top, text='執行進度', font=('Segoe UI', 10, 'bold'),
                 fg=C['accent'], bg=C['bg_panel']).pack(side='left')
        self._pct = tk.Label(top, text='0 %',
                              font=('Consolas', 10, 'bold'),
                              fg=C['accent'], bg=C['bg_panel'])
        self._pct.pack(side='right')
        self._bar = ttk.Progressbar(c,
                                     style='Bar.Horizontal.TProgressbar',
                                     mode='determinate')
        self._bar.pack(fill='x', pady=(6, 5))
        self._stlbl = tk.Label(c,
                                text='就緒 — 請輸入成品料號後點擊「開始檢索」',
                                font=('Segoe UI', 8), fg=C['text_dim'],
                                bg=C['bg_panel'], anchor='w')
        self._stlbl.pack(fill='x')

    def _log_card(self, p):
        tk.Label(p, text='執行日誌 (Real-time Log)',
                 font=('Segoe UI', 9, 'bold'),
                 fg=C['accent'], bg=C['bg_dark']).pack(anchor='w', pady=(0,3))
        self._log = scrolledtext.ScrolledText(
            p, font=('Consolas', 8),
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
                 font=('Segoe UI', 9, 'bold'),
                 fg=C['accent'], bg=C['bg_dark']).pack(side='left')
        tk.Label(hdr, text='  雙擊列可開啟 GeDCC 查詢',
                 font=('Segoe UI', 8), fg=C['text_dim'],
                 bg=C['bg_dark']).pack(side='left')

        # 產品識別 banner（搜尋後更新）
        self._fg_banner = tk.Label(p,
            text='',
            font=('Segoe UI', 9), fg=C['warning'],
            bg=C['bg_dark'], anchor='w')
        self._fg_banner.pack(fill='x', pady=(0,3))

        wrap = tk.Frame(p, bg=C['bg_panel'])
        wrap.pack(fill='both', expand=True)

        cols = ('status','fg','pcb')
        self._tv = ttk.Treeview(wrap, columns=cols,
                                 show='headings', style='T.Treeview')
        for col, hd, w, anc in [
            ('status', '狀態',                    90, 'center'),
            ('fg',     '母階 料號/名稱',          420, 'w'),
            ('pcb',    '1199 PCB 料號/名稱（雙擊開啟 GeDCC 查詢）', 620, 'w'),
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
        if row:
            self._tv.config(cursor='hand2')
        else:
            self._tv.config(cursor='')

    def _dbl(self, event):
        sel = self._tv.selection()
        if not sel: return
        vals = self._tv.item(sel[0])['values']
        if not vals or len(vals) < 3: return
        pcb_cell = str(vals[2])
        m = re.search(r'(1199\d{6,})', pcb_cell)
        if not m: return
        pcb_num = m.group(1)
        threading.Thread(target=self._open_gedcc, args=(pcb_num,), daemon=True).start()

    def _open_gedcc(self, pcb_num: str):
        """雙擊觸發：開啟 GeDCC，登入後自動填入料號並按 Enter 查詢"""
        try:
            # 建立或重用 GeDCC 瀏覽器
            need_new = False
            if self.dmp_drv is None:
                need_new = True
            else:
                try:
                    _ = self.dmp_drv.current_url
                except Exception:
                    need_new = True

            if need_new:
                self.dmp_drv = webdriver.Edge(
                    service=self._edge_service(), options=self._edge_opts())
                self._q(type='dot', sys='dmp', ok=False)

            drv = self.dmp_drv

            # 若未登入，先導向登入頁
            if 'DMP/private' not in drv.current_url:
                drv.get(GEDCC_URL)
                time.sleep(1)
                self._ql(f'⚠ 請在 GeDCC 視窗登入', 'WARN')
                self._q(type='popup', sys='GeDCC (DMP)')
                WebDriverWait(drv, 180).until(
                    lambda d: 'DMP/private' in d.current_url)
                self._ql('✅ GeDCC 登入成功', 'OK')
                self._q(type='dot', sys='dmp', ok=True)

            # 確保在首頁
            if 'user_index' not in drv.current_url:
                drv.get(GEDCC_HOME)
                time.sleep(1.5)

            # 選「料號」
            try:
                sel_el = WebDriverWait(drv, 5).until(EC.presence_of_element_located((
                    By.XPATH,
                    '//select[option[normalize-space()="料號"]]'
                    '|//select[contains(@name,"type") or contains(@id,"type")]'
                )))
                Select(sel_el).select_by_visible_text('料號')
                time.sleep(0.3)
            except Exception:
                pass

            # 填入料號並按 Enter
            for xp in [
                '//input[@name="bqKeyword"]',
                '//input[@id="bqKeyword"]',
                '//input[contains(@name,"Keyword") or contains(@id,"Keyword")]',
                '//input[@type="text" and contains(@name,"keyword")]',
                '//form//input[@type="text"][1]',
            ]:
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
        self._fg_info.config(text='—')
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
            w.writerow(['狀態','母階 料號/名稱','1199 PCB 料號/名稱'])
            for i in items: w.writerow(self._tv.item(i)['values'])
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

    # ══════════════════════════════════════════════════════════════════
    #  訊息佇列
    # ══════════════════════════════════════════════════════════════════
    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                t = m.get('type')
                if t == 'log':
                    self._log_w(m['text'], m.get('lv','INFO'))
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
                        values=(m['st'], m.get('fgtxt',''), m.get('pcbtxt', m['pcb'])),
                        tags=(m.get('tag','ok'),))
                    self._iid_map[m['pcb']] = iid
                    self._upd_stats()
                elif t == 'banner':
                    self._fg_banner.config(text=m.get('text',''))
                    self._fg_info.config(text=m.get('info','—'))
                elif t == 'popup':
                    messagebox.showinfo('需要手動登入',
                        f'請在彈出的 {m["sys"]} 瀏覽器視窗中完成帳密輸入。\n\n'
                        '登入後程式將自動繼續。')
                elif t == 'done':
                    self._on_done()
        except queue.Empty:
            pass
        self.root.after(80, self._poll)

    # ── 快捷發訊 ──────────────────────────────────────────────────────
    def _q(self, **kw):             self.q.put(kw)
    def _log_w(self, t, lv='INFO'):
        self._log.config(state='normal')
        ts = datetime.now().strftime('%H:%M:%S')
        self._log.insert('end', f'[{ts}] ', 'TS')
        self._log.insert('end', f'{t}\n', lv)
        self._log.see('end')
        self._log.config(state='disabled')
    def _ql(self, t, lv='INFO'):    self._q(type='log', text=t, lv=lv)
    def _qp(self, v, s=''):         self._q(type='prog', v=v, s=s)

    # ══════════════════════════════════════════════════════════════════
    #  自動化主流程（背景執行緒）
    # ══════════════════════════════════════════════════════════════════
    def _run(self, fg_pn: str):
        try:
            self._ql(f'🚀 開始檢索：{fg_pn}', 'STEP')
            self._qp(5, '正在初始化瀏覽器...')

            # ① 啟動 PDM 瀏覽器（Edge，Selenium 自動管理 Driver）
            self._qp(8, '正在開啟 PDM (Windchill) 瀏覽器 (Edge)...')
            try:
                self.pdm_drv = webdriver.Edge(
                    service=self._edge_service(), options=self._edge_opts())
            except Exception as e:
                self._ql(f'❌ Edge 啟動失敗：{e}', 'ERROR')
                self._q(type='done'); return

            # ③ PDM 登入（HTTP Basic Auth → 使用者手動輸入）
            self._qp(10, '正在開啟 PDM 登入頁面...')
            self.pdm_drv.get(PDM_URL)
            self._ql('⚠ 請在彈出的 PDM 瀏覽器視窗輸入帳號密碼', 'WARN')
            self._q(type='popup', sys='PDM (Windchill)')

            if not self._wait_pdm_login():
                self._ql('❌ PDM 登入等候逾時（120 秒）', 'ERROR')
                self._q(type='done'); return

            self._q(type='dot', sys='pdm', ok=True)
            self._ql('✅ PDM 登入成功', 'OK')
            self._qp(20, f'正在 PDM 搜尋 {fg_pn}...')

            # ④ 搜尋 FG 料號，擷取 BOM 中所有 1199 料號
            pcb_list = self._pdm_find_1199(fg_pn)
            if not pcb_list:
                self._ql('❌ BOM 中未找到任何 1199 PCB 料號', 'ERROR')
                self._q(type='done'); return

            self._ql(f'✅ 共找到 {len(pcb_list)} 個 1199 PCB 料號', 'OK')
            self._qp(55, f'共 {len(pcb_list)} 個 PCB 料號')

            # 讀取 FG 料號的名稱，顯示在 banner
            fg_name = self._get_fg_name(fg_pn)
            info_txt = f'{fg_pn}\n{fg_name}' if fg_name else fg_pn
            self._q(type='banner',
                    text=f'🔎  {fg_pn}  {fg_name}'.strip(),
                    info=info_txt)

            if not self.running: self._q(type='done'); return

            for p in pcb_list:
                # 統一檢查並移除多餘的「版本：」與「名稱：」字眼，並截斷「狀態」及其後的文字
                pcb_txt = p.get('text') or p['num']
                fg_txt = p.get('parentText') or fg_pn
                pcb_txt = re.sub(r'(版本|名稱)\s*[:：]?\s*', ' ', pcb_txt).split('狀態')[0].strip()
                fg_txt = re.sub(r'(版本|名稱)\s*[:：]?\s*', ' ', fg_txt).split('狀態')[0].strip()

                self._q(type='add',
                         st='✅ 找到',
                         pcb=p['num'],
                         pcbtxt=pcb_txt,
                         fgtxt=fg_txt,
                         tag='ok')

            self._qp(100, '✅ 檢索完成！')
            self._ql('🎉 掃描完畢！雙擊任一列可開啟 GeDCC 查詢。', 'OK')

        except Exception as exc:
            self._ql(f'❌ 執行例外：{exc}', 'ERROR')
        finally:
            self._q(type='done')

    # ── Edge 選項（含靜音 EdgeDriver 系統訊息）────────────────────────
    @staticmethod
    def _edge_opts():
        o = EdgeOptions()
        o.add_argument('--ignore-certificate-errors')
        o.add_argument('--ignore-ssl-errors')
        o.add_argument('--log-level=3')          # 靜音 Edge 系統訊息
        o.add_argument('--silent')
        o.add_experimental_option('excludeSwitches',
                                   ['enable-logging', 'enable-automation'])
        o.add_experimental_option('prefs', {
            'protocol_handler.excluded_schemes': {'wnjp': False}
        })
        return o

    @staticmethod
    def _edge_service():
        """EdgeDriver Service，將 driver log 導向 DEVNULL"""
        import subprocess as _sp
        return EdgeService(log_output=_sp.DEVNULL)

    # ══════════════════════════════════════════════════════════════════
    #  PDM (Windchill) 自動化
    # ══════════════════════════════════════════════════════════════════
    def _wait_pdm_login(self, timeout=180) -> bool:
        """
        等候 Windchill 真正登入完成。
        判斷依據：ID=gloabalSearchField 的搜尋框出現在頁面上。
        （HTTP Basic Auth 彈窗關閉、頁面渲染完成後才會出現此元素）
        """
        drv = self.pdm_drv
        self._ql('等候 PDM 登入完成（請在 Edge 輸入帳密）...', 'WARN')
        try:
            WebDriverWait(drv, timeout).until(
                EC.presence_of_element_located((By.ID, 'gloabalSearchField'))
            )
            self._ql('✅ 偵測到搜尋框，PDM 登入完成', 'OK')
            time.sleep(1)
            return True
        except TimeoutException:
            self._ql('❌ 等候 PDM 登入逾時（180 秒）', 'ERROR')
            return False

    def _get_fg_name(self, fg_pn: str) -> str:
        """從 Windchill 屬性面板 td[attrid="name"] 讀取產品名稱"""
        try:
            drv = self.pdm_drv
            result = drv.execute_script("""
                // 主文件直接找
                var el = document.querySelector('td[attrid="name"]');
                if (el) {
                    var v = (el.innerText || el.textContent || '').trim();
                    if (v) return v;
                }
                // 穿透 iframe 找
                var frames = document.querySelectorAll('iframe, frame');
                for (var i = 0; i < frames.length; i++) {
                    try {
                        var fdoc = frames[i].contentDocument || frames[i].contentWindow.document;
                        var fel = fdoc.querySelector('td[attrid="name"]');
                        if (fel) {
                            var fv = (fel.innerText || fel.textContent || '').trim();
                            if (fv) return fv;
                        }
                    } catch(e) {}
                }
                return '';
            """)
            return (result or '').strip()
        except Exception:
            return ''

    def _pdm_find_1199(self, fg_pn: str) -> list:
        """
        搜尋 FG 料號 → 結構頁籤 → 收集 1199 PCB 料號
        Windchill 右上角搜尋流程：
          1. 下拉選「全部類型」
          2. 在搜尋框輸入料號
          3. 點放大鏡按鈕
        """
        drv   = self.pdm_drv
        wait  = WebDriverWait(drv, 40)
        parts = []

        try:
            self._ql(f'頁面標題：{drv.title[:50]}', 'INFO')
            self._qp(22, '等候 Windchill 首頁穩定...')
            time.sleep(3)

            # ── Step 1: 確保「全部類型」下拉已選 ─────────────────
            self._qp(24, '設定搜尋類型為「全部類型」...')
            self._ql('尋找類型下拉選單...', 'INFO')
            try:
                # Windchill 類型下拉（select 或 combobox）
                type_drop = WebDriverWait(drv, 8).until(
                    EC.presence_of_element_located((By.XPATH,
                        '//select[option[contains(.,"全部類型")]]'
                        '|//select[option[contains(.,"All Types")]]'
                    )))
                Select(type_drop).select_by_index(0)   # 第一項通常是全部類型
                self._ql('已選擇「全部類型」', 'OK')
            except (TimeoutException, NoSuchElementException):
                self._ql('（類型下拉不需設定或已預設全部類型）', 'INFO')

            # ── Step 2: 直接用 ID 定位搜尋框並輸入料號 ───────────
            # ID = gloabalSearchField（Windchill 原始拼法，非筆誤）
            self._qp(25, '定位搜尋框 #gloabalSearchField ...')
            self._ql(f'輸入料號：{fg_pn}', 'STEP')

            sb = WebDriverWait(drv, 15).until(
                EC.element_to_be_clickable((By.ID, 'gloabalSearchField')))
            sb.click()
            time.sleep(0.5)
            # ExtJS 搜尋框：先 triple-click 全選，再 send_keys 輸入
            sb.send_keys(Keys.CONTROL + 'a')
            time.sleep(0.2)
            sb.send_keys(Keys.DELETE)
            time.sleep(0.2)
            sb.send_keys(fg_pn)
            self._ql(f'料號已填入搜尋框：{sb.get_attribute("value")}', 'OK')
            time.sleep(0.5)

            # ── Step 3: 按 Enter 觸發搜尋 ─────────────────────────
            self._ql('按 Enter 觸發搜尋...', 'INFO')
            sb.send_keys(Keys.RETURN)
            self._qp(28, '等候搜尋結果...')
            time.sleep(6)

            # ── Step 4: 點擊搜尋結果中的料號連結 ─────────────────
            self._ql(f'在搜尋結果中尋找 {fg_pn}...', 'INFO')
            self._ql(f'目前 URL：{drv.current_url[:80]}', 'INFO')

            result_link = None
            xpaths = [
                f'//a[normalize-space(text())="{fg_pn}"]',
                f'//a[contains(normalize-space(text()),"{fg_pn}")]',
                f'//a[contains(@title,"{fg_pn}")]',
                f'//td[contains(.,"{fg_pn}")]//a[1]',
                f'//*[contains(@id,"searchResult")]//a[contains(.,"{fg_pn}")]',
                f'//span[contains(.,"{fg_pn}")]//ancestor::a[1]',
                f'//div[contains(@class,"listViewRow")]//a[contains(.,"{fg_pn}")]',
            ]
            # 先等 3 秒讓結果渲染完
            time.sleep(3)
            for attempt in range(2):
                for xp in xpaths:
                    try:
                        result_link = WebDriverWait(drv, 10).until(
                            EC.element_to_be_clickable((By.XPATH, xp)))
                        self._ql(f'找到連結（XPath#{xpaths.index(xp)}），點擊...', 'INFO')
                        break
                    except TimeoutException:
                        continue
                if result_link:
                    break
                # 若第一輪沒找到，記錄頁面標題再等一下
                self._ql(f'第 {attempt+1} 輪未找到，頁面標題：{drv.title[:50]}', 'WARN')
                time.sleep(4)

            if not result_link:
                self._ql(f'❌ 搜尋結果中找不到 {fg_pn}，請確認料號正確', 'ERROR')
                self._ql(f'頁面標題：{drv.title}', 'INFO')
                self._ql(f'URL：{drv.current_url[:100]}', 'INFO')
                return parts

            result_link.click()
            self._qp(35, '開啟料號詳細頁面...')
            time.sleep(5)
            self._ql(f'料號頁面：{drv.title or drv.current_url[:50]}', 'INFO')

            # ── Step 3: 點擊「結構」頁籤 ──────────────────────────
            self._qp(38, '尋找並點擊「結構」頁籤...')
            struct_tab = None
            for xp in [
                '//span[normalize-space()="結構"]',
                '//a[normalize-space()="結構"]',
                '//*[@role="tab" and contains(.,"結構")]',
                '//li[contains(.,"結構")]//a',
            ]:
                try:
                    struct_tab = WebDriverWait(drv, 10).until(
                        EC.element_to_be_clickable((By.XPATH, xp)))
                    break
                except TimeoutException:
                    continue

            if not struct_tab:
                self._ql('❌ 找不到「結構」頁籤', 'ERROR')
                return parts

            struct_tab.click()
            self._ql('已點擊「結構」頁籤，等候 BOM 載入...', 'INFO')
            self._qp(42, '正在載入 BOM 結構（最多等候 20 秒）...')
            time.sleep(10)

            # ── Step 4: 展開全部 ───────────────────────────────────
            for xp in [
                '//a[normalize-space()="全部"]',
                '//button[normalize-space()="全部"]',
                '//*[normalize-space()="全部" and (@href or @onclick)]',
            ]:
                try:
                    btn = drv.find_element(By.XPATH, xp)
                    btn.click()
                    self._ql('已點擊「全部展開」', 'INFO')
                    time.sleep(5)
                    break
                except NoSuchElementException:
                    continue

            # ── Step 5: 用結構搜尋框輸入 1199，逐一讀取結果 ──────
            self._qp(48, '在結構搜尋框輸入 1199...')
            parts = self._scan_1199_via_structure_search()

        except TimeoutException as e:
            self._ql(f'PDM 操作逾時（可能頁面載入太慢）：{str(e)[:60]}', 'ERROR')
        except Exception as e:
            self._ql(f'PDM 操作錯誤：{str(e)[:80]}', 'ERROR')

        return parts

    def _scan_1199(self) -> list:
        """掃描 Windchill 結構樹，以 regex 抓取所有 1199 料號"""
        drv   = self.pdm_drv
        seen  = set()
        parts = []
        pat   = re.compile(r'\b(1199\d{8,})\b')

        # 從整頁 body 文字掃描（最可靠）
        try:
            body_text = drv.find_element(By.TAG_NAME, 'body').text
            for m in pat.finditer(body_text):
                num = m.group(1)
                if num not in seen:
                    seen.add(num)
        except Exception:
            pass

        # 再從個別元素取名稱
        elems = drv.find_elements(By.XPATH,
            '//*[contains(normalize-space(text()),"1199")'
            ' and not(self::script) and not(self::style)]')

        for el in elems:
            try:
                txt = el.text.strip()
                if not txt:
                    continue
                for m in pat.finditer(txt):
                    num = m.group(1)
                    # 加入 seen 中還沒取得名稱的
                    ci   = txt.find(',')
                    name = txt[ci+1:].strip() if ci >= 0 else txt
                    # 存入 parts（去重）
                    if not any(p['num'] == num for p in parts):
                        parts.append({'num': num, 'name': name[:80]})
                        self._ql(f'  → {num}  {name[:48]}', 'INFO')
            except StaleElementReferenceException:
                continue

        return parts

    def _scan_1199_via_structure_search(self) -> list:
        """
        在結構頁籤的「在結構內搜尋」框輸入 1199 + Enter，
        用 DOM 樹狀層級找到每個 1199 的直接母階料號，去重後回傳。
        """
        drv  = self.pdm_drv
        seen = set()
        parts = []

        # JS 掃描邏輯：用 DOM 樹狀層級 (ExtJS TreePanel: li > ul > li) 找母階
        _js_scan = r"""
            var results = [];
            var seen1199 = {};

            function getIdentText(el, num) {
                var cur = el;
                var best = num;
                for (var k = 0; k < 10; k++) {
                    if (!cur || cur === document.body) break;
                    var t = (cur.innerText || cur.textContent || '')
                                .replace(/[\r\n\t]+/g, ' ')
                                .replace(/  +/g, ' ').trim();
                    var idx = t.indexOf(num);
                    if (idx >= 0) {
                        var after = t.substring(idx);
                        var cut = after.search(/\s\d{10,}/);
                        var frag = (cut > 0 ? after.substring(0, cut) : after).trim();
                        
                        // 移除「版本：」和「名稱：」
                        frag = frag.replace(/(版本|名稱)\s*[:：]?\s*/g, ' ').replace(/  +/g, ' ').trim();

                        if (frag.length > best.length) best = frag;
                        if (frag.indexOf(',') > 0 && frag.length > num.length + 2) break;
                        if (/[\u4e00-\u9fff]/.test(frag)) break;
                        if (/[A-Za-z]/.test(frag) && frag.length > num.length + 2) break;
                    }
                    cur = cur.parentElement;
                }
                return best;
            }


            // === 從每個 1199 DOM 元素往上爬 DOM 層級，找 3x (BO/SRAW) 母階 ===
            // Windchill ExtJS 樹狀結構：<li> 包 <ul> 包 <li>，母階是上層 <li>

            // 輔助：從一個 DOM 元素取出所有文字（遞迴子節點），用於判斷是否含 3x
            function getAllText(el) {
                return (el.innerText || el.textContent || '').replace(/[\r\n\t]+/g, ' ').replace(/  +/g, ' ').trim();
            }

            // 輔助：檢查文字是否包含 3x 料號 + BO/SRAW 關鍵字，回傳整理後的母階描述
            function extract3xParent(text) {
                var m = text.match(/(3\d{9,})/);
                if (!m) return null;
                var upper = text.toUpperCase();
                if (upper.indexOf('SRAW') < 0 && upper.indexOf('BO ') < 0 &&
                    upper.indexOf(', BO') < 0 && upper.indexOf('RAW') < 0) return null;
                // 擷取 3x 料號及後續描述
                var m3 = text.match(/(3\d{9,}[^\x00-\x1F]*)/);
                if (!m3) return m[1];
                var result = m3[1].trim();
                var idx = result.indexOf('狀態');
                result = idx > 0 ? result.substring(0, idx) : result.substring(0, 100);
                return result.replace(/(版本|名稱)\s*[:：]?\s*/g, ' ').replace(/  +/g, ' ').trim();
            }

            // 方法 1：Ext JS Store — 找到 1199 record 後用 parentNode 往上爬
            try {
                if (window.Ext && window.Ext.ComponentQuery) {
                    var stores = [];
                    var cmpList = Ext.ComponentQuery.query('*');
                    for (var m = 0; m < cmpList.length; m++) {
                        var cmp = cmpList[m];
                        if (cmp && typeof cmp.getStore === 'function') {
                            if (cmp.isHidden && cmp.isHidden()) continue;
                            try {
                                var s = cmp.getStore();
                                if (s) stores.push(s);
                            } catch(e) {}
                        }
                    }

                    for (var si = 0; si < stores.length; si++) {
                        var store = stores[si];
                        // 嘗試遍歷 store 所有 record
                        var allRecords = [];
                        try {
                            if (store.each) {
                                store.each(function(r) { allRecords.push(r); });
                            } else if (store.getRange) {
                                allRecords = store.getRange();
                            } else if (store.data && store.data.items) {
                                allRecords = store.data.items;
                            }
                        } catch(e) { continue; }

                        for (var ri = 0; ri < allRecords.length; ri++) {
                            var rec = allRecords[ri];
                            var dataStr = '';
                            var rd = rec.data || rec.attributes || {};
                            for (var k in rd) {
                                if (typeof rd[k] === 'string') dataStr += ' ' + rd[k];
                            }
                            if (typeof rec.get === 'function') {
                                try { dataStr += ' ' + String(rec.get('number') || ''); } catch(e){}
                                try { dataStr += ' ' + String(rec.get('name') || ''); } catch(e){}
                            }
                            var plain = dataStr.replace(/<[^>]+>/g, ' ').replace(/&nbsp;/ig, ' ');
                            var mNum = plain.match(/(1199\d{6,})/);
                            if (!mNum) continue;
                            var num = mNum[1];
                            if (seen1199[num]) continue;
                            seen1199[num] = true;

                            var txt1199 = num;
                            var m1 = plain.match(/(1199\d{6,}[^\x00-\x1F]*)/);
                            if (m1) {
                                txt1199 = m1[1].trim();
                                var idx2 = txt1199.indexOf('狀態');
                                txt1199 = idx2 > 0 ? txt1199.substring(0, idx2) : txt1199.substring(0, 100);
                                txt1199 = txt1199.replace(/(版本|名稱)\s*[:：]?\s*/g, ' ').replace(/  +/g, ' ').trim();
                            }

                            // 用 parentNode 往上爬找 3x 母階
                            var parent3x = '';
                            var pNode = rec.parentNode || null;
                            for (var depth = 0; depth < 20 && pNode; depth++) {
                                var pStr = '';
                                var pd = pNode.data || pNode.attributes || {};
                                for (var pk in pd) {
                                    if (typeof pd[pk] === 'string') pStr += ' ' + pd[pk];
                                }
                                if (typeof pNode.get === 'function') {
                                    try { pStr += ' ' + String(pNode.get('number') || ''); } catch(e){}
                                    try { pStr += ' ' + String(pNode.get('name') || ''); } catch(e){}
                                }
                                var found3x = extract3xParent(pStr);
                                if (found3x) { parent3x = found3x; break; }
                                pNode = pNode.parentNode || null;
                            }

                            results.push({ num: num, text: txt1199, parentText: parent3x });
                        }
                        if (results.length > 0) return results;
                    }
                }
            } catch(e) {}

            // 方法 2：DOM 掃描 — 找到 1199 文字後往上爬 <li> 層級找 3x 母階
            var walker = document.createTreeWalker(
                document.body, NodeFilter.SHOW_TEXT, null, false);
            var tNode;

            while (tNode = walker.nextNode()) {
                var txt = tNode.textContent.trim();
                var m1199 = txt.match(/(1199\d{6,})/);
                if (!m1199) continue;
                var num = m1199[1];
                if (seen1199[num]) continue;

                var el = tNode.parentElement;
                if (!el) continue;
                var rect = el.getBoundingClientRect();
                if (rect.height === 0 && rect.width === 0) continue;

                seen1199[num] = true;
                var txt1199 = getIdentText(el, num);

                // 從 1199 元素往上爬 DOM 層級找 3x 母階
                // ExtJS tree: <li class="x-tree-node"> → <ul> → <li class="x-tree-node"> (parent)
                var parent3x = '';
                var climb = el;
                for (var depth = 0; depth < 30 && climb; depth++) {
                    climb = climb.parentElement;
                    if (!climb || climb === document.body) break;

                    // 只在 <li> 或含 tree-node class 的元素上檢查
                    var tag = climb.tagName ? climb.tagName.toUpperCase() : '';
                    var cls = climb.className || '';
                    if (tag !== 'LI' && tag !== 'TR' && cls.indexOf('tree-node') < 0 && cls.indexOf('x-grid-row') < 0) continue;

                    var parentTxt = getAllText(climb);
                    var found3x = extract3xParent(parentTxt);
                    if (found3x) {
                        parent3x = found3x;
                        break;
                    }
                }

                results.push({ num: num, text: txt1199, parentText: parent3x });
            }

            return results;
        """

        try:
            # ① 找結構搜尋框（ExtJS x-form-text，非標準 input）
            self._ql('尋找結構搜尋框（ExtJS）...', 'INFO')
            sb = None

            drv.switch_to.default_content()
            self._ql('切換到 psbIFrame...', 'INFO')
            try:
                psb = WebDriverWait(drv, 10).until(
                    EC.presence_of_element_located((By.ID, 'psbIFrame')))
                drv.switch_to.frame(psb)
                self._ql('✅ 已切換到 psbIFrame', 'OK')
            except TimeoutException:
                self._ql('❌ 找不到 psbIFrame', 'ERROR')
                return self._scan_1199()

            self._ql('定位結構搜尋框（StructureBrowserFindToolbar）...', 'INFO')
            sb = None
            for xp in [
                '//*[@id="StructureBrowserFindToolbar"]//input',
                '//*[@id="x-auto-83"]//input',
                '//div[contains(@id,"StructureBrowserFind")]//input',
            ]:
                try:
                    sb = WebDriverWait(drv, 8).until(
                        EC.presence_of_element_located((By.XPATH, xp)))
                    self._ql(f'✅ 找到結構搜尋框', 'OK')
                    break
                except TimeoutException:
                    continue

            if sb is None:
                self._ql('❌ 找不到結構搜尋框', 'ERROR')
                drv.switch_to.default_content()
                return self._scan_1199()

            # ② 清空 + 輸入 1199 + Enter
            drv.execute_script("arguments[0].value = '';", sb)
            sb.click(); time.sleep(0.3)
            sb.send_keys(Keys.CONTROL + 'a')
            sb.send_keys(Keys.DELETE)
            time.sleep(0.2)
            sb.send_keys('1199')
            time.sleep(0.5)
            sb.send_keys(Keys.RETURN)
            self._ql('輸入 1199 並按 Enter', 'OK')
            time.sleep(1.5)

            # ③ 從 psbIFrame 掃描 1199 項目，並從每個 1199 往上爬找 3x 母階
            self._ql('掃描 1199 項目，並往上追溯 3x 母階...', 'INFO')
            time.sleep(2.0)

            nodes_data = drv.execute_script(_js_scan)
            self._ql(f'JS 掃描找到 {len(nodes_data)} 個 1199 節點', 'INFO')

            # 若 JS 掃描沒結果，改用 XPath
            if not nodes_data:
                node_xp = (
                    '//*[contains(@class,"x-tree-node-el") and contains(.,"1199")]'
                    '|//span[starts-with(normalize-space(.),"1199")]'
                    '|//td[starts-with(normalize-space(.),"1199")]'
                )
                node_els = drv.find_elements(By.XPATH, node_xp)
                nodes_data = []
                for ne in node_els:
                    try:
                        txt = ne.text.strip().split('\n')[0]
                        if re.search(r'1199\d{6,}', txt):
                            m = re.search(r'(1199\d{6,})', txt)
                            nodes_data.append({'num': m.group(1), 'text': txt, 'parentText': ''})
                    except Exception:
                        continue
                self._ql(f'XPath 補掃找到 {len(nodes_data)} 個節點', 'INFO')

            def _collect(items):
                found = False
                for item in (items or []):
                    num = item.get('num', '')
                    if not num or num in seen:
                        continue
                    seen.add(num)
                    pt = item.get('parentText', '')
                    parts.append({
                        'num':        num,
                        'text':       item.get('text', num),
                        'parentText': pt,
                    })
                    self._ql(
                        f'  [{len(parts)}] {item.get("text", num)[:60]}'
                        + (f'  ← {pt[:40]}' if pt else ''),
                        'OK')
                    found = True
                return found

            _collect(nodes_data)

            if sb:
                # Enter 循環補掃：每按 Enter 跳到下一個搜尋結果
                self._ql('Enter 循環補掃開始，尋找其餘 1199 料號...', 'INFO')
                no_new_count = 0
                for _iter in range(25):
                    if not self.running:
                        break
                    sb.send_keys(Keys.RETURN)
                    time.sleep(1.2)
                    if not _collect(drv.execute_script(_js_scan)):
                        no_new_count += 1
                        if no_new_count >= 4:
                            self._ql('連續 4 次無新料號，補掃完成', 'INFO')
                            break
                    else:
                        no_new_count = 0

        except Exception as e:
            self._ql(f'結構搜尋例外：{str(e)[:80]}', 'ERROR')
            if not parts:
                parts = self._scan_1199()
        finally:
            try: drv.switch_to.default_content()
            except Exception: pass

        self._ql(f'共收集到 {len(parts)} 個不重複 1199 料號', 'OK')
        return parts


    def _read_right_panel(self, label: str) -> str:
        """
        用 JS 穿透所有 iframe，找到右側屬性面板中 label 對應的值。
        """
        drv = self.pdm_drv
        result = drv.execute_script("""
            var label = arguments[0];

            function findInDoc(doc) {
                // 找所有包含 label 文字的元素
                var all = doc.querySelectorAll('*');
                for (var el of all) {
                    var t = (el.innerText || el.textContent || '').trim();
                    // 匹配「編號:」或「編號：」
                    if (t === label + ':' || t === label + '\uff1a' || t === label) {
                        // 找同層下一個兄弟元素
                        var sib = el.nextElementSibling;
                        if (sib) {
                            var val = (sib.innerText || sib.textContent || '').trim();
                            if (val) return val;
                        }
                        // 找父元素的下一個 td
                        var parent = el.parentElement;
                        if (parent) {
                            var nextTd = parent.nextElementSibling;
                            if (nextTd) {
                                var val2 = (nextTd.innerText || nextTd.textContent || '').trim();
                                if (val2) return val2;
                            }
                        }
                    }
                }
                return null;
            }

            // 先掃主文件
            var v = findInDoc(document);
            if (v) return v;

            // 掃所有 iframe（最多 2 層）
            var frames = document.querySelectorAll('iframe');
            for (var fr of frames) {
                try {
                    var fdoc = fr.contentDocument || fr.contentWindow.document;
                    var v2 = findInDoc(fdoc);
                    if (v2) return v2;
                    // 第 2 層
                    var inner = fdoc.querySelectorAll('iframe');
                    for (var jfr of inner) {
                        try {
                            var jdoc = jfr.contentDocument || jfr.contentWindow.document;
                            var v3 = findInDoc(jdoc);
                            if (v3) return v3;
                        } catch(e) {}
                    }
                } catch(e) {}
            }
            return '';
        """, label)
        return (result or '').strip()

    # ══════════════════════════════════════════════════════════════════
    #  GeDCC (DMP) 自動化
    # ══════════════════════════════════════════════════════════════════
    def _dmp_login(self) -> bool:
        """開啟 GeDCC，等候使用者完成帳密登入"""
        drv = self.dmp_drv
        try:
            self._ql('開啟 GeDCC 登入頁...', 'INFO')
            drv.get(GEDCC_URL)
            time.sleep(2)
            # 關閉可能出現的 browser alert
            try: drv.switch_to.alert.dismiss()
            except: pass

            self._ql('⚠ 請在 GeDCC 視窗輸入帳號密碼', 'WARN')
            self._q(type='popup', sys='GeDCC (DMP)')

            # 等候登入完成（URL 轉至 private/ 即成功）
            WebDriverWait(drv, 180).until(
                lambda d: 'DMP/private' in d.current_url)
            self._ql('✅ GeDCC 登入成功', 'OK')
            return True
        except TimeoutException:
            return False
        except Exception as e:
            self._ql(f'GeDCC 登入錯誤：{e}', 'ERROR')
            return False

    def _dmp_get_link(self, pcb_num: str):
        """
        在 GeDCC 搜尋 pcb_num：
        1. 確保在首頁 → 選「料號」→ 輸入料號 → 點查詢
        2. 點擊結果列中的藍色文件名稱連結（class=w01_11px_blue）
        3. 取得文件頁的 URL 作為線路圖連結
        """
        drv  = self.dmp_drv
        wait = WebDriverWait(drv, 15)
        lnk  = None
        try:
            # ① 確保在首頁
            if 'user_index' not in drv.current_url:
                self._ql(f'  導向 GeDCC 首頁...', 'INFO')
                drv.get(GEDCC_HOME)
                time.sleep(2)

            self._ql(f'  搜尋 GeDCC：{pcb_num}', 'INFO')

            # ② 選「料號」類型下拉
            try:
                sel_el = WebDriverWait(drv, 5).until(EC.presence_of_element_located((
                    By.XPATH,
                    '//select[option[normalize-space()="料號"]]'
                    '|//select[contains(@name,"type") or contains(@id,"type")]'
                )))
                Select(sel_el).select_by_visible_text('料號')
                time.sleep(0.4)
                self._ql('  已選「料號」', 'INFO')
            except Exception:
                self._ql('  未找到料號下拉（略過）', 'INFO')

            # ③ 找搜尋輸入框（多重 fallback）
            inp = None
            for xp in [
                '//input[@name="bqKeyword"]',
                '//input[@id="bqKeyword"]',
                '//input[contains(@name,"Keyword") or contains(@id,"Keyword")]',
                '//input[@type="text" and (contains(@name,"keyword") or contains(@id,"keyword"))]',
                '//input[@type="text" and (contains(@name,"search") or contains(@id,"search"))]',
                '//form//input[@type="text"][1]',
            ]:
                try:
                    inp = WebDriverWait(drv, 3).until(
                        EC.element_to_be_clickable((By.XPATH, xp)))
                    self._ql(f'  輸入框 XPath: {xp[:50]}', 'INFO')
                    break
                except Exception:
                    continue

            if inp is None:
                self._ql('  ❌ 找不到 GeDCC 搜尋輸入框', 'ERROR')
                # Log page inputs for debugging
                try:
                    all_inp = drv.find_elements(By.XPATH, '//input[@type="text"]')
                    self._ql(f'  頁面 text input 數量：{len(all_inp)}', 'INFO')
                    for ai in all_inp[:3]:
                        self._ql(f'    name={ai.get_attribute("name")} id={ai.get_attribute("id")}', 'INFO')
                except Exception:
                    pass
                return None

            drv.execute_script("arguments[0].value = '';", inp)
            inp.click()
            time.sleep(0.2)
            inp.send_keys(pcb_num)

            # ④ 點擊查詢按鈕
            try:
                btn = WebDriverWait(drv, 5).until(EC.element_to_be_clickable((By.XPATH,
                    '//input[@value="查詢"]|//button[contains(.,"查詢")]'
                    '|//a[contains(.,"查詢")]|//input[@type="submit"]'
                )))
                drv.execute_script("arguments[0].click();", btn)
                self._ql('  已點查詢', 'INFO')
            except Exception:
                inp.send_keys(Keys.RETURN)
                self._ql('  Enter 送出查詢', 'INFO')
            time.sleep(2.5)

            # ⑤ 找藍色文件連結（GeDCC 結果可能在 frameset 子 frame 內）
            self._ql('  等待搜尋結果（含 frame 掃描）...', 'INFO')
            time.sleep(2.0)

            # 收集所有 frame/iframe（先回到 default content）
            drv.switch_to.default_content()
            all_frames = drv.find_elements(By.XPATH, '//frame|//iframe')
            self._ql(f'  frame 數量：{len(all_frames)}', 'INFO')

            # 依序嘗試：default content → frame 0 → frame 1 → ...
            res_found = None
            checked_frame = -1
            for fi in range(-1, len(all_frames)):
                try:
                    drv.switch_to.default_content()
                    if fi >= 0:
                        drv.switch_to.frame(fi)
                    candidates = drv.find_elements(By.XPATH,
                        '//a[contains(@class,"w01_11px_blue")]'
                        '|//a[contains(@href,"DocumentBrowse")]'
                        '|//a[contains(@href,"userBrowse")]'
                    )
                    if candidates:
                        res_found = candidates[0]
                        checked_frame = fi
                        self._ql(f'  ✅ 找到藍色連結 (frame={fi})：{res_found.text[:40]}', 'OK')
                        break
                except Exception:
                    drv.switch_to.default_content()
                    continue

            if res_found is None:
                # Debug: log links in each frame
                self._ql('  找不到藍色連結，輸出各 frame 連結...', 'WARN')
                for fi in range(-1, len(all_frames)):
                    try:
                        drv.switch_to.default_content()
                        if fi >= 0:
                            drv.switch_to.frame(fi)
                        flinks = drv.find_elements(By.XPATH, '//a[@href]')
                        self._ql(f'  frame={fi} 連結數：{len(flinks)}', 'INFO')
                        for al in flinks[:4]:
                            href = al.get_attribute('href') or ''
                            cls  = al.get_attribute('class') or ''
                            self._ql(f'    [{cls[:15]}] {href[:55]}', 'INFO')
                    except Exception:
                        pass
                drv.switch_to.default_content()
                self._ql(f'  找不到 {pcb_num} 的 GeDCC 文件', 'WARN')
            else:
                # 點擊藍色連結，進入文件頁面，取得 URL
                raw_href = res_found.get_attribute('href') or ''
                pre_lnk  = raw_href if raw_href.startswith('http') \
                           else 'http://global-gedcc.moxa.com' + raw_href
                drv.execute_script("arguments[0].click();", res_found)
                time.sleep(2.5)
                drv.switch_to.default_content()
                lnk = drv.current_url
                # 若 URL 沒變（frameset 內部換頁），改用 href
                if 'user_index' in lnk or lnk == pre_lnk:
                    lnk = pre_lnk
                self._ql(f'  文件 URL：{lnk[:70]}', 'OK')

        except Exception as e:
            self._ql(f'  DMP 例外 [{pcb_num}]：{str(e)[:80]}', 'WARN')
        finally:
            try:
                drv.switch_to.default_content()
                if 'user_index' not in drv.current_url:
                    drv.get(GEDCC_HOME)
                    time.sleep(1)
            except Exception:
                pass
        return lnk


# ══════════════════════════════════════════════════════════════════════
#  進入點
# ══════════════════════════════════════════════════════════════════════
def main():
    root = tk.Tk()
    app  = App(root)

    # 置中視窗
    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    ww, wh = root.winfo_width(), root.winfo_height()
    root.geometry(f'{ww}x{wh}+{(sw-ww)//2}+{(sh-wh)//2}')

    root.mainloop()


if __name__ == '__main__':
    main()
