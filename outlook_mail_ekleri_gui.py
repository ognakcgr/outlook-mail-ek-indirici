# -*- coding: utf-8 -*-
"""
Outlook Mail Arama ve Ek İndirme Uygulaması
============================================

Outlook'ta seçilen klasörde filtreleme yaparak arama yapar
ve ekleri (Excel, PDF veya tümü) indirir.

Gereksinimler:
    pip install pywin32

Kullanım:
    python outlook_mail_ekleri_gui.py
"""

from __future__ import annotations

import ctypes
import json
import os
import queue
import sys
import threading
import time
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk
from typing import Optional, List, Dict, Any

try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

APP_VERSION = "1.2"
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outlook_ayarlar.json")

# --------------------------------------------------------------------------
# Tema - Mevcut uygulamayla aynı koyu tema
# --------------------------------------------------------------------------
class T:
    BG = "#0b0e13"
    BG_SOFT = "#0f141b"
    CARD = "#141a23"
    CARD_HI = "#18202b"
    BORDER = "#222c39"
    BORDER_HI = "#2e3b4d"
    TEXT = "#e8eef7"
    TEXT_SOFT = "#a7b4c6"
    MUTED = "#6f7f94"
    ACCENT = "#3b82f6"
    ACCENT_HI = "#5b9bff"
    ACCENT_DIM = "#1d3a6b"
    OK = "#2ecc71"
    OK_DIM = "#123524"
    WARN = "#f0b429"
    WARN_DIM = "#3a2b0a"
    ERR = "#f05252"
    ERR_DIM = "#3a1414"
    LOG_BG = "#0a0d12"

    FONT = "Segoe UI"
    FONT_SEMI = "Segoe UI Semibold"
    MONO = "Consolas"


def ui(size: int, weight: str = "normal") -> tuple:
    return (T.FONT_SEMI if weight == "bold" else T.FONT, size)


# --------------------------------------------------------------------------
# Yardımcı çizim fonksiyonları
# --------------------------------------------------------------------------
def round_rect(canvas: tk.Canvas, x1, y1, x2, y2, r, **kwargs):
    r = max(1, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    points = [
        x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
        x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
        x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)


class RoundedButton(tk.Canvas):
    """Yumuşak köşeli, hover efektli düğme."""

    def __init__(self, master, text: str, command=None, kind: str = "primary",
                 width: int = 200, height: int = 44, radius: int = 10, font_size: int = 10, **kw):
        super().__init__(master, width=width, height=height, bg=kw.pop("bg", T.CARD),
                         highlightthickness=0, bd=0, cursor="hand2", **kw)
        self._text = text
        self._command = command
        self._kind = kind
        self._radius = radius
        self._bw, self._bh = width, height
        self._enabled = True
        self._hover = False
        self._pressed = False
        self._font = ui(font_size, "bold")
        self._shape = None
        self._label = None
        self._draw()
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Configure>", self._on_configure)

    def _palette(self):
        if not self._enabled:
            fill = T.CARD if self._kind == "ghost" else T.CARD_HI
            return fill, "#4c5a6d", T.BORDER
        if self._kind == "primary":
            fill = T.ACCENT_HI if self._hover else T.ACCENT
            if self._pressed:
                fill = "#2f6fd0"
            return fill, "#ffffff", fill
        if self._kind == "danger":
            fill = "#ff6b6b" if self._hover else T.ERR
            if self._pressed:
                fill = "#c93f3f"
            return fill, "#ffffff", fill
        if self._kind == "ghost":
            if self._hover:
                return T.CARD_HI, T.TEXT, T.BORDER_HI
            return T.CARD, T.TEXT, T.BORDER_HI
        fill = T.BORDER_HI if self._hover else T.CARD_HI
        return fill, T.TEXT, T.BORDER_HI if self._hover else T.BORDER

    def _draw(self):
        self.delete("all")
        fill, fg, outline = self._palette()
        self._shape = round_rect(self, 1, 1, self._bw - 1, self._bh - 1, self._radius,
                                 fill=fill, outline=outline)
        self._label = self.create_text(self._bw / 2, self._bh / 2 + 1, text=self._text,
                                       fill=fg, font=self._font)

    def _on_configure(self, event):
        self._bw, self._bh = event.width, event.height
        self._draw()

    def _on_enter(self, _):
        self._hover = True
        self._draw()

    def _on_leave(self, _):
        self._hover = False
        self._pressed = False
        self._draw()

    def _on_press(self, _):
        if not self._enabled:
            return
        self._pressed = True
        self._draw()

    def _on_release(self, _):
        if not self._enabled:
            return
        was = self._pressed
        self._pressed = False
        self._draw()
        if was and self._command:
            self._command()

    def set_enabled(self, enabled: bool) -> None:
        self._enabled = enabled
        self.configure(cursor="hand2" if enabled else "arrow")
        self._draw()

    def set_text(self, text: str) -> None:
        self._text = text
        self._draw()


class Switch(tk.Frame):
    """Modern aç/kapa anahtarı + etiket."""

    def __init__(self, master, text: str, value: bool = False, command=None, bg: str = T.CARD, **kw):
        super().__init__(master, bg=bg, **kw)
        self._value = value
        self._command = command
        self._bgcolor = bg
        self._onvar = tk.BooleanVar(value=value)

        self.canvas = tk.Canvas(self, width=48, height=26, bg=bg, highlightthickness=0, cursor="hand2")
        self.canvas.pack(side="left", pady=2)
        self.label = tk.Label(self, text=text, bg=bg, fg=T.TEXT_SOFT, font=ui(9), cursor="hand2")
        self.label.pack(side="left", padx=(10, 0))

        for w in (self.canvas, self.label):
            w.bind("<Button-1>", self._toggle)
        self._draw()

    def _draw(self):
        self.canvas.delete("all")
        track = T.ACCENT if self._value else T.BORDER_HI
        round_rect(self.canvas, 1, 2, 47, 24, 11, fill=track, outline=track)
        x = 35 if self._value else 13
        self.canvas.create_oval(x - 9, 4, x + 9, 20, fill="#ffffff", outline="")

    def _toggle(self, _=None):
        self._value = not self._value
        self._onvar.set(self._value)
        self._draw()
        if self._command:
            self._command(self._value)

    def get(self) -> bool:
        return self._value

    def set(self, value: bool) -> None:
        self._value = value
        self._draw()


class Card(tk.Frame):
    """Başlıklı, ince çerçeveli kart."""

    def __init__(self, master, title: str = "", subtitle: str = "", **kw):
        super().__init__(master, bg=T.CARD, highlightthickness=1, highlightbackground=T.BORDER, **kw)
        self.body = tk.Frame(self, bg=T.CARD)
        self.body.pack(fill="both", expand=True, padx=18, pady=14)

        if title:
            head = tk.Frame(self.body, bg=T.CARD)
            head.pack(fill="x", pady=(0, 12))
            tk.Label(head, text=title, bg=T.CARD, fg=T.TEXT, font=ui(11, "bold")).pack(side="left")
            if subtitle:
                tk.Label(head, text=subtitle, bg=T.CARD, fg=T.MUTED, font=ui(9)).pack(side="left", padx=(10, 0), pady=(2, 0))


class ProgressBar(tk.Canvas):
    def __init__(self, master, height: int = 8, bg: str = T.CARD, **kw):
        super().__init__(master, height=height, bg=bg, highlightthickness=0, **kw)
        self._mode = "idle"
        self._value = 0.0
        self._offset = 0.0
        self._job = None
        self.bind("<Configure>", lambda e: self.redraw())
        self.after(50, self.redraw)

    def set_mode(self, mode: str) -> None:
        self._mode = mode
        if mode == "indeterminate":
            self._animate()
        self.redraw()

    def set_value(self, value: float) -> None:
        self._value = max(0.0, min(1.0, value))
        self.redraw()

    def redraw(self):
        if not self.winfo_exists():
            return
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 1:
            return
        round_rect(self, 0, 0, w, h, h / 2, fill=T.BG_SOFT, outline=T.BG_SOFT)
        if self._mode == "indeterminate":
            seg = max(60, w * 0.28)
            x = -seg + (w + seg) * self._offset
            round_rect(self, x, 0, x + seg, h, h / 2, fill=T.ACCENT, outline=T.ACCENT)
        elif self._mode == "determinate" and self._value > 0:
            round_rect(self, 0, 0, max(h, w * self._value), h, h / 2, fill=T.ACCENT, outline=T.ACCENT)

    def _animate(self):
        if not self.winfo_exists() or self._mode != "indeterminate":
            return
        self._offset += 0.012
        if self._offset > 1.0:
            self._offset = 0.0
        self.redraw()
        self._job = self.after(25, self._animate)


# --------------------------------------------------------------------------
# Outlook Klasör Seçim Dialogu
# --------------------------------------------------------------------------
class FolderSelectDialog(tk.Toplevel):
    """Outlook klasör hiyerarşisinden klasör seçme dialogu."""

    def __init__(self, master, namespace, current_folder_entry_id: str = ""):
        super().__init__(master)
        self.title("Outlook Klasörü Seç")
        self.configure(bg=T.BG)
        self.geometry("520x650")
        self.minsize(480, 500)
        self.resizable(True, True)
        self.transient(master)
        self.namespace = namespace
        self.result = None
        self.result_entry_id = current_folder_entry_id
        self.folder_map = {}  # tree item id -> (folder, entry_id)
        self._com_initialized = False
        self._master = master  # Store master for centering
        self._build_ui()
        self._load_folders()
        self._select_current()
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.wait_window()

    def _on_close(self):
        if self._com_initialized:
            pythoncom.CoUninitialize()
        self.destroy()

    def _build_ui(self):
        wrap = tk.Frame(self, bg=T.BG, padx=16, pady=16)
        wrap.pack(fill="both", expand=True)
        wrap.rowconfigure(1, weight=1)
        wrap.columnconfigure(0, weight=1)

        # Başlık
        tk.Label(wrap, text="Arama yapılacak Outlook klasörünü seçin",
                 bg=T.BG, fg=T.TEXT, font=ui(12, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 12))

        # Treeview
        tree_frame = tk.Frame(wrap, bg=T.CARD, highlightthickness=1, highlightbackground=T.BORDER)
        tree_frame.grid(row=1, column=0, sticky="nsew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tree_frame, show="tree", selectmode="browse")
        self.tree.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

        # Scrollbar
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        vsb.grid(row=0, column=1, sticky="ns", pady=4)
        self.tree.configure(yscrollcommand=vsb.set)

        # Style
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview",
                        background=T.CARD,
                        foreground=T.TEXT,
                        fieldbackground=T.CARD,
                        borderwidth=0,
                        font=ui(9))
        style.configure("Treeview.Heading",
                        background=T.CARD_HI,
                        foreground=T.TEXT,
                        font=ui(9, "bold"))
        style.map("Treeview",
                  background=[("selected", T.ACCENT_DIM)],
                  foreground=[("selected", T.ACCENT_HI)])

        # Butonlar
        btn_frame = tk.Frame(wrap, bg=T.BG)
        btn_frame.grid(row=2, column=0, sticky="ew", pady=(16, 0))
        RoundedButton(btn_frame, "Vazgeç", self._on_close, kind="secondary", width=110, height=38, font_size=10, bg=T.BG).pack(side="right", padx=(8, 0))
        RoundedButton(btn_frame, "Seç", self._on_select, kind="primary", width=110, height=38, font_size=10, bg=T.BG).pack(side="right")

        # Çift tıkla seç
        self.tree.bind("<Double-1>", lambda e: self._on_select())

        self.update_idletasks()
        self._center(self._master)

    def _center(self, master):
        w, h = self.winfo_width(), self.winfo_height()
        x = master.winfo_rootx() + (master.winfo_width() - w) // 2
        y = master.winfo_rooty() + (master.winfo_height() - h) // 3
        self.geometry(f"+{max(0, x)}+{max(0, y)}")

    def _load_folders(self):
        """Tüm Outlook klasörlerini ağaç yapısına yükle."""
        try:
            pythoncom.CoInitialize()
            self._com_initialized = True

            # Kök depoları (Store'lar) al
            stores = self.namespace.Stores
            for i in range(1, stores.Count + 1):
                try:
                    store = stores.Item(i)
                    root_folder = store.GetRootFolder()
                    self._add_folder_to_tree(root_folder, "")
                except Exception as e:
                    print(f"Store yükleme hatası: {e}")
                    continue
        except Exception as e:
            print(f"Klasör yükleme hatası: {e}")
            # Hata durumunda boş tree kalır, kullanıcı kapatır

    def _add_folder_to_tree(self, folder, parent_id: str):
        """Klasörü ve alt klasörlerini ağaca ekle."""
        try:
            name = folder.Name
            entry_id = folder.EntryID
            # Tüm klasörleri göster (filtreleme yok - kullanıcı karar versin)
            item_id = self.tree.insert(parent_id, "end", text=f"📁 {name}", open=False)
            self.folder_map[item_id] = (folder, entry_id)

            # Alt klasörler
            try:
                for j in range(1, folder.Folders.Count + 1):
                    subfolder = folder.Folders.Item(j)
                    self._add_folder_to_tree(subfolder, item_id)
            except Exception:
                pass
        except Exception as e:
            print(f"Klasör ekleme hatası ({getattr(folder, 'Name', '?')}): {e}")
            pass

    def _select_current(self):
        """Mevcut seçili klasörü ağaçta vurgula."""
        if not self.result_entry_id:
            return
        for item_id, (_, entry_id) in self.folder_map.items():
            if entry_id == self.result_entry_id:
                self.tree.selection_set(item_id)
                self.tree.see(item_id)
                # Ebeveynleri aç
                parent = self.tree.parent(item_id)
                while parent:
                    self.tree.item(parent, open=True)
                    parent = self.tree.parent(parent)
                break

    def _on_select(self):
        selection = self.tree.selection()
        if selection:
            item_id = selection[0]
            folder, entry_id = self.folder_map[item_id]
            self.result = folder
            self.result_entry_id = entry_id
        if self._com_initialized:
            pythoncom.CoUninitialize()
            self._com_initialized = False
        self.destroy()


# --------------------------------------------------------------------------
# Outlook İşlemci
# --------------------------------------------------------------------------
class OutlookProcessor:
    """Outlook bağlantısı ve mail/ek işlemlerini yönetir."""

    def __init__(self, reporter):
        self.reporter = reporter
        self.outlook = None
        self.namespace = None

    def connect(self) -> bool:
        """Outlook'a bağlanır."""
        if win32com is None:
            self.reporter.log("pywin32 kütüphanesi yüklü değil. 'pip install pywin32' çalıştırın.", "error")
            return False

        try:
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.reporter.log("Outlook'a başarıyla bağlanıldı.", "ok")
            return True
        except Exception as e:
            self.reporter.log(f"Outlook bağlantı hatası: {e}", "error")
            return False

    def get_sent_folder(self):
        """Gönderilenler klasörünü döndürür (varsayılan)."""
        try:
            return self.namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
        except Exception as e:
            self.reporter.log(f"Gönderilenler klasörü alınamadı: {e}", "error")
            return None

    def get_folder_by_entry_id(self, entry_id: str):
        """EntryID ile klasör getir."""
        try:
            return self.namespace.GetFolderFromID(entry_id)
        except Exception as e:
            self.reporter.log(f"Klasör alınamadı (EntryID): {e}", "error")
            return None

    def search_mails(self, folder, filters: Dict[str, Any]) -> List:
        """Filtrelere göre mailleri arar."""
        mails = []
        try:
            items = folder.Items
            # Sıralama: SentOn varsa (gönderilenler), ReceivedTime varsa (gelen kutusu)
            sort_field = "[SentOn]" if "Sent" in str(folder.Name).lower() or folder.DefaultItemType == 0 else "[ReceivedTime]"
            try:
                items.Sort(sort_field, True)  # En yeni önce
            except Exception:
                pass

            # Filtre string'i oluştur - SADECE TARİH FİLTRESİ Restrict ile
            date_filter_parts = []

            # Tarih aralığı - hem SentOn hem ReceivedTime dene
            if filters.get("date_from"):
                dt_from = filters["date_from"].strftime("%m/%d/%Y %H:%M %p")
                date_filter_parts.append(f"[SentOn] >= '{dt_from}'")
                date_filter_parts.append(f"[ReceivedTime] >= '{dt_from}'")
            if filters.get("date_to"):
                dt_to = filters["date_to"].strftime("%m/%d/%Y %H:%M %p")
                date_filter_parts.append(f"[SentOn] <= '{dt_to}'")
                date_filter_parts.append(f"[ReceivedTime] <= '{dt_to}'")

            # Metin filtreleri (manuel uygulanacak)
            subject_filter = filters.get("subject", "").lower() if filters.get("subject") else None
            to_filter = filters.get("to", "").lower() if filters.get("to") else None
            sender_filter = filters.get("sender", "").lower() if filters.get("sender") else None

            # Debug: filtre değerlerini logla
            self.reporter.log(f"Metin filtreleri - Konu: '{subject_filter}', Alıcı: '{to_filter}', Gönderen: '{sender_filter}'", "debug")
            if not subject_filter and not to_filter and not sender_filter:
                self.reporter.log("UYARI: Hiçbir metin filtresi girilmemiş! Tüm mailler geçecek.", "warning")

            # Sadece tarih filtresi ile Restrict uygula
            restricted_items = items
            if date_filter_parts:
                date_filter_str = " AND ".join(date_filter_parts)
                self.reporter.log(f"Tarih filtresi (Restrict): {date_filter_str}", "debug")
                try:
                    restricted_items = items.Restrict(date_filter_str)
                    try:
                        restricted_count = restricted_items.Count
                        self.reporter.log(f"Tarih filtresi sonrası öğe sayısı: {restricted_count}", "debug")
                    except Exception:
                        pass
                except Exception as e:
                    self.reporter.log(f"Tarih Restrict başarısız ({e}), tüm öğeler taranacak", "warning")
                    restricted_items = items
            else:
                self.reporter.log("Tarih filtresi yok, tüm öğeler taranacak", "debug")

            # Mailleri topla - METİN FİLTRELERİ MANUEL UYGULA
            count = 0
            max_mails = filters.get("max_mails", 500)
            checked = 0
            skipped_by_text_filter = 0
            skipped_not_mail = 0
            sample_logged = 0
            for item in restricted_items:
                checked += 1
                if count >= max_mails:
                    break
                try:
                    if item.Class == 43:  # olMail
                        # İlk 3 öğeyi örnek olarak logla
                        if sample_logged < 3:
                            sample_logged += 1
                            self.reporter.log(f"Örnek {sample_logged}: Konu='{(item.Subject or '')[:50]}', Alıcı='{(item.To or '')[:30]}', Gönderen='{(item.SenderName or '')[:30]}'", "debug")

                        # Metin filtrelerini manuel uygula
                        if subject_filter or to_filter or sender_filter:
                            subject = (item.Subject or "").lower()
                            to = (item.To or "").lower()
                            sender = (item.SenderName or "").lower()

                            match = True
                            if subject_filter and subject_filter not in subject:
                                match = False
                            if to_filter and to_filter not in to:
                                match = False
                            if sender_filter and sender_filter not in sender:
                                match = False

                            if not match:
                                skipped_by_text_filter += 1
                                if skipped_by_text_filter <= 3:
                                    self.reporter.log(f"Elendi: Konu='{(item.Subject or '')[:50]}' (filtre: konu='{subject_filter}', to='{to_filter}', sender='{sender_filter}')", "debug")
                                continue

                        mails.append(item)
                        count += 1
                    else:
                        skipped_not_mail += 1
                except Exception as e:
                    self.reporter.debug(f"Mail işleme hatası: {e}")
                    continue

            self.reporter.log(f"Taranan: {checked}, Mail öğesi: {checked - skipped_not_mail}, Metin filtresi elenen: {skipped_by_text_filter}, Bulunan: {len(mails)}", "debug")

            self.reporter.log(f"{len(mails)} mail bulundu ({folder.Name}).", "ok")
            return mails

        except Exception as e:
            self.reporter.log(f"Arama hatası: {e}", "error")
            return []

    def get_attachments_info(self, mail) -> List[Dict]:
        """Mailin ek bilgilerini döndürür."""
        attachments = []
        try:
            for att in mail.Attachments:
                if att.Type == 1:  # olByValue = normal dosya eki
                    attachments.append({
                        "filename": att.FileName,
                        "size": att.Size,
                        "index": att.Index,
                        "attachment": att
                    })
        except Exception as e:
            self.reporter.debug(f"Ek bilgisi alma hatası: {e}")
        return attachments

    def filter_attachments(self, attachments: List[Dict], filter_type: str) -> List[Dict]:
        """Ekleri dosya türüne göre filtreler."""
        if filter_type == "all":
            return attachments

        filtered = []
        for att in attachments:
            ext = os.path.splitext(att["filename"])[1].lower()
            if filter_type == "excel" and ext in [".xls", ".xlsx", ".xlsm", ".xlsb", ".csv"]:
                filtered.append(att)
            elif filter_type == "pdf" and ext == ".pdf":
                filtered.append(att)
        return filtered

    def download_attachments(self, mails: List, download_folder: str, attachment_filter: str,
                             create_subfolders: bool, progress_callback=None) -> Dict[str, int]:
        """Ekleri indirir."""
        stats = {"total": 0, "downloaded": 0, "skipped": 0, "errors": 0}
        pythoncom.CoInitialize()

        try:
            for i, mail in enumerate(mails):
                if progress_callback:
                    progress_callback(i + 1, len(mails), mail.Subject or "Konu yok")

                try:
                    attachments = self.get_attachments_info(mail)
                    filtered_atts = self.filter_attachments(attachments, attachment_filter)

                    if not filtered_atts:
                        continue

                    # Alt klasör oluştur (tarih + konu)
                    if create_subfolders:
                        # SentOn veya ReceivedTime
                        mail_date = getattr(mail, 'SentOn', None) or getattr(mail, 'ReceivedTime', None)
                        if hasattr(mail_date, 'strftime'):
                            date_str = mail_date.strftime("%Y-%m-%d")
                        else:
                            date_str = str(mail_date)[:10] if mail_date else "tarih_yok"
                        subject_clean = "".join(c for c in (mail.Subject or "KonuYok") if c.isalnum() or c in " -_")[:50]
                        subfolder = os.path.join(download_folder, f"{date_str}_{subject_clean}")
                    else:
                        subfolder = download_folder

                    os.makedirs(subfolder, exist_ok=True)

                    for att_info in filtered_atts:
                        stats["total"] += 1
                        try:
                            att = att_info["attachment"]
                            filename = att_info["filename"]
                            filepath = os.path.join(subfolder, filename)

                            # Aynı isimde dosya varsa yeniden adlandır
                            base, ext = os.path.splitext(filename)
                            counter = 1
                            while os.path.exists(filepath):
                                filepath = os.path.join(subfolder, f"{base}_{counter}{ext}")
                                counter += 1

                            att.SaveAsFile(filepath)
                            stats["downloaded"] += 1
                            self.reporter.log(f"İndirildi: {os.path.basename(filepath)}", "ok")
                        except Exception as e:
                            stats["errors"] += 1
                            self.reporter.log(f"İndirme hatası ({att_info['filename']}): {e}", "error")

                except Exception as e:
                    self.reporter.log(f"Mail işleme hatası: {e}", "error")
                    stats["errors"] += 1

        finally:
            pythoncom.CoUninitialize()

        return stats


# --------------------------------------------------------------------------
# Bildirim (Reporter) - GUI ile iletişim
# --------------------------------------------------------------------------
class GuiReporter:
    def __init__(self, events: queue.Queue, cancel_event: threading.Event, verbose: bool = False):
        self.events = events
        self.cancel_event = cancel_event
        self.verbose = verbose

    def log(self, msg: str, level: str = "info") -> None:
        self.events.put(("log", level, msg))

    def debug(self, msg: str) -> None:
        if self.verbose:
            self.events.put(("log", "debug", msg))

    def progress(self, current: int, total: int, subject: str = "") -> None:
        self.events.put(("progress", current, total, subject))

    def finished(self, stats: Dict) -> None:
        self.events.put(("finished", stats))

    def error(self, msg: str) -> None:
        self.events.put(("error", msg))

    def cancelled(self) -> bool:
        return self.cancel_event.is_set()


# --------------------------------------------------------------------------
# Tarih seçici yardımcıları
# --------------------------------------------------------------------------
class DateEntry(tk.Frame):
    """Tarih ve saat girişi için bileşik widget."""

    def __init__(self, master, label: str, default_days_ago: int = 0, **kw):
        super().__init__(master, bg=T.CARD, **kw)
        self.var_date = tk.StringVar()
        self.var_time = tk.StringVar()

        default_date = datetime.now() - timedelta(days=default_days_ago)
        self.var_date.set(default_date.strftime("%d.%m.%Y"))
        self.var_time.set("00:00" if default_days_ago > 0 else "23:59")

        tk.Label(self, text=label, bg=T.CARD, fg=T.MUTED, font=ui(9)).pack(anchor="w")

        row = tk.Frame(self, bg=T.CARD)
        row.pack(fill="x", pady=(4, 0))

        date_frame = tk.Frame(row, bg=T.BG_SOFT, highlightthickness=1, highlightbackground=T.BORDER)
        date_frame.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entry_date = tk.Entry(date_frame, textvariable=self.var_date, width=12,
                                   bg=T.BG_SOFT, fg=T.TEXT, insertbackground=T.TEXT,
                                   relief="flat", font=ui(9), bd=0, justify="center")
        self.entry_date.pack(padx=10, pady=6)

        time_frame = tk.Frame(row, bg=T.BG_SOFT, highlightthickness=1, highlightbackground=T.BORDER)
        time_frame.pack(side="left", fill="x", expand=True)
        self.entry_time = tk.Entry(time_frame, textvariable=self.var_time, width=8,
                                   bg=T.BG_SOFT, fg=T.TEXT, insertbackground=T.TEXT,
                                   relief="flat", font=ui(9), bd=0, justify="center")
        self.entry_time.pack(padx=10, pady=6)

    def get_datetime(self) -> Optional[datetime]:
        try:
            date_str = self.var_date.get().strip()
            time_str = self.var_time.get().strip()
            return datetime.strptime(f"{date_str} {time_str}", "%d.%m.%Y %H:%M")
        except Exception:
            return None

    def set_datetime(self, dt: datetime) -> None:
        self.var_date.set(dt.strftime("%d.%m.%Y"))
        self.var_time.set(dt.strftime("%H:%M"))


# --------------------------------------------------------------------------
# Ek türü seçim dialogu
# --------------------------------------------------------------------------
class AttachmentFilterDialog(tk.Toplevel):
    def __init__(self, master, current_filter: str):
        super().__init__(master)
        self.title("Ek Türü Seç")
        self.configure(bg=T.BG)
        self.resizable(False, False)
        self.transient(master)
        self.result = current_filter
        self.applied = False
        self._master = master

        wrap = tk.Frame(self, bg=T.BG, padx=22, pady=20)
        wrap.pack(fill="both", expand=True)
        tk.Label(wrap, text="İndirilecek ek türünü seçin", bg=T.BG, fg=T.TEXT, font=ui(13, "bold")).pack(anchor="w")
        tk.Label(wrap, text="Seçim, bulunan maillerdeki ekler üzerinde filtre uygular.", bg=T.BG, fg=T.MUTED, font=ui(9)).pack(anchor="w", pady=(2, 14))

        self.filter_var = tk.StringVar(value=current_filter)

        options = [
            ("all", "📎 Tüm Ekler", "Tüm dosya türlerindeki ekleri indir"),
            ("excel", "📊 Excel Dosyaları", ".xls, .xlsx, .xlsm, .xlsb, .csv"),
            ("pdf", "📄 PDF Dosyaları", ".pdf dosyaları"),
        ]

        for value, label, desc in options:
            rb_frame = tk.Frame(wrap, bg=T.CARD, highlightthickness=1, highlightbackground=T.BORDER)
            rb_frame.pack(fill="x", pady=4)
            rb = tk.Radiobutton(rb_frame, text=label, variable=self.filter_var, value=value,
                                bg=T.CARD, fg=T.TEXT, font=ui(10, "bold"),
                                activebackground=T.CARD, activeforeground=T.TEXT,
                                selectcolor=T.ACCENT_DIM, bd=0, anchor="w", padx=12, pady=8)
            rb.pack(fill="x")
            tk.Label(rb_frame, text=desc, bg=T.CARD, fg=T.MUTED, font=ui(8), anchor="w").pack(fill="x", padx=30)

        actions = tk.Frame(wrap, bg=T.BG)
        actions.pack(fill="x", pady=(16, 0))
        RoundedButton(actions, "Vazgeç", self.destroy, kind="secondary", width=110, height=38, font_size=10, bg=T.BG).pack(side="right", padx=(8, 0))
        RoundedButton(actions, "Uygula", self._apply, kind="primary", width=110, height=38, font_size=10, bg=T.BG).pack(side="right")

        self.bind("<Escape>", lambda e: self.destroy())
        self.update_idletasks()
        self._center(self._master)
        self.grab_set()

    def _apply(self):
        self.result = self.filter_var.get()
        self.applied = True
        self.destroy()

    def _center(self, master):
        w, h = self.winfo_width(), self.winfo_height()
        x = master.winfo_rootx() + (master.winfo_width() - w) // 2
        y = master.winfo_rooty() + (master.winfo_height() - h) // 3
        self.geometry(f"+{max(0, x)}+{max(0, y)}")


# --------------------------------------------------------------------------
# Ana Pencere
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook Mail Arama ve Ek İndirme")
        self.configure(bg=T.BG)
        self.minsize(1050, 780)
        self.geometry("1150x850")
        self._center_window(1150, 850)

        self.events: queue.Queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker: Optional[threading.Thread] = None
        self.started_at: Optional[float] = None
        self._running = False
        self.processor = None
        self.namespace = None  # Outlook namespace (klasör seçimi için)

        # Ayarlar
        self.download_folder = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "Downloads", "Outlook_Ekler"))
        self.subject_filter = tk.StringVar()
        self.to_filter = tk.StringVar()
        self.sender_filter = tk.StringVar()
        self.max_mails = tk.IntVar(value=200)
        self.attachment_filter = tk.StringVar(value="all")
        self.create_subfolders = tk.BooleanVar(value=True)
        self.open_folder_after = tk.BooleanVar(value=True)
        self.debug_mode = tk.BooleanVar(value=False)
        # Klasör seçimi ayarları
        self.selected_folder_entry_id = tk.StringVar(value="")
        self.selected_folder_name = tk.StringVar(value="Gönderilenler (Varsayılan)")

        self._load_settings()
        self._build()
        self._append_log("Hazır. Klasör seçin, filtreleri ayarlayıp 'Ara ve İndir' butonuna basın.", "info")
        self.after(80, self._drain_events)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _center_window(self, w: int, h: int) -> None:
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{max(0, (sw - w) // 2)}+{max(0, (sh - h) // 3)}")

    def _on_close(self) -> None:
        if self._running:
            if not messagebox.askyesno("İşlem sürüyor", "İndirme işlemi devam ediyor. İptal edip çıkmak istiyor musunuz?"):
                return
            self.cancel_event.set()
        self._save_settings()
        self.destroy()

    def _build(self) -> None:
        root = tk.Frame(self, bg=T.BG, padx=22, pady=16)
        root.pack(fill="both", expand=True)

        self._build_header(root)

        body = tk.Frame(root, bg=T.BG)
        body.pack(fill="both", expand=True, pady=(16, 0))
        body.columnconfigure(0, weight=0, minsize=460)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        # SOL PANEL - Kaydırılabilir (Canvas + Scrollbar)
        left_outer = tk.Frame(body, bg=T.BG)
        left_outer.grid(row=0, column=0, sticky="nsew", padx=(0, 16))
        left_outer.rowconfigure(0, weight=1)
        left_outer.columnconfigure(0, weight=1)

        left_canvas = tk.Canvas(left_outer, bg=T.BG, highlightthickness=0, bd=0)
        left_canvas.grid(row=0, column=0, sticky="nsew")

        left_scrollbar = tk.Scrollbar(left_outer, orient="vertical", command=left_canvas.yview,
                                       bg=T.CARD, troughcolor=T.BG, activebackground=T.BORDER_HI,
                                       bd=0, width=10, relief="flat")
        left_scrollbar.grid(row=0, column=1, sticky="ns")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)

        left = tk.Frame(left_canvas, bg=T.BG)
        left_canvas.create_window((0, 0), window=left, anchor="nw", tags="left_frame")

        def _on_left_configure(event):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))
            left_canvas.itemconfig("left_frame", width=event.width)
        left.bind("<Configure>", _on_left_configure)
        left_canvas.bind("<Configure>", lambda e: left_canvas.itemconfig("left_frame", width=e.width))

        # Mouse wheel scroll - sadece canvas üzerindeyken çalışsın
        def _on_mousewheel(event):
            left_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        def _bind_wheel(_):
            left_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        def _unbind_wheel(_):
            left_canvas.unbind_all("<MouseWheel>")
        left_canvas.bind("<Enter>", _bind_wheel)
        left_canvas.bind("<Leave>", _unbind_wheel)

        self._build_filters(left)
        self._build_progress(left)

        right = tk.Frame(body, bg=T.BG)
        right.grid(row=0, column=1, sticky="nsew")
        self._build_log(right)
        self._build_summary(right)

        self._build_footer(root)

    def _build_header(self, parent) -> None:
        head = tk.Frame(parent, bg=T.BG)
        head.pack(fill="x")

        brand = tk.Canvas(head, width=44, height=44, bg=T.BG, highlightthickness=0)
        brand.pack(side="left")
        round_rect(brand, 2, 2, 42, 42, 12, fill=T.ACCENT_DIM, outline=T.ACCENT)
        brand.create_text(22, 22, text="📧", fill=T.ACCENT_HI, font=ui(17, "bold"))

        titles = tk.Frame(head, bg=T.BG)
        titles.pack(side="left", padx=(14, 0))
        tk.Label(titles, text="Outlook Mail Arama ve Ek İndirme", bg=T.BG, fg=T.TEXT, font=ui(19, "bold")).pack(anchor="w")
        tk.Label(titles, text="Seçilen Outlook klasöründe filtreleme yapıp ekleri (Excel/PDF/Tümü) toplu indirin",
                 bg=T.BG, fg=T.MUTED, font=ui(9)).pack(anchor="w", pady=(3, 0))

        self.status_pill = tk.Canvas(head, width=180, height=32, bg=T.BG, highlightthickness=0)
        self.status_pill.pack(side="right", pady=(6, 0))
        self._draw_pill("Hazır", T.MUTED, T.CARD_HI)

    def _draw_pill(self, text: str, fg: str, bg: str) -> None:
        c = self.status_pill
        c.delete("all")
        round_rect(c, 1, 1, 179, 31, 15, fill=bg, outline=fg)
        c.create_oval(13, 12, 21, 20, fill=fg, outline="")
        c.create_text(90, 16, text=text, fill=fg, font=ui(9, "bold"))

    def _build_filters(self, parent) -> None:
        card = Card(parent, "Arama Filtreleri", "Outlook klasörünü ve kriterleri seçin")
        card.pack(fill="x")

        # --- KLASÖR SEÇİMİ ---
        tk.Label(card.body, text="Outlook Klasörü", bg=T.CARD, fg=T.MUTED, font=ui(9)).pack(anchor="w")
        folder_row = tk.Frame(card.body, bg=T.CARD)
        folder_row.pack(fill="x", pady=(4, 12))
        folder_row.columnconfigure(0, weight=1)

        entry_wrap = tk.Frame(folder_row, bg=T.BG_SOFT, highlightthickness=1, highlightbackground=T.BORDER)
        entry_wrap.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.folder_entry = tk.Entry(entry_wrap, textvariable=self.selected_folder_name,
                                      bg=T.BG_SOFT, fg=T.TEXT, insertbackground=T.TEXT,
                                      relief="flat", font=ui(9), bd=0, state="readonly")
        self.folder_entry.pack(fill="x", padx=10, pady=8)

        RoundedButton(folder_row, "Klasör Seç…", self._pick_folder, kind="secondary", width=140, height=38, font_size=10, bg=T.CARD).grid(row=0, column=1)

        # Konu filtresi
        tk.Label(card.body, text="Konu içeriyor", bg=T.CARD, fg=T.MUTED, font=ui(9)).pack(anchor="w")
        subj_wrap = tk.Frame(card.body, bg=T.BG_SOFT, highlightthickness=1, highlightbackground=T.BORDER)
        subj_wrap.pack(fill="x", pady=(4, 12))
        self.entry_subject = tk.Entry(subj_wrap, textvariable=self.subject_filter, bg=T.BG_SOFT, fg=T.TEXT,
                                       insertbackground=T.TEXT, relief="flat", font=ui(9), bd=0)
        self.entry_subject.pack(fill="x", padx=10, pady=8)

        # Alıcı filtresi
        tk.Label(card.body, text="Alıcı (To) içeriyor", bg=T.CARD, fg=T.MUTED, font=ui(9)).pack(anchor="w")
        to_wrap = tk.Frame(card.body, bg=T.BG_SOFT, highlightthickness=1, highlightbackground=T.BORDER)
        to_wrap.pack(fill="x", pady=(4, 12))
        self.entry_to = tk.Entry(to_wrap, textvariable=self.to_filter, bg=T.BG_SOFT, fg=T.TEXT,
                                  insertbackground=T.TEXT, relief="flat", font=ui(9), bd=0)
        self.entry_to.pack(fill="x", padx=10, pady=8)

        # Gönderen filtresi
        tk.Label(card.body, text="Gönderen adı içeriyor", bg=T.CARD, fg=T.MUTED, font=ui(9)).pack(anchor="w")
        sender_wrap = tk.Frame(card.body, bg=T.BG_SOFT, highlightthickness=1, highlightbackground=T.BORDER)
        sender_wrap.pack(fill="x", pady=(4, 12))
        self.entry_sender = tk.Entry(sender_wrap, textvariable=self.sender_filter, bg=T.BG_SOFT, fg=T.TEXT,
                                      insertbackground=T.TEXT, relief="flat", font=ui(9), bd=0)
        self.entry_sender.pack(fill="x", padx=10, pady=8)

        # Tarih aralığı
        tk.Label(card.body, text="Tarih Aralığı (Gönderim/Alım)", bg=T.CARD, fg=T.MUTED, font=ui(9)).pack(anchor="w", pady=(6, 0))
        date_row = tk.Frame(card.body, bg=T.CARD)
        date_row.pack(fill="x", pady=(4, 12))
        date_row.columnconfigure(0, weight=1)
        date_row.columnconfigure(1, weight=1)

        self.date_from = DateEntry(date_row, "Başlangıç", default_days_ago=30)
        self.date_from.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        self.date_to = DateEntry(date_row, "Bitiş", default_days_ago=0)
        self.date_to.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        # Ek türü seçimi
        tk.Label(card.body, text="İndirilecek Ek Türü", bg=T.CARD, fg=T.MUTED, font=ui(9)).pack(anchor="w", pady=(6, 0))
        att_row = tk.Frame(card.body, bg=T.CARD)
        att_row.pack(fill="x", pady=(6, 12))

        self.att_filter_label = tk.Label(att_row, text=self._filter_type_text(), bg=T.CARD, fg=T.TEXT_SOFT, font=ui(9))
        self.att_filter_label.pack(side="left", fill="x", expand=True)

        RoundedButton(att_row, "Tür Seç…", self._pick_attachment_filter, kind="secondary", width=120, height=38, font_size=10, bg=T.CARD).pack(side="right")

        # İleri ayarlar
        adv = tk.Frame(card.body, bg=T.CARD)
        adv.pack(fill="x", pady=(6, 0))
        adv.columnconfigure(0, weight=1)

        max_row = tk.Frame(adv, bg=T.CARD)
        max_row.grid(row=0, column=0, sticky="ew", pady=3)
        tk.Label(max_row, text="Maksimum mail sayısı", bg=T.CARD, fg=T.MUTED, font=ui(9)).pack(side="left")
        spin = tk.Spinbox(max_row, from_=10, to=2000, textvariable=self.max_mails, width=6,
                          bg=T.BG_SOFT, fg=T.TEXT, buttonbackground=T.CARD_HI, relief="flat",
                          insertbackground=T.TEXT, font=ui(9), justify="center")
        spin.pack(side="left", padx=(8, 4))

        Switch(adv, "Her mail için alt klasör oluştur (Tarih_Konu)", self.create_subfolders.get(),
               lambda v: self.create_subfolders.set(v), bg=T.CARD).grid(row=1, column=0, sticky="w", pady=3)
        Switch(adv, "İndirme sonrası klasörü aç", self.open_folder_after.get(),
               lambda v: self.open_folder_after.set(v), bg=T.CARD).grid(row=2, column=0, sticky="w", pady=3)
        Switch(adv, "Ayrıntılı kayıt (debug)", self.debug_mode.get(),
               lambda v: self.debug_mode.set(v), bg=T.CARD).grid(row=3, column=0, sticky="w", pady=3)

        # İndirme klasörü
        tk.Label(card.body, text="İndirme Klasörü", bg=T.CARD, fg=T.MUTED, font=ui(9)).pack(anchor="w", pady=(12, 0))
        folder_row2 = tk.Frame(card.body, bg=T.CARD)
        folder_row2.pack(fill="x", pady=(6, 0))

        entry_wrap2 = tk.Frame(folder_row2, bg=T.BG_SOFT, highlightthickness=1, highlightbackground=T.BORDER)
        entry_wrap2.pack(side="left", fill="x", expand=True)
        self.entry_folder = tk.Entry(entry_wrap2, textvariable=self.download_folder, bg=T.BG_SOFT, fg=T.TEXT,
                                      insertbackground=T.TEXT, relief="flat", font=ui(9), bd=0)
        self.entry_folder.pack(fill="x", padx=10, pady=8)

        RoundedButton(folder_row2, "Gözat…", self._browse_folder, kind="secondary", width=110, height=38, font_size=10, bg=T.CARD).pack(side="left", padx=(10, 0))

    def _build_progress(self, parent) -> None:
        card = Card(parent, "İlerleme")
        card.pack(fill="both", expand=True, pady=(14, 0))

        top = tk.Frame(card.body, bg=T.CARD)
        top.pack(fill="x")
        self.activity = tk.Label(top, text="Hazır. Klasör ve filtreleri seçip 'Ara ve İndir'e basın.",
                                  bg=T.CARD, fg=T.TEXT_SOFT, font=ui(9), anchor="w")
        self.activity.pack(side="left", fill="x", expand=True)
        self.elapsed_label = tk.Label(top, text="00:00", bg=T.CARD, fg=T.MUTED, font=(T.MONO, 10))
        self.elapsed_label.pack(side="right")

        self.bar = ProgressBar(card.body, height=8, bg=T.CARD)
        self.bar.pack(fill="x", pady=(8, 12))

        self.progress_detail = tk.Label(card.body, text="", bg=T.CARD, fg=T.TEXT_SOFT, font=ui(8), anchor="w")
        self.progress_detail.pack(fill="x", pady=(0, 8))

        actions = tk.Frame(card.body, bg=T.CARD)
        actions.pack(fill="x", pady=(8, 0))
        self.start_btn = RoundedButton(actions, "Ara ve İndir", self._start, kind="primary",
                                       width=220, height=48, radius=12, font_size=11, bg=T.CARD)
        self.start_btn.pack(side="left")
        self.cancel_btn = RoundedButton(actions, "İptal", self._cancel, kind="danger",
                                        width=130, height=48, radius=12, bg=T.CARD)
        self.cancel_btn.pack(side="left", padx=(12, 0))
        self.cancel_btn.set_enabled(False)
        self.open_dir_btn = RoundedButton(actions, "Klasörü Aç", self._open_folder, kind="ghost",
                                          width=140, height=48, radius=12, bg=T.CARD)
        self.open_dir_btn.pack(side="right")
        self.open_dir_btn.set_enabled(False)

    def _build_log(self, parent) -> None:
        card = Card(parent, "Canlı Günlük")
        card.pack(fill="both", expand=True)

        content = tk.Frame(card.body, bg=T.CARD)
        content.pack(fill="both", expand=True)
        content.rowconfigure(1, weight=1)
        content.columnconfigure(0, weight=1)

        head = tk.Frame(content, bg=T.CARD)
        head.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        head.columnconfigure(0, weight=1)
        tk.Label(head, text="Outlook bağlantısı ve indirme işlemleri", bg=T.CARD, fg=T.MUTED, font=ui(8)).grid(row=0, column=0, sticky="w")
        RoundedButton(head, "Temizle", self._clear_log, kind="ghost", width=90, height=30, font_size=9, bg=T.CARD).grid(row=0, column=1, sticky="e")

        wrap = tk.Frame(content, bg=T.LOG_BG, highlightthickness=1, highlightbackground=T.BORDER)
        wrap.grid(row=1, column=0, sticky="nsew")
        self.log = tk.Text(wrap, bg=T.LOG_BG, fg=T.TEXT_SOFT, relief="flat", bd=0,
                           font=(T.MONO, 9), wrap="word", padx=12, pady=10,
                           insertbackground=T.TEXT, state="disabled")
        scroll = tk.Scrollbar(wrap, command=self.log.yview, bg=T.CARD, troughcolor=T.LOG_BG,
                              activebackground=T.BORDER_HI, bd=0, width=10, relief="flat")
        self.log.configure(yscrollcommand=scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.log.tag_configure("time", foreground=T.MUTED)
        self.log.tag_configure("info", foreground="#c2cbd8")
        self.log.tag_configure("ok", foreground=T.OK)
        self.log.tag_configure("warning", foreground=T.WARN)
        self.log.tag_configure("error", foreground=T.ERR)
        self.log.tag_configure("step", foreground=T.ACCENT_HI)
        self.log.tag_configure("debug", foreground=T.MUTED)

    def _build_summary(self, parent) -> None:
        self.summary_card = Card(parent, "Sonuç Özeti")
        self.summary_card.pack(fill="x", pady=(14, 0))

        grid = tk.Frame(self.summary_card.body, bg=T.CARD)
        grid.pack(fill="x")
        for i in range(4):
            grid.columnconfigure(i, weight=1)

        self.summary_values = {}
        for i, (key, label) in enumerate((("mails_found", "BULUNAN MAİL"), ("attachments", "EK SAYISI"),
                                           ("downloaded", "İNDİRİLEN"), ("errors", "HATA"))):
            cell = tk.Frame(grid, bg=T.CARD)
            cell.grid(row=0, column=i, sticky="w", padx=(0, 12))
            tk.Label(cell, text=label, bg=T.CARD, fg=T.MUTED, font=ui(7)).pack(anchor="w")
            value = tk.Label(cell, text="—", bg=T.CARD, fg=T.TEXT, font=(T.FONT_SEMI, 15))
            value.pack(anchor="w", pady=(2, 0))
            self.summary_values[key] = value

        self.summary_note = tk.Label(self.summary_card.body, text="Henüz arama yapılmadı.",
                                      bg=T.CARD, fg=T.MUTED, font=ui(8), anchor="w")
        self.summary_note.pack(fill="x", pady=(10, 0))

    def _build_footer(self, parent) -> None:
        foot = tk.Frame(parent, bg=T.BG)
        foot.pack(fill="x", pady=(10, 0))
        tk.Label(foot, text=f"v{APP_VERSION}  ·  Outlook Mail Ek İndirici",
                 bg=T.BG, fg=T.MUTED, font=ui(8)).pack(side="left")
        tk.Label(foot, text="Outlook açık ve oturum açmış olmalıdır.",
                 bg=T.BG, fg=T.MUTED, font=ui(8)).pack(side="right")

    def _filter_type_text(self) -> str:
        texts = {"all": "📎 Tüm Ekler", "excel": "📊 Excel Dosyaları (.xls, .xlsx, .csv)", "pdf": "📄 PDF Dosyaları"}
        return texts.get(self.attachment_filter.get(), "📎 Tüm Ekler")

    def _browse_folder(self) -> None:
        current = self.download_folder.get().strip()
        path = filedialog.askdirectory(title="İndirme Klasörü Seç", initialdir=current or os.path.expanduser("~"))
        if path:
            self.download_folder.set(path)

    def _pick_folder(self) -> None:
        """Outlook klasör seçim dialogunu aç."""
        if win32com is None:
            messagebox.showerror("Eksik kütüphane", "pywin32 yüklü değil.\n'pip install pywin32' çalıştırın.")
            return

        # Outlook'a bağlanıp namespace al - CoUninitialize YAPMAYIZ, dialog kullanacak
        try:
            pythoncom.CoInitialize()
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
        except Exception as e:
            messagebox.showerror("Bağlantı hatası", f"Outlook'a bağlanılamadı:\n{e}")
            pythoncom.CoUninitialize()
            return

        # Dialog aç - dialog kapanınca CoUninitialize yapacak
        dlg = FolderSelectDialog(self, namespace, self.selected_folder_entry_id.get())
        if dlg.result is not None:
            self.namespace = namespace  # Worker için sakla
            self.selected_folder_entry_id.set(dlg.result_entry_id)
            self.selected_folder_name.set(dlg.result.Name)
            self.folder_entry.configure(state="normal")
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, dlg.result.Name)
            self.folder_entry.configure(state="readonly")
            self._append_log(f"Klasör seçildi: {dlg.result.Name}", "ok")
        else:
            # İptal edildiyse COM'u temizle
            pythoncom.CoUninitialize()

    def _pick_attachment_filter(self) -> None:
        dlg = AttachmentFilterDialog(self, self.attachment_filter.get())
        self.wait_window(dlg)
        if dlg.applied:
            self.attachment_filter.set(dlg.result)
            self.att_filter_label.configure(text=self._filter_type_text())

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _open_folder(self) -> None:
        path = self.download_folder.get().strip()
        if not path or not os.path.isdir(path):
            messagebox.showwarning("Klasör yok", "İndirme klasörü bulunamadı.")
            return
        try:
            os.startfile(path)
        except Exception:
            import subprocess
            subprocess.Popen(["explorer", path])

    def _append_log(self, message: str, level: str = "info", stamp: Optional[str] = None) -> None:
        stamp = stamp or datetime.now().strftime("%H:%M:%S")
        self.log.configure(state="normal")
        self.log.insert("end", f"{stamp}  ", ("time",))
        self.log.insert("end", f"{message}\n", (level,))
        self.log.configure(state="disabled")
        self.log.see("end")

    def _set_running(self, running: bool) -> None:
        self._running = running
        self.start_btn.set_enabled(not running)
        self.cancel_btn.set_enabled(running)
        self.start_btn.set_text("İşlem sürüyor…" if running else "Ara ve İndir")
        self.entry_folder.configure(state="disabled" if running else "normal")
        self.entry_subject.configure(state="disabled" if running else "normal")
        self.entry_to.configure(state="disabled" if running else "normal")
        self.entry_sender.configure(state="disabled" if running else "normal")
        self.folder_entry.configure(state="disabled" if running else "readonly")

    def _reset_ui(self) -> None:
        self.bar.set_mode("idle")
        self.bar.set_value(0)
        self.elapsed_label.configure(text="00:00")
        self.progress_detail.configure(text="")
        self.open_dir_btn.set_enabled(False)
        for key in self.summary_values:
            self.summary_values[key].configure(text="—")
        self.summary_note.configure(text="Arama ve indirme işlemi başlatılıyor…", fg=T.MUTED)

    def _start(self) -> None:
        if self._running:
            return

        folder = self.download_folder.get().strip()
        if not folder:
            messagebox.showwarning("Eksik bilgi", "Lütfen indirme klasörünü seçin.")
            return

        os.makedirs(folder, exist_ok=True)

        date_from = self.date_from.get_datetime()
        date_to = self.date_to.get_datetime()

        if date_from and date_to and date_from > date_to:
            messagebox.showwarning("Tarih hatası", "Başlangıç tarihi bitiş tarihinden sonra olamaz.")
            return

        if win32com is None:
            messagebox.showerror("Eksik kütüphane", "pywin32 yüklü değil.\nKomut satırında 'pip install pywin32' çalıştırın.")
            return

        filters = {
            "subject": self.subject_filter.get().strip() or None,
            "to": self.to_filter.get().strip() or None,
            "sender": self.sender_filter.get().strip() or None,
            "date_from": date_from,
            "date_to": date_to,
            "max_mails": self.max_mails.get(),
        }

        self.cancel_event.clear()
        self._reset_ui()
        self._set_running(True)
        self._draw_pill("Çalışıyor", T.ACCENT, T.ACCENT_DIM)
        self.activity.configure(text="Outlook'a bağlanılıyor…", fg=T.TEXT_SOFT)
        self.bar.set_mode("indeterminate")
        self.started_at = time.time()
        self._append_log("İşlem başlatıldı.", "step")
        self._tick_elapsed()

        # Seçili klasör EntryID'sini worker'a geç
        selected_entry_id = self.selected_folder_entry_id.get().strip() or None
        self.worker = threading.Thread(target=self._worker, args=(filters, selected_entry_id), daemon=True)
        self.worker.start()

    def _worker(self, filters: Dict, selected_folder_entry_id: Optional[str]) -> None:
        reporter = GuiReporter(self.events, self.cancel_event, verbose=self.debug_mode.get())
        processor = OutlookProcessor(reporter)
        self.processor = processor

        try:
            if not processor.connect():
                reporter.error("Outlook'a bağlanılamadı.")
                return

            # Klasörü belirle
            if selected_folder_entry_id:
                target_folder = processor.get_folder_by_entry_id(selected_folder_entry_id)
                if target_folder is None:
                    reporter.error("Seçili klasör bulunamadı. Varsayılan (Gönderilenler) kullanılacak.")
                    target_folder = processor.get_sent_folder()
            else:
                target_folder = processor.get_sent_folder()

            if target_folder is None:
                reporter.error("Hiçbir klasör seçilemedi.")
                return

            reporter.log(f"Seçili klasör: {target_folder.Name}", "info")

            # Arama
            reporter.log("Mailler aranıyor…", "step")
            mails = processor.search_mails(target_folder, filters)

            if not mails:
                reporter.log("Filtreye uygun mail bulunamadı.", "warning")
                reporter.finished({"mails_found": 0, "attachments": 0, "downloaded": 0, "errors": 0})
                return

            reporter.log(f"{len(mails)} mail bulundu. Ekler kontrol ediliyor…", "ok")

            total_atts = 0
            for mail in mails:
                atts = processor.get_attachments_info(mail)
                filtered = processor.filter_attachments(atts, self.attachment_filter.get())
                total_atts += len(filtered)

            if total_atts == 0:
                reporter.log("Filtreye uygun ek bulunamadı.", "warning")
                reporter.finished({"mails_found": len(mails), "attachments": 0, "downloaded": 0, "errors": 0})
                return

            reporter.log(f"{total_atts} ek indirilecek. İndirme başlıyor…", "step")

            def progress_cb(current, total, subject):
                if reporter.cancelled():
                    raise CancelledError()
                reporter.progress(current, total, subject)

            stats = processor.download_attachments(
                mails,
                self.download_folder.get().strip(),
                self.attachment_filter.get(),
                self.create_subfolders.get(),
                progress_cb
            )
            stats["mails_found"] = len(mails)
            reporter.finished(stats)

        except CancelledError:
            reporter.events.put(("cancelled", None))
        except Exception as e:
            reporter.error(f"Beklenmeyen hata: {e}")

    def _cancel(self) -> None:
        if not self._running:
            return
        self.cancel_event.set()
        self.cancel_btn.set_enabled(False)
        self._append_log("İptal isteği gönderildi…", "warning")
        self.activity.configure(text="İptal ediliyor…", fg=T.WARN)

    def _drain_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        self.after(70, self._drain_events)

    def _handle_event(self, event: tuple) -> None:
        kind = event[0]
        if kind == "log":
            _, level, msg = event
            self._append_log(msg, level)
        elif kind == "progress":
            _, current, total, subject = event
            self.bar.set_mode("determinate")
            self.bar.set_value(current / total if total > 0 else 0)
            self.activity.configure(text=f"İşleniyor: {current}/{total}", fg=T.TEXT_SOFT)
            self.progress_detail.configure(text=f"{subject[:60]}" + ("…" if len(subject) > 60 else ""))
        elif kind == "finished":
            self._on_finished(event[1])
        elif kind == "cancelled":
            self._finish("İptal edildi", T.WARN, T.WARN_DIM)
            self._append_log("İşlem kullanıcı tarafından iptal edildi.", "warning")
        elif kind == "error":
            self._finish("Hata", T.ERR, T.ERR_DIM)
            for line in str(event[1]).splitlines():
                self._append_log(line, "error")
            self.summary_note.configure(text="İşlem tamamlanamadı. Ayrıntı için günlüğe bakın.", fg=T.ERR)

    def _on_finished(self, stats: Dict) -> None:
        self._finish("Tamamlandı", T.OK, T.OK_DIM)
        self.bar.set_mode("determinate")
        self.bar.set_value(1.0)

        self.summary_values["mails_found"].configure(text=str(stats.get("mails_found", 0)))
        self.summary_values["attachments"].configure(text=str(stats.get("attachments", 0)))
        self.summary_values["downloaded"].configure(text=str(stats.get("downloaded", 0)))
        self.summary_values["errors"].configure(text=str(stats.get("errors", 0)))

        downloaded = stats.get("downloaded", 0)
        errors = stats.get("errors", 0)
        if errors == 0:
            note = f"{downloaded} dosya başarıyla indirildi."
            color = T.OK
        else:
            note = f"{downloaded} dosya indirildi, {errors} hata oluştu."
            color = T.WARN if downloaded > 0 else T.ERR

        self.summary_note.configure(
            text=f"{note} · {datetime.now():%d.%m.%Y %H:%M}",
            fg=color,
        )
        self._append_log(f"İşlem tamamlandı: {downloaded} indirildi, {errors} hata.", "ok" if errors == 0 else "warning")
        self.open_dir_btn.set_enabled(True)

    def _finish(self, pill: str, color: str, bg: str, duration: Optional[float] = None) -> None:
        if duration is None and self.started_at is not None:
            duration = time.time() - self.started_at
        self._set_running(False)
        self.bar.set_mode("idle")
        self._draw_pill(pill, color, bg)
        self.started_at = None
        if duration is not None:
            self.elapsed_label.configure(text=self._fmt_duration(duration))
        self.activity.configure(
            text={"Tamamlandı": "İşlem başarıyla tamamlandı.",
                  "Hata": "İşlem sırasında hata oluştu.",
                  "İptal edildi": "İşlem iptal edildi."}.get(pill, pill),
            fg=color,
        )
        self._save_settings()

    def _tick_elapsed(self) -> None:
        if self.started_at is None:
            return
        elapsed = time.time() - self.started_at
        self.elapsed_label.configure(text=self._fmt_duration(elapsed))
        self.after(200, self._tick_elapsed)

    @staticmethod
    def _fmt_duration(seconds: float) -> str:
        seconds = int(max(0, seconds))
        if seconds < 60:
            return f"{seconds} sn"
        return f"{seconds // 60:02d}:{seconds % 60:02d}"

    def _load_settings(self) -> None:
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            self.download_folder.set(data.get("download_folder") or self.download_folder.get())
            self.subject_filter.set(data.get("subject_filter", ""))
            self.to_filter.set(data.get("to_filter", ""))
            self.sender_filter.set(data.get("sender_filter", ""))
            self.max_mails.set(int(data.get("max_mails", 200)))
            self.attachment_filter.set(data.get("attachment_filter", "all"))
            self.create_subfolders.set(bool(data.get("create_subfolders", True)))
            self.open_folder_after.set(bool(data.get("open_folder_after", True)))
            self.debug_mode.set(bool(data.get("debug_mode", False)))
            self.selected_folder_entry_id.set(data.get("selected_folder_entry_id", ""))
            self.selected_folder_name.set(data.get("selected_folder_name", "Gönderilenler (Varsayılan)"))
        except Exception:
            pass

    def _save_settings(self) -> None:
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as fh:
                json.dump({
                    "download_folder": self.download_folder.get().strip(),
                    "subject_filter": self.subject_filter.get().strip(),
                    "to_filter": self.to_filter.get().strip(),
                    "sender_filter": self.sender_filter.get().strip(),
                    "max_mails": int(self.max_mails.get()),
                    "attachment_filter": self.attachment_filter.get(),
                    "create_subfolders": self.create_subfolders.get(),
                    "open_folder_after": self.open_folder_after.get(),
                    "debug_mode": self.debug_mode.get(),
                    "selected_folder_entry_id": self.selected_folder_entry_id.get(),
                    "selected_folder_name": self.selected_folder_name.get(),
                }, fh, ensure_ascii=False, indent=2)
        except Exception:
            pass


class CancelledError(Exception):
    """Kullanıcı işlemi iptal etti."""
    pass


def enable_dpi() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def apply_dark_titlebar(window) -> None:
    """Windows başlık çubuğunu koyu temaya çevirir."""
    try:
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())
        enabled = ctypes.c_int(1)
        for attribute in (20, 19):
            ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd, attribute, ctypes.byref(enabled), ctypes.sizeof(enabled)
            )
        color = ctypes.c_int(0x00000000 | (0x13 << 16) | (0x0e << 8) | 0x0b)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd, 34, ctypes.byref(color), ctypes.sizeof(color))
    except Exception:
        pass


def apply_icon(window) -> None:
    """Pencere ikonunu çizer."""
    try:
        from PIL import Image, ImageDraw, ImageTk
    except Exception:
        return
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle((2, 2, size - 2, size - 2), radius=16, fill=(59, 130, 246, 255))
    draw.text((16, 14), "📧", fill=(255, 255, 255, 255), font_size=28)
    photo = ImageTk.PhotoImage(img)
    window.iconphoto(True, photo)
    window._icon_ref = photo


def main() -> None:
    enable_dpi()
    app = App()
    apply_dark_titlebar(app)
    apply_icon(app)
    app.mainloop()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass