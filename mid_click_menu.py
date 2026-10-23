import sys, os, math, subprocess, zipfile, shutil, threading, time, traceback, json
import webbrowser
from pathlib import Path

from PyQt5.QtWidgets import (QApplication, QWidget, QSystemTrayIcon, QMenu as QMenuW,
                             QAction as QActionW, QMessageBox, QDialog, QVBoxLayout,
                             QHBoxLayout, QLabel, QPushButton, QRubberBand, QInputDialog,
                             QFileDialog, QFileIconProvider, QStyle, QLineEdit,
                             QListWidget, QListWidgetItem, QTabWidget, QFormLayout,
                             QColorDialog, QSpinBox, QComboBox)
from PyQt5.QtCore import (Qt, QPoint, QTimer, QRectF, QRect, QSize, QFileInfo,
                          pyqtSignal, QObject, QEvent)
from PyQt5.QtGui import (QPainter, QColor, QPen, QFont, QPainterPath, QPixmap, QIcon,
                         QLinearGradient, QRadialGradient)

from PIL import Image

try:
    import pythoncom
    import win32clipboard, win32con, win32api, win32gui
    import win32com.client
    import win32com.shell.shell as shell
    import win32com.shell.shellcon as shellcon
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

try:
    from pynput import mouse as pynput_mouse
    from pynput import keyboard as pynput_keyboard
    from pynput.keyboard import Controller as PynController, Key as PynKey
    HAS_PYNPUT = True
except ImportError:
    HAS_PYNPUT = False

try:
    import keyboard as kb
    HAS_KB = True
except ImportError:
    HAS_KB = False

IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp', '.tiff'}
CONFIG_FILE = Path.home() / ".radial_apps_v2.json"

def _excepthook(typ, val, tb):
    print("\n=== UNHANDLED EXCEPTION ===")
    traceback.print_exception(typ, val, tb)
    print("===========================\n")

sys.excepthook = _excepthook


# ---------- ФАЙЛОВЫЕ ФУНКЦИИ (возвращены) ----------

def get_clipboard_files():
    files = []
    if HAS_WIN32:
        try:
            win32clipboard.OpenClipboard()
            try:
                if win32clipboard.IsClipboardFormatAvailable(win32con.CF_HDROP):
                    files = list(win32clipboard.GetClipboardData(win32con.CF_HDROP))
            finally:
                win32clipboard.CloseClipboard()
        except Exception:
            pass
    if not files:
        for line in QApplication.clipboard().text().splitlines():
            line = line.strip()
            if line and os.path.exists(line):
                files.append(line)
    return files


def get_selected_files_from_explorer():
    if not HAS_WIN32:
        return []
    try:
        hwnd = win32gui.GetForegroundWindow()
        cls = win32gui.GetClassName(hwnd)
        if cls not in ('CabinetWClass', 'ExploreWClass'):
            return []
        pythoncom.CoInitialize()
        try:
            sh = win32com.client.Dispatch("Shell.Application")
            for window in sh.Windows():
                try:
                    if window.HWND == hwnd:
                        items = window.Document.SelectedItems()
                        return [item.Path for item in items]
                except Exception:
                    continue
        finally:
            pythoncom.CoUninitialize()
    except Exception as e:
        print("[shell]", e)
    return []


_pynput_kb = None
def _kb_ctl():
    global _pynput_kb
    if _pynput_kb is None and HAS_PYNPUT:
        _pynput_kb = PynController()
    return _pynput_kb


def send_copy():
    if sys.platform == 'win32' and HAS_WIN32:
        try:
            win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
            time.sleep(0.02)
            win32api.keybd_event(0x43, 0, 0, 0)
            time.sleep(0.02)
            win32api.keybd_event(0x43, 0, win32con.KEYEVENTF_KEYUP, 0)
            time.sleep(0.02)
            win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
            return
        except Exception:
            pass
    ctl = _kb_ctl()
    if ctl is not None:
        try:
            with ctl.pressed(PynKey.ctrl):
                ctl.press('c'); ctl.release('c')
        except Exception:
            pass


def act_copy_path(files):
    QApplication.clipboard().setText("\n".join(files))

def act_reveal(files):
    if not files: return
    f = os.path.normpath(files[0])
    if sys.platform == 'win32':
        subprocess.Popen(['explorer', '/select,', f])
    elif sys.platform == 'darwin':
        subprocess.Popen(['open', '-R', f])
    else:
        subprocess.Popen(['xdg-open', os.path.dirname(f)])

def act_delete(files):
    if not files: return
    if QMessageBox.question(None, "Удалить",
                            f"Отправить в корзину {len(files)} объект(ов)?") != QMessageBox.Yes:
        return
    for f in files:
        try:
            if sys.platform == 'win32' and HAS_WIN32:
                shell.SHFileOperation((0, shellcon.FO_DELETE, f, None,
                                       shellcon.FOF_ALLOWUNDO | shellcon.FOF_NOCONFIRMATION,
                                       None, None))
            else:
                shutil.rmtree(f) if os.path.isdir(f) else os.remove(f)
        except Exception as e:
            print("[delete]", e)

def act_zip(files):
    if not files: return
    base = files[0].rstrip('/\\') if len(files) == 1 else \
           os.path.join(os.path.dirname(files[0]), "archive")
    out = base + ".zip"
    i = 1
    while os.path.exists(out):
        out = f"{base}_{i}.zip"; i += 1
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            if os.path.isdir(f):
                for root, _, names in os.walk(f):
                    for n in names:
                        full = os.path.join(root, n)
                        z.write(full, os.path.relpath(full, os.path.dirname(f)))
            else:
                z.write(f, os.path.basename(f))
    QMessageBox.information(None, "Готово", out)

def act_resize(files):
    imgs = [f for f in files if Path(f).suffix.lower() in IMAGE_EXTS]
    if not imgs:
        QMessageBox.warning(None, "Ошибка", "Нужно изображение"); return
    size, ok = QInputDialog.getItem(None, "Изменить размер", "Ширина (px):",
                                    ["1920", "1280", "800", "400", "200"], 0, True)
    if not ok: return
    try: new_w = int(size)
    except ValueError: return
    for path in imgs:
        img = Image.open(path)
        if img.width <= new_w: continue
        new_h = int(img.height * new_w / img.width)
        img = img.resize((new_w, new_h), Image.LANCZOS)
        p = Path(path)
        out = p.with_name(f"{p.stem}_{new_w}{p.suffix}")
        i = 1
        while out.exists():
            out = p.with_name(f"{p.stem}_{new_w}_{i}{p.suffix}"); i += 1
        img.save(out)
    QMessageBox.information(None, "Готово", f"Обработано: {len(imgs)}")

def act_duplicate(files):
    for f in files:
        p = Path(f)
        if p.is_dir():
            out = str(p) + "_copy"; i = 1
            while os.path.exists(out):
                out = f"{p}_copy_{i}"; i += 1
            shutil.copytree(f, out)
        else:
            out = p.with_name(p.stem + "_copy" + p.suffix); i = 1
            while out.exists():
                out = p.with_name(f"{p.stem}_copy_{i}{p.suffix}"); i += 1
            shutil.copy2(f, out)


class CropDialog(QDialog):
    def __init__(self, image_path):
        super().__init__()
        self.setWindowTitle(f"Обрезать: {Path(image_path).name}")
        self.image_path = image_path
        self.origin = None
        self.original = QPixmap(image_path)
        if self.original.isNull():
            QMessageBox.warning(self, "Ошибка", "Не удалось загрузить изображение")
            QTimer.singleShot(0, self.reject)
            return

        scr = QApplication.primaryScreen().availableGeometry()
        max_w = int(scr.width() * 0.7)
        max_h = int(scr.height() * 0.6)
        scaled = self.original
        if scaled.width() > max_w or scaled.height() > max_h:
            scaled = scaled.scaled(max_w, max_h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.scale = self.original.width() / scaled.width() if scaled.width() else 1.0

        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(6)

        self.view = QLabel()
        self.view.setPixmap(scaled)
        self.view.setFixedSize(scaled.size())
        self.view.setCursor(Qt.CrossCursor)
        self.view.setStyleSheet("background:#141414;")
        lay.addWidget(self.view, 0, Qt.AlignCenter)

        hint = QLabel(f"Выделите область мышкой · {self.original.width()}×{self.original.height()} px")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("color:#777;font-size:11px;")
        lay.addWidget(hint)

        h = QHBoxLayout()
        ok = QPushButton("Обрезать"); ok.setDefault(True)
        cancel = QPushButton("Отмена")
        h.addStretch(); h.addWidget(ok); h.addWidget(cancel); h.addStretch()
        lay.addLayout(h)
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)

        self.rubber = QRubberBand(QRubberBand.Rectangle, self.view)
        self.rubber.setStyleSheet(
            "QRubberBand{background:rgba(80,140,220,90);border:1px solid #5a9cf0;}")
        self.view.installEventFilter(self)

        self.adjustSize()
        g = self.frameGeometry(); g.moveCenter(scr.center()); self.move(g.topLeft())

    def eventFilter(self, obj, ev):
        if obj is self.view:
            if ev.type() == QEvent.MouseButtonPress and ev.button() == Qt.LeftButton:
                self.origin = ev.pos()
                self.rubber.setGeometry(QRect(self.origin, QSize()))
                self.rubber.show()
                return True
            if ev.type() == QEvent.MouseMove and self.origin is not None:
                self.rubber.setGeometry(QRect(self.origin, ev.pos()).normalized())
                return True
            if ev.type() == QEvent.MouseButtonRelease and ev.button() == Qt.LeftButton:
                self.origin = None
                return True
        return super().eventFilter(obj, ev)

    def crop_box(self):
        if not self.rubber.isVisible():
            return None
        r = self.rubber.geometry()
        if r.width() < 5 or r.height() < 5:
            return None
        s = self.scale
        x1 = max(0, int(r.x() * s)); y1 = max(0, int(r.y() * s))
        x2 = min(self.original.width(), int((r.x() + r.width()) * s))
        y2 = min(self.original.height(), int((r.y() + r.height()) * s))
        if x2 - x1 < 2 or y2 - y1 < 2:
            return None
        return (x1, y1, x2, y2)


def act_crop(files):
    imgs = [f for f in files if Path(f).suffix.lower() in IMAGE_EXTS]
    if not imgs:
        QMessageBox.warning(None, "Ошибка", "Нужно изображение"); return
    path = imgs[0]
    dlg = CropDialog(path)
    if dlg.exec_() != QDialog.Accepted:
        return
    box = dlg.crop_box()
    if not box:
        QMessageBox.information(None, "Отмена", "Область не выделена"); return
    img = Image.open(path)
    p = Path(path)
    out = p.with_name(f"{p.stem}_crop{p.suffix}")
    i = 1
    while out.exists():
        out = p.with_name(f"{p.stem}_crop_{i}{p.suffix}"); i += 1
    img.crop(box).save(out)
    QMessageBox.information(None, "Готово", str(out))


# ---------- Конфигурация ----------

DEFAULT_CONFIG = {
    "apps": [
        {"type": "app", "path": r"C:\Windows\System32\notepad.exe", "name": "Блокнот", "icon": ""},
        {"type": "app", "path": r"C:\Windows\System32\calc.exe", "name": "Калькулятор", "icon": ""},
        {"type": "url", "path": "https://google.com", "name": "Google", "icon": ""}
    ],
    "settings": {
        "radius": 130, "inner": 52, "bg_color": "#26262D",
        "accent_color": "#2D5AB4", "opacity": 240
    }
}

def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text("utf-8"))
        except Exception as e:
            print("[config] read:", e)
    return json.loads(json.dumps(DEFAULT_CONFIG))

def save_config(cfg):
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), "utf-8")
    except Exception as e:
        print("[config] write:", e)


def icon_for_item(item):
    if item.get("icon") and os.path.exists(item["icon"]):
        return QIcon(item["icon"])
    if item.get("type") == "url":
        return QApplication.style().standardIcon(QStyle.SP_DriveNetIcon)
    try:
        return QFileIconProvider().icon(QFileInfo(item["path"]))
    except Exception:
        return QApplication.style().standardIcon(QStyle.SP_FileIcon)

def icon_plus():
    pm = QPixmap(32, 32); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    p.setPen(QPen(QColor(120, 200, 140), 3))
    p.drawLine(16, 6, 16, 26); p.drawLine(6, 16, 26, 16)
    p.end(); return QIcon(pm)

def icon_minus():
    pm = QPixmap(32, 32); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
    p.setPen(QPen(QColor(220, 100, 100), 3))
    p.drawLine(6, 16, 26, 16)
    p.end(); return QIcon(pm)


# ---------- Диалоги настроек ----------

class EditItemDialog(QDialog):
    def __init__(self, item, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Редактировать элемент")
        self.resize(450, 250)
        self.item = item.copy()

        lay = QFormLayout(self)
        self.type_combo = QComboBox()
        self.type_combo.addItems(["Приложение", "Ссылка"])
        self.type_combo.setCurrentIndex(0 if item.get("type") == "app" else 1)
        lay.addRow("Тип:", self.type_combo)

        self.name_edit = QLineEdit(item.get("name", ""))
        lay.addRow("Название:", self.name_edit)

        path_lay = QHBoxLayout()
        self.path_edit = QLineEdit(item.get("path", ""))
        self.path_btn = QPushButton("...")
        self.path_btn.clicked.connect(self._pick_path)
        path_lay.addWidget(self.path_edit)
        path_lay.addWidget(self.path_btn)
        lay.addRow("Путь/URL:", path_lay)

        icon_lay = QHBoxLayout()
        self.icon_edit = QLineEdit(item.get("icon", ""))
        self.icon_btn = QPushButton("...")
        self.icon_btn.clicked.connect(self._pick_icon)
        icon_lay.addWidget(self.icon_edit)
        icon_lay.addWidget(self.icon_btn)
        lay.addRow("Иконка:", icon_lay)

        btns = QHBoxLayout()
        ok = QPushButton("Сохранить"); cancel = QPushButton("Отмена")
        btns.addStretch(); btns.addWidget(ok); btns.addWidget(cancel)
        lay.addRow(btns)
        ok.clicked.connect(self.accept)
        cancel.clicked.connect(self.reject)

    def _pick_path(self):
        if self.type_combo.currentIndex() == 0:
            p, _ = QFileDialog.getOpenFileName(self, "Выбрать файл", "C:\\", "Все файлы (*.*)")
            if p: self.path_edit.setText(p)
        else:
            p, ok = QInputDialog.getText(self, "URL", "Введите ссылку:", text=self.path_edit.text())
            if ok: self.path_edit.setText(p)

    def _pick_icon(self):
        p, _ = QFileDialog.getOpenFileName(self, "Выбрать иконку", "C:\\", "Иконки (*.png *.ico *.jpg *.jpeg)")
        if p: self.icon_edit.setText(p)

    def get_item(self):
        self.item["type"] = "app" if self.type_combo.currentIndex() == 0 else "url"
        self.item["name"] = self.name_edit.text().strip() or "Без имени"
        self.item["path"] = self.path_edit.text().strip()
        self.item["icon"] = self.icon_edit.text().strip()
        return self.item


class SettingsDialog(QDialog):
    def __init__(self, config, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки Radial Menu")
        self.resize(700, 550)
        self.config = json.loads(json.dumps(config))
        self.saved = False

        self.tabs = QTabWidget(self)
        lay = QVBoxLayout(self)
        lay.addWidget(self.tabs)

        # Tab 1: Items
        self.tab_items = QWidget()
        items_lay = QVBoxLayout(self.tab_items)
        self.list = QListWidget()
        self.list.setIconSize(QSize(24, 24))
        items_lay.addWidget(self.list)
        
        item_btns = QHBoxLayout()
        b_add = QPushButton("Добавить"); b_edit = QPushButton("Изменить"); b_del = QPushButton("Удалить")
        item_btns.addWidget(b_add); item_btns.addWidget(b_edit); item_btns.addWidget(b_del); item_btns.addStretch()
        items_lay.addLayout(item_btns)
        b_add.clicked.connect(self._add_item)
        b_edit.clicked.connect(self._edit_item)
        b_del.clicked.connect(self._del_item)
        self.list.itemDoubleClicked.connect(lambda _: self._edit_item())
        self.tabs.addTab(self.tab_items, "Приложения и ссылки")

        # Tab 2: Global Settings
        self.tab_style = QWidget()
        style_lay = QFormLayout(self.tab_style)
        s = self.config.get("settings", {})
        self.spin_radius = QSpinBox(); self.spin_radius.setRange(80, 300); self.spin_radius.setValue(s.get("radius", 130))
        self.spin_inner = QSpinBox(); self.spin_inner.setRange(20, 150); self.spin_inner.setValue(s.get("inner", 52))
        self.spin_opacity = QSpinBox(); self.spin_opacity.setRange(100, 255); self.spin_opacity.setValue(s.get("opacity", 240))
        self.btn_bg = QPushButton(s.get("bg_color", "#26262D"))
        self.btn_accent = QPushButton(s.get("accent_color", "#2D5AB4"))
        self.btn_bg.clicked.connect(lambda: self._pick_color(self.btn_bg))
        self.btn_accent.clicked.connect(lambda: self._pick_color(self.btn_accent))
        
        style_lay.addRow("Радиус меню:", self.spin_radius)
        style_lay.addRow("Радиус центра:", self.spin_inner)
        style_lay.addRow("Прозрачность (100-255):", self.spin_opacity)
        style_lay.addRow("Цвет фона:", self.btn_bg)
        style_lay.addRow("Цвет акцента:", self.btn_accent)
        self.tabs.addTab(self.tab_style, "Внешний вид")

        btm = QHBoxLayout()
        b_save = QPushButton("Сохранить"); b_cancel = QPushButton("Отмена")
        btm.addStretch(); btm.addWidget(b_save); btm.addWidget(b_cancel)
        lay.addLayout(btm)
        b_save.clicked.connect(self._save)
        b_cancel.clicked.connect(self.reject)

        self._populate_list()

    def _populate_list(self):
        self.list.clear()
        for i, item in enumerate(self.config["apps"]):
            icon = icon_for_item(item)
            lw = QListWidgetItem(icon, f"[{item['type']}] {item['name']}")
            lw.setData(Qt.UserRole, i)
            self.list.addItem(lw)

    def _add_item(self):
        dlg = EditItemDialog({"type": "app", "path": "", "name": "Новый", "icon": ""})
        if dlg.exec_() == QDialog.Accepted:
            self.config["apps"].append(dlg.get_item())
            self._populate_list()

    def _edit_item(self):
        item = self.list.currentItem()
        if not item: return
        idx = item.data(Qt.UserRole)
        dlg = EditItemDialog(self.config["apps"][idx])
        if dlg.exec_() == QDialog.Accepted:
            self.config["apps"][idx] = dlg.get_item()
            self._populate_list()

    def _del_item(self):
        item = self.list.currentItem()
        if not item: return
        idx = item.data(Qt.UserRole)
        del self.config["apps"][idx]
        self._populate_list()

    def _pick_color(self, btn):
        c = QColorDialog.getColor(QColor(btn.text()), self)
        if c.isValid():
            btn.setText(c.name())

    def _save(self):
        self.config["settings"] = {
            "radius": self.spin_radius.value(),
            "inner": self.spin_inner.value(),
            "bg_color": self.btn_bg.text(),
            "accent_color": self.btn_accent.text(),
            "opacity": self.spin_opacity.value()
        }
        self.saved = True
        self.accept()


# ---------- Радиальное меню ----------

class RadialMenu(QWidget):
    chosen = pyqtSignal(str, object, object)
    submenu_requested = pyqtSignal(int, QPoint)

    def __init__(self, items, center, mode, config, files=None, title="МЕНЮ"):
        super().__init__()
        self.setWindowFlags(Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.NoDropShadowWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground, True)
        self.setAttribute(Qt.WA_NoSystemBackground, True)
        self.setAttribute(Qt.WA_DeleteOnClose)
        self.setMouseTracking(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet("background: transparent;")

        self.mode = mode
        self.files = files or []
        self.items = items
        self.title = title
        self.hovered = -1
        
        s = config.get("settings", {})
        self.RADIUS = s.get("radius", 130)
        self.INNER = s.get("inner", 52)
        self.OUTER_PAD = 10
        self.bg_color = QColor(s.get("bg_color", "#26262D"))
        self.accent_color = QColor(s.get("accent_color", "#2D5AB4"))
        self.opacity = s.get("opacity", 240)

        size = (self.RADIUS + self.OUTER_PAD) * 2
        self.setFixedSize(size, size)

        scr = QApplication.screenAt(center) or QApplication.primaryScreen()
        sg = scr.availableGeometry()
        x = max(sg.x(), min(center.x() - size // 2, sg.x() + sg.width() - size))
        y = max(sg.y(), min(center.y() - size // 2, sg.y() + sg.height() - size))
        self.move(x, y)
        self.center_pt = QPoint(size // 2, size // 2)

        self.submenu_timer = QTimer(self)
        self.submenu_timer.setSingleShot(True)
        self.submenu_timer.setInterval(500)
        self.submenu_timer.timeout.connect(self._trigger_submenu)

        self._build_cache()

    def _trigger_submenu(self):
        if self.hovered >= 0:
            _, _, kind, _ = self.items[self.hovered]
            if kind == "submenu":
                step = 2 * math.pi / len(self.items)
                ang = -math.pi / 2 + self.hovered * step
                tr = (self.RADIUS + self.INNER) / 2
                gx = self.x() + self.center_pt.x() + tr * math.cos(ang)
                gy = self.y() + self.center_pt.y() + tr * math.sin(ang)
                self.submenu_requested.emit(self.hovered, QPoint(int(gx), int(gy)))
                self.hovered = -2

    def _sector_path(self, s, e):
        c = self.center_pt
        gap = 0.03
        s += gap; e -= gap
        path = QPainterPath()
        path.moveTo(c.x() + self.INNER * math.cos(s), c.y() + self.INNER * math.sin(s))
        outer = QRectF(c.x() - self.RADIUS, c.y() - self.RADIUS, self.RADIUS * 2, self.RADIUS * 2)
        path.arcTo(outer, -math.degrees(s), -math.degrees(e - s))
        inner = QRectF(c.x() - self.INNER, c.y() - self.INNER, self.INNER * 2, self.INNER * 2)
        path.arcTo(inner, -math.degrees(e), math.degrees(e - s))
        path.closeSubpath()
        return path

    def _hit(self, pos):
        dx = pos.x() - self.center_pt.x()
        dy = pos.y() - self.center_pt.y()
        r = math.hypot(dx, dy)
        if r < self.INNER or r > self.RADIUS:
            return -1
        ang = (math.atan2(dy, dx) + math.pi / 2) % (2 * math.pi)
        n = len(self.items)
        step = 2 * math.pi / n
        return int((ang + step / 2) / step) % n

    def _build_cache(self):
        self._cache = QPixmap(self.size())
        self._cache.fill(Qt.transparent)
        p = QPainter(self._cache)
        p.setRenderHint(QPainter.Antialiasing)
        c = self.center_pt
        n = len(self.items)
        if n == 0: p.end(); return

        grad = QRadialGradient(c, self.RADIUS + self.OUTER_PAD)
        bg = QColor(self.bg_color); bg.setAlpha(self.opacity)
        bg2 = QColor(self.bg_color).darker(150); bg2.setAlpha(self.opacity)
        grad.setColorAt(0.0, bg)
        grad.setColorAt(1.0, bg2)
        
        p.setPen(QPen(QColor(60, 60, 70, 200), 1.5))
        p.setBrush(grad)
        p.drawEllipse(c, self.RADIUS + self.OUTER_PAD - 1, self.RADIUS + self.OUTER_PAD - 1)

        step = 2 * math.pi / n
        for i in range(n):
            ang = -math.pi / 2 + i * step
            path = self._sector_path(ang - step / 2, ang + step / 2)
            sec_color = QColor(self.bg_color).lighter(110)
            sec_color.setAlpha(self.opacity)
            p.setBrush(sec_color)
            p.setPen(QPen(QColor(70, 70, 85, 150), 1))
            p.drawPath(path)

        p.setBrush(QColor(20, 20, 24, 255))
        p.setPen(QPen(self.accent_color, 1.5))
        p.drawEllipse(c, self.INNER, self.INNER)
        p.end()

    def paintEvent(self, ev):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.drawPixmap(0, 0, self._cache)

        c = self.center_pt
        n = len(self.items)
        if n == 0: return
        step = 2 * math.pi / n

        if 0 <= self.hovered < n:
            ang = -math.pi / 2 + self.hovered * step
            path = self._sector_path(ang - step / 2, ang + step / 2)
            grad = QLinearGradient(c.x(), c.y(), c.x() + self.RADIUS * math.cos(ang), c.y() + self.RADIUS * math.sin(ang))
            grad.setColorAt(0, self.accent_color.darker(130))
            grad.setColorAt(1, self.accent_color.lighter(130))
            p.setBrush(grad)
            p.setPen(QPen(self.accent_color.lighter(160), 1.5))
            p.drawPath(path)

        tr = (self.RADIUS + self.INNER) / 2
        for i, (icon, label, _, _) in enumerate(self.items):
            ang = -math.pi / 2 + i * step
            tx = c.x() + tr * math.cos(ang)
            ty = c.y() + tr * math.sin(ang)

            icon_size = 26
            ir = QRectF(tx - icon_size/2, ty - icon_size/2 - 8, icon_size, icon_size)
            if isinstance(icon, QIcon):
                icon.paint(p, ir.toRect(), Qt.AlignCenter, QIcon.Normal, QIcon.Off)

            p.setFont(QFont("Segoe UI", 8, QFont.Medium))
            p.setPen(QColor(230, 235, 245))
            p.drawText(QRectF(tx - 60, ty + 9, 120, 20), Qt.AlignCenter, label)

        p.setFont(QFont("Segoe UI", 8, QFont.Bold))
        p.setPen(QColor(180, 195, 220))
        p.drawText(QRectF(c.x() - self.INNER + 4, c.y() - 10, self.INNER * 2 - 8, 20), Qt.AlignCenter, self.title)

    def mouseMoveEvent(self, ev):
        h = self._hit(ev.pos())
        if h != self.hovered:
            self.hovered = h
            self.submenu_timer.stop()
            if h >= 0:
                _, _, kind, _ = self.items[h]
                if kind == "submenu":
                    self.submenu_timer.start()
            self.update()

    def leaveEvent(self, ev):
        if self.hovered != -1:
            self.hovered = -1
            self.submenu_timer.stop()
            self.update()

    def mousePressEvent(self, ev):
        if ev.button() == Qt.LeftButton:
            h = self._hit(ev.pos())
            if h >= 0:
                _, _, kind, payload = self.items[h]
                if kind != "submenu":
                    self.chosen.emit(kind, payload, self.files)
            self.close()
        elif ev.button() == Qt.RightButton:
            self.close()

    def keyPressEvent(self, ev):
        if ev.key() == Qt.Key_Escape:
            self.close()


# ---------- Хоткей ----------

class HotkeyManager(QObject):
    triggered = pyqtSignal(QPoint)
    any_click = pyqtSignal(QPoint, int, bool)

    def __init__(self):
        super().__init__(); self.listener = None

    def start(self):
        if not HAS_PYNPUT: return
        def on_click(x, y, button, pressed):
            try:
                if button == pynput_mouse.Button.middle and not pressed:
                    self.triggered.emit(QPoint(int(x), int(y)))
                btn = 0
                if button == pynput_mouse.Button.left: btn = 1
                elif button == pynput_mouse.Button.middle: btn = 2
                elif button == pynput_mouse.Button.right: btn = 3
                self.any_click.emit(QPoint(int(x), int(y)), btn, pressed)
            except Exception: pass
        try:
            self.listener = pynput_mouse.Listener(on_click=on_click)
            self.listener.daemon = True; self.listener.start()
        except Exception: pass

    def stop(self):
        if self.listener: self.listener.stop()


# ---------- Приложение ----------

class App(QObject):
    def __init__(self):
        super().__init__()
        self.config = load_config()
        self.tray = QSystemTrayIcon(self._make_icon())
        self.tray.setToolTip("Radial Menu · СКМ")
        self.tray.show()

        self._tray_menu = QMenuW()
        a_settings = QActionW("Настройки", self)
        a_settings.triggered.connect(lambda: self._open_settings())
        self._tray_menu.addAction(a_settings)
        self._tray_menu.addSeparator()
        a_quit = QActionW("Выход", self)
        a_quit.triggered.connect(self.quit)
        self._tray_menu.addAction(a_quit)
        self.tray.setContextMenu(self._tray_menu)

        self.hotkey = HotkeyManager()
        self.hotkey.triggered.connect(self._on_hotkey)
        self.hotkey.any_click.connect(self._on_any_click)
        self.hotkey.start()

        self.menu = None
        self.submenu = None
        self.last_pos = None
        self.esc_listener = None

    def _make_icon(self):
        pm = QPixmap(64, 64); pm.fill(Qt.transparent)
        p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing)
        grad = QRadialGradient(32, 32, 30)
        grad.setColorAt(0, QColor(90, 150, 230)); grad.setColorAt(1, QColor(35, 55, 90))
        p.setPen(QPen(QColor(180, 210, 250), 2)); p.setBrush(grad)
        p.drawEllipse(6, 6, 52, 52)
        p.setBrush(QColor(255, 255, 255))
        for i in range(6):
            ang = i * math.pi / 3
            x = 32 + 20 * math.cos(ang); y = 32 + 20 * math.sin(ang)
            p.drawEllipse(QPoint(int(x), int(y)), 3, 3)
        p.end(); return QIcon(pm)

    def _on_hotkey(self, phys_pos):
        scr = QApplication.primaryScreen()
        dpr = scr.devicePixelRatio() or 1.0
        pos = QPoint(int(phys_pos.x() / dpr), int(phys_pos.y() / dpr))
        self.last_pos = pos

        files = get_selected_files_from_explorer()
        if not files:
            files = [f for f in get_clipboard_files() if os.path.exists(f)]
        
        if files:
            self._show_file_menu(pos, files)
        else:
            self._show_launcher(pos)

    def _on_any_click(self, pos, btn_num, pressed):
        if not pressed: return
        scr = QApplication.primaryScreen()
        dpr = scr.devicePixelRatio() or 1.0
        lpos = QPoint(int(pos.x() / dpr), int(pos.y() / dpr))
        
        in_menu = self.menu and self.menu.geometry().contains(lpos)
        in_sub = self.submenu and self.submenu.geometry().contains(lpos)
        
        if not in_menu and not in_sub:
            self._close_all()

    def _close_all(self):
        if self.submenu:
            try: self.submenu.close()
            except: pass
            self.submenu = None
        if self.menu:
            try: self.menu.close()
            except: pass
            self.menu = None
        self._stop_esc_listener()

    def _start_esc_listener(self):
        self._stop_esc_listener()
        def on_press(key):
            if key == pynput_keyboard.Key.esc:
                QTimer.singleShot(0, self._close_all)
                return False
        try:
            self.esc_listener = pynput_keyboard.Listener(on_press=on_press)
            self.esc_listener.daemon = True; self.esc_listener.start()
        except: pass

    def _stop_esc_listener(self):
        if self.esc_listener:
            try: self.esc_listener.stop()
            except: pass
            self.esc_listener = None

    def _show_file_menu(self, pos, files):
        style = QApplication.style()
        exts = {Path(f).suffix.lower() for f in files if os.path.isfile(f)}
        is_image = bool(exts) and exts.issubset(IMAGE_EXTS)

        items = [(style.standardIcon(QStyle.SP_DirOpenIcon), "Открыть", "file", act_reveal)]
        if is_image:
            items.append((style.standardIcon(QStyle.SP_FileDialogContentsView), "Обрезать", "file", act_crop))
            items.append((style.standardIcon(QStyle.SP_FileDialogDetailedView), "Размер", "file", act_resize))
        items.append((style.standardIcon(QStyle.SP_FileIcon), "Путь", "file", act_copy_path))
        items.append((style.standardIcon(QStyle.SP_DriveHDIcon), "ZIP", "file", act_zip))
        items.append((style.standardIcon(QStyle.SP_FileDialogNewFolder), "Копия", "file", act_duplicate))
        items.append((style.standardIcon(QStyle.SP_TrashIcon), "Удалить", "file", act_delete))
        
        title = Path(files[0]).name if len(files) == 1 else f"×{len(files)}"
        if len(title) > 12: title = title[:11] + "…"
        self._spawn_menu(items, pos, "file", title, files)

    def _show_launcher(self, pos):
        apps = self.config.get("apps", [])
        items = []
        for app in apps:
            items.append((icon_for_item(app), app["name"], "launch", app))
        
        items.append((icon_plus(), "Управление", "submenu", None))
        self._spawn_menu(items, pos, "app", "ПРИЛОЖЕНИЯ", None)

    def _spawn_submenu(self, parent_idx, global_pos):
        items = [
            (icon_plus(), "Добавить", "add_dialog", None),
            (QApplication.style().standardIcon(QStyle.SP_TrashIcon), "Удалить", "remove_dialog", None),
            (QApplication.style().standardIcon(QStyle.SP_FileDialogInfoView), "Настройки", "settings", None)
        ]
        
        if self.submenu:
            try: self.submenu.close()
            except: pass
            self.submenu = None

        self.submenu = RadialMenu(items, global_pos, "submenu", self.config, title="УПРАВЛЕНИЕ")
        self.submenu.chosen.connect(self._on_submenu_chosen)
        self.submenu.destroyed.connect(lambda *_: setattr(self, "submenu", None))
        
        self.submenu.show()
        self.submenu.raise_()
        self.submenu.activateWindow()

    def _spawn_menu(self, items, pos, mode, title, files):
        self._close_all()

        self.menu = RadialMenu(items, pos, mode, self.config, files=files, title=title)
        self.menu.chosen.connect(self._on_chosen)
        self.menu.submenu_requested.connect(self._spawn_submenu)
        self.menu.destroyed.connect(lambda *_: setattr(self, "menu", None))
        
        self.menu.show()
        self.menu.raise_()
        self.menu.activateWindow()
        self.menu.setFocus()
        self._start_esc_listener()

    def _on_chosen(self, kind, payload, files):
        self._close_all()
        QTimer.singleShot(30, lambda: self._dispatch(kind, payload, files))

    def _on_submenu_chosen(self, kind, payload, files):
        self._close_all()
        QTimer.singleShot(30, lambda: self._dispatch(kind, payload, files))

    def _dispatch(self, kind, payload, files):
        try:
            if kind == "file":
                payload(files)
            elif kind == "launch":
                if payload["type"] == "url":
                    webbrowser.open(payload["path"])
                else:
                    os.startfile(payload["path"])
            elif kind == "add_dialog":
                self._open_settings(tab_index=0)
            elif kind == "remove_dialog":
                self._open_settings(tab_index=0)
            elif kind == "settings":
                self._open_settings()
        except Exception:
            traceback.print_exc()
            QMessageBox.critical(None, "Ошибка", traceback.format_exc())

    def _open_settings(self, tab_index=0):
        dlg = SettingsDialog(self.config)
        dlg.tabs.setCurrentIndex(tab_index)
        if dlg.exec_() == QDialog.Accepted and dlg.saved:
            self.config = dlg.config
            save_config(self.config)
            if self.last_pos:
                self._show_launcher(self.last_pos)

    def quit(self):
        self._close_all()
        self.hotkey.stop()
        QApplication.quit()


def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    app.instance = App()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()