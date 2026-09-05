"""
AnotAI - Desktop Application
============================
Drop in a PDF or Word file, get an Excel report of every annotation -
reference, author, date, type, text, a high-resolution detail snapshot
and a full-page overview showing exactly where it sits.

Fully offline. No Azure, no SharePoint, no internet required.

Powered by Maehan Solutions.
"""

import os
import sys
import logging
import subprocess
from pathlib import Path
from typing import Optional

from PyQt5.QtCore import Qt, QObject, QThread, pyqtSignal, QTimer, QUrl, QRectF
from PyQt5.QtGui import (QFont, QDesktopServices, QPainter, QColor, QPixmap,
                         QLinearGradient, QBrush, QPen, QFontMetrics)
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QProgressBar, QTextEdit,
    QMessageBox, QCheckBox, QFrame, QSizePolicy, QSplashScreen,
    QGraphicsDropShadowEffect, QDialog, QDialogButtonBox
)

import brand
from annotation_processor import AnnotationExtractor, Cancelled

SUPPORTED = (".pdf", ".docx")
MAX_LOG_BLOCKS = 400

logger = logging.getLogger(brand.APP_NAME)
logger.setLevel(logging.INFO)


# ======================================================================
# THREAD-SAFE LOG BRIDGE
# ======================================================================
class LogBridge(QObject):
    """
    Carries log records from any thread onto the GUI thread.

    A logging.Handler can fire on the worker thread, and touching a
    QWidget from a non-GUI thread freezes Qt. Records are emitted as a
    signal so delivery always happens on the GUI thread.
    """
    message = pyqtSignal(str)


class BridgeHandler(logging.Handler):
    def __init__(self, bridge: LogBridge):
        super().__init__()
        self.bridge = bridge

    def emit(self, record):
        try:
            self.bridge.message.emit(self.format(record))
        except Exception:
            pass


# ======================================================================
# SPLASH
# ======================================================================
def build_splash_pixmap(w: int = 620, h: int = 360) -> QPixmap:
    app = QApplication.instance()
    ratio = app.primaryScreen().devicePixelRatio() if app else 1.0

    pm = QPixmap(int(w * ratio), int(h * ratio))
    pm.setDevicePixelRatio(ratio)
    pm.fill(Qt.transparent)

    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    p.setRenderHint(QPainter.SmoothPixmapTransform, True)

    grad = QLinearGradient(0, 0, w, h)
    grad.setColorAt(0.0, QColor(brand.DARK_BLUE))
    grad.setColorAt(0.55, QColor(brand.DARK_BLUE_2))
    grad.setColorAt(1.0, QColor(brand.DARK_BLUE))
    p.setBrush(QBrush(grad))
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(QRectF(0, 0, w, h), 16, 16)

    p.setBrush(QColor(79, 163, 217, 34))
    p.drawEllipse(QRectF(w - 190, -110, 300, 300))
    p.setBrush(QColor(127, 196, 236, 22))
    p.drawEllipse(QRectF(-90, h - 150, 240, 240))

    p.setPen(QPen(QColor(brand.SKY), 3))
    p.drawLine(int(w / 2 - 46), 236, int(w / 2 + 46), 236)

    mark = brand.logo_pixmap(104)
    p.drawPixmap(int((w - mark.width() / mark.devicePixelRatio()) / 2), 40, mark)

    f_name = QFont("Segoe UI", 34, QFont.Bold)
    f_name.setLetterSpacing(QFont.AbsoluteSpacing, 1.5)
    p.setFont(f_name)
    p.setPen(QColor(brand.WHITE))
    fm = QFontMetrics(f_name)
    p.drawText(int((w - fm.horizontalAdvance(brand.APP_NAME)) / 2), 218, brand.APP_NAME)

    f_strap = QFont("Segoe UI", 11)
    f_strap.setLetterSpacing(QFont.AbsoluteSpacing, 2.6)
    p.setFont(f_strap)
    p.setPen(QColor(brand.SKY_LIGHT))
    strap = brand.APP_STRAP.upper()
    fm = QFontMetrics(f_strap)
    p.drawText(int((w - fm.horizontalAdvance(strap)) / 2), 266, strap)

    f_by = QFont("Segoe UI", 11, QFont.Bold)
    p.setFont(f_by)
    p.setPen(QColor(brand.CREAM))
    fm = QFontMetrics(f_by)
    p.drawText(int((w - fm.horizontalAdvance(brand.APP_TAGLINE)) / 2), 306,
               brand.APP_TAGLINE)

    f_v = QFont("Segoe UI", 8)
    p.setFont(f_v)
    p.setPen(QColor(127, 196, 236, 165))
    ver = f"Version {brand.APP_VERSION}"
    fm = QFontMetrics(f_v)
    p.drawText(int((w - fm.horizontalAdvance(ver)) / 2), 332, ver)

    p.end()
    return pm


class BrandSplash(QSplashScreen):
    def __init__(self):
        super().__init__(build_splash_pixmap())
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowFlags(
            Qt.SplashScreen | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)

    def note(self, text: str):
        self.showMessage(text, Qt.AlignBottom | Qt.AlignHCenter,
                         QColor(brand.SKY_LIGHT))
        QApplication.processEvents()


# ======================================================================
# HEADER BANNER
# ======================================================================
class BrandHeader(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(84)
        self.setStyleSheet("QFrame{border:none;}")
        self._mark = None

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing, True)
        p.setRenderHint(QPainter.SmoothPixmapTransform, True)
        w, h = self.width(), self.height()

        grad = QLinearGradient(0, 0, w, 0)
        grad.setColorAt(0.0, QColor(brand.DARK_BLUE))
        grad.setColorAt(1.0, QColor(brand.DARK_BLUE_2))
        p.setBrush(QBrush(grad))
        p.setPen(Qt.NoPen)
        p.drawRoundedRect(QRectF(0, 0, w, h), 12, 12)
        p.drawRect(QRectF(0, h - 12, w, 12))

        p.setBrush(QColor(79, 163, 217, 30))
        p.drawEllipse(QRectF(w - 150, -80, 220, 220))

        if self._mark is None:
            self._mark = brand.logo_pixmap(50)
        p.drawPixmap(18, 17, self._mark)

        f_name = QFont("Segoe UI", 19, QFont.Bold)
        f_name.setLetterSpacing(QFont.AbsoluteSpacing, 1.0)
        p.setFont(f_name)
        p.setPen(QColor(brand.WHITE))
        p.drawText(82, 42, brand.APP_NAME)

        f_sub = QFont("Segoe UI", 8)
        f_sub.setLetterSpacing(QFont.AbsoluteSpacing, 1.3)
        p.setFont(f_sub)
        p.setPen(QColor(brand.SKY_LIGHT))
        p.drawText(84, 61, brand.APP_STRAP.upper())

        f_by = QFont("Segoe UI", 9, QFont.Bold)
        p.setFont(f_by)
        p.setPen(QColor(brand.CREAM))
        fm = QFontMetrics(f_by)
        p.drawText(w - fm.horizontalAdvance(brand.APP_TAGLINE) - 20, 46,
                   brand.APP_TAGLINE)

        f_v = QFont("Segoe UI", 8)
        p.setFont(f_v)
        p.setPen(QColor(127, 196, 236, 170))
        ver = f"v{brand.APP_VERSION}"
        fm = QFontMetrics(f_v)
        p.drawText(w - fm.horizontalAdvance(ver) - 20, 63, ver)
        p.end()


# ======================================================================
# ADVANCED SETTINGS DIALOG
# ======================================================================
class SettingsDialog(QDialog):
    """Rarely-needed switches, kept out of the main window."""

    def __init__(self, parent, template_path, include_noise, keep_markup):
        super().__init__(parent)
        self.setWindowTitle(f"{brand.APP_NAME} - Advanced settings")
        self.setWindowIcon(brand.app_icon())
        self.setMinimumWidth(560)

        self.template_path = template_path

        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 16)
        lay.setSpacing(14)

        head = QLabel("Advanced settings")
        head.setFont(QFont("Segoe UI", 13, QFont.Bold))
        head.setStyleSheet("color:%s;" % brand.DARK_BLUE)
        lay.addWidget(head)

        # --- template -------------------------------------------------
        t_lbl = QLabel("Excel template (optional)")
        t_lbl.setStyleSheet("font-weight:700;color:%s;" % brand.DARK_BLUE)
        lay.addWidget(t_lbl)

        t_help = QLabel("Append rows into an existing workbook instead of "
                        "creating a new branded report.")
        t_help.setWordWrap(True)
        t_help.setStyleSheet("color:%s;font-size:9pt;" % brand.MUTED)
        lay.addWidget(t_help)

        t_row = QHBoxLayout()
        self.tpl_label = QLabel(Path(template_path).name if template_path
                                else "None - a fresh branded report will be created")
        self.tpl_label.setStyleSheet("color:%s;" % brand.INK)
        self.tpl_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        t_row.addWidget(self.tpl_label, 1)
        b_sel = QPushButton("Select")
        b_sel.setStyleSheet(brand.SUBTLE_BTN)
        b_sel.setCursor(Qt.PointingHandCursor)
        b_sel.clicked.connect(self._pick)
        t_row.addWidget(b_sel)
        b_rm = QPushButton("Remove")
        b_rm.setStyleSheet(brand.SUBTLE_BTN)
        b_rm.setCursor(Qt.PointingHandCursor)
        b_rm.clicked.connect(self._clear)
        t_row.addWidget(b_rm)
        lay.addLayout(t_row)

        # --- extraction switches --------------------------------------
        e_lbl = QLabel("Extraction")
        e_lbl.setStyleSheet("font-weight:700;color:%s;" % brand.DARK_BLUE)
        lay.addWidget(e_lbl)

        self.chk_markup = QCheckBox(
            "Include shapes and lines that contain no reviewer text")
        self.chk_markup.setChecked(keep_markup)
        lay.addWidget(self.chk_markup)

        m_help = QLabel("Off by default. Rectangles, arrows, clouds and leader "
                        "lines are usually pointers to a nearby comment rather "
                        "than findings in their own right.")
        m_help.setWordWrap(True)
        m_help.setStyleSheet("color:%s;font-size:9pt;margin-left:25px;" % brand.MUTED)
        lay.addWidget(m_help)

        self.chk_noise = QCheckBox("Include Link and Popup objects")
        self.chk_noise.setChecked(include_noise)
        lay.addWidget(self.chk_noise)

        n_help = QLabel("Off by default. These are PDF viewer scaffolding, "
                        "not review comments.")
        n_help.setWordWrap(True)
        n_help.setStyleSheet("color:%s;font-size:9pt;margin-left:25px;" % brand.MUTED)
        lay.addWidget(n_help)

        lay.addStretch(1)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setStyleSheet(brand.PRIMARY_BTN)
        btns.button(QDialogButtonBox.Cancel).setStyleSheet(brand.SUBTLE_BTN)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _pick(self):
        p, _ = QFileDialog.getOpenFileName(
            self, "Select an Excel template", str(Path.home()), "Excel files (*.xlsx)")
        if p:
            self.template_path = p
            self.tpl_label.setText(Path(p).name)

    def _clear(self):
        self.template_path = None
        self.tpl_label.setText("None - a fresh branded report will be created")

    def values(self):
        return (self.template_path,
                self.chk_noise.isChecked(),
                self.chk_markup.isChecked())


# ======================================================================
# BACKGROUND WORKER
# ======================================================================
class ProcessorWorker(QObject):
    progress = pyqtSignal(str, int)
    finished = pyqtSignal(str, int, int)     # path, count, skipped
    failed = pyqtSignal(str)
    cancelled = pyqtSignal()

    def __init__(self, file_path, output_dir, template_path,
                 include_noise, keep_markup):
        super().__init__()
        self.file_path = file_path
        self.output_dir = output_dir
        self.template_path = template_path
        self.include_noise = include_noise
        self.keep_markup = keep_markup
        self._cancel = False

    def request_cancel(self):
        """Called from the GUI thread; only ever sets a flag."""
        self._cancel = True

    def _is_cancelled(self) -> bool:
        return self._cancel

    def _report(self, payload: str):
        if "||" in payload:
            msg, _, pct = payload.rpartition("||")
            try:
                self.progress.emit(msg, int(pct))
                return
            except ValueError:
                pass
        self.progress.emit(payload, -1)

    def run(self):
        extractor = None
        try:
            self.progress.emit("Starting...", 1)
            extractor = AnnotationExtractor(
                logger=logger,
                include_noise=self.include_noise,
                keep_markup_only=self.keep_markup,
                progress_cb=self._report,
                cancel_cb=self._is_cancelled,
            )
            out_path, count, skipped = extractor.process(
                self.file_path, output_dir=self.output_dir,
                template_path=self.template_path)
            self.finished.emit(out_path, count, skipped)
        except Cancelled:
            self.cancelled.emit()
        except Exception as exc:
            logger.error(f"Processing failed: {exc}", exc_info=True)
            self.failed.emit(str(exc))
        finally:
            if extractor is not None:
                try:
                    extractor.cleanup()
                except Exception:
                    pass


# ======================================================================
# DROP ZONE
# ======================================================================
class DropZone(QFrame):
    fileDropped = pyqtSignal(str)

    IDLE = ("QFrame{background:%s;border:2px dashed %s;border-radius:12px;}"
            % (brand.WHITE, brand.SKY_LIGHT))
    HOVER = ("QFrame{background:%s;border:2px dashed %s;border-radius:12px;}"
             % (brand.SKY_PALE, brand.DARK_BLUE))
    LOADED = ("QFrame{background:%s;border:2px solid %s;border-radius:12px;}"
              % (brand.SKY_PALE, brand.SKY))

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setStyleSheet(self.IDLE)
        self.setMinimumHeight(104)

        lay = QVBoxLayout(self)
        lay.setAlignment(Qt.AlignCenter)
        lay.setSpacing(3)

        self.title = QLabel("Drop a PDF or Word file here")
        self.title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.title.setAlignment(Qt.AlignCenter)
        self.title.setStyleSheet(
            "color:%s;border:none;background:transparent;" % brand.DARK_BLUE)

        self.subtitle = QLabel("or click Browse below     \u00b7     supported: .pdf  .docx")
        self.subtitle.setAlignment(Qt.AlignCenter)
        self.subtitle.setStyleSheet(
            "color:%s;border:none;background:transparent;font-size:9.5pt;" % brand.MUTED)

        lay.addWidget(self.title)
        lay.addWidget(self.subtitle)

    def _valid(self, event) -> Optional[str]:
        if not event.mimeData().hasUrls():
            return None
        urls = event.mimeData().urls()
        if len(urls) != 1:
            return None
        path = urls[0].toLocalFile()
        if path and Path(path).suffix.lower() in SUPPORTED:
            return path
        return None

    def dragEnterEvent(self, event):
        if self._valid(event):
            event.acceptProposedAction()
            self.setStyleSheet(self.HOVER)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(self.IDLE)

    def dropEvent(self, event):
        path = self._valid(event)
        if path:
            event.acceptProposedAction()
            self.fileDropped.emit(path)
        else:
            self.setStyleSheet(self.IDLE)

    def show_file(self, path: str):
        self.setStyleSheet(self.LOADED)
        self.title.setText(Path(path).name)
        self.subtitle.setText("Ready to extract     \u00b7     %s file"
                              % Path(path).suffix.upper().lstrip("."))

    def reset(self):
        self.setStyleSheet(self.IDLE)
        self.title.setText("Drop a PDF or Word file here")
        self.subtitle.setText("or click Browse below     \u00b7     supported: .pdf  .docx")


# ======================================================================
# MAIN WINDOW
# ======================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{brand.APP_NAME}  \u00b7  {brand.APP_TAGLINE}")
        self.setWindowIcon(brand.app_icon())
        self.resize(900, 700)
        self.setMinimumSize(760, 560)

        self.input_file: Optional[str] = None
        self.template_file: Optional[str] = None
        self.include_noise = False
        self.keep_markup = False
        self.output_dir: str = str(Path.home() / f"{brand.APP_NAME}_Reports")
        self.last_report: Optional[str] = None
        self.thread: Optional[QThread] = None
        self.worker: Optional[ProcessorWorker] = None

        self.bridge = LogBridge()
        self.bridge.message.connect(self._append_log)

        self._build_ui()
        self._attach_logger()
        self.log(f"{brand.APP_NAME} v{brand.APP_VERSION} ready.")

    # ------------------------------------------------------------------
    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)
        outer.addWidget(BrandHeader())

        body = QWidget()
        lay = QVBoxLayout(body)
        lay.setContentsMargins(22, 14, 22, 12)
        lay.setSpacing(11)
        outer.addWidget(body, 1)

        intro = QLabel("Every annotation \u2014 reference, author, date, type, text, "
                       "a close-up snapshot and a page overview \u2014 in one Excel report.")
        intro.setWordWrap(True)
        intro.setStyleSheet("color:%s;font-size:10pt;" % brand.MUTED)
        lay.addWidget(intro)

        self.drop = DropZone()
        sh = QGraphicsDropShadowEffect(blurRadius=18, xOffset=0, yOffset=3)
        sh.setColor(QColor(11, 37, 69, 38))
        self.drop.setGraphicsEffect(sh)
        self.drop.fileDropped.connect(self.set_input_file)
        lay.addWidget(self.drop)

        row = QHBoxLayout()
        row.setSpacing(9)
        self.browse_btn = QPushButton("Browse for document")
        self.browse_btn.setStyleSheet(brand.DARK_BTN)
        self.browse_btn.setMinimumHeight(35)
        self.browse_btn.setCursor(Qt.PointingHandCursor)
        self.browse_btn.clicked.connect(self.pick_input)
        row.addWidget(self.browse_btn)

        self.clear_btn = QPushButton("Clear")
        self.clear_btn.setStyleSheet(brand.SUBTLE_BTN)
        self.clear_btn.setMinimumHeight(35)
        self.clear_btn.setMaximumWidth(110)
        self.clear_btn.setCursor(Qt.PointingHandCursor)
        self.clear_btn.clicked.connect(self.clear_input)
        row.addWidget(self.clear_btn)
        lay.addLayout(row)

        # ---- single-line output row (replaces the Options panel) ------
        out_row = QHBoxLayout()
        out_row.setSpacing(6)
        cap = QLabel("Save to:")
        cap.setStyleSheet("color:%s;font-weight:600;" % brand.DARK_BLUE)
        out_row.addWidget(cap)

        self.out_label = QLabel(self.output_dir)
        self.out_label.setStyleSheet("color:%s;" % brand.MUTED)
        self.out_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.out_label.setToolTip(self.output_dir)
        out_row.addWidget(self.out_label, 1)

        self.change_btn = QPushButton("Change")
        self.change_btn.setStyleSheet(brand.LINK_BTN)
        self.change_btn.setCursor(Qt.PointingHandCursor)
        self.change_btn.clicked.connect(self.pick_output_dir)
        out_row.addWidget(self.change_btn)

        self.settings_btn = QPushButton("Advanced")
        self.settings_btn.setStyleSheet(brand.LINK_BTN)
        self.settings_btn.setCursor(Qt.PointingHandCursor)
        self.settings_btn.setToolTip("Excel template and extraction switches")
        self.settings_btn.clicked.connect(self.open_settings)
        out_row.addWidget(self.settings_btn)
        lay.addLayout(out_row)

        # ---- action ---------------------------------------------------
        act = QHBoxLayout()
        act.setSpacing(9)
        self.run_btn = QPushButton("Extract Annotations")
        self.run_btn.setMinimumHeight(46)
        self.run_btn.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.run_btn.setStyleSheet(brand.PRIMARY_BTN)
        self.run_btn.setCursor(Qt.PointingHandCursor)
        rs = QGraphicsDropShadowEffect(blurRadius=16, xOffset=0, yOffset=3)
        rs.setColor(QColor(79, 163, 217, 110))
        self.run_btn.setGraphicsEffect(rs)
        self.run_btn.clicked.connect(self.start)
        act.addWidget(self.run_btn, 1)

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setMinimumHeight(46)
        self.cancel_btn.setMaximumWidth(130)
        self.cancel_btn.setFont(QFont("Segoe UI", 11, QFont.Bold))
        self.cancel_btn.setStyleSheet(brand.DANGER_BTN)
        self.cancel_btn.setCursor(Qt.PointingHandCursor)
        self.cancel_btn.setVisible(False)
        self.cancel_btn.clicked.connect(self.cancel_run)
        act.addWidget(self.cancel_btn)
        lay.addLayout(act)

        self.bar = QProgressBar()
        self.bar.setRange(0, 100)
        self.bar.setValue(0)
        lay.addWidget(self.bar)

        self.status = QLabel("Ready")
        self.status.setStyleSheet("color:%s;font-weight:600;" % brand.MUTED)
        lay.addWidget(self.status)

        ll = QLabel("Activity log")
        ll.setStyleSheet("font-weight:700;color:%s;" % brand.DARK_BLUE)
        lay.addWidget(ll)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.document().setMaximumBlockCount(MAX_LOG_BLOCKS)
        self.log_box.setMinimumHeight(90)
        self.log_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        lay.addWidget(self.log_box, 1)

        foot = QHBoxLayout()
        foot.setSpacing(9)
        self.open_report_btn = QPushButton("Open Report")
        self.open_report_btn.setStyleSheet(brand.PRIMARY_BTN)
        self.open_report_btn.setMinimumHeight(34)
        self.open_report_btn.setEnabled(False)
        self.open_report_btn.setCursor(Qt.PointingHandCursor)
        self.open_report_btn.clicked.connect(self.open_report)
        foot.addWidget(self.open_report_btn)

        self.open_folder_btn = QPushButton("Open Output Folder")
        self.open_folder_btn.setStyleSheet(brand.SUBTLE_BTN)
        self.open_folder_btn.setMinimumHeight(34)
        self.open_folder_btn.setCursor(Qt.PointingHandCursor)
        self.open_folder_btn.clicked.connect(self.open_folder)
        foot.addWidget(self.open_folder_btn)
        lay.addLayout(foot)

        credit = QLabel(f"{brand.APP_NAME} v{brand.APP_VERSION}   \u00b7   "
                        f"{brand.APP_TAGLINE}   \u00b7   Runs fully offline")
        credit.setAlignment(Qt.AlignCenter)
        credit.setStyleSheet("color:%s;font-size:8.5pt;" % brand.MUTED)
        lay.addWidget(credit)

    # ------------------------------------------------------------------
    def _attach_logger(self):
        h = BridgeHandler(self.bridge)
        h.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(h)

    def _append_log(self, message: str):
        """Always runs on the GUI thread."""
        self.log_box.append(message)

    def log(self, message: str):
        self.bridge.message.emit(message)

    # ------------------------------------------------------------------
    def set_input_file(self, path: str):
        self.input_file = path
        self.drop.show_file(path)
        self.status.setText(f"Selected: {Path(path).name}")
        self.status.setStyleSheet("color:%s;font-weight:600;" % brand.DARK_BLUE)
        self.log(f"Selected: {path}")

    def pick_input(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select a PDF or Word document", str(Path.home()),
            "Documents (*.pdf *.docx);;PDF files (*.pdf);;Word files (*.docx)")
        if path:
            self.set_input_file(path)

    def clear_input(self):
        self.input_file = None
        self.drop.reset()
        self.status.setText("Ready")
        self.status.setStyleSheet("color:%s;font-weight:600;" % brand.MUTED)
        self.bar.setValue(0)

    def pick_output_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Choose output folder",
                                             self.output_dir)
        if d:
            self.output_dir = d
            self.out_label.setText(d)
            self.out_label.setToolTip(d)

    def open_settings(self):
        dlg = SettingsDialog(self, self.template_file,
                             self.include_noise, self.keep_markup)
        if dlg.exec_() == QDialog.Accepted:
            self.template_file, self.include_noise, self.keep_markup = dlg.values()
            bits = []
            if self.template_file:
                bits.append(f"template: {Path(self.template_file).name}")
            if self.keep_markup:
                bits.append("including text-less markup")
            if self.include_noise:
                bits.append("including links/popups")
            self.log("Settings updated" + (f" ({', '.join(bits)})" if bits else ""))

    # ------------------------------------------------------------------
    def _set_busy(self, busy: bool):
        for w in (self.run_btn, self.browse_btn, self.clear_btn,
                  self.change_btn, self.settings_btn):
            w.setEnabled(not busy)
        self.drop.setAcceptDrops(not busy)
        self.cancel_btn.setVisible(busy)
        self.cancel_btn.setEnabled(busy)
        self.run_btn.setText("Working..." if busy else "Extract Annotations")

    def start(self):
        if not self.input_file:
            QMessageBox.warning(self, "No file selected",
                                "Please choose a PDF or Word document first.")
            return
        if not os.path.exists(self.input_file):
            QMessageBox.critical(self, "File missing",
                                 "That file no longer exists at the saved path.")
            self.clear_input()
            return

        self._set_busy(True)
        self.open_report_btn.setEnabled(False)
        self.bar.setValue(0)
        self.status.setText("Starting...")
        self.status.setStyleSheet("color:%s;font-weight:600;" % brand.WARNING)
        self.log("-" * 58)

        self.thread = QThread()
        self.worker = ProcessorWorker(self.input_file, self.output_dir,
                                      self.template_file, self.include_noise,
                                      self.keep_markup)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.cancelled.connect(self.on_cancelled)

        for sig in (self.worker.finished, self.worker.failed, self.worker.cancelled):
            sig.connect(self.thread.quit)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self._thread_done)
        self.thread.start()

    def cancel_run(self):
        if self.worker is not None:
            self.worker.request_cancel()
            self.cancel_btn.setEnabled(False)
            self.status.setText("Cancelling...")
            self.status.setStyleSheet("color:%s;font-weight:600;" % brand.DANGER)

    def _thread_done(self):
        self.thread = None
        self.worker = None

    # ------------------------------------------------------------------
    def on_progress(self, message: str, pct: int):
        self.status.setText(message)
        self.status.setStyleSheet("color:%s;font-weight:600;" % brand.WARNING)
        if pct >= 0:
            self.bar.setValue(min(100, max(0, pct)))
        self.log(message)

    def on_finished(self, out_path: str, count: int, skipped: int):
        self.bar.setValue(100)
        self.last_report = out_path
        self.status.setText(f"Done \u2014 {count} annotation(s) exported.")
        self.status.setStyleSheet("color:%s;font-weight:700;" % brand.SUCCESS)
        self.log(f"SUCCESS: {count} annotation(s) -> {out_path}")

        extra = ""
        if skipped:
            extra = (f"\n\n{skipped} shape/markup object(s) without reviewer text "
                     f"were skipped.\nTurn this off under Advanced if you need them.")
            self.log(f"Skipped {skipped} markup object(s) with no text.")

        self._set_busy(False)
        self.open_report_btn.setEnabled(True)
        QMessageBox.information(
            self, f"{brand.APP_NAME} \u2014 report ready",
            f"{count} annotation(s) exported.\n\nSaved to:\n{out_path}{extra}")

    def on_failed(self, message: str):
        self.bar.setValue(0)
        self.status.setText("Failed")
        self.status.setStyleSheet("color:%s;font-weight:700;" % brand.DANGER)
        self.log(f"ERROR: {message}")
        self._set_busy(False)
        QMessageBox.critical(self, f"{brand.APP_NAME} \u2014 could not complete", message)

    def on_cancelled(self):
        self.bar.setValue(0)
        self.status.setText("Cancelled")
        self.status.setStyleSheet("color:%s;font-weight:700;" % brand.DANGER)
        self.log("Run cancelled by user.")
        self._set_busy(False)

    # ------------------------------------------------------------------
    @staticmethod
    def _open_path(path: str):
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)  # noqa
            elif sys.platform == "darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def open_report(self):
        if self.last_report and os.path.exists(self.last_report):
            self._open_path(self.last_report)
        else:
            QMessageBox.information(self, "Not available",
                                    "No report has been generated yet.")

    def open_folder(self):
        Path(self.output_dir).mkdir(parents=True, exist_ok=True)
        self._open_path(self.output_dir)

    # ------------------------------------------------------------------
    def closeEvent(self, event):
        if self.thread is not None and self.thread.isRunning():
            reply = QMessageBox.question(
                self, "Still working",
                "Extraction is still running.\n\nCancel it and close?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                event.ignore()
                return
            if self.worker is not None:
                self.worker.request_cancel()
            self.thread.quit()
            if not self.thread.wait(5000):
                self.thread.terminate()
                self.thread.wait(1000)
        event.accept()


# ======================================================================
def main():
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s | %(levelname)s | %(message)s")

    if sys.platform.startswith("win"):
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                f"MaehanSolutions.{brand.APP_NAME}.{brand.APP_VERSION}")
        except Exception:
            pass

    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setApplicationName(brand.APP_NAME)
    app.setOrganizationName(brand.COMPANY)
    app.setWindowIcon(brand.app_icon())
    app.setStyleSheet(brand.APP_STYLESHEET)

    splash = BrandSplash()
    splash.show()
    splash.note("Loading engine...")
    app.processEvents()

    win = MainWindow()

    def reveal():
        win.show()
        win.raise_()
        win.activateWindow()
        splash.finish(win)
        if len(sys.argv) > 1 and Path(sys.argv[1]).suffix.lower() in SUPPORTED:
            win.set_input_file(sys.argv[1])

    QTimer.singleShot(1400, reveal)
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
