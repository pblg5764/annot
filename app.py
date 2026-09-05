"""
AnotAI - Desktop Application
============================
Drop in a PDF or Word file, get an Excel report of every annotation -
author, date, type, text, and a cropped snapshot showing exactly where
it sits on the page.

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
    QMessageBox, QCheckBox, QFrame, QSizePolicy, QGroupBox, QSplashScreen,
    QGraphicsDropShadowEffect
)

import brand
from annotation_processor import AnnotationExtractor

SUPPORTED = (".pdf", ".docx")

logger = logging.getLogger(brand.APP_NAME)
logger.setLevel(logging.INFO)


# ======================================================================
# SPLASH SCREEN
# ======================================================================
def build_splash_pixmap(w: int = 620, h: int = 360) -> QPixmap:
    """Dark-blue branded splash with the AnotAI mark and company line."""
    app = QApplication.instance()
    ratio = app.primaryScreen().devicePixelRatio() if app else 1.0

    pm = QPixmap(int(w * ratio), int(h * ratio))
    pm.setDevicePixelRatio(ratio)
    pm.fill(Qt.transparent)

    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing, True)
    p.setRenderHint(QPainter.SmoothPixmapTransform, True)

    # navy gradient panel
    grad = QLinearGradient(0, 0, w, h)
    grad.setColorAt(0.0, QColor(brand.DARK_BLUE))
    grad.setColorAt(0.55, QColor(brand.DARK_BLUE_2))
    grad.setColorAt(1.0, QColor(brand.DARK_BLUE))
    p.setBrush(QBrush(grad))
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(QRectF(0, 0, w, h), 16, 16)

    # sky accent arcs
    p.setBrush(QColor(79, 163, 217, 34))
    p.drawEllipse(QRectF(w - 190, -110, 300, 300))
    p.setBrush(QColor(127, 196, 236, 22))
    p.drawEllipse(QRectF(-90, h - 150, 240, 240))

    # sky rule
    p.setPen(QPen(QColor(brand.SKY), 3))
    p.drawLine(int(w / 2 - 46), 236, int(w / 2 + 46), 236)

    # logo mark
    mark = brand.logo_pixmap(104)
    p.drawPixmap(int((w - mark.width() / mark.devicePixelRatio()) / 2), 40, mark)

    # product name
    f_name = QFont("Segoe UI", 34, QFont.Bold)
    f_name.setLetterSpacing(QFont.AbsoluteSpacing, 1.5)
    p.setFont(f_name)
    p.setPen(QColor(brand.WHITE))
    fm = QFontMetrics(f_name)
    p.drawText(int((w - fm.horizontalAdvance(brand.APP_NAME)) / 2), 218,
               brand.APP_NAME)

    # strapline
    f_strap = QFont("Segoe UI", 11)
    f_strap.setLetterSpacing(QFont.AbsoluteSpacing, 2.6)
    p.setFont(f_strap)
    p.setPen(QColor(brand.SKY_LIGHT))
    strap = brand.APP_STRAP.upper()
    fm = QFontMetrics(f_strap)
    p.drawText(int((w - fm.horizontalAdvance(strap)) / 2), 266, strap)

    # powered by
    f_by = QFont("Segoe UI", 11, QFont.Bold)
    p.setFont(f_by)
    p.setPen(QColor(brand.CREAM))
    fm = QFontMetrics(f_by)
    p.drawText(int((w - fm.horizontalAdvance(brand.APP_TAGLINE)) / 2), 306,
               brand.APP_TAGLINE)

    # version
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
    """Navy banner across the top of the main window."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(92)
        self.setStyleSheet("QFrame{border:none;}")

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

        mark = brand.logo_pixmap(54)
        p.drawPixmap(18, 19, mark)

        f_name = QFont("Segoe UI", 20, QFont.Bold)
        f_name.setLetterSpacing(QFont.AbsoluteSpacing, 1.0)
        p.setFont(f_name)
        p.setPen(QColor(brand.WHITE))
        p.drawText(86, 46, brand.APP_NAME)

        f_sub = QFont("Segoe UI", 9)
        f_sub.setLetterSpacing(QFont.AbsoluteSpacing, 1.4)
        p.setFont(f_sub)
        p.setPen(QColor(brand.SKY_LIGHT))
        p.drawText(88, 66, brand.APP_STRAP.upper())

        f_by = QFont("Segoe UI", 9, QFont.Bold)
        p.setFont(f_by)
        p.setPen(QColor(brand.CREAM))
        fm = QFontMetrics(f_by)
        p.drawText(w - fm.horizontalAdvance(brand.APP_TAGLINE) - 20, 50,
                   brand.APP_TAGLINE)

        f_v = QFont("Segoe UI", 8)
        p.setFont(f_v)
        p.setPen(QColor(127, 196, 236, 170))
        ver = f"v{brand.APP_VERSION}"
        fm = QFontMetrics(f_v)
        p.drawText(w - fm.horizontalAdvance(ver) - 20, 68, ver)
        p.end()


# ======================================================================
# BACKGROUND WORKER
# ======================================================================
class ProcessorWorker(QObject):
    progress = pyqtSignal(str)
    finished = pyqtSignal(str, int)
    failed = pyqtSignal(str)

    def __init__(self, file_path: str, output_dir: str,
                 template_path: Optional[str], include_noise: bool):
        super().__init__()
        self.file_path = file_path
        self.output_dir = output_dir
        self.template_path = template_path
        self.include_noise = include_noise

    def run(self):
        extractor = None
        try:
            self.progress.emit("Starting...")
            extractor = AnnotationExtractor(
                logger=logger, include_noise=self.include_noise,
                progress_cb=self.progress.emit)
            out_path, count = extractor.process(
                self.file_path, output_dir=self.output_dir,
                template_path=self.template_path)
            self.finished.emit(out_path, count)
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
        self.setMinimumHeight(116)

        lay = QVBoxLayout(self)
        lay.setAlignment(Qt.AlignCenter)
        lay.setSpacing(4)

        self.title = QLabel("Drop a PDF or Word file here")
        self.title.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.title.setAlignment(Qt.AlignCenter)
        self.title.setStyleSheet(
            "color:%s;border:none;background:transparent;" % brand.DARK_BLUE)

        self.subtitle = QLabel("or click Browse below     ·     supported: .pdf  .docx")
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
        self.subtitle.setText(
            "Ready to extract     ·     %s file"
            % Path(path).suffix.upper().lstrip("."))

    def reset(self):
        self.setStyleSheet(self.IDLE)
        self.title.setText("Drop a PDF or Word file here")
        self.subtitle.setText("or click Browse below     ·     supported: .pdf  .docx")


# ======================================================================
# MAIN WINDOW
# ======================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{brand.APP_NAME}  ·  {brand.APP_TAGLINE}")
        self.setWindowIcon(brand.app_icon())
        self.resize(920, 830)
        self.setMinimumSize(800, 680)

        self.input_file: Optional[str] = None
        self.template_file: Optional[str] = None
        self.output_dir: str = str(Path.home() / f"{brand.APP_NAME}_Reports")
        self.last_report: Optional[str] = None
        self.thread: Optional[QThread] = None
        self.worker: Optional[ProcessorWorker] = None

        self._build_ui()
        self._attach_logger()
        self.log(f"{brand.APP_NAME} v{brand.APP_VERSION} ready.")
        self.log("Drop a PDF or Word document above to begin.")

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
        lay.setContentsMargins(24, 18, 24, 18)
        lay.setSpacing(13)
        outer.addWidget(body, 1)

        intro = QLabel(
            "Extract every annotation — author, date, type, text and a snapshot "
            "of its exact location — into one Excel report.")
        intro.setWordWrap(True)
        intro.setStyleSheet("color:%s;font-size:10.5pt;" % brand.MUTED)
        lay.addWidget(intro)

        self.drop = DropZone()
        shadow = QGraphicsDropShadowEffect(blurRadius=18, xOffset=0, yOffset=3)
        shadow.setColor(QColor(11, 37, 69, 38))
        self.drop.setGraphicsEffect(shadow)
        self.drop.fileDropped.connect(self.set_input_file)
        lay.addWidget(self.drop)

        row = QHBoxLayout()
        row.setSpacing(10)
        browse = QPushButton("Browse for document")
        browse.setStyleSheet(brand.DARK_BTN)
        browse.setMinimumHeight(36)
        browse.setCursor(Qt.PointingHandCursor)
        browse.clicked.connect(self.pick_input)
        row.addWidget(browse)

        clear = QPushButton("Clear")
        clear.setStyleSheet(brand.SUBTLE_BTN)
        clear.setMinimumHeight(36)
        clear.setMaximumWidth(120)
        clear.setCursor(Qt.PointingHandCursor)
        clear.clicked.connect(self.clear_input)
        row.addWidget(clear)
        lay.addLayout(row)

        opts = QGroupBox("Options")
        ol = QVBoxLayout(opts)
        ol.setSpacing(9)

        out_row = QHBoxLayout()
        lbl_out = QLabel("Save reports to:")
        lbl_out.setStyleSheet("font-weight:600;color:%s;" % brand.DARK_BLUE)
        lbl_out.setMinimumWidth(150)
        out_row.addWidget(lbl_out)
        self.out_label = QLabel(self.output_dir)
        self.out_label.setStyleSheet("color:%s;" % brand.INK)
        self.out_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        out_row.addWidget(self.out_label, 1)
        out_btn = QPushButton("Change")
        out_btn.setStyleSheet(brand.SUBTLE_BTN)
        out_btn.setMaximumWidth(110)
        out_btn.setCursor(Qt.PointingHandCursor)
        out_btn.clicked.connect(self.pick_output_dir)
        out_row.addWidget(out_btn)
        ol.addLayout(out_row)

        tpl_row = QHBoxLayout()
        lbl_tpl = QLabel("Excel template:")
        lbl_tpl.setStyleSheet("font-weight:600;color:%s;" % brand.DARK_BLUE)
        lbl_tpl.setMinimumWidth(150)
        tpl_row.addWidget(lbl_tpl)
        self.tpl_label = QLabel("None — a fresh branded report will be created")
        self.tpl_label.setStyleSheet("color:%s;" % brand.MUTED)
        self.tpl_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        tpl_row.addWidget(self.tpl_label, 1)
        tpl_btn = QPushButton("Select")
        tpl_btn.setStyleSheet(brand.SUBTLE_BTN)
        tpl_btn.setMaximumWidth(110)
        tpl_btn.setCursor(Qt.PointingHandCursor)
        tpl_btn.clicked.connect(self.pick_template)
        tpl_row.addWidget(tpl_btn)
        tpl_clear = QPushButton("Remove")
        tpl_clear.setStyleSheet(brand.SUBTLE_BTN)
        tpl_clear.setMaximumWidth(110)
        tpl_clear.setCursor(Qt.PointingHandCursor)
        tpl_clear.clicked.connect(self.clear_template)
        tpl_row.addWidget(tpl_clear)
        ol.addLayout(tpl_row)

        self.chk_noise = QCheckBox(
            "Also capture Link and Popup objects (usually not review comments)")
        ol.addWidget(self.chk_noise)
        lay.addWidget(opts)

        self.run_btn = QPushButton("Extract Annotations")
        self.run_btn.setMinimumHeight(48)
        self.run_btn.setFont(QFont("Segoe UI", 12, QFont.Bold))
        self.run_btn.setStyleSheet(brand.PRIMARY_BTN)
        self.run_btn.setCursor(Qt.PointingHandCursor)
        run_shadow = QGraphicsDropShadowEffect(blurRadius=16, xOffset=0, yOffset=3)
        run_shadow.setColor(QColor(79, 163, 217, 110))
        self.run_btn.setGraphicsEffect(run_shadow)
        self.run_btn.clicked.connect(self.start)
        lay.addWidget(self.run_btn)

        self.bar = QProgressBar()
        lay.addWidget(self.bar)

        self.status = QLabel("Ready")
        self.status.setStyleSheet("color:%s;font-weight:600;" % brand.MUTED)
        lay.addWidget(self.status)

        log_lbl = QLabel("Activity log")
        log_lbl.setStyleSheet("font-weight:700;color:%s;" % brand.DARK_BLUE)
        lay.addWidget(log_lbl)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        lay.addWidget(self.log_box, 1)

        foot = QHBoxLayout()
        foot.setSpacing(10)
        self.open_report_btn = QPushButton("Open Report")
        self.open_report_btn.setStyleSheet(brand.PRIMARY_BTN)
        self.open_report_btn.setMinimumHeight(36)
        self.open_report_btn.setEnabled(False)
        self.open_report_btn.setCursor(Qt.PointingHandCursor)
        self.open_report_btn.clicked.connect(self.open_report)
        foot.addWidget(self.open_report_btn)

        self.open_folder_btn = QPushButton("Open Output Folder")
        self.open_folder_btn.setStyleSheet(brand.SUBTLE_BTN)
        self.open_folder_btn.setMinimumHeight(36)
        self.open_folder_btn.setCursor(Qt.PointingHandCursor)
        self.open_folder_btn.clicked.connect(self.open_folder)
        foot.addWidget(self.open_folder_btn)
        lay.addLayout(foot)

        credit = QLabel(f"{brand.APP_NAME} v{brand.APP_VERSION}   ·   "
                        f"{brand.APP_TAGLINE}   ·   Runs fully offline")
        credit.setAlignment(Qt.AlignCenter)
        credit.setStyleSheet("color:%s;font-size:8.5pt;" % brand.MUTED)
        lay.addWidget(credit)

    # ------------------------------------------------------------------
    def _attach_logger(self):
        parent = self

        class GuiHandler(logging.Handler):
            def emit(self, record):
                try:
                    parent.log(self.format(record))
                except Exception:
                    pass

        h = GuiHandler()
        h.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(h)

    def log(self, message: str):
        self.log_box.append(message)
        sb = self.log_box.verticalScrollBar()
        sb.setValue(sb.maximum())

    # ------------------------------------------------------------------
    def set_input_file(self, path: str):
        self.input_file = path
        self.drop.show_file(path)
        self.status.setText(f"Selected: {Path(path).name}")
        self.status.setStyleSheet("color:%s;font-weight:600;" % brand.DARK_BLUE)
        self.log(f"Selected file: {path}")

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
        d = QFileDialog.getExistingDirectory(
            self, "Choose output folder", self.output_dir)
        if d:
            self.output_dir = d
            self.out_label.setText(d)
            self.log(f"Output folder set to: {d}")

    def pick_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select an Excel template", str(Path.home()),
            "Excel files (*.xlsx)")
        if path:
            self.template_file = path
            self.tpl_label.setText(Path(path).name)
            self.tpl_label.setStyleSheet("color:%s;" % brand.INK)
            self.log(f"Template selected: {path}")

    def clear_template(self):
        self.template_file = None
        self.tpl_label.setText("None — a fresh branded report will be created")
        self.tpl_label.setStyleSheet("color:%s;" % brand.MUTED)

    # ------------------------------------------------------------------
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

        self.run_btn.setEnabled(False)
        self.open_report_btn.setEnabled(False)
        self.bar.setRange(0, 0)
        self.status.setText("Processing...")
        self.status.setStyleSheet("color:%s;font-weight:600;" % brand.WARNING)
        self.log("-" * 62)

        self.thread = QThread()
        self.worker = ProcessorWorker(
            self.input_file, self.output_dir,
            self.template_file, self.chk_noise.isChecked())
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.progress.connect(self.on_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.failed.connect(self.on_failed)
        self.worker.finished.connect(self.thread.quit)
        self.worker.failed.connect(self.thread.quit)
        self.thread.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def on_progress(self, message: str):
        self.status.setText(message)
        self.status.setStyleSheet("color:%s;font-weight:600;" % brand.WARNING)

    def _reset_bar(self, value: int):
        self.bar.setRange(0, 100)
        self.bar.setValue(value)

    def on_finished(self, out_path: str, count: int):
        self._reset_bar(100)
        self.last_report = out_path
        self.status.setText(f"Done — {count} annotation(s) exported.")
        self.status.setStyleSheet("color:%s;font-weight:700;" % brand.SUCCESS)
        self.log(f"SUCCESS: {count} annotation(s) written to {out_path}")
        self.run_btn.setEnabled(True)
        self.open_report_btn.setEnabled(True)
        QMessageBox.information(
            self, f"{brand.APP_NAME} — report ready",
            f"{count} annotation(s) exported.\n\nSaved to:\n{out_path}")

    def on_failed(self, message: str):
        self._reset_bar(0)
        self.status.setText("Failed")
        self.status.setStyleSheet("color:%s;font-weight:700;" % brand.DANGER)
        self.log(f"ERROR: {message}")
        self.run_btn.setEnabled(True)
        QMessageBox.critical(self, f"{brand.APP_NAME} — could not complete", message)

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
                self, "Still working", "Extraction is still running. Close anyway?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                event.ignore()
                return
            self.thread.quit()
            self.thread.wait(3000)
        event.accept()


# ======================================================================
def main():
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s | %(levelname)s | %(message)s")

    # Correct taskbar grouping + icon on Windows
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
        splash.note("Ready")
        win.show()
        splash.finish(win)
        if len(sys.argv) > 1 and Path(sys.argv[1]).suffix.lower() in SUPPORTED:
            win.set_input_file(sys.argv[1])

    QTimer.singleShot(1900, reveal)
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
