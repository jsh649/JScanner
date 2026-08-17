import sys
import os
import re
from datetime import datetime
from pathlib import Path
import threading
import time
import subprocess
import platform

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog

import cv2
import numpy as np
from PIL import Image
import tempfile

# WIA library for Windows scanner support
try:
    import win32com.client
    import pythoncom
    WIA_AVAILABLE = True
except ImportError:
    WIA_AVAILABLE = False
    print("Warning: pywin32 not installed. Scanner selection disabled.")


class ScannerDevice:
    """Scanner device information"""
    def __init__(self, device_id, name, device_type="WIA"):
        self.device_id = device_id
        self.name = name
        self.device_type = device_type
        
    def __str__(self):
        return self.name


class ScanWorker(QThread):
    """Thread for scanning without freezing UI"""
    progress_updated = pyqtSignal(int)
    status_updated = pyqtSignal(str)
    scan_finished = pyqtSignal(str, str)  # filepath, message
    scan_error = pyqtSignal(str)
    
    # Preset name -> (scan DPI, JPEG save quality)
    QUALITY_PRESETS = {
        "compact": (150, 75),
        "balanced": (200, 85),
        "high": (300, 95),
    }
    
    def __init__(self, save_dir, national_code, full_name, doc_type, 
                 custom_doc_type=None, scanner_device=None, quality_preset="balanced"):
        super().__init__()
        self.save_dir = save_dir
        self.national_code = national_code
        self.full_name = full_name
        self.doc_type = doc_type
        self.custom_doc_type = custom_doc_type
        self.scanner_device = scanner_device
        self.is_running = True
        self.dpi, self.jpeg_quality = self.QUALITY_PRESETS.get(
            quality_preset, self.QUALITY_PRESETS["balanced"]
        )
        
    def run(self):
        try:
            self.status_updated.emit("در حال آماده‌سازی اسکنر...")
            self.progress_updated.emit(10)
            
            # Simulate scanner initialization if no scanner selected
            if not self.scanner_device or not WIA_AVAILABLE:
                time.sleep(0.5)
            else:
                # Initialize real scanner
                self.status_updated.emit(f"اتصال به اسکنر: {self.scanner_device.name}")
                self.progress_updated.emit(15)
                time.sleep(1)
            
            # Create directory structure
            self.status_updated.emit("ایجاد پوشه‌های مورد نیاز...")
            self.progress_updated.emit(20)
            
            # Create national code folder - all documents for this person
            # are saved directly inside this single folder
            national_folder = self.save_dir / self.national_code
            national_folder.mkdir(parents=True, exist_ok=True)
            
            self.status_updated.emit("آماده‌سازی برای اسکن...")
            self.progress_updated.emit(30)
            
            # Generate filename (document type is encoded in the filename)
            filename = self.generate_filename()
            filepath = national_folder / filename
            
            # Scan the document
            if self.scanner_device and WIA_AVAILABLE:
                self.status_updated.emit("در حال اسکن سند... (این عمل چند ثانیه طول می‌کشد)")
                self.progress_updated.emit(40)
                
                # Perform actual scan
                scanned_image = self.scan_with_wia()
                
                if scanned_image is None:
                    raise Exception("اسکن با خطا مواجه شد. لطفاً اسکنر را بررسی کنید.")
                
                # Save scanned image (optimize=True + chosen quality keeps
                # size down while preserving readable quality)
                scanned_image.save(
                    str(filepath), 'JPEG',
                    quality=self.jpeg_quality, optimize=True
                )
                
            else:
                # Simulate scanning process for demo
                self.status_updated.emit("در حال اسکن سند... (حالت شبیه‌سازی)")
                self.progress_updated.emit(50)
                
                # Simulate scanning time
                for i in range(50, 90, 10):
                    if not self.is_running:
                        return
                    time.sleep(0.3)
                    self.progress_updated.emit(i)
                
                # Create a sample scanned image (for demonstration)
                self.create_sample_image(filepath)
            
            self.progress_updated.emit(95)
            self.status_updated.emit("پردازش نهایی سند...")
            time.sleep(0.3)
            
            self.progress_updated.emit(100)
            self.scan_finished.emit(str(filepath), "اسکن با موفقیت انجام شد!")
            
        except Exception as e:
            self.scan_error.emit(f"خطا در اسکن: {str(e)}")
    
    def scan_with_wia(self):
        """Scan using WIA (Windows Image Acquisition) via WIA.DeviceManager.

        Note: this intentionally avoids WIA.CommonDialog / WIA.ShowScan, which
        depend on the legacy "WIA Automation Layer" (wiaaut.dll). That
        component is not registered by default on Windows 10/11 and causes
        the "(-2147221005, 'Invalid class string', None, None)" error. Using
        WIA.DeviceManager + Item.Transfer works with the WIA drivers that
        ship with Windows out of the box.
        """
        WIA_FORMAT_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        temp_path = None
        try:
            pythoncom.CoInitialize()
            manager = win32com.client.Dispatch("WIA.DeviceManager")

            device = None
            if self.scanner_device:
                # Connect to the specific device the user selected
                for device_info in manager.DeviceInfos:
                    if device_info.DeviceID == self.scanner_device.device_id:
                        device = device_info.Connect()
                        break
                if device is None:
                    raise Exception(
                        "اسکنر انتخاب‌شده یافت نشد. لطفاً دوباره اسکنرها را جستجو کنید."
                    )
            else:
                # No specific scanner chosen: use the first available device
                if manager.DeviceInfos.Count == 0:
                    raise Exception("هیچ اسکنری شناسایی نشد.")
                device = manager.DeviceInfos.Item(1).Connect()

            # Flatbed/scan item is normally the first item
            item = device.Items(1)

            # Best-effort scan settings (not all scanners support every
            # property, so failures here are ignored)
            self._set_wia_property(item.Properties, "6146", 1)         # Color intent: color
            self._set_wia_property(item.Properties, "6147", self.dpi)  # Horizontal DPI
            self._set_wia_property(item.Properties, "6148", self.dpi)  # Vertical DPI

            # Perform the actual scan/transfer
            image_file = item.Transfer(WIA_FORMAT_JPEG)

            # WIA returns a COM ImageFile object; save it to a temp file and
            # load that with PIL
            temp_path = os.path.join(
                tempfile.gettempdir(), f"wia_scan_{int(time.time())}.jpg"
            )
            if os.path.exists(temp_path):
                os.remove(temp_path)
            image_file.SaveFile(temp_path)

            image = Image.open(temp_path)
            image.load()  # force read into memory before temp file cleanup
            return image

        except Exception as e:
            raise Exception(f"خطا در اسکن با WIA: {str(e)}")
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    @staticmethod
    def _set_wia_property(properties, prop_id, value):
        """Set a WIA item property by ID, ignoring unsupported properties"""
        try:
            properties(prop_id).Value = value
        except Exception:
            pass
    
    def get_doc_folder_name(self):
        """Get document type folder name"""
        doc_type_map = {
            "روی کارت ملی": "Front_Card",
            "پشت کارت ملی": "Back_Card",
            "شناسنامه 1": "Birth_Cert_1",
            "شناسنامه 2": "Birth_Cert_2",
            "شناسنامه 3": "Birth_Cert_3",
            "شناسنامه 4": "Birth_Cert_4",
            "شناسنامه 5": "Birth_Cert_5",
            "سایر": self.custom_doc_type if self.custom_doc_type else "Other"
        }
        return doc_type_map.get(self.doc_type, "Other")
    
    def generate_filename(self):
        """Generate unique filename with counter if exists"""
        # Clean national code for filename
        clean_national = re.sub(r'[^\w]', '', self.national_code)
        
        # Base filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = f"{clean_national}_{timestamp}"
        
        # Add document type to filename
        doc_short = self.doc_type.replace(" ", "_")
        if self.doc_type == "سایر" and self.custom_doc_type:
            doc_short = self.custom_doc_type.replace(" ", "_")
            
        base_name = f"{base_name}_{doc_short}"
        
        # Check if file exists and add counter
        national_folder = self.save_dir / self.national_code
        counter = 1
        final_name = f"{base_name}.jpg"
        
        while (national_folder / final_name).exists():
            final_name = f"{base_name}_{counter}.jpg"
            counter += 1
            
        return final_name
    
    def create_sample_image(self, filepath):
        """Create a sample image for demonstration"""
        # Create a blank A4 size image with some text
        width, height = 1240, 1754  # A4 at 150 DPI
        
        # Create white background
        img = np.ones((height, width, 3), dtype=np.uint8) * 255
        
        # Add some text for demonstration
        font = cv2.FONT_HERSHEY_SIMPLEX
        
        # Add document info
        texts = [
            ("Document Scan", (width//2 - 100, 100)),
            (f"National Code: {self.national_code}", (50, 200)),
            (f"Name: {self.full_name if self.full_name else 'Not provided'}", (50, 250)),
            (f"Document Type: {self.doc_type}", (50, 300)),
            (f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", (50, 350)),
            (f"Folder: {self.get_doc_folder_name()}", (50, 400)),
            ("Sample Document", (width//2 - 80, height//2)),
        ]
        
        for text, pos in texts:
            cv2.putText(img, text, pos, font, 0.8, (0, 0, 0), 2)
        
        # Add a border
        cv2.rectangle(img, (10, 10), (width-10, height-10), (0, 0, 0), 2)
        
        # Add some decorative lines
        cv2.line(img, (10, 150), (width-10, 150), (200, 200, 200), 1)
        cv2.line(img, (10, 450), (width-10, 450), (200, 200, 200), 1)
        
        # Save image
        cv2.imwrite(str(filepath), img, [cv2.IMWRITE_JPEG_QUALITY, 95])


class ScannerUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.scan_worker = None
        self.save_path = Path.home() / "Documents" / "Scanned_Documents"
        self.selected_scanner = None
        self.scanner_list = []
        # Persists last-used settings (save path, scanner, doc type, quality)
        # across app restarts. National code and name are never saved here.
        self.settings = QSettings("JScannerApp", "DocumentScanner")
        self.init_ui()
        self.scan_for_scanners()
        self.load_saved_settings()
        
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle("اسکنر اسناد - Document Scanner")
        self.setGeometry(100, 100, 800, 650)
        
        # Set application icon
        self.setWindowIcon(self.create_icon())
        
        # Central widget with scrolling
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Create menu bar
        self.create_menu_bar()
        
        # Title
        title_label = QLabel("اسکنر اسناد")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont("Arial", 18, QFont.Bold)
        title_label.setFont(title_font)
        main_layout.addWidget(title_label)
        
        # Settings group
        settings_group = QGroupBox("تنظیمات اسکن")
        settings_layout = QGridLayout()
        settings_group.setLayout(settings_layout)
        
        # Save path
        settings_layout.addWidget(QLabel("مسیر ذخیره‌سازی:"), 0, 0)
        self.save_path_edit = QLineEdit(str(self.save_path))
        self.save_path_edit.setReadOnly(True)
        settings_layout.addWidget(self.save_path_edit, 0, 1)
        
        browse_btn = QPushButton("انتخاب مسیر")
        browse_btn.clicked.connect(self.browse_save_path)
        settings_layout.addWidget(browse_btn, 0, 2)
        
        # Scanner selection
        settings_layout.addWidget(QLabel("اسکنر:"), 1, 0)
        self.scanner_combo = QComboBox()
        self.scanner_combo.addItem("-- انتخاب اسکنر --")
        self.scanner_combo.currentIndexChanged.connect(self.on_scanner_selected)
        settings_layout.addWidget(self.scanner_combo, 1, 1)
        
        refresh_scanner_btn = QPushButton("🔍")
        refresh_scanner_btn.setToolTip("جستجوی اسکنرها")
        refresh_scanner_btn.clicked.connect(self.scan_for_scanners)
        refresh_scanner_btn.setMaximumWidth(40)
        settings_layout.addWidget(refresh_scanner_btn, 1, 2)
        
        # A4 default setting
        self.a4_check = QCheckBox("استاندارد A4 (پیش‌فرض)")
        self.a4_check.setChecked(True)
        self.a4_check.setEnabled(False)
        settings_layout.addWidget(self.a4_check, 2, 0, 1, 3)
        
        # Scan quality / file size preset
        settings_layout.addWidget(QLabel("کیفیت اسکن:"), 3, 0)
        self.quality_combo = QComboBox()
        self.quality_combo.addItem("فشرده - کمترین حجم (150dpi)", "compact")
        self.quality_combo.addItem("متعادل - پیشنهادی (200dpi)", "balanced")
        self.quality_combo.addItem("کیفیت بالا - بیشترین حجم (300dpi)", "high")
        self.quality_combo.setCurrentIndex(1)  # Balanced by default
        self.quality_combo.setToolTip(
            "رزولوشن و فشرده‌سازی اسکن را تنظیم می‌کند. حالت متعادل کیفیت "
            "مناسب برای خواندن سند را با حجم فایل پایین‌تر ترکیب می‌کند."
        )
        self.quality_combo.currentIndexChanged.connect(lambda _: self.save_settings())
        settings_layout.addWidget(self.quality_combo, 3, 1, 1, 2)
        
        main_layout.addWidget(settings_group)
        
        # Document info group
        doc_group = QGroupBox("اطلاعات سند")
        doc_layout = QGridLayout()
        doc_group.setLayout(doc_layout)
        
        # National code (required)
        doc_layout.addWidget(QLabel("کد ملی (اجباری):"), 0, 0)
        self.national_code_edit = QLineEdit()
        self.national_code_edit.setPlaceholderText("مثال: 1234567890")
        self.national_code_edit.textChanged.connect(self.validate_national_code)
        doc_layout.addWidget(self.national_code_edit, 0, 1)
        
        # Full name (optional)
        doc_layout.addWidget(QLabel("نام و نام خانوادگی (اختیاری):"), 1, 0)
        self.full_name_edit = QLineEdit()
        self.full_name_edit.setPlaceholderText("مثال: علی محمدی")
        doc_layout.addWidget(self.full_name_edit, 1, 1)
        
        # Document type combo
        doc_layout.addWidget(QLabel("نوع مدرک:"), 2, 0)
        self.doc_type_combo = QComboBox()
        doc_types = [
            "روی کارت ملی",
            "پشت کارت ملی",
            "شناسنامه 1",
            "شناسنامه 2",
            "شناسنامه 3",
            "شناسنامه 4",
            "شناسنامه 5",
            "سایر"
        ]
        self.doc_type_combo.addItems(doc_types)
        self.doc_type_combo.currentTextChanged.connect(self.on_doc_type_changed)
        doc_layout.addWidget(self.doc_type_combo, 2, 1)
        
        # Custom document type
        self.custom_doc_label = QLabel("متن دلخواه:")
        self.custom_doc_label.setVisible(False)
        doc_layout.addWidget(self.custom_doc_label, 3, 0)
        
        self.custom_doc_edit = QLineEdit()
        self.custom_doc_edit.setPlaceholderText("نوع مدرک دلخواه را وارد کنید...")
        self.custom_doc_edit.setVisible(False)
        doc_layout.addWidget(self.custom_doc_edit, 3, 1)
        
        main_layout.addWidget(doc_group)
        
        # Scan button
        scan_btn_layout = QHBoxLayout()
        self.scan_btn = QPushButton("شروع اسکن")
        self.scan_btn.setIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))
        self.scan_btn.setMinimumHeight(50)
        self.scan_btn.clicked.connect(self.start_scan)
        self.scan_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
                padding: 10px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        scan_btn_layout.addWidget(self.scan_btn)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        scan_btn_layout.addWidget(self.progress_bar)
        
        main_layout.addLayout(scan_btn_layout)
        
        # Status label
        self.status_label = QLabel("آماده اسکن - لطفاً سند را روی اسکنر قرار دهید")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("""
            QLabel {
                padding: 10px;
                background-color: #f0f0f0;
                border-radius: 5px;
                border: 1px solid #cccccc;
            }
        """)
        main_layout.addWidget(self.status_label)
        
        # Info label
        info_label = QLabel("نکته: کد ملی اجباری بوده و باید عددی ۱۰ رقمی باشد")
        info_label.setStyleSheet("color: #666666; font-size: 10px;")
        info_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(info_label)
        
        # Apply RTL layout for Persian text
        self.setLayoutDirection(Qt.RightToLeft)
        
        # Set tab order
        self.setTabOrder(self.national_code_edit, self.full_name_edit)
        self.setTabOrder(self.full_name_edit, self.doc_type_combo)
        self.setTabOrder(self.doc_type_combo, self.custom_doc_edit)
        self.setTabOrder(self.custom_doc_edit, self.scan_btn)
        
        # Check WIA availability
        if not WIA_AVAILABLE:
            self.status_label.setText("⚠️ کتابخانه اسکنر نصب نیست. نصب کنید: pip install pywin32")
            self.status_label.setStyleSheet("""
                QLabel {
                    padding: 10px;
                    background-color: #ffffcc;
                    border-radius: 5px;
                    border: 1px solid #ffaa00;
                }
            """)
        
    def create_icon(self):
        """Load the bundled icon.ico if available, otherwise fall back to
        a simple drawn icon so the app still has an icon either way."""
        icon_path = self.get_resource_path("icon.ico")
        if icon_path and os.path.exists(icon_path):
            return QIcon(icon_path)
        
        icon = QPixmap(64, 64)
        icon.fill(Qt.transparent)
        painter = QPainter(icon)
        painter.setPen(QPen(Qt.blue, 3))
        painter.drawRect(10, 10, 44, 44)
        painter.drawLine(10, 25, 54, 25)
        painter.drawLine(10, 40, 54, 40)
        painter.drawLine(25, 10, 25, 54)
        painter.drawLine(40, 10, 40, 54)
        painter.end()
        return QIcon(icon)
    
    @staticmethod
    def get_resource_path(filename):
        """Resolve a resource path that works both when run as a plain
        .py script and when bundled into a PyInstaller --onefile exe."""
        if hasattr(sys, "_MEIPASS"):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, filename)
    
    def create_menu_bar(self):
        """Create menu bar"""
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("فایل")
        
        open_action = QAction("باز کردن پوشه اسکن", self)
        open_action.triggered.connect(self.open_scan_folder)
        file_menu.addAction(open_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("خروج", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Scanner menu
        scanner_menu = menubar.addMenu("اسکنر")
        
        scan_action = QAction("جستجوی اسکنرها", self)
        scan_action.triggered.connect(self.scan_for_scanners)
        scanner_menu.addAction(scan_action)
        
        # Settings menu
        settings_menu = menubar.addMenu("تنظیمات")
        
        settings_action = QAction("تنظیمات پیشرفته", self)
        settings_action.triggered.connect(self.show_settings)
        settings_menu.addAction(settings_action)
        
        # Help menu
        help_menu = menubar.addMenu("راهنما")
        
        about_action = QAction("درباره برنامه", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
    
    def scan_for_scanners(self):
        """Detect available scanners using WIA (simplified method)"""
        self.scanner_combo.clear()
        self.scanner_combo.addItem("-- انتخاب اسکنر --")
        self.scanner_list = []
        
        if not WIA_AVAILABLE:
            self.status_label.setText("⚠️ لطفاً pywin32 را نصب کنید: pip install pywin32")
            return
        
        try:
            self.status_label.setText("در حال جستجوی اسکنرها...")
            QApplication.processEvents()
            
            import pythoncom
            pythoncom.CoInitialize()
            
            # Try to find scanners using WIA DeviceManager
            scanner_found = False
            try:
                manager = win32com.client.Dispatch("WIA.DeviceManager")
                device_infos = manager.DeviceInfos
                
                for device_info in device_infos:
                    try:
                        # Get device name
                        name = device_info.Properties("Name").Value
                        
                        # Try to connect to device
                        try:
                            device = device_info.Connect()
                            if device:
                                # Check if it has scanner capabilities
                                try:
                                    items = device.Items
                                    if items and len(items) > 0:
                                        scanner = ScannerDevice(
                                            device_info.DeviceID,
                                            name
                                        )
                                        self.scanner_list.append(scanner)
                                        self.scanner_combo.addItem(name)
                                        scanner_found = True
                                        print(f"Scanner found: {name}")
                                except:
                                    pass
                        except:
                            pass
                    except:
                        continue
                        
            except Exception as e:
                print(f"Error in scanner detection: {e}")
            
            pythoncom.CoUninitialize()
            
            if scanner_found:
                self.status_label.setText(f"✅ {len(self.scanner_list)} اسکنر پیدا شد.")
            else:
                self.status_label.setText("⚠️ هیچ اسکنری یافت نشد. لطفاً اسکنر را متصل کنید و دوباره تلاش کنید.")
                self.scanner_combo.addItem("-- هیچ اسکنری یافت نشد --")
                
        except Exception as e:
            self.status_label.setText(f"⚠️ خطا در جستجوی اسکنر: {str(e)}")
            self.scanner_combo.addItem("-- خطا در جستجو --")
    
    def on_scanner_selected(self, index):
        """Handle scanner selection"""
        if index > 0 and index <= len(self.scanner_list):
            self.selected_scanner = self.scanner_list[index - 1]
            self.status_label.setText(f"اسکنر انتخاب شد: {self.selected_scanner.name}")
        else:
            self.selected_scanner = None
        self.save_settings()
    
    def load_saved_settings(self):
        """Restore last-used settings from a previous run.

        Only general preferences are restored - national code and full
        name are per-document and are intentionally never saved.
        """
        saved_path = self.settings.value("save_path", "")
        if saved_path:
            path = Path(saved_path)
            if path.exists():
                self.save_path = path
                self.save_path_edit.setText(str(self.save_path))
        
        saved_doc_type = self.settings.value("doc_type", "")
        if saved_doc_type:
            idx = self.doc_type_combo.findText(saved_doc_type)
            if idx >= 0:
                self.doc_type_combo.setCurrentIndex(idx)
        
        saved_quality = self.settings.value("quality_preset", "")
        if saved_quality:
            idx = self.quality_combo.findData(saved_quality)
            if idx >= 0:
                self.quality_combo.setCurrentIndex(idx)
        
        saved_scanner_name = self.settings.value("scanner_name", "")
        if saved_scanner_name:
            idx = self.scanner_combo.findText(saved_scanner_name)
            if idx > 0:
                self.scanner_combo.setCurrentIndex(idx)
    
    def save_settings(self):
        """Persist current general settings for the next run"""
        self.settings.setValue("save_path", str(self.save_path))
        self.settings.setValue("doc_type", self.doc_type_combo.currentText())
        self.settings.setValue("quality_preset", self.quality_combo.currentData())
        if self.selected_scanner:
            self.settings.setValue("scanner_name", self.selected_scanner.name)
    
    def validate_national_code(self, text):
        """Validate national code (basic validation)"""
        # Allow only digits
        digits_only = ''.join(filter(str.isdigit, text))
        if text != digits_only:
            cursor_pos = self.national_code_edit.cursorPosition()
            self.national_code_edit.setText(digits_only)
            self.national_code_edit.setCursorPosition(cursor_pos - 1)
        
        # Style based on length
        if len(digits_only) == 10:
            self.national_code_edit.setStyleSheet("border: 2px solid green;")
        elif len(digits_only) > 0:
            self.national_code_edit.setStyleSheet("border: 2px solid orange;")
        else:
            self.national_code_edit.setStyleSheet("")
    
    def on_doc_type_changed(self, text):
        """Handle document type change"""
        if text == "سایر":
            self.custom_doc_label.setVisible(True)
            self.custom_doc_edit.setVisible(True)
            self.custom_doc_edit.setFocus()
        else:
            self.custom_doc_label.setVisible(False)
            self.custom_doc_edit.setVisible(False)
            self.custom_doc_edit.clear()
        self.save_settings()
    
    def browse_save_path(self):
        """Browse for save directory"""
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "انتخاب مسیر ذخیره‌سازی",
            str(self.save_path),
            QFileDialog.ShowDirsOnly
        )
        if dir_path:
            self.save_path = Path(dir_path)
            self.save_path_edit.setText(str(self.save_path))
            self.update_status(f"مسیر ذخیره‌سازی تغییر کرد: {self.save_path}", "info")
            self.save_settings()
    
    def start_scan(self):
        """Start scanning process"""
        # Validate national code
        national_code = self.national_code_edit.text().strip()
        if not national_code:
            QMessageBox.warning(self, "خطا", "لطفاً کد ملی را وارد کنید")
            self.national_code_edit.setFocus()
            return
        
        if not national_code.isdigit() or len(national_code) != 10:
            QMessageBox.warning(self, "خطا", "کد ملی باید ۱۰ رقم باشد")
            self.national_code_edit.setFocus()
            return
        
        # Check if scanner is selected
        if not self.selected_scanner and WIA_AVAILABLE:
            reply = QMessageBox.question(
                self,
                "انتخاب اسکنر",
                "هیچ اسکنری انتخاب نشده است. آیا می‌خواهید از حالت شبیه‌سازی استفاده کنید؟",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
        
        # Get other fields
        full_name = self.full_name_edit.text().strip()
        doc_type = self.doc_type_combo.currentText()
        
        # Validate custom document type if "سایر" is selected
        custom_doc_type = None
        if doc_type == "سایر":
            custom_doc_type = self.custom_doc_edit.text().strip()
            if not custom_doc_type:
                QMessageBox.warning(self, "خطا", "لطفاً نوع مدرک دلخواه را وارد کنید")
                self.custom_doc_edit.setFocus()
                return
        
        # Disable scan button and show progress
        self.scan_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # Get chosen scan quality / file size preset
        quality_preset = self.quality_combo.currentData()
        
        # Start scanning thread
        self.scan_worker = ScanWorker(
            self.save_path,
            national_code,
            full_name,
            doc_type,
            custom_doc_type,
            self.selected_scanner,
            quality_preset
        )
        
        # Connect signals
        self.scan_worker.progress_updated.connect(self.update_progress)
        self.scan_worker.status_updated.connect(self.update_status)
        self.scan_worker.scan_finished.connect(self.on_scan_finished)
        self.scan_worker.scan_error.connect(self.on_scan_error)
        
        # Start the thread
        self.scan_worker.start()
    
    def update_progress(self, value):
        """Update progress bar"""
        self.progress_bar.setValue(value)
    
    def update_status(self, message, status_type="info"):
        """Update status label"""
        if status_type == "error":
            self.status_label.setStyleSheet("""
                QLabel {
                    padding: 10px;
                    background-color: #ffcccc;
                    border-radius: 5px;
                    border: 1px solid #ff0000;
                }
            """)
        elif status_type == "success":
            self.status_label.setStyleSheet("""
                QLabel {
                    padding: 10px;
                    background-color: #ccffcc;
                    border-radius: 5px;
                    border: 1px solid #00cc00;
                }
            """)
        else:
            self.status_label.setStyleSheet("""
                QLabel {
                    padding: 10px;
                    background-color: #f0f0f0;
                    border-radius: 5px;
                    border: 1px solid #cccccc;
                }
            """)
        
        self.status_label.setText(message)
    
    def on_scan_finished(self, filepath, message):
        """Handle scan completion"""
        self.progress_bar.setValue(100)
        self.update_status(f"{message} - فایل ذخیره شد: {filepath}", "success")
        self.scan_btn.setEnabled(True)
        
        # Show success message
        QMessageBox.information(
            self,
            "اسکن موفق",
            f"سند با موفقیت اسکن و ذخیره شد!\n\nمسیر: {filepath}"
        )
        
        # Reset progress bar after a delay
        QTimer.singleShot(3000, self.reset_progress)
    
    def on_scan_error(self, error_message):
        """Handle scan error"""
        self.update_status(error_message, "error")
        self.scan_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        QMessageBox.critical(self, "خطا", error_message)
    
    def reset_progress(self):
        """Reset progress bar"""
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)
    
    def open_scan_folder(self):
        """Open the scan folder in file explorer"""
        if self.save_path.exists():
            if platform.system() == "Windows":
                os.startfile(str(self.save_path))
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", str(self.save_path)])
            else:  # Linux
                subprocess.run(["xdg-open", str(self.save_path)])
        else:
            QMessageBox.warning(self, "خطا", "پوشه مورد نظر وجود ندارد")
    
    def show_settings(self):
        """Show settings dialog"""
        dialog = QDialog(self)
        dialog.setWindowTitle("تنظیمات پیشرفته")
        dialog.setLayout(QVBoxLayout())
        dialog.resize(400, 300)
        
        # Add settings options
        settings_group = QGroupBox("تنظیمات اسکن")
        settings_layout = QGridLayout()
        settings_group.setLayout(settings_layout)
        
        # Resolution
        settings_layout.addWidget(QLabel("وضوح تصویر (DPI):"), 0, 0)
        resolution_combo = QComboBox()
        resolution_combo.addItems(["150", "200", "300", "600"])
        resolution_combo.setCurrentText("300")
        settings_layout.addWidget(resolution_combo, 0, 1)
        
        # Color mode
        settings_layout.addWidget(QLabel("حالت رنگ:"), 1, 0)
        color_combo = QComboBox()
        color_combo.addItems(["رنگی", "سیاه و سفید", "خاکستری"])
        settings_layout.addWidget(color_combo, 1, 1)
        
        # File format
        settings_layout.addWidget(QLabel("فرمت فایل:"), 2, 0)
        format_combo = QComboBox()
        format_combo.addItems(["JPEG", "PNG", "PDF"])
        settings_layout.addWidget(format_combo, 2, 1)
        
        dialog.layout().addWidget(settings_group)
        
        # Buttons
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        dialog.layout().addWidget(button_box)
        
        if dialog.exec_() == QDialog.Accepted:
            QMessageBox.information(self, "تنظیمات", "تنظیمات با موفقیت ذخیره شد")
    
    def show_about(self):
        """Show about dialog"""
        QMessageBox.about(
            self,
            "درباره برنامه",
            """
            <h2>اسکنر اسناد</h2>
            <p>نسخه 1.0.0</p>
            <p>برنامه اسکن اسناد با قابلیت مدیریت خودکار فایل‌ها</p>
            <p>توسعه‌دهنده: <b>شیخ زاده</b></p>
            <p>ویژگی‌ها:</p>
            <ul>
                <li>اسکن با کیفیت بالا</li>
                <li>سازماندهی خودکار بر اساس کد ملی</li>
                <li>پشتیبانی از انواع مختلف مدارک</li>
                <li>مدیریت نام‌گذاری خودکار فایل‌ها</li>
                <li>پشتیبانی از اسکنرهای WIA در ویندوز</li>
            </ul>
            """
        )
    
    def closeEvent(self, event):
        """Handle window close event"""
        if self.scan_worker and self.scan_worker.isRunning():
            reply = QMessageBox.question(
                self,
                "در حال اسکن",
                "اسکن در حال انجام است. آیا مطمئنید می‌خواهید خارج شوید؟",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.No:
                event.ignore()
                return
            else:
                self.scan_worker.is_running = False
                self.scan_worker.wait()
        
        self.save_settings()
        event.accept()


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    # Set application font
    font = QFont("Arial", 10)
    app.setFont(font)
    
    # Set application style
    app.setStyleSheet("""
        QMainWindow {
            background-color: #f5f5f5;
        }
        QGroupBox {
            font-weight: bold;
            border: 2px solid #cccccc;
            border-radius: 8px;
            margin-top: 10px;
            padding-top: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 5px 0 5px;
        }
        QLineEdit, QComboBox {
            padding: 5px;
            border: 1px solid #cccccc;
            border-radius: 4px;
            min-height: 25px;
        }
        QLineEdit:focus, QComboBox:focus {
            border: 2px solid #4CAF50;
        }
        QPushButton {
            min-height: 30px;
            padding: 5px 15px;
            border-radius: 4px;
        }
    """)
    
    window = ScannerUI()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()