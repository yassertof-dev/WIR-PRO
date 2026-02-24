import sys
import os
import re
import datetime
import time
import tempfile
import threading
import traceback
import unicodedata
import pythoncom
import json
from docxtpl import DocxTemplate
from win32com.client import Dispatch
from pypdf import PdfWriter

from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QLabel, QLineEdit, QPushButton,
                             QTabWidget, QDateEdit, QSpinBox, QFileDialog,
                             QMessageBox, QScrollArea, QGroupBox, QFormLayout,
                             QTextEdit, QProgressBar, QListWidget, QListWidgetItem,
                             QFrame, QMenu, QSizePolicy, QStyleFactory,
                             QCalendarWidget, QDialog, QAbstractItemView, QCompleter)
from PyQt5.QtCore import (Qt, QDate, QThread, pyqtSignal, QLocale, QPoint, QMimeData, QObject, QStringListModel)
from PyQt5.QtGui import QFont, QTextCharFormat, QColor, QDragEnterEvent, QDropEvent, QWheelEvent, QIcon

# كتم رسائل المكتبات مرة واحدة عند بداية البرنامج
import logging as _logging
_logging.getLogger("pypdf").setLevel(_logging.CRITICAL)
_logging.getLogger("docxtpl").setLevel(_logging.CRITICAL)

# -------------------- QDateEdit مخصص بسهمين ظاهرين --------------------
class ArrowDateEdit(QDateEdit):
    """QDateEdit بدون Fusion style - يستخدم Windows native style"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setButtonSymbols(QDateEdit.UpDownArrows)
        # لا نستخدم Fusion style لأنه يؤثر على الـ widgets المجاورة

# -------------------- دالة المسار (لـ PyInstaller) --------------------
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# -------------------- دوال مساعدة --------------------
# جدول التحويل يُبنى مرة واحدة عند التحميل
_ARABIC_TO_ENGLISH = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")

def to_english_digits(text) -> str:
    """تحويل الأرقام العربية إلى إنجليزية مع تطبيع النص"""
    if not text:
        return ""
    return unicodedata.normalize('NFKC', str(text)).translate(_ARABIC_TO_ENGLISH)

def clean_digits_only(text) -> str:
    """تحويل وإبقاء الأرقام فقط"""
    return ''.join(c for c in to_english_digits(text) if c.isdigit())

def apply_rtl_lock(text):
    if not text:
        return ""
    return f"\u202b{text}\u202c"

def convert_docx_to_pdf(word_app, docx_path, pdf_path):
    docx_path = os.path.abspath(docx_path)
    pdf_path = os.path.abspath(pdf_path)
    doc = word_app.Documents.Open(docx_path)
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    return True

# -------------------- الإعدادات العامة --------------------
PROJECT_PREFIX = "TOL-ADW-WIR"
DISC_LIST = [('AR', 'معماري'), ('CV', 'مدني'), ('MECH', 'ميكانيكا'), ('ELEC', 'كهرباء')]
ARABIC_DAYS = {
    "Monday": "الاثنين", "Tuesday": "الثلاثاء", "Wednesday": "الأربعاء",
    "Thursday": "الخميس", "Friday": "الجمعة", "Saturday": "السبت", "Sunday": "الأحد"
}
def log_error(error_message):
    try:
        log_dir = "Logs"
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, f"error_{datetime.datetime.now().strftime('%Y%m%d')}.log")
        with open(log_file, 'a', encoding='utf-8') as f:
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {error_message}\n")
            f.write("-" * 50 + "\n")
    except:
        pass

# -------------------- Smart Suggestions System --------------------
class SuggestionsDB:
    """قاعدة بيانات الاقتراحات الذكية مع حفظ مؤجل لتحسين الأداء"""

    VALID_CODES = {'AR', 'CV', 'MECH', 'ELEC'}

    def __init__(self, db_path="suggestions_db.json"):
        self.db_path = db_path
        self._dirty = False          # علامة لتجنب الحفظ الزائد
        self.data = self._load()

    def _load(self) -> dict:
        if os.path.exists(self.db_path):
            try:
                with open(self.db_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        return self._init_structure()

    def _init_structure(self) -> dict:
        return {code: {} for code in self.VALID_CODES}

    def save(self):
        """حفظ فوري لقاعدة البيانات"""
        try:
            with open(self.db_path, 'w', encoding='utf-8') as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
            self._dirty = False
        except Exception as e:
            log_error(f"Failed to save suggestions DB: {e}")

    def save_if_dirty(self):
        """حفظ فقط إذا كان هناك تغييرات"""
        if self._dirty:
            self.save()

    def _upsert(self, bucket: dict, key: str, suffix: str, attachments: list):
        """إضافة أو تحديث اقتراح في bucket معين"""
        if key not in bucket:
            bucket[key] = {'count': 0, 'suffix': suffix, 'attachments': []}
        entry = bucket[key]
        entry['count'] += 1
        entry['suffix'] = suffix
        # دمج المرفقات بدون تكرار بكفاءة باستخدام set
        existing = set(entry['attachments'])
        for att in (attachments or []):
            if att not in existing:
                entry['attachments'].append(att)
                existing.add(att)

    def add_suggestion(self, discipline_code: str, plot_number, description: str,
                       suffix: str = "", attachments: list = None):
        """إضافة اقتراح - الحفظ مؤجل حتى save_if_dirty()"""
        desc = (description or "").strip()
        if not desc or discipline_code not in self.VALID_CODES:
            return

        disc = self.data.setdefault(discipline_code, {})
        plot_key = str(plot_number)

        # تحديث bucket القطعة المحددة
        self._upsert(disc.setdefault(plot_key, {}), desc, suffix, attachments)

        # تحديث bucket العام
        self._upsert(disc.setdefault('all', {}), desc, suffix, attachments)

        self._dirty = True
        # لا نحفظ هنا - الحفظ يتم مرة واحدة بعد انتهاء كل العمليات

    def _parse_entry(self, desc: str, data, source: str) -> dict:
        """تحويل entry (قديم أو جديد) لصيغة موحدة"""
        if isinstance(data, dict):
            return {
                'text': desc,
                'count': data.get('count', 0),
                'suffix': data.get('suffix', ''),
                'attachments': data.get('attachments', []),
                'source': source,
            }
        return {'text': desc, 'count': data, 'suffix': '', 'attachments': [], 'source': source}

    def get_suggestions(self, discipline_code: str, plot_number=None, limit: int = 5) -> list:
        """جلب الاقتراحات مرتبة حسب التكرار"""
        disc = self.data.get(discipline_code)
        if not disc:
            return []

        seen: set = set()
        suggestions = []

        # أولاً: اقتراحات القطعة المحددة
        if plot_number:
            for desc, data in disc.get(str(plot_number), {}).items():
                seen.add(desc)
                suggestions.append(self._parse_entry(desc, data, f'قطعة {plot_number}'))

        # ثانياً: الاقتراحات العامة (بدون تكرار)
        for desc, data in disc.get('all', {}).items():
            if desc not in seen:
                suggestions.append(self._parse_entry(desc, data, 'عام'))

        suggestions.sort(key=lambda x: x['count'], reverse=True)
        return suggestions[:limit]

    def remove_suggestion(self, discipline_code: str, description: str):
        """حذف اقتراح من جميع القطع"""
        disc = self.data.get(discipline_code)
        if not disc:
            return
        for bucket in disc.values():
            bucket.pop(description, None)
        self.save()

# Instance عامة
suggestions_db = SuggestionsDB()

# -------------------- Worker Thread --------------------
class WorkerThread(QThread):
    task_done  = pyqtSignal(int, str, str)
    task_error = pyqtSignal(int, str)

    def __init__(self, tasks, data, stop_event):
        super().__init__()
        self.tasks      = tasks
        self.data       = data
        self.stop_event = stop_event  # threading.Event بدل list

    def _build_final_path(self, output_dir, ref, rev, plot, suffix):
        base = f"{ref}-REV{rev:02d}" if rev > 0 else ref
        name = f"{base}-{plot}-{suffix}.pdf" if suffix else f"{base}-{plot}.pdf"
        return os.path.join(output_dir, name)

    def _merge_pdfs(self, temp_pdf, attach_paths, final_pdf):
        """دمج PDFs بطريقة آمنة مع كتم رسائل pypdf"""
        merger = PdfWriter()
        try:
            merger.append(temp_pdf)
            for ap in attach_paths:
                if ap and os.path.exists(ap):
                    merger.append(ap)
            with open(final_pdf, 'wb') as f:
                merger.write(f)
        finally:
            merger.close()

    def _cleanup(self, *paths):
        for p in paths:
            try:
                if p and os.path.exists(p):
                    os.remove(p)
            except OSError:
                pass

    def run(self):
        pythoncom.CoInitialize()
        word_app  = None
        temp_docx = None
        temp_pdf  = None
        try:
            word_app = Dispatch("Word.Application")
            word_app.Visible       = False
            word_app.DisplayAlerts = False

            template_path = resource_path("template.docx")
            temp_dir = tempfile.gettempdir()
            thread_id = threading.get_ident()

            for task_index, task in self.tasks:
                if self.stop_event.is_set():
                    return

                temp_docx = temp_pdf = None
                try:
                    ref          = task['ref']
                    name         = task['name']
                    desc         = task['desc']
                    attach_paths = task['attach_paths']
                    rev          = task['rev']
                    plot         = task['plot']
                    suffix       = task.get('suffix', '')

                    secured_desc = apply_rtl_lock(f"{desc} قطعة رقم {plot}")
                    output_dir   = os.path.join("Output", name)
                    os.makedirs(output_dir, exist_ok=True)

                    uid       = f"{thread_id}_{int(time.time()*1e6)}"
                    temp_docx = os.path.join(temp_dir, f"tmp_{uid}.docx")
                    temp_pdf  = os.path.join(temp_dir, f"tmp_{uid}.pdf")

                    doc = DocxTemplate(template_path)
                    doc.render({
                        'REF': ref, 'DATE': self.data['date'],
                        'TIME': self.data['time'], 'DESC': secured_desc,
                        'PLOT': plot, 'REV': f"{rev:02d}"
                    })
                    doc.save(temp_docx)

                    if self.stop_event.is_set():
                        self._cleanup(temp_docx)
                        return

                    convert_docx_to_pdf(word_app, temp_docx, temp_pdf)

                    if self.stop_event.is_set():
                        self._cleanup(temp_docx, temp_pdf)
                        return

                    final_pdf = self._build_final_path(output_dir, ref, rev, plot, suffix)
                    self._merge_pdfs(temp_pdf, attach_paths, final_pdf)
                    self._cleanup(temp_docx, temp_pdf)

                    self.task_done.emit(task_index, f"تم: {os.path.basename(final_pdf)}", final_pdf)

                except Exception as e:
                    self._cleanup(temp_docx, temp_pdf)
                    error_msg = f"خطأ في {task.get('ref','')}-{task.get('plot','')}: {e}"
                    log_error(error_msg + "\n" + traceback.format_exc())
                    self.task_error.emit(task_index, error_msg)
                    return

        except Exception as e:
            error_msg = f"خطأ في تشغيل Word: {e}"
            log_error(error_msg + "\n" + traceback.format_exc())
            self.task_error.emit(-1, error_msg)
        finally:
            if word_app:
                try: word_app.Quit()
                except: pass
            pythoncom.CoUninitialize()

class ProcessThread(QThread):
    progress_update = pyqtSignal(int, int, str, str)
    finished = pyqtSignal(bool, str)

    NUM_WORKERS = 4

    def __init__(self, data):
        super().__init__()
        self.data          = data
        self._stop_event   = threading.Event()   # أسرع وأأمن من list
        self.created_files = []
        self._lock         = threading.Lock()
        self._processed    = 0
        self._total        = 0
        self._error        = None
        self.start_time    = None

    def stop(self):
        self._stop_event.set()

    def run(self):
        pythoncom.CoInitialize()
        self.start_time = time.time()
        try:
            global_plots = self.data['plots']
            tasks = []

            for tab_data in self.data['tabs']:
                code = tab_data['code']
                name = tab_data['name']
                serial = tab_data['serial']

                for row in tab_data['rows']:
                    manual_mode = row.get('manual_mode', False)
                    manual_ref = row.get('manual_ref', None)
                    desc = row['desc']
                    attach_paths = row['attach_paths']
                    rev = row['revision']
                    suffix_raw = row.get('suffix', '')

                    # تنظيف suffix من الأرقام العربية
                    suffix_clean = to_english_digits(suffix_raw)

                    if rev > 0:
                        plots_list = [row['manual_plot']]
                    else:
                        plots_list = global_plots if row['plots'] is None else row['plots']

                    for p_num in plots_list:
                        if not p_num:
                            continue
                        p_num_clean = clean_digits_only(str(p_num))
                        ref_clean   = to_english_digits(
                            manual_ref if (manual_mode and manual_ref)
                            else f"{PROJECT_PREFIX}-{code}-{serial:03d}"
                        )
                        tasks.append({
                            'ref': ref_clean, 'code': code, 'name': name,
                            'desc': desc, 'attach_paths': attach_paths,
                            'rev': rev, 'plot': p_num_clean, 'suffix': suffix_clean
                        })
                        if not manual_mode:
                            serial += 1

            self._total = len(tasks)
            if self._total == 0:
                self.finished.emit(False, "لا توجد ملفات لتوليدها")
                return

            num_workers = min(self.NUM_WORKERS, self._total)
            chunks = [[] for _ in range(num_workers)]
            for i, task in enumerate(tasks):
                chunks[i % num_workers].append((i, task))

            workers = []
            for chunk in chunks:
                if chunk:
                    w = WorkerThread(chunk, self.data, self._stop_event)
                    w.task_done.connect(self._on_task_done)
                    w.task_error.connect(self._on_task_error)
                    workers.append(w)

            for w in workers:
                w.start()

            for w in workers:
                w.wait()

            if self._stop_event.is_set():
                for f in self.created_files:
                    try:
                        if os.path.exists(f):
                            os.remove(f)
                    except OSError:
                        pass
                self.finished.emit(False, "تم إيقاف التوليد وحذف الملفات المنشأة")
            elif self._error:
                self.finished.emit(False, self._error)
            else:
                self.progress_update.emit(self._total, self._total, "اكتمل توليد جميع الملفات ✓", "")

                # حفظ الاقتراحات الذكية
                self._save_suggestions()
                self.finished.emit(True, f"تم توليد {self._total} ملف بنجاح!")

        except Exception as e:
            error_msg = f"خطأ عام: {str(e)}"
            log_error(error_msg + "\n" + traceback.format_exc())
            self.finished.emit(False, error_msg)
        finally:
            pythoncom.CoUninitialize()

    def _on_task_done(self, task_index, message, final_file):
        with self._lock:
            self._processed += 1
            self.created_files.append(final_file)
            current = self._processed
            total = self._total

            elapsed = time.time() - self.start_time
            if current > 0:
                self.avg_time_per_file = elapsed / current
                remaining_seconds = self.avg_time_per_file * (total - current)
                if remaining_seconds < 60:
                    time_str = f"{int(remaining_seconds)} ثانية"
                elif remaining_seconds < 3600:
                    minutes = int(remaining_seconds // 60)
                    seconds = int(remaining_seconds % 60)
                    time_str = f"{minutes} دقيقة {seconds} ثانية"
                else:
                    hours = int(remaining_seconds // 3600)
                    minutes = int((remaining_seconds % 3600) // 60)
                    time_str = f"{hours} ساعة {minutes} دقيقة"
            else:
                time_str = "جاري الحساب..."

        self.progress_update.emit(current, total, message, time_str)

    def _on_task_error(self, task_index, error):
        with self._lock:
            self._error = error
        self._stop_event.set()

    def _save_suggestions(self):
        """حفظ الاقتراحات الذكية مرة واحدة بعد انتهاء التوليد"""
        try:
            for tab_data in self.data['tabs']:
                code = tab_data['code']
                for row in tab_data['rows']:
                    desc = (row.get('desc') or "").strip()
                    if not desc:
                        continue
                    suffix      = row.get('suffix', '')
                    attachments = row.get('attach_paths', [])
                    plots_list  = (
                        [row.get('manual_plot')] if row.get('manual_mode')
                        else (self.data['plots'] if row.get('plots') is None else row.get('plots', []))
                    )
                    for plot_num in plots_list:
                        if plot_num:
                            suggestions_db.add_suggestion(code, plot_num, desc, suffix, attachments)
            # حفظ مرة واحدة في النهاية بدل حفظ في كل add_suggestion
            suggestions_db.save_if_dirty()
        except Exception as e:
            log_error(f"Failed to save suggestions: {e}")

# -------------------- قائمة الاقتراحات المخصصة لمنع تمرير الخلفية --------------------
class SuggestionListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFocusPolicy(Qt.StrongFocus)

    def wheelEvent(self, event: QWheelEvent):
        """
        منع الحدث من الانتشار إلى القائمة الرئيسية للطلبات.
        إذا كان شريط التمرير مرئيًا، نسمح بالتمرير داخل القائمة.
        وفي جميع الأحوال، نقبل الحدث لمنع انتقاله للأب.
        """
        if self.verticalScrollBar().isVisible():
            super().wheelEvent(event)
        # منع انتشار الحدث إلى الأب (QScrollArea الرئيسي)
        event.accept()

# -------------------- Drop Zone للمرفقات --------------------
class DropZoneWidget(QWidget):
    """QWidget يستقبل ملفات PDF مسحوبة من خارج البرنامج"""
    filesDropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self._normal_style = ""
        self._hover_style  = ""

    def setStyles(self, normal, hover):
        self._normal_style = normal
        self._hover_style  = hover

    def _has_pdf_urls(self, mime_data):
        """التحقق إن الـ mime data فيها ملف PDF واحد على الأقل"""
        if not mime_data.hasUrls():
            return False
        return any(u.toLocalFile().lower().endswith('.pdf') for u in mime_data.urls())

    def dragEnterEvent(self, event):
        if self._has_pdf_urls(event.mimeData()):
            if self._hover_style:
                self.setStyleSheet(self._hover_style)
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if self._has_pdf_urls(event.mimeData()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        if self._normal_style:
            self.setStyleSheet(self._normal_style)
        event.accept()

    def dropEvent(self, event):
        if self._normal_style:
            self.setStyleSheet(self._normal_style)
        if self._has_pdf_urls(event.mimeData()):
            paths = [u.toLocalFile() for u in event.mimeData().urls()
                     if u.toLocalFile().lower().endswith('.pdf')]
            if paths:
                self.filesDropped.emit(paths)
            event.acceptProposedAction()
        else:
            event.ignore()

# -------------------- فلتر لتمرير Drag events عبر QScrollArea --------------------
class ScrollAreaDropFilter(QObject):
    """يعترض drag/drop events على الـ scroll viewport ويمررها للـ child الصح"""
    def __init__(self, scroll_area, target_widget):
        super().__init__(scroll_area)
        self._target = target_widget
        scroll_area.setAcceptDrops(True)
        scroll_area.viewport().setAcceptDrops(True)
        scroll_area.viewport().installEventFilter(self)

    def eventFilter(self, obj, event):
        from PyQt5.QtCore import QEvent
        if event.type() not in (QEvent.DragEnter, QEvent.DragMove,
                                 QEvent.DragLeave, QEvent.Drop):
            return False

        mime = event.mimeData()

        # تجاهل أي drag ليس فيه PDF على الأقل
        if event.type() in (QEvent.DragEnter, QEvent.DragMove, QEvent.Drop):
            has_pdf = (mime.hasUrls() and
                       any(u.toLocalFile().lower().endswith('.pdf')
                           for u in mime.urls()))
            if not has_pdf:
                event.ignore()
                return True

        target_pos = self._target.mapFromGlobal(obj.mapToGlobal(event.pos()))

        if event.type() == QEvent.DragLeave:
            QApplication.sendEvent(self._target, event)
            return False

        if not self._target.rect().contains(target_pos):
            event.ignore()
            return True

        from PyQt5.QtGui import QDragEnterEvent, QDragMoveEvent, QDropEvent
        if event.type() == QEvent.DragEnter:
            new_event = QDragEnterEvent(
                target_pos, event.possibleActions(),
                mime, event.mouseButtons(), event.keyboardModifiers()
            )
            QApplication.sendEvent(self._target, new_event)
            if new_event.isAccepted():
                event.acceptProposedAction()
            else:
                event.ignore()
            return True

        elif event.type() == QEvent.DragMove:
            new_event = QDragMoveEvent(
                target_pos, event.possibleActions(),
                mime, event.mouseButtons(), event.keyboardModifiers()
            )
            QApplication.sendEvent(self._target, new_event)
            if new_event.isAccepted():
                event.acceptProposedAction()
            else:
                event.ignore()
            return True

        elif event.type() == QEvent.Drop:
            new_event = QDropEvent(
                target_pos, event.possibleActions(),
                mime, event.mouseButtons(), event.keyboardModifiers()
            )
            QApplication.sendEvent(self._target, new_event)
            if new_event.isAccepted():
                event.setDropAction(new_event.dropAction())
                event.accept()
            else:
                event.ignore()
            return True

        return False

# -------------------- AttachListWidget (list فقط بدون drop) --------------------
class AttachListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)

        # Drag & Drop ترتيب داخلي
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QAbstractItemView.SingleSelection)

        # تعطيل التركيز (يزيل الخطوط الغريبة)
        self.setFocusPolicy(Qt.NoFocus)

        # تحسين السحب
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragEnabled(True)

        # ستايل نظيف
        self.setStyleSheet("""
        QListWidget {
            background-color: #f8fafc;
            border: 2px dashed #cbd5e1;
            border-radius: 10px;
        }

        QListWidget::item {
            background-color: white;
            border: none;
            border-radius: 8px;
            padding: 6px 12px;
            margin: 6px;
            color: #1e293b;
        }

        QListWidget::item:hover {
            background-color: #eff6ff;
        }

        QListWidget::item:selected {
            background-color: #dbeafe;
            color: #1e40af;
        }
        """)

# -------------------- كلاس التبويب (DisciplineTab) --------------------
class DisciplineTab(QWidget):
    SUFFIX_HISTORY_FILE = "suffix_history.json"
    SUFFIX_HISTORY_MAX = 30

    def __init__(self, code, name, main_window=None):
        super().__init__()
        self.code, self.name, self.rows = code, name, []
        self.main_window = main_window
        self.suffix_history = []
        self._load_suffix_history()
        self.suffix_completer_model = QStringListModel()
        self.suffix_completer_model.setStringList(self.suffix_history)
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(8, 8, 8, 8)
        self.main_layout.setSpacing(8)

        # شريط الأدوات العلوي
        top_bar = QWidget()
        top_bar.setStyleSheet("""
            QWidget {
                background-color: #f8f9fa;
                border: 1px solid #e9ecef;
                border-radius: 8px;
            }
        """)
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(10, 6, 10, 6)

        self.btn_add = QPushButton(f"＋  إضافة طلب {name}")
        self.btn_add.clicked.connect(self.add_row)
        self.btn_add.setCursor(Qt.PointingHandCursor)
        self.btn_add.setStyleSheet("""
            QPushButton {
                background-color: #2e7d32;
                color: white;
                font-weight: bold;
                font-size: 14px;
                padding: 6px 14px;
                border-radius: 6px;
                border: none;
            }
            QPushButton:hover { background-color: #1b5e20; }
            QPushButton:pressed { background-color: #388e3c; }
        """)

        # زر مسح كل الطلبات في هذا التبويب
        self.btn_clear_tab = QPushButton("🗑️  مسح كل الطلبات")
        self.btn_clear_tab.setCursor(Qt.PointingHandCursor)
        self.btn_clear_tab.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                font-weight: bold;
                font-size: 14px;
                padding: 6px 14px;
                border-radius: 6px;
                border: none;
                margin-right: 8px;
            }
            QPushButton:hover { background-color: #c82333; }
            QPushButton:pressed { background-color: #bd2130; }
        """)
        self.btn_clear_tab.clicked.connect(self.clear_tab)

        serial_label = QLabel("بداية الترقيم:")
        serial_label.setStyleSheet("color: #555; font-size: 14px; background: transparent; border: none;")

        self.serial_input = QLineEdit("1")
        self.serial_input.setFixedWidth(55)
        self.serial_input.setFixedHeight(18)
        self.serial_input.setAlignment(Qt.AlignCenter)
        self.serial_input.setStyleSheet("""
            QLineEdit {
                border: 1px solid #ced4da;
                border-radius: 5px;
                padding: 2px 5px;
                font-size: 14px;
                background: white;
            }
            QLineEdit:focus { border-color: #80bdff; }
        """)

        top_layout.addWidget(self.btn_add)
        top_layout.addWidget(self.btn_clear_tab)
        top_layout.addStretch()
        top_layout.addWidget(serial_label)
        top_layout.addWidget(self.serial_input)
        self.main_layout.addWidget(top_bar)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        self.scroll_content = QWidget()
        self.scroll_content.setStyleSheet("background: transparent;")
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setAlignment(Qt.AlignTop)
        self.scroll_layout.setSpacing(8)
        self.scroll_layout.setContentsMargins(2, 2, 2, 2)
        self.scroll.setWidget(self.scroll_content)
        self.main_layout.addWidget(self.scroll)
        self.add_row()

    def add_row(self):
        # الحاوية الرئيسية للطلب - كارد أبيض بظل خفيف
        row_container = QWidget()
        row_container.setObjectName("requestCard")
        row_container.setStyleSheet("""
            QWidget#requestCard {
                background-color: white;
                border: 1px solid #dee2e6;
                border-radius: 10px;
            }
        """)
        main_layout = QVBoxLayout(row_container)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # ===== شريط العنوان =====
        header_widget = QWidget()
        header_widget.setObjectName("cardHeader")
        header_widget.setStyleSheet("""
            QWidget#cardHeader {
                background-color: #f1f3f5;
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
                border-bottom: 1px solid #dee2e6;
            }
        """)
        header_widget.setCursor(Qt.PointingHandCursor)
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(12, 8, 12, 8)
        header_layout.setSpacing(8)

        expand_btn = QPushButton("▼")
        expand_btn.setFixedSize(22, 22)
        expand_btn.setStyleSheet("""
            QPushButton {
                background-color: #e9ecef;
                border: 1px solid #ced4da;
                border-radius: 11px;
                font-size: 14px;
                color: #495057;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #dee2e6; }
        """)
        expand_btn.setCheckable(True)
        expand_btn.setChecked(True)

        summary_label = QPushButton("طلب جديد")
        summary_label.setStyleSheet("""
            QPushButton {
                border: none;
                background: transparent;
                font-weight: bold;
                font-size: 14px;
                color: #1e293b;
                text-align: right;
            }
            QPushButton:hover { color: #0d6efd; }
        """)
        summary_label.setCursor(Qt.PointingHandCursor)

        header_layout.addWidget(expand_btn)
        header_layout.addWidget(summary_label, 1)

        summary_label.clicked.connect(expand_btn.toggle)

        # ===== محتوى الطلب =====
        details_widget = QWidget()
        details_widget.setStyleSheet("background: white; border-radius: 0 0 10px 10px;")
        details_layout = QFormLayout(details_widget)
        details_layout.setContentsMargins(16, 12, 16, 12)
        details_layout.setSpacing(10)
        details_layout.setLabelAlignment(Qt.AlignRight)

        # ستايل موحد للحقول - ارتفاع مصغّر (18 بكسل)
        FIELD_STYLE = """
            QLineEdit {
                border: 1.5px solid #94a3b8;
                border-radius: 6px;
                padding: 2px 6px;
                font-size: 14px;
                background: white;
                color: #0f172a;
                min-height: 18px;
            }
            QLineEdit:focus {
                border: 2px solid #3b82f6;
                background: white;
            }
            QLineEdit:disabled {
                background: #e9ecef;
                color: #9ca3af;
                border: 1.5px dashed #cbd5e1;
            }
        """
        SPIN_STYLE = """
            QSpinBox {
                border: 1.5px solid #94a3b8;
                border-radius: 6px;
                padding: 1px 4px;
                font-size: 14px;
                background: white;
                color: #0f172a;
                min-height: 18px;
            }
            QSpinBox:focus { border: 2px solid #3b82f6; }
            QSpinBox:disabled {
                background: #e9ecef;
                color: #9ca3af;
                border: 1.5px dashed #cbd5e1;
            }
        """
        LABEL_STYLE = "font-size: 14px; color: #334155; font-weight: 600;"

        # وصف الأعمال — مربع أعلى و padding أوضح لئلا يظهر النص متآكل
        FIELD_STYLE_TALL = """
            QLineEdit {
                border: 1.5px solid #94a3b8;
                border-radius: 6px;
                padding: 6px 10px;
                font-size: 14px;
                background: white;
                color: #0f172a;
                min-height: 22px;
            }
            QLineEdit:focus { border: 2px solid #3b82f6; background: white; }
            QLineEdit:disabled { background: #e9ecef; color: #9ca3af; border: 1.5px dashed #cbd5e1; }
        """
        desc_edit = QLineEdit()
        desc_edit.setPlaceholderText("اكتب وصف الأعمال هنا...")
        desc_edit.setStyleSheet(FIELD_STYLE_TALL)
        desc_edit.setFixedHeight(36)
        desc_edit.setMinimumWidth(220)
        desc_edit.textChanged.connect(self.update_counter)

        # واجهة الاقتراحات الذكية
        suggestions_container = QWidget()
        suggestions_container.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        suggestions_layout = QVBoxLayout(suggestions_container)
        suggestions_layout.setContentsMargins(0, 0, 0, 0)
        suggestions_layout.setSpacing(2)

        suggestions_label = QLabel("💡 اقتراحات ذكية:")
        suggestions_label.setStyleSheet("font-size: 14px; color: #6c757d; font-weight: 600;")
        suggestions_layout.addWidget(suggestions_label)

        suggestions_list = SuggestionListWidget()
        suggestions_list.setMaximumHeight(80)  # تصغير الارتفاع الأقصى
        suggestions_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Maximum)
        suggestions_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #dee2e6;
                border-radius: 6px;
                background-color: #fff;
                font-size: 14px;
                color: #1e293b;
            }
            QListWidget::item {
                padding: 4px 8px;
                border-bottom: 1px solid #f1f3f5;
                color: #1e293b;
            }
            QListWidget::item:hover { background-color: #e7f5ff; }
            QListWidget::item:selected { background-color: #339af0; color: white; }
        """)
        suggestions_list.setVisible(False)
        suggestions_layout.addWidget(suggestions_list)

        # تخزين رقم الصف الحالي للاستخدام في الدوال الداخلية
        current_row_index = len(self.rows)

        # دالة تحديث الاقتراحات لهذا الصف
        def update_suggestions_for_row():
            suggestions_list.clear()
            # الحصول على أرقام القطع من الخانة العامة
            raw_plots = self.main_window.plots_input.text() if self.main_window else ""
            plots = [p.strip() for p in re.split(r'[,\-\s.*]+', raw_plots) if p.strip().isdigit()]

            # جلب الاقتراحات
            plot_num = plots[0] if plots else None
            suggestions = suggestions_db.get_suggestions(self.code, plot_num, limit=5)

            if suggestions:
                suggestions_list.setVisible(True)
                for sug in suggestions:
                    # إنشاء عنصر مخصص بجانبه زر حذف
                    item = QListWidgetItem(suggestions_list)

                    # إنشاء widget للعنصر
                    item_widget = QWidget()
                    item_layout = QHBoxLayout(item_widget)
                    item_layout.setContentsMargins(2, 2, 2, 2)
                    item_layout.setSpacing(5)

                    # نص الاقتراح الأساسي (بدون عداد التكرار)
                    item_text = sug['text']
                    if sug.get('suffix'):
                        item_text += f" ({sug['suffix']})"

                    # إضافة نص المرفقات إذا وجدت
                    if sug.get('attachments') and len(sug['attachments']) > 0:
                        item_text += f" (مرفقات: {len(sug['attachments'])} ملف)"

                    label = QLabel(item_text)
                    label.setStyleSheet("font-size: 14px;")

                    # إضافة النص
                    item_layout.addWidget(label)

                    # إضافة مسافة مرنة لدفع زر الحذف إلى أقصى اليمين
                    item_layout.addStretch()

                    # زر الحذف (✕ باللون الأحمر) - بحجم أصغر
                    delete_btn = QPushButton("✕")
                    delete_btn.setFixedSize(18, 18)
                    delete_btn.setStyleSheet("""
                        QPushButton {
                            background-color: transparent;
                            border: none;
                            color: #dc3545;
                            font-size: 12px;
                            font-weight: bold;
                            border-radius: 2px;
                        }
                        QPushButton:hover {
                            background-color: #ffebee;
                        }
                    """)
                    delete_btn.setToolTip("حذف هذا الاقتراح")

                    # ربط زر الحذف بدالة الحذف مع تمرير النص الأصلي
                    delete_btn.clicked.connect(
                        lambda checked, text=sug['text']: self.delete_suggestion(text, current_row_index)
                    )

                    item_layout.addWidget(delete_btn)

                    # تعيين الـ widget للعنصر
                    suggestions_list.setItemWidget(item, item_widget)

                    # تخزين بيانات الاقتراح في العنصر لاستخدامها عند النقر (اختيار الاقتراح)
                    item.setData(Qt.UserRole, sug['text'])
                    item.setData(Qt.UserRole + 1, sug.get('suffix', ''))
                    item.setData(Qt.UserRole + 2, sug.get('attachments', []))
            else:
                suggestions_list.setVisible(False)

        # عند التركيز على حقل الوصف
        def on_desc_focus(event):
            update_suggestions_for_row()
            QLineEdit.focusInEvent(desc_edit, event)

        desc_edit.focusInEvent = on_desc_focus

        # عند اختيار اقتراح (النقر على العنصر نفسه)
        def on_suggestion_selected(item):
            original_text = item.data(Qt.UserRole)
            suffix_text = item.data(Qt.UserRole + 1)
            attachments = item.data(Qt.UserRole + 2)

            # ملء الوصف
            desc_edit.setText(original_text)

            # ملء suffix
            if suffix_text:
                suffix_edit.setText(suffix_text)

            # إضافة المرفقات عبر الدالة الموحدة
            if attachments:
                for att_path in attachments:
                    if os.path.exists(att_path):
                        add_attachment_file(att_path)

            suggestions_list.setVisible(False)

        suggestions_list.itemClicked.connect(on_suggestion_selected)

        # Revision
        revision_layout = QHBoxLayout()
        revision_spin = QSpinBox()
        revision_spin.setRange(0, 99)
        revision_spin.setValue(0)
        revision_spin.setPrefix("Rev ")
        revision_spin.setFixedWidth(85)
        revision_spin.setFixedHeight(18)
        revision_spin.setAlignment(Qt.AlignCenter)
        revision_spin.setStyleSheet(SPIN_STYLE)
        revision_spin.valueChanged.connect(self.update_counter)
        revision_layout.addWidget(revision_spin)
        revision_layout.addStretch()

        # حقول الإدخال اليدوي
        manual_ref_layout = QHBoxLayout()
        manual_ref_label = QLabel("رقم الطلب (آخر 3 أرقام):")
        manual_ref_input = QLineEdit()
        manual_ref_input.setPlaceholderText("مثال: 015")
        manual_ref_input.setMaxLength(3)
        manual_ref_input.setFixedWidth(100)
        manual_ref_input.setFixedHeight(18)
        manual_ref_input.setEnabled(False)
        manual_ref_input.setStyleSheet(FIELD_STYLE)
        manual_ref_input.textChanged.connect(self.update_counter)

        def validate_manual_ref_input(text):
            cursor_pos = manual_ref_input.cursorPosition()
            filtered_text = ''.join(c for c in text if c.isdigit())
            if filtered_text != text:
                manual_ref_input.setText(filtered_text)
                manual_ref_input.setCursorPosition(min(cursor_pos, len(filtered_text)))

        manual_ref_input.textChanged.connect(validate_manual_ref_input)
        manual_ref_layout.addWidget(manual_ref_label)
        manual_ref_layout.addWidget(manual_ref_input)
        manual_ref_layout.addStretch()

        manual_plot_layout = QHBoxLayout()
        manual_plot_label = QLabel("رقم القطعة يدوياً:")
        manual_plot_input = QLineEdit()
        manual_plot_input.setPlaceholderText("مثال: 102")
        manual_plot_input.setEnabled(False)
        manual_plot_input.setFixedHeight(18)
        manual_plot_input.setStyleSheet(FIELD_STYLE)
        manual_plot_input.textChanged.connect(self.update_counter)
        manual_plot_input.textChanged.connect(self.update_current_row_summary)

        def validate_manual_plot_input(text):
            cursor_pos = manual_plot_input.cursorPosition()
            filtered_text = ''.join(c for c in text if c.isdigit())
            if filtered_text != text:
                manual_plot_input.setText(filtered_text)
                manual_plot_input.setCursorPosition(min(cursor_pos, len(filtered_text)))

        manual_plot_input.textChanged.connect(validate_manual_plot_input)
        manual_plot_layout.addWidget(manual_plot_label)
        manual_plot_layout.addWidget(manual_plot_input)

        # اختيار نطاق القطع
        plot_scope_layout = QHBoxLayout()
        radio_all = QPushButton("كل القطع")
        radio_all.setCheckable(True)
        radio_all.setChecked(True)
        radio_all.setCursor(Qt.PointingHandCursor)
        radio_all.setStyleSheet("""
            QPushButton {
                padding: 5px 12px;
                border: 1px solid #ced4da;
                border-radius: 6px;
                font-size: 14px;
                color: #495057;
                background: white;
            }
            QPushButton:checked {
                background-color: #0d6efd;
                color: white;
                border-color: #0d6efd;
                font-weight: bold;
            }
            QPushButton:hover:!checked { background-color: #f8f9fa; }
            QPushButton:disabled { background-color: #f8f9fa; color: #6c757d; }
        """)

        radio_specific = QPushButton("قطع محددة")
        radio_specific.setCheckable(True)
        radio_specific.setCursor(Qt.PointingHandCursor)
        radio_specific.setStyleSheet("""
            QPushButton {
                padding: 5px 12px;
                border: 1px solid #ced4da;
                border-radius: 6px;
                font-size: 14px;
                color: #495057;
                background: white;
            }
            QPushButton:checked {
                background-color: #0d6efd;
                color: white;
                border-color: #0d6efd;
                font-weight: bold;
            }
            QPushButton:hover:!checked { background-color: #f8f9fa; }
            QPushButton:disabled { background-color: #f8f9fa; color: #6c757d; }
        """)

        plots_input = QLineEdit()
        plots_input.setPlaceholderText("مثال: 101 105 110 أو 101-105-110")
        plots_input.setEnabled(False)
        plots_input.setStyleSheet(FIELD_STYLE)
        plots_input.setFixedHeight(26)   # مربع أكبر ليناسب النص
        plots_input.textChanged.connect(self.update_counter)
        plots_input.textChanged.connect(self.update_current_row_summary)

        def validate_plots_input(text):
            cursor_pos = plots_input.cursorPosition()
            # السماح بالأرقام والمسافات والشرطات والنقاط والنجوم والفواصل (العربية والإنجليزية)
            filtered_text = ''.join(c for c in text if c.isdigit() or c in ' -.,*،')
            if filtered_text != text:
                plots_input.setText(filtered_text)
                plots_input.setCursorPosition(min(cursor_pos, len(filtered_text)))

        plots_input.textChanged.connect(validate_plots_input)

        def toggle_plots_input():
            if radio_specific.isChecked():
                radio_all.setChecked(False)
                plots_input.setEnabled(True)
            else:
                radio_all.setChecked(True)
                plots_input.setEnabled(False)
                plots_input.clear()
            self.update_counter()
            self.update_current_row_summary()

        def toggle_all():
            if radio_all.isChecked():
                radio_specific.setChecked(False)
                plots_input.setEnabled(False)
                plots_input.clear()
            else:
                radio_specific.setChecked(True)
                plots_input.setEnabled(True)
            self.update_counter()
            self.update_current_row_summary()

        radio_specific.clicked.connect(toggle_plots_input)
        radio_all.clicked.connect(toggle_all)

        def on_rev_changed(value):
            if value > 0:
                manual_ref_input.setEnabled(True)
                manual_plot_input.setEnabled(True)
                manual_ref_label.setStyleSheet("color: red; font-weight: bold;")
                manual_plot_label.setStyleSheet("color: red; font-weight: bold;")
                radio_all.setEnabled(False)
                radio_specific.setEnabled(False)
                plots_input.setEnabled(False)
            else:
                manual_ref_input.setEnabled(False)
                manual_ref_input.clear()
                manual_plot_input.setEnabled(False)
                manual_plot_input.clear()
                manual_ref_label.setStyleSheet("")
                manual_plot_label.setStyleSheet("")
                radio_all.setEnabled(True)
                radio_specific.setEnabled(True)
                if radio_specific.isChecked():
                    plots_input.setEnabled(True)
            self.update_current_row_summary()

        revision_spin.valueChanged.connect(on_rev_changed)

        plot_scope_layout.addWidget(radio_all)
        plot_scope_layout.addWidget(radio_specific)
        plot_scope_layout.addWidget(plots_input)
        plot_scope_layout.addStretch()

        # حقل suffix — اقتراحات أثناء الكتابة (مثل شريط العناوين في المتصفح)
        suffix_edit = QLineEdit()
        suffix_edit.setPlaceholderText("اختياري: لاحقة لاسم الملف")
        suffix_edit.setMaxLength(50)
        suffix_edit.setStyleSheet(FIELD_STYLE_TALL)
        suffix_edit.setFixedHeight(36)
        suffix_edit.setMinimumWidth(220)
        suffix_edit.textChanged.connect(self.update_counter)
        suffix_edit.textChanged.connect(self.update_current_row_summary)
        suffix_edit.editingFinished.connect(lambda: self._add_to_suffix_history(suffix_edit.text().strip()))
        # ===== Dropdown تلقائي عند الضغط على suffix_edit =====
        from PyQt5.QtCore import QSize as _QSize, QTimer as _QTimer

        # نبني الـ dropdown كـ QFrame داخل النافذة الرئيسية (مش Popup منفصل)
        suffix_dropdown_frame = QFrame(self.main_window if self.main_window else self)
        suffix_dropdown_frame.setWindowFlags(Qt.ToolTip)
        suffix_dropdown_frame.setObjectName("suffixDropdown")
        suffix_dropdown_frame.setStyleSheet("""
            QFrame#suffixDropdown {
                background: white;
                border: 1.5px solid #3b82f6;
                border-radius: 8px;
            }
            QFrame#suffixDropdown QWidget {
                border: none;
                outline: none;
            }
        """)
        suffix_dropdown_frame.setVisible(False)

        _dd_layout = QVBoxLayout(suffix_dropdown_frame)
        _dd_layout.setContentsMargins(2, 2, 2, 2)
        _dd_layout.setSpacing(0)

        suffix_dd_list = QListWidget()
        suffix_dd_list.setFocusPolicy(Qt.NoFocus)
        suffix_dd_list.setSelectionMode(QAbstractItemView.NoSelection)
        suffix_dd_list.setStyleSheet("""
            QListWidget {
                border: none;
                background: transparent;
                font-size: 14px;
                color: #1e293b;
                outline: none;
            }
            QListWidget::item {
                padding: 0px;
                margin: 0px;
                border: none;
                border-bottom: 1px solid #f1f5f9;
                color: #1e293b;
                background: white;
            }
            QListWidget::item:hover { background-color: #eff6ff; }
            QListWidget::item:selected { background-color: transparent; border: none; }
            QListWidget::item:focus { border: none; outline: none; }
        """)
        _dd_layout.addWidget(suffix_dd_list)

        def _populate_suffix_dd(filter_text=""):
            suffix_dd_list.clear()
            history = self.suffix_history
            for s in history:
                if filter_text and filter_text.lower() not in s.lower():
                    continue
                item = QListWidgetItem()
                item_w = QWidget()
                item_w.setStyleSheet("background: transparent; border: none; outline: none;")
                item_l = QHBoxLayout(item_w)
                item_l.setContentsMargins(6, 2, 6, 2)
                item_l.setSpacing(6)
                lbl = QLabel(s)
                lbl.setStyleSheet("font-size: 14px; color: #1e293b; background: transparent;")
                lbl.setAttribute(Qt.WA_TransparentForMouseEvents)
                item_l.addWidget(lbl, 1)
                del_btn = QPushButton("✕")
                del_btn.setFixedSize(18, 18)
                del_btn.setFocusPolicy(Qt.NoFocus)
                del_btn.setStyleSheet("""
                    QPushButton {
                        background: transparent; border: none;
                        color: #94a3b8; font-size: 11px; font-weight: bold;
                    }
                    QPushButton:hover { color: #dc2626; background: #fee2e2; border-radius: 3px; }
                """)
                del_btn.setToolTip("حذف")
                del_btn.clicked.connect(lambda _, t=s: _dd_delete(t))
                item_l.addWidget(del_btn)
                suffix_dd_list.addItem(item)
                item.setSizeHint(_QSize(0, 32))
                suffix_dd_list.setItemWidget(item, item_w)
                item.setData(Qt.UserRole, s)

        def _dd_delete(text):
            self._remove_from_suffix_history(text)
            cur = suffix_edit.text()
            _populate_suffix_dd(cur)
            if suffix_dd_list.count() == 0:
                suffix_dropdown_frame.hide()
            else:
                _dd_reposition()

        def _dd_reposition():
            if not self.main_window:
                return
            from PyQt5.QtWidgets import QApplication as _QApp
            pos = suffix_edit.mapToGlobal(suffix_edit.rect().bottomLeft())
            w = max(suffix_edit.width(), 240)
            cnt = suffix_dd_list.count()
            h = min(cnt, 8) * 34 + 8
            screen = _QApp.screenAt(pos)
            if screen is None:
                screen = _QApp.primaryScreen()
            sr = screen.availableGeometry()
            x = pos.x()
            y = pos.y() + 2
            if y + h > sr.bottom():
                y = suffix_edit.mapToGlobal(suffix_edit.rect().topLeft()).y() - h - 2
            if x + w > sr.right():
                x = sr.right() - w - 4
            suffix_dropdown_frame.setGeometry(x, y, w, h)

        def _show_suffix_dd():
            if not self.suffix_history:
                return
            _populate_suffix_dd(suffix_edit.text())
            if suffix_dd_list.count() == 0:
                return
            _dd_reposition()
            suffix_dropdown_frame.show()
            suffix_dropdown_frame.raise_()

        def _hide_suffix_dd():
            suffix_dropdown_frame.hide()

        def _on_suffix_focus_in(event):
            QLineEdit.focusInEvent(suffix_edit, event)
            _show_suffix_dd()

        def _on_suffix_focus_out(event):
            QLineEdit.focusOutEvent(suffix_edit, event)
            _QTimer.singleShot(200, _hide_suffix_dd)

        def _on_suffix_item_clicked(item):
            text = item.data(Qt.UserRole)
            if text:
                suffix_edit.blockSignals(True)
                suffix_edit.setText(text)
                suffix_edit.blockSignals(False)
                _hide_suffix_dd()

        def _on_suffix_text_changed_dropdown(text):
            if not suffix_edit.hasFocus():
                return
            if not text.strip():
                # النص اتمسح - اعرض كل الـ history
                _show_suffix_dd()
            else:
                _populate_suffix_dd(text)
                if suffix_dd_list.count() == 0:
                    _hide_suffix_dd()
                else:
                    _dd_reposition()

        suffix_edit.focusInEvent  = _on_suffix_focus_in
        suffix_edit.focusOutEvent = _on_suffix_focus_out
        suffix_dd_list.itemClicked.connect(_on_suffix_item_clicked)
        suffix_edit.textChanged.connect(_on_suffix_text_changed_dropdown)

        suffix_edit.setContextMenuPolicy(Qt.CustomContextMenu)
        def _suffix_context_menu(pos):
            menu = QMenu(self)
            menu.addAction("إدارة السوابق المحفوظة...", lambda: self._show_suffix_history_dialog(suffix_edit))
            menu.exec_(suffix_edit.mapToGlobal(pos))
        suffix_edit.customContextMenuRequested.connect(_suffix_context_menu)

        # المرفقات
        attach_layout = QVBoxLayout()
        attach_layout.setSpacing(5)
        attach_buttons_layout = QHBoxLayout()
        btn_add_attach = QPushButton("＋  إضافة مرفق PDF")
        btn_add_attach.setCursor(Qt.PointingHandCursor)
        btn_add_attach.setStyleSheet("""
            QPushButton {
                background-color: #0288d1;
                color: white;
                font-weight: bold;
                padding: 5px 10px;
                border-radius: 5px;
                border: none;
                font-size: 14px;
            }
            QPushButton:hover { background-color: #0277bd; }
            QPushButton:pressed { background-color: #01579b; }
        """)
        btn_clear_attach = QPushButton("🗑  مسح الكل")
        btn_clear_attach.setCursor(Qt.PointingHandCursor)
        btn_clear_attach.setStyleSheet("""
            QPushButton {
                background-color: #eceff1;
                color: #546e7a;
                padding: 5px 10px;
                border-radius: 5px;
                border: 1px solid #cfd8dc;
                font-size: 14px;
            }
            QPushButton:hover { background-color: #ffebee; color: #c62828; border-color: #ef9a9a; }
            QPushButton:pressed { background-color: #ffcdd2; }
        """)
        attach_buttons_layout.addWidget(btn_add_attach)
        attach_buttons_layout.addWidget(btn_clear_attach)
        attach_buttons_layout.addStretch()

        # ===== منطقة المرفقات مع إعادة الترتيب =====
        attach_paths = []

        # قائمة المرفقات المخصصة
        # DropZone يلف الـ list ويستقبل الـ drops
        drop_zone = DropZoneWidget()
        drop_zone_layout = QVBoxLayout(drop_zone)
        drop_zone_layout.setContentsMargins(0, 0, 0, 0)
        drop_zone_layout.setSpacing(0)

        attach_list_widget = AttachListWidget()
        # ربط إشارة نقل الصفوف لتحديث الترتيب والمصفوفة بعد السحب
        attach_list_widget.model().rowsMoved.connect(
            lambda: (_rebuild_widgets(), _sync_paths_from_list())
        )
        attach_list_widget.setMinimumHeight(70)
        attach_list_widget.setMaximumHeight(180)
        attach_list_widget.setDragDropMode(QAbstractItemView.InternalMove)
        attach_list_widget.setDefaultDropAction(Qt.MoveAction)
        attach_list_widget.setFocusPolicy(Qt.NoFocus)
        attach_list_widget.setStyleSheet("""
        QListWidget {
            background-color: #f8fafc;
            border: 2px dashed #cbd5e1;
            border-radius: 10px;
            outline: none;
        }

        QListWidget::item {
            background-color: white;
            border: none;
            border-radius: 8px;
            padding: 6px 12px;
            margin: 6px;
            color: #1e293b;
        }

        QListWidget::item:hover {
            background-color: #eff6ff;
        }

        QListWidget::item:selected {
            background-color: #dbeafe;
            color: #1e40af;
        }
        """)
        attach_list_widget.setFocusPolicy(Qt.NoFocus)

        # placeholder item داخل الـ list نفسها
        _PLACEHOLDER_KEY = "__placeholder__"

        def _show_placeholder():
            for i in range(attach_list_widget.count()):
                if attach_list_widget.item(i).data(Qt.UserRole) == _PLACEHOLDER_KEY:
                    return
            item = QListWidgetItem("📂  اسحب ملفات PDF هنا أو اضغط ＋")
            item.setData(Qt.UserRole, _PLACEHOLDER_KEY)
            item.setFlags(Qt.NoItemFlags)
            item.setTextAlignment(Qt.AlignCenter)
            item.setForeground(QColor("#78909c"))
            attach_list_widget.insertItem(0, item)

        def _hide_placeholder():
            for i in range(attach_list_widget.count()):
                if attach_list_widget.item(i).data(Qt.UserRole) == _PLACEHOLDER_KEY:
                    attach_list_widget.takeItem(i)
                    return

        def _has_real_files():
            for i in range(attach_list_widget.count()):
                if attach_list_widget.item(i).data(Qt.UserRole) != _PLACEHOLDER_KEY:
                    return True
            return False

        _show_placeholder()

        def _sync_paths_from_list():
            """مزامنة attach_paths مع ترتيب الـ list"""
            attach_paths.clear()
            for i in range(attach_list_widget.count()):
                item = attach_list_widget.item(i)
                if item:
                    d = item.data(Qt.UserRole)
                    if d and d != _PLACEHOLDER_KEY:
                        attach_paths.append(d)

        DZ_NORMAL = """
        QWidget {
            background: #f8fafc;
            border: 2px dashed #cbd5e1;
            border-radius: 10px;
        }
        """

        DZ_HAS = """
        QWidget {
            background: #ecfeff;
            border: 2px solid #67e8f9;
            border-radius: 10px;
        }
        """

        DZ_HOVER = """
        QWidget {
            background: #eff6ff;
            border: 2px dashed #3b82f6;
            border-radius: 10px;
        }
        """
        LIST_STYLE = """
            QListWidget {
                background: transparent;
                border: none;
                outline: none;
                font-size: 13px;
            }
            QListWidget::item {
                background: white;
                border: 1px solid #e2e8f0;
                border-radius: 5px;
                padding: 5px 8px;
                margin: 2px 4px;
                color: #1e293b;
            }
            QListWidget::item:hover { border-color: #93c5fd; background: #eff6ff; }
            QListWidget::item:selected { border-color: #3b82f6; background: #dbeafe; color: #1e40af; }
        """
        attach_list_widget.setStyleSheet(LIST_STYLE)
        drop_zone.setStyles(DZ_NORMAL, DZ_HOVER)
        drop_zone.setStyleSheet(DZ_NORMAL)
        # فلتر يمرر الـ drag events عبر الـ QScrollArea
        ScrollAreaDropFilter(self.scroll, drop_zone)

        def _refresh_zone_style():
            if _has_real_files():
                drop_zone.setStyleSheet(DZ_HAS)
            else:
                drop_zone.setStyleSheet(DZ_NORMAL)

        BTN_ARROW = """
        QPushButton {
            background: transparent;
            border: none;
            font-size: 11px;
            color: #64748b;
            min-width: 20px;
            max-width: 20px;
            min-height: 16px;
            max-height: 16px;
        }

        QPushButton:hover {
            background-color: #e0f2fe;
            border-radius: 4px;
            color: #1d4ed8;
        }

        QPushButton:pressed {
            background-color: #bfdbfe;
        }

        QPushButton:focus {
            outline: none;
        }
        """

        def _move_item(row, direction):
            """تحريك العنصر لأعلى (-1) أو لأسفل (+1)"""
            new_row = row + direction
            count = attach_list_widget.count()
            # تخطي placeholder
            while 0 <= new_row < count:
                if attach_list_widget.item(new_row).data(Qt.UserRole) != _PLACEHOLDER_KEY:
                    break
                new_row += direction
            if new_row < 0 or new_row >= count:
                return
            # swap data بدل نقل الـ item
            it_a = attach_list_widget.item(row)
            it_b = attach_list_widget.item(new_row)
            path_a = it_a.data(Qt.UserRole)
            path_b = it_b.data(Qt.UserRole)
            tip_a  = it_a.toolTip()
            tip_b  = it_b.toolTip()
            it_a.setData(Qt.UserRole, path_b)
            it_a.setToolTip(tip_b)
            it_b.setData(Qt.UserRole, path_a)
            it_b.setToolTip(tip_a)
            attach_list_widget.setCurrentRow(new_row)
            _rebuild_widgets()
            _renumber_items()
            _sync_paths_from_list()

        def _make_item_widget(row):
            w = QWidget()
            w.setStyleSheet("background: transparent; border: none;")
            # إضافة tooltip للمسار الكامل
            it = attach_list_widget.item(row)
            path = it.data(Qt.UserRole) if it else ""
            if path and path != _PLACEHOLDER_KEY:
                w.setToolTip(path)  # ← هنا إضافة tooltip للصف كامل

            l = QHBoxLayout(w)
            l.setContentsMargins(4, 0, 4, 0)
            l.setSpacing(4)

            lbl_num = QLabel()
            lbl_num.setFixedWidth(18)
            lbl_num.setAlignment(Qt.AlignCenter)
            lbl_num.setStyleSheet("font-size: 11px; color: #64748b; background: transparent;")

            ico = QLabel("📄")
            ico.setStyleSheet("background: transparent; font-size: 12px;")
            ico.setFixedWidth(18)
            ico.setAttribute(Qt.WA_TransparentForMouseEvents)  # ← شفاف للماوس

            name_lbl = QLabel(os.path.basename(path))
            name_lbl.setToolTip(path)
            name_lbl.setStyleSheet("font-size: 12px; color: #1e293b; background: transparent;")
            name_lbl.setAttribute(Qt.WA_TransparentForMouseEvents)  # ← شفاف للماوس

            btn_up   = QPushButton("▲")
            btn_up.setStyleSheet(BTN_ARROW)
            btn_up.setCursor(Qt.PointingHandCursor)
            btn_up.setToolTip("تحريك لأعلى")
            btn_up.clicked.connect(lambda _, r=row: _move_item(r, -1))

            btn_down = QPushButton("▼")
            btn_down.setStyleSheet(BTN_ARROW)
            btn_down.setCursor(Qt.PointingHandCursor)
            btn_down.setToolTip("تحريك لأسفل")
            btn_down.clicked.connect(lambda _, r=row: _move_item(r, +1))

            l.addWidget(lbl_num)
            l.addWidget(ico)
            l.addWidget(name_lbl, 1)
            l.addWidget(btn_up)
            l.addWidget(btn_down)

            # تجاهل أحداث الضغط على الخلفية لتمريرها للقائمة (تبدأ السحب)
            def ignore_mouse_press(event):
                event.ignore()
            w.mousePressEvent = ignore_mouse_press

            return w

        def _rebuild_widgets():
            """إعادة بناء كل الـ widgets بعد تغيير الترتيب"""
            num = 1
            for i in range(attach_list_widget.count()):
                it = attach_list_widget.item(i)
                if not it or it.data(Qt.UserRole) == _PLACEHOLDER_KEY:
                    continue
                w = _make_item_widget(i)
                # تحديث رقم العنصر
                w.layout().itemAt(0).widget().setText(str(num))
                num += 1
                attach_list_widget.setItemWidget(it, w)

        def add_attachment_file(file_path):
            if not file_path.lower().endswith('.pdf'):
                return
            for i in range(attach_list_widget.count()):
                if attach_list_widget.item(i).data(Qt.UserRole) == file_path:
                    return
            _hide_placeholder()
            item = QListWidgetItem()
            item.setData(Qt.UserRole, file_path)
            item.setToolTip(file_path)
            # تمكين السحب + التحديد + التمكين
            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled)
            item.setSizeHint(__import__('PyQt5.QtCore', fromlist=['QSize']).QSize(0, 32))
            attach_list_widget.addItem(item)
            _rebuild_widgets()
            _renumber_items()
            _sync_paths_from_list()
            _refresh_zone_style()

        def _renumber_items():
            """تحديث نصوص الـ items (للمزامنة الداخلية فقط)"""
            _sync_paths_from_list()

        def add_attachment():
            fnames, _ = QFileDialog.getOpenFileNames(self, 'اختر ملفات PDF', '', "PDF Files (*.pdf)")
            for fname in fnames:
                add_attachment_file(fname)

        def remove_attachment(fname=None, widget=None):
            """حذف العنصر المحدد أو المحدد بـ fname"""
            if fname:
                for i in range(attach_list_widget.count()):
                    if attach_list_widget.item(i).data(Qt.UserRole) == fname:
                        attach_list_widget.takeItem(i)
                        break
            else:
                row = attach_list_widget.currentRow()
                item = attach_list_widget.item(row) if row >= 0 else None
                if item and item.data(Qt.UserRole) != _PLACEHOLDER_KEY:
                    attach_list_widget.takeItem(row)
            if not _has_real_files():
                _show_placeholder()
            else:
                _rebuild_widgets()
            _sync_paths_from_list()
            _refresh_zone_style()

        def clear_attachments():
            attach_list_widget.clear()
            attach_paths.clear()
            _show_placeholder()
            _refresh_zone_style()

        # حذف بزرار Delete من الكيبورد
        def keyPressEvent_list(event):
            if event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
                row = attach_list_widget.currentRow()
                item = attach_list_widget.item(row) if row >= 0 else None
                if item and item.data(Qt.UserRole) != _PLACEHOLDER_KEY:
                    remove_attachment()
            else:
                QListWidget.keyPressEvent(attach_list_widget, event)

        attach_list_widget.keyPressEvent = keyPressEvent_list

        # ربط signal السحب الخارجي
        def _on_files_dropped(paths):
            for p in paths:
                add_attachment_file(p)
            _refresh_zone_style()

        drop_zone.filesDropped.connect(_on_files_dropped)

        btn_add_attach.clicked.connect(add_attachment)
        btn_clear_attach.clicked.connect(clear_attachments)

        # زرار حذف العنصر المحدد
        btn_remove_selected = QPushButton("✕  حذف المحدد")
        btn_remove_selected.setCursor(Qt.PointingHandCursor)
        btn_remove_selected.setStyleSheet("""
            QPushButton {
                background: transparent; color: #64748b;
                border: 1px solid #e2e8f0; border-radius: 5px;
                font-size: 13px; padding: 4px 8px;
            }
            QPushButton:hover { background: #fef2f2; color: #dc2626; border-color: #fca5a5; }
        """)
        btn_remove_selected.clicked.connect(lambda: remove_attachment())
        attach_buttons_layout.addWidget(btn_remove_selected)

        attach_layout.addLayout(attach_buttons_layout)
        drop_zone_layout.addWidget(attach_list_widget)
        attach_layout.addWidget(drop_zone)

        # hint الترتيب - ثابت تحت الـ list
        order_hint = QLabel("☰  اسحب للترتيب  |  Del للحذف")
        order_hint.setAlignment(Qt.AlignCenter)
        order_hint.setStyleSheet("color: #64748b; font-size: 13px; background: transparent; padding: 1px;")
        order_hint.setVisible(False)
        attach_layout.addWidget(order_hint)

        orig_refresh = _refresh_zone_style
        def _refresh_zone_style_ext():
            orig_refresh()
            order_hint.setVisible(_has_real_files() and attach_list_widget.count() > 2)
        _refresh_zone_style = _refresh_zone_style_ext

        attach_container = QWidget()
        attach_container.setLayout(attach_layout)

        delete_layout = QHBoxLayout()
        btn_delete = QPushButton("🗑  حذف هذا الطلب")
        btn_delete.setCursor(Qt.PointingHandCursor)
        btn_delete.setStyleSheet("""
            QPushButton {
                background-color: transparent;
                color: #dc3545;
                border: 1px solid #dc3545;
                border-radius: 6px;
                font-size: 13px;
                padding: 5px 12px;
            }
            QPushButton:hover { background-color: #dc3545; color: white; }
            QPushButton:pressed { background-color: #c82333; color: white; }
        """)
        btn_delete.clicked.connect(lambda: self.delete_row(row_container, {
            'desc': desc_edit, 'radio_all': radio_all,
            'plots_input': plots_input, 'revision': revision_spin,
            'manual_ref_input': manual_ref_input, 'manual_plot_input': manual_plot_input,
            'suffix': suffix_edit
        }))
        delete_layout.addStretch()
        delete_layout.addWidget(btn_delete)

        # إضافة الحقول إلى تخطيط التفاصيل
        details_layout.addRow("وصف الأعمال:", desc_edit)
        details_layout.addRow("", suggestions_container)
        details_layout.addRow("رقم المراجعة:", revision_layout)
        details_layout.addRow("", manual_ref_layout)
        details_layout.addRow("", manual_plot_layout)
        details_layout.addRow("نطاق القطع:", plot_scope_layout)
        details_layout.addRow("اللاحقة (suffix):", suffix_edit)
        details_layout.addRow("المرفقات (PDF):", attach_container)
        details_layout.addRow("", delete_layout)

        # تجميع الأجزاء
        main_layout.addWidget(header_widget)
        main_layout.addWidget(details_widget)

        # ربط زر التوسيع/الطي
        def toggle_details(checked):
            details_widget.setVisible(checked)
            expand_btn.setText("▼" if checked else "▶")
            if checked:
                header_widget.setStyleSheet("""
                    QWidget#cardHeader {
                        background-color: #f1f3f5;
                        border-top-left-radius: 10px;
                        border-top-right-radius: 10px;
                        border-bottom: 1px solid #dee2e6;
                    }
                """)
            else:
                header_widget.setStyleSheet("""
                    QWidget#cardHeader {
                        background-color: #f1f3f5;
                        border-radius: 10px;
                    }
                """)
        expand_btn.toggled.connect(toggle_details)

        # إضافة الحاوية إلى التمرير
        self.scroll_layout.addWidget(row_container)

        # تخزين المراجع
        self.rows.append({
            'container': row_container,
            'summary_label': summary_label,
            'expand_btn': expand_btn,
            'details': details_widget,
            'desc': desc_edit,
            'suggestions_list': suggestions_list,
            'attach_paths': attach_paths,
            'radio_all': radio_all,
            'plots_input': plots_input,
            'revision': revision_spin,
            'manual_ref_input': manual_ref_input,
            'manual_plot_input': manual_plot_input,
            'suffix': suffix_edit,
            'attach_list_widget': attach_list_widget,
            'update_suggestions': update_suggestions_for_row
        })

        # تحديث الملخص الأولي
        self.update_row_summary(len(self.rows) - 1)

    def _load_suffix_history(self):
        """تحميل قائمة سوابق الـ suffix من الملف"""
        try:
            if os.path.exists(self.SUFFIX_HISTORY_FILE):
                with open(self.SUFFIX_HISTORY_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.suffix_history = data if isinstance(data, list) else []
            else:
                self.suffix_history = []
        except Exception:
            self.suffix_history = []

    def _save_suffix_history(self):
        """حفظ قائمة سوابق الـ suffix في الملف"""
        try:
            with open(self.SUFFIX_HISTORY_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.suffix_history, f, ensure_ascii=False, indent=2)
        except Exception as e:
            log_error(f"حفظ سوابق الـ suffix: {e}")

    def _add_to_suffix_history(self, text: str):
        """إضافة نص إلى السوابق (بدون تكرار) وتحديث الاقتراحات"""
        text = (text or "").strip()
        if not text:
            return
        while text in self.suffix_history:
            self.suffix_history.remove(text)
        self.suffix_history.insert(0, text)
        self.suffix_history = self.suffix_history[: self.SUFFIX_HISTORY_MAX]
        self._save_suffix_history()
        self.suffix_completer_model.setStringList(self.suffix_history)

    def _remove_from_suffix_history(self, text: str):
        """حذف سطر من السوابق وتحديث الاقتراحات"""
        if text in self.suffix_history:
            self.suffix_history.remove(text)
            self._save_suffix_history()
            self.suffix_completer_model.setStringList(self.suffix_history)

    def _show_suffix_history_dialog(self, suffix_edit_ref=None):
        """نافذة إدارة السوابق: عرض القائمة وحذف أي سطر"""
        dlg = QDialog(self)
        dlg.setWindowTitle("إدارة السوابق المحفوظة")
        dlg.setLayoutDirection(Qt.RightToLeft)
        layout = QVBoxLayout(dlg)
        lst = QListWidget()
        lst.setStyleSheet("QListWidget { font-size: 13px; } QListWidget::item { padding: 6px; }")
        for s in self.suffix_history:
            item = QListWidgetItem(s)
            item.setData(Qt.UserRole, s)
            lst.addItem(item)
        def on_item_clicked(item):
            if item and suffix_edit_ref and item.data(Qt.UserRole):
                suffix_edit_ref.setText(item.data(Qt.UserRole))
                dlg.accept()
        lst.itemDoubleClicked.connect(on_item_clicked)
        layout.addWidget(QLabel("انقر مرتين لملء خانة اللاحقة • زر الماوس الأيمن لحذف سطر:"))
        layout.addWidget(lst)
        lst.setContextMenuPolicy(Qt.CustomContextMenu)
        def on_context(pos):
            item = lst.itemAt(pos)
            if not item:
                return
            s = item.data(Qt.UserRole) or item.text()
            menu = QMenu(dlg)
            menu.addAction("حذف هذا السطر", lambda: self._remove_from_suffix_history(s) or lst.takeItem(lst.row(item)))
            menu.exec_(lst.mapToGlobal(pos))
        lst.customContextMenuRequested.connect(on_context)
        btn_close = QPushButton("إغلاق")
        btn_close.clicked.connect(dlg.accept)
        layout.addWidget(btn_close)
        dlg.exec_()

    def refresh_suggestions(self):
        """تحديث قوائم الاقتراحات في جميع الصفوف"""
        for row in self.rows:
            if 'update_suggestions' in row:
                row['update_suggestions']()

    def delete_suggestion(self, suggestion_text: str, row_index: int):
        """حذف الاقتراح المحدد من قاعدة البيانات وتحديث القوائم"""
        # تأكيد الحذف
        reply = QMessageBox.question(self, "تأكيد الحذف",
                                     f"هل أنت متأكد من حذف الاقتراح:\n\n{suggestion_text}",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        # حذف من قاعدة البيانات
        suggestions_db.remove_suggestion(self.code, suggestion_text)

        # تحديث قائمة الاقتراحات في جميع الصفوف (قد تكون موجودة في عدة صفوف)
        for row in self.rows:
            if 'update_suggestions' in row:
                row['update_suggestions']()

        # لا نعرض رسالة "تم الحذف" - نكتفي بالتأكيد

    def update_counter(self):
        if self.main_window:
            self.main_window.calculate_expected_files()

    def update_current_row_summary(self):
        for i in range(len(self.rows)):
            self.update_row_summary(i)

    def delete_row(self, row_container, row_data):
        reply = QMessageBox.question(self, 'تأكيد الحذف',
                                     'هل أنت متأكد من حذف هذا الطلب؟',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            row_container.setParent(None)
            self.rows = [r for r in self.rows if r.get('container') != row_container]
            self.update_row_numbers()
            self.update_counter()

    def update_row_numbers(self):
        for i in range(len(self.rows)):
            self.update_row_summary(i)

    def update_row_summary(self, row_index):
        if row_index >= len(self.rows):
            return
        row = self.rows[row_index]
        if 'summary_label' not in row:
            return

        base_title = f"طلب رقم {row_index + 1}"
        summary_parts = []
        rev_value = row['revision'].value()

        if rev_value > 0:
            summary_parts.append(f"REV{rev_value:02d}")
            manual_plot = row['manual_plot_input'].text().strip()
            if manual_plot:
                summary_parts.append(f"قطعة: {manual_plot}")
        else:
            if row['radio_all'].isChecked():
                summary_parts.append("كل القطع")
            else:
                plots_text = row['plots_input'].text().strip()
                if plots_text:
                    summary_parts.append(f"قطع: {plots_text}")

        suffix_text = row['suffix'].text().strip()
        if suffix_text:
            summary_parts.append(f"({suffix_text})")

        full_title = f"{base_title} [{' | '.join(summary_parts)}]" if summary_parts else base_title
        row['summary_label'].setText(full_title)

    def clear_tab(self):
        """مسح جميع الطلبات في هذا التبويب وإضافة طلب فارغ واحد"""
        reply = QMessageBox.question(
            self, "تأكيد المسح",
            f"هل أنت متأكد من مسح جميع الطلبات في تبويب {self.name}؟",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        # حذف جميع الصفوف الموجودة
        for row in self.rows[:]:  # نسخة لتجنب مشاكل التعديل أثناء التكرار
            row['container'].setParent(None)
            row['container'].deleteLater()
        self.rows.clear()

        # إعادة تعيين بداية الترقيم
        self.serial_input.setText("1")

        # إضافة صف فارغ جديد
        self.add_row()

        # تحديث العداد في النافذة الرئيسية
        if self.main_window:
            self.main_window.calculate_expected_files()

# -------------------- النافذة الرئيسية --------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("WIR Toilets Project by Yasser Hamouda v3.4")  # تم تغيير الإصدار
        self.setWindowIcon(QIcon(resource_path("icon.ico")))  # تعيين أيقونة البرنامج
        self.setGeometry(100, 100, 1400, 900)  # تم تعديل الارتفاع إلى 900
        self.setLayoutDirection(Qt.RightToLeft)
        self.process_thread = None

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_h_layout = QHBoxLayout(central_widget)
        main_h_layout.setContentsMargins(0, 0, 0, 0)
        main_h_layout.setSpacing(0)

        main_content_widget = QWidget()
        main_layout = QVBoxLayout(main_content_widget)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(8)

        # ===== البيانات العامة - Modern Card =====
        CTRL_H = 18   # ارتفاع موحد للعناصر (مصغّر)
        FIELD_STYLE_MAIN = """
            QLineEdit {
                border: 1.5px solid #94a3b8;
                border-radius: 6px;
                padding: 2px 6px;
                font-size: 14px;
                background: #ffffff;
                color: #0f172a;
                min-height: 18px;
            }
            QLineEdit:focus {
                border: 2px solid #3b82f6;
                background: white;
            }
            QLineEdit:disabled {
                background: #e9ecef;
                color: #9ca3af;
                border: 1.5px dashed #cbd5e1;
            }
        """
        SPIN_STYLE_MAIN = """
            QSpinBox {
                border: 1.5px solid #94a3b8;
                border-radius: 6px;
                padding: 1px 4px;
                font-size: 14px;
                background: white;
                color: #0f172a;
                min-height: 18px;
            }
            QSpinBox:focus { border: 2px solid #3b82f6; }
            QSpinBox:disabled {
                background: #e9ecef;
                color: #9ca3af;
                border: 1.5px dashed #cbd5e1;
            }
        """
        LBL_STYLE = "font-size: 14px; font-weight: 700; color: #334155;"

        header_group = QWidget()
        header_group.setObjectName("headerCard")
        header_group.setStyleSheet("""
            QWidget#headerCard {
                background-color: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 12px;
            }
        """)
        header_outer = QVBoxLayout(header_group)
        header_outer.setContentsMargins(0, 0, 0, 0)
        header_outer.setSpacing(0)

        # شريط العنوان
        hdr_title_bar = QWidget()
        hdr_title_bar.setObjectName("hdrTitleBar")
        hdr_title_bar.setStyleSheet("""
            QWidget#hdrTitleBar {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 #1e293b, stop:1 #334155);
                border-top-left-radius: 12px;
                border-top-right-radius: 12px;
            }
        """)
        hdr_title_bar.setFixedHeight(40)
        hdr_title_layout = QHBoxLayout(hdr_title_bar)
        hdr_title_layout.setContentsMargins(16, 0, 16, 0)

        hdr_icon = QLabel("🗂")
        hdr_icon.setStyleSheet("font-size: 15px; background: transparent;")
        hdr_text = QLabel("البيانات العامة")
        hdr_text.setStyleSheet("color: #f1f5f9; font-size: 15px; font-weight: bold; background: transparent;")
        hdr_title_layout.addWidget(hdr_icon)
        hdr_title_layout.addWidget(hdr_text)
        hdr_title_layout.addStretch()
        header_outer.addWidget(hdr_title_bar)

        # المحتوى
        hdr_body = QWidget()
        hdr_body.setStyleSheet("background: transparent;")
        hdr_body_layout = QVBoxLayout(hdr_body)
        hdr_body_layout.setContentsMargins(16, 14, 16, 14)
        hdr_body_layout.setSpacing(14)

        # --- صف أرقام القطع ---
        plots_row = QHBoxLayout()
        plots_row.setSpacing(10)

        plots_lbl = QLabel("🏘  أرقام القطع")
        plots_lbl.setStyleSheet(LBL_STYLE)
        plots_lbl.setFixedWidth(130)   # عرض كافٍ لظهور العنوان كاملاً

        self.plots_input = QLineEdit()
        self.plots_input.setPlaceholderText("مثال: 101 102 105  أو  101-105-110")
        self.plots_input.setFixedHeight(26)   # أكبر من CTRL_H ليتسع النص
        self.plots_input.setStyleSheet(FIELD_STYLE_MAIN)
        self.plots_input.textChanged.connect(self.calculate_expected_files)

        def validate_main_plots_input(text):
            cursor_pos = self.plots_input.cursorPosition()
            filtered_text = ''.join(c for c in text if c.isdigit() or c in ' -.,*،')
            if filtered_text != text:
                self.plots_input.setText(filtered_text)
                self.plots_input.setCursorPosition(min(cursor_pos, len(filtered_text)))

        self.plots_input.textChanged.connect(validate_main_plots_input)

        plots_row.addWidget(plots_lbl)
        plots_row.addWidget(self.plots_input, 1)
        hdr_body_layout.addLayout(plots_row)

        # فاصل خفيف
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #f1f5f9; background: #f1f5f9; border: none; max-height: 1px;")
        hdr_body_layout.addWidget(sep)

        # --- صف التاريخ والوقت ---
        dt_row = QHBoxLayout()
        dt_row.setSpacing(16)
        dt_row.setAlignment(Qt.AlignVCenter)

        # --- بلوك التاريخ ---
        date_block = QHBoxLayout()
        date_block.setSpacing(6)
        date_block.setAlignment(Qt.AlignVCenter)

        date_lbl = QLabel("📅  التاريخ")
        date_lbl.setStyleSheet(LBL_STYLE + f" line-height: {CTRL_H}px; padding: 2px 0 0 0; margin: 0;")
        date_lbl.setFixedWidth(90)   # عرض كافٍ لظهور "التاريخ" كاملاً
        date_lbl.setFixedHeight(CTRL_H)
        date_lbl.setAlignment(Qt.AlignVCenter | Qt.AlignLeading)

        # حقل التاريخ بدون Fusion style
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(False)
        self.date_edit.setButtonSymbols(QDateEdit.NoButtons)
        self.date_edit.setDisplayFormat("dd/MM/yyyy")
        self.date_edit.setFixedWidth(108)
        self.date_edit.setFixedHeight(CTRL_H)
        self.date_edit.setCurrentSection(QDateEdit.DaySection)
        self.date_edit.setStyleSheet("""
            QDateEdit {
                border: 1.5px solid #94a3b8;
                border-radius: 7px;
                padding: 4px 8px;
                font-size: 14px;
                background: white;
                color: #0f172a;
                min-height: 18px;
            }
            QDateEdit:focus { border: 2px solid #3b82f6; }
        """)
        self.date_edit.dateChanged.connect(self.update_day_label)

        # أسهم اليوم (فوق/تحت) كزرارين منفصلين أنيقين
        arrows_widget = QWidget()
        arrows_widget.setFixedSize(20, CTRL_H)
        arrows_layout = QVBoxLayout(arrows_widget)
        arrows_layout.setContentsMargins(0, 0, 0, 0)
        arrows_layout.setSpacing(1)

        btn_day_up = QPushButton("▲")
        btn_day_up.setFixedHeight(CTRL_H // 2 - 1)
        btn_day_up.setCursor(Qt.PointingHandCursor)
        btn_day_up.setToolTip("يوم للأمام")
        btn_day_up.setStyleSheet("""
            QPushButton {
                background: #f1f5f9; border: 1px solid #e2e8f0;
                border-radius: 3px; font-size: 7px; color: #475569;
                padding: 0;
            }
            QPushButton:hover { background: #dbeafe; color: #1d4ed8; }
            QPushButton:pressed { background: #bfdbfe; }
        """)
        btn_day_up.clicked.connect(lambda: self.date_edit.setDate(self.date_edit.date().addDays(1)))

        btn_day_down = QPushButton("▼")
        btn_day_down.setFixedHeight(CTRL_H // 2 - 1)
        btn_day_down.setCursor(Qt.PointingHandCursor)
        btn_day_down.setToolTip("يوم للخلف")
        btn_day_down.setStyleSheet("""
            QPushButton {
                background: #f1f5f9; border: 1px solid #e2e8f0;
                border-radius: 3px; font-size: 7px; color: #475569;
                padding: 0;
            }
            QPushButton:hover { background: #dbeafe; color: #1d4ed8; }
            QPushButton:pressed { background: #bfdbfe; }
        """)
        btn_day_down.clicked.connect(lambda: self.date_edit.setDate(self.date_edit.date().addDays(-1)))

        arrows_layout.addWidget(btn_day_up)
        arrows_layout.addWidget(btn_day_down)

        self.btn_calendar = QPushButton("📆")
        self.btn_calendar.setFixedSize(CTRL_H, CTRL_H)
        self.btn_calendar.setToolTip("فتح التقويم")
        self.btn_calendar.setCursor(Qt.PointingHandCursor)
        self.btn_calendar.setStyleSheet("""
            QPushButton {
                background: #f1f5f9;
                border: 1px solid #e2e8f0;
                border-radius: 7px;
                font-size: 14px;
            }
            QPushButton:hover { background: #dbeafe; border-color: #93c5fd; }
        """)

        def open_calendar():
            from PyQt5.QtWidgets import QApplication as _QApp
            dlg = QDialog(self)
            dlg.setWindowTitle("اختر التاريخ")
            dlg.setWindowFlags(Qt.Popup | Qt.FramelessWindowHint)
            dlg.setStyleSheet("""
                QDialog {
                    background: white;
                    border: 1px solid #cbd5e1;
                    border-radius: 10px;
                }
            """)
            layout = QVBoxLayout(dlg)
            layout.setContentsMargins(6, 6, 6, 6)
            cal = QCalendarWidget(dlg)
            cal.setSelectedDate(self.date_edit.date())
            cal.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
            cal.setFixedSize(380, 260)
            self.customize_calendar_widget(cal)
            # stylesheet لأسماء الأيام
            cal.setStyleSheet("""
                QCalendarWidget QWidget { font-size: 12px; }
                QCalendarWidget QAbstractItemView {
                    font-size: 13px;
                    selection-background-color: #3b82f6;
                    selection-color: white;
                }
                QCalendarWidget QToolButton {
                    font-size: 13px; font-weight: bold;
                    color: white; background: transparent;
                }
                QCalendarWidget QMenu { font-size: 12px; }
                QCalendarWidget #qt_calendar_navigationbar {
                    background: #1e293b;
                    border-radius: 8px 8px 0 0;
                    padding: 4px;
                }
                QCalendarWidget #qt_calendar_prevmonth,
                QCalendarWidget #qt_calendar_nextmonth {
                    color: white; font-size: 16px;
                    background: transparent; border: none;
                    padding: 2px 8px;
                }
                QCalendarWidget #qt_calendar_prevmonth:hover,
                QCalendarWidget #qt_calendar_nextmonth:hover {
                    background: #334155; border-radius: 4px;
                }
            """)
            layout.addWidget(cal)
            def on_date_selected(date):
                self.date_edit.setDate(date)
                dlg.accept()
            cal.clicked.connect(on_date_selected)

            # حساب الموضع مع ضمان بقاء التقويم داخل الشاشة
            dlg.adjustSize()
            btn_pos = self.btn_calendar.mapToGlobal(self.btn_calendar.rect().bottomLeft())
            screen = _QApp.screenAt(btn_pos)
            if screen is None:
                screen = _QApp.primaryScreen()
            screen_rect = screen.availableGeometry()

            x = btn_pos.x()
            y = btn_pos.y() + 4

            # تصحيح لو بيخرج من يمين الشاشة
            if x + dlg.width() > screen_rect.right():
                x = screen_rect.right() - dlg.width() - 4

            # تصحيح لو بيخرج من أسفل الشاشة
            if y + dlg.height() > screen_rect.bottom():
                y = self.btn_calendar.mapToGlobal(self.btn_calendar.rect().topLeft()).y() - dlg.height() - 4

            # تصحيح لو بيخرج من شمال الشاشة
            if x < screen_rect.left():
                x = screen_rect.left() + 4

            dlg.move(x, y)
            dlg.exec_()
        self.btn_calendar.clicked.connect(open_calendar)

        self.lbl_day = QLabel()
        self.lbl_day.setFixedSize(82, 26)   # مربع اليوم أكبر لظهور اسم اليوم بوضوح
        self.lbl_day.setAlignment(Qt.AlignCenter)
        self._day_style_normal  = "color:#1e293b; background:#f8fafc; border:1px solid #e2e8f0; border-radius:6px; padding:3px 8px; font-size: 14px; font-weight:bold;"
        self._day_style_friday  = "color:#dc2626; background:#fef2f2; border:1px solid #fca5a5; border-radius:6px; padding:3px 8px; font-size: 14px; font-weight:bold;"
        self.update_day_label()

        date_block.addWidget(date_lbl, 0, Qt.AlignVCenter)
        date_block.addWidget(self.date_edit, 0, Qt.AlignVCenter)
        date_block.addWidget(arrows_widget, 0, Qt.AlignVCenter)
        date_block.addWidget(self.btn_calendar, 0, Qt.AlignVCenter)
        date_block.addWidget(self.lbl_day, 0, Qt.AlignVCenter)

        # فاصل عمودي
        v_sep = QFrame()
        v_sep.setFrameShape(QFrame.VLine)
        v_sep.setStyleSheet("color: #e2e8f0; background: #e2e8f0; border: none; max-width: 1px;")

        # --- بلوك الوقت ---
        time_block = QHBoxLayout()
        time_block.setSpacing(6)
        time_block.setAlignment(Qt.AlignVCenter)

        time_lbl = QLabel("🕐  الوقت")
        time_lbl.setStyleSheet(LBL_STYLE + f" line-height: {CTRL_H}px; padding: 2px 0 0 0; margin: 0;")
        time_lbl.setFixedWidth(75)   # عرض كافٍ لظهور "الوقت" كاملاً
        time_lbl.setFixedHeight(CTRL_H)
        time_lbl.setAlignment(Qt.AlignVCenter | Qt.AlignLeading)

        self.time_h = QSpinBox()
        self.time_h.setRange(0, 23)
        self.time_h.setMinimumWidth(100)
        self.time_h.setFixedHeight(CTRL_H)
        self.time_h.setValue(datetime.datetime.now().hour)
        self.time_h.setStyleSheet(SPIN_STYLE_MAIN)
        self.time_h.setSuffix("  ساعة")

        colon_lbl = QLabel(":")
        colon_lbl.setStyleSheet("font-size: 14px; font-weight: bold; color: #64748b;")
        colon_lbl.setFixedSize(10, CTRL_H)
        colon_lbl.setAlignment(Qt.AlignCenter)

        self.time_m = QSpinBox()
        self.time_m.setRange(0, 59)
        self.time_m.setMinimumWidth(100)
        self.time_m.setFixedHeight(CTRL_H)
        self.time_m.setValue(datetime.datetime.now().minute)
        self.time_m.setStyleSheet(SPIN_STYLE_MAIN)
        self.time_m.setSuffix("  دقيقة")

        time_block.addWidget(time_lbl, 0, Qt.AlignVCenter)
        time_block.addWidget(self.time_h, 0, Qt.AlignVCenter)
        time_block.addWidget(colon_lbl, 0, Qt.AlignVCenter)
        time_block.addWidget(self.time_m, 0, Qt.AlignVCenter)
        time_block.addStretch()

        dt_row.addLayout(date_block, 0)
        dt_row.addWidget(v_sep, 0, Qt.AlignVCenter)
        dt_row.addLayout(time_block, 0)
        dt_row.addStretch()
        hdr_body_layout.addLayout(dt_row)

        header_outer.addWidget(hdr_body)
        main_layout.addWidget(header_group)

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #dee2e6;
                border-radius: 0 8px 8px 8px;
                background: #f8f9fa;
            }
            QTabBar::tab {
                background: #e9ecef;
                color: #334155;
                padding: 9px 20px;
                font-size: 14px;
                font-weight: 600;
                border: 1px solid #dee2e6;
                border-bottom: none;
                border-radius: 6px 6px 0 0;
                margin-left: 2px;
            }
            QTabBar::tab:selected {
                background: white;
                color: #0d6efd;
                font-weight: bold;
                border-color: #dee2e6;
                border-bottom: 2px solid white;
            }
            QTabBar::tab:hover:!selected { background: #dee2e6; color: #1e293b; }
        """)
        self.disc_widgets = []
        for code, name in DISC_LIST:
            tab = DisciplineTab(code, name, self)
            self.tabs.addTab(tab, name)
            self.disc_widgets.append(tab)
        main_layout.addWidget(self.tabs, 1)  # stretch=1 يخليها تملأ المساحة

        self.btn_run = QPushButton("▶️ بدء توليد كافة الملفات")
        self.btn_run.setFixedHeight(60)
        self.btn_run.setCursor(Qt.PointingHandCursor)
        self.btn_run.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #0d6efd, stop:1 #0a58ca);
                color: white;
                font-size: 18px;
                font-weight: bold;
                border-radius: 10px;
                border: 3px solid #084298;
                padding: 10px;
                text-align: center;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #0b5ed7, stop:1 #0a58ca);
                border: 3px solid #0a58ca;
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #0a58ca, stop:1 #084298);
                padding-top: 12px;
                padding-bottom: 8px;
            }
        """)
        self.btn_run.clicked.connect(self.run_or_stop)
        main_layout.addWidget(self.btn_run)

        main_h_layout.addWidget(main_content_widget, 1)  # يمتد مع النافذة

        # الجزء الأيمن (شريط التقدم)
        # ===== لوحة الحالة - Modern & Clean =====
        progress_widget = QWidget()
        progress_widget.setMaximumWidth(420)
        progress_widget.setMinimumWidth(320)
        progress_widget.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border-left: 1px solid #e9ecef;
            }
        """)
        progress_layout = QVBoxLayout(progress_widget)
        progress_layout.setSpacing(0)
        progress_layout.setContentsMargins(0, 0, 0, 0)

        # ----- Header -----
        header_widget = QWidget()
        header_widget.setStyleSheet("""
            QWidget {
                background-color: #1e293b;
                border: none;
            }
        """)
        header_widget.setFixedHeight(52)
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(16, 0, 16, 0)

        header_icon = QLabel("⚙")
        header_icon.setStyleSheet("color: #64748b; font-size: 16px; background: transparent;")
        header_title = QLabel("حالة التوليد")
        header_title.setStyleSheet("""
            color: #f1f5f9;
            font-size: 14px;
            font-weight: bold;
            background: transparent;
        """)
        header_layout.addStretch()
        header_layout.addWidget(header_title)
        header_layout.addWidget(header_icon)
        header_layout.addStretch()
        progress_layout.addWidget(header_widget)

        # ----- Body (مساحة بيضاء بـ padding) -----
        body_widget = QWidget()
        body_widget.setStyleSheet("background: white; border: none;")
        body_layout = QVBoxLayout(body_widget)
        body_layout.setContentsMargins(14, 14, 14, 10)
        body_layout.setSpacing(10)

        # --- عداد الملفات ---
        counter_card = QWidget()
        counter_card.setStyleSheet("""
            QWidget {
                background-color: #eff6ff;
                border: 1px solid #bfdbfe;
                border-radius: 8px;
            }
        """)
        counter_layout = QHBoxLayout(counter_card)
        counter_layout.setContentsMargins(12, 8, 12, 8)

        counter_icon = QLabel("📁")
        counter_icon.setStyleSheet("font-size: 18px; background: transparent; border: none;")
        counter_icon.setFixedWidth(28)

        self.file_counter_label = QLabel("إجمالي الملفات المتوقعة: 0 ملف")
        self.file_counter_label.setStyleSheet("""
            font-size: 14px;
            font-weight: bold;
            color: #1d4ed8;
            background: transparent;
            border: none;
        """)
        self.file_counter_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        # التوجيه: الآيقونة أقصى اليسار، النص لليمين
        counter_layout.addWidget(counter_icon)
        counter_layout.addStretch()
        counter_layout.addWidget(self.file_counter_label)
        body_layout.addWidget(counter_card)

        # --- معلومات الحالة فوق الـ progress bar ---
        status_layout = QHBoxLayout()
        status_layout.setSpacing(6)

        self.progress_label = QLabel("في انتظار البدء...")
        self.progress_label.setStyleSheet("font-size: 14px; color: #334155;")
        self.progress_label.setAlignment(Qt.AlignRight)

        self.pct_label = QLabel("0%")
        self.pct_label.setStyleSheet("""
            font-size: 13px;
            font-weight: bold;
            color: #3b82f6;
            min-width: 36px;
        """)
        self.pct_label.setAlignment(Qt.AlignLeft)

        status_layout.addWidget(self.pct_label)
        status_layout.addStretch()
        status_layout.addWidget(self.progress_label)
        body_layout.addLayout(status_layout)

        # --- Progress Bar ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(10)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: none;
                border-radius: 5px;
                background-color: #e2e8f0;
            }
            QProgressBar::chunk {
                border-radius: 5px;
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #3b82f6, stop:1 #06b6d4);
            }
        """)
        body_layout.addWidget(self.progress_bar)

        # --- بطاقتا الوقت ---
        time_cards_layout = QHBoxLayout()
        time_cards_layout.setSpacing(8)

        def _make_time_card(icon, title, title_align=Qt.AlignCenter):
            card = QWidget()
            card.setStyleSheet("""
                QWidget {
                    background: #f8fafc;
                    border: 1px solid #e2e8f0;
                    border-radius: 8px;
                }
            """)
            cl = QVBoxLayout(card)
            cl.setContentsMargins(10, 8, 10, 8)
            cl.setSpacing(2)
            top = QLabel(f"{icon}  {title}")
            top.setStyleSheet("font-size: 13px; color: #64748b; background: transparent; border: none;")
            top.setAlignment(title_align)
            val = QLabel("--")
            val.setStyleSheet("font-size: 14px; font-weight: bold; color: #1e293b; background: transparent; border: none;")
            val.setAlignment(Qt.AlignCenter)
            cl.addWidget(top)
            cl.addWidget(val)
            return card, val

        elapsed_card, self.elapsed_time_label   = _make_time_card("⏱", "الوقت المنقضي", Qt.AlignRight)
        remaining_card, self.time_remaining_label = _make_time_card("⌛", "الوقت المتبقي", Qt.AlignLeft)

        # الصورة: الصندوق اليسار = الوقت المنقضي، اليمين = الوقت المتبقي
        time_cards_layout.addWidget(remaining_card)   # يظهر لليمين في RTL
        time_cards_layout.addWidget(elapsed_card)    # يظهر لليسار في RTL
        body_layout.addLayout(time_cards_layout)

        # --- قائمة الملفات ---
        files_header = QLabel("الملفات المنشأة")
        files_header.setStyleSheet("""
            font-size: 13px;
            font-weight: 600;
            color: #475569;
            padding-top: 4px;
        """)
        files_header.setAlignment(Qt.AlignRight)
        files_header_row = QHBoxLayout()
        files_header_row.addWidget(files_header)
        files_header_row.addStretch()
        body_layout.addLayout(files_header_row)

        self.files_list = QListWidget()
        self.files_list.setStyleSheet("""
            QListWidget {
                background-color: #f8fafc;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                font-size: 13px;
                font-family: 'Segoe UI';
                outline: none;
                color: #1e293b;
            }
            QListWidget::item {
                padding: 6px 10px;
                border-bottom: 1px solid #f1f5f9;
                color: #1e293b;
            }
            QListWidget::item:hover { background-color: #eff6ff; }
            QListWidget::item:selected {
                background-color: #dbeafe;
                color: #1e40af;
            }
        """)
        body_layout.addWidget(self.files_list, 1)

        # --- زرار فتح المجلد ---
        self.btn_open_output = QPushButton("📂  فتح مجلد الإخراج")
        self.btn_open_output.setCursor(Qt.PointingHandCursor)
        self.btn_open_output.setFixedHeight(40)
        self.btn_open_output.setStyleSheet("""
            QPushButton {
                background-color: #16a34a;
                color: white;
                font-weight: bold;
                font-size: 13px;
                border-radius: 8px;
                border: none;
            }
            QPushButton:hover   { background-color: #15803d; }
            QPushButton:pressed { background-color: #166534; }
            QPushButton:disabled {
                background-color: #e2e8f0;
                color: #64748b;
            }
        """)
        self.btn_open_output.clicked.connect(self.open_output_folder)
        self.btn_open_output.setEnabled(False)
        body_layout.addWidget(self.btn_open_output)

        progress_layout.addWidget(body_widget, 1)

        main_h_layout.addWidget(progress_widget, 0)  # عرض ثابت بدون stretch

        self.tabs.setCurrentIndex(0)

        # -------------------- استعادة الجلسة السابقة --------------------
        self.load_session()

    def customize_calendar(self):
        # لا شيء - التقويم أصبح منفصلاً
        pass

    def customize_calendar_widget(self, cal):
        cal.setLayoutDirection(Qt.RightToLeft)
        cal.setLocale(QLocale(QLocale.Arabic, QLocale.Egypt))
        cal.setFirstDayOfWeek(Qt.Saturday)

        red_f = QTextCharFormat()
        red_f.setForeground(QColor("#dc2626"))
        red_f.setFontWeight(700)

        norm_f = QTextCharFormat()
        norm_f.setForeground(QColor("#1e293b"))

        hdr_f = QTextCharFormat()
        hdr_f.setForeground(QColor("#475569"))
        hdr_f.setFontWeight(600)

        # تلوين أيام الأسبوع
        cal.setWeekdayTextFormat(Qt.Friday,    red_f)
        cal.setWeekdayTextFormat(Qt.Saturday,  norm_f)
        cal.setWeekdayTextFormat(Qt.Sunday,    norm_f)
        cal.setWeekdayTextFormat(Qt.Monday,    norm_f)
        cal.setWeekdayTextFormat(Qt.Tuesday,   norm_f)
        cal.setWeekdayTextFormat(Qt.Wednesday, norm_f)
        cal.setWeekdayTextFormat(Qt.Thursday,  norm_f)

        # تنسيق اليوم الحالي
        today_f = QTextCharFormat()
        today_f.setFontWeight(700)
        today_f.setForeground(QColor("#0d6efd"))
        cal.setDateTextFormat(QDate.currentDate(), today_f)

    def update_day_label(self):
        qdate   = self.date_edit.date()
        py_date = datetime.date(qdate.year(), qdate.month(), qdate.day())
        day_ar  = ARABIC_DAYS.get(py_date.strftime("%A"), "")
        self.lbl_day.setText(day_ar)
        is_friday = py_date.weekday() == 4
        style = getattr(self, '_day_style_friday' if is_friday else '_day_style_normal',
                        "font-size: 13px; font-weight:bold;")
        self.lbl_day.setStyleSheet(style)

    def calculate_expected_files(self):
        """حساب عدد الملفات مع debounce لتجنب الحسابات المتكررة"""
        if not hasattr(self, '_calc_timer'):
            from PyQt5.QtCore import QTimer
            self._calc_timer = QTimer()
            self._calc_timer.setSingleShot(True)
            self._calc_timer.timeout.connect(self._do_calculate)
        self._calc_timer.start(100)  # تأخير 100ms

    def _do_calculate(self):
        try:
            raw_plots = self.plots_input.text()
            global_plots = [p.strip() for p in re.split(r'[,\-\s.*]+', raw_plots) if p.strip().isdigit()]
            total = 0
            for tab in self.disc_widgets:
                for row in tab.rows:
                    if not row['desc'].text().strip():
                        continue
                    rev = row['revision'].value()
                    if rev > 0:
                        if row['manual_plot_input'].text().strip():
                            total += 1
                    else:
                        if row['radio_all'].isChecked():
                            total += len(global_plots)
                        else:
                            sp = row['plots_input'].text()
                            pl = [p.strip() for p in re.split(r'[,\-\s.*]+', sp) if p.strip().isdigit()]
                            total += len(pl)
            self.file_counter_label.setText(f"إجمالي الملفات المتوقعة: {total} ملف")
        except:
            self.file_counter_label.setText("إجمالي الملفات المتوقعة: -- ملف")

    def run_or_stop(self):
        if self.process_thread and self.process_thread.isRunning():
            reply = QMessageBox.warning(self, "تأكيد الإيقاف",
                                        "هل تريد إيقاف التوليد؟\nسيتم حذف الملفات المنشأة في هذه العملية فقط.",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.process_thread.stop()
                self.btn_run.setEnabled(False)
                self.btn_run.setText("⏳ جاري الإيقاف...")
                self.btn_run.setStyleSheet("""
                    QPushButton {
                        background-color: #6c757d;
                        color: white;
                        font-size: 18px;
                        font-weight: bold;
                        border-radius: 10px;
                        border: 3px solid #6c757d;
                        padding: 10px;
                    }
                """)
        else:
            self.run_process()

    def run_process(self):
        qdate = self.date_edit.date()
        py_date = datetime.date(qdate.year(), qdate.month(), qdate.day())
        if py_date.weekday() == 4:
            reply = QMessageBox.warning(self, "⚠️ تحذير: يوم الجمعة", 
                                        "⚠️ تنبيه: التاريخ المحدد هو يوم الجمعة (عطلة رسمية)\n\n"
                                        "هل تريد الاستمرار في توليد الملفات بهذا التاريخ؟",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                return

        raw_plots = self.plots_input.text()
        plots = [p.strip() for p in re.split(r'[,\-\s.*]+', raw_plots) if p.strip().isdigit()]

        tabs_data = []
        # متغيرات للتحقق
        has_any_auto = False          # وجود أي طلب أوتوماتيكي (Rev 00) يعتمد على الحقل العام
        has_auto_without_plots = False # وجود طلب أوتوماتيكي يعتمد على الحقل العام لكن الحقل العام فارغ
        has_specific_without_plots = False  # وجود طلب "قطع محددة" بدون أرقام مكتوبة

        for tab in self.disc_widgets:
            try:
                serial = int(tab.serial_input.text())
            except:
                continue
            rows_data = []
            for row in tab.rows:
                desc = row['desc'].text().strip()
                if not desc:
                    continue
                rev = row['revision'].value()
                if rev > 0:  # طلب يدوي
                    man_ref = row['manual_ref_input'].text().strip()
                    man_plot = row['manual_plot_input'].text().strip()
                    if not man_ref or not man_plot:
                        continue
                    try:
                        full_ref = f"{PROJECT_PREFIX}-{tab.code}-{int(man_ref):03d}"
                    except:
                        continue
                    rows_data.append({
                        'desc': desc,
                        'attach_paths': list(row['attach_paths']),
                        'revision': rev,
                        'manual_mode': True,
                        'manual_ref': full_ref,
                        'manual_plot': man_plot,
                        'plots': None,
                        'suffix': row['suffix'].text().strip()
                    })
                else:  # طلب أوتوماتيكي
                    if row['radio_all'].isChecked():
                        # يعتمد على الحقل العام
                        has_any_auto = True
                        if not plots:
                            has_auto_without_plots = True
                            continue
                        rows_data.append({
                            'desc': desc,
                            'attach_paths': list(row['attach_paths']),
                            'revision': 0,
                            'manual_mode': False,
                            'plots': None,
                            'suffix': row['suffix'].text().strip()
                        })
                    else:
                        # يعتمد على حقل القطع المحددة الخاص به
                        sp = row['plots_input'].text()
                        pl = [p.strip() for p in re.split(r'[,\-\s.*]+', sp) if p.strip().isdigit()]
                        if not pl:
                            has_specific_without_plots = True
                            continue
                        rows_data.append({
                            'desc': desc,
                            'attach_paths': list(row['attach_paths']),
                            'revision': 0,
                            'manual_mode': False,
                            'plots': pl,
                            'suffix': row['suffix'].text().strip()
                        })
            if rows_data:
                tabs_data.append({
                    'code': tab.code,
                    'name': tab.name,
                    'serial': serial,
                    'rows': rows_data
                })

        if not tabs_data:
            if has_specific_without_plots:
                QMessageBox.warning(self, "تنبيه", "لم يتم إدخال أرقام القطع المحددة")
            elif has_auto_without_plots:
                QMessageBox.warning(self, "تنبيه", "لم يتم إدخال أرقام القطع لطلبات الاستلام")
            else:
                QMessageBox.warning(self, "تنبيه", "لا توجد طلبات صالحة للتوليد")
            return

        # إذا كان هناك طلبات أوتوماتيكية تعتمد على الحقل العام ولكن الحقل العام فارغ
        if has_any_auto and not plots and has_auto_without_plots:
            QMessageBox.warning(self, "تنبيه", "برجاء إدخال أرقام القطع في الحقل العام للطلبات التي تستخدم 'كل القطع'")
            return

        data = {
            'date': py_date.strftime("%d/%m/%Y"),
            'time': f"{self.time_h.value():02d}:{self.time_m.value():02d}",
            'plots': plots,
            'tabs': tabs_data
        }

        self.progress_bar.setValue(0)
        self.pct_label.setText("0%")
        self.files_list.clear()
        self.progress_label.setText("جاري التوليد...")
        self.elapsed_time_label.setText("--")
        self.time_remaining_label.setText("جاري الحساب...")
        self.btn_open_output.setEnabled(False)
        self.btn_run.setText("⏹ إيقاف التوليد")
        self.btn_run.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #6c757d, stop:1 #5a6268);
                color: white;
                font-size: 18px;
                font-weight: bold;
                border-radius: 10px;
                border: 3px solid #495057;
                padding: 10px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #5a6268, stop:1 #545b62);
                border: 3px solid #545b62;
            }
        """)

        self.process_thread = ProcessThread(data)
        self.process_thread.progress_update.connect(self.on_progress_update)
        self.process_thread.finished.connect(self.on_process_finished)
        self.process_thread.start()

    def on_progress_update(self, current, total, message, time_remaining):
        if total > 0:
            progress = 99 if current >= total else int(current / total * 100)
            self.progress_bar.setValue(progress)
            self.pct_label.setText(f"{progress}%")
            self.progress_label.setText(f"الملف {current} من {total}")
        if time_remaining:
            self.time_remaining_label.setText(time_remaining)

        if self.process_thread and self.process_thread.start_time:
            elapsed = time.time() - self.process_thread.start_time
            if elapsed < 60:
                elapsed_str = f"{int(elapsed)} ث"
            elif elapsed < 3600:
                elapsed_str = f"{int(elapsed//60)} د {int(elapsed%60)} ث"
            else:
                elapsed_str = f"{int(elapsed//3600)} س {int((elapsed%3600)//60)} د"
            self.elapsed_time_label.setText(elapsed_str)

        if message.startswith("تم: "):
            self.files_list.addItem(message[4:])
            self.files_list.scrollToBottom()

    def on_process_finished(self, success, message):
        self.btn_run.setEnabled(True)
        self.btn_run.setText("▶️ بدء توليد كافة الملفات")
        self.btn_run.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #0d6efd, stop:1 #0a58ca);
                color: white;
                font-size: 18px;
                font-weight: bold;
                border-radius: 10px;
                border: 3px solid #084298;
                padding: 10px;
                text-align: center;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #0b5ed7, stop:1 #0a58ca);
                border: 3px solid #0a58ca;
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #0a58ca, stop:1 #084298);
                padding-top: 12px;
                padding-bottom: 8px;
            }
        """)

        if self.process_thread and self.process_thread.start_time:
            elapsed = time.time() - self.process_thread.start_time
            if elapsed < 60:
                elapsed_str = f"{int(elapsed)} ث"
            elif elapsed < 3600:
                elapsed_str = f"{int(elapsed//60)} د {int(elapsed%60)} ث"
            else:
                elapsed_str = f"{int(elapsed//3600)} س {int((elapsed%3600)//60)} د"
            self.elapsed_time_label.setText(elapsed_str)

        if success:
            self.progress_bar.setValue(100)
            self.pct_label.setText("100%")
            self.progress_label.setText("اكتمل ✓")
            self.time_remaining_label.setText("0 ث")
            self.btn_open_output.setEnabled(True)
            QMessageBox.information(self, "نجاح", message)
            # تحديث جميع قوائم الاقتراحات
            self.refresh_all_suggestions()
        else:
            QMessageBox.critical(self, "خطأ", message)
            self.progress_label.setText("تم الإيقاف!" if "إيقاف" in message else "فشل!")
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        item = QListWidgetItem(f"[{timestamp}] {'✓' if success else '✗'} {message}")
        item.setForeground(QColor("green") if success else QColor("red"))
        self.files_list.addItem(item)
        self.files_list.scrollToBottom()

    def refresh_all_suggestions(self):
        """تحديث قوائم الاقتراحات في جميع التبويبات"""
        for tab in self.disc_widgets:
            tab.refresh_suggestions()

    def open_output_folder(self):
        output_path = os.path.abspath("Output")
        if os.path.exists(output_path):
            os.startfile(output_path)
        else:
            QMessageBox.warning(self, "تحذير", "مجلد الإخراج غير موجود بعد!")

    # -------------------- حفظ واستعادة الجلسة --------------------
    SESSION_FILE = "session.json"

    def save_session(self):
        """حفظ الجلسة الحالية كاملةً في ملف JSON"""
        try:
            session = {
                'version': 1,
                'plots': self.plots_input.text(),
                'time_h': self.time_h.value(),
                'time_m': self.time_m.value(),
                'active_tab': self.tabs.currentIndex(),
                'disciplines': []
            }

            for tab in self.disc_widgets:
                # حفظ بداية الترقيم
                try:
                    serial_start = int(tab.serial_input.text())
                except:
                    serial_start = 1

                rows_data = []
                for row in tab.rows:
                    row_state = {
                        'desc': row['desc'].text(),
                        'suffix': row['suffix'].text(),
                        'revision': row['revision'].value(),
                        'radio_all': row['radio_all'].isChecked(),
                        'plots_input': row['plots_input'].text(),
                        'manual_ref_input': row['manual_ref_input'].text(),
                        'manual_plot_input': row['manual_plot_input'].text(),
                        'attach_paths': list(row['attach_paths']),
                        # حفظ حالة التوسيع/الطي
                        'expanded': row['expand_btn'].isChecked(),
                    }
                    rows_data.append(row_state)

                session['disciplines'].append({
                    'code': tab.code,
                    'serial_start': serial_start,
                    'rows': rows_data
                })

            with open(self.SESSION_FILE, 'w', encoding='utf-8') as f:
                json.dump(session, f, ensure_ascii=False, indent=2)

        except Exception as e:
            log_error(f"فشل حفظ الجلسة: {str(e)}\n{traceback.format_exc()}")

    def load_session(self):
        """استعادة آخر جلسة محفوظة عند فتح البرنامج"""
        if not os.path.exists(self.SESSION_FILE):
            return

        try:
            with open(self.SESSION_FILE, 'r', encoding='utf-8') as f:
                session = json.load(f)

            if session.get('version') != 1:
                return

            # إيقاف الـ signals أثناء التحميل لتجنب إعادة الرسم المتكررة
            for w in (self.plots_input, self.time_h, self.time_m):
                w.blockSignals(True)

            self.plots_input.setText(session.get('plots', ''))
            self.time_h.setValue(session.get('time_h', datetime.datetime.now().hour))
            self.time_m.setValue(session.get('time_m', datetime.datetime.now().minute))

            for w in (self.plots_input, self.time_h, self.time_m):
                w.blockSignals(False)

            disc_map = {tab.code: tab for tab in self.disc_widgets}

            for disc_data in session.get('disciplines', []):
                code = disc_data.get('code')
                if code not in disc_map:
                    continue
                tab = disc_map[code]

                tab.serial_input.blockSignals(True)
                tab.serial_input.setText(str(disc_data.get('serial_start', 1)))
                tab.serial_input.blockSignals(False)

                saved_rows = disc_data.get('rows', [])
                if not saved_rows:
                    continue

                # إزالة الصف الافتراضي الفارغ
                while tab.rows:
                    container = tab.rows[0]['container']
                    container.setParent(None)
                    tab.rows.pop(0)

                for row_state in saved_rows:
                    tab.add_row()
                    row = tab.rows[-1]

                    # إيقاف signals لكل الحقول دفعة واحدة
                    for wk in ('desc', 'suffix', 'revision', 'manual_ref_input',
                               'manual_plot_input', 'plots_input', 'radio_all', 'expand_btn'):
                        row[wk].blockSignals(True)

                    row['desc'].setText(row_state.get('desc', ''))
                    row['suffix'].setText(row_state.get('suffix', ''))
                    row['revision'].setValue(row_state.get('revision', 0))
                    row['manual_ref_input'].setText(row_state.get('manual_ref_input', ''))
                    row['manual_plot_input'].setText(row_state.get('manual_plot_input', ''))

                    if row_state.get('radio_all', True):
                        row['radio_all'].setChecked(True)
                        row['plots_input'].setEnabled(False)
                        row['plots_input'].clear()
                    else:
                        row['radio_all'].setChecked(False)
                        row['plots_input'].setEnabled(True)
                        row['plots_input'].setText(row_state.get('plots_input', ''))

                    # استعادة signals
                    for wk in ('desc', 'suffix', 'revision', 'manual_ref_input',
                               'manual_plot_input', 'plots_input', 'radio_all', 'expand_btn'):
                        row[wk].blockSignals(False)

                    # استعادة المرفقات مباشرة بدون بناء widgets معقدة
                    from PyQt5.QtCore import QSize as _QS
                    aw = row['attach_list_widget']
                    for att_path in row_state.get('attach_paths', []):
                        if os.path.exists(att_path) and att_path not in row['attach_paths']:
                            # إخفاء placeholder
                            for i in range(aw.count()):
                                it = aw.item(i)
                                if it and it.data(Qt.UserRole) == "__placeholder__":
                                    aw.takeItem(i)
                                    break
                            item = QListWidgetItem()
                            item.setData(Qt.UserRole, att_path)
                            item.setToolTip(att_path)
                            item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsDragEnabled)
                            item.setSizeHint(_QS(0, 32))
                            aw.addItem(item)
                            iw = QWidget()
                            iw.setStyleSheet("background: transparent; border: none;")
                            il = QHBoxLayout(iw)
                            il.setContentsMargins(6, 2, 6, 2)
                            lbl = QLabel(f"📄  {os.path.basename(att_path)}")
                            lbl.setStyleSheet("font-size: 12px; color: #1e293b; background: transparent;")
                            lbl.setToolTip(att_path)
                            il.addWidget(lbl, 1)
                            aw.setItemWidget(item, iw)
                            row['attach_paths'].append(att_path)

                    # استعادة حالة التوسيع/الطي
                    expanded = row_state.get('expanded', True)
                    row['details'].setVisible(expanded)
                    row['expand_btn'].setChecked(expanded)
                    row['expand_btn'].setText("▼" if expanded else "▶")

            active_tab = session.get('active_tab', 0)
            if 0 <= active_tab < self.tabs.count():
                self.tabs.setCurrentIndex(active_tab)

            self.calculate_expected_files()

        except Exception as e:
            log_error(f"فشل استعادة الجلسة: {str(e)}\n{traceback.format_exc()}")

    def closeEvent(self, event):
        """حفظ الجلسة عند إغلاق البرنامج"""
        # إذا كان التوليد شغال نسأل المستخدم
        if self.process_thread and self.process_thread.isRunning():
            reply = QMessageBox.question(
                self, "تأكيد الإغلاق",
                "التوليد لا يزال جارياً!\n\nهل تريد إيقافه وإغلاق البرنامج؟",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No
            )
            if reply == QMessageBox.No:
                event.ignore()
                return
            self.process_thread.stop()
            self.process_thread.wait(3000)

        self.save_session()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Segoe UI", 11))
    app.setStyleSheet("""
        QToolTip {
            background-color: #2b2b2b;
            color: #ffffff;
            border: 1px solid #555;
            border-radius: 4px;
            padding: 5px 8px;
            font-size: 13px;
        }
    """)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
