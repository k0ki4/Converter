import os
import sys
import tempfile
from pathlib import Path
from PyQt6 import uic
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt6.QtCore import Qt, QThread, pyqtSignal
import traceback


class ConversionThread(QThread):
    #Поток для выполнения конвертации в фоне
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)  # путь к сохраненному файлу
    error = pyqtSignal(str)

    def __init__(self, input_path, output_format, conversion_func, save_path):
        super().__init__()
        self.input_path = input_path
        self.output_format = output_format
        self.conversion_func = conversion_func
        self.save_path = save_path

    def run(self):
        try:
            self.progress.emit(10)
            # Выполняем конвертацию
            self.conversion_func(self.input_path, self.save_path)
            self.progress.emit(100)
            self.finished.emit(self.save_path)
        except Exception as e:
            self.error.emit(f"Ошибка: {str(e)}\n{traceback.format_exc()}")


# ==================== ФУНКЦИИ КОНВЕРТАЦИИ ====================

def convert_image_pillow(input_path, output_path):
    # Конвертация изображений с помощью Pillow
    from PIL import Image

    img = Image.open(input_path)
    output_ext = output_path.split('.')[-1].lower()

    # Обработка прозрачности для JPEG/BMP
    if output_ext in ['jpg', 'jpeg', 'bmp'] and img.mode in ('RGBA', 'LA', 'P'):
        background = Image.new('RGB', img.size, (255, 255, 255))
        if img.mode == 'P':
            img = img.convert('RGBA')
        background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
        img = background

    # ОСОБАЯ ОБРАБОТКА ДЛЯ ICO
    elif output_ext == 'ico':
        # ICO требует специальной обработки
        # Конвертируем в нужный формат (RGBA для поддержки прозрачности)
        if img.mode != 'RGBA':
            img = img.convert('RGBA')

        # ICO файлы обычно содержат несколько размеров
        # Создаем список размеров для иконки (стандартные размеры иконок Windows)
        sizes = [(16, 16), (32, 32), (48, 48), (64, 64)]

        # Создаем список изображений разных размеров
        icon_images = []
        for size in sizes:
            # Масштабируем изображение с антиалиасингом
            resized_img = img.resize(size, Image.Resampling.LANCZOS)
            icon_images.append(resized_img)

        # Сохраняем как ICO с несколькими размерами
        icon_images[0].save(output_path, format='ICO', sizes=[(img.width, img.height) for img in icon_images])
        return

    # Для GIF с анимацией
    elif output_ext == 'gif' and hasattr(img, 'is_animated') and img.is_animated:
        frames = []
        try:
            while True:
                frames.append(img.copy())
                img.seek(len(frames))
        except EOFError:
            pass
        if frames:
            frames[0].save(output_path, save_all=True, append_images=frames[1:])
            return

    # Для всех остальных форматов
    img.save(output_path)

def convert_audio_pydub(input_path, output_path):
    # Конвертация аудио с помощью pydub
    from pydub import AudioSegment

    input_ext = input_path.split('.')[-1].lower()
    output_ext = output_path.split('.')[-1].lower()

    try:
        audio = AudioSegment.from_file(input_path, format=input_ext)
    except:
        audio = AudioSegment.from_file(input_path)

    audio.export(output_path, format=output_ext)


def convert_video_moviepy(input_path, output_path):
    # Конвертация видео с помощью moviepy
    from moviepy import VideoFileClip

    clip = VideoFileClip(input_path)
    output_ext = output_path.split('.')[-1].lower()

    if output_ext == 'gif':
        clip.write_gif(output_path)
    else:
        clip.write_videofile(output_path)

    clip.close()


def extract_audio_from_video_moviepy(video_path, audio_path):
    # Извлечение аудио из видео
    from moviepy import VideoFileClip

    clip = VideoFileClip(video_path)
    audio = clip.audio
    audio.write_audiofile(audio_path)
    clip.close()


def convert_document_pypandoc(input_path, output_path):
    # Конвертация документов с помощью pypandoc
    import pypandoc

    output_ext = output_path.split('.')[-1].lower()
    format_map = {
        'docx': 'docx', 'pdf': 'pdf', 'html': 'html', 'txt': 'plain',
        'md': 'markdown', 'epub': 'epub', 'rtf': 'rtf', 'odt': 'odt'
    }

    output_format = format_map.get(output_ext, output_ext)
    pypandoc.convert_file(input_path, output_format, outputfile=output_path)


def convert_table_pandas(input_path, output_path):
    # Конвертация таблиц с помощью pandas
    import pandas as pd

    input_ext = input_path.split('.')[-1].lower()
    output_ext = output_path.split('.')[-1].lower()

    # Чтение
    if input_ext == 'csv':
        df = pd.read_csv(input_path)
    elif input_ext in ['xlsx', 'xls']:
        df = pd.read_excel(input_path)
    elif input_ext == 'json':
        df = pd.read_json(input_path)
    elif input_ext == 'xml':
        df = pd.read_xml(input_path)
    else:
        raise ValueError(f"Неподдерживаемый формат: {input_ext}")

    # Запись
    if output_ext == 'csv':
        df.to_csv(output_path, index=False)
    elif output_ext in ['xlsx', 'xls']:
        df.to_excel(output_path, index=False)
    elif output_ext == 'json':
        df.to_json(output_path, indent=2)
    elif output_ext == 'html':
        df.to_html(output_path, index=False)
    else:
        raise ValueError(f"Неподдерживаемый выходной формат: {output_ext}")


def convert_pdf_to_image_pdf2image(pdf_path, output_path):
    # Конвертация PDF в изображение
    from pdf2image import convert_from_path

    images = convert_from_path(pdf_path, dpi=150, first_page=1, last_page=1)
    if images:
        images[0].save(output_path)


def convert_docx_to_image(docx_path, output_path):
    # Конвертация DOCX в изображение через временный PDF
    import pypandoc
    from pdf2image import convert_from_path

    # Создаем временный PDF
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
        temp_pdf = tmp.name

    try:
        # DOCX → PDF
        pypandoc.convert_file(docx_path, 'pdf', outputfile=temp_pdf)
        # PDF → изображение
        images = convert_from_path(temp_pdf, dpi=150, first_page=1, last_page=1)
        if images:
            images[0].save(output_path)
    finally:
        if os.path.exists(temp_pdf):
            os.remove(temp_pdf)


# ==================== ОСНОВНОЙ КЛАСС ====================

class FileConverter(QMainWindow):
    def __init__(self):
        super().__init__()

        # Загружаем интерфейс
        uic.loadUi('uic_storage/main2.ui', self)
        self.setWindowIcon(QIcon("ico/icon.ico"))

        self.setup_tooltips()
        self.ex_format = None
        self.upload_file = False
        self.current_file_path = None
        self.conversion_thread = None

        # Матрица конвертаций
        self.matrix = {
            # ========== ИЗОБРАЖЕНИЯ ==========
            'JPG': ['PNG', 'GIF', 'BMP', 'TIFF', 'PDF', 'WEBP', 'ICO', 'PPM'],
            'JPEG': ['PNG', 'GIF', 'BMP', 'TIFF', 'PDF', 'WEBP', 'ICO', 'PPM'],
            'PNG': ['JPG', 'GIF', 'BMP', 'TIFF', 'PDF', 'WEBP', 'ICO', 'PPM'],
            'GIF': ['JPG', 'PNG', 'BMP', 'TIFF', 'PDF', 'WEBP', 'ICO', 'PPM'],
            'BMP': ['JPG', 'PNG', 'GIF', 'TIFF', 'PDF', 'WEBP', 'ICO', 'PPM'],
            'TIFF': ['JPG', 'PNG', 'GIF', 'BMP', 'PDF', 'WEBP', 'ICO', 'PPM'],
            'WEBP': ['JPG', 'PNG', 'GIF', 'BMP', 'TIFF', 'PDF', 'ICO', 'PPM'],
            'ICO': ['JPG', 'PNG', 'GIF', 'BMP', 'TIFF', 'PDF', 'WEBP', 'PPM'],
            'PPM': ['JPG', 'PNG', 'GIF', 'BMP', 'TIFF', 'PDF', 'WEBP', 'ICO'],

            # ========== ВИДЕО ==========
            'MP4': ['AVI', 'GIF', 'MKV', 'WEBM', 'MOV', 'WMV', 'MP3', 'WAV'],
            'AVI': ['MP4', 'GIF', 'MKV', 'WEBM', 'MOV', 'WMV', 'MP3', 'WAV'],
            'MKV': ['MP4', 'AVI', 'GIF', 'WEBM', 'MOV', 'WMV', 'MP3', 'WAV'],
            'WEBM': ['MP4', 'AVI', 'GIF', 'MKV', 'MOV', 'WMV', 'MP3', 'WAV'],
            'MOV': ['MP4', 'AVI', 'GIF', 'MKV', 'WEBM', 'WMV', 'MP3', 'WAV'],
            'WMV': ['MP4', 'AVI', 'GIF', 'MKV', 'WEBM', 'MOV', 'MP3', 'WAV'],
            'FLV': ['MP4', 'AVI', 'GIF', 'MKV', 'WEBM', 'MOV', 'WMV', 'MP3', 'WAV'],
            'M4V': ['MP4', 'AVI', 'GIF', 'MKV', 'WEBM', 'MOV', 'WMV', 'MP3', 'WAV'],
            '3GP': ['MP4', 'AVI', 'GIF', 'MKV', 'WEBM', 'MOV', 'WMV', 'MP3', 'WAV'],

            # ========== АУДИО ==========
            'MP3': ['WAV', 'FLAC', 'OGG', 'AAC', 'AIFF', 'M4A'],
            'WAV': ['MP3', 'FLAC', 'OGG', 'AAC', 'AIFF', 'M4A'],
            'FLAC': ['MP3', 'WAV', 'OGG', 'AAC', 'AIFF', 'M4A'],
            'OGG': ['MP3', 'WAV', 'FLAC', 'AAC', 'AIFF', 'M4A'],
            'AAC': ['MP3', 'WAV', 'FLAC', 'OGG', 'AIFF', 'M4A'],
            'AIFF': ['MP3', 'WAV', 'FLAC', 'OGG', 'AAC', 'M4A'],
            'M4A': ['MP3', 'WAV', 'FLAC', 'OGG', 'AAC', 'AIFF'],
            'WMA': ['MP3', 'WAV', 'FLAC', 'OGG', 'AAC', 'AIFF', 'M4A'],
            'AMR': ['MP3', 'WAV', 'FLAC', 'OGG', 'AAC', 'AIFF', 'M4A'],

            # ========== ДОКУМЕНТЫ ==========
            'DOCX': ['PDF', 'HTML', 'TXT', 'MD', 'RTF', 'EPUB'],
            'DOC': ['PDF', 'HTML', 'TXT', 'MD', 'RTF', 'DOCX', 'EPUB'],
            'PDF': ['DOCX', 'HTML', 'TXT', 'MD', 'RTF', 'JPG', 'PNG', 'EPUB'],
            'HTML': ['PDF', 'DOCX', 'TXT', 'MD', 'RTF', 'EPUB'],
            'HTM': ['PDF', 'DOCX', 'TXT', 'MD', 'RTF', 'HTML', 'EPUB'],
            'TXT': ['PDF', 'DOCX', 'HTML', 'MD', 'RTF', 'EPUB'],
            'MD': ['PDF', 'DOCX', 'HTML', 'TXT', 'RTF', 'EPUB'],
            'MARKDOWN': ['PDF', 'DOCX', 'HTML', 'TXT', 'RTF', 'EPUB'],
            'RTF': ['PDF', 'DOCX', 'HTML', 'TXT', 'MD', 'EPUB'],
            'EPUB': ['PDF', 'DOCX', 'HTML', 'TXT', 'MD', 'RTF'],
            'ODT': ['PDF', 'DOCX', 'HTML', 'TXT', 'MD', 'RTF', 'EPUB'],
            'LATEX': ['PDF', 'TXT', 'HTML'],
            'TEX': ['PDF', 'TXT', 'HTML'],

            # ========== ТАБЛИЦЫ ==========
            'CSV': ['XLSX', 'JSON', 'HTML', 'XML', 'PDF', 'PARQUET'],
            'XLSX': ['CSV', 'JSON', 'HTML', 'XML', 'PDF', 'PARQUET'],
            'XLS': ['CSV', 'XLSX', 'JSON', 'HTML', 'XML', 'PDF', 'PARQUET'],
            'JSON': ['CSV', 'XLSX', 'HTML', 'XML', 'PDF', 'PARQUET'],
            'XML': ['CSV', 'XLSX', 'JSON', 'HTML', 'PDF', 'PARQUET'],
            'PARQUET': ['CSV', 'XLSX', 'JSON', 'HTML', 'XML', 'PDF'],
            'ODS': ['CSV', 'XLSX', 'JSON', 'HTML', 'XML', 'PDF', 'PARQUET'],

            # ========== ПРЕЗЕНТАЦИИ ==========
            'PPTX': ['PDF', 'HTML', 'JPG', 'PNG'],
            'PPT': ['PDF', 'HTML', 'JPG', 'PNG', 'PPTX'],
            'ODP': ['PDF', 'HTML', 'JPG', 'PNG', 'PPTX'],

            # ========== АРХИВЫ ==========
            'ZIP': ['TAR', 'GZ', 'RAR'],
            'TAR': ['ZIP', 'GZ'],
            'GZ': ['ZIP', 'TAR'],
            'RAR': ['ZIP', 'TAR'],
            '7Z': ['ZIP', 'TAR'],
        }

        self.special_conversions = {
            'MP4': ['MP3', 'WAV'],
            'AVI': ['MP3', 'WAV'],
            'MKV': ['MP3', 'WAV'],
            'PDF': ['JPG', 'PNG', 'TIFF'],
            'DOCX': ['JPG', 'PNG'],
            'PPTX': ['JPG', 'PNG'],
        }

        # Запуск функции
        self.initUI()

    def initUI(self):
        self.download_bar.setValue(0)
        self.rd_show_off_ex.toggled.connect(self.show_or_off_all_formats)
        self.upload_button.clicked.connect(self.open_file_explorer_advanced)
        self.dowlnload_button.clicked.connect(self.start_conversion_process)

        if not self.upload_file:
            self.setTextSelectedFileExtension(["Выберите файл"])

    def setup_tooltips(self):
        tooltip_text = """
        <b>Что значит "Показывать недоступные"?</b><br><br>
        При включенной опции в списке форматов будут<br>
        отображаться все возможные форматы, включая те,<br>
        которые недоступны для текущего файла.<br><br>
        При выключенной опции будут показаны только<br>
        те форматы, в которые можно конвертировать<br>
        выбранный файл.
        """
        self.what_show_button.setEnabled(False)
        self.what_show_button.setToolTip(tooltip_text)

    def open_file_explorer_advanced(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл",
            "",
            "Все файлы (*.*)",
        )

        if file_path:
            self.upload_file = True
            self.current_file_path = file_path
            self.rd_show_off_ex.setChecked(False)
            self.process_selected_file(file_path)
        else:
            print('Файл не выбран')

    def process_selected_file(self, file_path):
        file_name = os.path.basename(file_path).split('.')[0]
        file_extension = os.path.basename(file_path).split('.')[-1].upper()

        print(f"Выбран файл: {file_name}")
        print(f"Расширение: {file_extension}")

        self.update_ui_after_file_selection(file_name, file_extension)

    def update_ui_after_file_selection(self, file_name, file_extension):
        self.upload_button.setText(f"{file_name}.{file_extension}")
        need_format_list = self.get_output_formats(file_extension)
        if need_format_list:
            self.setTextSelectedFileExtension(need_format_list, file_extension.upper())
            self.ex_format = file_extension.upper()
        else:
            self.setTextSelectedFileExtension(['Неизвестный формат'])
            self.ex_format = None
            self.rd_show_off_ex.setChecked(False)

    def setTextSelectedFileExtension(self, args: list, ext=None):
        self.selected_file_extension.clear()
        for i in args:
            if not ext:
                self.selected_file_extension.addItem(i)
            else:
                if i != ext:
                    self.selected_file_extension.addItem(i)

    def show_or_off_all_formats(self, checked):
        if checked:
            if self.ex_format:
                buff = []
                for i in self.matrix:
                    for j in self.matrix[i]:
                        if j not in buff:
                            buff.append(j)
                self.setTextSelectedFileExtension(buff)
        else:
            if self.ex_format is not None:
                sort_list = self.get_output_formats(self.ex_format)
                self.setTextSelectedFileExtension(sort_list, self.ex_format)

    def get_output_formats(self, input_format):
        input_format = input_format.upper()
        base_formats = self.matrix.get(input_format, [])
        special_formats = self.special_conversions.get(input_format, [])
        all_formats = list(set(base_formats + special_formats))
        return sorted(all_formats)

    def get_converter_type(self, input_format, output_format):
        input_format = input_format.upper()
        output_format = output_format.upper()

        image_formats = ['JPG', 'JPEG', 'PNG', 'GIF', 'BMP', 'TIFF', 'WEBP', 'ICO', 'PPM']
        video_formats = ['MP4', 'AVI', 'MKV', 'WEBM', 'MOV', 'WMV', 'FLV', 'M4V', '3GP']
        audio_formats = ['MP3', 'WAV', 'FLAC', 'OGG', 'AAC', 'AIFF', 'M4A', 'WMA', 'AMR']
        doc_formats = ['DOCX', 'DOC', 'PDF', 'HTML', 'HTM', 'TXT', 'MD', 'MARKDOWN', 'RTF', 'EPUB', 'ODT', 'LATEX',
                       'TEX']
        table_formats = ['CSV', 'XLSX', 'XLS', 'JSON', 'XML', 'PARQUET', 'ODS']
        presentation_formats = ['PPTX', 'PPT', 'ODP']

        if input_format in image_formats and output_format in image_formats:
            return 'image'
        if input_format in audio_formats and output_format in audio_formats:
            return 'audio'
        if input_format in video_formats and output_format in video_formats:
            return 'video'
        if input_format in video_formats and output_format in audio_formats:
            return 'video_to_audio'
        if input_format in doc_formats and output_format in doc_formats:
            return 'document'
        if input_format in table_formats and output_format in table_formats:
            return 'table'
        if input_format in presentation_formats and output_format in presentation_formats:
            return 'presentation'
        if (input_format in doc_formats + presentation_formats and output_format in image_formats):
            return 'doc_to_image'
        return 'unknown'

    # ==================== НОВАЯ ФУНКЦИЯ КОНВЕРТАЦИИ ====================

    def start_conversion_process(self):
        # Запускает процесс конвертации
        if not self.current_file_path:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите файл!")
            return

        if not self.ex_format:
            QMessageBox.warning(self, "Ошибка", "Не удалось определить формат файла!")
            return

        output_format = self.selected_file_extension.currentText()
        if not output_format or output_format in ["Выберите файл", "Неизвестный формат"]:
            QMessageBox.warning(self, "Ошибка", "Выберите формат для конвертации!")
            return

        # Определяем тип конвертации
        conv_type = self.get_converter_type(self.ex_format, output_format)
        if conv_type == 'unknown':
            QMessageBox.warning(self, "Ошибка",
                                f"Конвертация из .{self.ex_format} в .{output_format} не поддерживается!")
            return

        # Определяем функцию конвертации
        conversion_func = self.select_conversion_function(conv_type)
        if not conversion_func:
            QMessageBox.warning(self, "Ошибка",
                                f"Для конвертации типа '{conv_type}' нет функции!")
            return

        # Спрашиваем куда сохранить файл
        base_name = Path(self.current_file_path).stem
        default_name = f"{base_name}_converted.{output_format.lower()}"

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить файл как",
            os.path.join(os.path.dirname(self.current_file_path), default_name),
            f"{output_format} файлы (*.{output_format.lower()})"
        )

        if not save_path:  # Пользователь отменил
            return

        # Блокируем кнопки на время конвертации
        self.set_buttons_enabled(False)

        # Запускаем конвертацию в отдельном потоке
        self.conversion_thread = ConversionThread(
            self.current_file_path,
            output_format,
            conversion_func,
            save_path
        )

        self.conversion_thread.progress.connect(self.update_progress_bar)
        self.conversion_thread.finished.connect(self.on_conversion_finished)
        self.conversion_thread.error.connect(self.on_conversion_error)

        self.conversion_thread.start()

    def select_conversion_function(self, conv_type):
        # Выбирает функцию конвертации по типу
        func_map = {
            'image': convert_image_pillow,
            'audio': convert_audio_pydub,
            'video': convert_video_moviepy,
            'video_to_audio': extract_audio_from_video_moviepy,
            'document': convert_document_pypandoc,
            'table': convert_table_pandas,
            'doc_to_image': self.select_doc_to_image_function,
        }

        func = func_map.get(conv_type)
        if conv_type == 'doc_to_image':
            return self.select_doc_to_image_function
        return func

    def select_doc_to_image_function(self, input_path, output_path):
        # Выбирает функцию конвертации документа в изображение
        input_ext = Path(input_path).suffix.lower()

        if input_ext == '.pdf':
            convert_pdf_to_image_pdf2image(input_path, output_path)
        elif input_ext == '.docx':
            convert_docx_to_image(input_path, output_path)
        else:
            raise ValueError(f"Конвертация {input_ext} в изображение не поддерживается")

    def update_progress_bar(self, value):
        # Обновляет прогресс-бар
        self.download_bar.setValue(value)

    def on_conversion_finished(self, output_path):
        # Обработка успешного завершения конвертации
        self.set_buttons_enabled(True)
        self.download_bar.setValue(100)

        # Спрашиваем, открыть ли папку с файлом
        reply = QMessageBox.question(
            self,
            "Конвертация завершена",
            f"Файл успешно сохранен:\n{output_path}\n\nОткрыть папку с файлом?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.open_file_in_explorer(output_path)

        # Сбрасываем прогресс-бар через 2 секунды
        QApplication.processEvents()
        import time
        time.sleep(2)
        self.download_bar.setValue(0)

    def on_conversion_error(self, error_message):
        # Обработка ошибки конвертации
        self.set_buttons_enabled(True)
        self.download_bar.setValue(0)

        QMessageBox.critical(self, "Ошибка конвертации",
                             f"Произошла ошибка:\n\n{error_message[:500]}...")

    def set_buttons_enabled(self, enabled):
        # Включает/выключает кнопки
        self.upload_button.setEnabled(enabled)
        self.dowlnload_button.setEnabled(enabled)
        self.selected_file_extension.setEnabled(enabled)
        self.rd_show_off_ex.setEnabled(enabled)

    def open_file_in_explorer(self, file_path):
        # Открывает папку с файлом в проводнике
        folder_path = os.path.dirname(file_path)

        if sys.platform == "win32":
            os.startfile(folder_path)
        elif sys.platform == "darwin":
            import subprocess
            subprocess.run(["open", folder_path])
        else:
            import subprocess
            subprocess.run(["xdg-open", folder_path])

    def closeEvent(self, event):
        # Обработка закрытия окна
        if self.conversion_thread and self.conversion_thread.isRunning():
            reply = QMessageBox.question(
                self,
                "Конвертация выполняется",
                "Конвертация все еще выполняется. Закрыть программу?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.conversion_thread.terminate()
                self.conversion_thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def except_hook(cls, exception, traceback):
    sys.__excepthook__(cls, exception, traceback)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = FileConverter()
    window.show()
    sys.excepthook = except_hook
    sys.exit(app.exec())