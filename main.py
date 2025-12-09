import os
import sys
from PyQt6 import uic
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog


class FileConverter(QMainWindow):
    def __init__(self):
        super().__init__()

        # Загружаем интерфейс из .ui файла
        uic.loadUi('uic_storage/main2.ui', self)

        # Добавляем всплывающую подсказку для кнопки "?"
        self.setup_tooltips()

        # Переменнаы статуса
        self.ex_format = None

        self.upload_file = False
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

        # Специальные кейсы
        self.special_conversions = {
            # Видео в аудио (извлечение звука)
            'MP4': ['MP3', 'WAV'],
            'AVI': ['MP3', 'WAV'],
            'MKV': ['MP3', 'WAV'],
            # Документы в изображения (рендер страниц)
            'PDF': ['JPG', 'PNG', 'TIFF'],
            'DOCX': ['JPG', 'PNG'],  # через print to PDF -> image
            # Презентации в изображения
            'PPTX': ['JPG', 'PNG'],
        }
        # Запуск функции
        self.initUI()

    # Инициализация Классов PyQT
    def initUI(self):
        self.download_bar.setValue(0)

        self.rd_show_off_ex.toggled.connect(self.show_or_off_all_formats)

        self.upload_button.clicked.connect(self.open_file_explorer_advanced)
        if not self.upload_file:
            self.setTextSelectedFileExtension(["Выберите файл"])

    def setup_tooltips(self):
        # Подсказка для кнопки "?" (what_show_button)
        tooltip_text = """
        <b>Что значит "Показывать недоступные"?</b><br><br>
        При включенной опции в списке форматов будут<br>
        отображаться все возможные форматы, включая те,<br>
        которые недоступны для текущего файла.<br><br>
        При выключенной опции будут показаны только<br>
        те форматы, в которые можно конвертировать<br>
        выбранный файл.
        """

        # Установка кнопки в неактивное состояние
        self.what_show_button.setEnabled(False)
        # Применение текста для подсказки
        self.what_show_button.setToolTip(tooltip_text)


    def open_file_explorer_advanced(self):
        #Проводник с дополнительными опциями
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл",
            "",
            "Все файлы (*.*)",
        )

        if file_path:
            self.upload_file = True
            self.rd_show_off_ex.setChecked(False)
            self.process_selected_file(file_path)
        else:
            print('Файл не выбран')

    def process_selected_file(self, file_path):
        # Обработка выбранного файла
        # Получаем информацию о файле
        file_name = os.path.basename(file_path).split('.')[0]
        file_extension = os.path.basename(file_path).split('.')[1]

        print(f"Выбран файл: {file_name}")
        print(f"Расширение: {file_extension}")
        print(f"Полный путь: {file_path}")

        # Обновляем интерфейс
        self.update_ui_after_file_selection(file_name, file_extension)

    def update_ui_after_file_selection(self, file_name, file_extension):
        self.upload_button.setText(f"{file_name}.{file_extension}")
        need_format_list = self.get_output_formats(file_extension)
        if need_format_list:
            self.setTextSelectedFileExtension(need_format_list, file_extension.upper())
            self.ex_format = file_extension
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
        # Получение доступных форматов длч конвертации

        input_format = input_format.upper()

        # Базовые конвертации
        base_formats = self.matrix.get(input_format, [])

        # Добавляем специальные конвертации
        special_formats = self.special_conversions.get(input_format, [])

        # Объединяем и убираем дубликаты
        all_formats = list(set(base_formats + special_formats))

        return sorted(all_formats)

    def get_converter_type(self, input_format, output_format):
        # Определить тип конвертера для форматов
        input_format = input_format.upper()
        output_format = output_format.upper()

        # Изображения
        image_formats = ['JPG', 'JPEG', 'PNG', 'GIF', 'BMP', 'TIFF', 'WEBP', 'ICO', 'PPM']
        if input_format in image_formats and output_format in image_formats:
            return 'image'

        # Видео
        video_formats = ['MP4', 'AVI', 'MKV', 'WEBM', 'MOV', 'WMV', 'FLV', 'M4V', '3GP']
        if input_format in video_formats and output_format in video_formats:
            return 'video'

        # Аудио
        audio_formats = ['MP3', 'WAV', 'FLAC', 'OGG', 'AAC', 'AIFF', 'M4A', 'WMA', 'AMR']
        if input_format in audio_formats and output_format in audio_formats:
            return 'audio'

        # Видео в аудио
        if input_format in video_formats and output_format in audio_formats:
            return 'video_to_audio'

        # Документы
        doc_formats = ['DOCX', 'DOC', 'PDF', 'HTML', 'HTM', 'TXT', 'MD', 'MARKDOWN', 'RTF', 'EPUB', 'ODT', 'LATEX',
                       'TEX']
        if input_format in doc_formats and output_format in doc_formats:
            return 'document'

        # Таблицы
        table_formats = ['CSV', 'XLSX', 'XLS', 'JSON', 'XML', 'PARQUET', 'ODS']
        if input_format in table_formats and output_format in table_formats:
            return 'table'

        # Презентации
        presentation_formats = ['PPTX', 'PPT', 'ODP']
        if input_format in presentation_formats and output_format in presentation_formats:
            return 'presentation'

        # Документы/презентации в изображения
        if (input_format in doc_formats + presentation_formats and
                output_format in image_formats):
            return 'doc_to_image'

        return 'unknown'


def expect_hook(cls, expection, traceback):
    sys.__excepthook__(cls, expection, traceback)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = FileConverter()
    window.show()
    sys.excepthook = expect_hook
    sys.exit(app.exec())
