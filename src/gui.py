"""
Графический интерфейс приложения на PySide6.
"""
import sys
import logging
from pathlib import Path
from typing import Optional


def get_config_path() -> Path:
    """
    Вернуть путь к config.json.
    - В dev-режиме (запуск из исходников): рядом с main.py
    - В собранном .app: ~/Library/Application Support/АнализТК/ (macOS)
      или ~/.analiz_tk/ (Linux/Windows)
    """
    import sys, os

    # Если запущен как PyInstaller bundle — sys.frozen = True
    if getattr(sys, 'frozen', False):
        if sys.platform == 'darwin':
            data_dir = Path.home() / 'Library' / 'Application Support' / 'АнализТК'
        elif sys.platform == 'win32':
            data_dir = Path(os.environ.get('APPDATA', Path.home())) / 'АнализТК'
        else:
            data_dir = Path.home() / '.analiz_tk'
        data_dir.mkdir(parents=True, exist_ok=True)
        config_path = data_dir / 'config.json'
        # При первом запуске скопировать example если конфига нет
        if not config_path.exists():
            example = Path(sys._MEIPASS) / 'config.example.json'
            if example.exists():
                import shutil
                shutil.copy(example, config_path)
        return config_path
    else:
        # dev-режим: ищем config.json рядом с корнем проекта
        return Path(__file__).parent.parent / 'config.json'

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QComboBox, QCheckBox, QSpinBox, QDoubleSpinBox,
    QTextEdit, QFileDialog, QTabWidget, QTableWidget, QTableWidgetItem,
    QGroupBox, QMessageBox, QProgressBar, QStatusBar
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont

from src.models import AppConfig, CompetitorConfig
from src.excel_processor import ExcelProcessor

logger = logging.getLogger(__name__)


class ProcessingThread(QThread):
    """Поток для обработки Excel файлов в фоне."""
    progress = Signal(str, bool)  # competitor_name, is_done
    finished = Signal(list)  # results

    def __init__(self, processor: ExcelProcessor):
        super().__init__()
        self.processor = processor

    def run(self):
        results = self.processor.process_all(self.progress.emit)
        self.finished.emit(results)


class MainWindow(QMainWindow):
    """Главное окно приложения."""

    def __init__(self):
        super().__init__()
        self.config_path = get_config_path()
        self.config = AppConfig.load(self.config_path)
        self.processor = ExcelProcessor(self.config)
        self.processing_thread: Optional[ProcessingThread] = None

        self.init_ui()
        self.load_config_to_ui()

        # Настройка логирования
        self.setup_logging()

    def init_ui(self):
        """Инициализация интерфейса."""
        self.setWindowTitle("Анализ цен конкурентов v2.0")
        self.setMinimumSize(1000, 700)

        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Главный layout
        main_layout = QVBoxLayout(central_widget)

        # Вкладки
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # Создать вкладки
        self.create_main_tab()
        self.create_competitor_tab()
        self.create_cities_tab()
        self.create_preview_tab()
        self.create_log_tab()

        # Панель управления внизу
        control_panel = self.create_control_panel()
        main_layout.addLayout(control_panel)

        # Статус бар
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готов к работе")

        # Прогресс бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)

        # Загрузить конфигурацию в UI (после создания всех элементов)
        self.load_config_to_ui()

    def create_main_tab(self):
        """Вкладка с основными настройками."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Группа файлов
        files_group = QGroupBox("Файлы")
        files_layout = QVBoxLayout(files_group)

        # Шаблон файла
        template_layout = QHBoxLayout()
        template_layout.addWidget(QLabel("Шаблон Excel:"))
        self.template_path_edit = QLineEdit()
        self.template_path_edit.setReadOnly(True)
        template_layout.addWidget(self.template_path_edit)
        self.template_browse_btn = QPushButton("Обзор...")
        self.template_browse_btn.clicked.connect(self.browse_template)
        template_layout.addWidget(self.template_browse_btn)
        files_layout.addLayout(template_layout)

        # Выходной файл
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("Путь для сохранения:"))
        self.output_path_edit = QLineEdit()
        self.output_path_edit.setReadOnly(True)
        output_layout.addWidget(self.output_path_edit)
        self.output_browse_btn = QPushButton("Обзор...")
        self.output_browse_btn.clicked.connect(self.browse_output)
        output_layout.addWidget(self.output_browse_btn)
        files_layout.addLayout(output_layout)

        layout.addWidget(files_group)

        # Группа настроек выходного файла
        output_cfg_group = QGroupBox("Параметры файла")
        output_cfg_layout = QVBoxLayout(output_cfg_group)

        title_layout = QHBoxLayout()
        title_layout.addWidget(QLabel("Заголовок:"))
        self.output_title_edit = QLineEdit("Стоимость доставки")
        title_layout.addWidget(self.output_title_edit)
        output_cfg_layout.addLayout(title_layout)

        subtitle_layout = QHBoxLayout()
        subtitle_layout.addWidget(QLabel("Подзаголовок:"))
        self.output_subtitle_edit = QLineEdit()
        subtitle_layout.addWidget(self.output_subtitle_edit)
        output_cfg_layout.addLayout(subtitle_layout)

        start_row_layout = QHBoxLayout()
        start_row_layout.addWidget(QLabel("Начало данных (строка):"))
        self.start_row_spin = QSpinBox()
        self.start_row_spin.setRange(1, 100)
        self.start_row_spin.setValue(3)
        start_row_layout.addWidget(self.start_row_spin)
        start_row_layout.addStretch()
        output_cfg_layout.addLayout(start_row_layout)

        self.include_average_check = QCheckBox("Включить строку \"Среднее значение\"")
        self.include_average_check.setChecked(True)
        output_cfg_layout.addWidget(self.include_average_check)


        self.markups_sheet_check = QCheckBox("Создать отдельный лист с наценками")
        self.markups_sheet_check.setChecked(True)
        output_cfg_layout.addWidget(self.markups_sheet_check)

        layout.addWidget(output_cfg_group)

        # Группа собственной компании
        own_group = QGroupBox("Собственная компания (строка после среднего)")
        own_layout = QVBoxLayout(own_group)

        self.own_enabled_check = QCheckBox("Добавлять строку собственной компании")
        self.own_enabled_check.setChecked(True)
        own_layout.addWidget(self.own_enabled_check)

        own_name_layout = QHBoxLayout()
        own_name_layout.addWidget(QLabel("Название:"))
        self.own_name_edit = QLineEdit("Новая Витэка")
        own_name_layout.addWidget(self.own_name_edit)
        own_layout.addLayout(own_name_layout)

        own_markups_label = QLabel("Наценки на среднее значение (%):")
        own_layout.addWidget(own_markups_label)

        own_markups_layout = QHBoxLayout()
        self.own_markup_fields = {}
        for field_key, field_name in [
            ('convert', 'Конверт'), ('minimum_1', 'Посылка 10кг'),
            ('minimum_2', '1 место 30кг'), ('volume', '0,5 куба'),
            ('weight_100', 'Груз 100кг'), ('weight_3000', 'Груз 3000кг')
        ]:
            col = QVBoxLayout()
            col.addWidget(QLabel(field_name + ':'))
            spin = QDoubleSpinBox()
            spin.setRange(-100, 1000)
            spin.setValue(0)
            spin.setSingleStep(0.5)
            spin.setMaximumWidth(75)
            col.addWidget(spin)
            own_markups_layout.addLayout(col)
            self.own_markup_fields[field_key] = spin
        own_markups_layout.addStretch()
        own_layout.addLayout(own_markups_layout)

        layout.addWidget(own_group)

        # Информация
        info_group = QGroupBox("Информация")
        info_layout = QVBoxLayout(info_group)

        self.info_label = QLabel()
        self.info_label.setWordWrap(True)
        info_layout.addWidget(self.info_label)

        layout.addWidget(info_group)
        layout.addStretch()

        self.tabs.addTab(tab, "Основное")

    def create_competitor_tab(self):
        """Вкладка настройки конкурентов."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Выбор конкурента
        competitor_select_layout = QHBoxLayout()
        competitor_select_layout.addWidget(QLabel("Конкурент:"))
        self.competitor_combo = QComboBox()
        self.competitor_combo.currentTextChanged.connect(self.on_competitor_changed)
        competitor_select_layout.addWidget(self.competitor_combo)

        self.add_competitor_btn = QPushButton("➕ Добавить")
        self.add_competitor_btn.clicked.connect(self.add_competitor)
        competitor_select_layout.addWidget(self.add_competitor_btn)

        self.delete_competitor_btn = QPushButton("🗑 Удалить")
        self.delete_competitor_btn.clicked.connect(self.delete_competitor)
        competitor_select_layout.addWidget(self.delete_competitor_btn)

        self.move_up_btn = QPushButton("▲ Вверх")
        self.move_up_btn.clicked.connect(self.move_competitor_up)
        competitor_select_layout.addWidget(self.move_up_btn)

        self.move_down_btn = QPushButton("▼ Вниз")
        self.move_down_btn.clicked.connect(self.move_competitor_down)
        competitor_select_layout.addWidget(self.move_down_btn)

        competitor_select_layout.addStretch()
        layout.addLayout(competitor_select_layout)

        # Настройки конкурента
        settings_group = QGroupBox("Настройки")
        settings_layout = QVBoxLayout(settings_group)

        # Файл
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("Файл:"))
        self.competitor_file_edit = QLineEdit()
        self.competitor_file_edit.setReadOnly(True)
        file_layout.addWidget(self.competitor_file_edit)
        self.competitor_file_btn = QPushButton("Обзор...")
        self.competitor_file_btn.clicked.connect(self.browse_competitor_file)
        file_layout.addWidget(self.competitor_file_btn)
        settings_layout.addLayout(file_layout)

        # Включен
        self.competitor_enabled_check = QCheckBox("Обрабатывать этого конкурента")
        self.competitor_enabled_check.setChecked(True)
        settings_layout.addWidget(self.competitor_enabled_check)

        self.competitor_bold_check = QCheckBox("Выделять строку жирным шрифтом в итоговом файле")
        self.competitor_bold_check.setChecked(False)
        settings_layout.addWidget(self.competitor_bold_check)

        # Колонки источника — все в одну строку
        source_cols_group = QGroupBox("Колонки в файле конкурента")
        source_cols_layout = QHBoxLayout(source_cols_group)

        for label, attr, default in [
            ("Город",       "src_city_edit",      "A"),
            ("Конверт",     "src_convert_edit",   "D"),
            ("Минималка 1", "src_min1_edit",       "E"),
            ("Минималка 2", "src_min2_edit",       "F"),
            ("Объем",       "src_volume_edit",     "O"),
            ("Вес 100",     "src_weight100_edit",  "P"),
            ("Вес 3000",    "src_weight3000_edit", "Q"),
        ]:
            source_cols_layout.addWidget(QLabel(label + ":"))
            edit = QLineEdit(default)
            edit.setMaximumWidth(45)
            source_cols_layout.addWidget(edit)
            setattr(self, attr, edit)

        source_cols_layout.addStretch()
        settings_layout.addWidget(source_cols_group)

        # Смещения строк — все в одну строку
        offsets_group = QGroupBox("Смещения строк")
        offsets_layout = QHBoxLayout(offsets_group)

        for label, attr in [
            ("Общее",       "offset_app_spin"),
            ("Конверт",     "offset_convert_spin"),
            ("Минималка 1", "offset_min1_spin"),
            ("Минималка 2", "offset_min2_spin"),
            ("Объем",       "offset_volume_spin"),
            ("Вес 100",     "offset_weight100_spin"),
            ("Вес 3000",    "offset_weight3000_spin"),
        ]:
            offsets_layout.addWidget(QLabel(label + ":"))
            spin = QSpinBox()
            spin.setRange(-100, 100)
            spin.setMaximumWidth(55)
            offsets_layout.addWidget(spin)
            setattr(self, attr, spin)

        offsets_layout.addStretch()
        settings_layout.addWidget(offsets_group)

        # Наценки — все в одну строку
        markups_group = QGroupBox("Наценки на цены (%)")
        markups_layout = QHBoxLayout(markups_group)

        for label, attr in [
            ("Конверт",     "markup_convert_spin"),
            ("Минималка 1", "markup_min1_spin"),
            ("Минималка 2", "markup_min2_spin"),
            ("Объем",       "markup_volume_spin"),
            ("Вес 100",     "markup_weight100_spin"),
            ("Вес 3000",    "markup_weight3000_spin"),
        ]:
            markups_layout.addWidget(QLabel(label + ":"))
            spin = QDoubleSpinBox()
            spin.setRange(-100, 1000)
            spin.setValue(0)
            spin.setSingleStep(0.5)
            spin.setMaximumWidth(65)
            markups_layout.addWidget(spin)
            setattr(self, attr, spin)

        markups_layout.addStretch()

        settings_layout.addWidget(markups_group)

        # Порог совпадения
        threshold_layout = QHBoxLayout()
        threshold_layout.addWidget(QLabel("Порог совпадения города (%):"))
        self.threshold_spin = QSpinBox()
        self.threshold_spin.setRange(50, 100)
        self.threshold_spin.setValue(95)
        threshold_layout.addWidget(self.threshold_spin)
        threshold_layout.addStretch()
        settings_layout.addLayout(threshold_layout)

        layout.addWidget(settings_group)

        # Группа дополнительных строк с наценками
        mk_rows_group = QGroupBox("Дополнительные строки с наценкой (перед «Среднее значение»)")
        mk_rows_layout = QVBoxLayout(mk_rows_group)

        self.markup_rows_table = QTableWidget()
        self.markup_rows_table.setColumnCount(2)
        self.markup_rows_table.setHorizontalHeaderLabels(["Название строки", "Наценка (%)"])
        self.markup_rows_table.horizontalHeader().setStretchLastSection(True)
        self.markup_rows_table.setMaximumHeight(150)
        mk_rows_layout.addWidget(self.markup_rows_table)

        mk_btn_layout = QHBoxLayout()
        add_mk_btn = QPushButton("➕ Добавить")
        add_mk_btn.clicked.connect(self.add_markup_row)
        mk_btn_layout.addWidget(add_mk_btn)

        del_mk_btn = QPushButton("➖ Удалить")
        del_mk_btn.clicked.connect(self.remove_markup_row)
        mk_btn_layout.addWidget(del_mk_btn)
        mk_btn_layout.addStretch()
        mk_rows_layout.addLayout(mk_btn_layout)

        layout.addWidget(mk_rows_group)

        # Кнопка сохранения
        save_btn = QPushButton("💾 Сохранить настройки конкурента")
        save_btn.clicked.connect(self.save_competitor_config)
        layout.addWidget(save_btn)

        self.tabs.addTab(tab, "Конкуренты")

    def create_cities_tab(self):
        """Вкладка управления городами."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        layout.addWidget(QLabel(
            "Список городов. Псевдонимы — доп. варианты написания через запятую (напр.: Астана, Нур-Султан, Astana):"
        ))

        self.cities_table = QTableWidget()
        self.cities_table.setColumnCount(3)
        self.cities_table.setHorizontalHeaderLabels(["Город", "Строка", "Псевдонимы (через запятую)"])
        self.cities_table.horizontalHeader().setStretchLastSection(True)
        self.cities_table.setColumnWidth(0, 160)
        self.cities_table.setColumnWidth(1, 60)
        layout.addWidget(self.cities_table)

        # Кнопки управления
        buttons_layout = QHBoxLayout()

        add_city_btn = QPushButton("➕ Добавить город")
        add_city_btn.clicked.connect(self.add_city)
        buttons_layout.addWidget(add_city_btn)

        remove_city_btn = QPushButton("➖ Удалить город")
        remove_city_btn.clicked.connect(self.remove_city)
        buttons_layout.addWidget(remove_city_btn)

        load_cities_btn = QPushButton("📂 Загрузить из JSON")
        load_cities_btn.clicked.connect(self.load_cities_from_json)
        buttons_layout.addWidget(load_cities_btn)

        save_cities_btn = QPushButton("💾 Сохранить")
        save_cities_btn.clicked.connect(self.save_cities)
        buttons_layout.addWidget(save_cities_btn)

        buttons_layout.addStretch()
        layout.addLayout(buttons_layout)

        self.tabs.addTab(tab, "Города")

    def create_preview_tab(self):
        """Вкладка предпросмотра данных."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Выбор конкурента
        select_layout = QHBoxLayout()
        select_layout.addWidget(QLabel("Конкурент:"))
        self.preview_competitor_combo = QComboBox()
        select_layout.addWidget(self.preview_competitor_combo)

        preview_btn = QPushButton("🔍 Просмотр")
        preview_btn.clicked.connect(self.preview_data)
        select_layout.addWidget(preview_btn)
        select_layout.addStretch()

        layout.addLayout(select_layout)

        # Таблица предпросмотра
        self.preview_table = QTableWidget()
        self.preview_table.setColumnCount(8)
        self.preview_table.setHorizontalHeaderLabels([
            "Строка", "Город", "Конверт", "Мин. 1", "Мин. 2",
            "Объем", "Вес 100", "Вес 3000"
        ])
        layout.addWidget(self.preview_table)

        self.tabs.addTab(tab, "Предпросмотр")

    def create_log_tab(self):
        """Вкладка логов."""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        layout.addWidget(QLabel("Журнал выполнения:"))

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Courier", 10))
        layout.addWidget(self.log_text)

        clear_log_btn = QPushButton("🗑 Очистить")
        clear_log_btn.clicked.connect(self.log_text.clear)
        layout.addWidget(clear_log_btn)

        self.tabs.addTab(tab, "Журнал")

    def create_control_panel(self) -> QHBoxLayout:
        """Создать панель управления."""
        layout = QHBoxLayout()

        self.run_btn = QPushButton("▶️ Запустить обработку")
        self.run_btn.setStyleSheet("QPushButton { font-size: 14px; padding: 10px; }")
        self.run_btn.clicked.connect(self.run_processing)
        layout.addWidget(self.run_btn)

        save_config_btn = QPushButton("💾 Сохранить конфигурацию")
        save_config_btn.clicked.connect(self.save_config)
        layout.addWidget(save_config_btn)

        layout.addStretch()

        return layout

    def setup_logging(self):
        """Настроить логирование в текстовое поле."""
        handler = QTextEditLogger(self.log_text)
        handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)

    def load_config_to_ui(self):
        """Загрузить конфигурацию в интерфейс."""
        # Основные настройки
        self.template_path_edit.setText(self.config.template_file)
        self.output_path_edit.setText(self.config.output_file)
        self.output_title_edit.setText(self.config.output_config.title)
        self.output_subtitle_edit.setText(self.config.output_config.subtitle)
        self.start_row_spin.setValue(self.config.output_config.start_row)
        self.include_average_check.setChecked(self.config.output_config.include_average)
        self.markups_sheet_check.setChecked(self.config.output_config.markups_sheet)

        # Собственная компания
        self.own_enabled_check.setChecked(self.config.own_company.enabled)
        self.own_name_edit.setText(self.config.own_company.name)
        for field_key, spin in self.own_markup_fields.items():
            spin.setValue(getattr(self.config.own_company.markups, field_key, 0.0))

        # Конкуренты
        self.competitor_combo.clear()
        self.preview_competitor_combo.clear()
        for name in self.config.competitors.keys():
            self.competitor_combo.addItem(name)
            self.preview_competitor_combo.addItem(name)

        # Города
        self.load_cities_to_table()

        # Информация
        self.update_info_label()

    def update_info_label(self):
        """Обновить информационный лейбл."""
        enabled_count = sum(1 for c in self.config.competitors.values() if c.enabled)
        total_count = len(self.config.competitors)
        cities_count = len(self.config.cities)

        info_text = f"""
        <b>Конкуренты:</b> {enabled_count} активных из {total_count}<br>
        <b>Городов:</b> {cities_count}<br>
        <b>Выходной файл:</b> {'✅ Указан' if self.config.output_file else '❌ Не указан'}
        """
        self.info_label.setText(info_text)


    def browse_output(self):
        """Выбрать выходной файл."""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Укажите выходной файл", "", "Excel Files (*.xlsx)"
        )
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            self.output_path_edit.setText(file_path)
            self.config.output_file = file_path
            self.update_info_label()

    def browse_template(self):
        """Выбрать шаблон Excel файла."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Укажите шаблон Excel файла", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.template_path_edit.setText(file_path)
            self.config.template_file = file_path
            self.update_info_label()

    def on_competitor_changed(self, name: str):
        """Обработка смены конкурента."""
        if not name or name not in self.config.competitors:
            return

        competitor = self.config.competitors[name]

        # Загрузить настройки в UI
        self.competitor_file_edit.setText(competitor.file_path)
        self.competitor_enabled_check.setChecked(competitor.enabled)
        self.competitor_bold_check.setChecked(competitor.bold)

        # Загрузить строки наценок
        self.markup_rows_table.setRowCount(len(competitor.markup_rows))
        for i, mk in enumerate(competitor.markup_rows):
            self.markup_rows_table.setItem(i, 0, QTableWidgetItem(mk.name))
            self.markup_rows_table.setItem(i, 1, QTableWidgetItem(str(mk.percent)))

        # Исходные колонки
        self.src_city_edit.setText(competitor.source_columns.city)
        self.src_convert_edit.setText(competitor.source_columns.convert)
        self.src_min1_edit.setText(competitor.source_columns.minimum_1)
        self.src_min2_edit.setText(competitor.source_columns.minimum_2)
        self.src_volume_edit.setText(competitor.source_columns.volume)
        self.src_weight100_edit.setText(competitor.source_columns.weight_100)
        self.src_weight3000_edit.setText(competitor.source_columns.weight_3000)

        # Смещения
        self.offset_app_spin.setValue(competitor.row_offsets.row_app)
        self.offset_convert_spin.setValue(competitor.row_offsets.row_2)
        self.offset_min1_spin.setValue(competitor.row_offsets.row_3)
        self.offset_min2_spin.setValue(competitor.row_offsets.row_4)
        self.offset_volume_spin.setValue(competitor.row_offsets.row_5)
        self.offset_weight100_spin.setValue(competitor.row_offsets.row_6)
        self.offset_weight3000_spin.setValue(competitor.row_offsets.row_7)

        # Наценки
        self.markup_convert_spin.setValue(competitor.markups.convert)
        self.markup_min1_spin.setValue(competitor.markups.minimum_1)
        self.markup_min2_spin.setValue(competitor.markups.minimum_2)
        self.markup_volume_spin.setValue(competitor.markups.volume)
        self.markup_weight100_spin.setValue(competitor.markups.weight_100)
        self.markup_weight3000_spin.setValue(competitor.markups.weight_3000)

        # Порог
        self.threshold_spin.setValue(competitor.fuzzy_match_threshold)

    def add_competitor(self):
        """Добавить нового конкурента."""
        from PySide6.QtWidgets import QInputDialog

        name, ok = QInputDialog.getText(self, "Новый конкурент", "Название конкурента:")
        if ok and name:
            if name in self.config.competitors:
                QMessageBox.warning(self, "Ошибка", "Конкурент с таким именем уже существует")
                return

            self.config.competitors[name] = CompetitorConfig(name=name)
            self.competitor_combo.addItem(name)
            self.preview_competitor_combo.addItem(name)
            self.competitor_combo.setCurrentText(name)
            self.update_info_label()

    def delete_competitor(self):
        """Удалить конкурента."""
        current = self.competitor_combo.currentText()
        if not current:
            return

        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить конкурента '{current}'?",
            QMessageBox.Yes | QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            del self.config.competitors[current]
            index = self.competitor_combo.currentIndex()
            self.competitor_combo.removeItem(index)
            self.preview_competitor_combo.removeItem(
                self.preview_competitor_combo.findText(current)
            )
            self.update_info_label()

    def _shift_competitor(self, direction: int):
        """Сдвинуть текущего конкурента на direction позиций (+1 вниз, -1 вверх)."""
        current = self.competitor_combo.currentText()
        if not current:
            return

        keys = list(self.config.competitors.keys())
        idx = keys.index(current)
        new_idx = idx + direction

        if new_idx < 0 or new_idx >= len(keys):
            return

        # Переставить в dict (Python 3.7+ сохраняет порядок)
        keys[idx], keys[new_idx] = keys[new_idx], keys[idx]
        self.config.competitors = {k: self.config.competitors[k] for k in keys}

        # Обновить ComboBox
        self.competitor_combo.blockSignals(True)
        self.competitor_combo.clear()
        self.preview_competitor_combo.clear()
        for name in keys:
            self.competitor_combo.addItem(name)
            self.preview_competitor_combo.addItem(name)
        self.competitor_combo.setCurrentText(current)
        self.competitor_combo.blockSignals(False)

    def move_competitor_up(self):
        """Переместить конкурента вверх."""
        self._shift_competitor(-1)

    def move_competitor_down(self):
        """Переместить конкурента вниз."""
        self._shift_competitor(1)

    def browse_competitor_file(self):
        """Выбрать файл конкурента."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл конкурента", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.competitor_file_edit.setText(file_path)

    def save_competitor_config(self):
        """Сохранить настройки текущего конкурента."""
        current = self.competitor_combo.currentText()
        if not current:
            return

        competitor = self.config.competitors[current]

        # Обновить настройки
        competitor.file_path = self.competitor_file_edit.text()
        competitor.enabled = self.competitor_enabled_check.isChecked()
        competitor.bold = self.competitor_bold_check.isChecked()

        # Исходные колонки
        competitor.source_columns.city = self.src_city_edit.text()
        competitor.source_columns.convert = self.src_convert_edit.text()
        competitor.source_columns.minimum_1 = self.src_min1_edit.text()
        competitor.source_columns.minimum_2 = self.src_min2_edit.text()
        competitor.source_columns.volume = self.src_volume_edit.text()
        competitor.source_columns.weight_100 = self.src_weight100_edit.text()
        competitor.source_columns.weight_3000 = self.src_weight3000_edit.text()

        # Смещения
        competitor.row_offsets.row_app = self.offset_app_spin.value()
        competitor.row_offsets.row_2 = self.offset_convert_spin.value()
        competitor.row_offsets.row_3 = self.offset_min1_spin.value()
        competitor.row_offsets.row_4 = self.offset_min2_spin.value()
        competitor.row_offsets.row_5 = self.offset_volume_spin.value()
        competitor.row_offsets.row_6 = self.offset_weight100_spin.value()
        competitor.row_offsets.row_7 = self.offset_weight3000_spin.value()

        # Наценки
        competitor.markups.convert = self.markup_convert_spin.value()
        competitor.markups.minimum_1 = self.markup_min1_spin.value()
        competitor.markups.minimum_2 = self.markup_min2_spin.value()
        competitor.markups.volume = self.markup_volume_spin.value()
        competitor.markups.weight_100 = self.markup_weight100_spin.value()
        competitor.markups.weight_3000 = self.markup_weight3000_spin.value()

        # Порог
        competitor.fuzzy_match_threshold = self.threshold_spin.value()

        # Строки наценок
        from src.models import MarkupRow
        mk_rows = []
        for i in range(self.markup_rows_table.rowCount()):
            name_item = self.markup_rows_table.item(i, 0)
            pct_item = self.markup_rows_table.item(i, 1)
            if name_item and pct_item:
                try:
                    mk_rows.append(MarkupRow(
                        name=name_item.text(),
                        percent=float(pct_item.text())
                    ))
                except ValueError:
                    pass
        competitor.markup_rows = mk_rows

        self.status_bar.showMessage(f"Настройки '{current}' сохранены", 3000)
        self.update_info_label()

    def load_cities_to_table(self):
        """Загрузить города в таблицу."""
        self.cities_table.setRowCount(len(self.config.cities))

        for i, (city, row) in enumerate(self.config.cities.items()):
            self.cities_table.setItem(i, 0, QTableWidgetItem(city))
            self.cities_table.setItem(i, 1, QTableWidgetItem(str(row)))
            aliases = self.config.city_aliases.get(city, [])
            self.cities_table.setItem(i, 2, QTableWidgetItem(", ".join(aliases)))

    def add_city(self):
        """Добавить город."""
        from PySide6.QtWidgets import QInputDialog

        city, ok = QInputDialog.getText(self, "Новый город", "Название города:")
        if not ok or not city:
            return

        row, ok = QInputDialog.getInt(self, "Строка", "Номер строки в выходном файле:", 1, 1, 10000)
        if ok:
            self.config.cities[city] = row
            self.load_cities_to_table()

    def remove_city(self):
        """Удалить выбранный город."""
        current_row = self.cities_table.currentRow()
        if current_row < 0:
            return

        city = self.cities_table.item(current_row, 0).text()
        del self.config.cities[city]
        self.load_cities_to_table()

    def add_markup_row(self):
        """Добавить строку наценки для текущего конкурента."""
        row = self.markup_rows_table.rowCount()
        self.markup_rows_table.insertRow(row)
        self.markup_rows_table.setItem(row, 0, QTableWidgetItem("+10%"))
        self.markup_rows_table.setItem(row, 1, QTableWidgetItem("10.0"))

    def remove_markup_row(self):
        """Удалить выбранную строку наценки."""
        current = self.markup_rows_table.currentRow()
        if current >= 0:
            self.markup_rows_table.removeRow(current)

    def save_cities(self):
        """Сохранить изменения в городах."""
        new_cities = {}
        new_aliases = {}
        for i in range(self.cities_table.rowCount()):
            city_item = self.cities_table.item(i, 0)
            row_item = self.cities_table.item(i, 1)
            aliases_item = self.cities_table.item(i, 2)
            if city_item and row_item:
                try:
                    city = city_item.text().strip()
                    new_cities[city] = int(row_item.text())
                    if aliases_item and aliases_item.text().strip():
                        aliases = [a.strip() for a in aliases_item.text().split(",") if a.strip()]
                        if aliases:
                            new_aliases[city] = aliases
                except ValueError:
                    pass

        self.config.cities = new_cities
        self.config.city_aliases = new_aliases
        self.status_bar.showMessage("Города сохранены", 3000)

    def load_cities_from_json(self):
        """Загрузить города из JSON файла."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите JSON файл с городами", "", "JSON Files (*.json)"
        )
        if file_path:
            try:
                import json
                with open(file_path, 'r', encoding='utf-8') as f:
                    cities = json.load(f)

                if isinstance(cities, dict):
                    self.config.cities.update(cities)
                    self.load_cities_to_table()
                    self.status_bar.showMessage("Города загружены", 3000)
                else:
                    QMessageBox.warning(self, "Ошибка", "Неверный формат JSON файла")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл: {e}")

    def preview_data(self):
        """Предпросмотр данных конкурента."""
        competitor_name = self.preview_competitor_combo.currentText()
        if not competitor_name or competitor_name not in self.config.competitors:
            return

        competitor = self.config.competitors[competitor_name]
        data = self.processor.preview_data(competitor, max_rows=20)

        self.preview_table.setRowCount(len(data))

        for i, row_data in enumerate(data):
            self.preview_table.setItem(i, 0, QTableWidgetItem(str(row_data['row'])))
            self.preview_table.setItem(i, 1, QTableWidgetItem(str(row_data.get('city', ''))))
            self.preview_table.setItem(i, 2, QTableWidgetItem(str(row_data.get('convert', ''))))
            self.preview_table.setItem(i, 3, QTableWidgetItem(str(row_data.get('minimum_1', ''))))
            self.preview_table.setItem(i, 4, QTableWidgetItem(str(row_data.get('minimum_2', ''))))
            self.preview_table.setItem(i, 5, QTableWidgetItem(str(row_data.get('volume', ''))))
            self.preview_table.setItem(i, 6, QTableWidgetItem(str(row_data.get('weight_100', ''))))
            self.preview_table.setItem(i, 7, QTableWidgetItem(str(row_data.get('weight_3000', ''))))

    def save_config(self):
        """Сохранить конфигурацию."""
        try:
            self.config.save(self.config_path)
            self.status_bar.showMessage("Конфигурация сохранена", 3000)
            logger.info("Конфигурация сохранена")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить конфигурацию: {e}")

    def run_processing(self):
        """Запустить обработку."""
        # Проверки
        if not self.config.output_file:
            QMessageBox.warning(self, "Ошибка", "Не указан выходной файл")
            return

        if not any(c.enabled for c in self.config.competitors.values()):
            QMessageBox.warning(self, "Ошибка", "Нет активных конкурентов")
            return

        if not self.config.cities:
            QMessageBox.warning(self, "Ошибка", "Не указаны города")
            return

        # Проверить что все файлы конкурентов существуют
        for competitor in self.config.competitors.values():
            if competitor.enabled and not competitor.file_path:
                QMessageBox.warning(
                    self, "Ошибка",
                    f"Не указан файл для конкурента '{competitor.name}'"
                )
                return

        # Сохранить конфигурацию (включая параметры выходного файла)
        self.config.template_file = self.template_path_edit.text()
        self.config.output_file = self.output_path_edit.text()
        self.config.output_config.title = self.output_title_edit.text()
        self.config.output_config.subtitle = self.output_subtitle_edit.text()
        self.config.output_config.start_row = self.start_row_spin.value()
        self.config.output_config.include_average = self.include_average_check.isChecked()
        self.config.output_config.markups_sheet = self.markups_sheet_check.isChecked()

        # Собственная компания
        self.config.own_company.enabled = self.own_enabled_check.isChecked()
        self.config.own_company.name = self.own_name_edit.text()
        for field_key, spin in self.own_markup_fields.items():
            setattr(self.config.own_company.markups, field_key, spin.value())

        self.save_config()

        # Обновить процессор
        self.processor = ExcelProcessor(self.config)

        # Запустить обработку в отдельном потоке
        self.run_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(0)  # Indeterminate progress

        self.processing_thread = ProcessingThread(self.processor)
        self.processing_thread.progress.connect(self.on_processing_progress)
        self.processing_thread.finished.connect(self.on_processing_finished)
        self.processing_thread.start()

        self.status_bar.showMessage("Обработка запущена...")
        logger.info("=" * 50)
        logger.info("Начало обработки")

    def on_processing_progress(self, competitor_name: str, is_done: bool):
        """Обработка прогресса."""
        if is_done:
            logger.info(f"✅ {competitor_name} - обработан")
            self.status_bar.showMessage(f"Обработан: {competitor_name}")
        else:
            logger.info(f"⏳ Обработка: {competitor_name}")
            self.status_bar.showMessage(f"Обработка: {competitor_name}")

    def on_processing_finished(self, results: list):
        """Обработка завершена."""
        self.run_btn.setEnabled(True)
        self.progress_bar.setVisible(False)

        # Подсчет статистики
        total = len(results)
        successful = sum(1 for r in results if r['success'])
        total_cities = sum(r['processed_cities'] for r in results)

        logger.info("=" * 50)
        logger.info(f"Обработка завершена!")
        logger.info(f"Обработано конкурентов: {successful}/{total}")
        logger.info(f"Всего обработано городов: {total_cities}")

        # Показать ошибки если есть
        errors = []
        for r in results:
            if r['errors']:
                errors.extend([f"{r['competitor']}: {e}" for e in r['errors']])

        if errors:
            logger.warning("Ошибки:")
            for error in errors:
                logger.warning(f"  - {error}")

        # Показать сообщение
        if successful == total:
            QMessageBox.information(
                self, "Готово",
                f"Обработка завершена успешно!\n\n"
                f"Конкурентов: {total}\n"
                f"Городов: {total_cities}\n\n"
                f"Результат сохранен в:\n{self.config.output_file}"
            )
        else:
            QMessageBox.warning(
                self, "Завершено с ошибками",
                f"Обработано: {successful}/{total} конкурентов\n"
                f"Городов: {total_cities}\n\n"
                f"Проверьте журнал для деталей"
            )

        self.status_bar.showMessage("Готово", 5000)


class QTextEditLogger(logging.Handler):
    """Хэндлер для вывода логов в QTextEdit."""

    def __init__(self, text_edit: QTextEdit):
        super().__init__()
        self.text_edit = text_edit

    def emit(self, record):
        msg = self.format(record)
        self.text_edit.append(msg)


def main():
    """Главная функция запуска приложения."""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()

