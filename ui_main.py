import sys
import json
import os
import pandas as pd
import math
from collections import Counter, defaultdict
from datetime import datetime, timedelta
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTextEdit, QGroupBox,
    QSplitter, QMessageBox, QProgressBar, QFrame, QListWidget,
    QListWidgetItem, QCheckBox, QTableWidget, QTableWidgetItem,
    QTabWidget, QHeaderView, QStackedWidget, QTreeWidget, 
    QTreeWidgetItem, QDoubleSpinBox, QScrollArea, QLineEdit,
    QComboBox, QInputDialog, QGraphicsView, QGraphicsScene,
    QAbstractItemView, QDateEdit
)
from PyQt5.QtCore import Qt, QPointF, QRectF, QThread, QTimer, pyqtSignal, QDate
from PyQt5.QtGui import (
    QFont, QFontMetrics, QPalette, QColor, QBrush, QPen,
    QImage, QPainter, QPainterPath, QPolygonF, QTransform
)
from excel_parser import (
    parse_excel_to_json, get_excel_info, ExcelParseError,
    detect_column_indices, normalize_cell_value, parse_datetime
)
from process_insights import (
    action_name, delete_records, duplicate_indices_except_first, find_case_timeline,
    find_duplicate_records, find_missing_records, format_timedelta,
    group_processes, parse_dt, common_final_action_ids,
    load_json_data, most_likely_value, process_map,
    save_json_data, set_record_value, suggest_action_replacements,
    time_dynamics
)

class ParseWorker(QThread):
    """Фоновый поток для парсинга Excel файла"""
    finished = pyqtSignal(str)  # Сигнал успешного завершения с результатом
    error = pyqtSignal(str)     # Сигнал ошибки с сообщением
    progress = pyqtSignal(str)  # Сигнал прогресса

    def __init__(self, filepath, custom_action_order=None):
        super().__init__()
        self.filepath = filepath
        self.custom_action_order = custom_action_order

    def run(self):
        try:
            self.progress.emit("Читаю файл...")
            json_result = parse_excel_to_json(self.filepath, self.custom_action_order)
            self.finished.emit(json_result)
        except Exception as e:
            self.error.emit(str(e))


class ProcessMapView(QGraphicsView):
    """Canvas-like view: full graph at open, left-button panning after that."""

    def __init__(self, scene):
        super().__init__(scene)
        self._fit_pending = True
        self._fit_rect = None
        self._is_panning = False
        self._last_mouse_position = None
        self._pan_start_scene_position = None
        self._pan_start_center = None
        self._min_zoom = 0.08
        self._max_zoom = 3.5
        self.setRenderHint(QPainter.Antialiasing)
        self.setDragMode(QGraphicsView.NoDrag)
        self.setInteractive(False)
        self.setTransformationAnchor(QGraphicsView.NoAnchor)
        self.setResizeAnchor(QGraphicsView.AnchorViewCenter)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setCursor(Qt.OpenHandCursor)

    def set_content_rect(self, rect):
        self._fit_rect = rect
        padding_x = max(rect.width() * 1.5, 900)
        padding_y = max(rect.height() * 1.5, 700)
        self.scene().setSceneRect(rect.adjusted(-padding_x, -padding_y, padding_x, padding_y))
        self.request_fit()

    def request_fit(self):
        self._fit_pending = True
        QTimer.singleShot(0, self.fit_scene_once)

    def fit_scene_once(self):
        target_rect = self._fit_rect or self.sceneRect()
        if not self.isVisible() or self.viewport().width() < 20 or self.viewport().height() < 20:
            self._fit_pending = True
            return
        if self.scene() and not target_rect.isEmpty():
            self.resetTransform()
            self.fitInView(target_rect, Qt.KeepAspectRatio)
        self._fit_pending = False

    def showEvent(self, event):
        super().showEvent(event)
        if self._fit_pending:
            self.request_fit()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self._fit_pending:
            self.request_fit()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._is_panning = True
            self._last_mouse_position = event.pos()
            self._pan_start_scene_position = self.mapToScene(event.pos())
            self._pan_start_center = self.mapToScene(self.viewport().rect().center())
            self.setCursor(Qt.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if (
            self._is_panning
            and self._pan_start_scene_position is not None
            and self._pan_start_center is not None
        ):
            current_scene_position = self.mapToScene(event.pos())
            delta = self._pan_start_scene_position - current_scene_position
            self.centerOn(self._pan_start_center + delta)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._is_panning:
            self._is_panning = False
            self._last_mouse_position = None
            self._pan_start_scene_position = None
            self._pan_start_center = None
            self.setCursor(Qt.OpenHandCursor)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def wheelEvent(self, event):
        delta = event.angleDelta().y()
        if not delta:
            event.ignore()
            return

        cursor_position = event.pos()
        scene_before = self.mapToScene(cursor_position)
        current_zoom = self.transform().m11()
        zoom_in = delta > 0
        factor = 1.16 if zoom_in else 1 / 1.16
        target_zoom = current_zoom * factor

        if target_zoom < self._min_zoom:
            factor = self._min_zoom / current_zoom
        elif target_zoom > self._max_zoom:
            factor = self._max_zoom / current_zoom

        if factor != 1:
            self.scale(factor, factor)
            scene_after = self.mapToScene(cursor_position)
            scene_delta = scene_after - scene_before
            self.translate(scene_delta.x(), scene_delta.y())
        event.accept()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Анализ бизнес-процессов — Парсинг')
        self.setGeometry(100, 100, 1000, 700)
        self.current_file = None
        self.custom_action_order = None
        self.json_result = None
        self.json_file_path = None
        self.display_mode = 'parsing'  # 'parsing' или 'statistics'
        self.statistics_widget = None
        self.quality_widget = None
        self.graphics_widget = None
        self.export_widget = None
        self.cached_quality_data = None
        self.cached_graphics_data = None
        self.cached_export_data = None
        self.cached_statistics_data = None  # Кэш данных статистики
        self.statistics_tab_widget = None
        self.case_detail_tab_index = None
        self.case_detail_input = None
        self.case_detail_refresh = None
        self.quality_summary_label = None
        self.quality_tabs = None
        self.export_checkboxes = {}
        self.export_controls = {}
        self.process_map_lane_spacing = 65.0
        self.actions_cache_key = None
        self.actions_cache_actions = set()
        self.actions_cache_records = []
        self.actions_cache_auto_order = None
        self.actions_cache_alpha_order = None
        self.analysis_cache_key = None
        self.analysis_data = None
        self.analysis_missing_records = None
        self.analysis_statistics_analyzer = None
        self.analysis_results_cache = {}
        self.analysis_time_dynamics_cache = {}
        self.analysis_outliers_cache = {}
        self.init_ui()

    def init_ui(self):
        # Создаем центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Основной layout
        main_layout = QVBoxLayout()

        # Панель управления (верхняя часть)
        control_layout = QHBoxLayout()

        # Кнопки
        self.btn_select_file = QPushButton('Выбрать Excel файл')
        self.btn_select_file.setMinimumHeight(35)
        self.btn_select_file.clicked.connect(self.select_file)

        self.btn_parse = QPushButton('Парсить данные')
        self.btn_parse.setMinimumHeight(35)
        self.btn_parse.setEnabled(False)
        self.btn_parse.clicked.connect(self.start_parsing)

        self.btn_select_json = QPushButton('Выбрать готовый JSON')
        self.btn_select_json.setMinimumHeight(35)
        self.btn_select_json.clicked.connect(self.select_json_file)

        self.btn_clear = QPushButton('Очистить')
        self.btn_clear.setMinimumHeight(35)
        self.btn_clear.setMinimumWidth(110)
        self.btn_clear.clicked.connect(self.clear_results)

        self.btn_statistics = QPushButton('Статистика>>>')
        self.btn_statistics.setMinimumHeight(35)
        self.btn_statistics.setEnabled(False)
        self.btn_statistics.clicked.connect(self.show_statistics)

        self.btn_back_to_parsing = QPushButton('<<<Парсинг')
        self.btn_back_to_parsing.setMinimumHeight(35)
        self.btn_back_to_parsing.clicked.connect(self.show_parsing)
        self.btn_back_to_parsing.setVisible(False)  # Изначально скрыта

        self.btn_to_quality = QPushButton('Качество данных>>>')
        self.btn_to_quality.setMinimumHeight(35)
        self.btn_to_quality.setEnabled(False)
        self.btn_to_quality.clicked.connect(self.show_quality)

        self.btn_back_to_quality = QPushButton('<<<Качество данных')
        self.btn_back_to_quality.setMinimumHeight(35)
        self.btn_back_to_quality.clicked.connect(self.show_quality)
        self.btn_back_to_quality.setVisible(False)

        self.btn_to_graphics = QPushButton('Графика>>>')
        self.btn_to_graphics.setMinimumHeight(35)
        self.btn_to_graphics.clicked.connect(self.show_graphics)
        self.btn_to_graphics.setVisible(False)

        self.btn_back_to_statistics = QPushButton('<<<Статистика')
        self.btn_back_to_statistics.setMinimumHeight(35)
        self.btn_back_to_statistics.clicked.connect(self.show_statistics)
        self.btn_back_to_statistics.setVisible(False)

        self.btn_to_export = QPushButton('Экспорт>>>')
        self.btn_to_export.setMinimumHeight(35)
        self.btn_to_export.clicked.connect(self.show_export)
        self.btn_to_export.setVisible(False)

        self.btn_back_to_graphics = QPushButton('<<<Графика')
        self.btn_back_to_graphics.setMinimumHeight(35)
        self.btn_back_to_graphics.clicked.connect(self.show_graphics)
        self.btn_back_to_graphics.setVisible(False)

        control_layout.addWidget(self.btn_back_to_parsing)
        control_layout.addWidget(self.btn_back_to_quality)
        control_layout.addWidget(self.btn_back_to_statistics)
        control_layout.addWidget(self.btn_back_to_graphics)
        control_layout.addWidget(self.btn_select_file)
        control_layout.addWidget(self.btn_parse)
        control_layout.addWidget(self.btn_select_json)
        control_layout.addWidget(self.btn_clear)
        control_layout.addStretch()
        control_layout.addWidget(self.btn_to_quality)
        control_layout.addWidget(self.btn_statistics)
        control_layout.addWidget(self.btn_to_graphics)
        control_layout.addWidget(self.btn_to_export)

        # === СОЗДАЕМ STACKED WIDGET ДЛЯ ПЕРЕКЛЮЧЕНИЯ МЕЖДУ ПАРСИНГОМ И СТАТИСТИКОЙ ===
        self.stacked_widget = QStackedWidget()

        # === СТРАНИЦА ПАРСИНГА (индекс 0) ===
        self.parsing_page = QWidget()
        parsing_layout = QVBoxLayout()

        # Информационная панель
        self.info_group = QGroupBox('Информация о файле')
        info_layout = QVBoxLayout()

        self.lbl_file_info = QLabel('Файл не выбран')
        self.lbl_file_info.setWordWrap(True)

        self.lbl_file_details = QLabel('')
        self.lbl_file_details.setWordWrap(True)

        info_layout.addWidget(self.lbl_file_info)
        info_layout.addWidget(self.lbl_file_details)
        self.info_group.setLayout(info_layout)

        # Панель управления порядком действий
        self.actions_group = QGroupBox('Порядок действий')
        actions_layout = QVBoxLayout()

        # Чекбокс для автоматической сортировки
        self.chk_auto_sort = QCheckBox('Автоматическая сортировка действий')
        self.chk_auto_sort.setChecked(True)  # По умолчанию включена
        self.chk_auto_sort.stateChanged.connect(self.toggle_auto_sort)

        # Список действий
        self.list_actions = QListWidget()
        self.list_actions.setMaximumHeight(150)

        # Кнопки управления
        buttons_layout = QHBoxLayout()
        self.btn_move_up = QPushButton('↑ Вверх')
        self.btn_move_up.clicked.connect(self.move_action_up)
        self.btn_move_up.setMaximumWidth(80)

        self.btn_move_down = QPushButton('↓ Вниз')
        self.btn_move_down.clicked.connect(self.move_action_down)
        self.btn_move_down.setMaximumWidth(80)

        self.btn_reset_order = QPushButton('Сбросить')
        self.btn_reset_order.clicked.connect(self.reset_action_order)
        self.btn_reset_order.setMaximumWidth(80)

        buttons_layout.addWidget(self.btn_move_up)
        buttons_layout.addWidget(self.btn_move_down)
        buttons_layout.addWidget(self.btn_reset_order)
        buttons_layout.addStretch()

        actions_layout.addWidget(self.chk_auto_sort)
        actions_layout.addWidget(self.list_actions)
        actions_layout.addLayout(buttons_layout)
        self.actions_group.setLayout(actions_layout)

        # Панель прогресса
        progress_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.lbl_progress = QLabel('')
        self.lbl_progress.setVisible(False)

        progress_layout.addWidget(self.lbl_progress)
        progress_layout.addWidget(self.progress_bar)

        # Панель результатов
        self.results_group = QGroupBox('Результаты парсинга')
        results_layout = QVBoxLayout()

        self.text_results = QTextEdit()
        self.text_results.setReadOnly(True)
        self.text_results.setFont(QFont('Courier New', 10))

        results_layout.addWidget(self.text_results)
        self.results_group.setLayout(results_layout)

        # Разделитель для вертикального разделения на странице парсинга
        self.parsing_splitter = QSplitter(Qt.Vertical)

        # Верхняя часть (информация + действия + прогресс)
        top_widget = QWidget()
        top_layout = QVBoxLayout()
        top_layout.addWidget(self.info_group)
        top_layout.addWidget(self.actions_group)
        top_layout.addLayout(progress_layout)
        top_widget.setLayout(top_layout)

        # Добавляем виджеты в разделитель
        self.parsing_splitter.addWidget(top_widget)
        self.parsing_splitter.addWidget(self.results_group)

        # Устанавливаем пропорции разделителя
        self.parsing_splitter.setSizes([200, 500])

        parsing_layout.addWidget(self.parsing_splitter)
        self.parsing_page.setLayout(parsing_layout)

        # === СТРАНИЦА СТАТИСТИКИ (индекс 1) ===
        self.quality_page = QWidget()
        self.quality_layout = QVBoxLayout()
        self.quality_page.setLayout(self.quality_layout)

        self.statistics_page = QWidget()
        self.statistics_layout = QVBoxLayout()
        self.statistics_page.setLayout(self.statistics_layout)

        self.graphics_page = QWidget()
        self.graphics_layout = QVBoxLayout()
        self.graphics_page.setLayout(self.graphics_layout)

        self.export_page = QWidget()
        self.export_layout = QVBoxLayout()
        self.export_page.setLayout(self.export_layout)

        # Добавляем страницы в stacked widget
        self.stacked_widget.addWidget(self.parsing_page)      # индекс 0
        self.stacked_widget.addWidget(self.quality_page)      # индекс 1
        self.stacked_widget.addWidget(self.statistics_page)   # индекс 2
        self.stacked_widget.addWidget(self.graphics_page)     # индекс 3
        self.stacked_widget.addWidget(self.export_page)       # индекс 4

        # Добавляем все в основной layout
        main_layout.addLayout(control_layout)
        main_layout.addWidget(self.stacked_widget)

        central_widget.setLayout(main_layout)

        # Настраиваем цвета и стили
        self.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 1ex;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 10px 0 10px;
            }
            QPushButton {
                padding: 8px 16px;
                font-size: 12px;
            }
            QPushButton:disabled {
                background-color: #f0f0f0;
                color: #999999;
            }
        """)
        self._set_navigation_mode('parsing')

    def select_file(self):
        """Выбор Excel файла"""
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            'Выбрать Excel файл',
            '',
            'Excel файлы (*.xlsx *.xls);;Все файлы (*)'
        )

        if filepath:
            self.current_file = filepath
            self.lbl_file_info.setText(f'Выбран файл: {filepath}')

            # Показываем информацию о файле
            try:
                file_info = get_excel_info(filepath)
                details = f"""
Обнаружено столбцов: {len(file_info['columns'])}
Всего строк: {file_info['total_rows']}

Определенные столбцы:
• ID кейса: {file_info['detected_columns']['case_id']}
• Действия: {file_info['detected_columns']['action']}
• Дата/время: {file_info['detected_columns']['datetime']}

Пример первой строки:
{json.dumps(file_info['sample_data']['first_row'], indent=2, ensure_ascii=False) if file_info['sample_data']['first_row'] else 'Нет данных'}
                """.strip()

                self.lbl_file_details.setText(details)

                # Заполняем список действий
                self.populate_actions_list(filepath)

                self.btn_parse.setEnabled(True)

            except ExcelParseError as e:
                self.lbl_file_details.setText(f'Ошибка анализа файла: {e}')
                self.list_actions.clear()
                self.btn_parse.setEnabled(False)
        else:
            self.lbl_file_info.setText('Файл не выбран')
            self.lbl_file_details.setText('')
            self.list_actions.clear()
            self.btn_parse.setEnabled(False)

    def start_parsing(self):
        """Запуск парсинга в фоновом потоке"""
        if not self.current_file:
            return

        # Показываем прогресс
        self.progress_bar.setVisible(True)
        self.lbl_progress.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Неопределенный прогресс
        self.lbl_progress.setText('Подготовка к парсингу...')

        # Отключаем кнопки
        self.btn_select_file.setEnabled(False)
        self.btn_parse.setEnabled(False)
        self.btn_clear.setEnabled(False)

        # Передаем в парсер ровно тот порядок, который пользователь видит в UI.
        visible_action_order = [
            self.list_actions.item(index).text()
            for index in range(self.list_actions.count())
        ]
        self.custom_action_order = visible_action_order or None

        # Создаем и запускаем поток парсинга
        self.worker = ParseWorker(self.current_file, self.custom_action_order)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.parsing_finished)
        self.worker.error.connect(self.parsing_error)
        self.worker.start()

    def update_progress(self, message):
        """Обновление прогресса"""
        self.lbl_progress.setText(message)

    def parsing_finished(self, json_result):
        """Обработка успешного завершения парсинга"""
        # Сохраняем результат для возможности сохранения
        self.json_result = json_result

        # Сбрасываем кэш статистики, т.к. данные могли измениться
        self.cached_statistics_data = None
        self.cached_quality_data = None
        self.cached_graphics_data = None
        self.cached_export_data = None
        self.statistics_widget = None
        self.quality_widget = None
        self.graphics_widget = None
        self.export_widget = None
        self._reset_analysis_cache()

        # Очищаем страницу статистики при новом парсинге
        while self.statistics_layout.count():
            item = self.statistics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self.quality_summary_label = None
        self.quality_tabs = None
        self._clear_layout(self.quality_layout)
        self._clear_layout(self.graphics_layout)
        self._clear_layout(self.export_layout)

        try:
            # Парсим JSON для красивого отображения
            data = json.loads(json_result)

            # Форматируем результат для отображения
            formatted_result = f"""# РЕЗУЛЬТАТЫ ПАРСИНГА

## Статистика
- Всего записей: {data['metadata']['total_records']}
- Уникальных действий: {data['metadata']['unique_actions']}

## Определенные столбцы
- ID кейса: {data['metadata']['columns_detected']['case_id_column']}
- Действия: {data['metadata']['columns_detected']['action_column']}
- Дата/время: {data['metadata']['columns_detected']['datetime_column']}

## Список действий
{chr(10).join(f'{i+1}. {action["name"]}{" (тупик)" if action.get("is_dead_end") else ""}' for i, action in enumerate(data['actions']))}

## Анализ последовательностей
"""

            # Добавляем информацию о последовательностях если она есть
            if 'sequence_analysis' in data['metadata'] and data['metadata']['sequence_analysis']:
                seq_info = data['metadata']['sequence_analysis']
                formatted_result += f"""Всего процессов: {seq_info.get('total_processes', 0)}
Средняя длина процесса: {seq_info.get('average_process_length', 0):.1f} действий
Стартовые действия: {', '.join(seq_info.get('start_actions', []))}
Финальные действия: {', '.join(seq_info.get('end_actions', []))}
"""

            formatted_result += f"""
## Примеры записей
"""

            # Добавляем первые 5 записей
            for i, record in enumerate(data['records'][:5]):
                action_id = record.get('action_id')
                action_meta = (
                    data['actions'][action_id]
                    if isinstance(action_id, int) and 0 <= action_id < len(data['actions'])
                    else {}
                )
                resolved_action_name = action_meta.get('name', 'Нет действия')
                is_dead_end = action_meta.get('is_dead_end', False)

                formatted_result += f"""
{i+1}. Case {record['case_id']} - {resolved_action_name}{" (тупик)" if is_dead_end else ""}
   Время: {record['datetime']}
"""

            if len(data['records']) > 5:
                formatted_result += f"\n... и еще {len(data['records']) - 5} записей\n"

            # Создаем ограниченную версию JSON с первыми 100 записями
            limited_data = data.copy()
            if len(data['records']) > 100:
                limited_data['records'] = data['records'][:100]
                limited_data['metadata'] = data['metadata'].copy()
                limited_data['metadata']['total_records'] = 100
                limited_data['metadata']['note'] = f'Показаны первые 100 из {len(data["records"])} записей'

            limited_json = json.dumps(limited_data, indent=2, ensure_ascii=False)
            formatted_result += f"\n## JSON результат (первые 100 записей)\n{limited_json}"


            self.text_results.setText(formatted_result)

            # Автоматически сохраняем JSON файл
            if self.current_file:
                import os
                base_name = os.path.splitext(os.path.basename(self.current_file))[0]
                auto_save_path = f"{base_name}_parsed.json"

                try:
                    with open(auto_save_path, 'w', encoding='utf-8') as f:
                        f.write(json_result)

                    QMessageBox.information(
                        self,
                        'Парсинг завершен',
                        f'Файл успешно обработан!\n\nJSON сохранен как: {auto_save_path}'
                    )

                    # Включаем кнопку статистики
                    self.btn_statistics.setEnabled(True)
                    self.json_file_path = auto_save_path
                    self.btn_to_quality.setEnabled(True)
                    self._set_navigation_mode('parsing')

                except Exception as e:
                    QMessageBox.warning(
                        self,
                        'Парсинг завершен',
                        f'Файл успешно обработан!\n\nПредупреждение: Не удалось автоматически сохранить JSON файл: {e}'
                    )
            else:
                QMessageBox.information(
                    self,
                    'Парсинг завершен',
                    'Файл успешно обработан!'
                )

        except Exception as e:
            self.text_results.setText(f'Ошибка обработки результатов: {e}\n\nПолный JSON:\n{json_result}')

        # Скрываем прогресс и включаем кнопки
        self.hide_progress()

    def parsing_error(self, error_message):
        """Обработка ошибки парсинга"""
        self.text_results.setText(f'ОШИБКА ПАРСИНГА:\n\n{error_message}')

        QMessageBox.critical(
            self,
            'Ошибка парсинга',
            f'Произошла ошибка при обработке файла:\n\n{error_message}'
        )

        # Скрываем прогресс и включаем кнопки
        self.hide_progress()

    def hide_progress(self):
        """Скрытие индикатора прогресса и включение кнопок"""
        self.progress_bar.setVisible(False)
        self.lbl_progress.setVisible(False)
        self.lbl_progress.setText('')

        self.btn_select_file.setEnabled(True)
        self.btn_parse.setEnabled(bool(self.current_file))
        self.btn_clear.setEnabled(True)

    def _actions_cache_key(self, filepath):
        try:
            return (filepath, os.path.getmtime(filepath), os.path.getsize(filepath))
        except OSError:
            return (filepath, None, None)

    def _reset_actions_cache(self):
        self.actions_cache_key = None
        self.actions_cache_actions = set()
        self.actions_cache_records = []
        self.actions_cache_auto_order = None
        self.actions_cache_alpha_order = None

    def _load_actions_cache(self, filepath):
        cache_key = self._actions_cache_key(filepath)
        if self.actions_cache_key == cache_key:
            return

        # Читаем Excel один раз для списка действий и дальнейших переключений сортировки.
        df = pd.read_excel(filepath, dtype=str)
        column_indices = detect_column_indices(df.columns)

        actions_set = set()
        records = []
        for _, row in df.iterrows():
            action = normalize_cell_value(row.iloc[column_indices['action']])
            if action:
                actions_set.add(action)

            case_raw = normalize_cell_value(row.iloc[column_indices['case_id']])
            dt_raw = normalize_cell_value(row.iloc[column_indices['datetime']])
            try:
                dt = parse_datetime(dt_raw)
                if dt is None or not action:
                    continue
                records.append({
                    'case_id': case_raw,
                    'action': action,
                    'datetime': dt
                })
            except Exception:
                pass

        self.actions_cache_key = cache_key
        self.actions_cache_actions = actions_set
        self.actions_cache_records = records
        self.actions_cache_auto_order = None
        self.actions_cache_alpha_order = None

    def _get_alphabetical_action_order(self):
        if self.actions_cache_alpha_order is None:
            self.actions_cache_alpha_order = sorted(self.actions_cache_actions)
        return list(self.actions_cache_alpha_order)

    def _get_auto_action_order(self):
        if self.actions_cache_auto_order is None:
            if self.actions_cache_records:
                try:
                    from sequence_analyzer import sort_actions_automatically
                    actions_list = [
                        action for action in sort_actions_automatically(self.actions_cache_records)
                        if action in self.actions_cache_actions
                    ]
                    actions_list.extend(sorted(self.actions_cache_actions - set(actions_list)))
                    if not actions_list:
                        actions_list = self._get_alphabetical_action_order()
                except Exception:
                    actions_list = self._get_alphabetical_action_order()
            else:
                actions_list = self._get_alphabetical_action_order()
            self.actions_cache_auto_order = list(actions_list)
        return list(self.actions_cache_auto_order)

    def populate_actions_list(self, filepath):
        """Заполнение списка действий из Excel файла"""
        try:
            self._load_actions_cache(filepath)

            # Определяем порядок действий
            if self.chk_auto_sort.isChecked():
                actions_list = self._get_auto_action_order()
            else:
                actions_list = self.custom_action_order or self._get_auto_action_order()

            # Заполняем список
            self.list_actions.clear()
            for action in actions_list:
                self.list_actions.addItem(action)

            # Сбрасываем пользовательский порядок
            self.custom_action_order = None

        except Exception as e:
            self.list_actions.clear()
            QMessageBox.warning(self, 'Ошибка', f'Не удалось загрузить действия: {e}')

    def clear_results(self):
        """Очистка результатов"""
        self.text_results.clear()
        self.lbl_file_info.setText('Файл не выбран')
        self.lbl_file_details.setText('')
        self.list_actions.clear()
        self.current_file = None
        self.custom_action_order = None
        self._reset_actions_cache()
        self.json_result = None
        self.json_file_path = None
        self.display_mode = 'parsing'
        self.cached_statistics_data = None
        self.statistics_widget = None
        self.cached_quality_data = None
        self.cached_graphics_data = None
        self.cached_export_data = None
        self.quality_widget = None
        self.graphics_widget = None
        self.export_widget = None
        self._reset_analysis_cache()

        # Восстанавливаем заголовок окна
        self.setWindowTitle('Анализ бизнес-процессов — Парсинг')

        # Очищаем страницу статистики
        while self.statistics_layout.count():
            item = self.statistics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        self._clear_layout(self.export_layout)

        # Показываем все элементы управления (режим парсинга)
        self.btn_select_file.setVisible(True)
        self.btn_parse.setVisible(True)
        self.btn_select_json.setVisible(True)
        self.btn_clear.setVisible(True)
        self.btn_statistics.setVisible(True)
        self.btn_back_to_parsing.setVisible(False)

        self.btn_parse.setEnabled(False)
        self.btn_statistics.setEnabled(False)
        self.btn_to_quality.setEnabled(False)
        self._set_navigation_mode('parsing')

        # Переключаем на страницу парсинга
        self.stacked_widget.setCurrentIndex(0)

    def _set_navigation_mode(self, mode):
        has_json = bool(self.json_file_path)
        self.display_mode = mode

        self.btn_back_to_parsing.setVisible(mode == 'quality')
        self.btn_back_to_quality.setVisible(mode == 'statistics')
        self.btn_back_to_statistics.setVisible(mode == 'graphics')
        self.btn_back_to_graphics.setVisible(mode == 'export')

        parsing_controls_visible = mode == 'parsing'
        self.btn_select_file.setVisible(parsing_controls_visible)
        self.btn_parse.setVisible(parsing_controls_visible)
        self.btn_select_json.setVisible(parsing_controls_visible)
        self.btn_clear.setVisible(parsing_controls_visible)

        self.btn_to_quality.setVisible(mode == 'parsing')
        self.btn_statistics.setVisible(mode == 'quality')
        self.btn_to_graphics.setVisible(mode == 'statistics')
        self.btn_to_export.setVisible(mode == 'graphics')

        self.btn_to_quality.setEnabled(has_json)
        self.btn_statistics.setEnabled(has_json)
        self.btn_to_graphics.setEnabled(has_json)
        self.btn_to_export.setEnabled(has_json)

    def _clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

    def _invalidate_analysis_pages(self):
        self.cached_statistics_data = None
        self.cached_quality_data = None
        self.cached_graphics_data = None
        self.cached_export_data = None
        self.statistics_widget = None
        self.quality_widget = None
        self.graphics_widget = None
        self.export_widget = None
        self.quality_summary_label = None
        self.quality_tabs = None
        self._reset_analysis_cache()
        self._clear_layout(self.statistics_layout)
        self._clear_layout(self.quality_layout)
        self._clear_layout(self.graphics_layout)
        self._clear_layout(self.export_layout)

    def _json_cache_key(self):
        if not self.json_file_path:
            return None
        try:
            return (
                os.path.abspath(self.json_file_path),
                os.path.getmtime(self.json_file_path),
                os.path.getsize(self.json_file_path),
            )
        except OSError:
            return (os.path.abspath(self.json_file_path), None, None)

    def _reset_analysis_cache(self):
        self.analysis_cache_key = None
        self.analysis_data = None
        self.analysis_missing_records = None
        self.analysis_statistics_analyzer = None
        self.analysis_results_cache = {}
        self.analysis_time_dynamics_cache = {}
        self.analysis_outliers_cache = {}

    def _ensure_analysis_cache(self):
        cache_key = self._json_cache_key()
        if self.analysis_cache_key != cache_key:
            self.analysis_cache_key = cache_key
            self.analysis_data = None
            self.analysis_missing_records = None
            self.analysis_statistics_analyzer = None
            self.analysis_results_cache = {}
            self.analysis_time_dynamics_cache = {}
            self.analysis_outliers_cache = {}

    def _get_analysis_data(self):
        self._ensure_analysis_cache()
        if self.analysis_data is None:
            self.analysis_data = load_json_data(self.json_file_path)
        return self.analysis_data

    def _get_missing_records_cached(self):
        self._ensure_analysis_cache()
        if self.analysis_missing_records is None:
            self.analysis_missing_records = find_missing_records(self._get_analysis_data())
        return self.analysis_missing_records

    def _get_statistics_analyzer(self):
        self._ensure_analysis_cache()
        if self.analysis_statistics_analyzer is None:
            from statistics_analyzer import StatisticsAnalyzer
            self.analysis_statistics_analyzer = StatisticsAnalyzer(self.json_file_path)
        return self.analysis_statistics_analyzer

    def _get_statistics_result(self, name, factory):
        self._ensure_analysis_cache()
        if name not in self.analysis_results_cache:
            self.analysis_results_cache[name] = factory()
        return self.analysis_results_cache[name]

    def _get_time_dynamics_cached(self, group_by):
        self._ensure_analysis_cache()
        if group_by not in self.analysis_time_dynamics_cache:
            self.analysis_time_dynamics_cache[group_by] = time_dynamics(self._get_analysis_data(), group_by)
        return self.analysis_time_dynamics_cache[group_by]

    def _get_outliers_cached(self, sigma_threshold):
        self._ensure_analysis_cache()
        threshold_key = round(float(sigma_threshold), 3)
        if threshold_key not in self.analysis_outliers_cache:
            from outlier_analyzer import OutlierAnalyzer
            analyzer = OutlierAnalyzer(self.json_file_path, sigma_threshold)
            outliers = analyzer.find_process_outliers()
            too_long = sum(1 for item in outliers if item.is_too_long)
            summary = {
                'total_processes': len(analyzer.processes_data),
                'process_outliers_count': len(outliers),
                'too_long_count': too_long,
                'too_short_count': len(outliers) - too_long,
                'total_action_outliers': sum(len(item.action_outliers) for item in outliers),
                'sigma_threshold': sigma_threshold,
                'process_avg': analyzer.process_avg,
                'process_median': analyzer.process_median,
                'process_std': analyzer.process_std,
                'process_mad': analyzer.process_mad,
            }
            self.analysis_outliers_cache[threshold_key] = (outliers, summary)
        return self.analysis_outliers_cache[threshold_key]

    def _commit_json_changes(self, data, message):
        if not self.json_file_path:
            return
        save_json_data(self.json_file_path, data)
        self.json_result = json.dumps(data, indent=2, ensure_ascii=False)
        self._invalidate_analysis_pages()
        QMessageBox.information(self, 'Данные сохранены', message)
        self.show_quality()

    def _save_json_without_quality_rebuild(self, data, message):
        if not self.json_file_path:
            return
        save_json_data(self.json_file_path, data)
        self.json_result = json.dumps(data, indent=2, ensure_ascii=False)
        self._reset_analysis_cache()
        self.cached_statistics_data = None
        self.cached_graphics_data = None
        self.cached_export_data = None
        self.statistics_widget = None
        self.graphics_widget = None
        self.export_widget = None
        self._clear_layout(self.statistics_layout)
        self._clear_layout(self.graphics_layout)
        self._clear_layout(self.export_layout)
        self._refresh_quality_summary()
        QMessageBox.information(self, 'Данные сохранены', message)

    def _refresh_quality_summary(self):
        if not self.quality_summary_label or not self.json_file_path:
            return
        data = self._get_analysis_data()
        missing_count = len(self._get_missing_records_cached())
        duplicates_count = sum(len(group['indices']) - 1 for group in find_duplicate_records(data))
        try:
            self.quality_summary_label.setText(
                f"Файл: {self.json_file_path}\n"
                f"Записей: {len(data.get('records', []))}. "
                f"Пропусков: {missing_count}. Дублирующих записей: {duplicates_count}."
            )
        except RuntimeError:
            self.quality_summary_label = None

    def show_quality(self):
        if not self.json_file_path:
            QMessageBox.warning(self, 'Ошибка', 'Сначала загрузите или сформируйте JSON файл')
            return

        self.setWindowTitle('Анализ бизнес-процессов — Качество данных')
        self._set_navigation_mode('quality')
        self._build_quality_page()
        self.stacked_widget.setCurrentIndex(1)

    def _build_quality_page(self):
        if self.cached_quality_data == self.json_file_path and self.quality_widget is not None:
            return
        self._clear_layout(self.quality_layout)
        self.quality_widget = self.create_quality_widget()
        self.quality_layout.addWidget(self.quality_widget)
        self.cached_quality_data = self.json_file_path

    def create_quality_widget(self):
        widget = QWidget()
        layout = QVBoxLayout()

        data = self._get_analysis_data()
        missing_count = len(self._get_missing_records_cached())
        duplicates_count = sum(len(group['indices']) - 1 for group in find_duplicate_records(data))

        summary = QLabel(
            f"Файл: {self.json_file_path}\n"
            f"Записей: {len(data.get('records', []))}. "
            f"Пропусков: {missing_count}. Дублирующих записей: {duplicates_count}."
        )
        self.quality_summary_label = summary
        summary.setWordWrap(True)
        layout.addWidget(summary)

        note = QLabel(
            "Автозамена использует контекст поля: дата ставится между ближайшими известными событиями "
            "того же кейса, а действие восстанавливается по полным траекториям процессов с учетом "
            "соседних шагов и непрерывных блоков пропусков."
        )
        note.setWordWrap(True)
        note.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(note)

        tabs = QTabWidget()
        self.quality_tabs = tabs
        tabs.addTab(self.create_missing_values_tab(), "Пропуски")
        tabs.addTab(self.create_duplicates_tab(), "Дубли")
        layout.addWidget(tabs)

        widget.setLayout(layout)
        return widget

    def create_missing_values_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.missing_table = QTableWidget()
        self.missing_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.missing_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.missing_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.missing_table.setColumnCount(6)
        self.missing_table.setHorizontalHeaderLabels([
            "Строка JSON", "Поле", "Текущее значение", "Case", "Действие", "Дата"
        ])
        self.missing_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.missing_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.missing_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.missing_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.missing_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.missing_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)

        buttons = QHBoxLayout()
        btn_replace_selected = QPushButton("Автозаменить выбранные")
        btn_replace_all = QPushButton("Автозаменить все")
        btn_replace_all.setStyleSheet("font-weight: bold;")
        btn_manual = QPushButton("Ввести вручную")
        btn_delete_cases_selected = QPushButton("Удалить кейсы")
        btn_delete_cases_all = QPushButton("Удалить все кейсы с пропусками")
        btn_delete_cases_all.setStyleSheet("font-weight: bold;")

        btn_replace_selected.clicked.connect(lambda: self.replace_missing_values(False))
        btn_replace_all.clicked.connect(lambda: self.replace_missing_values(True))
        btn_manual.clicked.connect(self.fill_missing_manually)
        btn_delete_cases_selected.clicked.connect(lambda: self.delete_missing_cases(False))
        btn_delete_cases_all.clicked.connect(lambda: self.delete_missing_cases(True))

        buttons.addWidget(btn_replace_all)
        buttons.addWidget(btn_replace_selected)
        buttons.addWidget(btn_manual)
        buttons.addStretch()
        buttons.addWidget(btn_delete_cases_all)
        buttons.addWidget(btn_delete_cases_selected)

        layout.addWidget(self.missing_table)
        layout.addLayout(buttons)
        tab.setLayout(layout)
        self.populate_missing_table()
        return tab

    def create_duplicates_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.duplicates_table = QTableWidget()
        self.duplicates_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.duplicates_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.duplicates_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.duplicates_table.setColumnCount(6)
        self.duplicates_table.setHorizontalHeaderLabels([
            "Группа", "Строка JSON", "Статус", "Case", "Действие", "Дата"
        ])
        self.duplicates_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.duplicates_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.duplicates_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.duplicates_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.duplicates_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.duplicates_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)

        buttons = QHBoxLayout()
        btn_delete_selected = QPushButton("Удалить выбранные")
        btn_delete_all = QPushButton("Удалить все дубли")
        btn_delete_selected.clicked.connect(self.delete_selected_duplicates)
        btn_delete_all.clicked.connect(self.delete_all_duplicates)
        buttons.addStretch()
        buttons.addWidget(btn_delete_selected)
        buttons.addWidget(btn_delete_all)

        layout.addWidget(self.duplicates_table)
        layout.addLayout(buttons)
        tab.setLayout(layout)
        self.populate_duplicates_table()
        return tab

    def populate_missing_table(self):
        issues = self._get_missing_records_cached()
        self.missing_table.setRowCount(len(issues))
        for row, issue in enumerate(issues):
            item = QTableWidgetItem(str(issue['record_index']))
            item.setData(Qt.UserRole, (issue['record_index'], issue['field']))
            item.setData(Qt.UserRole + 1, False)
            item.setData(Qt.UserRole + 2, issue['case_id'])
            self.missing_table.setItem(row, 0, item)
            self.missing_table.setItem(row, 1, QTableWidgetItem(issue['field']))
            self.missing_table.setItem(row, 2, QTableWidgetItem(str(issue['current_value'])))
            self.missing_table.setItem(row, 3, QTableWidgetItem(str(issue['case_id'])))
            self.missing_table.setItem(row, 4, QTableWidgetItem(issue['action']))
            self.missing_table.setItem(row, 5, QTableWidgetItem(str(issue['datetime'])))

    def mark_missing_issue_resolved(self, data, issue, value):
        if not hasattr(self, 'missing_table') or self.missing_table is None:
            return

        display_value = str(value)

        for row in range(self.missing_table.rowCount()):
            item = self.missing_table.item(row, 0)
            if not item or item.data(Qt.UserRole) != (issue['record_index'], issue['field']):
                continue

            item.setData(Qt.UserRole, (issue['record_index'], issue['field']))
            item.setData(Qt.UserRole + 1, True)
            item.setData(Qt.UserRole + 2, issue.get('case_id'))
            field_item = self.missing_table.item(row, 1)
            value_item = self.missing_table.item(row, 2)
            if field_item:
                field_item.setText(issue['field'])
            if value_item:
                value_item.setText(display_value)

            for column in range(self.missing_table.columnCount()):
                cell = self.missing_table.item(row, column)
                if cell:
                    cell.setBackground(QColor(220, 245, 220))
            break

    def populate_duplicates_table(self):
        data = load_json_data(self.json_file_path)
        groups = find_duplicate_records(data)
        rows = sum(len(group['indices']) for group in groups)
        self.duplicates_table.setRowCount(rows)
        table_row = 0
        records = data.get('records', [])
        for group_number, group in enumerate(groups, 1):
            for position, record_index in enumerate(group['indices']):
                record = records[record_index]
                item = QTableWidgetItem(str(group_number))
                item.setData(Qt.UserRole, record_index)
                self.duplicates_table.setItem(table_row, 0, item)
                self.duplicates_table.setItem(table_row, 1, QTableWidgetItem(str(record_index)))
                self.duplicates_table.setItem(table_row, 2, QTableWidgetItem("Оригинал" if position == 0 else "Дубль"))
                self.duplicates_table.setItem(table_row, 3, QTableWidgetItem(str(record.get('case_id'))))
                self.duplicates_table.setItem(table_row, 4, QTableWidgetItem(action_name(data, record.get('action_id'))))
                self.duplicates_table.setItem(table_row, 5, QTableWidgetItem(str(record.get('datetime'))))
                table_row += 1

    def selected_missing_issues(self, include_resolved=False):
        rows = sorted(set(index.row() for index in self.missing_table.selectedIndexes()))
        issues = []
        for row in rows:
            item = self.missing_table.item(row, 0)
            if item and item.data(Qt.UserRole):
                if item.data(Qt.UserRole + 1) and not include_resolved:
                    continue
                record_index, field = item.data(Qt.UserRole)
                issues.append({"record_index": record_index, "field": field})
        return issues

    def unresolved_missing_issues_from_table(self):
        issues = []
        if not hasattr(self, 'missing_table') or self.missing_table is None:
            return issues
        for row in range(self.missing_table.rowCount()):
            item = self.missing_table.item(row, 0)
            if not item or not item.data(Qt.UserRole):
                continue
            if item.data(Qt.UserRole + 1):
                continue
            record_index, field = item.data(Qt.UserRole)
            issues.append({"record_index": record_index, "field": field})
        return issues

    def remove_missing_rows_from_table(self, record_indices):
        if not hasattr(self, 'missing_table') or self.missing_table is None:
            return
        to_remove = set(record_indices)
        for row in range(self.missing_table.rowCount() - 1, -1, -1):
            item = self.missing_table.item(row, 0)
            if not item or not item.data(Qt.UserRole):
                continue
            record_index, _ = item.data(Qt.UserRole)
            is_resolved = bool(item.data(Qt.UserRole + 1))
            if record_index in to_remove and not is_resolved:
                self.missing_table.removeRow(row)

    def remove_missing_cases_from_table(self, case_ids):
        if not hasattr(self, 'missing_table') or self.missing_table is None:
            return
        normalized_case_ids = {str(case_id) for case_id in case_ids}
        for row in range(self.missing_table.rowCount() - 1, -1, -1):
            item = self.missing_table.item(row, 0)
            if not item:
                continue
            case_id = item.data(Qt.UserRole + 2)
            if case_id is None:
                case_item = self.missing_table.item(row, 3)
                case_id = case_item.text() if case_item else None
            if str(case_id) in normalized_case_ids:
                self.missing_table.removeRow(row)

    def replace_missing_values(self, replace_all):
        data = load_json_data(self.json_file_path)
        issues = find_missing_records(data) if replace_all else self.selected_missing_issues()
        if not issues:
            QMessageBox.information(self, 'Пропуски', 'Нет выбранных пропусков для замены')
            return
        full_issues = find_missing_records(data)
        issue_map = {(i['record_index'], i['field']): i for i in full_issues}
        action_replacements = suggest_action_replacements(
            data,
            [issue['record_index'] for issue in issues if issue['field'] == 'action_id']
        )
        changed = 0
        replacements = []
        for issue in issues:
            full_issue = issue_map.get((issue['record_index'], issue['field']), issue)
            if issue['field'] == 'action_id':
                value = action_replacements.get(issue['record_index'], "")
            else:
                value = most_likely_value(data, full_issue)
            if value == "":
                continue
            set_record_value(data, issue['record_index'], issue['field'], value)
            actual_value = data['records'][issue['record_index']].get(issue['field'])
            replacements.append((full_issue, actual_value))
            changed += 1
        if not changed:
            QMessageBox.information(
                self,
                'Автозамена',
                'Для выбранных пропусков не нашлось достаточно надежного контекста для автозамены.'
            )
            return
        self._save_json_without_quality_rebuild(data, f"Заменено пропусков: {changed}")
        for issue, value in replacements:
            self.mark_missing_issue_resolved(data, issue, value)

    def fill_missing_manually(self):
        data = load_json_data(self.json_file_path)
        issues = self.selected_missing_issues(include_resolved=True)
        if not issues:
            QMessageBox.information(self, 'Пропуски', 'Выберите одну или несколько ячеек с пропусками')
            return
        changed = 0
        replacements = []
        for issue in issues:
            value, ok = QInputDialog.getText(
                self,
                'Ручное значение',
                f"Строка {issue['record_index']}, поле {issue['field']}:"
            )
            if ok:
                set_record_value(data, issue['record_index'], issue['field'], value)
                actual_value = data['records'][issue['record_index']].get(issue['field'])
                replacements.append((issue, actual_value))
                changed += 1
        if changed:
            self._save_json_without_quality_rebuild(data, f"Вручную заполнено значений: {changed}")
            for issue, value in replacements:
                self.mark_missing_issue_resolved(data, issue, value)

    def has_resolved_missing_rows(self):
        if not hasattr(self, 'missing_table') or self.missing_table is None:
            return False
        for row in range(self.missing_table.rowCount()):
            item = self.missing_table.item(row, 0)
            if item and item.data(Qt.UserRole + 1):
                return True
        return False

    def delete_missing_records(self, delete_all):
        data = load_json_data(self.json_file_path)
        issues = self.unresolved_missing_issues_from_table() if delete_all else self.selected_missing_issues()
        if delete_all and not issues:
            issues = find_missing_records(data)
        indices = sorted(set(issue['record_index'] for issue in issues), reverse=True)
        if not indices:
            QMessageBox.information(self, 'Пропуски', 'Нет строк для удаления')
            return
        delete_records(data, indices)
        if delete_all:
            self._save_json_without_quality_rebuild(data, f"Удалено строк с пропусками: {len(indices)}")
            self.remove_missing_rows_from_table(indices)
        else:
            self._commit_json_changes(data, f"Удалено строк с пропусками: {len(indices)}")

    def delete_missing_cases(self, delete_all):
        if delete_all and not self.has_resolved_missing_rows():
            result = QMessageBox.warning(
                self,
                'Сначала попробуйте автозамену',
                'Перед удалением всех кейсов с пропусками стоит сначала попробовать автозаменить значения.\n\n'
                'Продолжить удаление кейсов?',
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if result != QMessageBox.Yes:
                return

        data = load_json_data(self.json_file_path)
        issues = self.unresolved_missing_issues_from_table() if delete_all else self.selected_missing_issues()
        if delete_all and not issues:
            issues = find_missing_records(data)
        records = data.get("records", [])
        case_ids = {
            records[issue['record_index']].get('case_id')
            for issue in issues
            if 0 <= issue.get('record_index', -1) < len(records)
        }
        case_ids.discard(None)
        if not case_ids:
            QMessageBox.information(self, 'Пропуски', 'Нет кейсов для удаления')
            return
        original_count = len(records)
        data["records"] = [
            record for record in records
            if record.get("case_id") not in case_ids
        ]
        deleted_rows = original_count - len(data["records"])
        message = f"Удалено кейсов с пропусками: {len(case_ids)}\nУдалено строк: {deleted_rows}"
        if delete_all:
            self._save_json_without_quality_rebuild(data, message)
            self.remove_missing_cases_from_table(case_ids)
        else:
            self._commit_json_changes(data, message)

    def selected_duplicate_indices(self):
        rows = sorted(set(index.row() for index in self.duplicates_table.selectedIndexes()))
        indices = []
        for row in rows:
            item = self.duplicates_table.item(row, 0)
            if item:
                indices.append(item.data(Qt.UserRole))
        return indices

    def delete_selected_duplicates(self):
        data = load_json_data(self.json_file_path)
        indices = sorted(set(self.selected_duplicate_indices()), reverse=True)
        if not indices:
            QMessageBox.information(self, 'Дубли', 'Выберите дублирующие записи для удаления')
            return
        delete_records(data, indices)
        self._save_json_without_quality_rebuild(data, f"Удалено выбранных записей: {len(indices)}")
        self.populate_duplicates_table()

    def delete_all_duplicates(self):
        data = load_json_data(self.json_file_path)
        indices = sorted(duplicate_indices_except_first(data), reverse=True)
        if not indices:
            QMessageBox.information(self, 'Дубли', 'Дубли не найдены')
            return
        delete_records(data, indices)
        self._save_json_without_quality_rebuild(data, f"Удалено дублей: {len(indices)}")
        self.populate_duplicates_table()

    def create_statistics_widget(self):
        """Создает виджет статистики"""
        # Создаем основной виджет
        widget = QWidget()
        layout = QVBoxLayout()

        # Создаем вкладки для разных разделов статистики
        tab_widget = QTabWidget()
        self.statistics_tab_widget = tab_widget

        # Вкладка "Общая статистика"
        general_tab = self.create_general_statistics_tab()
        tab_widget.addTab(general_tab, "Общая статистика")

        dynamics_tab = self.create_time_dynamics_tab()
        tab_widget.addTab(dynamics_tab, "Временная динамика")

        case_tab = self.create_case_detail_tab()
        self.case_detail_tab_index = tab_widget.addTab(case_tab, "Детализация кейса")

        actions_tab = self.create_actions_statistics_tab()
        tab_widget.addTab(actions_tab, "Статистика действий")

        # Вкладка "Переходы"
        transitions_tab = self.create_transitions_statistics_tab()
        tab_widget.addTab(transitions_tab, "Переходы")

        # Вкладка "Поиск сбоев"
        failed_tab = self.create_failed_processes_tab()
        tab_widget.addTab(failed_tab, "Поиск сбоев")

        # Вкладка "Поиск выбросов"
        outliers_tab = self.create_outliers_tab()
        tab_widget.addTab(outliers_tab, "Поиск выбросов")

        layout.addWidget(tab_widget)
        widget.setLayout(layout)

        return widget

    def create_general_statistics_tab(self):
        """Создает вкладку общей статистики"""
        widget = QWidget()
        layout = QVBoxLayout()

        # Создаем анализатор
        analyzer = self._get_statistics_analyzer()

        # Основные метрики процессов
        process_stats = self._get_statistics_result(
            'process_statistics',
            analyzer.get_process_statistics
        )

        if process_stats:
            # Группа основных метрик
            metrics_group = QGroupBox("Основные метрики процессов")
            metrics_layout = QVBoxLayout()

            metrics_layout.addWidget(QLabel(f"Всего процессов: {process_stats['total_processes']}"))
            metrics_layout.addWidget(QLabel(f"Всего действий: {process_stats['total_actions']}"))
            metrics_layout.addWidget(QLabel(f"Средняя длительность: {process_stats['avg_process_duration']}"))
            metrics_layout.addWidget(QLabel(f"Минимальная длительность: {process_stats['min_process_duration']}"))
            metrics_layout.addWidget(QLabel(f"Максимальная длительность: {process_stats['max_process_duration']}"))
            metrics_layout.addWidget(QLabel(f"Стандартное отклонение длительности: {process_stats['std_process_duration']}"))
            metrics_layout.addWidget(QLabel(f"Медианное отклонение длительности: {process_stats['mad_process_duration']}"))
            metrics_layout.addWidget(QLabel(f"Среднее количество действий: {process_stats['avg_actions_per_process']}"))
            metrics_layout.addWidget(QLabel(f"Стандартное отклонение кол-ва действий: {process_stats['std_actions_per_process']}"))
            metrics_layout.addWidget(QLabel(f"Медианное отклонение кол-ва действий: {process_stats['mad_actions_per_process']}"))

            metrics_group.setLayout(metrics_layout)
            layout.addWidget(metrics_group)

        # Распределение по времени
        time_dist = self._get_statistics_result(
            'time_distribution',
            analyzer.get_time_distribution
        )
        if time_dist:
            time_group = QGroupBox("Распределение по времени выполнения")
            time_layout = QVBoxLayout()

            for interval, data in time_dist.items():
                time_layout.addWidget(QLabel(f"{interval}: {data['count']} процессов ({data['percentage']}%)"))

            time_group.setLayout(time_layout)
            layout.addWidget(time_group)

        layout.addStretch()
        widget.setLayout(layout)
        return widget

    def create_time_dynamics_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()

        controls = QHBoxLayout()
        controls.addWidget(QLabel("Группировка:"))
        combo = QComboBox()
        combo.addItem("По месяцам", "month")
        combo.addItem("По неделям", "week")
        combo.addItem("По дням", "day")
        self.stats_time_dynamics_combo = combo
        controls.addWidget(combo)
        controls.addStretch()

        table = QTreeWidget()
        headers = ["Период", "Стартовало", "Без сбоев", "Сбоев", "% сбоев", "Средняя длительность"]
        table.setHeaderLabels(headers)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setAlternatingRowColors(True)
        table.header().setSectionResizeMode(0, QHeaderView.Stretch)
        for column in range(1, len(headers)):
            table.header().setSectionResizeMode(column, QHeaderView.ResizeToContents)

        def refresh_dynamics(*args):
            rows = self._get_time_dynamics_cached(combo.currentData())
            table.clear()
            for item in rows:
                period_item = QTreeWidgetItem([
                    item['period'],
                    str(item['started']),
                    str(item['completed']),
                    str(item['failed']),
                    str(item['failed_rate']),
                    item['avg_duration'],
                ])
                for case in item.get('cases', []):
                    case_id = case.get('case_id')
                    failed = bool(case.get('failed'))
                    case_item = QTreeWidgetItem(period_item, [
                        f"Case_{case_id} - {'сбой' if failed else 'без сбоя'}",
                        "",
                        "",
                        "Да" if failed else "Нет",
                        "",
                        "",
                    ])
                    case_item.setData(0, Qt.UserRole, case_id)
                table.addTopLevelItem(period_item)

        table.itemDoubleClicked.connect(self.open_case_from_tree_item)
        combo.currentIndexChanged.connect(refresh_dynamics)
        layout.addLayout(controls)
        layout.addWidget(table)
        widget.setLayout(layout)
        refresh_dynamics()
        return widget

    def create_case_detail_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()

        controls = QHBoxLayout()
        controls.addWidget(QLabel("Case ID:"))
        case_input = QLineEdit()
        case_input.setPlaceholderText("Например: 123 или Case_123")
        btn_search = QPushButton("Найти")
        controls.addWidget(case_input)
        controls.addWidget(btn_search)

        table = QTableWidget()
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        headers = ["Шаг", "Дата", "Действие", "Время с прошлого шага", "Строка JSON"]
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)

        timeline_scene = QGraphicsScene()
        timeline_view = self._create_scene_view(timeline_scene, minimum_height=250)

        def draw_timeline(rows):
            timeline_scene.clear()
            if not rows:
                self._draw_empty_scene(timeline_scene, "Введите Case ID и нажмите «Найти», чтобы построить таймлайн")
                self._set_scene_content(timeline_view, timeline_scene)
                return

            parsed_rows = []
            for item in rows:
                dt = parse_dt(item.get("datetime"))
                parsed_rows.append({**item, "dt": dt})

            dated_rows = [item for item in parsed_rows if item["dt"] is not None]
            left, top = 80, 70
            width = max(920, len(parsed_rows) * 160)
            baseline_y = top + 88
            node_radius = 13

            case_label = case_input.text().strip() or "Case"
            self._add_scene_text(timeline_scene, f"Таймлайн {case_label}", left, 0, size=13, bold=True)
            self._add_scene_text(
                timeline_scene,
                "Промежуток красный, если длится дольше среднего для этого действия по файлу; синий, если короче среднего.",
                left,
                22,
                color=QColor(90, 95, 100),
                max_width=900,
            )

            if dated_rows:
                min_dt = dated_rows[0]["dt"]
                max_dt = dated_rows[-1]["dt"]
                span_seconds = max(1, (max_dt - min_dt).total_seconds())
            else:
                min_dt = None
                max_dt = None
                span_seconds = max(1, len(parsed_rows) - 1)

            try:
                action_avg_seconds = {
                    action: metrics.avg_duration.total_seconds()
                    for action, metrics in self._get_statistics_analyzer().action_metrics.items()
                }
            except Exception:
                action_avg_seconds = {}

            points = []
            for index, item in enumerate(parsed_rows):
                if min_dt is not None and item["dt"] is not None:
                    x = left + ((item["dt"] - min_dt).total_seconds() / span_seconds) * width
                else:
                    x = left + (width * index / max(1, len(parsed_rows) - 1))
                points.append((x, baseline_y, item))

            previous_point = None
            previous_dt = None
            for point in points:
                x, y, item = point
                current_dt = item["dt"]
                if previous_point is not None:
                    interval = None
                    if current_dt is not None and previous_dt is not None:
                        interval = max(0, (current_dt - previous_dt).total_seconds())
                    previous_action = previous_point[2].get("action", "")
                    avg_interval = action_avg_seconds.get(previous_action)
                    is_longer_than_avg = interval is not None and avg_interval is not None and interval > avg_interval
                    is_shorter_than_avg = interval is not None and avg_interval is not None and interval < avg_interval
                    line_color = QColor(190, 85, 70) if is_longer_than_avg else QColor(66, 116, 175) if is_shorter_than_avg else QColor(120, 125, 130)
                    label_color = QColor(170, 65, 55) if is_longer_than_avg else QColor(45, 90, 155) if is_shorter_than_avg else QColor(70, 80, 90)
                    pen = QPen(line_color, 4 if is_longer_than_avg else 3)
                    timeline_scene.addLine(previous_point[0], baseline_y, x, baseline_y, pen)
                    if interval is not None:
                        mid_x = (previous_point[0] + x) / 2
                        label = self._duration_label_from_seconds(interval)
                        label_item = timeline_scene.addText(label, QFont("Arial", 8))
                        label_item.setDefaultTextColor(label_color)
                        label_item.setPos(mid_x - label_item.boundingRect().width() / 2, baseline_y - 42)
                previous_point = point
                if current_dt is not None:
                    previous_dt = current_dt

            for index, (x, y, item) in enumerate(points):
                fill = QColor(74, 130, 180) if index not in {0, len(points) - 1} else QColor(58, 145, 98)
                if index == len(points) - 1:
                    fill = QColor(182, 92, 77)
                circle = timeline_scene.addEllipse(
                    x - node_radius,
                    y - node_radius,
                    node_radius * 2,
                    node_radius * 2,
                    QPen(QColor(255, 255, 255), 2),
                    QBrush(fill),
                )
                circle.setToolTip(f"{item.get('datetime', '')}\n{item.get('action', '')}")

                step_text = timeline_scene.addText(str(item.get("order", index + 1)), QFont("Arial", 8))
                step_text.setDefaultTextColor(QColor(255, 255, 255))
                step_text.setPos(x - step_text.boundingRect().width() / 2, y - step_text.boundingRect().height() / 2)

                action_label = item.get("action", "")
                label = timeline_scene.addText(action_label, self._angled_label_font(8))
                label.setDefaultTextColor(QColor(35, 45, 55))
                label.setTransformOriginPoint(label.boundingRect().left(), label.boundingRect().top())
                label.setRotation(28)
                label.setPos(x + 18, y + 42)

                if item.get("datetime"):
                    date_label = timeline_scene.addText(str(item.get("datetime"))[:16], self._angled_label_font(7))
                    date_label.setDefaultTextColor(QColor(85, 90, 95))
                    date_label.setRotation(-25)
                    date_label.setPos(x - 42, y - 74)

            if min_dt is not None and max_dt is not None:
                self._add_scene_text(timeline_scene, min_dt.strftime("%Y-%m-%d %H:%M"), left, baseline_y + 230, size=8)
                self._add_scene_text(timeline_scene, max_dt.strftime("%Y-%m-%d %H:%M"), left + width - 110, baseline_y + 230, size=8)

            self._set_scene_content(timeline_view, timeline_scene)

        def refresh_case(*args):
            data = self._get_analysis_data()
            rows = find_case_timeline(data, case_input.text())
            table.setRowCount(len(rows))
            for row, item in enumerate(rows):
                table.setItem(row, 0, QTableWidgetItem(str(item['order'])))
                table.setItem(row, 1, QTableWidgetItem(str(item['datetime'])))
                table.setItem(row, 2, QTableWidgetItem(item['action']))
                table.setItem(row, 3, QTableWidgetItem(item['from_previous']))
                table.setItem(row, 4, QTableWidgetItem(str(item['record_index'])))
            draw_timeline(rows)

        btn_search.clicked.connect(refresh_case)
        case_input.returnPressed.connect(refresh_case)
        self.case_detail_input = case_input
        self.case_detail_refresh = refresh_case
        layout.addLayout(controls)

        splitter = QSplitter(Qt.Vertical)
        splitter.addWidget(table)
        splitter.addWidget(timeline_view)
        splitter.setSizes([280, 260])
        layout.addWidget(splitter)
        widget.setLayout(layout)
        draw_timeline([])
        return widget

    def open_case_detail(self, case_id):
        if case_id is None:
            return
        self.show_statistics()
        if self.statistics_tab_widget is None or self.case_detail_tab_index is None:
            return
        self.statistics_tab_widget.setCurrentIndex(self.case_detail_tab_index)
        if self.case_detail_input is not None:
            self.case_detail_input.setText(str(case_id))
        if self.case_detail_refresh is not None:
            self.case_detail_refresh()

    def open_case_from_tree_item(self, item, column):
        case_id = item.data(0, Qt.UserRole)
        if case_id is None:
            case_id = item.data(3, Qt.UserRole)
        if case_id is not None:
            self.open_case_detail(case_id)

    def open_case_from_table_item(self, item):
        if item is None:
            return
        case_id = item.data(Qt.UserRole)
        if case_id is None:
            table = item.tableWidget()
            case_item = table.item(item.row(), 0) if table is not None else None
            case_id = case_item.text() if case_item is not None else None
        if case_id is not None:
            self.open_case_detail(case_id)

    def create_actions_statistics_tab(self):
        """Создает вкладку статистики действий"""
        widget = QWidget()
        layout = QVBoxLayout()

        # Создаем таблицу действий
        table = QTableWidget()
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)

        # Заголовки таблицы
        headers = ["Действие", "Выполнений", "Процессов", "Ср. длительность", "Ст. откл.", "Мед. откл.", "Мин. время", "Макс. время"]
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)

        # Настройка размеров колонок
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)  # Действие
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Выполнений
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Процессов
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Ср. длительность
        table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Ст. откл.
        table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Мед. откл.
        table.horizontalHeader().setSectionResizeMode(6, QHeaderView.ResizeToContents)  # Мин. время
        table.horizontalHeader().setSectionResizeMode(7, QHeaderView.ResizeToContents)  # Макс. время

        # Создаем анализатор и заполняем таблицу
        analyzer = self._get_statistics_analyzer()
        action_stats = self._get_statistics_result(
            'action_statistics',
            analyzer.get_action_statistics
        )

        table.setRowCount(len(action_stats))

        for row, stat in enumerate(action_stats):
            table.setItem(row, 0, QTableWidgetItem(stat['action']))
            table.setItem(row, 1, QTableWidgetItem(str(stat['occurrences'])))
            table.setItem(row, 2, QTableWidgetItem(str(stat['processes'])))
            table.setItem(row, 3, QTableWidgetItem(stat['avg_duration']))
            table.setItem(row, 4, QTableWidgetItem(stat['duration_std']))
            table.setItem(row, 5, QTableWidgetItem(stat['duration_mad']))
            table.setItem(row, 6, QTableWidgetItem(stat['min_duration']))
            table.setItem(row, 7, QTableWidgetItem(stat['max_duration']))

        layout.addWidget(table)
        widget.setLayout(layout)
        return widget

    def create_transitions_statistics_tab(self):
        """Создает вкладку статистики переходов"""
        widget = QWidget()
        layout = QVBoxLayout()

        # Создаем анализатор
        analyzer = self._get_statistics_analyzer()
        flow_analysis = self._get_statistics_result(
            'process_flow_analysis',
            analyzer.get_process_flow_analysis
        )

        # Топ действий
        if flow_analysis.get('action_frequency'):
            actions_group = QGroupBox("Топ действий по частоте")
            actions_layout = QVBoxLayout()

            for i, action in enumerate(flow_analysis['action_frequency'][:10], 1):
                actions_layout.addWidget(QLabel(f"{i}. {action['action']}: {action['count']} раз ({action['percentage']}%)"))

            actions_group.setLayout(actions_layout)
            layout.addWidget(actions_group)

        # Топ переходов
        if flow_analysis.get('top_transitions'):
            transitions_group = QGroupBox("Наиболее частые переходы")
            transitions_layout = QVBoxLayout()

            for i, transition in enumerate(flow_analysis['top_transitions'][:10], 1):
                transitions_layout.addWidget(QLabel(f"{i}. {transition['from']} → {transition['to']}: {transition['count']} раз ({transition['percentage']}%)"))

            transitions_group.setLayout(transitions_layout)
            layout.addWidget(transitions_group)

        layout.addStretch()
        widget.setLayout(layout)
        return widget

    def create_failed_processes_tab(self):
        """Создает вкладку поиска незавершенных процессов"""
        widget = QWidget()
        layout = QVBoxLayout()

        # Создаем анализатор и получаем список незавершенных процессов
        analyzer = self._get_statistics_analyzer()
        failed_processes = self._get_statistics_result(
            'failed_processes',
            analyzer.find_failed_processes
        )

        # Группа с информацией о сбоях
        info_group = QGroupBox(f"Найдено незавершенных процессов: {len(failed_processes)}")
        info_layout = QVBoxLayout()

        if failed_processes:
            # Создаем таблицу для отображения результатов
            table = QTableWidget()
            table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            table.setColumnCount(6)
            table.setHorizontalHeaderLabels([
                "ID процесса", "Дата начала", "Дата последнего действия",
                "Последнее действие", "Ожидаемое действие", "Количество действий"
            ])

            # Настройка размеров колонок
            table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)  # ID процесса
            table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)  # Дата начала
            table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)  # Дата последнего
            table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)  # Действие
            table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)  # Ожидаемое действие
            table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Количество

            table.setRowCount(len(failed_processes))

            for row, process in enumerate(failed_processes):
                values = [
                    str(process['case_id']),
                    process['start_date'],
                    process['last_date'],
                    process['last_action'],
                    process['correct_end_action'],
                    str(process['actions_count']),
                ]
                for column, value in enumerate(values):
                    item = QTableWidgetItem(value)
                    item.setData(Qt.UserRole, process['case_id'])
                    table.setItem(row, column, item)

            table.itemDoubleClicked.connect(self.open_case_from_table_item)

            info_layout.addWidget(table)

            # Дополнительная информация
            summary_label = QLabel(
                f"Эти процессы остановились на действиях, которые не являются общепринятыми конечными действиями.\n"
                f"Возможные причины: прерывание процесса, ошибка в данных, незавершенные операции."
            )
            summary_label.setWordWrap(True)
            summary_label.setStyleSheet("color: #666; font-style: italic;")
            info_layout.addWidget(summary_label)

        else:
            # Если нет незавершенных процессов
            no_failed_label = QLabel(
                "Все процессы завершились общепринятыми конечными действиями.\n"
                "Незавершенных процессов не найдено."
            )
            no_failed_label.setWordWrap(True)
            no_failed_label.setStyleSheet("color: #008000; font-weight: bold;")
            info_layout.addWidget(no_failed_label)

        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        widget.setLayout(layout)
        return widget

    def create_outliers_tab(self):
        """Создает вкладку поиска выбросов (аномалий)"""
        widget = QWidget()
        layout = QVBoxLayout()

        # Панель настроек порога
        settings_group = QGroupBox("Настройки поиска выбросов")
        settings_layout = QHBoxLayout()

        settings_layout.addWidget(QLabel("Порог отклонения (σ):"))
        
        self.sigma_spinbox = QDoubleSpinBox()
        self.sigma_spinbox.setRange(1.0, 5.0)
        self.sigma_spinbox.setValue(3.0)
        self.sigma_spinbox.setSingleStep(0.5)
        self.sigma_spinbox.setDecimals(1)
        self.sigma_spinbox.setMaximumWidth(80)
        settings_layout.addWidget(self.sigma_spinbox)

        self.btn_refresh_outliers = QPushButton("Обновить")
        self.btn_refresh_outliers.setMaximumWidth(100)
        self.btn_refresh_outliers.clicked.connect(self.refresh_outliers_tree)
        settings_layout.addWidget(self.btn_refresh_outliers)

        settings_layout.addStretch()
        settings_group.setLayout(settings_layout)
        layout.addWidget(settings_group)

        # Статистика выбросов
        self.outliers_stats_group = QGroupBox("Статистика выбросов")
        self.outliers_stats_layout = QVBoxLayout()
        self.outliers_stats_group.setLayout(self.outliers_stats_layout)
        layout.addWidget(self.outliers_stats_group)

        # Дерево процессов-выбросов
        outliers_group = QGroupBox("Процессы-выбросы")
        outliers_layout = QVBoxLayout()

        self.outliers_tree = QTreeWidget()
        self.outliers_tree.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.outliers_tree.setHeaderLabels([
            "Процесс",
            "Длительность", 
            "Среднее",
            "Ст. откл.",
            "Мед. откл.",
            "Отклонение",
            "Тип"
        ])
        self.outliers_tree.itemDoubleClicked.connect(self.open_case_from_tree_item)
        
        # Настройка размеров колонок
        self.outliers_tree.setColumnWidth(0, 200)
        self.outliers_tree.setColumnWidth(1, 100)
        self.outliers_tree.setColumnWidth(2, 100)
        self.outliers_tree.setColumnWidth(3, 100)
        self.outliers_tree.setColumnWidth(4, 100)
        self.outliers_tree.setColumnWidth(5, 150)
        self.outliers_tree.setColumnWidth(6, 100)

        outliers_layout.addWidget(self.outliers_tree)
        outliers_group.setLayout(outliers_layout)
        layout.addWidget(outliers_group)

        # Заполняем данными
        self._populate_outliers_tree(3.0)

        widget.setLayout(layout)
        return widget

    def _populate_outliers_tree(self, sigma_threshold: float):
        """Заполняет дерево выбросов данными"""
        from outlier_analyzer import format_timedelta, format_deviation

        self.outliers_tree.clear()
        
        # Очищаем старую статистику
        while self.outliers_stats_layout.count():
            item = self.outliers_stats_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        try:
            outliers, summary = self._get_outliers_cached(sigma_threshold)

            # Обновляем статистику
            self.outliers_stats_layout.addWidget(QLabel(
                f"Всего процессов: {summary['total_processes']}"
            ))
            self.outliers_stats_layout.addWidget(QLabel(
                f"Найдено процессов-выбросов: {summary['process_outliers_count']} "
                f"(слишком долгих: {summary['too_long_count']}, слишком быстрых: {summary['too_short_count']})"
            ))
            self.outliers_stats_layout.addWidget(QLabel(
                f"Всего действий-выбросов: {summary['total_action_outliers']}"
            ))
            self.outliers_stats_layout.addWidget(QLabel(
                f"Норма процессов — среднее: {format_timedelta(summary['process_avg'])}, "
                f"медиана: {format_timedelta(summary['process_median'])}"
            ))
            self.outliers_stats_layout.addWidget(QLabel(
                f"Ст. откл.: {format_timedelta(summary['process_std'])}, "
                f"Мед. откл.: {format_timedelta(summary['process_mad'])}"
            ))
            self.outliers_stats_layout.addWidget(QLabel(
                f"Порог: {sigma_threshold}σ"
            ))

            # Заполняем дерево
            for proc_outlier in outliers:
                # Создаем элемент процесса
                proc_item = QTreeWidgetItem()
                proc_item.setText(0, f"Case_{proc_outlier.case_id}")
                proc_item.setData(0, Qt.UserRole, proc_outlier.case_id)
                proc_item.setText(1, format_timedelta(proc_outlier.duration))
                proc_item.setText(2, format_timedelta(proc_outlier.avg_duration))
                proc_item.setText(3, format_timedelta(proc_outlier.std_duration))
                proc_item.setText(4, format_timedelta(proc_outlier.mad_duration))
                proc_item.setText(5, format_deviation(proc_outlier.deviation_abs, proc_outlier.deviation_percent))
                
                type_text = "⬆️ Долгий" if proc_outlier.is_too_long else "⬇️ Быстрый"
                proc_item.setText(6, type_text)

                # Устанавливаем цвет фона в зависимости от типа
                if proc_outlier.is_too_long:
                    for col in range(7):
                        proc_item.setBackground(col, QColor(255, 200, 200))  # Светло-красный
                else:
                    for col in range(7):
                        proc_item.setBackground(col, QColor(200, 200, 255))  # Светло-синий

                # Добавляем действия-выбросы как дочерние элементы
                if proc_outlier.action_outliers:
                    for action_outlier in proc_outlier.action_outliers:
                        action_item = QTreeWidgetItem(proc_item)
                        action_item.setText(0, f"  └ {action_outlier.action_name}")
                        action_item.setData(0, Qt.UserRole, proc_outlier.case_id)
                        action_item.setText(1, format_timedelta(action_outlier.duration))
                        action_item.setText(2, format_timedelta(action_outlier.avg_duration))
                        action_item.setText(3, format_timedelta(action_outlier.std_duration))
                        action_item.setText(4, format_timedelta(action_outlier.mad_duration))
                        action_item.setText(5, format_deviation(action_outlier.deviation_abs, action_outlier.deviation_percent))
                        
                        action_type = "⬆️ Долгий" if action_outlier.is_too_long else "⬇️ Быстрый"
                        action_item.setText(6, action_type)

                        # Более светлый оттенок для действий
                        if action_outlier.is_too_long:
                            for col in range(7):
                                action_item.setBackground(col, QColor(255, 230, 230))
                        else:
                            for col in range(7):
                                action_item.setBackground(col, QColor(230, 230, 255))

                self.outliers_tree.addTopLevelItem(proc_item)

            # Если нет выбросов
            if not outliers:
                no_outliers_item = QTreeWidgetItem()
                no_outliers_item.setText(0, "Выбросов не обнаружено")
                no_outliers_item.setForeground(0, QColor(0, 128, 0))
                self.outliers_tree.addTopLevelItem(no_outliers_item)

        except Exception as e:
            error_item = QTreeWidgetItem()
            error_item.setText(0, f"Ошибка анализа: {str(e)}")
            error_item.setForeground(0, QColor(255, 0, 0))
            self.outliers_tree.addTopLevelItem(error_item)

    def refresh_outliers_tree(self):
        """Обновляет дерево выбросов с новым порогом"""
        sigma_threshold = self.sigma_spinbox.value()
        self._populate_outliers_tree(sigma_threshold)

    def show_graphics(self):
        if not self.json_file_path:
            QMessageBox.warning(self, 'Ошибка', 'Нет данных для построения графики')
            return

        self.setWindowTitle('Анализ бизнес-процессов — Графика')
        self._set_navigation_mode('graphics')
        self._build_graphics_page()
        self.stacked_widget.setCurrentIndex(3)

    def show_export(self):
        if not self.json_file_path:
            QMessageBox.warning(self, 'Ошибка', 'Нет данных для экспорта')
            return

        self.setWindowTitle('Анализ бизнес-процессов — Экспорт')
        self._set_navigation_mode('export')
        self._build_export_page()
        self.stacked_widget.setCurrentIndex(4)

    def _build_export_page(self):
        current_cache_key = self._json_cache_key()
        if self.cached_export_data == current_cache_key and self.export_widget is not None:
            return
        self._clear_layout(self.export_layout)
        self.export_widget = self.create_export_widget()
        self.export_layout.addWidget(self.export_widget)
        self.cached_export_data = current_cache_key

    def create_export_widget(self):
        widget = QWidget()
        layout = QVBoxLayout()
        self.export_checkboxes = {}
        self.export_controls = {}

        title = QLabel("Экспорт отчета в PDF")
        title.setFont(QFont("Arial", 14, QFont.Bold))
        layout.addWidget(title)

        description = QLabel(
            "Выберите вкладки, которые нужно напечатать в отчет. Параметры по умолчанию "
            "совпадают с настройками соответствующих вкладок, их можно изменить только для отчета."
        )
        description.setWordWrap(True)
        description.setStyleSheet("color: #666;")
        layout.addWidget(description)

        title_row = QHBoxLayout()
        title_row.addWidget(QLabel("Название отчета:"))
        report_title = QLineEdit("Отчет по бизнес-процессу")
        title_row.addWidget(report_title)
        layout.addLayout(title_row)
        self.export_controls["report_title"] = report_title

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content = QWidget()
        content_layout = QVBoxLayout()

        data = self._get_analysis_data()
        min_process_date, max_process_date = self._process_date_bounds(data)
        min_qdate = QDate(min_process_date.year, min_process_date.month, min_process_date.day)
        max_qdate = QDate(max_process_date.year, max_process_date.month, max_process_date.day)

        self._add_export_option(content_layout, "general_statistics", "Общая статистика")

        time_controls = QHBoxLayout()
        time_combo = self._create_grouping_combo(self._current_combo_data("stats_time_dynamics_combo", "month"))
        time_controls.addWidget(QLabel("Периодичность:"))
        time_controls.addWidget(time_combo)
        time_controls.addStretch()
        self.export_controls["time_statistics_group"] = time_combo
        self._add_export_option(content_layout, "time_statistics", "Временная статистика", time_controls)

        self._add_export_option(content_layout, "action_statistics", "Статистика действий")
        self._add_export_option(content_layout, "transitions", "Переходы")
        self._add_export_option(content_layout, "failed_processes", "Поиск сбоев")

        outlier_controls = QHBoxLayout()
        outlier_sigma = self._create_sigma_spinbox(self._current_spinbox_value("sigma_spinbox", 3.0))
        outlier_controls.addWidget(QLabel("Порог отклонения (σ):"))
        outlier_controls.addWidget(outlier_sigma)
        outlier_controls.addStretch()
        self.export_controls["outliers_sigma"] = outlier_sigma
        self._add_export_option(content_layout, "outliers", "Поиск выбросов", outlier_controls)

        map_controls = QHBoxLayout()
        map_min = QDoubleSpinBox()
        map_min.setRange(0, 1000000)
        map_min.setSingleStep(1)
        map_min.setDecimals(0)
        map_min.setValue(self._current_spinbox_value("process_map_min_transitions_spinbox", 0))
        map_min.setMaximumWidth(90)
        map_controls.addWidget(QLabel("Связи от"))
        map_controls.addWidget(map_min)
        map_controls.addWidget(QLabel("переходов"))
        map_controls.addStretch()
        self.export_controls["process_map_min"] = map_min
        self._add_export_option(content_layout, "process_map", "Карта процессов", map_controls)

        self._add_export_option(content_layout, "transition_matrix", "Матрица переходов")

        dynamics_controls = QHBoxLayout()
        dynamics_combo = self._create_grouping_combo(self._current_combo_data("graphics_dynamics_combo", "month"))
        date_from = self._create_export_date_edit(min_qdate, max_qdate, self._current_date_edit_value("graphics_dynamics_date_from", min_qdate))
        date_to = self._create_export_date_edit(min_qdate, max_qdate, self._current_date_edit_value("graphics_dynamics_date_to", max_qdate))
        dynamics_controls.addWidget(QLabel("Группировка:"))
        dynamics_controls.addWidget(dynamics_combo)
        dynamics_controls.addSpacing(12)
        dynamics_controls.addWidget(QLabel("с"))
        dynamics_controls.addWidget(date_from)
        dynamics_controls.addWidget(QLabel("по"))
        dynamics_controls.addWidget(date_to)
        dynamics_controls.addStretch()
        self.export_controls["dynamics_group"] = dynamics_combo
        self.export_controls["dynamics_date_from"] = date_from
        self.export_controls["dynamics_date_to"] = date_to
        dynamics_warning = QLabel("Внимание: при выборе больше 30 дней или недель график в PDF может получиться мелким")
        dynamics_warning.setStyleSheet(
            "color: #8a5a00; background: #fff3cd; border: 1px solid #f0d98c; "
            "border-radius: 4px; padding: 3px 6px;"
        )
        dynamics_warning.setVisible(False)
        self._add_export_option(
            content_layout,
            "dynamics",
            "Динамика",
            dynamics_controls,
            checkbox_suffix=dynamics_warning,
        )

        def update_dynamics_warning(*args):
            group_by = dynamics_combo.currentData()
            days = abs(date_from.date().daysTo(date_to.date())) + 1
            periods = days if group_by == "day" else math.ceil(days / 7) if group_by == "week" else 0
            checked = self.export_checkboxes.get("dynamics") and self.export_checkboxes["dynamics"].isChecked()
            dynamics_warning.setVisible(bool(checked and group_by in {"day", "week"} and periods > 30))

        dynamics_combo.currentIndexChanged.connect(update_dynamics_warning)
        date_from.dateChanged.connect(update_dynamics_warning)
        date_to.dateChanged.connect(update_dynamics_warning)
        self.export_checkboxes["dynamics"].toggled.connect(update_dynamics_warning)
        update_dynamics_warning()

        duration_controls = QHBoxLayout()
        duration_sigma = self._create_sigma_spinbox(self._current_spinbox_value("duration_distribution_sigma_spinbox", 3.0))
        duration_controls.addWidget(QLabel("Порог отклонения (σ):"))
        duration_controls.addWidget(duration_sigma)
        duration_controls.addStretch()
        self.export_controls["duration_distribution_sigma"] = duration_sigma
        self._add_export_option(content_layout, "duration_distribution", "Распределение длительности", duration_controls)

        self._add_export_option(content_layout, "bottleneck_pareto", "Парето узких мест")

        content_layout.addStretch()
        content.setLayout(content_layout)
        scroll.setWidget(content)
        layout.addWidget(scroll)

        buttons = QHBoxLayout()
        select_all = QPushButton("Выбрать все")
        clear_all = QPushButton("Снять все")
        export_button = QPushButton("Сформировать PDF")
        export_button.setMinimumHeight(36)
        select_all.clicked.connect(lambda: self._set_all_export_sections(True))
        clear_all.clicked.connect(lambda: self._set_all_export_sections(False))
        export_button.clicked.connect(self.export_selected_pdf)
        buttons.addWidget(select_all)
        buttons.addWidget(clear_all)
        buttons.addStretch()
        buttons.addWidget(export_button)
        layout.addLayout(buttons)

        widget.setLayout(layout)
        return widget

    def _add_export_option(self, layout, key, label, controls_layout=None, checkbox_suffix=None):
        group = QGroupBox()
        group_layout = QVBoxLayout()
        checkbox = QCheckBox(label)
        checkbox.setChecked(True)
        checkbox.setStyleSheet("font-weight: bold;")
        checkbox_row = QHBoxLayout()
        checkbox_row.addWidget(checkbox)
        if checkbox_suffix is not None:
            checkbox_row.addWidget(checkbox_suffix)
        checkbox_row.addStretch()
        group_layout.addLayout(checkbox_row)
        if controls_layout is not None:
            group_layout.addLayout(controls_layout)
            controlled_widgets = []
            for index in range(controls_layout.count()):
                item = controls_layout.itemAt(index)
                if item.widget():
                    controlled_widgets.append(item.widget())

            def sync_controls(checked):
                for control in controlled_widgets:
                    control.setEnabled(checked)

            checkbox.toggled.connect(sync_controls)
            sync_controls(checkbox.isChecked())
        group.setLayout(group_layout)
        layout.addWidget(group)
        self.export_checkboxes[key] = checkbox

    def _set_all_export_sections(self, checked):
        for checkbox in self.export_checkboxes.values():
            checkbox.setChecked(checked)

    def _create_grouping_combo(self, current_value):
        combo = QComboBox()
        combo.addItem("По месяцам", "month")
        combo.addItem("По неделям", "week")
        combo.addItem("По дням", "day")
        index = combo.findData(current_value)
        if index >= 0:
            combo.setCurrentIndex(index)
        return combo

    def _create_sigma_spinbox(self, value):
        spinbox = QDoubleSpinBox()
        spinbox.setRange(1.0, 5.0)
        spinbox.setSingleStep(0.5)
        spinbox.setDecimals(1)
        spinbox.setValue(float(value))
        spinbox.setMaximumWidth(80)
        return spinbox

    def _create_export_date_edit(self, min_qdate, max_qdate, current_qdate):
        date_edit = QDateEdit()
        date_edit.setCalendarPopup(True)
        date_edit.setDisplayFormat("dd.MM.yyyy")
        date_edit.setDateRange(min_qdate, max_qdate)
        date_edit.setDate(current_qdate if current_qdate.isValid() else min_qdate)
        date_edit.setMinimumWidth(115)
        return date_edit

    def _current_combo_data(self, attr_name, default):
        combo = getattr(self, attr_name, None)
        if combo is not None:
            try:
                return combo.currentData() or default
            except RuntimeError:
                return default
        return default

    def _current_spinbox_value(self, attr_name, default):
        spinbox = getattr(self, attr_name, None)
        if spinbox is not None:
            try:
                return spinbox.value()
            except RuntimeError:
                return default
        return default

    def _current_date_edit_value(self, attr_name, default):
        date_edit = getattr(self, attr_name, None)
        if date_edit is not None:
            try:
                return date_edit.date()
            except RuntimeError:
                return default
        return default

    def _qdate_to_iso(self, qdate):
        return f"{qdate.year():04d}-{qdate.month():02d}-{qdate.day():02d}"

    def _collect_export_options(self):
        order = [
            "general_statistics",
            "time_statistics",
            "action_statistics",
            "transitions",
            "failed_processes",
            "outliers",
            "process_map",
            "transition_matrix",
            "dynamics",
            "duration_distribution",
            "bottleneck_pareto",
        ]
        labels = {
            "general_statistics": "Общая статистика",
            "time_statistics": "Временная статистика",
            "action_statistics": "Статистика действий",
            "transitions": "Переходы",
            "failed_processes": "Поиск сбоев",
            "outliers": "Поиск выбросов",
            "process_map": "Карта процессов",
            "transition_matrix": "Матрица переходов",
            "dynamics": "Динамика",
            "duration_distribution": "Распределение длительности",
            "bottleneck_pareto": "Парето узких мест",
        }
        options = {
            "order": order,
            "title": self.export_controls["report_title"].text().strip() or "Отчет по бизнес-процессу",
        }
        for key in order:
            options[key] = {
                "label": labels[key],
                "selected": self.export_checkboxes[key].isChecked(),
            }

        options["time_statistics"]["group_by"] = self.export_controls["time_statistics_group"].currentData()
        options["outliers"]["sigma"] = self.export_controls["outliers_sigma"].value()
        options["process_map"]["min_transitions"] = self.export_controls["process_map_min"].value()
        options["dynamics"]["group_by"] = self.export_controls["dynamics_group"].currentData()
        options["dynamics"]["date_from"] = self._qdate_to_iso(self.export_controls["dynamics_date_from"].date())
        options["dynamics"]["date_to"] = self._qdate_to_iso(self.export_controls["dynamics_date_to"].date())
        options["duration_distribution"]["sigma"] = self.export_controls["duration_distribution_sigma"].value()
        return options

    def _render_scene_image_for_report(self, scene, padding=30):
        rect = scene.itemsBoundingRect().adjusted(-padding, -padding, padding, padding)
        if rect.isEmpty():
            rect = scene.sceneRect().adjusted(-padding, -padding, padding, padding)
        if rect.isEmpty():
            return None

        items = scene.items()
        if items:
            content_rect = items[0].sceneBoundingRect()
            for item in items[1:]:
                content_rect = content_rect.united(item.sceneBoundingRect())
            rect.setLeft(content_rect.left() - 4)
            rect.setTop(content_rect.top() - padding)
            rect.setRight(content_rect.right() + padding)
            rect.setBottom(content_rect.bottom() + padding)

        max_side = 14000
        scale = min(1.0, max_side / max(1.0, rect.width()), max_side / max(1.0, rect.height()))
        width = max(1, int(rect.width() * scale))
        height = max(1, int(rect.height() * scale))

        image = QImage(width, height, QImage.Format_ARGB32)
        image.fill(QColor(255, 255, 255))
        painter = QPainter(image)
        painter.setRenderHint(QPainter.Antialiasing)
        scene.render(painter, QRectF(0, 0, width, height), rect)
        painter.end()
        return image

    def _graphics_tab_image_for_report(self, factory, configure=None, padding=30):
        widget = factory()
        try:
            if configure is not None:
                configure(widget)
                QApplication.processEvents()
            views = widget.findChildren(ProcessMapView)
            if not views:
                return None
            return self._render_scene_image_for_report(views[0].scene(), padding=padding)
        finally:
            widget.deleteLater()

    def _build_report_graphics_images(self, options):
        images = {}

        if options["process_map"]["selected"]:
            def configure_map(widget):
                spinbox = getattr(self, "process_map_min_transitions_spinbox", None)
                if spinbox is not None:
                    spinbox.setValue(options["process_map"]["min_transitions"])

            image = self._graphics_tab_image_for_report(
                self.create_process_map_graphics_tab,
                configure=configure_map,
                padding=45,
            )
            if image is not None:
                images["process_map"] = image.transformed(
                    QTransform().rotate(90),
                    Qt.SmoothTransformation,
                )

        if options["transition_matrix"]["selected"]:
            image = self._graphics_tab_image_for_report(
                self.create_transition_heatmap_tab,
                padding=45,
            )
            if image is not None:
                images["transition_matrix"] = image

        if options["dynamics"]["selected"]:
            def configure_dynamics(widget):
                combo = getattr(self, "graphics_dynamics_combo", None)
                if combo is not None:
                    index = combo.findData(options["dynamics"]["group_by"])
                    if index >= 0:
                        combo.setCurrentIndex(index)
                for attr_name, value in (
                    ("graphics_dynamics_date_from", options["dynamics"]["date_from"]),
                    ("graphics_dynamics_date_to", options["dynamics"]["date_to"]),
                ):
                    date_edit = getattr(self, attr_name, None)
                    if date_edit is not None:
                        try:
                            year, month, day = [int(part) for part in str(value).split("-")]
                            date_edit.setDate(QDate(year, month, day))
                        except Exception:
                            pass

            image = self._graphics_tab_image_for_report(
                self.create_time_dynamics_graphics_tab,
                configure=configure_dynamics,
                padding=45,
            )
            if image is not None:
                images["dynamics"] = image

        if options["duration_distribution"]["selected"]:
            def configure_duration(widget):
                spinbox = getattr(self, "duration_distribution_sigma_spinbox", None)
                if spinbox is not None:
                    spinbox.setValue(options["duration_distribution"]["sigma"])

            image = self._graphics_tab_image_for_report(
                self.create_duration_distribution_graphics_tab,
                configure=configure_duration,
                padding=45,
            )
            if image is not None:
                images["duration_distribution"] = image

        if options["bottleneck_pareto"]["selected"]:
            image = self._graphics_tab_image_for_report(
                self.create_pareto_bottlenecks_graphics_tab,
                padding=45,
            )
            if image is not None:
                images["bottleneck_pareto"] = image

        return images

    def export_selected_pdf(self):
        options = self._collect_export_options()
        if not any(options[key]["selected"] for key in options["order"]):
            QMessageBox.information(self, "Экспорт", "Выберите хотя бы один раздел для отчета")
            return

        default_name = os.path.splitext(os.path.basename(self.json_file_path))[0] + "_report.pdf"
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить PDF отчет",
            os.path.join(os.path.dirname(self.json_file_path), default_name),
            "PDF файлы (*.pdf)",
        )
        if not output_path:
            return
        if not output_path.lower().endswith(".pdf"):
            output_path += ".pdf"

        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            from report_exporter import ReportExporter

            graphics_images = self._build_report_graphics_images(options)
            ReportExporter(self.json_file_path, options, graphics_images).export_pdf(output_path)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка экспорта", f"Не удалось сформировать PDF:\n{e}")
            return
        finally:
            QApplication.restoreOverrideCursor()

        QMessageBox.information(self, "Экспорт", f"PDF отчет сохранен:\n{output_path}")

    def _build_graphics_page(self):
        if self.cached_graphics_data == self.json_file_path and self.graphics_widget is not None:
            return
        self._clear_layout(self.graphics_layout)
        self.graphics_widget = self.create_graphics_widget()
        self.graphics_layout.addWidget(self.graphics_widget)
        self.cached_graphics_data = self.json_file_path

    def create_graphics_widget(self):
        widget = QWidget()
        layout = QVBoxLayout()

        tabs = QTabWidget()
        tabs.addTab(self.create_process_map_graphics_tab(), "Карта процессов")
        tabs.addTab(self.create_transition_heatmap_tab(), "Матрица переходов")
        tabs.addTab(self.create_time_dynamics_graphics_tab(), "Динамика")
        tabs.addTab(self.create_duration_distribution_graphics_tab(), "Распределение длительности")
        tabs.addTab(self.create_pareto_bottlenecks_graphics_tab(), "Парето узких мест")
        tabs.currentChanged.connect(lambda index: self._fit_visible_graphics_tab(tabs, index))
        layout.addWidget(tabs)

        widget.setLayout(layout)
        QTimer.singleShot(0, lambda: self._fit_visible_graphics_tab(tabs, tabs.currentIndex()))
        return widget

    def _fit_visible_graphics_tab(self, tabs, index):
        current_widget = tabs.widget(index)
        if current_widget is None:
            return
        for view in current_widget.findChildren(ProcessMapView):
            view.request_fit()

    def _create_scene_view(self, scene, minimum_height=520):
        view = ProcessMapView(scene)
        view.setMinimumHeight(minimum_height)
        return view

    def _set_scene_content(self, view, scene, padding=70):
        rect = scene.itemsBoundingRect().adjusted(-padding, -padding, padding, padding)
        if rect.isEmpty():
            rect = scene.sceneRect().adjusted(-padding, -padding, padding, padding)
        view.set_content_rect(rect)

    def _add_scene_text(self, scene, text, x, y, size=9, color=None, bold=False, max_width=None):
        font = QFont("Arial", size)
        font.setBold(bold)
        item = scene.addText(str(text), font)
        item.setDefaultTextColor(color or QColor(45, 54, 64))
        if max_width:
            item.setTextWidth(max_width)
        item.setPos(x, y)
        return item

    def _angled_label_font(self, size=8):
        font = QFont("Segoe UI", size + 2)
        font.setBold(True)
        return font

    def _draw_empty_scene(self, scene, message):
        self._add_scene_text(scene, message, 40, 40, size=12, color=QColor(95, 95, 95), bold=True)

    def _blend_color(self, low, high, ratio):
        ratio = max(0.0, min(1.0, ratio))
        return QColor(
            int(low.red() + (high.red() - low.red()) * ratio),
            int(low.green() + (high.green() - low.green()) * ratio),
            int(low.blue() + (high.blue() - low.blue()) * ratio),
        )

    def _duration_label_from_seconds(self, seconds):
        return format_timedelta(timedelta(seconds=max(0, int(seconds))))

    def _process_rows(self, data):
        rows = []
        for case_id, records in group_processes(data).items():
            if len(records) < 2:
                continue
            start = parse_dt(records[0].get("datetime"))
            end = parse_dt(records[-1].get("datetime"))
            if not start or not end:
                continue
            rows.append({
                "case_id": case_id,
                "records": records,
                "start": start,
                "end": end,
                "duration_seconds": max(0, (end - start).total_seconds()),
                "action_count": len(records),
            })
        return rows

    def _transition_duration_rows(self, data):
        rows = []
        for case_id, records in group_processes(data).items():
            for index in range(len(records) - 1):
                current = records[index]
                next_record = records[index + 1]
                start = parse_dt(current.get("datetime"))
                end = parse_dt(next_record.get("datetime"))
                if not start or not end:
                    continue
                seconds = max(0, (end - start).total_seconds())
                source = action_name(data, current.get("action_id"))
                target = action_name(data, next_record.get("action_id"))
                rows.append({
                    "case_id": case_id,
                    "from": source,
                    "to": target,
                    "duration_seconds": seconds,
                })
        return rows

    def _process_date_bounds(self, data):
        start_dates = []
        for records in group_processes(data).values():
            if not records:
                continue
            start_dt = parse_dt(records[0].get("datetime"))
            if start_dt is not None:
                start_dates.append(start_dt.date())
        if not start_dates:
            today = datetime.now().date()
            return today, today
        return min(start_dates), max(start_dates)

    def _filter_data_by_process_start_date(self, data, start_date, end_date):
        records = []
        for process_records in group_processes(data).values():
            if not process_records:
                continue
            start_dt = parse_dt(process_records[0].get("datetime"))
            if start_dt is None:
                continue
            process_date = start_dt.date()
            if start_date <= process_date <= end_date:
                records.extend(process_records)
        filtered_data = dict(data)
        filtered_data["records"] = [
            {key: value for key, value in record.items() if key != "_row_index"}
            for record in records
        ]
        return filtered_data

    def _seconds_to_scene_y(self, seconds, min_seconds, max_seconds, top, height):
        if max_seconds <= min_seconds:
            return top + height / 2
        ratio = (seconds - min_seconds) / (max_seconds - min_seconds)
        return top + height - ratio * height

    def create_transition_heatmap_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        data = load_json_data(self.json_file_path)
        actions = [action_name(data, index) for index, _ in enumerate(data.get("actions", []))]
        transitions = Counter()
        for records in group_processes(data).values():
            names = [action_name(data, record.get("action_id")) for record in records]
            for source, target in zip(names, names[1:]):
                transitions[(source, target)] += 1

        scene = QGraphicsScene()
        view = self._create_scene_view(scene)

        if not actions or not transitions:
            self._draw_empty_scene(scene, "Нет переходов для построения матрицы")
        else:
            max_count = max(transitions.values())
            cell = 52
            left = 260
            top = 210
            font = QFont("Arial", 10)
            metrics = QFontMetrics(font)
            max_label_width = 220

            self._add_scene_text(scene, "Матрица переходов", left, 30, size=16, bold=True)
            self._add_scene_text(
                scene,
                "Цвет ячейки показывает частоту перехода от действия в строке к действию в столбце.",
                left,
                58,
                size=11,
                color=QColor(90, 95, 100),
            )

            for index, name in enumerate(actions):
                y = top + index * cell
                display = metrics.elidedText(name, Qt.ElideRight, max_label_width)
                self._add_scene_text(scene, display, 28, y + 14, size=10, max_width=max_label_width)

                x = left + index * cell
                rotated = scene.addText(metrics.elidedText(name, Qt.ElideRight, 150), font)
                rotated.setDefaultTextColor(QColor(45, 54, 64))
                rotated.setTransformOriginPoint(rotated.boundingRect().left(), rotated.boundingRect().bottom())
                rotated.setRotation(-45)
                rotated.setPos(x + 6, top - 18)

            low = QColor(235, 244, 248)
            high = QColor(29, 102, 150)
            for row, source in enumerate(actions):
                for col, target in enumerate(actions):
                    count = transitions.get((source, target), 0)
                    ratio = count / max_count if max_count else 0
                    color = self._blend_color(low, high, math.sqrt(ratio)) if count else QColor(248, 250, 252)
                    x = left + col * cell
                    y = top + row * cell
                    rect = scene.addRect(x, y, cell, cell, QPen(QColor(218, 226, 234)), QBrush(color))
                    rect.setToolTip(f"{source} -> {target}: {count}")
                    if count:
                        text = scene.addText(str(count), QFont("Arial", 10))
                        text.setDefaultTextColor(QColor(255, 255, 255) if ratio > 0.55 else QColor(35, 45, 55))
                        bounds = text.boundingRect()
                        text.setPos(x + (cell - bounds.width()) / 2, y + (cell - bounds.height()) / 2)

            legend_x = left + len(actions) * cell + 60
            legend_y = top
            self._add_scene_text(scene, "Частота", legend_x, legend_y - 28, bold=True)
            for i in range(8):
                color = self._blend_color(low, high, i / 7)
                scene.addRect(legend_x, legend_y + i * 28, 34, 22, QPen(Qt.NoPen), QBrush(color))
            self._add_scene_text(scene, f"0", legend_x + 44, legend_y + 2)
            self._add_scene_text(scene, f"{max_count}", legend_x + 44, legend_y + 7 * 28 + 2)

        self._set_scene_content(view, scene)
        layout.addWidget(view)
        widget.setLayout(layout)
        return widget

    def create_time_dynamics_graphics_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        controls = QHBoxLayout()
        controls.addWidget(QLabel("Группировка:"))
        combo = QComboBox()
        combo.addItem("По месяцам", "month")
        combo.addItem("По неделям", "week")
        combo.addItem("По дням", "day")
        self.graphics_dynamics_combo = combo
        controls.addWidget(combo)
        data = self._get_analysis_data()
        min_process_date, max_process_date = self._process_date_bounds(data)
        min_qdate = QDate(min_process_date.year, min_process_date.month, min_process_date.day)
        max_qdate = QDate(max_process_date.year, max_process_date.month, max_process_date.day)

        controls.addSpacing(18)
        controls.addWidget(QLabel("Диапазон: с"))
        date_from = QDateEdit()
        date_from.setCalendarPopup(True)
        date_from.setDisplayFormat("dd.MM.yyyy")
        date_from.setDateRange(min_qdate, max_qdate)
        date_from.setDate(min_qdate)
        date_from.setMinimumWidth(115)
        date_from.calendarWidget().setCurrentPage(min_process_date.year, min_process_date.month)
        self.graphics_dynamics_date_from = date_from
        controls.addWidget(date_from)

        controls.addWidget(QLabel("по"))
        date_to = QDateEdit()
        date_to.setCalendarPopup(True)
        date_to.setDisplayFormat("dd.MM.yyyy")
        date_to.setDateRange(min_qdate, max_qdate)
        date_to.setDate(max_qdate)
        date_to.setMinimumWidth(115)
        date_to.calendarWidget().setCurrentPage(min_process_date.year, min_process_date.month)
        self.graphics_dynamics_date_to = date_to
        controls.addWidget(date_to)
        controls.addStretch()

        scene = QGraphicsScene()
        view = self._create_scene_view(scene)

        def draw():
            scene.clear()
            start_qdate = date_from.date()
            end_qdate = date_to.date()
            start_date = datetime(start_qdate.year(), start_qdate.month(), start_qdate.day()).date()
            end_date = datetime(end_qdate.year(), end_qdate.month(), end_qdate.day()).date()
            if start_date > end_date:
                start_date, end_date = end_date, start_date
            filtered_data = self._filter_data_by_process_start_date(data, start_date, end_date)
            rows = time_dynamics(filtered_data, combo.currentData())
            if not rows:
                self._draw_empty_scene(scene, "Нет данных для построения динамики")
                self._set_scene_content(view, scene)
                return

            left, top = 90, 80
            width, height = max(760, len(rows) * 90), 360
            max_volume = max(max(row["started"], row["completed"]) for row in rows) or 1
            max_fail_rate = max(row["failed_rate"] for row in rows) or 1
            fail_axis_max = min(100, max(10, int(math.ceil(max_fail_rate / 10.0) * 10)))
            bar_width = max(18, min(32, width / len(rows) * 0.28))
            step = width / len(rows)

            self._add_scene_text(scene, "Временная динамика процесса", left, 25, size=16, bold=True)
            self._add_scene_text(scene, "Столбцы: старт/без сбоев, красная линия: доля сбоев.", left, 52, size=11, color=QColor(90, 95, 100))
            scene.addLine(left, top, left, top + height, QPen(QColor(95, 105, 115), 1))
            scene.addLine(left, top + height, left + width, top + height, QPen(QColor(95, 105, 115), 1))
            for tick_index in range(6):
                tick_value = max_volume * tick_index / 5
                y = top + height - height * tick_value / max_volume
                scene.addLine(left - 8, y, left, y, QPen(QColor(95, 105, 115), 1))
                tick_label = scene.addText(f"{tick_value:.0f}", QFont("Arial", 10))
                tick_label.setDefaultTextColor(QColor(70, 75, 80))
                tick_label.setPos(left - 14 - tick_label.boundingRect().width(), y - tick_label.boundingRect().height() / 2)
            self._add_scene_text(scene, "Процессы", left - 72, top - 28, color=QColor(70, 75, 80), bold=True)

            right_axis_x = left + width
            scene.addLine(right_axis_x, top, right_axis_x, top + height, QPen(QColor(190, 70, 70), 1))
            for tick_index in range(6):
                tick_value = fail_axis_max * tick_index / 5
                y = top + height - height * tick_value / fail_axis_max
                scene.addLine(right_axis_x, y, right_axis_x + 8, y, QPen(QColor(190, 70, 70), 1))
                if tick_index > 0:
                    scene.addLine(left, y, right_axis_x, y, QPen(QColor(226, 232, 238), 1, Qt.DashLine))
                tick_label = scene.addText(f"{tick_value:.0f}%", QFont("Arial", 10))
                tick_label.setDefaultTextColor(QColor(170, 55, 55))
                tick_label.setPos(right_axis_x + 12, y - tick_label.boundingRect().height() / 2)
            self._add_scene_text(scene, "Доля сбоев", right_axis_x + 12, top - 28, color=QColor(170, 55, 55), bold=True)

            fail_points = []
            for index, row in enumerate(rows):
                center_x = left + step * index + step / 2
                started_h = height * row["started"] / max_volume
                completed_h = height * row["completed"] / max_volume
                scene.addRect(center_x - bar_width - 2, top + height - started_h, bar_width, started_h, QPen(Qt.NoPen), QBrush(QColor(79, 129, 189)))
                scene.addRect(center_x + 2, top + height - completed_h, bar_width, completed_h, QPen(Qt.NoPen), QBrush(QColor(91, 155, 112)))

                fail_y = top + height - height * row["failed_rate"] / fail_axis_max
                fail_points.append((center_x, fail_y, row))
                label = scene.addText(row["period"], self._angled_label_font(9))
                label.setDefaultTextColor(QColor(70, 75, 80))
                label.setTransformOriginPoint(label.boundingRect().left(), label.boundingRect().top())
                label.setRotation(34)
                label.setPos(center_x - 12, top + height + 18)

            for p1, p2 in zip(fail_points, fail_points[1:]):
                scene.addLine(p1[0], p1[1], p2[0], p2[1], QPen(QColor(190, 70, 70), 2))
            for x, y, row in fail_points:
                dot = scene.addEllipse(x - 4, y - 4, 8, 8, QPen(QColor(150, 45, 45)), QBrush(QColor(210, 83, 83)))
                dot.setToolTip(f"{row['period']}: {row['failed_rate']}% сбоев")

            self._add_scene_text(scene, f"Макс. объем: {max_volume}", left + width + 35, top + 10)
            self._add_scene_text(scene, f"Макс. % сбоев: {max_fail_rate}", left + width + 35, top + 34)
            scene.addRect(left + width + 35, top + 72, 18, 12, QPen(Qt.NoPen), QBrush(QColor(79, 129, 189)))
            self._add_scene_text(scene, "Стартовало", left + width + 60, top + 66)
            scene.addRect(left + width + 35, top + 100, 18, 12, QPen(Qt.NoPen), QBrush(QColor(91, 155, 112)))
            self._add_scene_text(scene, "Без сбоев", left + width + 60, top + 94)
            scene.addLine(left + width + 35, top + 134, left + width + 53, top + 134, QPen(QColor(190, 70, 70), 2))
            self._add_scene_text(scene, "% сбоев", left + width + 60, top + 122)
            self._set_scene_content(view, scene)

        combo.currentIndexChanged.connect(draw)
        date_from.dateChanged.connect(draw)
        date_to.dateChanged.connect(draw)
        layout.addLayout(controls)
        layout.addWidget(view)
        widget.setLayout(layout)
        draw()
        return widget

    def create_duration_distribution_graphics_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        controls = QHBoxLayout()
        controls.addWidget(QLabel("Порог отклонения (σ):"))
        sigma_spinbox = QDoubleSpinBox()
        sigma_spinbox.setRange(1.0, 5.0)
        sigma_spinbox.setValue(3.0)
        sigma_spinbox.setSingleStep(0.5)
        sigma_spinbox.setDecimals(1)
        sigma_spinbox.setMaximumWidth(80)
        self.duration_distribution_sigma_spinbox = sigma_spinbox
        controls.addWidget(sigma_spinbox)
        controls.addStretch()

        data = self._get_analysis_data()
        rows = self._process_rows(data)
        scene = QGraphicsScene()
        view = self._create_scene_view(scene)

        def draw():
            scene.clear()
            durations = sorted(row["duration_seconds"] for row in rows)
            if not durations:
                self._draw_empty_scene(scene, "Нет длительностей процессов для распределения")
                self._set_scene_content(view, scene)
                return

            all_avg = sum(durations) / len(durations)
            if len(durations) > 1:
                all_std = (sum((value - all_avg) ** 2 for value in durations) / (len(durations) - 1)) ** 0.5
            else:
                all_std = 0
            sigma_threshold = sigma_spinbox.value()
            if all_std > 0:
                visible_durations = [
                    value for value in durations
                    if abs(value - all_avg) / all_std < sigma_threshold
                ]
            else:
                visible_durations = durations
            if not visible_durations:
                visible_durations = durations
            hidden_count = len(durations) - len(visible_durations)

            left, top = 90, 90
            width, height = 820, 360
            bin_count = min(12, max(5, int(math.sqrt(len(visible_durations)))))
            min_d, max_d = min(visible_durations), max(visible_durations)
            span = max(1, max_d - min_d)
            bins = [0] * bin_count
            for value in visible_durations:
                index = min(bin_count - 1, int((value - min_d) / span * bin_count))
                bins[index] += 1
            max_bin = max(bins) or 1
            bar_gap = 10
            bar_width = (width - bar_gap * (bin_count - 1)) / bin_count
            avg = sum(visible_durations) / len(visible_durations)
            median = (
                visible_durations[len(visible_durations) // 2]
                if len(visible_durations) % 2
                else (visible_durations[len(visible_durations) // 2 - 1] + visible_durations[len(visible_durations) // 2]) / 2
            )

            self._add_scene_text(scene, "Распределение длительности процессов", left, 8, size=16, bold=True)
            subtitle = (
                "Гистограмма показывает, сколько процессов попадает в интервалы длительности. "
                f"Порог: {sigma_threshold}σ."
            )
            if hidden_count:
                subtitle += f" Скрыто выбросов: {hidden_count}."
            self._add_scene_text(scene, subtitle, left, 35, size=11, color=QColor(90, 95, 100), max_width=900)
            scene.addLine(left, top, left, top + height, QPen(QColor(95, 105, 115)))
            scene.addLine(left, top + height, left + width, top + height, QPen(QColor(95, 105, 115)))

            for index, count in enumerate(bins):
                x = left + index * (bar_width + bar_gap)
                bar_h = height * count / max_bin
                scene.addRect(x, top + height - bar_h, bar_width, bar_h, QPen(Qt.NoPen), QBrush(QColor(88, 142, 174)))
                label_start = min_d + span * index / bin_count
                label_end = min_d + span * (index + 1) / bin_count
                label = scene.addText(f"{self._duration_label_from_seconds(label_start)}\n-\n{self._duration_label_from_seconds(label_end)}", QFont("Arial", 9))
                label.setDefaultTextColor(QColor(70, 75, 80))
                label.setPos(x, top + height + 16)
                count_text = scene.addText(str(count), QFont("Arial", 10))
                count_text.setDefaultTextColor(QColor(45, 54, 64))
                count_text.setPos(x + bar_width / 2 - count_text.boundingRect().width() / 2, top + height - bar_h - 22)

            def marker(value, color, label, label_y):
                x = left + (value - min_d) / span * width if span else left
                scene.addLine(x, top - 8, x, top + height, QPen(color, 2))
                self._add_scene_text(scene, f"{label}: {self._duration_label_from_seconds(value)}", x + 8, label_y, color=color, bold=True)

            marker(avg, QColor(192, 95, 65), "Среднее", top - 36)
            marker(median, QColor(67, 125, 92), "Медиана", top - 16)

            self._set_scene_content(view, scene)

        sigma_spinbox.valueChanged.connect(draw)
        layout.addLayout(controls)
        layout.addWidget(view)
        widget.setLayout(layout)
        draw()
        return widget

    def create_outlier_scatter_graphics_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        data = load_json_data(self.json_file_path)
        rows = self._process_rows(data)
        final_ids = common_final_action_ids(data)
        scene = QGraphicsScene()
        view = self._create_scene_view(scene)

        if not rows:
            self._draw_empty_scene(scene, "Нет кейсов для диаграммы выбросов")
        else:
            durations = [row["duration_seconds"] for row in rows]
            avg = sum(durations) / len(durations)
            std = (sum((value - avg) ** 2 for value in durations) / max(1, len(durations) - 1)) ** 0.5 if len(durations) > 1 else 0
            display_rows = rows
            if len(rows) > 4000:
                important_rows = [
                    row for row in rows
                    if (
                        (std and abs(row["duration_seconds"] - avg) / std >= 3)
                        or (final_ids and row["records"][-1].get("action_id") not in final_ids)
                    )
                ]
                important_rows.sort(key=lambda row: abs(row["duration_seconds"] - avg), reverse=True)
                important_rows = important_rows[:4000]
                important_ids = {row["case_id"] for row in important_rows}
                remaining_limit = max(0, 4000 - len(important_rows))
                step = max(1, len(rows) // max(1, remaining_limit))
                sampled_rows = [
                    row for index, row in enumerate(rows)
                    if index % step == 0 and row["case_id"] not in important_ids
                ][:remaining_limit]
                display_rows = important_rows + sampled_rows
            min_start = min(row["start"] for row in rows)
            max_start = max(row["start"] for row in rows)
            min_d, max_d = min(durations), max(durations)
            span_time = max(1, (max_start - min_start).total_seconds())
            span_d = max(1, max_d - min_d)
            left, top = 90, 80
            width, height = 860, 390

            self._add_scene_text(scene, "Кейсы-выбросы", left, 25, size=14, bold=True)
            subtitle = "X: дата старта, Y: длительность, размер точки: количество действий, красный: сбой."
            if len(display_rows) < len(rows):
                subtitle += f" Показано {len(display_rows)} из {len(rows)} кейсов: выборка плюс все сбои/3σ."
            self._add_scene_text(scene, subtitle, left, 52, color=QColor(90, 95, 100), max_width=900)
            scene.addLine(left, top, left, top + height, QPen(QColor(95, 105, 115)))
            scene.addLine(left, top + height, left + width, top + height, QPen(QColor(95, 105, 115)))

            threshold = avg + 3 * std if std else None
            if threshold is not None:
                y = self._seconds_to_scene_y(threshold, min_d, max_d, top, height)
                scene.addLine(left, y, left + width, y, QPen(QColor(190, 70, 70), 1, Qt.DashLine))
                self._add_scene_text(scene, f"3σ: {self._duration_label_from_seconds(threshold)}", left + width + 18, y - 12, color=QColor(170, 55, 55))

            max_actions = max(row["action_count"] for row in rows) or 1
            for row in display_rows:
                x = left + ((row["start"] - min_start).total_seconds() / span_time) * width
                y = self._seconds_to_scene_y(row["duration_seconds"], min_d, max_d, top, height)
                failed = bool(final_ids and row["records"][-1].get("action_id") not in final_ids)
                is_outlier = bool(std and abs(row["duration_seconds"] - avg) / std >= 3)
                radius = 4 + 8 * row["action_count"] / max_actions
                color = QColor(205, 78, 78) if failed else QColor(65, 125, 171)
                if is_outlier:
                    color = QColor(218, 116, 42) if not failed else QColor(170, 52, 52)
                dot = scene.addEllipse(x - radius, y - radius, radius * 2, radius * 2, QPen(QColor(255, 255, 255), 1), QBrush(color))
                dot.setToolTip(
                    f"Case_{row['case_id']}: {self._duration_label_from_seconds(row['duration_seconds'])}, "
                    f"{row['action_count']} действий"
                )

            self._add_scene_text(scene, self._duration_label_from_seconds(max_d), 18, top - 8)
            self._add_scene_text(scene, self._duration_label_from_seconds(min_d), 18, top + height - 10)
            self._add_scene_text(scene, min_start.strftime("%Y-%m-%d"), left, top + height + 18)
            self._add_scene_text(scene, max_start.strftime("%Y-%m-%d"), left + width - 80, top + height + 18)

        self._set_scene_content(view, scene)
        layout.addWidget(view)
        widget.setLayout(layout)
        return widget

    def create_pareto_bottlenecks_graphics_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()
        data = load_json_data(self.json_file_path)
        durations = defaultdict(float)
        counts = Counter()
        for row in self._transition_duration_rows(data):
            key = f"{row['from']} -> {row['to']}"
            durations[key] += row["duration_seconds"]
            counts[key] += 1
        top_items = sorted(durations.items(), key=lambda item: item[1], reverse=True)[:12]
        total = sum(durations.values()) or 1

        scene = QGraphicsScene()
        view = self._create_scene_view(scene)
        if not top_items:
            self._draw_empty_scene(scene, "Нет переходов с длительностью для Парето")
        else:
            left, top = 90, 85
            width, height = 860, 360
            bar_gap = 10
            bar_width = (width - bar_gap * (len(top_items) - 1)) / len(top_items)
            max_value = max(value for _, value in top_items) or 1
            cumulative = 0
            cumulative_points = []

            self._add_scene_text(scene, "Парето узких мест", left, 25, size=16, bold=True)
            self._add_scene_text(scene, "Столбцы: суммарное время ожидания перехода, линия: накопленная доля от общего времени.", left, 52, size=11, color=QColor(90, 95, 100))
            scene.addLine(left, top, left, top + height, QPen(QColor(95, 105, 115)))
            scene.addLine(left, top + height, left + width, top + height, QPen(QColor(95, 105, 115)))

            for index, (name, seconds) in enumerate(top_items):
                x = left + index * (bar_width + bar_gap)
                bar_h = height * seconds / max_value
                scene.addRect(x, top + height - bar_h, bar_width, bar_h, QPen(Qt.NoPen), QBrush(QColor(79, 129, 189)))
                cumulative += seconds
                cumulative_y = top + height - height * cumulative / total
                cumulative_points.append((x + bar_width / 2, cumulative_y))
                label = scene.addText(name, self._angled_label_font(9))
                label.setDefaultTextColor(QColor(70, 75, 80))
                label.setTransformOriginPoint(label.boundingRect().left(), label.boundingRect().top())
                label.setRotation(32)
                label.setPos(x + 4, top + height + 28)
                value_label = scene.addText(self._duration_label_from_seconds(seconds), QFont("Arial", 9))
                value_label.setDefaultTextColor(QColor(45, 54, 64))
                value_label.setPos(x, top + height - bar_h - 20)

            for p1, p2 in zip(cumulative_points, cumulative_points[1:]):
                scene.addLine(p1[0], p1[1], p2[0], p2[1], QPen(QColor(190, 70, 70), 2))
            for x, y in cumulative_points:
                scene.addEllipse(x - 4, y - 4, 8, 8, QPen(QColor(160, 45, 45)), QBrush(QColor(210, 83, 83)))
            self._add_scene_text(scene, "100%", left + width + 20, top - 8, color=QColor(170, 55, 55))
            self._add_scene_text(scene, f"Макс.: {self._duration_label_from_seconds(max_value)}", left + width + 20, top + 28)

        self._set_scene_content(view, scene, padding=105)
        layout.addWidget(view)
        widget.setLayout(layout)
        return widget

    def create_process_map_graphics_tab(self):
        widget = QWidget()
        layout = QVBoxLayout()

        data = load_json_data(self.json_file_path)
        actions, transitions = process_map(data, limit=60)

        title = QLabel(
            "Карта процесса: центральная линия показывает наиболее сильный поток, "
            "ветви вынесены выше и ниже, толщина связей зависит от частоты перехода."
        )
        title.setWordWrap(True)
        layout.addWidget(title)

        controls = QHBoxLayout()
        controls.addWidget(QLabel("Отображать связи от"))
        min_transitions_spinbox = QDoubleSpinBox()
        min_transitions_spinbox.setRange(0, 1000000)
        min_transitions_spinbox.setValue(0)
        min_transitions_spinbox.setSingleStep(1)
        min_transitions_spinbox.setDecimals(0)
        min_transitions_spinbox.setMaximumWidth(90)
        self.process_map_min_transitions_spinbox = min_transitions_spinbox
        controls.addWidget(min_transitions_spinbox)
        controls.addWidget(QLabel("переходов"))

        controls.addStretch()
        layout.addLayout(controls)

        splitter = QSplitter(Qt.Vertical)
        scene = QGraphicsScene()
        view = ProcessMapView(scene)
        view.setMinimumHeight(470)
        table = QTreeWidget()
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setHeaderLabels(["Откуда", "Куда", "Переходов", "Case ID"])
        table.header().setSectionResizeMode(0, QHeaderView.Stretch)
        table.header().setSectionResizeMode(1, QHeaderView.Stretch)
        table.header().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        table.header().setSectionResizeMode(3, QHeaderView.ResizeToContents)

        declared_order = [
            item.get("name", str(item))
            if isinstance(item, dict) else str(item)
            for item in data.get("actions", [])
        ]

        def strongest_main_flow(ordered_actions, available_transitions):
            if not ordered_actions:
                return []
            order_index = {name: index for index, name in enumerate(ordered_actions)}
            forward_edges = [
                item for item in available_transitions
                if (
                    item["from"] in order_index
                    and item["to"] in order_index
                    and order_index[item["from"]] < order_index[item["to"]]
                )
            ]
            if not forward_edges:
                return ordered_actions[: min(8, len(ordered_actions))]

            best_score = {name: 0 for name in ordered_actions}
            best_length = {name: 1 for name in ordered_actions}
            predecessor = {}
            outgoing = {}
            for edge in forward_edges:
                outgoing.setdefault(edge["from"], []).append(edge)

            for name in ordered_actions:
                for edge in outgoing.get(name, []):
                    target = edge["to"]
                    candidate_score = best_score[name] + edge["count"]
                    candidate_length = best_length[name] + 1
                    current = (best_score[target], best_length[target])
                    candidate = (candidate_score, candidate_length)
                    if candidate > current:
                        best_score[target] = candidate_score
                        best_length[target] = candidate_length
                        predecessor[target] = name

            end_name = max(
                ordered_actions,
                key=lambda name: (best_score[name], best_length[name], order_index[name])
            )
            if best_length[end_name] <= 1:
                return ordered_actions[: min(8, len(ordered_actions))]

            flow = [end_name]
            while flow[0] in predecessor:
                flow.insert(0, predecessor[flow[0]])
            return flow[:12]

        def draw_process_map(*args):
            scene.clear()
            table.clear()
            min_transition_count = int(min_transitions_spinbox.value())
            filtered_transitions = [
                item for item in transitions
                if item["count"] >= min_transition_count
            ]
            filtered_action_names = []
            for transition in filtered_transitions:
                for name in (transition["from"], transition["to"]):
                    if name not in filtered_action_names:
                        filtered_action_names.append(name)

            if actions and filtered_transitions:
                draw_actions = filtered_action_names or actions
                self._draw_process_map_scene(
                    scene,
                    view,
                    data,
                    declared_order,
                    draw_actions,
                    filtered_transitions,
                    strongest_main_flow,
                )
            elif actions and min_transition_count <= 0:
                self._draw_process_map_scene(
                    scene,
                    view,
                    data,
                    declared_order,
                    actions,
                    filtered_transitions,
                    strongest_main_flow,
                )
            else:
                scene.addText("Нет связей, подходящих под фильтр")
                view.set_content_rect(scene.itemsBoundingRect().adjusted(-80, -80, 80, 80))

            for transition in filtered_transitions:
                transition_item = QTreeWidgetItem([
                    transition['from'],
                    transition['to'],
                    str(transition['count']),
                    "",
                ])
                for case_id in transition.get('case_ids', []):
                    case_item = QTreeWidgetItem(transition_item, ["", "", "", f"Case_{case_id}"])
                    case_item.setData(3, Qt.UserRole, case_id)
                table.addTopLevelItem(transition_item)

        min_transitions_spinbox.valueChanged.connect(draw_process_map)
        table.itemDoubleClicked.connect(self.open_case_from_tree_item)
        splitter.addWidget(view)
        splitter.addWidget(table)
        splitter.setSizes([520, 220])
        layout.addWidget(splitter)
        widget.setLayout(layout)
        draw_process_map()
        return widget

    def _draw_process_map_scene(self, scene, view, data, declared_order, actions, transitions, strongest_main_flow):
        if actions:
            ordered_actions = declared_order or actions
            main_flow = strongest_main_flow(ordered_actions, transitions)
            visible_actions = list(main_flow)
            for transition in transitions:
                for name in (transition["from"], transition["to"]):
                    if name not in visible_actions and len(visible_actions) < 24:
                        visible_actions.append(name)
            for name in ordered_actions + [name for name in actions if name not in ordered_actions]:
                if name not in visible_actions and len(visible_actions) < 24:
                    visible_actions.append(name)

            visible_set = set(visible_actions)
            visible_transitions = [
                item for item in transitions
                if item["from"] in visible_set and item["to"] in visible_set
            ]
            transition_max = max((item["count"] for item in visible_transitions), default=1)

            font = QFont("Arial", 9)
            metrics = QFontMetrics(font)
            widest_label = max(
                (metrics.horizontalAdvance(name) for name in visible_actions),
                default=120
            )
            node_width = max(170, min(320, widest_label + 42))
            node_height = 64
            horizontal_gap = max(230, node_width + 70)
            vertical_gap = 132
            left_margin = node_width / 2 + 80

            branch_meta = {}
            anchor_usage = {}
            main_index = {name: index for index, name in enumerate(main_flow)}
            for name in visible_actions:
                if name in main_index:
                    continue

                related = [
                    item for item in visible_transitions
                    if (
                        item["from"] == name and item["to"] in main_index
                    ) or (
                        item["to"] == name and item["from"] in main_index
                    )
                ]
                if related:
                    strongest = max(related, key=lambda item: item["count"])
                    anchor_name = strongest["to"] if strongest["to"] in main_index else strongest["from"]
                    anchor_index = main_index[anchor_name]
                elif main_flow:
                    anchor_index = min(len(main_flow) - 1, len(branch_meta) % len(main_flow))
                else:
                    anchor_index = 0

                usage = anchor_usage.get(anchor_index, 0)
                side = -1 if usage % 2 == 0 else 1
                lane = usage // 2 + 1
                anchor_usage[anchor_index] = usage + 1
                branch_meta[name] = {
                    "anchor_index": anchor_index,
                    "side": side,
                    "lane": lane,
                }

            max_lane = max((meta["lane"] for meta in branch_meta.values()), default=0)
            main_y = max(220, 120 + max_lane * vertical_gap)
            node_positions = {}
            for index, name in enumerate(main_flow):
                node_positions[name] = (
                    left_margin + index * horizontal_gap,
                    main_y,
                )
            for name, meta in branch_meta.items():
                anchor_x = left_margin + meta["anchor_index"] * horizontal_gap
                branch_x = anchor_x + (meta["lane"] - 1) * 24
                branch_y = main_y + meta["side"] * meta["lane"] * vertical_gap
                node_positions[name] = (branch_x, branch_y)

            if not main_flow:
                for index, name in enumerate(visible_actions):
                    node_positions[name] = (
                        left_margin + index * horizontal_gap,
                        main_y,
                    )

            main_edges = {
                (main_flow[index], main_flow[index + 1])
                for index in range(max(0, len(main_flow) - 1))
            }

            def port_point(node_name, side):
                x, y = node_positions[node_name]
                offset = node_width / 2
                return (x + offset, y) if side == "right" else (x - offset, y)

            def preferred_port_pair(source_name, target_name):
                source_is_main = source_name in main_index
                target_is_main = target_name in main_index
                source_x, _ = node_positions[source_name]
                target_x, _ = node_positions[target_name]

                if (source_name, target_name) in main_edges:
                    return "right", "left"

                # Для основной ветви фиксируем "смысловую" ориентацию.
                source_options = ["right"] if source_is_main else ["right", "left"]
                target_options = ["left"] if target_is_main else ["left", "right"]

                # Боковой узел может входить и в правый край, если это короче и визуально честнее.
                best_pair = None
                best_cost = None
                for source_side in source_options:
                    for target_side in target_options:
                        start_x, start_y = port_point(source_name, source_side)
                        end_x, end_y = port_point(target_name, target_side)
                        distance = abs(end_x - start_x) + abs(end_y - start_y)

                        # Небольшой штраф за порт, который ведет стрелку "через" узел по горизонтали.
                        if source_side == "right" and end_x < start_x:
                            distance += node_width * 0.35
                        if source_side == "left" and end_x > start_x:
                            distance += node_width * 0.35
                        if target_side == "left" and start_x > end_x:
                            distance += node_width * 0.35
                        if target_side == "right" and start_x < end_x:
                            distance += node_width * 0.35

                        if best_cost is None or distance < best_cost:
                            best_cost = distance
                            best_pair = (source_side, target_side)

                if best_pair is not None:
                    return best_pair

                # Фолбэк остается геометрически понятным.
                return (
                    "right" if target_x >= source_x else "left",
                    "left" if target_x >= source_x else "right",
                )

            def path_anchors(source_name, target_name):
                source_side, target_side = preferred_port_pair(source_name, target_name)
                start_x, start_y = port_point(source_name, source_side)
                end_x, end_y = port_point(target_name, target_side)
                return start_x, start_y, end_x, end_y, source_side, target_side

            def add_label(text, x, y):
                label = scene.addText(str(text), font)
                label.setDefaultTextColor(QColor(65, 65, 65))
                bounds = label.boundingRect()
                label.setPos(x - bounds.width() / 2, y - bounds.height() / 2)

            def add_arrowhead(path, color, line_width):
                end_point = path.pointAtPercent(1.0)
                near_point = path.pointAtPercent(0.96)
                dx = end_point.x() - near_point.x()
                dy = end_point.y() - near_point.y()
                length = (dx * dx + dy * dy) ** 0.5
                if length <= 0:
                    return

                ux = dx / length
                uy = dy / length
                arrow_length = max(14, 10 + line_width)
                arrow_half_width = max(6, 4 + line_width / 2)
                base_x = end_point.x() - ux * arrow_length
                base_y = end_point.y() - uy * arrow_length
                left = QPointF(
                    base_x - uy * arrow_half_width,
                    base_y + ux * arrow_half_width,
                )
                right = QPointF(
                    base_x + uy * arrow_half_width,
                    base_y - ux * arrow_half_width,
                )
                polygon = QPolygonF([end_point, left, right])
                scene.addPolygon(polygon, QPen(color), QBrush(color))

            node_obstacles = []
            for node_name, (node_x, node_y) in node_positions.items():
                node_obstacles.append({
                    "name": node_name,
                    "left": node_x - node_width / 2 - 18,
                    "right": node_x + node_width / 2 + 18,
                    "top": node_y - node_height / 2 - 18,
                    "bottom": node_y + node_height / 2 + 18,
                })

            edge_lanes = {
                -1: [],
                1: [],
            }

            def ranges_overlap(left_a, right_a, left_b, right_b):
                return not (right_a < left_b or right_b < left_a)

            def preferred_route_side(start_y, end_y, transition_index):
                if start_y < main_y or end_y < main_y:
                    return -1
                if start_y > main_y or end_y > main_y:
                    return 1
                return -1 if transition_index % 2 == 0 else 1

            def lane_hits_node(lane_y, span_left, span_right, source_name, target_name):
                clearance = 34
                for obstacle in node_obstacles:
                    if obstacle["name"] in {source_name, target_name}:
                        continue
                    if not ranges_overlap(span_left, span_right, obstacle["left"], obstacle["right"]):
                        continue
                    if obstacle["top"] - clearance <= lane_y <= obstacle["bottom"] + clearance:
                        return True
                return False

            def lane_hits_edge(side, lane_y, span_left, span_right):
                for used_lane in edge_lanes[side]:
                    if abs(used_lane["y"] - lane_y) < 30 and ranges_overlap(
                        span_left,
                        span_right,
                        used_lane["left"],
                        used_lane["right"],
                    ):
                        return True
                return False

            def choose_route_lane(source_name, target_name, start_x, start_y, end_x, end_y, transition_index):
                span_left = min(start_x, end_x)
                span_right = max(start_x, end_x)
                preferred_side = preferred_route_side(start_y, end_y, transition_index)
                side_candidates = [preferred_side, -preferred_side]
                lane_step = float(getattr(self, "process_map_lane_spacing", vertical_gap * 1.15))

                # 1. Лучший случай: не пересекаем ни действия, ни уже уложенные связи.
                for side in side_candidates:
                    for lane_index in range(1, 8):
                        lane_y = main_y + side * (lane_index * lane_step)
                        if (
                            not lane_hits_node(lane_y, span_left, span_right, source_name, target_name)
                            and not lane_hits_edge(side, lane_y, span_left, span_right)
                        ):
                            edge_lanes[side].append({
                                "y": lane_y,
                                "left": span_left,
                                "right": span_right,
                            })
                            return lane_y

                # 2. Если идеально не вышло, разрешаем пересечь другую связь,
                #    но все равно не прокладываем маршрут через действия.
                for side in side_candidates:
                    for lane_index in range(1, 8):
                        lane_y = main_y + side * (lane_index * lane_step)
                        if not lane_hits_node(lane_y, span_left, span_right, source_name, target_name):
                            edge_lanes[side].append({
                                "y": lane_y,
                                "left": span_left,
                                "right": span_right,
                            })
                            return lane_y

                fallback_y = main_y + preferred_side * (8 * lane_step)
                edge_lanes[preferred_side].append({
                    "y": fallback_y,
                    "left": span_left,
                    "right": span_right,
                })
                return fallback_y

            for transition_index, transition in enumerate(visible_transitions):
                source = transition["from"]
                target = transition["to"]
                if source not in node_positions or target not in node_positions:
                    continue

                width = 1 + int(7 * transition["count"] / transition_max)
                color = QColor(74, 110, 155) if (source, target) in main_edges else QColor(118, 132, 150)
                pen = QPen(color, width)
                start_x, start_y, end_x, end_y, source_side, target_side = path_anchors(source, target)

                if source == target:
                    center_x, center_y = node_positions[source]
                    path = QPainterPath()
                    path.moveTo(center_x + node_width / 4, center_y - node_height / 2)
                    path.cubicTo(
                        center_x + node_width / 2,
                        center_y - vertical_gap,
                        center_x - node_width / 2,
                        center_y - vertical_gap,
                        center_x - node_width / 4,
                        center_y - node_height / 2,
                    )
                    scene.addPath(path, pen)
                    add_arrowhead(path, color, width)
                    add_label(transition["count"], center_x, center_y - vertical_gap * 0.72)
                elif (source, target) in main_edges:
                    path = QPainterPath()
                    path.moveTo(start_x, start_y)
                    path.lineTo(end_x, end_y)
                    scene.addPath(path, pen)
                    add_arrowhead(path, color, width)
                    add_label(transition["count"], (start_x + end_x) / 2, start_y - 18)
                else:
                    path = QPainterPath()
                    path.moveTo(start_x, start_y)
                    start_direction = 1 if source_side == "right" else -1
                    end_direction = -1 if target_side == "left" else 1
                    horizontal_span = max(70, min(180, abs(end_x - start_x) * 0.35))
                    start_control_x = start_x + start_direction * horizontal_span
                    end_control_x = end_x + end_direction * horizontal_span
                    route_y = choose_route_lane(
                        source,
                        target,
                        start_x,
                        start_y,
                        end_x,
                        end_y,
                        transition_index,
                    )
                    if start_y == end_y:
                        path.cubicTo(
                            start_control_x,
                            route_y,
                            end_control_x,
                            route_y,
                            end_x,
                            end_y,
                        )
                        add_label(transition["count"], (start_x + end_x) / 2, route_y)
                    else:
                        mid_x = (start_control_x + end_control_x) / 2
                        path.cubicTo(
                            start_control_x,
                            route_y,
                            end_control_x,
                            route_y,
                            end_x,
                            end_y,
                        )
                        add_label(transition["count"], mid_x, route_y)
                    scene.addPath(path, pen)
                    add_arrowhead(path, color, width)

            for name, (x, y) in node_positions.items():
                is_main = name in main_index
                fill = QColor(219, 238, 255) if is_main else QColor(235, 242, 248)
                border = QColor(40, 92, 138) if is_main else QColor(86, 109, 128)
                ellipse = scene.addEllipse(
                    x - node_width / 2,
                    y - node_height / 2,
                    node_width,
                    node_height,
                    QPen(border, 2),
                    QBrush(fill),
                )
                ellipse.setToolTip(name)

                display_name = metrics.elidedText(name, Qt.ElideRight, int(node_width - 28))
                text_item = scene.addText(display_name, font)
                text_item.setDefaultTextColor(QColor(30, 47, 62))
                bounds = text_item.boundingRect()
                text_item.setPos(x - bounds.width() / 2, y - bounds.height() / 2)

            x_values = [point[0] for point in node_positions.values()]
            y_values = [point[1] for point in node_positions.values()]
            content_rect = scene.itemsBoundingRect().adjusted(-80, -80, 80, 80)
            scene.setSceneRect(
                min(x_values) - node_width,
                min(y_values) - node_height - 80,
                max(x_values) - min(x_values) + node_width * 2,
                max(y_values) - min(y_values) + node_height * 2 + 160,
            )
            view.set_content_rect(scene.sceneRect().united(content_rect))
        else:
            scene.addText("Нет данных для построения карты процесса")
            view.set_content_rect(scene.itemsBoundingRect().adjusted(-80, -80, 80, 80))

    def show_statistics(self):
        """Показать вкладку статистики"""
        if not hasattr(self, 'json_file_path') or not self.json_file_path:
            QMessageBox.warning(self, 'Ошибка', 'Нет данных для анализа статистики')
            return

        if self._get_missing_records_cached():
            QMessageBox.warning(
                self,
                'Требуется очистка данных',
                'Перед статистикой обработайте пропуски в разделе качества данных.'
            )
            self.show_quality()
            return

        try:
            self.display_mode = 'statistics'

            # Меняем заголовок окна
            self.setWindowTitle('Анализ бизнес-процессов — Статистика')

            # Переключаем видимость кнопок
            self.btn_statistics.setVisible(False)
            self.btn_back_to_parsing.setVisible(True)
            self.btn_select_file.setVisible(False)
            self.btn_parse.setVisible(False)
            self.btn_select_json.setVisible(False)
            self.btn_clear.setVisible(False)
            self._set_navigation_mode('statistics')

            # Создаем виджет статистики если его нет или данные изменились
            self._build_statistics_page()

            # Переключаем на страницу статистики
            self.stacked_widget.setCurrentIndex(2)

        except Exception as e:
            QMessageBox.critical(
                self,
                'Ошибка статистики',
                f'Не удалось проанализировать статистику:\n\n{str(e)}'
            )

    def show_parsing(self):
        """Показать вкладку парсинга"""
        self.display_mode = 'parsing'

        # Меняем заголовок окна
        self.setWindowTitle('Анализ бизнес-процессов — Парсинг')

        # Переключаем видимость кнопок
        self.btn_statistics.setVisible(True)
        self.btn_back_to_parsing.setVisible(False)
        self.btn_select_file.setVisible(True)
        self.btn_parse.setVisible(True)
        self.btn_select_json.setVisible(True)
        self.btn_clear.setVisible(True)
        self._set_navigation_mode('parsing')

        # Переключаем на страницу парсинга
        self.stacked_widget.setCurrentIndex(0)

    def _build_statistics_page(self):
        """Создает или обновляет содержимое страницы статистики"""
        current_cache_key = self._json_cache_key()
        # Проверяем, нужно ли пересоздать виджет статистики
        if self.cached_statistics_data == current_cache_key and self.statistics_widget is not None:
            # Данные не изменились, используем кэш
            return

        # Очищаем текущее содержимое страницы статистики
        while self.statistics_layout.count():
            item = self.statistics_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        # Создаем новый виджет статистики
        self.statistics_widget = self.create_statistics_widget()
        self.statistics_layout.addWidget(self.statistics_widget)

        # Сохраняем путь к файлу как кэш
        self.cached_statistics_data = current_cache_key

    def move_action_up(self):
        """Перемещение выбранного действия вверх"""
        current_row = self.list_actions.currentRow()
        if current_row > 0:
            current_item = self.list_actions.takeItem(current_row)
            self.list_actions.insertItem(current_row - 1, current_item)
            self.list_actions.setCurrentRow(current_row - 1)
            self.update_custom_order()

    def move_action_down(self):
        """Перемещение выбранного действия вниз"""
        current_row = self.list_actions.currentRow()
        if current_row < self.list_actions.count() - 1 and current_row >= 0:
            current_item = self.list_actions.takeItem(current_row)
            self.list_actions.insertItem(current_row + 1, current_item)
            self.list_actions.setCurrentRow(current_row + 1)
            self.update_custom_order()

    def reset_action_order(self):
        """Сброс порядка действий к автоматическому или алфавитному"""
        if self.current_file:
            self.populate_actions_list(self.current_file)

    def toggle_auto_sort(self):
        """Обработка изменения состояния чекбокса автоматической сортировки"""
        if not self.current_file:
            return
        if self.chk_auto_sort.isChecked():
            self.populate_actions_list(self.current_file)
        else:
            self.update_custom_order()

    def update_custom_order(self):
        """Обновление пользовательского порядка действий"""
        self.custom_action_order = []
        for i in range(self.list_actions.count()):
            self.custom_action_order.append(self.list_actions.item(i).text())

    def select_json_file(self):
        """Выбор готового JSON файла для просмотра"""
        filepath, _ = QFileDialog.getOpenFileName(
            self,
            'Выбрать JSON файл',
            '',
            'JSON файлы (*.json);;Все файлы (*)'
        )

        if filepath:
            try:
                # Проверяем, что файл существует и является валидным JSON
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                # Проверяем структуру JSON
                if not isinstance(data, dict) or 'records' not in data or 'actions' not in data:
                    QMessageBox.warning(self, 'Ошибка', 'Некорректный формат JSON файла')
                    return

                # Устанавливаем файл для дальнейшего использования
                self.json_file_path = filepath
                self.json_result = json.dumps(data, indent=2, ensure_ascii=False)

                # Сбрасываем кэш статистики
                self.cached_statistics_data = None
                self.statistics_widget = None
                self.cached_quality_data = None
                self.cached_graphics_data = None
                self.cached_export_data = None
                self.quality_widget = None
                self.graphics_widget = None
                self.export_widget = None
                self._reset_analysis_cache()
                self.analysis_cache_key = self._json_cache_key()
                self.analysis_data = data

                # Очищаем страницу статистики
                while self.statistics_layout.count():
                    item = self.statistics_layout.takeAt(0)
                    widget = item.widget()
                    if widget:
                        widget.deleteLater()

                # Показываем информацию о файле
                self._clear_layout(self.quality_layout)
                self._clear_layout(self.graphics_layout)
                self._clear_layout(self.export_layout)
                self.display_json_file_info(filepath, data)

                # Включаем кнопку статистики
                self.btn_statistics.setEnabled(True)
                self.btn_to_quality.setEnabled(True)
                self._set_navigation_mode('parsing')

                QMessageBox.information(
                    self,
                    'Файл загружен',
                    f'JSON файл успешно загружен:\n{filepath}\n\nТеперь вы можете перейти к статистике.'
                )

            except json.JSONDecodeError:
                QMessageBox.critical(self, 'Ошибка', 'Некорректный JSON файл')
            except Exception as e:
                QMessageBox.critical(
                    self,
                    'Ошибка загрузки',
                    f'Не удалось загрузить файл:\n{str(e)}'
                )

    def display_json_file_info(self, filepath, data):
        """Отображение информации о загруженном JSON файле"""
        import os

        # Форматируем результат для отображения
        formatted_result = f"""# ИНФОРМАЦИЯ О JSON ФАЙЛЕ

## Общая информация
- **Файл**: {os.path.basename(filepath)}
- **Полный путь**: {filepath}
- **Всего записей**: {data['metadata']['total_records']}
- **Уникальных действий**: {data['metadata']['unique_actions']}

## Определенные столбцы
- **ID кейса**: {data['metadata']['columns_detected']['case_id_column']}
- **Действия**: {data['metadata']['columns_detected']['action_column']}
- **Дата/время**: {data['metadata']['columns_detected']['datetime_column']}

## Список действий
{chr(10).join(f'{i+1}. {action["name"]}{" (тупик)" if action.get("is_dead_end") else ""}' for i, action in enumerate(data['actions']))}

## Первые 1000 записей
"""

        # Добавляем первые 1000 записей
        records_to_show = data['records'][:1000]
        for i, record in enumerate(records_to_show):
            action_id = record.get('action_id')
            action_meta = (
                data['actions'][action_id]
                if isinstance(action_id, int) and 0 <= action_id < len(data['actions'])
                else {}
            )
            resolved_action_name = action_meta.get('name', 'Нет действия')
            is_dead_end = action_meta.get('is_dead_end', False)
            formatted_result += f"{i+1}. Case {record['case_id']} - {resolved_action_name}{' (тупик)' if is_dead_end else ''} - {record['datetime']}\n"

        if len(data['records']) > 1000:
            formatted_result += f"\n... и еще {len(data['records']) - 1000} записей\n"

        self.text_results.setText(formatted_result)

def main():
    """Запуск приложения"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Современный стиль

    # Настраиваем палитру для лучшего вида
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(240, 240, 240))
    app.setPalette(palette)

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
