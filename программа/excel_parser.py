import pandas as pd
import re
import json
from datetime import datetime
from sequence_analyzer import sort_actions_automatically

# Конфигурация парсера
CASE_NUMBER_PATTERN = re.compile(r'(\d+)')

# Возможные форматы дат
DATETIME_FORMATS = [
    '%d.%m.%Y  %I:%M:%S',  # 02.05.2021  6:44:00
    '%Y-%m-%d %H:%M:%S',   # 2021-04-18 06:44:00
    '%d.%m.%Y %H:%M:%S',   # 02.05.2021 06:44:00
    '%Y-%m-%d  %I:%M:%S',  # 2021-04-18  6:44:00
    '%d.%m.%Y %H:%M',      # 02.05.2021 06:44
    '%Y-%m-%d %H:%M',      # 2021-04-18 06:44
    '%d.%m.%Y',            # 02.05.2021
    '%Y-%m-%d',            # 2021-04-18
    '%d %m %Y %H:%M:%S',   # 02 05 2021 06:44:00
    '%d %m %Y %H:%M',      # 02 05 2021 06:44
    '%d %m %Y',            # 02 05 2021
    '%Y %m %d %H:%M:%S',   # 2021 05 02 06:44:00
    '%Y %m %d %H:%M',      # 2021 05 02 06:44
    '%Y %m %d',            # 2021 05 02
    '%d/%m/%Y %H:%M:%S',   # 02/05/2021 06:44:00
    '%d/%m/%Y %H:%M',      # 02/05/2021 06:44
    '%d/%m/%Y',            # 02/05/2021
    '%Y/%m/%d %H:%M:%S',   # 2021/05/02 06:44:00
    '%Y/%m/%d %H:%M',      # 2021/05/02 06:44
    '%Y/%m/%d',            # 2021/05/02
    '%d-%m-%Y %H:%M:%S',   # 02-05-2021 06:44:00
    '%d-%m-%Y %H:%M',      # 02-05-2021 06:44
    '%d-%m-%Y',            # 02-05-2021
    '%Y-%m-%dT%H:%M:%S',   # 2021-05-02T06:44:00
    '%Y-%m-%dT%H:%M',      # 2021-05-02T06:44
]

# Возможные названия столбцов для автоматического определения
CASE_ID_HEADERS = ['case id', 'case_id', 'id кейса', 'caseid', 'кейс', 'case']
ACTION_HEADERS = ['статус', 'действие', 'название действия', 'action', 'status', 'операция']
DATETIME_HEADERS = ['дата', 'время', 'дата начала', 'дата и время', 'datetime', 'date', 'time']

class ExcelParseError(Exception):
    """Исключение для ошибок парсинга Excel файла"""
    pass

def normalize_cell_value(value):
    """Keeps missing Excel cells empty instead of turning them into 'nan'."""
    if pd.isna(value):
        return ''
    return str(value).strip()

def parse_datetime(dt_raw):
    """
    Парсит строку даты/времени, пытаясь разные форматы.

    Args:
        dt_raw: строка с датой/временем

    Returns:
        datetime: объект datetime

    Raises:
        ExcelParseError: если ни один формат не подошел
    """
    if dt_raw is None:
        return None

    dt_value = str(dt_raw).strip()
    if dt_value.lower() in {'', 'nan', 'none', 'null', 'nat'}:
        return None
    normalized_dt_value = ' '.join(dt_value.split())

    try:
        return datetime.fromisoformat(normalized_dt_value)
    except ValueError:
        pass

    for fmt in DATETIME_FORMATS:
        try:
            return datetime.strptime(normalized_dt_value, fmt)
        except ValueError:
            continue

    # Если ни один формат не подошел
    raise ExcelParseError(f'Некорректная дата/время: "{dt_raw}". '
                         f'Поддерживаемые форматы: {DATETIME_FORMATS}')

def detect_column_indices(headers):
    """
    Автоматически определяет индексы столбцов по их названиям.

    Args:
        headers: список названий столбцов

    Returns:
        dict: {'case_id': int, 'action': int, 'datetime': int}
    """
    headers_lower = [h.strip().lower() for h in headers]

    # Ищем индексы столбцов
    case_id_idx = None
    action_idx = None
    datetime_idx = None

    for i, header in enumerate(headers_lower):
        # Определяем case_id столбец
        if case_id_idx is None and any(keyword in header for keyword in CASE_ID_HEADERS):
            case_id_idx = i
        # Определяем action столбец
        elif action_idx is None and any(keyword in header for keyword in ACTION_HEADERS):
            action_idx = i
        # Определяем datetime столбец
        elif datetime_idx is None and any(keyword in header for keyword in DATETIME_HEADERS):
            datetime_idx = i

    # Если не нашли по названиям, используем позиционную логику (первые 3 столбца)
    if case_id_idx is None and len(headers) > 0:
        case_id_idx = 0
    if action_idx is None and len(headers) > 1:
        action_idx = 1
    if datetime_idx is None and len(headers) > 2:
        datetime_idx = 2

    # Проверяем, что нашли все необходимые столбцы
    if case_id_idx is None:
        raise ExcelParseError("Не найден столбец с ID кейса")
    if action_idx is None:
        raise ExcelParseError("Не найден столбец с действиями/статусом")
    if datetime_idx is None:
        raise ExcelParseError("Не найден столбец с датой/временем")

    return {
        'case_id': case_id_idx,
        'action': action_idx,
        'datetime': datetime_idx
    }

def parse_excel_to_json(filepath, custom_action_order=None):
    """
    Парсит Excel файл и преобразует в JSON формат.
    Автоматически определяет структуру файла по названиям столбцов.

    Args:
        filepath: путь к Excel файлу
        custom_action_order: список действий в желаемом порядке (опционально)

    Returns:
        str: JSON-строка с оптимизированной структурой данных

    Raises:
        ExcelParseError: при ошибках чтения или парсинга файла
    """
    try:
        # Читаем Excel файл
        df = pd.read_excel(filepath, dtype=str)
    except Exception as e:
        raise ExcelParseError(f'Ошибка при чтении Excel файла: {e}')

    # Проверяем, что есть хотя бы 3 столбца
    if len(df.columns) < 3:
        raise ExcelParseError(f'В файле должно быть минимум 3 столбца, найдено: {len(df.columns)}')

    # Определяем индексы столбцов
    column_indices = detect_column_indices(df.columns)

    records = []
    actions_set = set()

    # Собираем все уникальные действия и создаем записи
    for idx, row in df.iterrows():
        case_raw = normalize_cell_value(row.iloc[column_indices['case_id']])
        action = normalize_cell_value(row.iloc[column_indices['action']])
        dt_raw = normalize_cell_value(row.iloc[column_indices['datetime']])

        # Парсинг id из любой строки, содержащей число: Case_12, 12, "Кейс №12" и т.п.
        m = CASE_NUMBER_PATTERN.search(case_raw) if case_raw else None
        if case_raw and not m:
            raise ExcelParseError(f'Некорректный ID кейса в строке {idx+2}: "{case_raw}". Ожидается значение, содержащее число')

        case_id_num = int(m.group(1)) if m else None
        if action:
            actions_set.add(action)

        # Парсинг даты
        try:
            dt = parse_datetime(dt_raw)
        except ExcelParseError:
            raise  # Передаем исключение дальше

        records.append({
            'case_id': case_id_num,
            'case_raw': case_raw,
            'datetime': dt.isoformat() if dt else None,
            'action': action  # пока оставляем название, потом заменим на ID
        })

    # Проверяем, что все действия из records есть в actions_set
    record_actions = set(record['action'] for record in records if record['action'])
    missing_actions = record_actions - set(actions_set)
    if missing_actions:
        raise ExcelParseError(f"Найдены действия в записях, которых нет в наборе действий: {missing_actions}")

    # Сортировка действий: пользовательский порядок, автоматический или алфавит
    if custom_action_order:
        # Используем пользовательский порядок, добавляя в конец не указанные действия
        actions_list = []
        for action in custom_action_order:
            if action in actions_set:
                actions_list.append(action)
        # Добавляем оставшиеся действия по алфавиту
        remaining_actions = sorted(actions_set - set(actions_list))
        actions_list.extend(remaining_actions)
    else:
        # Пытаемся автоматически определить порядок на основе анализа последовательностей
        try:
            actions_list = [
                action for action in sort_actions_automatically(records)
                if action in actions_set
            ]
            actions_list.extend(sorted(actions_set - set(actions_list)))
            if not actions_list:
                # Если автоматическая сортировка не удалась, используем алфавит
                actions_list = sorted(list(actions_set))
        except Exception as e:
            # В случае ошибки используем алфавитную сортировку
            print(f"Warning: Automatic sorting failed: {e}")  # Временная отладка
            actions_list = sorted(list(actions_set))

    action_map = {action: idx for idx, action in enumerate(actions_list)}

    # Заменяем названия действий на ID в записях
    for record in records:
        try:
            action_name = record.pop('action')
            if not action_name:
                record['action_id'] = None
                continue
            if action_name not in action_map:
                raise ExcelParseError(f"Действие '{action_name}' не найдено в списке действий")
            record['action_id'] = action_map[action_name]
        except KeyError:
            raise ExcelParseError(f"В записи отсутствует поле 'action': {record}")

    # Получаем статистику последовательностей для метаданных
    sequence_stats = {}
    dead_end_actions = []
    try:
        from sequence_analyzer import SequenceAnalyzer
        analyzer = SequenceAnalyzer(records)
        sequence_stats = analyzer.get_statistics()
        # Определяем тупиковые действия (действия-возвраты)
        dead_end_actions = analyzer._find_dead_ends()
    except Exception:
        pass  # В случае ошибки оставляем пустую статистику

    # Создаем список действий с метаданными
    actions_with_metadata = []
    for action in actions_list:
        actions_with_metadata.append({
            'name': action,
            'is_dead_end': action in dead_end_actions
        })

    # Создаем финальную JSON структуру
    result = {
        'actions': actions_with_metadata,
        'records': records,
        'metadata': {
            'total_records': len(records),
            'unique_actions': len(actions_list),
            'columns_detected': {
                'case_id_column': df.columns[column_indices['case_id']],
                'action_column': df.columns[column_indices['action']],
                'datetime_column': df.columns[column_indices['datetime']]
            },
            'sequence_analysis': sequence_stats
        }
    }

    return json.dumps(result, indent=2, ensure_ascii=False)

def get_excel_info(filepath):
    """
    Получает информацию о структуре Excel файла без полного парсинга.

    Args:
        filepath: путь к Excel файлу

    Returns:
        dict: информация о файле
    """
    try:
        df = pd.read_excel(filepath, dtype=str)
        column_indices = detect_column_indices(df.columns)

        return {
            'columns': list(df.columns),
            'total_rows': len(df),
            'detected_columns': {
                'case_id': df.columns[column_indices['case_id']],
                'action': df.columns[column_indices['action']],
                'datetime': df.columns[column_indices['datetime']]
            },
            'sample_data': {
                'first_row': df.iloc[0].to_dict() if len(df) > 0 else None
            }
        }
    except Exception as e:
        raise ExcelParseError(f'Ошибка при анализе файла: {e}')
