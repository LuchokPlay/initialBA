#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Модуль статистического анализа данных бизнес-процессов.

Предоставляет комплексный анализ JSON данных, включая:
- Временные метрики процессов и действий
- Статистические распределения
- Анализ эффективности
- Метрики качества процессов
"""

import json
from datetime import datetime, timedelta
from typing import Dict, List, Any, Tuple
from collections import defaultdict, Counter
import statistics
from dataclasses import dataclass
from process_insights import common_final_action_ids

# Коэффициент для преобразования MAD в эквивалент стандартного отклонения
MAD_SCALE_FACTOR = 1.4826


@dataclass
class ProcessMetrics:
    """Метрики отдельного процесса"""
    case_id: str
    duration: timedelta
    action_count: int
    start_time: datetime
    end_time: datetime
    actions: List[str]


@dataclass
class ActionMetrics:
    """Метрики отдельного действия"""
    action_name: str
    total_occurrences: int
    avg_duration: timedelta
    min_duration: timedelta
    max_duration: timedelta
    duration_std: timedelta  # стандартное отклонение (от среднего)
    duration_mad: timedelta  # медианное отклонение (MAD × 1.4826)
    process_count: int  # в скольких процессах встречается


class StatisticsAnalyzer:
    """
    Анализатор статистики бизнес-процессов.

    Выполняет комплексный анализ JSON данных с расчетом
    временных и качественных метрик процессов.
    """

    def __init__(self, json_filepath: str):
        """
        Инициализация анализатора.

        Args:
            json_filepath: путь к JSON файлу с данными процессов
        """
        self.json_filepath = json_filepath
        self.data = None
        self.process_metrics = []
        self.action_metrics = {}

        self._load_data()
        self._analyze_processes()
        self._analyze_actions()

    def _load_data(self):
        """Загружает и парсит JSON данные"""
        try:
            with open(self.json_filepath, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
        except Exception as e:
            raise ValueError(f"Ошибка загрузки JSON файла: {e}")

        if not self.data or 'records' not in self.data:
            raise ValueError("Некорректный формат JSON файла")

    def _analyze_processes(self):
        """Анализирует метрики процессов"""
        # Группируем записи по case_id
        processes = defaultdict(list)
        for record in self.data['records']:
            processes[record['case_id']].append(record)

        # Анализируем каждый процесс
        for case_id, records in processes.items():
            # Сортируем по времени
            sorted_records = sorted(records, key=lambda x: x['datetime'])

            if len(sorted_records) < 2:
                continue  # Пропускаем процессы с одним действием

            start_time = datetime.fromisoformat(sorted_records[0]['datetime'])
            end_time = datetime.fromisoformat(sorted_records[-1]['datetime'])
            duration = end_time - start_time

            actions = [record['action_id'] for record in sorted_records]

            self.process_metrics.append(ProcessMetrics(
                case_id=case_id,
                duration=duration,
                action_count=len(actions),
                start_time=start_time,
                end_time=end_time,
                actions=actions
            ))

    def _analyze_actions(self):
        """Анализирует метрики действий"""
        if not self.data or 'actions' not in self.data:
            return

        action_names = {i: action['name'] for i, action in enumerate(self.data['actions'])}

        # Группируем записи по действиям
        action_durations = defaultdict(list)
        action_process_count = defaultdict(set)

        # Группируем по процессам для расчета длительностей
        processes = defaultdict(list)
        for record in self.data['records']:
            processes[record['case_id']].append(record)

        for case_id, records in processes.items():
            sorted_records = sorted(records, key=lambda x: x['datetime'])

            for i in range(len(sorted_records) - 1):
                current = sorted_records[i]
                next_record = sorted_records[i + 1]

                action_id = current['action_id']
                start_time = datetime.fromisoformat(current['datetime'])
                end_time = datetime.fromisoformat(next_record['datetime'])
                duration = end_time - start_time

                action_durations[action_id].append(duration)
                action_process_count[action_id].add(case_id)

        # Вычисляем метрики для каждого действия
        for action_id, durations in action_durations.items():
            if durations:
                avg_duration = sum(durations, timedelta()) / len(durations)
                min_duration = min(durations)
                max_duration = max(durations)

                # Вычисляем стандартное и медианное отклонение
                if len(durations) > 1:
                    duration_seconds = [d.total_seconds() for d in durations]
                    # Стандартное отклонение (от среднего)
                    std_seconds = statistics.stdev(duration_seconds)
                    duration_std = timedelta(seconds=std_seconds)
                    # Медианное отклонение (MAD × 1.4826)
                    median_seconds = statistics.median(duration_seconds)
                    mad_seconds = self._calculate_mad(duration_seconds, median_seconds) * MAD_SCALE_FACTOR
                    duration_mad = timedelta(seconds=mad_seconds)
                else:
                    duration_std = timedelta(0)
                    duration_mad = timedelta(0)

                self.action_metrics[action_names[action_id]] = ActionMetrics(
                    action_name=action_names[action_id],
                    total_occurrences=len(durations),
                    avg_duration=avg_duration,
                    min_duration=min_duration,
                    max_duration=max_duration,
                    duration_std=duration_std,
                    duration_mad=duration_mad,
                    process_count=len(action_process_count[action_id])
                )

    def _calculate_mad(self, values: List[float], median_value: float) -> float:
        """
        Вычисляет MAD (Median Absolute Deviation).
        
        MAD = median(|x_i - median(x)|)
        """
        if not values:
            return 0.0
        absolute_deviations = [abs(x - median_value) for x in values]
        return statistics.median(absolute_deviations)

    def get_process_statistics(self) -> Dict[str, Any]:
        """Возвращает статистику процессов"""
        if not self.process_metrics:
            return {}

        durations = [p.duration for p in self.process_metrics]
        action_counts = [p.action_count for p in self.process_metrics]

        # Вычисляем стандартное и медианное отклонение длительности процессов
        if len(durations) > 1:
            duration_seconds = [d.total_seconds() for d in durations]
            # Стандартное отклонение
            duration_std_seconds = statistics.stdev(duration_seconds)
            duration_std = self._format_timedelta(timedelta(seconds=duration_std_seconds))
            # Медианное отклонение (MAD × 1.4826)
            median_seconds = statistics.median(duration_seconds)
            duration_mad_seconds = self._calculate_mad(duration_seconds, median_seconds) * MAD_SCALE_FACTOR
            duration_mad = self._format_timedelta(timedelta(seconds=duration_mad_seconds))
        else:
            duration_std = "0сек"
            duration_mad = "0сек"

        # Вычисляем стандартное и медианное отклонение количества действий
        if len(action_counts) > 1:
            # Стандартное отклонение
            actions_std = round(statistics.stdev(action_counts), 2)
            # Медианное отклонение (MAD × 1.4826)
            median_actions = statistics.median(action_counts)
            actions_mad = round(self._calculate_mad(action_counts, median_actions) * MAD_SCALE_FACTOR, 2)
        else:
            actions_std = 0.0
            actions_mad = 0.0

        return {
            'total_processes': len(self.process_metrics),
            'avg_process_duration': self._format_timedelta(sum(durations, timedelta()) / len(durations)),
            'min_process_duration': self._format_timedelta(min(durations)),
            'max_process_duration': self._format_timedelta(max(durations)),
            'std_process_duration': duration_std,
            'mad_process_duration': duration_mad,
            'avg_actions_per_process': round(sum(action_counts) / len(action_counts), 2),
            'min_actions_per_process': min(action_counts),
            'max_actions_per_process': max(action_counts),
            'std_actions_per_process': actions_std,
            'mad_actions_per_process': actions_mad,
            'total_actions': sum(action_counts),
        }

    def get_action_statistics(self) -> List[Dict[str, Any]]:
        """Возвращает статистику действий"""
        stats = []
        for action_name, metrics in self.action_metrics.items():
            stats.append({
                'action': action_name,
                'occurrences': metrics.total_occurrences,
                'processes': metrics.process_count,
                'avg_duration': self._format_timedelta(metrics.avg_duration),
                'min_duration': self._format_timedelta(metrics.min_duration),
                'max_duration': self._format_timedelta(metrics.max_duration),
                'duration_std': self._format_timedelta(metrics.duration_std),
                'duration_mad': self._format_timedelta(metrics.duration_mad),
                'efficiency_score': self._calculate_action_efficiency(metrics)
            })

        # Сортируем по количеству выполнений
        stats.sort(key=lambda x: x['occurrences'], reverse=True)
        return stats

    def get_time_distribution(self) -> Dict[str, Any]:
        """Возвращает распределение по времени выполнения (динамические интервалы)"""
        if not self.process_metrics:
            return {}

        durations_hours = sorted([p.duration.total_seconds() / 3600 for p in self.process_metrics])

        if not durations_hours:
            return {}

        # Вычисляем динамические интервалы на основе данных
        min_duration = min(durations_hours)
        max_duration = max(durations_hours)

        # Создаем 5-6 интервалов
        range_size = max_duration - min_duration
        if range_size == 0:
            # Все процессы одинаковой длительности
            interval_size = 1
        else:
            interval_size = range_size / 5

        distribution = {}
        current_start = min_duration

        for i in range(6):  # 6 интервалов
            if i == 5:  # Последний интервал
                current_end = float('inf')
                label = f">{current_start:.1f}ч"
                count = sum(1 for d in durations_hours if d >= current_start)
            else:
                current_end = current_start + interval_size
                if current_end >= max_duration:
                    current_end = float('inf')
                    label = f">{current_start:.1f}ч"
                    count = sum(1 for d in durations_hours if d >= current_start)
                else:
                    label = f"{current_start:.1f}-{current_end:.1f}ч"
                    count = sum(1 for d in durations_hours if current_start <= d < current_end)

            distribution[label] = {
                'count': count,
                'percentage': round(count / len(durations_hours) * 100, 1)
            }

            if current_end == float('inf'):
                break

            current_start = current_end

        return distribution

    def get_process_flow_analysis(self) -> Dict[str, Any]:
        """Анализирует поток процессов"""
        if not self.data or 'records' not in self.data:
            return {}

        # Анализируем последовательности
        transitions = defaultdict(int)
        action_counts = defaultdict(int)

        processes = defaultdict(list)
        for record in self.data['records']:
            processes[record['case_id']].append(record)

        for case_id, records in processes.items():
            sorted_records = sorted(records, key=lambda x: x['datetime'])
            action_ids = [r['action_id'] for r in sorted_records]

            # Считаем переходы
            for i in range(len(action_ids) - 1):
                transition = (action_ids[i], action_ids[i + 1])
                transitions[transition] += 1

            # Считаем действия
            for action_id in action_ids:
                action_counts[action_id] += 1

        # Преобразуем в читаемые названия
        action_names = {i: action['name'] for i, action in enumerate(self.data['actions'])}

        top_transitions = []
        for (from_id, to_id), count in transitions.items():
            top_transitions.append({
                'from': action_names.get(from_id, f'Action_{from_id}'),
                'to': action_names.get(to_id, f'Action_{to_id}'),
                'count': count,
                'percentage': round(count / len(processes) * 100, 1)
            })

        top_transitions.sort(key=lambda x: x['count'], reverse=True)

        return {
            'total_transitions': len(transitions),
            'unique_transitions': len(set(transitions.keys())),
            'top_transitions': top_transitions[:10],  # Топ 10 переходов
            'action_frequency': [
                {
                    'action': action_names.get(action_id, f'Action_{action_id}'),
                    'count': count,
                    'percentage': round(count / sum(action_counts.values()) * 100, 1)
                }
                for action_id, count in sorted(action_counts.items(), key=lambda x: x[1], reverse=True)[:10]
            ]
        }

    def get_quality_metrics(self) -> Dict[str, Any]:
        """Возвращает метрики качества процессов"""
        if not self.process_metrics or not self.data:
            return {}

        final_action_ids = common_final_action_ids(self.data)
        unfinished_processes = 0
        for process in self.process_metrics:
            last_action_id = process.actions[-1]
            if final_action_ids and last_action_id not in final_action_ids:
                unfinished_processes += 1

        # Вычисляем метрики вариабельности
        durations_hours = [p.duration.total_seconds() / 3600 for p in self.process_metrics]
        if len(durations_hours) > 1:
            duration_std = statistics.stdev(durations_hours)
            duration_cv = duration_std / statistics.mean(durations_hours)  # Коэффициент вариации
        else:
            duration_cv = 0

        return {
            'completion_rate': round((len(self.process_metrics) - unfinished_processes) / len(self.process_metrics) * 100, 1) if self.process_metrics else 0,
            'dead_end_rate': round(unfinished_processes / len(self.process_metrics) * 100, 1) if self.process_metrics else 0,
            'duration_variability': round(duration_cv * 100, 1),  # в процентах
            'process_consistency': round((1 - duration_cv) * 100, 1),  # в процентах
        }

    def find_failed_processes(self) -> List[Dict[str, Any]]:
        """
        Находит процессы, которые не были завершены общепринятым конечным действием.

        Returns:
            список словарей с информацией о незавершенных процессах
        """
        if not self.data.get('actions'):
            return []

        correct_end_action_ids = common_final_action_ids(self.data)
        if not correct_end_action_ids:
            return []

        action_names = {i: action['name'] for i, action in enumerate(self.data['actions'])}
        correct_end_action_label = ", ".join(
            action_names.get(action_id, f'Action_{action_id}')
            for action_id in sorted(correct_end_action_ids)
        )

        # Анализируем каждый процесс
        failed_processes = []

        # Группируем записи по процессам
        processes = defaultdict(list)
        for record in self.data['records']:
            processes[record['case_id']].append(record)

        for case_id, records in processes.items():
            # Сортируем записи по времени
            sorted_records = sorted(records, key=lambda x: x['datetime'])

            if not sorted_records:
                continue

            # Проверяем, заканчивается ли процесс правильным конечным действием
            last_record = sorted_records[-1]
            last_action_id = last_record['action_id']

            if last_action_id not in correct_end_action_ids:
                # Это незавершенный процесс
                first_record = sorted_records[0]
                failed_processes.append({
                    'case_id': case_id,
                    'start_date': first_record['datetime'],
                    'last_date': last_record['datetime'],
                    'last_action': action_names.get(last_action_id, f'Action_{last_action_id}'),
                    'correct_end_action': correct_end_action_label,
                    'actions_count': len(sorted_records)
                })

        # Сортируем по дате начала (сначала самые свежие)
        failed_processes.sort(key=lambda x: x['start_date'], reverse=True)

        return failed_processes

    def get_comprehensive_report(self) -> str:
        """Генерирует полный отчет по всем метрикам"""
        report = []
        report.append("# СТАТИСТИЧЕСКИЙ ОТЧЕТ ПО БИЗНЕС-ПРОЦЕССАМ")
        report.append("")

        # Основная информация
        report.append("## ОСНОВНАЯ ИНФОРМАЦИЯ")
        process_stats = self.get_process_statistics()
        if process_stats:
            report.append(f"- Всего процессов: {process_stats['total_processes']}")
            report.append(f"- Всего действий: {process_stats['total_actions']}")
            report.append(f"- Средняя длительность процесса: {process_stats['avg_process_duration']}")
            report.append(f"- Среднее количество действий в процессе: {process_stats['avg_actions_per_process']}")
        report.append("")

        # Временные метрики
        report.append("## ВРЕМЕННЫЕ МЕТРИКИ")
        if process_stats:
            report.append(f"- Минимальная длительность: {process_stats['min_process_duration']}")
            report.append(f"- Максимальная длительность: {process_stats['max_process_duration']}")
        report.append("")

        # Распределение по времени
        report.append("## РАСПРЕДЕЛЕНИЕ ПРОЦЕССОВ ПО ВРЕМЕНИ ВЫПОЛНЕНИЯ")
        time_dist = self.get_time_distribution()
        for interval, data in time_dist.items():
            report.append(f"- {interval}: {data['count']} процессов ({data['percentage']}%)")
        report.append("")

        # Топ действий
        report.append("## ТОП ДЕЙСТВИЙ ПО ЧАСТОТЕ")
        flow_analysis = self.get_process_flow_analysis()
        for i, action in enumerate(flow_analysis.get('action_frequency', [])[:5], 1):
            report.append(f"{i}. {action['action']}: {action['count']} раз ({action['percentage']}%)")
        report.append("")

        # Топ переходов
        report.append("## НАИБОЛЕЕ ЧАСТЫЕ ПЕРЕХОДЫ")
        for i, transition in enumerate(flow_analysis.get('top_transitions', [])[:5], 1):
            report.append(f"{i}. {transition['from']} → {transition['to']}: {transition['count']} раз ({transition['percentage']}%)")
        report.append("")

        # Детальная статистика действий
        report.append("## ДЕТАЛЬНАЯ СТАТИСТИКА ДЕЙСТВИЙ")
        action_stats = self.get_action_statistics()
        if action_stats:
            # Исправленная таблица с правильными отступами
            report.append("| Действие | Выполнений | Процессов | Ср. длительность |")
            report.append("|----------|------------|-----------|------------------|")
            for stat in action_stats[:15]:  # Показываем топ 15
                report.append(f"| {stat['action'][:30]:<30} | {stat['occurrences']:<10} | {stat['processes']:<9} | {stat['avg_duration']:<16} |")
        report.append("")

        return "\n".join(report)

    def _format_timedelta(self, td: timedelta) -> str:
        """Форматирует timedelta в читаемый вид"""
        total_seconds = int(td.total_seconds())

        if total_seconds < 60:
            return f"{total_seconds}сек"
        elif total_seconds < 3600:
            minutes = total_seconds // 60
            seconds = total_seconds % 60
            return f"{minutes}мин {seconds}сек"
        elif total_seconds < 86400:
            hours = total_seconds // 3600
            minutes = (total_seconds % 3600) // 60
            return f"{hours}ч {minutes}мин"
        else:
            days = total_seconds // 86400
            hours = (total_seconds % 86400) // 3600
            return f"{days}д {hours}ч"

    def _calculate_action_efficiency(self, metrics: ActionMetrics) -> float:
        """Вычисляет коэффициент эффективности действия"""
        if metrics.total_occurrences == 0:
            return 0.0

        # Эффективность = 1 / (нормализованное отклонение + 1)
        # Чем меньше отклонение, тем выше эффективность
        avg_seconds = metrics.avg_duration.total_seconds()
        std_seconds = metrics.duration_std.total_seconds()

        if avg_seconds == 0:
            return 1.0

        cv = std_seconds / avg_seconds  # коэффициент вариации
        efficiency = 1.0 / (cv + 1.0)  # нормализуем к [0, 1]

        return round(efficiency, 2)


def analyze_statistics(json_filepath: str) -> str:
    """
    Основная функция для анализа статистики.

    Args:
        json_filepath: путь к JSON файлу

    Returns:
        str: полный отчет в формате Markdown
    """
    try:
        analyzer = StatisticsAnalyzer(json_filepath)
        return analyzer.get_comprehensive_report()
    except Exception as e:
        return f"# ❌ ОШИБКА АНАЛИЗА\n\nПроизошла ошибка при анализе файла: {e}"
