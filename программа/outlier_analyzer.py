#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Модуль поиска выбросов (аномалий) в бизнес-процессах.

Выявляет процессы и действия, длительность которых значительно
отличается от нормы (по правилу 3-х сигм от среднего).
"""

import json
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple
from collections import defaultdict
import statistics
from dataclasses import dataclass, field

# Коэффициент для преобразования MAD в эквивалент стандартного отклонения
MAD_SCALE_FACTOR = 1.4826


@dataclass
class ActionOutlier:
    """Информация о выбросе в действии"""
    action_id: int
    action_name: str
    duration: timedelta
    avg_duration: timedelta  # среднее
    median_duration: timedelta  # медиана
    std_duration: timedelta  # стандартное отклонение
    mad_duration: timedelta  # медианное отклонение (MAD × 1.4826)
    deviation_abs: timedelta  # абсолютное отклонение от среднего
    deviation_percent: float  # отклонение в процентах
    deviation_sigma: float  # отклонение в сигмах
    is_too_long: bool  # True если слишком долго, False если слишком быстро


@dataclass
class ProcessOutlier:
    """Информация о процессе-выбросе"""
    case_id: int
    duration: timedelta
    avg_duration: timedelta  # среднее по всем процессам
    median_duration: timedelta  # медиана по всем процессам
    std_duration: timedelta  # стандартное отклонение
    mad_duration: timedelta  # медианное отклонение (MAD × 1.4826)
    deviation_abs: timedelta  # абсолютное отклонение
    deviation_percent: float  # отклонение в процентах
    deviation_sigma: float  # отклонение в сигмах
    is_too_long: bool  # True если слишком долго, False если слишком быстро
    start_time: datetime
    end_time: datetime
    action_count: int
    action_outliers: List[ActionOutlier] = field(default_factory=list)


class OutlierAnalyzer:
    """
    Анализатор выбросов в бизнес-процессах.
    
    Использует правило N-сигм для определения аномалий:
    значение считается выбросом, если оно отклоняется от среднего
    более чем на N стандартных отклонений.
    """

    def __init__(self, json_filepath: str, sigma_threshold: float = 3.0):
        """
        Инициализация анализатора.

        Args:
            json_filepath: путь к JSON файлу с данными процессов
            sigma_threshold: порог в сигмах для определения выброса (по умолчанию 3)
        """
        self.json_filepath = json_filepath
        self.sigma_threshold = sigma_threshold
        self.data = None
        
        # Статистика по процессам
        self.process_durations: List[timedelta] = []
        self.process_avg: timedelta = timedelta(0)
        self.process_median: timedelta = timedelta(0)
        self.process_std: timedelta = timedelta(0)
        self.process_mad: timedelta = timedelta(0)  # медианное отклонение
        
        # Статистика по действиям (action_id -> list of durations)
        self.action_durations: Dict[int, List[timedelta]] = defaultdict(list)
        self.action_stats: Dict[int, Dict[str, timedelta]] = {}
        
        # Данные процессов
        self.processes_data: Dict[int, Dict] = {}
        
        self._load_data()
        self._calculate_statistics()

    def _load_data(self):
        """Загружает и парсит JSON данные"""
        try:
            with open(self.json_filepath, 'r', encoding='utf-8') as f:
                self.data = json.load(f)
        except Exception as e:
            raise ValueError(f"Ошибка загрузки JSON файла: {e}")

        if not self.data or 'records' not in self.data:
            raise ValueError("Некорректный формат JSON файла")

    def _calculate_statistics(self):
        """Вычисляет статистику по процессам и действиям"""
        # Группируем записи по case_id
        processes = defaultdict(list)
        for record in self.data['records']:
            processes[record['case_id']].append(record)

        # Анализируем каждый процесс
        for case_id, records in processes.items():
            sorted_records = sorted(records, key=lambda x: x['datetime'])

            if len(sorted_records) < 2:
                continue

            start_time = datetime.fromisoformat(sorted_records[0]['datetime'])
            end_time = datetime.fromisoformat(sorted_records[-1]['datetime'])
            duration = end_time - start_time

            self.process_durations.append(duration)
            
            # Сохраняем данные процесса
            self.processes_data[case_id] = {
                'duration': duration,
                'start_time': start_time,
                'end_time': end_time,
                'records': sorted_records,
                'action_count': len(sorted_records)
            }

            # Считаем длительности действий
            for i in range(len(sorted_records) - 1):
                current = sorted_records[i]
                next_record = sorted_records[i + 1]
                
                action_id = current['action_id']
                action_start = datetime.fromisoformat(current['datetime'])
                action_end = datetime.fromisoformat(next_record['datetime'])
                action_duration = action_end - action_start
                
                self.action_durations[action_id].append(action_duration)

        # Вычисляем статистику по процессам
        if self.process_durations:
            duration_seconds = [d.total_seconds() for d in self.process_durations]
            
            avg_seconds = statistics.mean(duration_seconds)
            self.process_avg = timedelta(seconds=avg_seconds)
            
            median_seconds = statistics.median(duration_seconds)
            self.process_median = timedelta(seconds=median_seconds)
            
            if len(duration_seconds) > 1:
                # Стандартное отклонение (от среднего)
                std_seconds = statistics.stdev(duration_seconds)
                self.process_std = timedelta(seconds=std_seconds)
                # Медианное отклонение (MAD × 1.4826)
                mad_seconds = self._calculate_mad(duration_seconds, median_seconds) * MAD_SCALE_FACTOR
                self.process_mad = timedelta(seconds=mad_seconds)

        # Вычисляем статистику по действиям
        for action_id, durations in self.action_durations.items():
            if durations:
                duration_seconds = [d.total_seconds() for d in durations]
                
                avg_seconds = statistics.mean(duration_seconds)
                median_seconds = statistics.median(duration_seconds)
                
                if len(duration_seconds) > 1:
                    std_seconds = statistics.stdev(duration_seconds)
                    mad_seconds = self._calculate_mad(duration_seconds, median_seconds) * MAD_SCALE_FACTOR
                else:
                    std_seconds = 0
                    mad_seconds = 0
                
                self.action_stats[action_id] = {
                    'avg': timedelta(seconds=avg_seconds),
                    'median': timedelta(seconds=median_seconds),
                    'std': timedelta(seconds=std_seconds),
                    'mad': timedelta(seconds=mad_seconds)
                }

    def _calculate_mad(self, values: List[float], median_value: float) -> float:
        """Вычисляет MAD (Median Absolute Deviation)."""
        if not values:
            return 0.0
        absolute_deviations = [abs(x - median_value) for x in values]
        return statistics.median(absolute_deviations)

    def get_action_name(self, action_id: int) -> str:
        """Возвращает название действия по ID"""
        if self.data and 'actions' in self.data:
            if action_id < len(self.data['actions']):
                return self.data['actions'][action_id]['name']
        return f"Action_{action_id}"

    def is_outlier(self, value: timedelta, avg: timedelta, std: timedelta) -> Tuple[bool, float]:
        """
        Проверяет, является ли значение выбросом.
        
        Returns:
            Tuple[bool, float]: (является ли выбросом, отклонение в сигмах)
        """
        if std.total_seconds() == 0:
            return False, 0.0
        
        deviation_sigma = abs(value.total_seconds() - avg.total_seconds()) / std.total_seconds()
        is_outlier = deviation_sigma >= self.sigma_threshold
        
        return is_outlier, deviation_sigma

    def find_process_outliers(self) -> List[ProcessOutlier]:
        """
        Находит процессы-выбросы.
        
        Returns:
            Список ProcessOutlier с информацией о выбросах
        """
        outliers = []
        
        if self.process_std.total_seconds() == 0:
            return outliers
        
        for case_id, proc_data in self.processes_data.items():
            duration = proc_data['duration']
            
            is_out, deviation_sigma = self.is_outlier(
                duration, self.process_avg, self.process_std
            )
            
            if is_out:
                deviation_abs = duration - self.process_avg
                is_too_long = duration > self.process_avg
                
                # Вычисляем процентное отклонение
                if self.process_avg.total_seconds() > 0:
                    deviation_percent = (deviation_abs.total_seconds() / self.process_avg.total_seconds()) * 100
                else:
                    deviation_percent = 0.0
                
                # Находим выбросы в действиях этого процесса
                action_outliers = self._find_action_outliers_in_process(proc_data['records'])
                
                outlier = ProcessOutlier(
                    case_id=case_id,
                    duration=duration,
                    avg_duration=self.process_avg,
                    median_duration=self.process_median,
                    std_duration=self.process_std,
                    mad_duration=self.process_mad,
                    deviation_abs=deviation_abs,
                    deviation_percent=deviation_percent,
                    deviation_sigma=deviation_sigma,
                    is_too_long=is_too_long,
                    start_time=proc_data['start_time'],
                    end_time=proc_data['end_time'],
                    action_count=proc_data['action_count'],
                    action_outliers=action_outliers
                )
                outliers.append(outlier)
        
        # Сортируем по абсолютному отклонению в сигмах (наибольшие отклонения первыми)
        outliers.sort(key=lambda x: x.deviation_sigma, reverse=True)
        
        return outliers

    def _find_action_outliers_in_process(self, records: List[Dict]) -> List[ActionOutlier]:
        """
        Находит действия-выбросы в конкретном процессе.
        
        Args:
            records: отсортированные записи процесса
            
        Returns:
            Список ActionOutlier
        """
        action_outliers = []
        
        for i in range(len(records) - 1):
            current = records[i]
            next_record = records[i + 1]
            
            action_id = current['action_id']
            action_start = datetime.fromisoformat(current['datetime'])
            action_end = datetime.fromisoformat(next_record['datetime'])
            action_duration = action_end - action_start
            
            # Получаем статистику для этого действия
            if action_id not in self.action_stats:
                continue
                
            stats = self.action_stats[action_id]
            
            if stats['std'].total_seconds() == 0:
                continue
            
            is_out, deviation_sigma = self.is_outlier(
                action_duration, stats['avg'], stats['std']
            )
            
            if is_out:
                deviation_abs = action_duration - stats['avg']
                is_too_long = action_duration > stats['avg']
                
                if stats['avg'].total_seconds() > 0:
                    deviation_percent = (deviation_abs.total_seconds() / stats['avg'].total_seconds()) * 100
                else:
                    deviation_percent = 0.0
                
                action_outlier = ActionOutlier(
                    action_id=action_id,
                    action_name=self.get_action_name(action_id),
                    duration=action_duration,
                    avg_duration=stats['avg'],
                    median_duration=stats['median'],
                    std_duration=stats['std'],
                    mad_duration=stats['mad'],
                    deviation_abs=deviation_abs,
                    deviation_percent=deviation_percent,
                    deviation_sigma=deviation_sigma,
                    is_too_long=is_too_long
                )
                action_outliers.append(action_outlier)
        
        # Сортируем по отклонению в сигмах
        action_outliers.sort(key=lambda x: x.deviation_sigma, reverse=True)
        
        return action_outliers

    def get_statistics_summary(self) -> Dict[str, Any]:
        """
        Возвращает сводную статистику для отображения.
        """
        outliers = self.find_process_outliers()
        
        too_long = sum(1 for o in outliers if o.is_too_long)
        too_short = len(outliers) - too_long
        
        total_action_outliers = sum(len(o.action_outliers) for o in outliers)
        
        return {
            'total_processes': len(self.processes_data),
            'process_outliers_count': len(outliers),
            'too_long_count': too_long,
            'too_short_count': too_short,
            'total_action_outliers': total_action_outliers,
            'sigma_threshold': self.sigma_threshold,
            'process_avg': self.process_avg,
            'process_median': self.process_median,
            'process_std': self.process_std,
            'process_mad': self.process_mad
        }

    def set_sigma_threshold(self, threshold: float):
        """Устанавливает новый порог сигм"""
        self.sigma_threshold = threshold


def format_timedelta(td: timedelta) -> str:
    """Форматирует timedelta в читаемый вид"""
    total_seconds = int(abs(td.total_seconds()))
    sign = "-" if td.total_seconds() < 0 else ""
    
    if total_seconds < 60:
        return f"{sign}{total_seconds}сек"
    elif total_seconds < 3600:
        minutes = total_seconds // 60
        seconds = total_seconds % 60
        return f"{sign}{minutes}мин {seconds}сек"
    elif total_seconds < 86400:
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        return f"{sign}{hours}ч {minutes}мин"
    else:
        days = total_seconds // 86400
        hours = (total_seconds % 86400) // 3600
        return f"{sign}{days}д {hours}ч"


def format_deviation(deviation_abs: timedelta, deviation_percent: float) -> str:
    """Форматирует отклонение в абсолютных значениях и процентах"""
    sign = "+" if deviation_abs.total_seconds() >= 0 else ""
    abs_str = format_timedelta(deviation_abs)
    if deviation_abs.total_seconds() < 0:
        abs_str = "-" + abs_str.lstrip("-")
    else:
        abs_str = "+" + abs_str
    
    percent_sign = "+" if deviation_percent >= 0 else ""
    return f"{abs_str} ({percent_sign}{deviation_percent:.1f}%)"

