"""
Модуль анализа последовательностей действий в бизнес-процессах.
Автоматически определяет логический порядок действий на основе статистики переходов.
"""

from collections import defaultdict, Counter
from typing import List, Dict, Tuple, Set
from dataclasses import dataclass



@dataclass
class Transition:
    """Переход между действиями"""
    from_action: str
    to_action: str
    count: int
    probability: float = 0.0


@dataclass
class ProcessSequence:
    """Последовательность действий одного процесса"""
    case_id: str
    actions: List[str]


class SequenceAnalyzer:
    """
    Анализатор последовательностей действий в бизнес-процессах.

    Автоматически определяет логическую последовательность действий
    на основе анализа переходов и статистических закономерностей.
    """

    def __init__(self, records: List[Dict]):
        """
        Инициализация анализатора.

        Args:
            records: список записей с полями case_id, action, datetime
        """
        self.records = records
        self.processes = self._build_processes()
        self.transitions = self._build_transitions()

    def _build_processes(self) -> List[ProcessSequence]:
        """Строит последовательности действий для каждого case_id"""
        processes_dict = defaultdict(list)

        # Группируем записи по case_id
        for record in self.records:
            processes_dict[record['case_id']].append(record)

        # Сортируем по времени и создаем последовательности
        processes = []
        for case_id, case_records in processes_dict.items():
            # Сортируем записи по времени
            sorted_records = sorted(case_records, key=lambda x: x['datetime'])
            actions = [record['action'] for record in sorted_records]

            processes.append(ProcessSequence(
                case_id=case_id,
                actions=actions
            ))

        return processes

    def _build_transitions(self) -> Dict[str, List[Transition]]:
        """Строит граф переходов между действиями"""
        transition_counts = defaultdict(lambda: defaultdict(int))

        # Считаем переходы
        for process in self.processes:
            for i in range(len(process.actions) - 1):
                from_action = process.actions[i]
                to_action = process.actions[i + 1]
                transition_counts[from_action][to_action] += 1

        # Преобразуем в объекты Transition
        transitions = {}
        for from_action, to_actions in transition_counts.items():
            total_transitions = sum(to_actions.values())
            transition_list = []

            for to_action, count in to_actions.items():
                transition_list.append(Transition(
                    from_action=from_action,
                    to_action=to_action,
                    count=count,
                    probability=count / total_transitions
                ))

            # Сортируем по вероятности (наиболее вероятные первыми)
            transition_list.sort(key=lambda x: x.probability, reverse=True)
            transitions[from_action] = transition_list

        return transitions

    def find_start_actions(self, threshold: float = 0.7) -> List[str]:
        """
        Находит стартовые действия (действия, с которых начинаются процессы).

        Args:
            threshold: минимальная доля процессов, начинающихся с действия

        Returns:
            список названий стартовых действий
        """
        if not self.processes:
            return []

        # Считаем частоту каждого действия на первой позиции
        start_counts = Counter()
        total_processes = len(self.processes)

        for process in self.processes:
            if process.actions:
                start_counts[process.actions[0]] += 1

        # Выбираем действия, которые встречаются чаще threshold
        start_actions = []
        for action, count in start_counts.items():
            if count / total_processes >= threshold:
                start_actions.append(action)

        return start_actions

    def find_end_actions(self, threshold: float = 0.7) -> List[str]:
        """
        Находит финальные действия (действия, которыми заканчиваются процессы).

        Args:
            threshold: минимальная доля процессов, заканчивающихся действием

        Returns:
            список названий финальных действий
        """
        if not self.processes:
            return []

        # Считаем частоту каждого действия на последней позиции
        end_counts = Counter()
        total_processes = len(self.processes)

        for process in self.processes:
            if process.actions:
                end_counts[process.actions[-1]] += 1

        # Выбираем действия, которые встречаются чаще threshold
        end_actions = []
        for action, count in end_counts.items():
            if count / total_processes >= threshold:
                end_actions.append(action)

        return end_actions

    def find_most_likely_sequence(self) -> List[str]:
        """
        Находит наиболее вероятную последовательность действий.
        Использует комбинацию анализа переходов и позиционного анализа.

        Returns:
            список действий в логическом порядке
        """
        if not self.processes:
            return []

        # Сначала пробуем позиционный анализ - он гарантированно включает все действия
        # и дает разумный порядок по средним позициям в процессах
        position_based_sequence = self._sequence_by_positions()

        # Пробуем улучшить порядок с помощью анализа переходов
        start_actions = self.find_start_actions(threshold=0.8)
        end_actions = self.find_end_actions(threshold=0.8)

        if start_actions and end_actions:
            # Выбираем наиболее частые старт и финиш
            start_action = self._get_most_frequent_action(start_actions, 'start')
            end_action = self._get_most_frequent_action(end_actions, 'end')

            # Пробуем найти путь
            main_path = self._find_path(start_action, end_action)

            if main_path and len(main_path) >= 2:
                # Пробуем вставить оставшиеся действия
                all_actions = set()
                for process in self.processes:
                    all_actions.update(process.actions)

                used_actions = set(main_path)
                remaining_actions = all_actions - used_actions

                if remaining_actions:
                    # Сначала обрабатываем финальные действия - они должны быть в конце
                    end_actions = self.find_end_actions(threshold=0.8)  # Высокий порог для финальных
                    final_actions = [action for action in remaining_actions if action in end_actions]
                    non_final_actions = remaining_actions - set(final_actions)

                    # Вставляем не-финальные действия
                    extended_path = main_path
                    if non_final_actions:
                        extended_path = self._insert_remaining_actions_smart(extended_path, non_final_actions)

                    # Добавляем финальные действия в конец
                    if final_actions:
                        extended_path.extend(final_actions)

                    # Если вставка удалась и все действия включены, используем расширенный путь
                    if len(extended_path) == len(all_actions):
                        return extended_path

        # В остальных случаях возвращаем позиционный анализ
        # (гарантированно включает все действия в разумном порядке)
        return position_based_sequence

    def _find_path(self, start_action: str, end_action: str) -> List[str]:
        """
        Находит наиболее вероятный путь от start_action до end_action.
        """
        if start_action == end_action:
            return [start_action]

        sequence = [start_action]
        current_action = start_action
        visited = set([start_action])
        max_steps = len(self.transitions) * 2  # предотвращаем бесконечные циклы

        while current_action != end_action and max_steps > 0:
            if current_action not in self.transitions:
                break

            # Выбираем наиболее вероятный переход, которого еще не посещали
            next_action = None
            for transition in self.transitions[current_action]:
                if transition.to_action not in visited:
                    next_action = transition.to_action
                    break

            if next_action is None:
                break

            if next_action in visited:
                break

            sequence.append(next_action)
            visited.add(next_action)
            current_action = next_action
            max_steps -= 1

        return sequence

    def _get_most_frequent_action(self, actions_list: List[str], action_type: str) -> str:
        """Возвращает наиболее частое действие из списка"""
        if len(actions_list) == 1:
            return actions_list[0]

        # Считаем частоту
        counts = {}
        for process in self.processes:
            if action_type == 'start' and process.actions:
                action = process.actions[0]
            elif action_type == 'end' and process.actions:
                action = process.actions[-1]
            else:
                continue

            if action in actions_list:
                counts[action] = counts.get(action, 0) + 1

        if counts:
            return max(counts.items(), key=lambda x: x[1])[0]

        return actions_list[0]  # Возвращаем первое, если ничего не нашли

    def _insert_remaining_actions_smart(self, main_sequence: List[str], remaining_actions: Set[str]) -> List[str]:
        """
        Вставляет оставшиеся действия между одинаковыми действиями на основе паттернов A->X->A.
        """
        result_sequence = main_sequence.copy()

        for action_x in remaining_actions:
            # Ищем паттерны для этого действия
            insertion_candidates = self._find_insertion_candidates(action_x)

            # Выбираем лучший кандидат для вставки
            if insertion_candidates:
                # Сортируем кандидатов по частоте
                sorted_candidates = sorted(insertion_candidates.items(), key=lambda x: x[1], reverse=True)

                for action_before, frequency in sorted_candidates:
                    # Проверяем, есть ли action_before в текущей последовательности
                    if action_before in result_sequence:
                        # Проверяем порог (разные пороги для разных типов паттернов)
                        total_occurrences = sum(1 for process in self.processes if action_x in process.actions)
                        threshold = 0.8 if frequency >= 2 else 0.9  # Меньший порог для сильных паттернов

                        if frequency / total_occurrences >= threshold:
                            # Вставляем X после action_before
                            result_sequence = self._insert_after_action(result_sequence, action_before, action_x)
                            break  # Выходим после первой успешной вставки

        return result_sequence

    def _find_insertion_candidates(self, action_x: str) -> Dict[str, int]:
        """
        Находит позиции для вставки X, анализируя паттерны возврата и циклов.
        Ищет паттерны A->X->A и более сложные цепочки с возвратами.
        Возвращает словарь action_before -> количество паттернов.
        """
        candidates = defaultdict(int)

        for process in self.processes:
            actions = process.actions

            # 1. Ищем паттерны A->X->A (простые циклы)
            for i in range(len(actions) - 2):
                if actions[i] == actions[i + 2] and actions[i + 1] == action_x:
                    candidates[actions[i]] += 2  # Высокий вес для идеальных паттернов

            # 2. Ищем паттерны с возвратом: ... -> Y -> X -> Z -> ...
            # где Z появляется раньше в последовательности (возврат назад)
            for i in range(len(actions)):
                if actions[i] == action_x:
                    # Ищем действие перед X
                    if i > 0:
                        action_before = actions[i - 1]
                        # Ищем действие после X, которое является возвратом
                        for j in range(i + 1, len(actions)):
                            action_after = actions[j]
                            # Проверяем, появлялось ли action_after раньше в процессе
                            if action_after in actions[:i]:
                                # Нашли паттерн возврата: action_before -> X -> action_after (где action_after - возврат)
                                candidates[action_before] += 1
                                break  # Берем первый найденный возврат

        return dict(candidates)

    def _insert_between_actions(self, sequence: List[str], action_a: str, action_x: str) -> List[str]:
        """
        Вставляет action_x между двумя последовательными action_a в последовательности.
        """
        result = []
        inserted = False

        for i in range(len(sequence)):
            result.append(sequence[i])

            # Если текущий элемент - action_a и следующий тоже action_a, вставляем X между ними
            if (not inserted and
                sequence[i] == action_a and
                i + 1 < len(sequence) and
                sequence[i + 1] == action_a):

                result.append(action_x)
                inserted = True

        return result

    def _insert_after_action(self, sequence: List[str], action_before: str, action_x: str) -> List[str]:
        """
        Вставляет action_x после action_before в последовательности.
        """
        result = []
        inserted = False

        for item in sequence:
            result.append(item)

            if not inserted and item == action_before:
                result.append(action_x)
                inserted = True

        # Если action_before не найден, добавляем в конец
        if not inserted:
            result.append(action_x)

        return result

    def _append_remaining_by_positions(self, main_sequence: List[str], remaining_actions: Set[str]) -> List[str]:
        """
        Добавляет оставшиеся действия в конец последовательности, отсортированные по средним позициям.
        """
        if not remaining_actions:
            return main_sequence

        # Сортируем оставшиеся действия по средним позициям
        position_stats = defaultdict(list)
        for process in self.processes:
            for pos, action in enumerate(process.actions):
                if action in remaining_actions:
                    position_stats[action].append(pos)

        # Вычисляем средние позиции
        avg_positions = {}
        for action, positions in position_stats.items():
            if positions:  # Проверяем, что список не пустой
                avg_positions[action] = sum(positions) / len(positions)
            else:
                avg_positions[action] = float('inf')  # Если нет позиций, ставим в конец

        # Сортируем по возрастанию средней позиции
        sorted_remaining = sorted(avg_positions.items(), key=lambda x: x[1])
        remaining_sequence = [action for action, _ in sorted_remaining]

        # Добавляем в конец основной последовательности
        result = main_sequence + remaining_sequence
        return result

    def _sequence_by_positions(self) -> List[str]:
        """
        Альтернативный метод: упорядочивание по средним позициям в процессах.
        """
        if not self.processes:
            return []

        # Считаем среднюю позицию для каждого действия
        position_stats = defaultdict(list)

        for process in self.processes:
            for pos, action in enumerate(process.actions):
                position_stats[action].append(pos)

        # Вычисляем средние позиции
        avg_positions = {}
        for action, positions in position_stats.items():
            avg_positions[action] = sum(positions) / len(positions)

        # Сортируем по возрастанию средней позиции
        sorted_actions = sorted(avg_positions.items(), key=lambda x: x[1])

        return [action for action, _ in sorted_actions]


    def _find_dead_ends(self) -> List[str]:
        """
        Находит тупиковые действия (действия, которые редко ведут к другим действиям или представляют собой точки завершения).

        Тупиковыми считаются действия, которые:
        1. Часто являются последними в процессах (>=50% случаев)
        2. Не имеют исходящих переходов (терминальные узлы)
        3. Имеют очень мало исходящих переходов и низкую частоту использования
        4. Имеют низкую вероятность успешного продолжения процесса
        5. Имеют мало связей с другими действиями (изолированные действия)

        Returns:
            список названий тупиковых действий
        """
        if not self.processes:
            return []

        dead_end_actions = set()

        # Собираем все уникальные действия
        all_actions = set()
        for process in self.processes:
            all_actions.update(process.actions)

        # Критерий 1: действия, часто являющиеся последними (>=50% процессов)
        end_actions = self.find_end_actions(threshold=0.5)
        dead_end_actions.update(end_actions)

        # Критерий 2: действия без исходящих переходов (терминальные узлы)
        actions_with_outgoing = set(self.transitions.keys())
        actions_without_outgoing = all_actions - actions_with_outgoing
        dead_end_actions.update(actions_without_outgoing)

        # Критерий 3: действия с очень низкой исходящей степенью и частотой использования
        if self.transitions:
            outgoing_counts = {action: len(transitions) for action, transitions in self.transitions.items()}

            if outgoing_counts:
                # Действия с 0 или 1 исходящим переходом, которые редко используются
                for action, count in outgoing_counts.items():
                    if count <= 1:  # 0 или 1 исходящий переход
                        # Проверяем частоту использования действия
                        action_frequency = sum(1 for process in self.processes if action in process.actions)
                        total_actions_in_processes = sum(len(process.actions) for process in self.processes)

                        # Если действие используется редко (< 5% от общего числа действий), оно тупиковое
                        if action_frequency / total_actions_in_processes < 0.05:
                            dead_end_actions.add(action)

        # Критерий 4: действия с низкой вероятностью успешного продолжения
        for from_action, transitions in self.transitions.items():
            if transitions:
                # Вычисляем максимальную вероятность перехода
                max_probability = max(t.probability for t in transitions)

                # Если максимальная вероятность перехода низкая (< 30%), действие может быть тупиковым
                if max_probability < 0.3:
                    # Проверяем, сколько процессов доходят до этого действия
                    processes_reaching_action = sum(1 for process in self.processes
                                                  if from_action in process.actions)

                    # Если это редкое действие (< 20% процессов), добавляем в тупиковые
                    if processes_reaching_action / len(self.processes) < 0.2:
                        dead_end_actions.add(from_action)

        # Критерий 5: изолированные действия (редко соединяются с другими)
        action_connections = defaultdict(set)
        for process in self.processes:
            for i in range(len(process.actions) - 1):
                action_connections[process.actions[i]].add(process.actions[i + 1])
                action_connections[process.actions[i + 1]].add(process.actions[i])

        for action in all_actions:
            connections = len(action_connections[action])
            # Если действие имеет мало связей (< 2), оно может быть тупиковым
            if connections < 2:
                # Проверяем частоту использования
                usage_count = sum(1 for process in self.processes if action in process.actions)
                if usage_count / len(self.processes) < 0.3:  # Менее 30% процессов
                    dead_end_actions.add(action)

        return list(dead_end_actions)

    def get_statistics(self) -> Dict:
        """
        Возвращает статистику по последовательностям.
        """
        if not self.processes:
            return {}

        total_processes = len(self.processes)
        action_counts = Counter()

        for process in self.processes:
            for action in process.actions:
                action_counts[action] += 1

        # Длина процессов
        process_lengths = [len(p.actions) for p in self.processes]
        avg_length = sum(process_lengths) / len(process_lengths) if process_lengths else 0

        return {
            'total_processes': total_processes,
            'unique_actions': len(action_counts),
            'average_process_length': round(avg_length, 2),
            'start_actions': self.find_start_actions(),
            'end_actions': self.find_end_actions(),
            'most_common_actions': action_counts.most_common(5)
        }

    


def sort_actions_automatically(records: List[Dict]) -> List[str]:
    """
    Основная функция для автоматической сортировки действий.

    Args:
        records: список записей из JSON (с полями case_id, action, datetime)

    Returns:
        список названий действий в логическом порядке
    """
    analyzer = SequenceAnalyzer(records)
    sequence = analyzer.find_most_likely_sequence()
    return sequence
