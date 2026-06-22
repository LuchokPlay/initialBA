import json
import math
import re
import statistics
from collections import Counter, defaultdict
from datetime import datetime, timedelta


MISSING_MARKERS = {"", "nan", "none", "null", "nat"}
CASE_NUMBER_PATTERN = re.compile(r"(\d+)")


def load_json_data(filepath):
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json_data(filepath, data):
    invalidate_record_indexes(data)
    metadata = data.setdefault("metadata", {})
    metadata["total_records"] = len(data.get("records", []))
    metadata["unique_actions"] = len(data.get("actions", []))
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def invalidate_record_indexes(data):
    for record in data.get("records", []):
        record.pop("_row_index", None)


def is_missing(value):
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if isinstance(value, str) and value.strip().lower() in MISSING_MARKERS:
        return True
    return False


def action_name(data, action_id):
    try:
        action_id = int(action_id)
        actions = data.get("actions", [])
        if 0 <= action_id < len(actions):
            action = actions[action_id]
            if isinstance(action, dict):
                return action.get("name", f"Action_{action_id}")
            return str(action)
    except Exception:
        pass
    return f"Action_{action_id}"


def parse_dt(value):
    if is_missing(value):
        return None
    try:
        return datetime.fromisoformat(str(value))
    except Exception:
        return None


def parse_case_id_value(value):
    if is_missing(value):
        return None
    match = CASE_NUMBER_PATTERN.search(str(value))
    if not match:
        return None
    return int(match.group(1))


def format_timedelta(td):
    if td is None:
        return ""
    total_seconds = int(abs(td.total_seconds()))
    sign = "-" if td.total_seconds() < 0 else ""
    if total_seconds < 60:
        return f"{sign}{total_seconds} сек"
    if total_seconds < 3600:
        minutes = total_seconds // 60
        seconds = total_seconds % 60
        return f"{sign}{minutes} мин {seconds} сек"
    if total_seconds < 86400:
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        return f"{sign}{hours} ч {minutes} мин"
    days = total_seconds // 86400
    hours = (total_seconds % 86400) // 3600
    return f"{sign}{days} д {hours} ч"


def group_processes(data):
    processes = defaultdict(list)
    for index, record in enumerate(data.get("records", [])):
        record["_row_index"] = index
        processes[record.get("case_id")].append(record)
    for records in processes.values():
        records.sort(key=lambda r: parse_dt(r.get("datetime")) or datetime.min)
    return processes


def common_final_action_ids(data, threshold=0.7):
    actions = data.get("actions", [])
    action_name_to_id = {}
    for action_id, action in enumerate(actions):
        if isinstance(action, dict):
            action_name_to_id[action.get("name", f"Action_{action_id}")] = action_id
        else:
            action_name_to_id[str(action)] = action_id

    sequence_info = data.get("metadata", {}).get("sequence_analysis", {}) or {}
    metadata_end_actions = sequence_info.get("end_actions", []) or []
    metadata_ids = {
        action_name_to_id[action]
        for action in metadata_end_actions
        if action in action_name_to_id
    }
    if metadata_ids:
        return metadata_ids

    terminal_counts = Counter()
    total_processes = 0
    for records in group_processes(data).values():
        if not records:
            continue
        final_action_id = records[-1].get("action_id")
        if is_missing(final_action_id):
            continue
        terminal_counts[final_action_id] += 1
        total_processes += 1

    if not terminal_counts or not total_processes:
        return set()

    dominant_ids = {
        action_id
        for action_id, count in terminal_counts.items()
        if count / total_processes >= threshold
    }
    if dominant_ids:
        return dominant_ids

    # Если явного общепринятого финала нет, не объявляем альтернативные завершения сбоями.
    return set(terminal_counts.keys())


def find_missing_records(data):
    issues = []
    required_fields = ("case_id", "case_raw", "datetime", "action_id")
    for index, record in enumerate(data.get("records", [])):
        for field in required_fields:
            if field not in record or is_missing(record.get(field)):
                issues.append({
                    "record_index": index,
                    "field": field,
                    "current_value": record.get(field),
                    "case_id": record.get("case_id"),
                    "datetime": record.get("datetime"),
                    "action": action_name(data, record.get("action_id")),
                })
    return issues


def _case_records_in_source_order(data):
    cases = defaultdict(list)
    for index, record in enumerate(data.get("records", [])):
        cases[record.get("case_id")].append((index, record))
    return cases


def _complete_action_sequences(data):
    sequences = []
    for case_records in _case_records_in_source_order(data).values():
        actions = [record.get("action_id") for _, record in case_records]
        if actions and all(not is_missing(action_id) for action_id in actions):
            sequences.append(tuple(actions))
    return sequences


def _missing_action_span(data, record_index):
    records = data.get("records", [])
    if not (0 <= record_index < len(records)):
        return None

    record = records[record_index]
    case_records = _case_records_in_source_order(data).get(record.get("case_id"), [])
    span_position = next(
        (position for position, (index, _) in enumerate(case_records) if index == record_index),
        None
    )
    if span_position is None:
        return None

    start = span_position
    while start > 0 and is_missing(case_records[start - 1][1].get("action_id")):
        start -= 1

    end = span_position
    while end + 1 < len(case_records) and is_missing(case_records[end + 1][1].get("action_id")):
        end += 1

    return case_records, start, end, span_position


def _exact_action_span_candidates(sequences, previous_action, next_action, missing_count):
    candidates = Counter()
    for sequence in sequences:
        if previous_action is not None and next_action is not None:
            required = missing_count + 2
            if len(sequence) < required:
                continue
            for start in range(len(sequence) - required + 1):
                window = sequence[start:start + required]
                if window[0] == previous_action and window[-1] == next_action:
                    candidates[tuple(window[1:-1])] += 1
            continue

        if previous_action is not None:
            required = missing_count + 1
            if len(sequence) < required:
                continue
            for start in range(len(sequence) - required + 1):
                window = sequence[start:start + required]
                if window[0] == previous_action:
                    candidates[tuple(window[1:])] += 1
            continue

        if next_action is not None:
            required = missing_count + 1
            if len(sequence) < required:
                continue
            for start in range(len(sequence) - required + 1):
                window = sequence[start:start + required]
                if window[-1] == next_action:
                    candidates[tuple(window[:-1])] += 1

    return candidates


def _transition_probabilities(sequences):
    counts = defaultdict(Counter)
    for sequence in sequences:
        for current_action, next_action in zip(sequence, sequence[1:]):
            counts[current_action][next_action] += 1

    probabilities = {}
    for current_action, next_counts in counts.items():
        total = sum(next_counts.values())
        probabilities[current_action] = {
            next_action: count / total
            for next_action, count in next_counts.items()
            if total
        }
    return probabilities


def _probabilistic_action_span_candidate(sequences, previous_action, next_action, missing_count):
    if previous_action is None or next_action is None or missing_count <= 0:
        return None

    probabilities = _transition_probabilities(sequences)
    if previous_action not in probabilities:
        return None

    beam = [(tuple(), previous_action, 1.0)]
    for _ in range(missing_count):
        expanded = []
        for path, current_action, score in beam:
            for candidate_action, probability in probabilities.get(current_action, {}).items():
                expanded.append((path + (candidate_action,), candidate_action, score * probability))
        if not expanded:
            return None
        expanded.sort(key=lambda item: item[2], reverse=True)
        beam = expanded[:8]

    scored_paths = []
    for path, last_action, score in beam:
        final_probability = probabilities.get(last_action, {}).get(next_action)
        if final_probability:
            scored_paths.append((path, score * final_probability))

    if not scored_paths:
        return None

    scored_paths.sort(key=lambda item: item[1], reverse=True)
    best_path, best_score = scored_paths[0]
    if best_score <= 0:
        return None

    if len(scored_paths) > 1:
        second_score = scored_paths[1][1]
        if second_score > 0 and best_score / second_score < 1.15:
            return None

    return best_path


def infer_action_span(data, record_index):
    span_info = _missing_action_span(data, record_index)
    if span_info is None:
        return None

    case_records, start, end, _ = span_info
    missing_count = end - start + 1
    previous_action = case_records[start - 1][1].get("action_id") if start > 0 else None
    next_action = case_records[end + 1][1].get("action_id") if end + 1 < len(case_records) else None
    previous_action = None if is_missing(previous_action) else previous_action
    next_action = None if is_missing(next_action) else next_action

    sequences = _complete_action_sequences(data)
    if not sequences:
        return None

    exact_candidates = _exact_action_span_candidates(
        sequences,
        previous_action,
        next_action,
        missing_count
    )
    if exact_candidates:
        best_path, best_count = exact_candidates.most_common(1)[0]
        tied_paths = [path for path, count in exact_candidates.items() if count == best_count]
        if len(tied_paths) == 1:
            return tuple(best_path)

    probabilistic_path = _probabilistic_action_span_candidate(
        sequences,
        previous_action,
        next_action,
        missing_count
    )
    if probabilistic_path is not None:
        return probabilistic_path

    return None


def suggest_action_replacements(data, record_indices):
    replacements = {}
    selected_indices = set(record_indices)
    processed_span_indices = set()

    for record_index in sorted(selected_indices):
        if record_index in processed_span_indices:
            continue
        records = data.get("records", [])
        if not (0 <= record_index < len(records)):
            continue
        if not is_missing(records[record_index].get("action_id")):
            continue

        span_info = _missing_action_span(data, record_index)
        if span_info is None:
            continue
        case_records, start, end, _ = span_info
        span_indices = [case_records[position][0] for position in range(start, end + 1)]
        processed_span_indices.update(span_indices)

        inferred_path = infer_action_span(data, record_index)
        if inferred_path is None or len(inferred_path) != len(span_indices):
            continue

        for index, action_id in zip(span_indices, inferred_path):
            if index in selected_indices:
                replacements[index] = action_id

    return replacements


def most_likely_value(data, issue):
    field = issue["field"]
    records = data.get("records", [])
    index = issue["record_index"]
    record = records[index]

    if field == "case_raw":
        case_id = record.get("case_id")
        if not is_missing(case_id):
            return f"Case_{case_id}"

    if field == "case_id":
        case_raw = record.get("case_raw")
        parsed_case_id = parse_case_id_value(case_raw)
        if parsed_case_id is not None:
            return parsed_case_id

    if field == "datetime":
        case_id = record.get("case_id")
        previous_dt = None
        next_dt = None

        for candidate in records[:index][::-1]:
            if candidate.get("case_id") != case_id:
                continue
            previous_dt = parse_dt(candidate.get("datetime"))
            if previous_dt is not None:
                break

        for candidate in records[index + 1:]:
            if candidate.get("case_id") != case_id:
                continue
            next_dt = parse_dt(candidate.get("datetime"))
            if next_dt is not None:
                break

        if previous_dt is not None and next_dt is not None and previous_dt <= next_dt:
            midpoint = previous_dt + (next_dt - previous_dt) / 2
            return midpoint.replace(microsecond=0).isoformat()

        return ""

    if field == "action_id":
        inferred_path = infer_action_span(data, index)
        span_info = _missing_action_span(data, index)
        if inferred_path is not None and span_info is not None:
            _, start, _, span_position = span_info
            offset = span_position - start
            if 0 <= offset < len(inferred_path):
                return inferred_path[offset]
        return ""

    values = [
        r.get(field)
        for r in records
        if r is not record and not is_missing(r.get(field))
    ]
    if not values:
        return ""
    return Counter(values).most_common(1)[0][0]


def set_record_value(data, record_index, field, value):
    records = data.get("records", [])
    if not (0 <= record_index < len(records)):
        return
    if field == "case_id":
        parsed_case_id = parse_case_id_value(value)
        if parsed_case_id is not None:
            value = parsed_case_id
    elif field == "action_id":
        try:
            value = int(value)
        except Exception:
            pass
    records[record_index][field] = value


def delete_records(data, indices):
    to_delete = set(indices)
    data["records"] = [
        record for index, record in enumerate(data.get("records", []))
        if index not in to_delete
    ]
    invalidate_record_indexes(data)


def duplicate_signature(record):
    return (
        record.get("case_id"),
        record.get("case_raw"),
        record.get("datetime"),
        record.get("action_id"),
    )


def find_duplicate_records(data):
    duplicate_groups = []
    groups = defaultdict(list)
    for index, record in enumerate(data.get("records", [])):
        groups[duplicate_signature(record)].append(index)
    for indices in groups.values():
        if len(indices) > 1:
            duplicate_groups.append(indices)

    for records in group_processes(data).values():
        run = []
        previous_action_id = object()
        for record in records:
            action_id = record.get("action_id")
            record_index = record.get("_row_index")
            if is_missing(action_id) or record_index is None:
                if len(run) > 1:
                    duplicate_groups.append(run)
                run = []
                previous_action_id = object()
                continue

            if action_id == previous_action_id:
                run.append(record_index)
            else:
                if len(run) > 1:
                    duplicate_groups.append(run)
                run = [record_index]
                previous_action_id = action_id

        if len(run) > 1:
            duplicate_groups.append(run)

    merged_groups = []
    for indices in duplicate_groups:
        candidate = set(indices)
        merged = True
        while merged:
            merged = False
            remaining = []
            for group in merged_groups:
                if candidate & group:
                    candidate |= group
                    merged = True
                else:
                    remaining.append(group)
            merged_groups = remaining
        merged_groups.append(candidate)

    records = data.get("records", [])
    result = []
    for group in merged_groups:
        ordered_indices = sorted(
            group,
            key=lambda index: (
                parse_dt(records[index].get("datetime")) or datetime.min,
                index,
            )
        )
        if len(ordered_indices) > 1:
            result.append({
                "signature": duplicate_signature(records[ordered_indices[0]]),
                "indices": ordered_indices,
            })
    result.sort(key=lambda item: item["indices"][0])
    return result


def duplicate_indices_except_first(data):
    indices = []
    for group in find_duplicate_records(data):
        indices.extend(group["indices"][1:])
    return indices


def time_dynamics(data, group_by="month"):
    processes = group_processes(data)
    buckets = defaultdict(lambda: {
        "started": 0,
        "completed": 0,
        "durations": [],
        "failed": 0,
        "cases": [],
    })
    final_action_ids = common_final_action_ids(data)

    for records in processes.values():
        if not records:
            continue
        start = parse_dt(records[0].get("datetime"))
        end = parse_dt(records[-1].get("datetime"))
        if not start:
            continue
        if group_by == "day":
            key = start.strftime("%Y-%m-%d")
        elif group_by == "week":
            year, week, _ = start.isocalendar()
            key = f"{year}-W{week:02d}"
        else:
            key = start.strftime("%Y-%m")

        bucket = buckets[key]
        bucket["started"] += 1
        is_failed = bool(final_action_ids and records[-1].get("action_id") not in final_action_ids)
        bucket["cases"].append({
            "case_id": records[0].get("case_id"),
            "failed": is_failed,
        })
        if is_failed:
            bucket["failed"] += 1
        else:
            bucket["completed"] += 1
        if end:
            bucket["durations"].append(end - start)

    rows = []
    for period, info in sorted(buckets.items()):
        durations = info["durations"]
        avg_duration = sum(durations, timedelta()) / len(durations) if durations else timedelta(0)
        rows.append({
            "period": period,
            "started": info["started"],
            "completed": info["completed"],
            "failed": info["failed"],
            "failed_rate": round(info["failed"] / info["started"] * 100, 1) if info["started"] else 0,
            "avg_duration": format_timedelta(avg_duration),
            "cases": sorted(info["cases"], key=lambda item: str(item["case_id"])),
        })
    return rows


def find_case_timeline(data, query):
    query = str(query).strip()
    if not query:
        return []
    if query.lower().startswith("case_"):
        query = query.split("_", 1)[1]
    processes = group_processes(data)
    try:
        key = int(query)
    except Exception:
        key = query
    records = processes.get(key) or processes.get(str(key)) or []

    rows = []
    previous_dt = None
    for order, record in enumerate(records, 1):
        current_dt = parse_dt(record.get("datetime"))
        rows.append({
            "order": order,
            "datetime": record.get("datetime", ""),
            "action": action_name(data, record.get("action_id")),
            "from_previous": format_timedelta(current_dt - previous_dt) if current_dt and previous_dt else "",
            "record_index": record.get("_row_index", ""),
        })
        if current_dt:
            previous_dt = current_dt
    return rows


def process_map(data, limit=40):
    processes = group_processes(data)
    transition_counts = Counter()
    transition_cases = defaultdict(set)
    action_counts = Counter()
    for case_id, records in processes.items():
        names = [action_name(data, r.get("action_id")) for r in records]
        action_counts.update(names)
        for i in range(len(names) - 1):
            pair = (names[i], names[i + 1])
            transition_counts[pair] += 1
            transition_cases[pair].add(case_id)

    transitions = [
        {
            "from": pair[0],
            "to": pair[1],
            "count": count,
            "case_ids": sorted(transition_cases[pair], key=lambda value: str(value)),
        }
        for pair, count in transition_counts.most_common(limit)
    ]
    actions = [name for name, _ in action_counts.most_common()]
    return actions, transitions
