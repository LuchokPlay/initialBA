import html
import math
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

from PyQt5.QtCore import QUrl
from PyQt5.QtGui import QTextDocument
from PyQt5.QtPrintSupport import QPrinter

from outlier_analyzer import OutlierAnalyzer, format_deviation
from process_insights import action_name, format_timedelta, group_processes, load_json_data, parse_dt, process_map, time_dynamics
from statistics_analyzer import StatisticsAnalyzer


MAX_TABLE_ROWS = 200


def _escape(value):
    return html.escape("" if value is None else str(value))


def _format_number(value, digits=1):
    try:
        return f"{float(value):.{digits}f}"
    except Exception:
        return str(value)


class ReportExporter:
    def __init__(self, json_filepath, options, graphics_images=None):
        self.json_filepath = json_filepath
        self.options = options or {}
        self.data = load_json_data(json_filepath)
        self.section_index = 0
        self.graphics_images = graphics_images or {}

    def export_pdf(self, output_path):
        report_html = self.build_html()
        document = QTextDocument()
        document.setDocumentMargin(0)
        for name, image in self.graphics_images.items():
            document.addResource(QTextDocument.ImageResource, QUrl(name), image)
        document.setHtml(report_html)
        document.setDocumentMargin(0)

        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(output_path)
        printer.setPageSize(QPrinter.A4)
        printer.setFullPage(True)
        printer.setPageMargins(0, 0, 0, 0, QPrinter.Millimeter)
        document.print_(printer)

    def build_html(self):
        title = _escape(self.options.get("title") or "Отчет по бизнес-процессу")
        source = _escape(self.json_filepath)
        now = datetime.now().strftime("%d.%m.%Y %H:%M")
        sections = []

        adders = {
            "general_statistics": self._section_general_statistics,
            "time_statistics": self._section_time_statistics,
            "action_statistics": self._section_action_statistics,
            "transitions": self._section_transitions,
            "failed_processes": self._section_failed_processes,
            "outliers": self._section_outliers,
            "process_map": self._section_process_map,
            "transition_matrix": self._section_transition_matrix,
            "dynamics": self._section_dynamics,
            "duration_distribution": self._section_duration_distribution,
            "bottleneck_pareto": self._section_bottleneck_pareto,
        }

        for key in self.options.get("order", []):
            if not self.options.get(key, {}).get("selected"):
                continue
            try:
                sections.append(adders[key]())
            except Exception as exc:
                sections.append(self._section_error(self.options[key].get("label", key), exc))

        if not sections:
            sections.append("<h2>Разделы не выбраны</h2><p>Для отчета не был выбран ни один раздел.</p>")

        return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
body {{
    font-family: "Arial", "DejaVu Sans", sans-serif;
    color: #1f2933;
    font-size: 9.5pt;
    margin: 0;
    padding: 0;
}}
h1 {{ font-size: 20pt; margin: 0 0 8px 0; }}
h2 {{
    font-size: 15pt;
    margin: 18px 0 8px 0;
    padding-bottom: 4px;
    border-bottom: 1px solid #aab7c4;
}}
h3 {{ font-size: 11pt; margin: 12px 0 6px 0; }}
p {{ margin: 4px 0 8px 0; }}
.meta {{ color: #65717d; margin-bottom: 16px; }}
.section {{ page-break-inside: auto; }}
.page-break {{ page-break-before: always; }}
table {{ width: 100%; border-collapse: collapse; margin: 6px 0 12px 0; }}
th, td {{ border: 1px solid #d6dde5; padding: 4px 5px; vertical-align: top; }}
th {{ background: #edf2f7; font-weight: bold; }}
tr:nth-child(even) td {{ background: #f8fafc; }}
.note {{ color: #65717d; font-style: italic; }}
.small {{ font-size: 8pt; color: #596673; }}
.ok {{ color: #166534; font-weight: bold; }}
.bad {{ color: #9f1239; font-weight: bold; }}
.bar-track {{ background: #eef2f7; width: 100%; height: 10px; }}
.bar {{ background: #3f78a8; height: 10px; }}
.bar-red {{ background: #c85b5b; height: 10px; }}
.graphic {{ margin: 0 0 8px 0; padding: 0; width: 100%; }}
.map-sheet {{ height: 820px; overflow: hidden; }}
.graphic img {{ display: block; margin: 0; }}
.map-sheet img {{ margin: 0 auto; }}
</style>
</head>
<body>
<h1>{title}</h1>
<p class="meta">Источник: {source}<br>Сформировано: {_escape(now)}</p>
{''.join(sections)}
</body>
</html>
"""

    def _next_section(self, title):
        css = "section"
        if self.section_index > 0:
            css += " page-break"
        self.section_index += 1
        return f'<div class="{css}"><h2>{_escape(title)}</h2>'

    def _section_error(self, title, exc):
        return (
            self._next_section(title)
            + f'<p class="bad">Раздел не удалось сформировать: {_escape(exc)}</p></div>'
        )

    def _table(self, headers, rows, limit=MAX_TABLE_ROWS):
        visible_rows = rows[:limit]
        html_rows = [
            "<tr>" + "".join(f"<th>{_escape(header)}</th>" for header in headers) + "</tr>"
        ]
        for row in visible_rows:
            html_rows.append("<tr>" + "".join(f"<td>{_escape(cell)}</td>" for cell in row) + "</tr>")
        note = ""
        if len(rows) > limit:
            note = f'<p class="note">Показаны первые {limit} строк из {len(rows)}.</p>'
        return "<table>" + "".join(html_rows) + "</table>" + note

    def _graphic_html(self, key, width=700, height=None, css_class="graphic"):
        if key not in self.graphics_images:
            return ""
        if width is not None and height is not None:
            size_attr = f'width="{width}" height="{height}"'
        elif height is not None:
            size_attr = f'height="{height}"'
        else:
            size_attr = f'width="{width}"'
        return f'<p class="{css_class}"><img src="{key}" {size_attr}></p>'

    def _metric_table(self, pairs):
        return self._table(["Показатель", "Значение"], pairs, limit=len(pairs) or 1)

    def _section_general_statistics(self):
        analyzer = StatisticsAnalyzer(self.json_filepath)
        stats = analyzer.get_process_statistics()
        time_dist = analyzer.get_time_distribution()
        rows = [
            ["Всего процессов", stats.get("total_processes", "")],
            ["Всего действий", stats.get("total_actions", "")],
            ["Средняя длительность", stats.get("avg_process_duration", "")],
            ["Минимальная длительность", stats.get("min_process_duration", "")],
            ["Максимальная длительность", stats.get("max_process_duration", "")],
            ["Стандартное отклонение длительности", stats.get("std_process_duration", "")],
            ["Медианное отклонение длительности", stats.get("mad_process_duration", "")],
            ["Среднее количество действий", stats.get("avg_actions_per_process", "")],
            ["Стандартное отклонение количества действий", stats.get("std_actions_per_process", "")],
            ["Медианное отклонение количества действий", stats.get("mad_actions_per_process", "")],
        ]
        content = self._next_section("Общая статистика")
        content += self._metric_table(rows)
        if time_dist:
            content += "<h3>Распределение по времени выполнения</h3>"
            content += self._table(
                ["Интервал", "Процессов", "Доля"],
                [[interval, item["count"], f'{item["percentage"]}%'] for interval, item in time_dist.items()],
                limit=len(time_dist),
            )
        return content + "</div>"

    def _section_time_statistics(self):
        cfg = self.options["time_statistics"]
        rows = time_dynamics(self.data, cfg.get("group_by", "month"))
        content = self._next_section("Временная статистика")
        content += f'<p class="note">Группировка: {_escape(self._group_label(cfg.get("group_by")))}.</p>'
        content += self._table(
            ["Период", "Стартовало", "Без сбоев", "Сбоев", "% сбоев", "Средняя длительность"],
            [
                [
                    row["period"],
                    row["started"],
                    row["completed"],
                    row["failed"],
                    row["failed_rate"],
                    row["avg_duration"],
                ]
                for row in rows
            ],
            limit=len(rows) or 1,
        )
        return content + "</div>"

    def _section_action_statistics(self):
        analyzer = StatisticsAnalyzer(self.json_filepath)
        rows = analyzer.get_action_statistics()
        content = self._next_section("Статистика действий")
        content += self._table(
            ["Действие", "Выполнений", "Процессов", "Ср. длительность", "Ст. откл.", "Мед. откл.", "Мин.", "Макс."],
            [
                [
                    row["action"],
                    row["occurrences"],
                    row["processes"],
                    row["avg_duration"],
                    row["duration_std"],
                    row["duration_mad"],
                    row["min_duration"],
                    row["max_duration"],
                ]
                for row in rows
            ],
        )
        return content + "</div>"

    def _section_transitions(self):
        analyzer = StatisticsAnalyzer(self.json_filepath)
        flow = analyzer.get_process_flow_analysis()
        content = self._next_section("Переходы")
        if flow.get("action_frequency"):
            content += "<h3>Топ действий по частоте</h3>"
            content += self._table(
                ["Действие", "Выполнений", "Доля"],
                [[row["action"], row["count"], f'{row["percentage"]}%'] for row in flow["action_frequency"]],
                limit=len(flow["action_frequency"]),
            )
        if flow.get("top_transitions"):
            content += "<h3>Наиболее частые переходы</h3>"
            content += self._table(
                ["Откуда", "Куда", "Переходов", "Доля процессов"],
                [
                    [row["from"], row["to"], row["count"], f'{row["percentage"]}%']
                    for row in flow["top_transitions"]
                ],
                limit=len(flow["top_transitions"]),
            )
        return content + "</div>"

    def _section_failed_processes(self):
        analyzer = StatisticsAnalyzer(self.json_filepath)
        rows = analyzer.find_failed_processes()
        content = self._next_section("Поиск сбоев")
        content += f"<p>Найдено незавершенных процессов: <b>{len(rows)}</b>.</p>"
        if rows:
            content += self._table(
                ["Case ID", "Дата начала", "Дата последнего действия", "Последнее действие", "Ожидаемое действие", "Действий"],
                [
                    [
                        row["case_id"],
                        row["start_date"],
                        row["last_date"],
                        row["last_action"],
                        row["correct_end_action"],
                        row["actions_count"],
                    ]
                    for row in rows
                ],
            )
        else:
            content += '<p class="ok">Все процессы завершились общепринятыми конечными действиями.</p>'
        return content + "</div>"

    def _section_outliers(self):
        cfg = self.options["outliers"]
        sigma = float(cfg.get("sigma", 3.0))
        analyzer = OutlierAnalyzer(self.json_filepath, sigma)
        rows = analyzer.find_process_outliers()
        too_long = sum(1 for row in rows if row.is_too_long)
        content = self._next_section("Поиск выбросов")
        content += self._metric_table([
            ["Порог отклонения", f"{sigma}σ"],
            ["Всего процессов", len(analyzer.processes_data)],
            ["Процессов-выбросов", len(rows)],
            ["Слишком долгих", too_long],
            ["Слишком быстрых", len(rows) - too_long],
            ["Средняя длительность процесса", format_timedelta(analyzer.process_avg)],
            ["Медианная длительность процесса", format_timedelta(analyzer.process_median)],
            ["Стандартное отклонение", format_timedelta(analyzer.process_std)],
        ])
        if rows:
            content += self._table(
                ["Case ID", "Длительность", "Отклонение", "Тип"],
                [
                    [
                        f"Case_{row.case_id}",
                        format_timedelta(row.duration),
                        format_deviation(row.deviation_abs, row.deviation_percent),
                        "Долгий" if row.is_too_long else "Быстрый",
                    ]
                    for row in rows
                ],
            )
        return content + "</div>"

    def _section_process_map(self):
        cfg = self.options["process_map"]
        threshold = int(float(cfg.get("min_transitions", 0)))
        actions, transitions = process_map(self.data, limit=60)
        transitions = [row for row in transitions if row["count"] >= threshold]
        content = self._next_section("Карта процессов")
        content += f'<p class="note">Показаны связи от {threshold} переходов.</p>'
        content += self._graphic_html("process_map", width=300, height=800, css_class="graphic map-sheet")
        content += '<div class="page-break"></div>'
        content += self._table(
            ["Откуда", "Куда", "Переходов"],
            [
                [
                    row["from"],
                    row["to"],
                    row["count"],
                ]
                for row in transitions
            ],
            limit=60,
        )
        return content + "</div>"

    def _section_transition_matrix(self):
        actions, matrix = self._transition_matrix()
        content = self._next_section("Матрица переходов")
        image_html = self._graphic_html("transition_matrix")
        if image_html:
            return content + image_html + "</div>"
        if not actions:
            return content + "<p>Нет переходов для матрицы.</p></div>"
        rows = []
        for source in actions:
            rows.append([source] + [matrix.get((source, target), 0) or "" for target in actions])
        content += self._table(["Откуда / куда"] + actions, rows, limit=len(rows))
        return content + "</div>"

    def _section_dynamics(self):
        cfg = self.options["dynamics"]
        group_by = cfg.get("group_by", "month")
        filtered_data = self._filter_by_range(cfg.get("date_from"), cfg.get("date_to"))
        rows = time_dynamics(filtered_data, group_by)
        content = self._next_section("Динамика")
        content += (
            f'<p class="note">Группировка: {_escape(self._group_label(group_by))}; '
            f'диапазон: {_escape(cfg.get("date_from", ""))} - {_escape(cfg.get("date_to", ""))}.</p>'
        )
        image_html = self._graphic_html("dynamics")
        if image_html:
            return content + image_html + "</div>"
        content += self._table(
            ["Период", "Стартовало", "Без сбоев", "Сбоев", "% сбоев", "Средняя длительность"],
            [
                [row["period"], row["started"], row["completed"], row["failed"], row["failed_rate"], row["avg_duration"]]
                for row in rows
            ],
            limit=len(rows) or 1,
        )
        return content + "</div>"

    def _section_duration_distribution(self):
        cfg = self.options["duration_distribution"]
        sigma = float(cfg.get("sigma", 3.0))
        rows = self._process_rows(self.data)
        durations = [row["duration_seconds"] for row in rows]
        content = self._next_section("Распределение длительности")
        content += f'<p class="note">Порог отклонения: {sigma}σ.</p>'
        content += self._graphic_html("duration_distribution")
        if not durations:
            return content + "<p>Нет длительностей процессов для распределения.</p></div>"

        avg = sum(durations) / len(durations)
        std = self._std(durations)
        if std > 0:
            visible = [value for value in durations if abs(value - avg) / std < sigma]
        else:
            visible = durations
        if not visible:
            visible = durations
        bins = self._duration_bins(visible)
        max_count = max((count for _, _, count in bins), default=1)
        content += self._metric_table([
            ["Процессов", len(durations)],
            ["Средняя длительность", format_timedelta(timedelta(seconds=avg))],
            ["Стандартное отклонение", format_timedelta(timedelta(seconds=std))],
            ["Скрыто выбросов гистограммы", len(durations) - len(visible)],
        ])
        content += self._table(
            ["Интервал", "Процессов", "Доля"],
            [
                [
                    f"{format_timedelta(timedelta(seconds=start))} - {format_timedelta(timedelta(seconds=end))}",
                    count,
                    self._bar_cell(count, max_count),
                ]
                for start, end, count in bins
            ],
            limit=len(bins),
        )
        if std > 0:
            outlier_rows = [
                row for row in rows
                if abs(row["duration_seconds"] - avg) / std >= sigma
            ]
            if outlier_rows:
                outlier_rows.sort(key=lambda row: abs(row["duration_seconds"] - avg), reverse=True)
                content += "<h3>Процессы за порогом</h3>"
                content += self._table(
                    ["Case ID", "Длительность", "Отклонение, σ"],
                    [
                        [
                            f"Case_{row['case_id']}",
                            format_timedelta(timedelta(seconds=row["duration_seconds"])),
                            _format_number(abs(row["duration_seconds"] - avg) / std, 2),
                        ]
                        for row in outlier_rows
                    ],
                )
        return content + "</div>"

    def _section_bottleneck_pareto(self):
        rows = self._transition_duration_rows(self.data)
        durations = defaultdict(float)
        counts = Counter()
        for row in rows:
            key = f"{row['from']} -> {row['to']}"
            durations[key] += row["duration_seconds"]
            counts[key] += 1
        top_items = sorted(durations.items(), key=lambda item: item[1], reverse=True)[:12]
        total = sum(durations.values()) or 1
        content = self._next_section("Парето узких мест")
        image_html = self._graphic_html("bottleneck_pareto")
        if image_html:
            return content + image_html + "</div>"
        if not top_items:
            return content + "<p>Нет переходов с длительностью для Парето.</p></div>"
        cumulative = 0
        table_rows = []
        for name, seconds in top_items:
            cumulative += seconds
            table_rows.append([
                name,
                counts[name],
                format_timedelta(timedelta(seconds=seconds)),
                f"{seconds / total * 100:.1f}%",
                f"{cumulative / total * 100:.1f}%",
            ])
        content += self._table(
            ["Переход", "Срабатываний", "Суммарное время", "Доля", "Накопленная доля"],
            table_rows,
            limit=len(table_rows),
        )
        return content + "</div>"

    def _transition_matrix(self):
        actions = [action_name(self.data, index) for index, _ in enumerate(self.data.get("actions", []))]
        transitions = Counter()
        for records in group_processes(self.data).values():
            names = [action_name(self.data, record.get("action_id")) for record in records]
            for source, target in zip(names, names[1:]):
                transitions[(source, target)] += 1
        return actions, transitions

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
                rows.append({
                    "case_id": case_id,
                    "from": action_name(data, current.get("action_id")),
                    "to": action_name(data, next_record.get("action_id")),
                    "duration_seconds": max(0, (end - start).total_seconds()),
                })
        return rows

    def _filter_by_range(self, date_from, date_to):
        start = self._parse_date(date_from)
        end = self._parse_date(date_to)
        if start is None or end is None:
            return self.data
        if start > end:
            start, end = end, start
        records = []
        for process_records in group_processes(self.data).values():
            if not process_records:
                continue
            start_dt = parse_dt(process_records[0].get("datetime"))
            if start_dt is None:
                continue
            process_date = start_dt.date()
            if start <= process_date <= end:
                records.extend(process_records)
        filtered = dict(self.data)
        filtered["records"] = [
            {key: value for key, value in record.items() if key != "_row_index"}
            for record in records
        ]
        return filtered

    def _parse_date(self, value):
        if isinstance(value, date):
            return value
        try:
            return datetime.strptime(str(value), "%Y-%m-%d").date()
        except Exception:
            return None

    def _duration_bins(self, values):
        values = sorted(values)
        if not values:
            return []
        bin_count = min(12, max(5, int(math.sqrt(len(values)))))
        min_d = min(values)
        max_d = max(values)
        span = max(1, max_d - min_d)
        counts = [0] * bin_count
        for value in values:
            index = min(bin_count - 1, int((value - min_d) / span * bin_count))
            counts[index] += 1
        bins = []
        for index, count in enumerate(counts):
            start = min_d + span * index / bin_count
            end = min_d + span * (index + 1) / bin_count
            bins.append((start, end, count))
        return bins

    def _std(self, values):
        if len(values) < 2:
            return 0
        avg = sum(values) / len(values)
        return (sum((value - avg) ** 2 for value in values) / (len(values) - 1)) ** 0.5

    def _bar_cell(self, value, max_value):
        ratio = 0 if not max_value else max(0, min(100, value / max_value * 100))
        return f"{value} ({ratio:.0f}%)"

    def _group_label(self, group_by):
        return {
            "day": "по дням",
            "week": "по неделям",
            "month": "по месяцам",
        }.get(group_by, group_by or "")
