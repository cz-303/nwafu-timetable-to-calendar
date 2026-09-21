"""Offline NWAFU list-export XLS to RFC 5545 iCalendar converter."""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path

CHINA = timezone(timedelta(hours=8))
HEADERS = ('课程号', '课程名', '课序号', '上课周次', '上课星期', '开始节次', '结束节次', '教室名称')


@dataclass(frozen=True)
class Lesson:
    code: str
    name: str
    section: str
    weeks: tuple[int, ...]
    weekday: int
    first: int
    last: int
    location: str


@dataclass(frozen=True)
class Event:
    lesson: Lesson
    week: int
    first: int
    last: int
    start: datetime
    end: datetime


def clean(value):
    return str(value).strip() if value is not None else ''


def parse_weeks(value):
    text = re.sub(r'\s+', '', clean(value)).replace('（', '(').replace('）', ')')
    text = re.sub(r'[－—–~～至]', '-', text)
    text = re.sub(r'[，、；;]', ',', text)
    # A final parity marker applies to the whole expression.
    parity = None
    match = re.search(r'(?:\(([单双])(?:周)?\)|([单双])(?:周)?)$', text)
    if match:
        parity = match[1] or match[2]
        text = text[:match.start()]
    text = text.replace('第', '').replace('周', '')
    if not re.fullmatch(r'\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*', text):
        raise ValueError(f'无法识别周次“{value}”，请使用 1-16周、1,3,5周 或 1-16周(单)。')
    weeks = set()
    for part in text.split(','):
        bounds = [int(n) for n in part.split('-')]
        lo, hi = bounds[0], bounds[-1]
        if not 1 <= lo <= hi <= 53:
            raise ValueError(f'周次范围无效：{value}')
        weeks.update(range(lo, hi + 1))
    if parity:
        weeks = {w for w in weeks if w % 2 == (1 if parity == '单' else 0)}
    if not weeks:
        raise ValueError(f'周次没有有效上课周：{value}')
    return tuple(sorted(weeks))


def parse_weekday(value):
    text = re.sub(r'\s+', '', clean(value))
    text = re.sub(r'^(星期|周|礼拜)', '', text)
    days = {'一': 0, '二': 1, '三': 2, '四': 3, '五': 4, '六': 5, '日': 6, '天': 6}
    if text in days:
        return days[text]
    if re.fullmatch(r'[1-7](?:\.0)?', text):
        return int(float(text)) - 1
    raise ValueError(f'无法识别星期：{value}')


def parse_period(value):
    text = clean(value)
    match = re.fullmatch(r'(?:第)?(\d+)(?:\.0)?(?:节)?', text)
    if not match or not 1 <= int(match[1]) <= 24:
        raise ValueError(f'无法识别节次：{value}')
    return int(match[1])


def parse_rows(rows):
    if not rows:
        raise ValueError('工作表为空。')
    header_index = next((i for i, row in enumerate(rows[:20]) if set(HEADERS).issubset({clean(v) for v in row})), None)
    if header_index is None:
        raise ValueError('未找到西农列表导出表头。请在“我的课表”点击“列表导出”，不要使用“导出打印”。')
    header = [clean(v) for v in rows[header_index]]
    if any(header.count(h) != 1 for h in HEADERS):
        raise ValueError('必要表头存在重复，无法确定列含义。')
    indexes = {h: header.index(h) for h in HEADERS}
    lessons, errors, seen = [], [], set()
    duplicate_count = 0
    for number, row in enumerate(rows[header_index + 1:], header_index + 2):
        if not any(clean(v) for v in row):
            continue
        record = {h: clean(row[i]) if i < len(row) else '' for h, i in indexes.items()}
        try:
            if not record['课程名'] or not record['课程号']:
                raise ValueError('课程名或课程号为空。')
            lesson = Lesson(record['课程号'], record['课程名'], record['课序号'],
                            parse_weeks(record['上课周次']), parse_weekday(record['上课星期']),
                            parse_period(record['开始节次']), parse_period(record['结束节次']), record['教室名称'])
            if lesson.first > lesson.last:
                raise ValueError('结束节次早于开始节次。')
            if lesson in seen:
                duplicate_count += 1
            else:
                lessons.append(lesson)
                seen.add(lesson)
        except ValueError as exc:
            errors.append(f'第 {number} 行：{exc}')
    if errors:
        raise ValueError('存在无法转换的记录，未导出任何课程：\n' + '\n'.join(errors[:20]))
    if not lessons:
        raise ValueError('表格没有可转换的课程。')
    return lessons, duplicate_count


def read_xls(path):
    import xlrd
    path = Path(path)
    if path.stat().st_size > 20 * 1024 * 1024:
        raise ValueError('文件超过 20 MB，请选择原始课表列表导出文件。')
    try:
        book = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as exc:
        raise ValueError('无法读取文件。当前版本支持原始 .xls 列表导出，不支持图片、打印版或 .xlsx。') from exc
    try:
        candidates = []
        for sheet in book.sheets():
            rows = [sheet.row_values(i) for i in range(sheet.nrows)]
            if any(set(HEADERS).issubset({clean(v) for v in row}) for row in rows[:20]):
                candidates.append(rows)
        if len(candidates) != 1:
            raise ValueError('需要且只能有一张包含完整课表列表表头的工作表。请重新进行“列表导出”。')
        return parse_rows(candidates[0])
    finally:
        book.release_resources()


def validate_settings(settings, lessons):
    try:
        monday = date.fromisoformat(settings.get('week1_monday', ''))
    except (TypeError, ValueError):
        raise ValueError('请填写第 1 教学周的周一，格式 YYYY-MM-DD。') from None
    if monday.weekday() != 0:
        raise ValueError('第 1 周起始日期必须是周一；请以本学期校历为准。')
    if not 2000 <= monday.year <= 2100:
        raise ValueError('学期年份应在 2000–2100 之间。')
    used = sorted({p for lesson in lessons for p in range(lesson.first, lesson.last + 1)})
    periods = {}
    for p in used:
        raw = settings.get('periods', {}).get(str(p))
        if not isinstance(raw, (list, tuple)) or len(raw) != 2:
            raise ValueError(f'请填写第 {p} 节上下课时间。')
        try:
            if not all(re.fullmatch(r'\d{2}:\d{2}', x) for x in raw):
                raise ValueError()
            start, end = (time.fromisoformat(x) for x in raw)
        except (TypeError, ValueError):
            raise ValueError(f'第 {p} 节时间格式应为 HH:MM。') from None
        if start >= end:
            raise ValueError(f'第 {p} 节下课时间必须晚于上课时间。')
        periods[p] = (start, end)
    for left, right in zip(used, used[1:]):
        if periods[left][1] > periods[right][0]:
            raise ValueError(f'第 {left} 节与第 {right} 节时间重叠或顺序错误。')
    reminder = settings.get('reminder_minutes', 15)
    if isinstance(reminder, bool) or not isinstance(reminder, int) or not 0 <= reminder <= 1440:
        raise ValueError('提醒分钟数必须是 0–1440 的整数；0 表示不提醒。')
    try:
        excluded = {date.fromisoformat(s) for s in settings.get('excluded_dates', [])}
    except (TypeError, ValueError):
        raise ValueError('停课日期应使用 YYYY-MM-DD，每行一个日期。') from None
    return monday, periods, excluded


def expand_events(lessons, settings):
    monday, periods, excluded = validate_settings(settings, lessons)
    changes = []
    previous = None
    for change in settings.get('clock_changes', []):
        try:
            effective = date.fromisoformat(change['from'])
            _, changed_periods, _ = validate_settings(dict(settings, periods=change['periods']), lessons)
        except (TypeError, KeyError, ValueError) as exc:
            raise ValueError(f'作息切换设置有误：{exc}') from None
        if previous is not None and effective <= previous:
            raise ValueError('作息切换日期必须按时间先后排列且不能重复。')
        changes.append((effective, changed_periods))
        previous = effective
    groups = {}
    for lesson in lessons:
        for week in lesson.weeks:
            day = monday + timedelta(days=(week - 1) * 7 + lesson.weekday)
            if day in excluded:
                continue
            key = (lesson.code, lesson.name, lesson.section, lesson.location, day, week)
            group = groups.setdefault(key, {'lesson': lesson, 'periods': set()})
            group['periods'].update(range(lesson.first, lesson.last + 1))
    events = []
    for key, group in groups.items():
        day, week = key[-2:]
        day_periods = periods
        for effective, changed_periods in changes:
            if day >= effective:
                day_periods = changed_periods
        blocks = []
        for p in sorted(group['periods']):
            if blocks:
                prev = blocks[-1][-1]
                gap = datetime.combine(day, day_periods[p][0]) - datetime.combine(day, day_periods[prev][1])
            if blocks and p == prev + 1 and gap <= timedelta(minutes=30):
                blocks[-1].append(p)
            else:
                blocks.append([p])
        for block in blocks:
            first, last = block[0], block[-1]
            events.append(Event(group['lesson'], week, first, last,
                                datetime.combine(day, day_periods[first][0], CHINA),
                                datetime.combine(day, day_periods[last][1], CHINA)))
    return sorted(events, key=lambda e: (e.start, e.lesson.code, e.lesson.section))


def conflicts(events):
    result = []
    for i, a in enumerate(events):
        for b in events[i + 1:]:
            if b.start >= a.end:
                break
            result.append(f'{a.start:%Y-%m-%d}：{a.lesson.name} 与 {b.lesson.name} 时间重叠')
    return result


def escape_ics(text):
    return str(text).replace('\\', '\\\\').replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\\n').replace(';', '\\;').replace(',', '\\,')


def fold_line(line):
    lines, current, length = [], '', 0
    for char in line:
        size = len(char.encode('utf-8'))
        if length + size > 75:
            lines.append(current)
            current, length = ' ', 1
        current += char
        length += size
    return '\r\n'.join(lines + [current])


def make_ics(events, settings):
    if not events:
        raise ValueError('没有可导出的日程。请检查课程与停课日期。')
    stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')
    lines = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//NWAFU Timetable to Calendar//CN',
             'CALSCALE:GREGORIAN', 'X-WR-CALNAME:' + escape_ics(settings.get('calendar_name') or '西农课表'),
             'X-WR-TIMEZONE:Asia/Shanghai']
    for event in events:
        lesson = event.lesson
        identity = '|'.join([lesson.code, lesson.section, lesson.name, lesson.location, str(event.start.date()), str(event.first), str(event.last)])
        uid = hashlib.sha256(identity.encode('utf-8')).hexdigest()[:32] + '@nwafu-timetable.local'
        description = f'课程号：{lesson.code}\n课序号：{lesson.section}\n第 {event.week} 周，第 {event.first}–{event.last} 节\n时间以北京时间为准。调课和选课变更请重新核对。'
        lines.extend(['BEGIN:VEVENT', 'UID:' + uid, 'DTSTAMP:' + stamp,
                      'DTSTART:' + event.start.astimezone(timezone.utc).strftime('%Y%m%dT%H%M%SZ'),
                      'DTEND:' + event.end.astimezone(timezone.utc).strftime('%Y%m%dT%H%M%SZ'),
                      'SUMMARY:' + escape_ics(lesson.name), 'LOCATION:' + escape_ics(lesson.location),
                      'DESCRIPTION:' + escape_ics(description), 'STATUS:CONFIRMED', 'TRANSP:OPAQUE'])
        reminder = settings.get('reminder_minutes', 15)
        if reminder:
            lines.extend(['BEGIN:VALARM', f'TRIGGER:-PT{reminder}M', 'ACTION:DISPLAY',
                          'DESCRIPTION:' + escape_ics(lesson.name + ' 即将开始'), 'END:VALARM'])
        lines.append('END:VEVENT')
    lines.append('END:VCALENDAR')
    return ('\r\n'.join(fold_line(line) for line in lines) + '\r\n').encode('utf-8')


def write_preview(events, path):
    with open(path, 'w', encoding='utf-8-sig', newline='') as handle:
        writer = csv.writer(handle)
        writer.writerow(['日期', '星期', '开始', '结束', '课程', '地点', '周次', '节次'])
        for e in events:
            values = [e.start.strftime('%Y-%m-%d'), '一二三四五六日'[e.start.weekday()], e.start.strftime('%H:%M'),
                      e.end.strftime('%H:%M'), e.lesson.name, e.lesson.location, e.week, f'{e.first}-{e.last}']
            writer.writerow(["'" + v if isinstance(v, str) and v.startswith(('=', '+', '-', '@', '\t')) else v for v in values])


def main():
    parser = argparse.ArgumentParser(description='离线将西农课表列表导出的 .xls 转为手机日历 .ics')
    parser.add_argument('xls')
    parser.add_argument('--settings', required=True, help='JSON 设置文件')
    parser.add_argument('--output', required=True, help='输出 .ics 路径')
    args = parser.parse_args()
    try:
        lessons, duplicates = read_xls(args.xls)
        settings = json.loads(Path(args.settings).read_text(encoding='utf-8-sig'))
        events = expand_events(lessons, settings)
        Path(args.output).write_bytes(make_ics(events, settings))
        print(f'已导出 {len(events)} 个日程，去除 {duplicates} 条重复记录。')
        for warning in conflicts(events):
            print('时间冲突：' + warning)
    except (ValueError, OSError) as exc:
        parser.exit(1, str(exc) + '\n')


if __name__ == '__main__':
    main()
