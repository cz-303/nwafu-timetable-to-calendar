import unittest
from datetime import date, timedelta
from converter import (HEADERS, Lesson, parse_weeks, parse_rows, parse_weekday,
                       expand_events, make_ics, conflicts, fold_line)


def fixture_settings():
    # Synthetic clock for regression tests; not a published university schedule.
    return {'week1_monday': '2026-09-07', 'reminder_minutes': 15,
            'periods': {'1': ['08:00', '08:45'], '2': ['08:55', '09:40'],
                        '3': ['10:00', '10:45'], '4': ['10:55', '11:40'],
                        '5': ['14:00', '14:45'], '6': ['14:55', '15:40'],
                        '7': ['16:00', '16:45'], '8': ['16:55', '17:40']}}


def lesson(**overrides):
    fields = dict(code='DEMO001', name='示例课程', section='01', weeks=(1, 3), weekday=0, first=1, last=2, location='示例教室')
    fields.update(overrides)
    return Lesson(**fields)


class ConverterTests(unittest.TestCase):
    def test_weeks_and_parity(self):
        self.assertEqual(parse_weeks('1-8周(单)'), (1, 3, 5, 7))
        self.assertEqual(parse_weeks('第2～8周（双周）'), (2, 4, 6, 8))
        self.assertEqual(parse_weeks('1-3,5,7-8周'), (1, 2, 3, 5, 7, 8))
        self.assertEqual(parse_weeks('12周'), (12,))

    def test_invalid_weeks_fail_closed(self):
        for text in ('待定', '0周', '54周', '8-2周', '1-8周另行通知', '1周(双)', '1,,2', '1-3周/5周'):
            with self.subTest(text=text), self.assertRaises(ValueError):
                parse_weeks(text)

    def test_weekdays(self):
        self.assertEqual(parse_weekday('星期日'), 6)
        self.assertEqual(parse_weekday('周一'), 0)
        with self.assertRaises(ValueError):
            parse_weekday('星期八')

    def test_rows_duplicates_and_errors(self):
        row = ['DEMO001', '示例课程', '01', '1-2周', '星期一', '第1节', '第2节', 'A101']
        lessons, duplicate_count = parse_rows([list(HEADERS), row, row, []])
        self.assertEqual((len(lessons), duplicate_count), (1, 1))
        invalid = row.copy()
        invalid[3] = '待定'
        with self.assertRaisesRegex(ValueError, '第 3 行'):
            parse_rows([list(HEADERS), row, invalid])

    def test_dates_and_sunday(self):
        events = expand_events([lesson(weeks=(12,), weekday=6)], fixture_settings())
        self.assertEqual(events[0].start.date(), date(2026, 11, 29))
        self.assertEqual(events[0].start.utcoffset(), timedelta(hours=8))

    def test_merge_and_lunch_split(self):
        rows = [lesson(weeks=(1,), first=p, last=p) for p in range(1, 9)]
        events = expand_events(rows + [rows[0]], fixture_settings())
        self.assertEqual([(e.first, e.last) for e in events], [(1, 4), (5, 8)])

    def test_nonconsecutive_not_merged(self):
        events = expand_events([lesson(weeks=(1,), first=p, last=p) for p in (1, 3)], fixture_settings())
        self.assertEqual(len(events), 2)

    def test_exclude_and_missing_clock(self):
        settings = fixture_settings()
        settings['excluded_dates'] = ['2026-09-07']
        self.assertEqual(len(expand_events([lesson()], settings)), 1)
        del settings['periods']['2']
        with self.assertRaisesRegex(ValueError, '第 2 节'):
            expand_events([lesson()], settings)

    def test_bad_monday_and_overlapping_clock(self):
        settings = fixture_settings()
        settings['week1_monday'] = '2026-09-08'
        with self.assertRaisesRegex(ValueError, '周一'):
            expand_events([lesson()], settings)
        settings = fixture_settings()
        settings['periods']['2'] = ['08:20', '09:20']
        with self.assertRaisesRegex(ValueError, '重叠'):
            expand_events([lesson()], settings)

    def test_conflicts_reported(self):
        events = expand_events([lesson(), lesson(code='DEMO002', name='另一门示例课')], fixture_settings())
        self.assertEqual(len(conflicts(events)), 2)

    def test_ics_encoding_utc_and_alarm(self):
        settings = fixture_settings()
        events = expand_events([lesson(name='中文课程,分号;换行\n' * 12)], settings)
        data = make_ics(events, settings)
        self.assertIn(b'DTSTART:20260907T000000Z', data)
        self.assertIn(b'TRIGGER:-PT15M', data)
        self.assertIn(b'\\,', data)
        self.assertTrue(all(len(line) <= 75 for line in data.split(b'\r\n')))
        self.assertEqual(data.count(b'BEGIN:VEVENT'), 2)
        self.assertTrue(data.endswith(b'END:VCALENDAR\r\n'))
        settings['reminder_minutes'] = 0
        self.assertNotIn(b'VALARM', make_ics(events, settings))

    def test_uid_stable_and_unique(self):
        settings = fixture_settings()
        events = expand_events([lesson(), lesson(location='另一教室')], settings)
        first = [line for line in make_ics(events, settings).split(b'\r\n') if line.startswith(b'UID:')]
        second = [line for line in make_ics(events, settings).split(b'\r\n') if line.startswith(b'UID:')]
        self.assertEqual(first, second)
        self.assertEqual(len(first), len(set(first)))

    def test_folding_unicode_round_trip(self):
        original = 'SUMMARY:' + '研究生示例课程' * 30
        folded = fold_line(original)
        self.assertEqual(folded.replace('\r\n ', ''), original)

    def test_season_switch_inclusive(self):
        from school_presets import preset
        settings = preset('summer')
        settings['clock_changes'] = [{'from': '2026-10-12', 'periods': preset('winter')['periods']}]
        events = expand_events([lesson(weeks=(5, 6), first=5, last=8)], settings)
        self.assertEqual([e.start.strftime('%H:%M') for e in events], ['14:30', '14:00'])
        self.assertEqual([e.end.strftime('%H:%M') for e in events], ['18:10', '17:40'])
        settings['clock_changes'][0]['from'] = '2026-10-13'
        events = expand_events([lesson(weeks=(5, 6), first=5, last=8)], settings)
        self.assertEqual([e.start.strftime('%H:%M') for e in events], ['14:30', '14:30'])


if __name__ == '__main__':
    unittest.main()
