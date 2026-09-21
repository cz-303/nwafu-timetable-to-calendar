"""Transcribed from NWAFU 2026–2027 academic calendar teaching-time table.

The calendar does not give seasonal switchover dates. A preset is a clock,
not a claim that the clock applies to all dates in a semester.
"""
WEEK1_MONDAY = '2026-09-07'
WINTER = [('08:00', '08:45'), ('08:55', '09:40'), ('10:10', '10:55'),
          ('11:05', '11:50'), ('14:00', '14:45'), ('14:55', '15:40'),
          ('16:00', '16:45'), ('16:55', '17:40'), ('19:00', '19:45'),
          ('19:50', '20:35'), ('20:40', '21:25')]
SUMMER = [('08:00', '08:45'), ('08:55', '09:40'), ('10:10', '10:55'),
          ('11:05', '11:50'), ('14:30', '15:15'), ('15:25', '16:10'),
          ('16:30', '17:15'), ('17:25', '18:10'), ('19:30', '20:15'),
          ('20:20', '21:05'), ('21:10', '21:55')]


def preset(season):
    clock = {'winter': WINTER, 'summer': SUMMER}[season]
    return {'week1_monday': WEEK1_MONDAY, 'calendar_name': '西农2026秋季课表',
            'reminder_minutes': 15, 'periods': {str(i): list(pair) for i, pair in enumerate(clock, 1)},
            'excluded_dates': []}
