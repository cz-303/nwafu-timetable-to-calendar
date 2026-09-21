"""Windows desktop interface. No network access or account credentials needed."""
import json
import sys
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from converter import read_xls, expand_events, conflicts, make_ics, write_preview
from school_presets import WINTER, SUMMER, preset


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('西农课表转手机日历')
        self.geometry('1120x820')
        self.minsize(940, 680)
        self.configure(bg='#f0f3f8')
        self.lessons = []
        self.events = []
        self.file_path = None
        self.preview_signature = None
        self.style = ttk.Style(self)
        self.style.theme_use('clam')
        self.option_add('*Font', ('Microsoft YaHei UI', 10))
        self.style.configure('TFrame', background='#f0f3f8')
        self.style.configure('TLabel', background='#f0f3f8', foreground='#14283f')
        self.style.configure('Title.TLabel', font=('Microsoft YaHei UI', 21, 'bold'))
        self.style.configure('Accent.TButton', background='#244cb0', foreground='white', padding=(14, 8))
        self.style.map('Accent.TButton', background=[('active', '#163985'), ('disabled', '#d4dbea')], foreground=[('disabled', '#66738b')])
        self.style.configure('Treeview', rowheight=32, font=('Microsoft YaHei UI', 10))
        self.style.configure('Treeview.Heading', font=('Microsoft YaHei UI', 10, 'bold'))
        self.style.configure('TButton', padding=(9, 6))
        outer = ttk.Frame(self, padding=22)
        outer.pack(fill='both', expand=True)
        ttk.Label(outer, text='西农课表 → 手机日历', style='Title.TLabel').pack(anchor='w')
        ttk.Label(outer, text='导入“我的课表 → 列表导出”的 XLS，核对后生成日历。全程在本机处理。').pack(anchor='w', pady=(7, 16))
        bar = ttk.Frame(outer)
        bar.pack(fill='x')
        ttk.Button(bar, text='1  选择课表 XLS', style='Accent.TButton', command=self.choose_file).pack(side='left')
        self.file_label = ttk.Label(bar, text='尚未选择文件')
        self.file_label.pack(side='left', padx=12)
        body = ttk.Frame(outer)
        body.pack(fill='both', expand=True, pady=18)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        left = ttk.Frame(body, width=340)
        left.grid(row=0, column=0, sticky='nsew', padx=(0, 20))
        notebook = ttk.Notebook(left, width=330)
        notebook.pack(fill='both', expand=True)
        settings_tab = ttk.Frame(notebook, padding=12)
        times_tab = ttk.Frame(notebook, padding=12)
        season_tab = ttk.Frame(notebook, padding=12)
        notebook.add(settings_tab, text='2  学期设置')
        notebook.add(times_tab, text='节次时间')
        notebook.add(season_tab, text='季节切换')
        self.monday = tk.StringVar(value='2026-09-07')
        self.calendar_name = tk.StringVar(value='西农2026秋季课表')
        self.reminder = tk.StringVar(value='15')
        self.confirmed = tk.BooleanVar(value=False)
        self.change_date = tk.StringVar()
        self.change_season = tk.StringVar(value='冬季')
        ttk.Label(season_tab, text='跨季节时可设置一次作息切换。\n不需要切换时，日期留空。', wraplength=280).pack(anchor='w', pady=(0, 18))
        ttk.Label(season_tab, text='从哪一天起使用新作息').pack(anchor='w')
        ttk.Entry(season_tab, textvariable=self.change_date).pack(fill='x', pady=6)
        ttk.Label(season_tab, text='格式 YYYY-MM-DD，包含当天。').pack(anchor='w', pady=(0, 16))
        ttk.Label(season_tab, text='切换后的作息').pack(anchor='w')
        ttk.Combobox(season_tab, textvariable=self.change_season, values=['冬季', '夏季'], state='readonly').pack(fill='x', pady=6)
        ttk.Label(season_tab, text='切换前使用“节次时间”中的时间；\n切换当天起使用校历里的目标季节时间。\n\n请按当年学校通知填写具体生效日期，不自动推算。', wraplength=280).pack(anchor='w', pady=15)
        ttk.Label(settings_tab, text='第 1 教学周的周一').pack(anchor='w')
        ttk.Entry(settings_tab, textvariable=self.monday).pack(fill='x', pady=(5, 2))
        ttk.Label(settings_tab, text='已填 2026 秋季校历；其他学期请修改').pack(anchor='w', pady=(0, 14))
        ttk.Label(settings_tab, text='日历名称').pack(anchor='w')
        ttk.Entry(settings_tab, textvariable=self.calendar_name).pack(fill='x', pady=(5, 14))
        ttk.Label(settings_tab, text='提前提醒（分钟，0 表示不提醒）').pack(anchor='w')
        ttk.Spinbox(settings_tab, from_=0, to=1440, textvariable=self.reminder, width=12).pack(anchor='w', pady=(5, 14))
        ttk.Label(settings_tab, text='整天停课日期（可选，每行一个）').pack(anchor='w')
        self.excluded = tk.Text(settings_tab, height=4, width=27, relief='solid', borderwidth=1)
        self.excluded.pack(fill='x', pady=(5, 5))
        ttk.Label(settings_tab, text='不自动推断节假日或调休。\n补课请在导入后按学校通知调整。', wraplength=285).pack(anchor='w', pady=(0, 14))
        ttk.Checkbutton(settings_tab, text='已核对校历和节次时间', variable=self.confirmed).pack(anchor='w', pady=8)
        ttk.Button(settings_tab, text='加载设置', command=self.load_settings).pack(fill='x', pady=(10, 5))
        ttk.Button(settings_tab, text='保存设置', command=self.save_settings).pack(fill='x')
        ttk.Label(settings_tab, text='设置文件仅保存日期与时间规则，\n不保存原始课表或个人账号。', wraplength=285).pack(anchor='w', pady=12)
        ttk.Label(times_tab, text='北京时间（24 小时制）\n只需填写课表使用到的节次。').grid(row=0, column=0, columnspan=3, sticky='w', pady=(0, 12))
        for col, name in enumerate(('节次', '上课', '下课')):
            ttk.Label(times_tab, text=name).grid(row=1, column=col, padx=6, sticky='w')
        self.times = {}
        for p in range(1, 13):
            start, end = tk.StringVar(), tk.StringVar()
            self.times[p] = (start, end)
            ttk.Label(times_tab, text=f'第 {p} 节').grid(row=p + 1, column=0, padx=6, pady=5)
            ttk.Entry(times_tab, textvariable=start, width=8).grid(row=p + 1, column=1, padx=6, pady=5)
            ttk.Entry(times_tab, textvariable=end, width=8).grid(row=p + 1, column=2, padx=6, pady=5)
        ttk.Label(times_tab, text='格式 HH:MM，例如 08:00。\n连续节次间隔超过 30 分钟时分段。', wraplength=285).grid(row=14, column=0, columnspan=3, sticky='w', pady=12)
        ttk.Button(times_tab, text='填入冬季作息', command=lambda: self.fill_clock(WINTER)).grid(row=15, column=0, columnspan=2, sticky='ew', padx=3)
        ttk.Button(times_tab, text='夏季作息', command=lambda: self.fill_clock(SUMMER)).grid(row=15, column=2, sticky='ew', padx=3)
        ttk.Label(times_tab, text='校历未注明切换日，请确认适用季节。', wraplength=285).grid(row=16, column=0, columnspan=3, sticky='w', pady=8)
        right = ttk.Frame(body)
        right.grid(row=0, column=1, sticky='nsew')
        self.summary = ttk.Label(right, text='选择文件后，先显示课程安排。', font=('Microsoft YaHei UI', 12, 'bold'), wraplength=640)
        self.summary.pack(anchor='w', pady=(0, 10))
        self.tree = ttk.Treeview(right, columns=('date', 'time', 'course', 'place'), show='headings', selectmode='browse')
        for name, title, width in [('date', '周次 / 日期', 150), ('time', '星期节次 / 时间', 165), ('course', '课程', 280), ('place', '地点', 125)]:
            self.tree.heading(name, text=title)
            self.tree.column(name, width=width, minwidth=80, stretch=True)
        scroll_y = ttk.Scrollbar(right, orient='vertical', command=self.tree.yview)
        scroll_y.pack(side='right', fill='y')
        scroll_x = ttk.Scrollbar(right, orient='horizontal', command=self.tree.xview)
        scroll_x.pack(side='bottom', fill='x')
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree.pack(fill='both', expand=True)
        bottom = ttk.Frame(outer)
        bottom.pack(fill='x')
        ttk.Button(bottom, text='3  生成日期预览', command=self.preview).pack(side='left')
        self.export_button = ttk.Button(bottom, text='4  导出手机日历 .ics', style='Accent.TButton', command=self.export, state='disabled')
        self.export_button.pack(side='left', padx=10)
        self.csv_button = ttk.Button(bottom, text='导出核对清单', command=self.export_csv, state='disabled')
        self.csv_button.pack(side='left')
        ttk.Button(bottom, text='安卓导入说明', command=self.help).pack(side='right')
        self.status = ttk.Label(outer, text='尚未生成日历。请先导入课表，并填写学期日期与节次时间。', wraplength=1020)
        self.status.pack(anchor='w', pady=(12, 0))
        for var in [self.monday, self.calendar_name, self.reminder, self.change_date, self.change_season] + [v for pair in self.times.values() for v in pair]:
            var.trace_add('write', self.invalidate)
        self.excluded.bind('<<Modified>>', self.text_changed)

    def text_changed(self, _=None):
        if self.excluded.edit_modified():
            self.invalidate()
            self.excluded.edit_modified(False)

    def fill_clock(self, clock):
        for p, (start, end) in self.times.items():
            pair = clock[p - 1] if p <= len(clock) else ('', '')
            start.set(pair[0])
            end.set(pair[1])
        self.confirmed.set(False)
        self.status.configure(text='已填入校历作息。请确认适用季节，并重新生成日期预览。')

    def invalidate(self, *_):
        self.preview_signature = None
        self.export_button.configure(state='disabled')
        self.csv_button.configure(state='disabled')
        if self.events:
            self.events = []
            self.show_lessons()
            self.status.configure(text='设置已变化，请重新生成日期预览。')

    def settings(self):
        try:
            reminder = int(self.reminder.get())
        except ValueError:
            raise ValueError('提醒分钟数必须是整数。') from None
        changes = []
        if self.change_date.get().strip():
            season = 'winter' if self.change_season.get() == '冬季' else 'summer'
            changes = [{'from': self.change_date.get().strip(), 'season': season, 'periods': preset(season)['periods']}]
        return {'week1_monday': self.monday.get().strip(), 'calendar_name': self.calendar_name.get().strip(),
                'reminder_minutes': reminder, 'periods': {str(p): [a.get().strip(), b.get().strip()] for p, (a, b) in self.times.items()},
                'excluded_dates': [line.strip() for line in self.excluded.get('1.0', 'end').splitlines() if line.strip()], 'clock_changes': changes}

    def choose_file(self):
        path = filedialog.askopenfilename(title='选择“列表导出”的 XLS', filetypes=[('Excel 97–2003 课表', '*.xls')])
        if path:
            self.import_file(path)

    def import_file(self, path):
        self.invalidate()
        self.lessons, self.file_path = [], None
        self.tree.delete(*self.tree.get_children())
        self.file_label.configure(text='尚未选择文件')
        self.summary.configure(text='正在读取课表…')
        try:
            lessons, duplicate_count = read_xls(path)
            if max(lesson.last for lesson in lessons) > 12:
                raise ValueError('界面支持第 1–12 节。更多节次可用命令行与自定义设置文件转换。')
            self.lessons, self.file_path = lessons, path
            self.file_label.configure(text=Path(path).name)
            self.show_lessons()
            self.status.configure(text=f'已读取 {len(lessons)} 条记录，去除 {duplicate_count} 条完全重复记录。请填写学期设置与节次时间。')
        except (OSError, ValueError) as exc:
            self.summary.configure(text='课表未导入')
            self.status.configure(text='请重新选择有效的列表导出文件。')
            messagebox.showerror('导入失败', str(exc))

    def show_lessons(self):
        self.tree.delete(*self.tree.get_children())
        grouped = {}
        for lesson in self.lessons:
            key = (lesson.code, lesson.name, lesson.section, lesson.weeks, lesson.weekday, lesson.location)
            grouped.setdefault(key, set()).update(range(lesson.first, lesson.last + 1))
        for key, periods in grouped.items():
            _, name, section, weeks, weekday, location = key
            self.tree.insert('', 'end', values=(','.join(map(str, weeks)) + ' 周', f'周{"一二三四五六日"[weekday]} / ' + ','.join(map(str, sorted(periods))) + ' 节', f'{name} [{section}]', location or '地点待定'))
        count = len({(l.code, l.section) for l in self.lessons})
        self.summary.configure(text=f'{count} 门课程 · {len(self.lessons)} 条原始节次记录')

    def preview(self):
        try:
            self.invalidate()
            if not self.lessons:
                raise ValueError('请先导入课表 XLS。')
            settings = self.settings()
            self.events = expand_events(self.lessons, settings)
            if not self.events:
                raise ValueError('所有课程日期已被排除，没有可导出的日程。')
            self.tree.delete(*self.tree.get_children())
            for event in self.events:
                self.tree.insert('', 'end', values=(event.start.strftime('%Y-%m-%d') + ' 周' + '一二三四五六日'[event.start.weekday()], f'{event.start:%H:%M}–{event.end:%H:%M}', event.lesson.name, event.lesson.location or '地点待定'))
            self.preview_signature = json.dumps(settings, sort_keys=True)
            self.summary.configure(text=f'{len(self.events)} 个日程 · {self.events[0].start:%m月%d日} 至 {self.events[-1].start:%m月%d日}')
            warnings = conflicts(self.events)
            self.status.configure(text=f'已生成北京时间预览。发现 {len(warnings)} 处时间冲突。' + (' 请核对后导出。' if warnings else ' 核对日期、教室和时间后即可导出。'))
            self.export_button.configure(state='normal')
            self.csv_button.configure(state='normal')
            if warnings:
                messagebox.showwarning('课程时间冲突', '\n'.join(warnings[:15]))
        except (ValueError, OSError) as exc:
            messagebox.showerror('无法生成预览', str(exc))

    def export(self):
        try:
            settings = self.settings()
            if not self.events or self.preview_signature != json.dumps(settings, sort_keys=True):
                raise ValueError('请先重新生成日期预览。')
            if not self.confirmed.get():
                raise ValueError('请核对日期和节次时间，并勾选“已核对校历和节次时间”。')
            data = make_ics(self.events, settings)
            path = filedialog.asksaveasfilename(title='保存手机日历', defaultextension='.ics', initialfile='西农课表.ics', filetypes=[('日历文件', '*.ics')])
            if path:
                Path(path).write_bytes(data)
                self.status.configure(text=f'已导出 {len(self.events)} 个日程：{path}')
                messagebox.showinfo('导出完成', '将 .ics 文件发送到手机后，用支持导入的日历应用打开。\n建议导入到单独的“西农课表”日历，避免与个人日程混在一起。\n详细方法见“安卓导入说明”。')
        except (ValueError, OSError) as exc:
            messagebox.showerror('导出失败', str(exc))

    def export_csv(self):
        if not self.events:
            return
        path = filedialog.asksaveasfilename(defaultextension='.csv', initialfile='课表核对清单.csv', filetypes=[('核对清单', '*.csv')])
        if path:
            try:
                write_preview(self.events, path)
            except OSError as exc:
                messagebox.showerror('保存失败', str(exc))

    def save_settings(self):
        try:
            settings = self.settings()
            path = filedialog.asksaveasfilename(defaultextension='.json', initialfile='我的学期设置.json', filetypes=[('学期设置', '*.json')])
            if path:
                Path(path).write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding='utf-8')
        except (ValueError, OSError) as exc:
            messagebox.showerror('保存失败', str(exc))

    def load_settings(self):
        path = filedialog.askopenfilename(filetypes=[('学期设置', '*.json')])
        if not path:
            return
        try:
            data = json.loads(Path(path).read_text(encoding='utf-8-sig'))
            if not isinstance(data, dict) or not isinstance(data.get('periods', {}), dict):
                raise ValueError('设置格式错误。')
            pairs = {p: data.get('periods', {}).get(str(p), ['', '']) for p in self.times}
            if any(not isinstance(pair, list) or len(pair) != 2 or not all(isinstance(v, str) for v in pair) for pair in pairs.values()):
                raise ValueError('节次时间必须为两个时间字符串。')
            excluded = data.get('excluded_dates', [])
            if not isinstance(excluded, list) or not all(isinstance(v, str) for v in excluded):
                raise ValueError('停课日期格式错误。')
            changes = data.get('clock_changes', [])
            if not isinstance(changes, list) or len(changes) > 1:
                raise ValueError('桌面界面支持一次季节切换；多次切换请使用命令行。')
            if changes and (not isinstance(changes[0], dict) or changes[0].get('season') not in ('winter', 'summer') or changes[0].get('periods') != preset(changes[0]['season'])['periods']):
                raise ValueError('桌面界面仅支持校历预设的季节切换；自定义切换时间请使用命令行。')
            self.change_date.set(changes[0].get('from', '') if changes else '')
            self.change_season.set('夏季' if changes and changes[0]['season'] == 'summer' else '冬季')
            self.monday.set(data.get('week1_monday', ''))
            self.calendar_name.set(data.get('calendar_name', '西农课表'))
            self.reminder.set(str(data.get('reminder_minutes', 15)))
            for p, (start, end) in self.times.items():
                start.set(pairs[p][0])
                end.set(pairs[p][1])
            self.excluded.delete('1.0', 'end')
            self.excluded.insert('1.0', '\n'.join(excluded))
            self.confirmed.set(False)
        except (ValueError, OSError, TypeError) as exc:
            messagebox.showerror('设置加载失败', str(exc))

    def help(self):
        messagebox.showinfo('安卓日历导入', '1. 导出 .ics 文件，通过 USB、文件传输或你自己的通信工具发送到手机。\n\n2. 在手机“文件管理”中打开文件，选择支持 ICS 导入的日历；或查找日历设置中的“导入日程”。不同品牌和版本入口不同，部分系统日历不支持直接导入。\n\n3. 若无入口，可在电脑端 Google 日历网页导入，再在手机同步同一 Google 账号（需能访问 Google 服务）。\n\n4. 建议使用单独的课表日历。重复导入可能产生重复事件；更新前先清空旧的课表日历，不要清空个人日历。\n\n5. 导入后检查一门普通课和一门周末课，并确认日历通知权限已开启。调课和新增课程不会自动同步。')


if __name__ == '__main__':
    app = App()
    if '--smoke-test' in sys.argv:
        app.withdraw()
        import xlrd
        app.update_idletasks()
        app.destroy()
        sys.exit(0)
    if len(sys.argv) > 1:
        app.after(100, lambda: app.import_file(sys.argv[1]))
    app.mainloop()
