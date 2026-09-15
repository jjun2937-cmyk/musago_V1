"""gui_app.py — 무사고전환 보고서 생성기 (데스크톱 GUI).
PyInstaller로 exe를 만들 때 이 파일을 진입점으로 사용한다.
파이썬 표준 라이브러리(tkinter)만 사용하므로 openpyxl 외 추가 의존성이 없다.
"""
import os
import queue
import sys
import threading
import traceback
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import calc_report_v7
import write_report_v7

APP_TITLE = '무사고전환 보고서 생성기'
MAX_PERIOD_ROWS = 7


class TextRedirector:
    """print() 출력을 큐를 통해 GUI 로그창으로 보낸다(스레드 안전)."""

    def __init__(self, msg_queue):
        self.msg_queue = msg_queue

    def write(self, s):
        if s:
            self.msg_queue.put(('log', s))

    def flush(self):
        pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('760x900')
        self.minsize(680, 700)

        self.data_files = []
        self.mapping_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.ref_date = tk.StringVar(value=date.today().strftime('%Y-%m'))
        today = date.today()
        self.new_sheet_years = tk.StringVar(value=f'{today.year - 1}~{today.year}')
        self.period_vars = []
        for i in range(MAX_PERIOD_ROWS):
            if i < len(calc_report_v7.PERIODS):
                _label, _key, start, end = calc_report_v7.PERIODS[i]
            else:
                start, end = '', ''
            self.period_vars.append((tk.StringVar(value=start), tk.StringVar(value=end)))
        self.msg_queue = queue.Queue()
        self.worker = None

        self._build_widgets()
        self.after(100, self._poll_queue)

    def _build_widgets(self):
        pad = {'padx': 10, 'pady': 6}

        frm_files = ttk.LabelFrame(self, text='1. 분석 데이터 파일 (여러 개 선택 가능)')
        frm_files.pack(fill='x', **pad)

        self.lst_files = tk.Listbox(frm_files, height=6, selectmode='extended')
        self.lst_files.pack(side='left', fill='both', expand=True, padx=(10, 0), pady=10)

        scr = ttk.Scrollbar(frm_files, orient='vertical', command=self.lst_files.yview)
        scr.pack(side='left', fill='y', pady=10)
        self.lst_files.config(yscrollcommand=scr.set)

        btns = ttk.Frame(frm_files)
        btns.pack(side='left', fill='y', padx=10, pady=10)
        ttk.Button(btns, text='파일 추가...', command=self._add_files).pack(fill='x', pady=(0, 4))
        ttk.Button(btns, text='선택 삭제', command=self._remove_selected).pack(fill='x', pady=(0, 4))
        ttk.Button(btns, text='전체 삭제', command=self._clear_files).pack(fill='x')

        frm_map = ttk.LabelFrame(self, text='2. 매핑테이블 파일 ("무사고전환매핑테이블" 시트가 있는 파일)')
        frm_map.pack(fill='x', **pad)
        row_map = ttk.Frame(frm_map)
        row_map.pack(fill='x', padx=10, pady=(10, 2))
        ent_map = ttk.Entry(row_map, textvariable=self.mapping_path)
        ent_map.pack(side='left', fill='x', expand=True)
        ttk.Button(row_map, text='찾아보기...', command=self._choose_mapping).pack(side='left', padx=(6, 0))
        ttk.Label(frm_map, text='비워두면 위 분석 데이터 중 첫 번째 파일을 매핑테이블로도 사용합니다.',
                  foreground='#666666').pack(fill='x', padx=10, pady=(0, 10))

        frm_ref = ttk.LabelFrame(self, text='3. 계산 기준 시점 (YYYY-MM)')
        frm_ref.pack(fill='x', **pad)
        row_ref = ttk.Frame(frm_ref)
        row_ref.pack(fill='x', padx=10, pady=(10, 2))
        ttk.Entry(row_ref, textvariable=self.ref_date, width=12).pack(side='left')
        ttk.Label(frm_ref,
                  text='"월별_부문별_new"/"참고" 시트가 기준으로 삼는 시점입니다. '
                       '기본값은 오늘 날짜이며, 과거 특정 시점 기준으로 계산하고 싶을 때만 바꾸세요.',
                  foreground='#666666').pack(fill='x', padx=10, pady=(0, 10))

        frm_years = ttk.LabelFrame(self, text='4. 월별_부문별_new 표시연도 (YYYY~YYYY)')
        frm_years.pack(fill='x', **pad)
        row_years = ttk.Frame(frm_years)
        row_years.pack(fill='x', padx=10, pady=(10, 2))
        ttk.Entry(row_years, textvariable=self.new_sheet_years, width=14).pack(side='left')
        ttk.Label(frm_years,
                  text='"월별_부문별_new" 시트에 나란히 보여줄 연도 범위입니다(양 끝 연도 포함). '
                       '예: 2025~2026, 2024~2026.',
                  foreground='#666666').pack(fill='x', padx=10, pady=(0, 10))

        frm_periods = ttk.LabelFrame(
            self, text='5. 체결기간별 구간 설정 (최대 7개, 시작~종료 YYYYMM, 비워두면 그 구간은 사용 안 함)')
        frm_periods.pack(fill='x', **pad)
        grid_periods = ttk.Frame(frm_periods)
        grid_periods.pack(fill='x', padx=10, pady=(10, 2))
        ttk.Label(grid_periods, text='구간').grid(row=0, column=0, padx=(0, 6), pady=2)
        ttk.Label(grid_periods, text='시작(YYYYMM)').grid(row=0, column=1, padx=(0, 6), pady=2)
        ttk.Label(grid_periods, text='종료(YYYYMM)').grid(row=0, column=2, pady=2)
        for i, (start_var, end_var) in enumerate(self.period_vars):
            ttk.Label(grid_periods, text=f'{i + 1}구간').grid(row=i + 1, column=0, padx=(0, 6), pady=2, sticky='w')
            ttk.Entry(grid_periods, textvariable=start_var, width=10).grid(row=i + 1, column=1, padx=(0, 6), pady=2)
            ttk.Entry(grid_periods, textvariable=end_var, width=10).grid(row=i + 1, column=2, pady=2)
        ttk.Label(frm_periods, text='보험기간 시작일 기준 12개월 단위 구간입니다. 채워진 구간만 계산에 사용됩니다.',
                  foreground='#666666').pack(fill='x', padx=10, pady=(4, 10))

        frm_out = ttk.LabelFrame(self, text='6. 결과 저장 위치')
        frm_out.pack(fill='x', **pad)
        ent = ttk.Entry(frm_out, textvariable=self.output_path)
        ent.pack(side='left', fill='x', expand=True, padx=(10, 0), pady=10)
        ttk.Button(frm_out, text='찾아보기...', command=self._choose_output).pack(side='left', padx=10, pady=10)

        frm_run = ttk.Frame(self)
        frm_run.pack(fill='x', **pad)
        self.btn_run = ttk.Button(frm_run, text='보고서 생성', command=self._run)
        self.btn_run.pack(side='left')
        self.progress = ttk.Progressbar(frm_run, mode='indeterminate')
        self.progress.pack(side='left', fill='x', expand=True, padx=10)

        frm_log = ttk.LabelFrame(self, text='진행 상황')
        frm_log.pack(fill='both', expand=True, **pad)
        self.txt_log = tk.Text(frm_log, height=12, state='disabled', wrap='word')
        self.txt_log.pack(fill='both', expand=True, padx=10, pady=10)

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title='데이터 파일 선택', filetypes=[('Excel 파일', '*.xlsx *.xlsm *.xls')])
        for p in paths:
            if p not in self.data_files:
                self.data_files.append(p)
                self.lst_files.insert('end', p)

    def _remove_selected(self):
        for i in reversed(self.lst_files.curselection()):
            del self.data_files[i]
            self.lst_files.delete(i)

    def _clear_files(self):
        self.data_files.clear()
        self.lst_files.delete(0, 'end')

    def _choose_mapping(self):
        path = filedialog.askopenfilename(
            title='매핑테이블 파일 선택', filetypes=[('Excel 파일', '*.xlsx *.xlsm *.xls')])
        if path:
            self.mapping_path.set(path)

    def _choose_output(self):
        path = filedialog.asksaveasfilename(
            title='결과 저장 위치', defaultextension='.xlsx',
            filetypes=[('Excel 파일', '*.xlsx')], initialfile='무사고전환_보고서.xlsx')
        if path:
            self.output_path.set(path)

    def _log(self, text):
        self.txt_log.config(state='normal')
        self.txt_log.insert('end', text)
        self.txt_log.see('end')
        self.txt_log.config(state='disabled')

    def _run(self):
        if not self.data_files:
            messagebox.showwarning(APP_TITLE, '데이터 파일을 하나 이상 선택해 주세요.')
            return
        out = self.output_path.get().strip()
        if not out:
            messagebox.showwarning(APP_TITLE, '결과 저장 위치를 지정해 주세요.')
            return

        mapping = self.mapping_path.get().strip() or self.data_files[0]

        ref_date_str = self.ref_date.get().strip()
        try:
            ref_year, ref_month = (int(p) for p in ref_date_str.split('-'))
            if not (1 <= ref_month <= 12):
                raise ValueError
        except ValueError:
            messagebox.showwarning(APP_TITLE, '계산 기준 시점은 "YYYY-MM" 형식으로 입력해 주세요. (예: 2026-09)')
            return

        years_str = self.new_sheet_years.get().strip()
        try:
            y_start, y_end = (int(p) for p in years_str.split('~'))
            if y_start > y_end:
                raise ValueError
        except ValueError:
            messagebox.showwarning(APP_TITLE, '표시연도는 "YYYY~YYYY" 형식으로 입력해 주세요. (예: 2025~2026)')
            return
        new_sheet_years = list(range(y_start, y_end + 1))

        period_pairs = []
        for i, (start_var, end_var) in enumerate(self.period_vars):
            start = start_var.get().strip()
            end = end_var.get().strip()
            if not start and not end:
                continue
            if not (len(start) == 6 and start.isdigit() and len(end) == 6 and end.isdigit() and start <= end):
                messagebox.showwarning(
                    APP_TITLE, f'{i + 1}구간의 시작/종료를 "YYYYMM" 형식으로 올바르게 입력해 주세요. (예: 202207)')
                return
            period_pairs.append((start, end))
        if not period_pairs:
            messagebox.showwarning(APP_TITLE, '체결기간별 구간을 하나 이상 입력해 주세요.')
            return

        self.btn_run.config(state='disabled')
        self.progress.start(12)
        self.txt_log.config(state='normal')
        self.txt_log.delete('1.0', 'end')
        self.txt_log.config(state='disabled')

        self.worker = threading.Thread(
            target=self._run_worker,
            args=(list(self.data_files), mapping, out, ref_year, ref_month, new_sheet_years, period_pairs),
            daemon=True)
        self.worker.start()

    def _run_worker(self, data_paths, mapping_path, out, ref_year, ref_month, new_sheet_years, period_pairs):
        old_stdout = sys.stdout
        sys.stdout = TextRedirector(self.msg_queue)
        try:
            write_report_v7.build(data_paths, mapping_path, out,
                                   reference_year=ref_year, current_month=ref_month,
                                   new_sheet_years=new_sheet_years, period_pairs=period_pairs)
            self.msg_queue.put(('done', out))
        except Exception:
            self.msg_queue.put(('error', traceback.format_exc()))
        finally:
            sys.stdout = old_stdout

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == 'log':
                    self._log(payload)
                elif kind == 'done':
                    self.progress.stop()
                    self.btn_run.config(state='normal')
                    messagebox.showinfo(APP_TITLE, f'보고서 생성이 완료됐습니다.\n\n{payload}')
                elif kind == 'error':
                    self.progress.stop()
                    self.btn_run.config(state='normal')
                    self._log('\n[오류]\n' + payload)
                    messagebox.showerror(APP_TITLE, '보고서 생성 중 오류가 발생했습니다.\n로그 창을 확인해 주세요.')
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)


def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
