"""가족 입출금 내역 정리 프로그램

은행 문자 스크린샷(농협/우리은행/하나은행)을 OCR로 읽어 엑셀로 정리한다.
"""
import os
import re
import sys
import glob
import datetime as dt
import threading
import traceback

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from PIL import Image
import openpyxl
from openpyxl import Workbook, load_workbook

import winocr
from winrt.windows.media.ocr import OcrEngine
from winrt.windows.globalization import Language


# ----------------------------------------------------------------------
# 파싱 규칙
#
# 실제 카카오톡 캡처로 확인한 특징:
#  - Windows OCR 결과는 줄바꿈 없이 공백으로 이어진 한 줄 텍스트로 나온다.
#  - 숫자 "2,"가 종종 한글 "기" 한 글자로 잘못 인식된다 (예: "2,873,850"->"기873,850").
#  - 시간의 콜론이 자주 사라지거나("15:34"->"1534"), 앞자리 숫자가 공백으로
#    분리된다("18:09"->"1 8:09").
#  - 화면에는 "확인된 발신번호", "거래내역 바로가기", "알림", 날짜 구분선
#    같은 채팅 UI 문구도 함께 인식된다.
# ----------------------------------------------------------------------

BANK_NAMES = {'우리': '우리은행', '하나': '하나은행', '농협': '농협'}
BANK_ORDER = ['농협', '우리은행', '하나은행']

ACCOUNT_RE = re.compile(
    r'(?:\*\d{5,8}|\d{2,3}\*{3,8}\d{3,6}|\d{2,4}-\*{2,6}-\d{2,6}-\d{1,4})'
)
DATE_RE = re.compile(r'(\d{1,2})/(\d{1,2})')
AMOUNT_TOKEN = r'(?:[\d,]|기)+'
TXN_RE = re.compile(r'(입금|출금)\s*(' + AMOUNT_TOKEN + r')\s*원')
BALANCE_RE = re.compile(r'잔액\s*(' + AMOUNT_TOKEN + r')\s*원')
YEAR_DIVIDER_RE = re.compile(r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일')

ANCHOR_RE = re.compile(
    r'(?:(우리|하나)(?=\s*\d{1,2}/\d{1,2}))'
    r'|(?:(농협)(?=\s*(?:입금|출금)\s*(?:[\d,]|기)+\s*원))'
)

TOKEN_BLOCKLIST_RE = re.compile(
    r'확인|발신|번호|바로가기|거래내역|^알림$|^오전$|^오후$|^입금$|^출금$|^잔액|요일$|^\d+월$|^\d+일$'
)


def normalize_amount(raw: str) -> int:
    """OCR이 '2,'를 '기'로 잘못 읽는 현상을 보정하여 정수 금액으로 변환."""
    fixed = raw.replace('기', '2,').replace(',', '')
    return int(fixed)


def extract_time(block: str, start: int):
    """날짜 매칭 바로 뒤 구간에서 시간을 찾는다. 못 찾으면 None.

    OCR이 콜론 앞뒤 숫자 사이에 공백을 끼워 넣는 경우(예: "1 8:09",
    "1 1 :1 5")가 흔해 각 숫자 사이 공백을 허용하는 패턴을 우선 시도하고,
    콜론 자체가 사라진 경우("1534")를 마지막에 시도한다.
    """
    window = block[start:start + 14]
    m = re.match(r'\s*(\d)\s*(\d)\s*[:•.]\s*(\d)\s*(\d)', window)
    if m:
        return f"{m.group(1)}{m.group(2)}:{m.group(3)}{m.group(4)}"
    m = re.match(r'\s*(\d{2})(\d{2})(?!\d)', window)
    if m:
        return f"{m.group(1)}:{m.group(2)}"
    return None


def extract_description(remainder: str) -> str:
    for tok in remainder.split():
        if TOKEN_BLOCKLIST_RE.search(tok):
            continue
        letters = re.sub(r'[^가-힣A-Za-z]', '', tok)
        if len(letters) < 2:
            continue
        return tok
    return ''


def find_year(full_text: str, pos: int, default_year: int) -> int:
    best = None
    for m in YEAR_DIVIDER_RE.finditer(full_text):
        if m.start() <= pos:
            best = int(m.group(1))
        else:
            break
    return best or default_year


def split_messages(text: str):
    """OCR 텍스트를 은행 토큰 기준 메시지 블록들로 분리한다."""
    matches = list(ANCHOR_RE.finditer(text))
    blocks = []
    for i, m in enumerate(matches):
        bank_token = m.group(1) or m.group(2)
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        blocks.append((bank_token, start, text[start:end]))
    return blocks


def parse_block(bank_token: str, block: str, block_start: int, full_text: str,
                 default_year: int, source_file: str):
    """블록 하나를 파싱한다. 성공 시 (record, None), 실패 시 (None, 사유)."""
    bank = BANK_NAMES[bank_token]

    dm = DATE_RE.search(block)
    if not dm:
        return None, '날짜를 찾지 못함'
    month, day = int(dm.group(1)), int(dm.group(2))

    txn = TXN_RE.search(block)
    if not txn:
        return None, '입금/출금 금액을 찾지 못함'

    bal = BALANCE_RE.search(block)
    if not bal:
        return None, '잔액을 찾지 못함'

    try:
        amount = normalize_amount(txn.group(2))
        balance = normalize_amount(bal.group(1))
    except ValueError:
        return None, '금액 숫자 변환 실패'

    time_str = extract_time(block, dm.end())
    acct_m = ACCOUNT_RE.search(block)
    account = acct_m.group(0) if acct_m else ''

    year = find_year(full_text, block_start, default_year)
    try:
        date_val = dt.date(year, month, day)
    except ValueError:
        return None, '날짜 값이 올바르지 않음'

    # 적요 추출을 위해 이미 인식한 구간을 공백으로 지운다
    chars = list(block)

    def blank(span):
        s, e = span
        for i in range(s, e):
            chars[i] = ' '

    blank((0, len(bank_token)))
    blank(dm.span())
    blank(txn.span())
    blank(bal.span())
    if acct_m:
        blank(acct_m.span())
    remainder = ''.join(chars)
    desc = extract_description(remainder)

    record = {
        '은행': bank,
        '날짜': date_val,
        '시간': time_str or '',
        '구분': txn.group(1),
        '금액': amount,
        '적요': desc,
        '잔액': balance,
        '계좌': account,
        '확인필요': '',
        '출처파일': source_file,
        'OCR원문': block.strip(),
    }
    return record, None


def parse_ocr_text(text: str, source_file: str, default_year: int):
    """이미지 한 장의 OCR 텍스트에서 레코드/미인식 목록을 뽑아낸다."""
    records = []
    unrecognized = []
    for bank_token, start, block in split_messages(text):
        record, reason = parse_block(bank_token, block, start, text, default_year, source_file)
        if record is None:
            unrecognized.append({
                '은행': BANK_NAMES.get(bank_token, bank_token),
                '사유': reason,
                '출처파일': source_file,
                'OCR원문': block.strip(),
            })
        else:
            records.append(record)
    return records, unrecognized


# ----------------------------------------------------------------------
# OCR
# ----------------------------------------------------------------------

def check_korean_ocr_available() -> bool:
    try:
        return bool(OcrEngine.is_language_supported(Language('ko')))
    except Exception:
        return False


def run_ocr_on_image(path: str) -> str:
    img = Image.open(path)
    result = winocr.recognize_pil_sync(img, lang='ko')
    return result.get('text', '')


# ----------------------------------------------------------------------
# 엑셀 저장
# ----------------------------------------------------------------------

HEADERS = ['은행', '날짜', '시간', '구분', '금액', '적요', '잔액', '계좌', '확인필요', '출처파일', 'OCR원문']
UNRECOG_HEADERS = ['은행', '사유', '출처파일', 'OCR원문']


def get_excel_path() -> str:
    folder = os.path.join(os.path.expanduser('~'), 'Documents', '가족_입출금정리')
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, '거래내역.xlsx')


def row_key(row: dict):
    return (row['은행'], row['날짜'], row['시간'], row['구분'], row['금액'], row['잔액'])


def load_existing_rows(wb) -> list:
    if '전체내역' not in wb.sheetnames:
        return []
    ws = wb['전체내역']
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        row = dict(zip(HEADERS, r))
        if isinstance(row['날짜'], dt.datetime):
            row['날짜'] = row['날짜'].date()
        rows.append(row)
    return rows


def load_existing_unrecognized(wb) -> list:
    if '미인식' not in wb.sheetnames:
        return []
    ws = wb['미인식']
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] is None:
            continue
        rows.append(dict(zip(UNRECOG_HEADERS, r)))
    return rows


def apply_reconciliation(rows: list):
    groups = {}
    for row in rows:
        key = (row['은행'], row['계좌']) if row['계좌'] else (row['은행'], '')
        groups.setdefault(key, []).append(row)

    for key, group_rows in groups.items():
        group_rows.sort(key=lambda r: (r['날짜'], r['시간'] or '99:99'))
        prev_balance = None
        for row in group_rows:
            if prev_balance is None:
                row['확인필요'] = ''
            else:
                delta = row['금액'] if row['구분'] == '입금' else -row['금액']
                expected = prev_balance + delta
                row['확인필요'] = '' if expected == row['잔액'] else '예'
            prev_balance = row['잔액']


def _reset_sheet(wb, name):
    if name in wb.sheetnames:
        del wb[name]
    return wb.create_sheet(name)


def write_all_sheet(wb, rows):
    ws = _reset_sheet(wb, '전체내역')
    ws.append(HEADERS)
    for row in sorted(rows, key=lambda r: (r['날짜'], r['시간'] or '', r['은행'])):
        ws.append([row[h] for h in HEADERS])
    _format_amount_columns(ws, ['금액', '잔액'])


def write_bank_sheets(wb, rows):
    for bank in BANK_ORDER:
        ws = _reset_sheet(wb, bank)
        ws.append(HEADERS)
        bank_rows = [r for r in rows if r['은행'] == bank]
        for row in sorted(bank_rows, key=lambda r: (r['날짜'], r['시간'] or '')):
            ws.append([row[h] for h in HEADERS])
        _format_amount_columns(ws, ['금액', '잔액'])


def write_daily_summary(wb, rows):
    ws = _reset_sheet(wb, '일별합계')
    header = ['날짜']
    for bank in BANK_ORDER:
        header += [f'{bank}_입금', f'{bank}_출금']
    header += ['합계_입금', '합계_출금']
    ws.append(header)

    totals = {}
    for row in rows:
        d = row['날짜']
        totals.setdefault(d, {b: {'입금': 0, '출금': 0} for b in BANK_ORDER})
        totals[d][row['은행']][row['구분']] += row['금액']

    for d in sorted(totals.keys()):
        line = [d]
        sum_in = sum_out = 0
        for bank in BANK_ORDER:
            v_in = totals[d][bank]['입금']
            v_out = totals[d][bank]['출금']
            sum_in += v_in
            sum_out += v_out
            line += [v_in, v_out]
        line += [sum_in, sum_out]
        ws.append(line)
    _format_amount_columns(ws, header[1:])


def write_desc_summary(wb, rows):
    ws = _reset_sheet(wb, '적요별합계')
    ws.append(['은행', '적요', '입금합계', '입금건수', '출금합계', '출금건수'])

    agg = {}
    for row in rows:
        desc = row['적요'] or '(미확인)'
        key = (row['은행'], desc)
        agg.setdefault(key, {'입금합계': 0, '입금건수': 0, '출금합계': 0, '출금건수': 0})
        if row['구분'] == '입금':
            agg[key]['입금합계'] += row['금액']
            agg[key]['입금건수'] += 1
        else:
            agg[key]['출금합계'] += row['금액']
            agg[key]['출금건수'] += 1

    items = sorted(agg.items(), key=lambda kv: (kv[0][0], -(kv[1]['입금합계'] + kv[1]['출금합계'])))
    for (bank, desc), v in items:
        ws.append([bank, desc, v['입금합계'], v['입금건수'], v['출금합계'], v['출금건수']])
    _format_amount_columns(ws, ['입금합계', '출금합계'])


def write_unrecognized_sheet(wb, rows):
    ws = _reset_sheet(wb, '미인식')
    ws.append(UNRECOG_HEADERS)
    for row in rows:
        ws.append([row[h] for h in UNRECOG_HEADERS])


def _format_amount_columns(ws, col_names):
    header = [c.value for c in ws[1]]
    for name in col_names:
        if name not in header:
            continue
        idx = header.index(name) + 1
        for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
            for cell in row:
                cell.number_format = '#,##0'


def save_all(new_records: list, new_unrecognized: list):
    path = get_excel_path()
    if os.path.exists(path):
        wb = load_workbook(path)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    existing_rows = load_existing_rows(wb)
    existing_keys = {row_key(r) for r in existing_rows}

    added = 0
    all_rows = existing_rows
    for rec in new_records:
        k = row_key(rec)
        if k in existing_keys:
            continue
        existing_keys.add(k)
        all_rows.append(rec)
        added += 1

    all_unrecognized = load_existing_unrecognized(wb) + new_unrecognized

    apply_reconciliation(all_rows)
    write_all_sheet(wb, all_rows)
    write_bank_sheets(wb, all_rows)
    write_daily_summary(wb, all_rows)
    write_desc_summary(wb, all_rows)
    write_unrecognized_sheet(wb, all_unrecognized)

    wb.save(path)
    review_count = sum(1 for r in all_rows if r['확인필요'] == '예')
    return added, review_count, path


# ----------------------------------------------------------------------
# GUI
# ----------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        root.title('가족 입출금 내역 정리')
        root.geometry('560x420')

        self.selected_files = []

        tk.Label(root, text='농협 / 우리은행 / 하나은행 문자 스크린샷을 선택해서 처리하세요.',
                 font=('맑은 고딕', 11)).pack(pady=(14, 4))

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=6)

        self.select_btn = tk.Button(btn_frame, text='이미지 선택', width=16,
                                     command=self.on_select)
        self.select_btn.grid(row=0, column=0, padx=6)

        self.process_btn = tk.Button(btn_frame, text='문자 저장하기', width=16,
                                      command=self.on_process, state='disabled')
        self.process_btn.grid(row=0, column=1, padx=6)

        self.open_btn = tk.Button(btn_frame, text='엑셀 열기', width=16,
                                   command=self.on_open_excel)
        self.open_btn.grid(row=0, column=2, padx=6)

        self.selected_label = tk.Label(root, text='선택된 이미지: 0개')
        self.selected_label.pack(pady=(4, 8))

        self.log = scrolledtext.ScrolledText(root, height=16, width=68, state='disabled')
        self.log.pack(padx=10, pady=6, fill='both', expand=True)

        self.log_line('한국어 OCR 기능을 확인하는 중...')
        if not check_korean_ocr_available():
            self.log_line('[오류] 한국어 OCR 언어팩이 없습니다.')
            self.log_line('Windows 설정 > 시간 및 언어 > 언어 및 지역 > "한국어" 옵션 > '
                           '"광학 문자 인식" 기능을 추가한 뒤 다시 실행해주세요.')
        else:
            self.log_line('준비 완료. "이미지 선택" 버튼을 눌러 시작하세요.')

    def log_line(self, text):
        self.log.configure(state='normal')
        self.log.insert('end', text + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def on_select(self):
        files = filedialog.askopenfilenames(
            title='은행 문자 스크린샷 선택',
            filetypes=[('이미지 파일', '*.jpg *.jpeg *.png *.bmp'), ('모든 파일', '*.*')],
        )
        if not files:
            return
        self.selected_files = list(files)
        self.selected_label.config(text=f'선택된 이미지: {len(self.selected_files)}개')
        self.process_btn.config(state='normal')
        self.log_line(f'{len(self.selected_files)}개 이미지를 선택했습니다.')

    def on_open_excel(self):
        path = get_excel_path()
        if not os.path.exists(path):
            messagebox.showinfo('안내', '아직 저장된 엑셀 파일이 없습니다. 먼저 문자를 저장해주세요.')
            return
        os.startfile(path)

    def on_process(self):
        if not self.selected_files:
            return
        self.select_btn.config(state='disabled')
        self.process_btn.config(state='disabled')
        self.log_line('처리를 시작합니다...')
        thread = threading.Thread(target=self._process_worker, daemon=True)
        thread.start()

    def _process_worker(self):
        try:
            files = self.selected_files
            default_year = dt.date.today().year
            all_records = []
            all_unrecognized = []
            for path in files:
                name = os.path.basename(path)
                try:
                    text = run_ocr_on_image(path)
                except Exception as e:
                    self.root.after(0, self.log_line, f'[{name}] OCR 실패: {e}')
                    continue
                records, unrecognized = parse_ocr_text(text, name, default_year)
                all_records.extend(records)
                all_unrecognized.extend(unrecognized)
                self.root.after(0, self.log_line,
                                 f'[{name}] 인식 {len(records)}건, 미인식 {len(unrecognized)}건')

            added, review_count, path = save_all(all_records, all_unrecognized)
            self.root.after(0, self._on_done, added, review_count, len(all_unrecognized))
        except Exception:
            err = traceback.format_exc()
            self.root.after(0, self._on_error, err)

    def _on_done(self, added, review_count, unrecog_count):
        self.log_line(f'완료: 신규 {added}건 저장, 확인 필요 {review_count}건, 미인식 {unrecog_count}건')
        self.select_btn.config(state='normal')
        self.process_btn.config(state='normal')
        messagebox.showinfo(
            '처리 완료',
            f'신규 거래 {added}건을 저장했습니다.\n'
            f'확인 필요(잔액 불일치 의심): {review_count}건\n'
            f'미인식(수동 확인 필요): {unrecog_count}건\n\n'
            f'엑셀 파일: {get_excel_path()}'
        )

    def _on_error(self, err):
        self.log_line('오류가 발생했습니다:')
        self.log_line(err)
        self.select_btn.config(state='normal')
        self.process_btn.config(state='normal')
        messagebox.showerror('오류', '처리 중 오류가 발생했습니다. 로그를 확인해주세요.')


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()
