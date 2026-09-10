# -*- coding: utf-8 -*-
"""配車入力から配車日報（実績帳票）を作る。

入力  配車入力.xlsx     シート『配車入力』『車両マスタ』『荷主マスタ』
出力  配車日報_YYYYMMDD.xlsx  シート『配車日報』『車番別集計』『荷主別集計』

車番をキーにドライバー名を引くため、ETC高速料金の集計
（etc_allocate_by_vehicle.py）と同じ車番で月次の突き合わせができる。

    python scripts/haisha_nippo.py --template   入力テンプレートを作る
    python scripts/haisha_nippo.py              入力にある日付ぶんの日報を作る
    python scripts/haisha_nippo.py --date 2026-09-10
"""
import argparse, collections, datetime, os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IN   = os.path.join(ROOT, '配車入力.xlsx')

FONT = '游ゴシック'
YEN  = '"￥"#,##0'
WEEK = '月火水木金土日'
GRAY = Side(style='thin', color='BFBFBF')

# 『配車入力』シートの列順。テンプレート作成と読み込みで共用する。
COLS = ['日付', '荷主', '受付No', '車番', '積地', '積時刻', '降地', '降時刻',
        '品目', '数量', '単位', '運賃', '高速料金', '備考']

# 日報の列順と幅。数量・単位は運賃と分けて持ち、月次で合算できるようにする。
HEAD = ['No', '車番', 'ドライバー', '荷主', '受付No', '積地', '積時刻',
        '降地', '降時刻', '品目', '数量', '単位', '運賃', '高速料金', '備考']
WIDTH = (5, 7, 12, 14, 10, 16, 8, 16, 8, 14, 8, 6, 11, 11, 20)

# テンプレートの『記入例』シートに入れる見本。列順は COLS と同じ。
EXAMPLES = [
    ('2026/9/10', '日進運輸', 'A-101', 46, '宇都宮工場', '08:00',
     '川崎倉庫', '13:30', '鋼材', 8, 't', 48000, 3200, ''),
    ('2026/9/10', '日進運輸', 'A-102', 46, '川崎倉庫', '14:30',
     '宇都宮工場', '19:00', '空', None, '', 12000, 3200, '復便'),
    ('2026/9/10', '北関ロジ', 'H-21', None, '足利工場', '10:00',
     '三郷DC', '14:00', '紙製品', 20, 't', 61000, 0, '車両手配中'),
]


def to_date(v):
    """セルの値を date にする。文字列 'YYYY/M/D' や 'YYYY-MM-DD' も受ける。"""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str) and v.strip():
        s = v.strip().replace('-', '/')
        y, m, d = (int(x) for x in s.split('/'))
        return datetime.date(y, m, d)
    return None


def fmt_time(v):
    """セルの値を 'HH:MM' にする。空欄はそのまま空欄。"""
    if isinstance(v, (datetime.datetime, datetime.time)):
        return '%02d:%02d' % (v.hour, v.minute)
    return str(v).strip() if v not in (None, '') else None


def to_int(v):
    """金額・数量セルを数値にする。空欄・文字列は 0 扱いにしない（None を返す）。"""
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        s = v.strip().replace(',', '').replace('￥', '').replace('¥', '')
        if s.lstrip('-').isdigit():
            return int(s)
    return None


def read_rows(ws, ncol, start=2):
    """ヘッダ行を飛ばして値だけを返す。全列空欄の行は読み飛ばす。"""
    for row in ws.iter_rows(min_row=start, max_col=ncol, values_only=True):
        if any(v not in (None, '') for v in row):
            yield list(row) + [None] * (ncol - len(row))


def load(path):
    """入力ブックを読み、明細と各マスタを返す。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    if '配車入力' not in wb.sheetnames:
        raise SystemExit('『配車入力』シートがありません: %s' % path)

    cars = {}
    if '車両マスタ' in wb.sheetnames:
        for car, driver, kind, load_t in read_rows(wb['車両マスタ'], 4):
            if car is not None:
                cars[car] = dict(driver=driver or '', kind=kind or '', load=load_t)

    owners = {}
    if '荷主マスタ' in wb.sheetnames:
        for short, full in read_rows(wb['荷主マスタ'], 2):
            if short is not None:
                owners[short] = full or short

    recs = []
    for i, r in enumerate(read_rows(wb['配車入力'], len(COLS)), 2):
        d = to_date(r[0])
        if d is None:
            raise SystemExit('配車入力 %d行目: 日付が読めません（%r）' % (i, r[0]))
        recs.append(dict(
            date=d, owner=r[1] or '', no=r[2], car=r[3],
            from_=r[4] or '', from_t=fmt_time(r[5]),
            to_=r[6] or '', to_t=fmt_time(r[7]),
            item=r[8] or '', qty=to_int(r[9]), unit=r[10] or '',
            fare=to_int(r[11]) or 0, toll=to_int(r[12]) or 0,
            note=r[13] or '', row=i))
    return recs, cars, owners


def car_key(car):
    """車番の並び順。数値は数値順、それ以外（未配車など）は最後にまとめる。"""
    return (0, car, '') if isinstance(car, int) else (1, 0, str(car))


def style(cell, b=False, sz=10, align=None, fmt=None, fill=None):
    cell.font = Font(name=FONT, sz=sz, b=b)
    if align:
        cell.alignment = Alignment(horizontal=align, vertical='center')
    if fmt:
        cell.number_format = fmt
    if fill:
        cell.fill = PatternFill('solid', fgColor=fill)
    return cell


def write_header(ws, head, row=1, fill='DDEBF7'):
    for i, h in enumerate(head, 1):
        c = style(ws.cell(row, i, h), b=True, sz=11, align='center', fill=fill)
        c.border = Border(bottom=GRAY)


def build_nippo(wb, day, recs, cars):
    """1日分の配車日報シートを作る。未配車の便は末尾に色付きで出す。"""
    ws = wb.create_sheet('配車日報')
    ws['A1'] = '配　車　日　報'
    style(ws['A1'], b=True, sz=16, align='center')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEAD))

    ws['A2'] = '%d年%d月%d日（%s）' % (day.year, day.month, day.day, WEEK[day.weekday()])
    style(ws['A2'], b=True, sz=12)
    ws.cell(2, len(HEAD), '作成 %s' % datetime.date.today().strftime('%Y/%m/%d'))
    style(ws.cell(2, len(HEAD)), sz=9, align='right')

    write_header(ws, HEAD, row=4)

    # 車番順（未配車は最後）→ 積時刻順。同じ車番はひとかたまりで見せる。
    assigned = [r for r in recs if r['car'] not in (None, '')]
    unassigned = [r for r in recs if r['car'] in (None, '')]
    assigned.sort(key=lambda r: (car_key(r['car']), r['from_t'] or '', str(r['no'] or '')))

    row, n = 5, 0
    prev_car = object()
    for r in assigned + unassigned:
        n += 1
        car = r['car']
        info = cars.get(car, {})
        note = r['note']
        if car in (None, ''):
            note = '／'.join(x for x in ('※未配車', note) if x)
        elif car not in cars:
            note = '／'.join(x for x in ('※車両マスタ未登録', note) if x)

        vals = [n, car, info.get('driver', ''), r['owner'], r['no'],
                r['from_'], r['from_t'], r['to_'], r['to_t'],
                r['item'], r['qty'], r['unit'], r['fare'], r['toll'], note]
        for i, v in enumerate(vals, 1):
            c = style(ws.cell(row, i, v), sz=10)
            if i in (1, 2, 7, 9, 12):
                c.alignment = Alignment(horizontal='center')
            if car in (None, ''):
                c.fill = PatternFill('solid', fgColor='FCE4E4')
            # 車番が変わる行の上に区切り線を引く
            if car != prev_car and row > 5:
                c.border = Border(top=GRAY)
        style(ws.cell(row, 13), fmt=YEN)
        style(ws.cell(row, 14), fmt=YEN)
        prev_car = car
        row += 1

    style(ws.cell(row, 12, '合計'), b=True, sz=11, align='right')
    style(ws.cell(row, 13, sum(r['fare'] for r in recs)), b=True, sz=11, fmt=YEN)
    style(ws.cell(row, 14, sum(r['toll'] for r in recs)), b=True, sz=11, fmt=YEN)
    for i in range(1, len(HEAD) + 1):
        ws.cell(row, i).border = Border(top=GRAY, bottom=GRAY)

    if unassigned:
        row += 1
        style(ws.cell(row, 1, '※未配車が %d 件あります' % len(unassigned)),
              b=True, sz=11).font = Font(name=FONT, sz=11, b=True, color='C00000')

    for i, w in enumerate(WIDTH, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A5'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = '4:4'
    return ws


def build_summary(wb, recs, cars):
    """車番別・荷主別の集計シートを作る。"""
    ws = wb.create_sheet('車番別集計')
    write_header(ws, ['車番', 'ドライバー', '便数', '運賃', '高速料金', '差引'])
    by_car = collections.defaultdict(list)
    for r in recs:
        by_car[r['car'] if r['car'] not in (None, '') else '未配車'].append(r)
    row = 2
    for car in sorted(by_car, key=car_key):
        rs = by_car[car]
        fare, toll = sum(r['fare'] for r in rs), sum(r['toll'] for r in rs)
        vals = [car, cars.get(car, {}).get('driver', ''), len(rs), fare, toll, fare - toll]
        for i, v in enumerate(vals, 1):
            style(ws.cell(row, i, v), sz=10)
        for i in (4, 5, 6):
            ws.cell(row, i).number_format = YEN
        row += 1
    style(ws.cell(row, 2, '合計'), b=True, sz=11, align='right')
    style(ws.cell(row, 3, len(recs)), b=True, sz=11)
    style(ws.cell(row, 4, sum(r['fare'] for r in recs)), b=True, sz=11, fmt=YEN)
    style(ws.cell(row, 5, sum(r['toll'] for r in recs)), b=True, sz=11, fmt=YEN)
    style(ws.cell(row, 6, sum(r['fare'] - r['toll'] for r in recs)), b=True, sz=11, fmt=YEN)
    for col, w in zip('ABCDEF', (8, 14, 7, 12, 12, 12)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'

    ws = wb.create_sheet('荷主別集計')
    write_header(ws, ['荷主', '便数', '運賃'])
    by_owner = collections.Counter()
    cnt = collections.Counter()
    for r in recs:
        by_owner[r['owner']] += r['fare']
        cnt[r['owner']] += 1
    row = 2
    for owner, fare in sorted(by_owner.items(), key=lambda kv: -kv[1]):
        for i, v in enumerate([owner, cnt[owner], fare], 1):
            style(ws.cell(row, i, v), sz=10)
        ws.cell(row, 3).number_format = YEN
        row += 1
    style(ws.cell(row, 1, '合計'), b=True, sz=11, align='right')
    style(ws.cell(row, 2, len(recs)), b=True, sz=11)
    style(ws.cell(row, 3, sum(by_owner.values())), b=True, sz=11, fmt=YEN)
    for col, w in zip('ABC', (20, 7, 12)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'


def make_template(path):
    """空の入力ブックを作る。6社ぶんの荷主マスタの雛形付き。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '配車入力'
    write_header(ws, COLS)
    for i, w in enumerate((11, 14, 10, 7, 16, 8, 16, 8, 14, 8, 6, 11, 11, 20), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'

    ws = wb.create_sheet('車両マスタ')
    write_header(ws, ['車番', 'ドライバー', '車種', '最大積載(kg)'], fill='E2EFDA')
    for col, w in zip('ABCD', (8, 14, 14, 13)):
        ws.column_dimensions[col].width = w

    ws = wb.create_sheet('荷主マスタ')
    write_header(ws, ['荷主(略称)', '正式名称'], fill='E2EFDA')
    for col, w in zip('AB', (16, 28)):
        ws.column_dimensions[col].width = w

    # 記入例は『配車入力』に置くと日報に混ざるため、別シートに分ける。
    ws = wb.create_sheet('記入例')
    write_header(ws, COLS, fill='FFF2CC')
    for r in EXAMPLES:
        ws.append(list(r))
    for i, w in enumerate((11, 14, 10, 7, 16, 8, 16, 8, 14, 8, 6, 11, 11, 20), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.append([])
    ws.append(['※このシートは見本です。実際の入力は『配車入力』シートへ。'])
    ws.append(['※車番を空欄にすると日報で「未配車」として赤く出ます（配車漏れ確認用）。'])

    wb.save(path)
    print('入力テンプレートを作りました: %s' % path)


def main():
    ap = argparse.ArgumentParser(description='配車入力から配車日報を作る')
    ap.add_argument('--input', default=IN, help='入力ブック（既定: 配車入力.xlsx）')
    ap.add_argument('--date', help='対象日 YYYY-MM-DD（既定: 入力にある全日付）')
    ap.add_argument('--outdir', default=ROOT, help='出力先ディレクトリ')
    ap.add_argument('--template', action='store_true', help='入力テンプレートを作る')
    args = ap.parse_args()

    if args.template:
        make_template(args.input)
        return

    recs, cars, owners = load(args.input)
    if not recs:
        raise SystemExit('『配車入力』シートにデータがありません: %s' % args.input)
    days = sorted({r['date'] for r in recs})
    if args.date:
        want = to_date(args.date)
        if want not in days:
            raise SystemExit('%s のデータが入力にありません' % want)
        days = [want]

    for day in days:
        todays = [r for r in recs if r['date'] == day]
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        build_nippo(wb, day, todays, cars)
        build_summary(wb, todays, cars)
        out = os.path.join(args.outdir, '配車日報_%s.xlsx' % day.strftime('%Y%m%d'))
        wb.save(out)
        un = sum(1 for r in todays if r['car'] in (None, ''))
        print('%s  %d便  運賃￥%s  高速￥%s%s' % (
            day, len(todays), format(sum(r['fare'] for r in todays), ','),
            format(sum(r['toll'] for r in todays), ','),
            '  ※未配車 %d件' % un if un else ''))
        print('  -> %s' % out)


if __name__ == '__main__':
    main()
