# -*- coding: utf-8 -*-
"""配車入力から配車日報（実績帳票）を作る。

入力  配車入力.xlsx     シート『配車入力』『乗務員マスタ』『得意先マスタ』
出力  配車日報_YYYYMMDD.xlsx  シート『配車日報』『乗務員別集計』『得意先別集計』

日報は実際の用紙と同じ左右2ブロック構成にする。
  左ブロック  … 1便目
  右ブロック  … 2便目以降（2便目が1行目、3便目が2行目）
乗務員の並び順は『乗務員マスタ』の表示順に従う。

車番は便ごとに持つ（同じ乗務員でも便で車を替えるため）。ETC高速料金の
車番別集計（etc_allocate_by_vehicle.py）とは車番で突き合わせできる。

    python scripts/haisha_nippo.py --template   入力テンプレートを作る
    python scripts/haisha_nippo.py              入力にある日付ぶんの日報を作る
    python scripts/haisha_nippo.py --date 2026-09-10 --tanto 藁科
"""
import argparse, collections, datetime, os, re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IN   = os.path.join(ROOT, '配車入力.xlsx')

FONT = '游ゴシック'
WEEK = '月火水木金土日'
BOX  = Border(*[Side(style='thin', color='000000')] * 4)

# 『配車入力』シートの列順。テンプレート作成と読み込みで共用する。
# B〜F列（車番・積荷・得意先・発地・着地）だけ打てば済むよう、この5つを隣に並べる。
# 日付は上の行から引き継ぎ、乗務員は車番マスタから、便は入力順から自動で決まる。
COLS = ['日付', '車番', '積荷', '得意先', '発地', '着地', '乗務員', '便', '備考']
TYPE_COLS = (2, 6)          # 打ち込む列の範囲（車番〜着地）
CWIDTH = (11, 8, 13, 14, 12, 14, 11, 5, 20)

# 日報1ブロックぶんの列と幅。これを左右2つ並べる。
BLOCK = ['乗務員', '車番', '積荷', '得意先', '発地', '着地', '備考']
BWIDTH = (10, 7, 12, 13, 11, 13, 6)

# テンプレートの『記入例』に入れる見本。列順は COLS と同じ。
EXAMPLES = [
    ('2026/9/10', 5022, '玉ねぎ', '鈴与', '日立', '豊洲', '', '', ''),
    ('', 5022, '食品', '鈴与', '桶川市', '常陸那珂', '', '', ''),
    ('', 1956, 'チーズ', '鈴与', '日立', '阿見', '', '', ''),
    ('', 1956, '代用乳', '鈴与', '高崎市', '日立', '', '', ''),
    ('', 1956, '紙製品', '鈴与', 'いわき市', '常陸那珂', '', '', ''),
    ('', 8816, 'トマト', '豊総合物流', '大洗', '', '', '', '着地未定'),
    ('', 2403, '雑貨', 'ロードリーム', '相模原市', '常陸那珂', '仲林', '', '代車のため乗務員を明記'),
    ('', '', '玉葱', '大晴通商', '日立', '豊洲', '', '', '車番未定→未配車に出る'),
]

# テンプレートに入れておくマスタの初期値。実際の日報から起こしたもの。
DRIVERS = ['関', '鶴田', '伊垣', '横須賀', '仲林', '海老澤', '芳賀', '砂押',
           '小林', '志賀', '宇佐美', '萩野間', '木村', '高橋', '山崎', '根本',
           '安野', '関根', '神長', '小堀内', '吉成', '山崎正二', '石川']

# 車番 -> 乗務員。車番を打てば乗務員が決まるので、入力は車番だけで済む。
# 代車で普段と違う人が乗る日は『配車入力』の乗務員欄に直接書けば、そちらが優先。
VEHICLES = [
    (5022, '関'), (2039, '鶴田'), (8816, '伊垣'), (2185, '横須賀'),
    (1464, '仲林'), (2403, '仲林'), (8815, '海老澤'), (3655, '海老澤'),
    (1956, '芳賀'), (1107, '砂押'), (8007, '小林'), (2145, '志賀'),
    (1957, '宇佐美'), (5135, '宇佐美'), (1925, '萩野間'), (9480, '木村'),
    (1983, '木村'), (5164, '高橋'), (1523, '山崎'), (5159, '根本'),
    (3134, '安野'), (1000, '関根'), (3069, '神長'), (6891, '小堀内'),
]
CUSTOMERS = ['鈴与', 'エアウォーター', '豊総合物流', 'ロードリーム',
             '大晴通商', '光洋運輸', '行方運送', '東亜物産', 'OOCL']


def to_date(v):
    """セルの値を date にする。文字列 'YYYY/M/D' や 'YYYY-MM-DD' も受ける。"""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, str) and v.strip():
        y, m, d = (int(x) for x in v.strip().replace('-', '/').split('/'))
        return datetime.date(y, m, d)
    return None


def wareki(d):
    """date -> '令和 8年  9月 10日 （木）'。令和1年 = 2019年。"""
    return '令和 %d年　%d月　%d日　（%s）' % (
        d.year - 2018, d.month, d.day, WEEK[d.weekday()])


def blank(v):
    return v in (None, '')


def read_rows(ws, ncol, start=2):
    """ヘッダ行を飛ばして値だけを返す。全列空欄の行は読み飛ばす。"""
    for row in ws.iter_rows(min_row=start, max_col=ncol, values_only=True):
        if any(not blank(v) for v in row):
            yield list(row) + [None] * (ncol - len(row))


def master(wb, name, ncol):
    """マスタシートを読む。『※』で始まる注意書きの行は読み飛ばす。"""
    if name not in wb.sheetnames:
        return []
    return [r for r in read_rows(wb[name], ncol)
            if not blank(r[0]) and not str(r[0]).startswith('※')]


def load(path, default_date=None):
    """入力ブックを読み、明細・乗務員の並び順・車番マスタ・得意先マスタを返す。

    打ち込みを減らすため、空欄は次のように補う。明示的に書けばそちらが優先。
      日付   … 直前の行から引き継ぐ（1日分なら先頭行に1回書けばよい）
      乗務員 … 車番マスタから引く（代車の日は乗務員欄に直接書く）
      便     … 同じ乗務員の中で出てきた順に1便目・2便目…とする
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    if '配車入力' not in wb.sheetnames:
        raise SystemExit('『配車入力』シートがありません: %s' % path)

    order = [r[0] for r in master(wb, '乗務員マスタ', 2)]
    vehicles = {str(r[0]).strip(): r[1] or '' for r in master(wb, '車番マスタ', 2)}
    customers = [r[0] for r in master(wb, '得意先マスタ', 1)]

    recs, last = [], default_date
    for i, r in enumerate(read_rows(wb['配車入力'], len(COLS)), 2):
        d = to_date(r[0]) or last
        if d is None:
            raise SystemExit('配車入力 %d行目: 日付が決まりません。'
                             '先頭行の日付欄を埋めるか --date を指定してください。' % i)
        last = d
        car = r[1]
        driver = r[6] if not blank(r[6]) else vehicles.get(str(car).strip(), '')
        recs.append(dict(
            date=d, car=car, item=r[2] or '', cust=r[3] or '',
            from_=r[4] or '', to_=r[5] or '', driver=driver or '',
            trip=r[7], note=r[8] or '', row=i))
    return recs, order, vehicles, customers


def read_text(path, vehicles, default_date=None):
    """読み取りテキストから明細を作る。写真をAIに読ませた結果を貼る想定。

        2026/9/10                       ← 日付行。以降この日付が続く
        5022  玉ねぎ  鈴与  日立  豊洲     ← 車番 積荷 得意先 発地 着地
        40ft  9:30  OOCL  大井  水戸市  吉成  6455   ← 6列目=乗務員 7列目=備考

    区切りはタブでも2文字以上の空白でもよい。空行と # 以降は無視する。
    """
    recs, day = [], default_date
    for i, raw in enumerate(open(path, encoding='utf-8'), 1):
        # 末尾の空欄（着地なしなど）を消さないよう、改行だけ落として分割する
        line = raw.split('#')[0].rstrip('\r\n')
        if not line.strip():
            continue
        cols = [c.strip() for c in (line.split('\t') if '\t' in line
                                    else re.split(r'\s{2,}', line.strip()))]
        if len(cols) == 1:
            d = to_date(cols[0])
            if d is None:
                raise SystemExit('%s %d行目: 日付として読めません（%r）' % (path, i, line))
            day = d
            continue
        if day is None:
            raise SystemExit('%s %d行目: 先に日付の行を書いてください。' % (path, i))
        if len(cols) < 5:
            raise SystemExit('%s %d行目: 車番 積荷 得意先 発地 着地 の5つが要ります（%r）'
                             % (path, i, line))
        cols = cols[:7] + [''] * max(0, 7 - len(cols))
        car = int(cols[0]) if cols[0].isdigit() else cols[0]
        recs.append(dict(
            date=day, car=car, item=cols[1], cust=cols[2], from_=cols[3],
            to_=cols[4], driver=cols[5] or vehicles.get(cols[0], ''),
            trip=None, note=cols[6], row=i))
    return recs


def group(recs, order):
    """乗務員ごとに便をまとめ、マスタの表示順に並べる。

    乗務員が空欄の便は『未配車』としてまとめ、必ず最後に置く。
    マスタに無い乗務員はマスタ掲載分の後ろに、入力に出てきた順で続ける。
    """
    trips = collections.defaultdict(list)
    seen = []
    for r in recs:
        name = r['driver'] if not blank(r['driver']) else None
        if name not in trips:
            seen.append(name)
        trips[name].append(r)

    for v in trips.values():
        # 便No が空欄でも入力順を保てるよう、行番号を第2キーにする
        v.sort(key=lambda r: (r['trip'] if isinstance(r['trip'], int) else 99, r['row']))

    ranked = [n for n in order if n in trips]
    ranked += [n for n in seen if n is not None and n not in order]
    if None in trips:
        ranked.append(None)
    return [(n, trips[n]) for n in ranked]


def style(cell, b=False, sz=10, align=None, fill=None, color=None):
    cell.font = Font(name=FONT, sz=sz, b=b, color=color)
    if align:
        cell.alignment = Alignment(horizontal=align, vertical='center')
    if fill:
        cell.fill = PatternFill('solid', fgColor=fill)
    return cell


def build_nippo(wb, day, recs, order, tanto, spacer):
    """実際の用紙と同じ左右2ブロックの日報シートを作る。"""
    ws = wb.create_sheet('配車日報')
    ncol = len(BLOCK) * 2

    ws['A1'] = '配　車　日　報'
    style(ws['A1'], b=True, sz=18, align='center')
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    ws.row_dimensions[1].height = 28

    ws['A2'] = wareki(day)
    style(ws['A2'], b=True, sz=14)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(BLOCK))
    if tanto:
        c = ws.cell(2, len(BLOCK) + 1, '担当　%s　様' % tanto)
        style(c, sz=11, align='right')
        ws.merge_cells(start_row=2, start_column=len(BLOCK) + 1,
                       end_row=2, end_column=ncol)
    ws.row_dimensions[2].height = 22

    for i, h in enumerate(BLOCK * 2, 1):
        c = style(ws.cell(3, i, h), b=True, sz=10, align='center', fill='F2F2F2')
        c.border = BOX
    ws.row_dimensions[3].height = 20

    row = 4
    for name, trips in group(recs, order):
        unassigned = name is None
        head, rest = trips[0], trips[1:]
        # ブロックの高さは「右に並べる2便目以降の数」で決まる。1便だけなら1行。
        height = max(1, len(rest))
        for k in range(height):
            left = head if k == 0 else None
            right = rest[k] if k < len(rest) else None
            for side, rec in ((0, left), (1, right)):
                base = side * len(BLOCK)
                vals = ['' if rec is None else v for v in (
                    name or '未配車', rec and rec['car'], rec and rec['item'],
                    rec and rec['cust'], rec and rec['from_'],
                    rec and rec['to_'], rec and rec['note'])]
                if rec is None:
                    vals = [''] * len(BLOCK)
                for i, v in enumerate(vals, base + 1):
                    c = style(ws.cell(row, i, v if v != '' else None), sz=10)
                    c.border = BOX
                    if i - base in (1, 2, 7):
                        c.alignment = Alignment(horizontal='center', vertical='center')
                    if unassigned and rec is not None:
                        c.fill = PatternFill('solid', fgColor='FCE4E4')
            ws.row_dimensions[row].height = 20
            row += 1
        if spacer:
            for i in range(1, ncol + 1):
                ws.cell(row, i).border = BOX
            ws.row_dimensions[row].height = 14
            row += 1

    n_un = sum(1 for r in recs if blank(r['driver']))
    if n_un:
        style(ws.cell(row, 1, '※未配車が %d 便あります' % n_un),
              b=True, sz=11, color='C00000')

    for i, w in enumerate(BWIDTH * 2, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A4'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1  # 1日分が必ず1ページに収まるようにする
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_title_rows = '3:3'
    return ws


def build_summary(wb, recs, order, cars):
    """乗務員別・得意先別の便数集計を作る。"""
    ws = wb.create_sheet('乗務員別集計')
    for i, h in enumerate(['乗務員', '車番', '便数', '積荷', '得意先'], 1):
        style(ws.cell(1, i, h), b=True, sz=11, align='center', fill='DDEBF7').border = BOX
    row = 2
    for name, trips in group(recs, order):
        used = sorted({str(t['car']) for t in trips if not blank(t['car'])})
        vals = [name or '未配車', '・'.join(used), len(trips),
                '／'.join(t['item'] for t in trips if t['item']),
                '／'.join(sorted({t['cust'] for t in trips if t['cust']}))]
        for i, v in enumerate(vals, 1):
            style(ws.cell(row, i, v), sz=10)
        row += 1
    style(ws.cell(row, 1, '合計'), b=True, sz=11, align='right')
    style(ws.cell(row, 3, len(recs)), b=True, sz=11)
    for col, w in zip('ABCDE', (12, 14, 7, 26, 26)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'

    ws = wb.create_sheet('得意先別集計')
    for i, h in enumerate(['得意先', '便数', '乗務員'], 1):
        style(ws.cell(1, i, h), b=True, sz=11, align='center', fill='DDEBF7').border = BOX
    cnt = collections.Counter(r['cust'] or '（未記入）' for r in recs)
    who = collections.defaultdict(list)
    for r in recs:
        d = r['driver'] or '未配車'
        if d not in who[r['cust'] or '（未記入）']:
            who[r['cust'] or '（未記入）'].append(d)
    row = 2
    for cust, n in cnt.most_common():
        for i, v in enumerate([cust, n, '・'.join(who[cust])], 1):
            style(ws.cell(row, i, v), sz=10)
        row += 1
    style(ws.cell(row, 1, '合計'), b=True, sz=11, align='right')
    style(ws.cell(row, 2, len(recs)), b=True, sz=11)
    for col, w in zip('ABC', (18, 7, 40)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A2'


def dropdown(ws, src_sheet, src_range, target):
    """マスタを参照するプルダウンを付ける。打ち間違いと打鍵数を減らす。"""
    # formula1 に '=' を付けると Excel が破損ファイルと判定するので付けない
    dv = DataValidation(type='list', allow_blank=True,
                        formula1="'%s'!%s" % (src_sheet, src_range))
    ws.add_data_validation(dv)
    dv.add(target)


def make_template(path):
    """空の入力ブックを作る。マスタは実際の日報から起こした初期値入り。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '配車入力'
    lo, hi = TYPE_COLS
    for i, h in enumerate(COLS, 1):
        # 打ち込む5列だけ濃い色にして、触る場所をひと目で分かるようにする
        fill = 'FFD966' if lo <= i <= hi else 'DDEBF7'
        style(ws.cell(1, i, h), b=True, sz=11, align='center', fill=fill).border = BOX
    for i, w in enumerate(CWIDTH, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    dropdown(ws, '車番マスタ', '$A$2:$A$200', 'B2:B500')
    dropdown(ws, '得意先マスタ', '$A$2:$A$100', 'D2:D500')

    ws = wb.create_sheet('乗務員マスタ')
    for i, h in enumerate(['乗務員', '備考'], 1):
        style(ws.cell(1, i, h), b=True, sz=11, align='center', fill='E2EFDA').border = BOX
    for name in DRIVERS:
        ws.append([name, None])
    for col, w in zip('AB', (12, 28)):
        ws.column_dimensions[col].width = w
    ws.append([])
    ws.append(['※この並び順が日報の行順になります。'])

    ws = wb.create_sheet('車番マスタ')
    for i, h in enumerate(['車番', '乗務員'], 1):
        style(ws.cell(1, i, h), b=True, sz=11, align='center', fill='E2EFDA').border = BOX
    for car, name in VEHICLES:
        ws.append([car, name])
    for col, w in zip('AB', (10, 12)):
        ws.column_dimensions[col].width = w
    ws.append([])
    ws.append(['※車番を打つと乗務員が自動で決まります。'])
    ws.append(['※代車で普段と違う人が乗る日は、配車入力の乗務員欄に直接書いてください。'])

    ws = wb.create_sheet('得意先マスタ')
    style(ws.cell(1, 1, '得意先'), b=True, sz=11, align='center', fill='E2EFDA').border = BOX
    for c in CUSTOMERS:
        ws.append([c])
    ws.column_dimensions['A'].width = 18

    # 記入例は『配車入力』に置くと日報に混ざるため、別シートに分ける。
    ws = wb.create_sheet('記入例')
    for i, h in enumerate(COLS, 1):
        fill = 'FFD966' if lo <= i <= hi else 'FFF2CC'
        style(ws.cell(1, i, h), b=True, sz=11, align='center', fill=fill).border = BOX
    for r in EXAMPLES:
        ws.append([v if v != '' else None for v in r])
    for i, w in enumerate(CWIDTH, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for msg in ('', '※打ち込むのは黄色の5列（車番・積荷・得意先・発地・着地）だけです。',
                '※日付は1日の先頭行に1回書けば、下の行は同じ日付になります。',
                '※乗務員は車番マスタから自動で入ります。代車の日だけ手で書いてください。',
                '※便は同じ乗務員の中で上から順に1便目・2便目…になります。',
                '※車番を空欄にすると日報の最後に「未配車」として赤く出ます。'):
        ws.append([msg])

    wb.save(path)
    print('入力テンプレートを作りました: %s' % path)


def append_input(path, recs):
    """読み取った明細を『配車入力』シートの末尾に追記する。"""
    wb = openpyxl.load_workbook(path)
    ws = wb['配車入力']
    last = None
    for r in recs:
        ws.append([r['date'] if r['date'] != last else None, r['car'], r['item'],
                   r['cust'], r['from_'], r['to_'], r['driver'], None, r['note']])
        ws.cell(ws.max_row, 1).number_format = 'yyyy/m/d'
        last = r['date']
    wb.save(path)
    print('配車入力に %d 便を追記しました: %s' % (len(recs), path))


def main():
    ap = argparse.ArgumentParser(description='配車入力から配車日報を作る')
    ap.add_argument('--input', default=IN, help='入力ブック（既定: 配車入力.xlsx）')
    ap.add_argument('--date', help='対象日 YYYY-MM-DD（既定: 入力にある全日付）')
    ap.add_argument('--outdir', default=ROOT, help='出力先ディレクトリ')
    ap.add_argument('--tanto', default='', help='担当者名（日報の右上に入る）')
    ap.add_argument('--spacer', action='store_true', help='乗務員ごとに空行を入れる')
    ap.add_argument('--text', help='読み取りテキストから日報を作る（配車入力は使わない）')
    ap.add_argument('--save-input', action='store_true',
                    help='--text の内容を配車入力シートにも追記する')
    ap.add_argument('--template', action='store_true', help='入力テンプレートを作る')
    args = ap.parse_args()

    if args.template:
        make_template(args.input)
        return

    want = to_date(args.date) if args.date else None
    recs, order, cars, customers = load(args.input, want)
    if args.text:
        # テキストを渡された日は、そちらを明細として使う（マスタはブックから）
        recs = read_text(args.text, cars, want)
        if args.save_input:
            append_input(args.input, recs)
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
        build_nippo(wb, day, todays, order, args.tanto, args.spacer)
        build_summary(wb, todays, order, cars)
        out = os.path.join(args.outdir, '配車日報_%s.xlsx' % day.strftime('%Y%m%d'))
        wb.save(out)
        un = sum(1 for r in todays if blank(r['driver']))
        drivers = len({r['driver'] for r in todays if not blank(r['driver'])})
        print('%s  %d便 / 乗務員%d名%s' % (
            day, len(todays), drivers, '  ※未配車 %d便' % un if un else ''))
        print('  -> %s' % out)


if __name__ == '__main__':
    main()
