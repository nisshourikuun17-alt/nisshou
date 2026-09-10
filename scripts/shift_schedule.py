#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ツーマン運行（2人1組）の月間出勤表を作成する。

  * 1組は必ず頭（運転の主担当）を1名含む。頭が余る日は頭が助手席に回る。
  * 号車は頭に固定（HEADS の並び順に ①②③…）。助手席の人はその日に乗る号車を書く。
  * 連続出勤は MAX_CONSECUTIVE 日まで（前月末からの連勤も CARRY_IN_STREAK で引き継ぐ）。
  * 希望休は必ず休みにする。
  * TARGET_AVG_WORK_DAYS を決めると、平日の組数を自動で調整して平均出勤日数を合わせる。
  * 出勤日数・土日出勤・組み合わせをメンバー間で平準化する。
"""

import calendar
import itertools
import random
from collections import defaultdict
from datetime import date

# ---------------------------------------------------------------- 設定 ----
YEAR = 2026
MONTH = 10

# 号車は現行の出勤表と同じ番号（船木①・和田②・永井③・山谷④・横山⑤）
HEADS = ['船木智一', '和田陽向太', '永井海里', '山谷大地', '横山凌']
ASSISTANTS = ['新田', '古川', '山口', '徳留']

# 出勤表に列だけ用意して、ローテーションには入れない人（空欄で出力）
EXTRA_COLUMNS = ['派遣']

TARGET_AVG_WORK_DAYS = 24   # 1人あたりの平均出勤日数。None なら CREWS_WEEKDAY を使う
CREWS_WEEKEND = 4           # 土日に出す組数（土日を多めに）
CREWS_WEEKDAY = 3           # TARGET_AVG_WORK_DAYS が None のときの平日の組数
MAX_CREWS = 5               # 車の台数
CREWS_OVERRIDE = {}         # 日にち -> 組数（祝日・繁忙日の個別指定）例: {12: 4}

# 希望休。氏名 -> 日にちのリスト。例: {'船木智一': [3, 20], '新田': [12]}
REQUESTED_OFF = {}

MAX_CONSECUTIVE = 6     # 連続出勤の上限（日）
CARRY_IN_STREAK = {}    # 氏名 -> 前月末時点で何連勤しているか。例: {'船木智一': 2}

SEED = 20261001         # 乱数種。変えると別パターンの出勤表が出る
TRIES = 300             # 生成の試行回数

OUT = f'{YEAR}年{MONTH}月_出勤表.xlsx'

WEEK_JA = ['月', '火', '水', '木', '金', '土', '日']
CIRCLED = '①②③④⑤⑥⑦⑧⑨⑩'
MEMBERS = HEADS + ASSISTANTS


# ------------------------------------------------------------ 日別の枠 ----


def build_days():
    """その月の各日について、曜日と出す組数を決める。"""
    last = calendar.monthrange(YEAR, MONTH)[1]
    days = []
    for d in range(1, last + 1):
        wd = date(YEAR, MONTH, d).weekday()
        days.append({'day': d, 'wd': wd, 'weekend': wd >= 5, 'target': 0})

    cap = min(MAX_CREWS, len(HEADS), len(MEMBERS) // 2)
    weekend = [s for s in days if s['weekend']]
    weekday = [s for s in days if not s['weekend']]
    for s in weekend:
        s['target'] = min(CREWS_WEEKEND, cap)

    if TARGET_AVG_WORK_DAYS is None:
        for s in weekday:
            s['target'] = min(CREWS_WEEKDAY, cap)
    else:
        # 延べ出勤人日 ÷ 2 が必要な組日数。土日分を引いた残りを平日に割り振る。
        total_crew_days = round(len(MEMBERS) * TARGET_AVG_WORK_DAYS / 2)
        rest = total_crew_days - sum(s['target'] for s in weekend)
        base, extra = divmod(max(rest, 0), len(weekday))
        base = min(base, cap)
        for s in weekday:
            s['target'] = base
        # 「base+1組」の日を月内に均等にばらけさせる
        if base < cap:
            for k in range(min(extra, len(weekday))):
                weekday[round(k * len(weekday) / extra)]['target'] = base + 1

    for s in days:
        if s['day'] in CREWS_OVERRIDE:
            s['target'] = min(CREWS_OVERRIDE[s['day']], cap)
    return days


def off_sets():
    return {m: set(REQUESTED_OFF.get(m, [])) for m in MEMBERS}


def cap_targets(days, offs):
    """希望休で人が足りない日は、その日の組数を実際に出せる数まで下げる。"""
    notes = []
    for spec in days:
        d = spec['day']
        avail = [m for m in MEMBERS if d not in offs[m]]
        avail_heads = sum(1 for m in avail if m in HEADS)
        capped = min(spec['target'], len(avail) // 2, avail_heads)
        if capped < spec['target']:
            notes.append(f'{MONTH}/{d}({WEEK_JA[spec["wd"]]}) 希望休のため {spec["target"]}組 → {capped}組')
            spec['target'] = capped
    return notes


# -------------------------------------------------------------- 割当て ----


def streak_ok(flags, carry):
    run = carry
    for worked in flags:
        run = run + 1 if worked else 0
        if run > MAX_CONSECUTIVE:
            return False
    return True


def max_run(flags, carry):
    run, best = carry, carry
    for w in flags:
        run = run + 1 if w else 0
        best = max(best, run)
    return best


def solve(days, offs, rng):
    """日ごとに出勤者（2×組数 名、うち頭が組数以上）を決める。"""
    n = len(days)
    avail = [[m for m in MEMBERS if days[i]['day'] not in offs[m]] for i in range(n)]
    need = [2 * s['target'] for s in days]
    # 出られる人数と必要人数が同じ日は、その人たちは出勤が確定する
    forced = [set(avail[i]) if need[i] >= len(avail[i]) else set() for i in range(n)]

    def forced_run(start, m):
        c, i = 0, start
        while i < n and m in forced[i]:
            c += 1
            i += 1
        return c

    work = {m: [False] * n for m in MEMBERS}
    worked = {m: 0 for m in MEMBERS}
    weekend_worked = {m: 0 for m in MEMBERS}
    streak = {m: CARRY_IN_STREAK.get(m, 0) for m in MEMBERS}
    shortages = []

    for i, spec in enumerate(days):
        crews = spec['target']
        cands = [m for m in avail[i]
                 if streak[m] + 1 + forced_run(i + 1, m) <= MAX_CONSECUTIVE]
        heads_avail = [m for m in cands if m in HEADS]
        if len(cands) < need[i] or len(heads_avail) < crews:
            crews = min(crews, len(cands) // 2, len(heads_avail))
            shortages.append((spec['day'], spec['target'], crews))
        want = 2 * crews

        if spec['weekend']:
            key = lambda m: (weekend_worked[m], worked[m], -streak[m], rng.random())
        else:
            key = lambda m: (worked[m], -streak[m], rng.random())
        order = sorted(cands, key=key)
        chosen = order[:want]
        # 頭が組数に満たなければ、出勤日数の多い非・頭と入れ替える
        while sum(1 for m in chosen if m in HEADS) < crews:
            out = max((m for m in chosen if m not in HEADS), key=key)
            inn = min((m for m in order if m in HEADS and m not in chosen), key=key)
            chosen[chosen.index(out)] = inn

        picked = set(chosen)
        for m in MEMBERS:
            if m in picked:
                work[m][i] = True
                worked[m] += 1
                if spec['weekend']:
                    weekend_worked[m] += 1
                streak[m] += 1
            else:
                streak[m] = 0

    weekend_idx = [i for i, s in enumerate(days) if s['weekend']]
    weekday_idx = [i for i, s in enumerate(days) if not s['weekend']]
    rebalance(days, work, offs, rng, weekend_idx)
    rebalance(days, work, offs, rng, weekday_idx)
    return work, shortages


def rebalance(days, work, offs, rng, idx_pool):
    """idx_pool の日に限って出勤日を交換し、その範囲の出勤日数を平準化する。"""
    pool = set(idx_pool)
    for _ in range(4000):
        counts = {m: sum(1 for i in pool if work[m][i]) for m in MEMBERS}
        hi = max(counts, key=lambda m: counts[m])
        lo = min(counts, key=lambda m: counts[m])
        if counts[hi] - counts[lo] <= 1:
            return
        idxs = [i for i in idx_pool
                if work[hi][i] and not work[lo][i] and days[i]['day'] not in offs[lo]]
        rng.shuffle(idxs)
        for i in idxs:
            heads_now = sum(1 for m in HEADS if work[m][i])
            after = heads_now - (hi in HEADS) + (lo in HEADS)
            if after < days[i]['target']:
                continue
            work[hi][i] = False
            work[lo][i] = True
            if (streak_ok(work[hi], CARRY_IN_STREAK.get(hi, 0))
                    and streak_ok(work[lo], CARRY_IN_STREAK.get(lo, 0))):
                break
            work[hi][i] = True
            work[lo][i] = False
        else:
            return


def spread(days, offs):
    """複数回作って、出勤日数・土日出勤のばらつきが最も小さいものを採用する。"""
    best = None
    for t in range(TRIES):
        rng = random.Random(SEED + t)
        work, shortages = solve(days, offs, rng)
        counts = [sum(work[m]) for m in MEMBERS]
        we = [sum(1 for i, s in enumerate(days) if s['weekend'] and work[m][i]) for m in MEMBERS]
        score = (len(shortages), max(we) - min(we), max(counts) - min(counts))
        if best is None or score < best[0]:
            best = (score, work, shortages)
        if score <= (0, 1, 1):
            break
    return best[1], best[2]


# ------------------------------------------------------------ ペア編成 ----


def make_pairs(days, work, rng):
    """毎日の号車ごとに「頭（運転）＋相方」を決める。頭が余る日は頭が助手席に回る。"""
    seen = defaultdict(int)      # 組み合わせ回数
    passenger = {h: 0 for h in HEADS}   # 頭が助手席に回った回数
    pairs_by_day = []
    for i, spec in enumerate(days):
        crews = spec['target']
        working = [m for m in MEMBERS if work[m][i]]
        heads_in = [m for m in HEADS if work[m][i]]
        # 助手席に回った回数が多い頭から順に運転（＝助手席役を輪番にする）
        leaders = sorted(heads_in, key=lambda h: (-passenger[h], rng.random()))[:crews]
        leaders = [h for h in HEADS if h in leaders]
        partners = [m for m in working if m not in leaders]
        for h in heads_in:
            if h not in leaders:
                passenger[h] += 1

        best, best_cost = None, None
        for perm in itertools.permutations(partners):
            cost = sum(seen[tuple(sorted((h, a)))] for h, a in zip(leaders, perm))
            cost += rng.random() * 0.01
            if best_cost is None or cost < best_cost:
                best, best_cost = perm, cost
        pairs = list(zip(leaders, best or ()))
        for h, a in pairs:
            seen[tuple(sorted((h, a)))] += 1
        pairs_by_day.append(pairs)
    return pairs_by_day, seen, passenger


def build_assignment(days, work, pairs_by_day):
    """各人・各日のセル内容（号車の丸数字 or 休）を決める。"""
    car_no = {h: i + 1 for i, h in enumerate(HEADS)}
    cells = {m: [''] * len(days) for m in MEMBERS}
    for i in range(len(days)):
        ride = {}
        for h, a in pairs_by_day[i]:
            ride[h] = car_no[h]
            ride[a] = car_no[h]
        for m in MEMBERS:
            if not work[m][i]:
                cells[m][i] = '休'
            else:
                n = ride.get(m)
                cells[m][i] = CIRCLED[n - 1] if n else '○'
    return cells


# ---------------------------------------------------------------- 出力 ----


def print_table(days, work, cells, pairs_by_day, shortages, seen, passenger):
    width = max(len(m) for m in MEMBERS) * 2 + 2
    hdr = '氏名'.ljust(width) + ' '.join(f'{s["day"]:>2}' for s in days) + '  出勤 土日'
    print(hdr)
    print('-' * len(hdr))
    print('曜日'.ljust(width) + ' '.join(f'{WEEK_JA[s["wd"]]:>2}' for s in days))
    print('車台数'.ljust(width - 2) + ' '.join(f'{len(pairs_by_day[i]):>2}' for i in range(len(days))))
    for label, group in (('【頭】', HEADS), ('【助手】', ASSISTANTS)):
        print(label)
        for m in group:
            line = ' '.join(f'{cells[m][i]:>2}' for i in range(len(days)))
            we = sum(1 for i, s in enumerate(days) if s['weekend'] and work[m][i])
            pad = width - len(m) * 2
            print(m + ' ' * max(pad, 1) + line + f'  {sum(work[m]):>3} {we:>3}')
    avg = sum(sum(work[m]) for m in MEMBERS) / len(MEMBERS)
    print(f'\n平均出勤日数: {avg:.1f}日 / {len(days)}日')
    print('連勤の最大: ' + ', '.join(
        f'{m}={max_run(work[m], CARRY_IN_STREAK.get(m, 0))}' for m in MEMBERS))
    print('頭が助手席に回った回数: ' + ', '.join(f'{h}={c}' for h, c in passenger.items()))
    if shortages:
        print('\n人数が足りず組数を減らした日:')
        for d, want, got in sorted(set(shortages)):
            print(f'  {MONTH}/{d}  {want}組 → {got}組')
    print('\n組み合わせ回数: ' + ', '.join(f'{a}×{b}={c}' for (a, b), c in sorted(seen.items())))


def write_xlsx(days, work, offs, cells, pairs_by_day, notes, shortages, passenger):
    """手書きの出勤表と同じ様式（縦＝日付、横＝氏名、セル＝号車 or 休）で書き出す。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = '出勤表'

    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center')
    sat_fill = PatternFill('solid', fgColor='DDEBF7')
    sun_fill = PatternFill('solid', fgColor='FCE4E4')
    off_fill = PatternFill('solid', fgColor='F2F2F2')
    req_fill = PatternFill('solid', fgColor='FFF2CC')
    name_fill = PatternFill('solid', fgColor='EFEFEF')
    big = Font(bold=True, size=14)

    columns = MEMBERS + EXTRA_COLUMNS
    ncols = 3 + len(columns)

    ws.cell(1, 3, f'{YEAR}年').font = big
    ws.cell(1, 3).alignment = center
    ws.cell(1, 3 + max(1, len(columns) // 3), '出勤表').font = big
    ws.cell(1, 3 + max(2, len(columns) * 2 // 3), f'{MONTH}月度').font = big

    HEAD_ROW = 2
    for j, title in enumerate(('', '', '車台数')):
        c = ws.cell(HEAD_ROW, 1 + j, title)
        c.font = Font(bold=True)
        c.alignment = center
        c.border = border
    for j, name in enumerate(columns):
        c = ws.cell(HEAD_ROW, 4 + j, name)
        c.font = Font(bold=True)
        c.alignment = center
        c.border = border
        c.fill = name_fill

    for i, spec in enumerate(days):
        r = HEAD_ROW + 1 + i
        c = ws.cell(r, 1, spec['day'])
        c.alignment = center
        c.border = border
        c = ws.cell(r, 2, WEEK_JA[spec['wd']])
        c.alignment = center
        c.border = border
        if spec['wd'] == 5:
            c.font = Font(color='0070C0', bold=True)
            c.fill = sat_fill
        elif spec['wd'] == 6:
            c.font = Font(color='C00000', bold=True)
            c.fill = sun_fill
        c = ws.cell(r, 3, len(pairs_by_day[i]))
        c.alignment = center
        c.border = border
        for j, name in enumerate(columns):
            cell = ws.cell(r, 4 + j)
            cell.alignment = center
            cell.border = border
            if name not in cells:
                continue
            cell.value = cells[name][i]
            requested = spec['day'] in offs[name]
            if cell.value == '休':
                cell.fill = req_fill if requested else off_fill
                if requested:
                    cell.font = Font(bold=True)
            elif spec['wd'] == 5:
                cell.fill = sat_fill
            elif spec['wd'] == 6:
                cell.fill = sun_fill

    total_row = HEAD_ROW + 1 + len(days)
    ws.cell(total_row, 1, '合計').font = Font(bold=True)
    ws.cell(total_row + 1, 1, '残業').font = Font(bold=True)
    for r in (total_row, total_row + 1):
        for col in range(1, ncols + 1):
            ws.cell(r, col).border = border
            ws.cell(r, col).alignment = center
    ws.cell(total_row, 3, sum(len(p) for p in pairs_by_day))
    for j, name in enumerate(columns):
        if name in cells:
            ws.cell(total_row, 4 + j, sum(work[name]))

    ws.freeze_panes = 'D3'
    ws.column_dimensions['A'].width = 4.5
    ws.column_dimensions['B'].width = 4.5
    ws.column_dimensions['C'].width = 7
    for j in range(len(columns)):
        ws.column_dimensions[ws.cell(1, 4 + j).column_letter].width = 11

    # --- 組み合わせシート ---
    ws2 = wb.create_sheet('組み合わせ')
    ws2.cell(1, 1, '日').font = Font(bold=True)
    ws2.cell(1, 2, '曜').font = Font(bold=True)
    ws2.cell(1, 3, '車台数').font = Font(bold=True)
    for k in range(len(HEADS)):
        ws2.cell(1, 4 + k, f'{CIRCLED[k]}{HEADS[k]}').font = Font(bold=True)
    for i, spec in enumerate(days):
        r = 2 + i
        ws2.cell(r, 1, f'{MONTH}/{spec["day"]}').border = border
        c = ws2.cell(r, 2, WEEK_JA[spec['wd']])
        c.border = border
        c.alignment = center
        if spec['wd'] == 5:
            c.fill = sat_fill
        elif spec['wd'] == 6:
            c.fill = sun_fill
        c = ws2.cell(r, 3, len(pairs_by_day[i]))
        c.border = border
        c.alignment = center
        by_car = {HEADS.index(h): a for h, a in pairs_by_day[i]}
        for k in range(len(HEADS)):
            cell = ws2.cell(r, 4 + k, by_car.get(k, '運休'))
            cell.border = border
            cell.alignment = center
            if k not in by_car:
                cell.fill = off_fill
    ws2.column_dimensions['A'].width = 8
    ws2.column_dimensions['B'].width = 5
    ws2.column_dimensions['C'].width = 7
    for k in range(len(HEADS)):
        ws2.column_dimensions[ws2.cell(1, 4 + k).column_letter].width = 16

    # --- 備考シート ---
    ws3 = wb.create_sheet('備考')
    r = 1
    ws3.cell(r, 1, '設定').font = Font(bold=True)
    r += 1
    for k, v in (('対象月', f'{YEAR}年{MONTH}月'),
                 ('目標の平均出勤日数', f'{TARGET_AVG_WORK_DAYS}日' if TARGET_AVG_WORK_DAYS else '指定なし'),
                 ('土日の組数', CREWS_WEEKEND),
                 ('連続出勤の上限', f'{MAX_CONSECUTIVE}日'),
                 ('頭（号車固定）', '、'.join(f'{CIRCLED[i]}{h}' for i, h in enumerate(HEADS))),
                 ('助手', '、'.join(ASSISTANTS))):
        ws3.cell(r, 1, k)
        ws3.cell(r, 2, v)
        r += 1
    r += 1
    ws3.cell(r, 1, '希望休').font = Font(bold=True)
    r += 1
    for m in MEMBERS:
        ds = REQUESTED_OFF.get(m)
        if ds:
            ws3.cell(r, 1, m)
            ws3.cell(r, 2, '、'.join(f'{MONTH}/{d}' for d in ds))
            r += 1
    r += 1
    ws3.cell(r, 1, '1人ごとの内訳').font = Font(bold=True)
    r += 1
    for m in MEMBERS:
        we = sum(1 for i, s in enumerate(days) if s['weekend'] and work[m][i])
        extra = f'／助手席{passenger[m]}日' if m in HEADS and passenger.get(m) else ''
        ws3.cell(r, 1, m)
        ws3.cell(r, 2, f'出勤{sum(work[m])}日／土日{we}日／公休{len(days) - sum(work[m])}日／'
                       f'最大{max_run(work[m], CARRY_IN_STREAK.get(m, 0))}連勤{extra}')
        r += 1
    if notes or shortages:
        r += 1
        ws3.cell(r, 1, '注意').font = Font(bold=True)
        r += 1
        for note in notes:
            ws3.cell(r, 1, note)
            r += 1
        for d, want, got in sorted(set(shortages)):
            ws3.cell(r, 1, f'{MONTH}/{d} 人数不足のため {want}組 → {got}組')
            r += 1
    ws3.column_dimensions['A'].width = 46
    ws3.column_dimensions['B'].width = 62

    wb.save(OUT)
    return OUT


def verify(days, work, offs, pairs_by_day):
    errors = []
    for m in MEMBERS:
        for i, spec in enumerate(days):
            if work[m][i] and spec['day'] in offs[m]:
                errors.append(f'{m}: 希望休 {MONTH}/{spec["day"]} に出勤している')
        if max_run(work[m], CARRY_IN_STREAK.get(m, 0)) > MAX_CONSECUTIVE:
            errors.append(f'{m}: 連続出勤が{MAX_CONSECUTIVE}日を超えている')
    for i, spec in enumerate(days):
        pairs = pairs_by_day[i]
        riders = [m for p in pairs for m in p]
        if len(riders) != len(set(riders)):
            errors.append(f'{MONTH}/{spec["day"]}: 同じ人が2台に乗っている')
        if sum(work[m][i] for m in MEMBERS) != 2 * len(pairs):
            errors.append(f'{MONTH}/{spec["day"]}: 出勤者と乗車人数が合わない')
        for h, a in pairs:
            if h not in HEADS:
                errors.append(f'{MONTH}/{spec["day"]}: 頭のいない組がある')
    return errors


def main():
    rng = random.Random(SEED)
    days = build_days()
    offs = off_sets()
    notes = cap_targets(days, offs)

    work, shortages = spread(days, offs)
    pairs_by_day, seen, passenger = make_pairs(days, work, rng)
    cells = build_assignment(days, work, pairs_by_day)

    print_table(days, work, cells, pairs_by_day, shortages, seen, passenger)
    for note in notes:
        print('注意:', note)

    errors = verify(days, work, offs, pairs_by_day)
    if errors:
        print('\n*** 制約違反 ***')
        for e in errors:
            print(' ', e)
    else:
        print('\n制約チェック: 希望休・連勤上限・組編成すべてOK')

    print('出力:', write_xlsx(days, work, offs, cells, pairs_by_day, notes, shortages, passenger))


if __name__ == '__main__':
    main()
