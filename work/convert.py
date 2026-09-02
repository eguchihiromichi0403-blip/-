# -*- coding: utf-8 -*-
"""令和08年シート → 工事台帳インポートCSV 変換"""
import openpyxl, datetime, csv, sys, unicodedata, re

SRC = 'work/source.xlsx'
SHEET = '令和08年 '
TODAY = datetime.date(2026, 9, 2)
EPOCH = datetime.date(1899, 12, 30)

HEADER = ['工事番号','現場名','得意先名','得意先コード','工期開始','工期終了','請負金額',
          '前期末未成工事支出金（材料費）','前期末未成工事支出金（労務費）',
          '前期末未成工事支出金（外注費）','前期末未成工事支出金（経費）','状態','既定の原価区分']

# 元請け名称の表記ゆれ統一（明らかな大小文字・全半角違いのみ）
CUSTOMER_ALIAS = {'k2': 'K2', 'ｋ２': 'K2'}

def norm_site(v):
    """現場名は原文を尊重し、前後の空白と連続空白の整理のみ"""
    if v is None: return ''
    return re.sub(r'[\s\u3000]+', ' ', str(v)).strip()

def norm_name(v):
    if v is None: return ''
    s = unicodedata.normalize('NFKC', str(v)).strip()
    s = re.sub(r'[\s　]+', ' ', s)
    return CUSTOMER_ALIAS.get(s, CUSTOMER_ALIAS.get(s.lower(), s))

def to_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    if isinstance(v, (int, float)) and 30000 < v < 60000:
        return EPOCH + datetime.timedelta(days=int(v))
    return None

def to_int(v):
    if isinstance(v, bool): return None
    if isinstance(v, (int, float)): return int(round(v))
    return None

def duration_days(v):
    if v is None: return None
    s = unicodedata.normalize('NFKC', str(v))
    m = re.search(r'(\d+)', s)
    return int(m.group(1)) if m else None

def add_year(dt):
    try: return dt.replace(year=dt.year + 1)
    except ValueError: return dt.replace(year=dt.year + 1, day=28)

def ym(dt):
    return f'{dt.year}/{dt.month:02d}' if dt else ''

def convert(seed_costs: bool):
    wb = openpyxl.load_workbook(SRC, data_only=True)
    ws = wb[SHEET]
    out, notes = [], []
    seq = 0
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or not any(c is not None and str(c).strip() for c in r):
            continue
        site = norm_site(r[1])
        if not site:
            continue
        seq += 1
        no = f'2026-{seq:03d}'
        cust = norm_name(r[0])

        start, end = to_date(r[2]), to_date(r[4])
        if r[4] is not None and end is None:
            notes.append(f'{no} {site}: 完了日「{r[4]}」を解釈できないため、着工日＋工事期間から補完')
        if start and end and end < start:
            if (start - end).days > 300:           # 年の入力ミス（1年前になっている）
                fixed = add_year(end)
            else:                                  # 日付の打ち間違い → 着工日＋工事期間で補完
                dur = duration_days(r[3])
                fixed = start + datetime.timedelta(days=dur) if dur is not None else start
            notes.append(f'{no} {site}: 完了日 {end} が着工日より前 → {fixed} に補正')
            end = fixed
        if start and end is None:                  # 完了日欠落は 着工日＋工事期間 で補完
            dur = duration_days(r[3])
            if dur is not None:
                end = start + datetime.timedelta(days=dur)

        amount = to_int(r[6])
        if r[6] is not None and amount is None:
            notes.append(f'{no} {site}: 請負金額が数値でないため空欄（元の値「{r[6]}」）')

        material = to_int(r[9]) or 0
        subcon = sum(to_int(r[c]) or 0 for c in (11, 13, 15, 17, 19, 21))

        if start is None:
            status = '見積'
        elif start > TODAY:
            status = '受注'
        elif end and end <= TODAY:
            status = '完成'
        else:
            status = '進行中'

        if subcon >= material and subcon > 0: cost_type = '外注費'
        elif material > 0:                    cost_type = '材料費'
        else:                                 cost_type = ''

        # 前期末未成工事支出金：未完成工事のみ既発生原価を期首残高として設定
        if seed_costs and status in ('進行中', '受注'):
            wip_mat, wip_sub = (material or ''), (subcon or '')
        else:
            wip_mat, wip_sub = '', ''

        out.append([no, site, cust, '', ym(start), ym(end),
                    amount if amount is not None else '',
                    wip_mat, '', wip_sub, '', status, cost_type])
    return HEADER, out, notes

def write_csv(path, header, rows):
    with open(path, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, lineterminator='\r\n')
        w.writerow(header)
        w.writerows(rows)

if __name__ == '__main__':
    for seed, path in ((False, 'koji_daicho_2026.csv'), (True, 'koji_daicho_2026_原価区分期首残高あり.csv')):
        h, rows, notes = convert(seed)
        write_csv(path, h, rows)
        print(f'{path}: {len(rows)}件')
    with open('work/conversion_notes.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(notes) + '\n')
    print('--- 補正・注意 ---')
    print('\n'.join(notes))
