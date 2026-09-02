# -*- coding: utf-8 -*-
"""工事一覧Excel → ひな型準拠の工事台帳インポートCSV

対象: 6月決算のため 工期終了が 2026/07/01 以降の工事のみ
      （完了日が未定の工事＝進行中・受注・見積 も当期以降完成として含める）
"""
import openpyxl, datetime, csv, re, unicodedata

SRC        = 'work/source.xlsx'
SHEETS     = ['令和08年 ', '令和07年']
OUT        = 'koji_daicho_2026.csv'
NOTES      = 'work/conversion_notes.txt'
TODAY      = datetime.date(2026, 9, 2)     # 状態の判定基準日
PERIOD_FROM = datetime.date(2026, 7, 1)    # 当期首（6月決算）
EPOCH      = datetime.date(1899, 12, 30)

HEADER = ['工事番号','現場名','得意先名','得意先コード','工期開始','工期終了','請負金額',
          '前期末未成工事支出金（材料費）','前期末未成工事支出金（労務費）',
          '前期末未成工事支出金（外注費）','前期末未成工事支出金（経費）','状態','既定の原価区分']

# シートごとの列位置（0始まり）: 元請け, 現場名, 着工日, 工事期間, 完了日, 請負金額, 材料金額, 外注費列
LAYOUT = {
    '令和08年 ': dict(cust=0, site=1, start=2, dur=3, end=4, amount=6, material=9,
                     subcon=(11, 13, 15, 17, 19, 21)),
    '令和07年':  dict(cust=0, site=2, start=3, dur=4, end=5, amount=7, material=None,
                     subcon=(9,)),
}

CUSTOMER_ALIAS = {'k2': 'K2', 'ｋ２': 'K2'}   # 大小文字・全半角の表記ゆれのみ統一

def norm_site(v):
    """現場名は原文を尊重し、前後・連続空白の整理のみ"""
    return re.sub(r'[\s　]+', ' ', str(v)).strip() if v is not None else ''

def norm_customer(v):
    if v is None: return ''
    s = re.sub(r'[\s　]+', ' ', unicodedata.normalize('NFKC', str(v))).strip()
    return CUSTOMER_ALIAS.get(s, CUSTOMER_ALIAS.get(s.lower(), s))

def to_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date):     return v
    if isinstance(v, (int, float)) and 30000 < v < 60000:
        return EPOCH + datetime.timedelta(days=int(v))
    return None

def to_int(v):
    if isinstance(v, bool): return None
    return int(round(v)) if isinstance(v, (int, float)) else None

def duration_days(v):
    if v is None: return None
    m = re.search(r'(\d+)', unicodedata.normalize('NFKC', str(v)))
    return int(m.group(1)) if m else None

def add_year(dt):
    try:    return dt.replace(year=dt.year + 1)
    except ValueError: return dt.replace(year=dt.year + 1, day=28)

def add_months(dt, n):
    total = dt.month - 1 + n
    return datetime.date(dt.year + total // 12, total % 12 + 1, 1)

def ym(dt):
    return f'{dt.year}/{dt.month:02d}' if dt else ''

def extract():
    """全シートから工事を抽出し、日付の破損を補正して返す"""
    wb = openpyxl.load_workbook(SRC, data_only=True)
    items = []
    for sheet in SHEETS:
        L = LAYOUT[sheet]
        for i, r in enumerate(wb[sheet].iter_rows(values_only=True)):
            if i == 0 or not any(c is not None and str(c).strip() for c in r):
                continue
            site = norm_site(r[L['site']])
            if not site:
                continue
            label = f'{sheet.strip()} {i + 1}行目「{site}」'
            memo = []
            start, end = to_date(r[L['start']]), to_date(r[L['end']])
            raw_end = r[L['end']]
            dur = duration_days(r[L['dur']])

            if raw_end is not None and end is None:
                memo.append('完了日「%s」を解釈できないため、着工日＋工事期間から補完' % raw_end)
            if start and end and end < start:
                if (start - end).days > 300:      # 年の入力ミス（1年前になっている）
                    end = add_year(end)
                else:                             # 日付の打ち間違い
                    end = start + datetime.timedelta(days=dur) if dur is not None else start
                memo.append(f'完了日 {to_date(raw_end)} が着工日より前 → {end} に補正')
            if start and end is None and dur is not None:
                end = start + datetime.timedelta(days=dur)

            amount = to_int(r[L['amount']])
            if r[L['amount']] is not None and amount is None:
                memo.append(f'請負金額が数値でないため空欄（元の値「{r[L["amount"]]}」）')

            material = (to_int(r[L['material']]) or 0) if L['material'] is not None else 0
            subcon = sum(to_int(r[c]) or 0 for c in L['subcon'])

            items.append(dict(sheet=sheet, label=label, site=site,
                              cust=norm_customer(r[L['cust']]), start=start, end=end,
                              amount=amount, material=material, subcon=subcon, memo=memo))
    return items

def build():
    items = extract()
    rows, notes = [], []
    for it in items:
        if it['end'] is not None and it['end'] < PERIOD_FROM:
            continue                              # 前期以前に完成 → 当期の台帳の対象外
        if it['end'] is None:
            it['memo'].append('完了日が未定のため当期の工事として取り込み')
        notes += [f'{it["label"]}: {m}' for m in it['memo']]

        start, end = it['start'], it['end']
        if start is None:                status = '見積'
        elif start > TODAY:              status = '受注'
        elif end and end <= TODAY:       status = '完成'
        else:                            status = '進行中'

        if it['subcon'] >= it['material'] and it['subcon'] > 0: cost_type = '外注費'
        elif it['material'] > 0:                                cost_type = '材料費'
        else:                                                   cost_type = ''

        # 工期終了が未定なら工期開始の2か月後を仮置き（状態・抽出条件の判定には使わない）
        end_out = end
        if end_out is None and start is not None:
            end_out = add_months(start, 2)
            notes.append(f'{it["label"]}: 工期終了が未定のため工期開始の2か月後 {ym(end_out)} を仮置き')

        rows.append([None, it['site'], it['cust'], '', ym(start), ym(end_out),
                     it['amount'] if it['amount'] is not None else '',
                     '', '', '', '',            # 前期末未成工事支出金は全件空欄
                     status, cost_type])

    rows.sort(key=lambda x: (x[4] == '', x[4], x[5]))   # 工期開始順、未定は末尾
    for n, row in enumerate(rows, 1):
        row[0] = f'2026-{n:03d}'
    return rows, notes

if __name__ == '__main__':
    rows, notes = build()
    with open(OUT, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, lineterminator='\r\n')
        w.writerow(HEADER)
        w.writerows(rows)
    with open(NOTES, 'w', encoding='utf-8') as f:
        f.write('\n'.join(notes) + '\n')
    print(f'{OUT}: {len(rows)}件')
    print('\n'.join(notes))
