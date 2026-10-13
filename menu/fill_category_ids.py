# エアレジからダウンロードした商品CSVのカテゴリーID欄を埋めて、更新用CSVを作る
import csv
from pathlib import Path

# カテゴリー設定画面の並び順（作成順）とダウンロードCSVの番号から判断
CATEGORY_IDS = {
    "アルコール": "0001",
    "ソフトドリンク": "0002",
    "ランチ": "0003",
    "ディナー": "0004",
}

out_dir = Path(__file__).parent

with open(out_dir / "airregi_menu_utf8.csv", encoding="utf-8-sig", newline="") as f:
    category_by_name = {row["商品名"]: row["カテゴリー"] for row in csv.DictReader(f)}

with open(out_dir / "airregi_download_20261013.csv", encoding="cp932", newline="") as f:
    rows = list(csv.reader(f))

for row in rows[2:]:
    name = row[4]
    # 管理画面で手入力済みのカテゴリーはそのまま残す
    if not row[3]:
        row[3] = CATEGORY_IDS[category_by_name[name]]

with open(out_dir / "airregi_menu_update.csv", "w", encoding="cp932", newline="") as f:
    csv.writer(f, quoting=csv.QUOTE_ALL, lineterminator="\r\n").writerows(rows)

with open(out_dir / "airregi_menu_update_utf8.csv", "w", encoding="utf-8-sig", newline="") as f:
    writer = csv.writer(f, lineterminator="\r\n")
    names = {v: k for k, v in CATEGORY_IDS.items()}
    writer.writerow(["商品ID", "カテゴリーID", "カテゴリー", "商品名", "価格"])
    for row in rows[2:]:
        writer.writerow([row[0], row[3], names[row[3]], row[4], row[14]])
