import sys
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

out = sys.argv[1]
pdfmetrics.registerFont(TTFont("G", "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf"))
W, H = landscape(A4)
LEFT, RIGHT = 85.1, W - 85.1
FS, PITCH = 12, 18
NOHEAD = "、。，．）」』】〕・ー：；？！ぁぃぅぇぉっゃゅょァィゥェォッャュョ)"

def wrap(text, fs=FS, indent=0):
    lines, cur = [], ""
    maxw = RIGHT - LEFT - indent
    for ch in text:
        if pdfmetrics.stringWidth(cur + ch, "G", fs) > maxw and ch not in NOHEAD:
            lines.append(cur); cur = ch
        else:
            cur += ch
    lines.append(cur)
    return lines

class Page:
    def __init__(self, c, top):
        self.c, self.y = c, H - top - FS * 0.88  # baseline from line top
    def para(self, text, indent=0, highlight=False):
        for i, ln in enumerate(wrap(text, indent=indent)):
            x = LEFT + (indent if i == 0 else 0)
            if highlight:
                w = pdfmetrics.stringWidth(ln, "G", FS)
                self.c.setFillColorRGB(1, 1, 0)
                self.c.rect(x, self.y - 3.2, w, 15.6, stroke=0, fill=1)
                self.c.setFillColorRGB(0, 0, 0)
            self.c.setFont("G", FS); self.c.drawString(x, self.y, ln)
            self.y -= PITCH
    def blank(self, n=1):
        self.y -= PITCH * n

c = canvas.Canvas(out, pagesize=(W, H))
c.setTitle("年末調整のお知らせ 令和8年度"); c.setAuthor("eguch")

# ---- page 1 ----
title = "～年末調整のお知らせ　令和８年度～"
c.setFont("G", 16); c.drawCentredString(W / 2, H - 95.1 - 14, title)
p = Page(c, 124.1)
p.para("＜提出書類＞")
p.para("1、令和８年分給与所得者の扶養控除等(異動)申告書")
p.para("税務署から年末調整の封筒が届いていると思いますが、その中に入っております。足りない分はコピーして使っていただくことになります。", indent=7.9)
p.para("(特に従業員様に変更がない場合は、改めて書いていただくこともないと思いますので臨機応変に対応してまいります。)")
p.para("その他、配偶者控除、特定親族特別控除についての申告書もございますが、こちらについては弊社の方で作成をさせて頂きます。")
p.blank()
p.para("２、生命保険料控除証明書（令和８年分）")
p.para("※23歳未満の扶養親族がいらっしゃる方は、今年から生命保険料控除が拡充されています(別紙④)。")
p.blank()
p.para("3、社会保険料控除証明書（令和８年分）")
p.para("就職される前に支払っていた国民健康保険や国民年金、または、ご家族の方の社会保険料を支払っている場合も提出をお願い致します。")
p.blank()
p.para("4、地震保険料控除証明書（令和８年分）")
p.blank()
p.para("5、前職の源泉徴収票（途中入社の方）")
p.blank()
p.para("6、住宅ローン残高証明書（住宅ローンがある方のみ）")
p.blank()
p.para("7、令和８年の税制改正内容(別紙)")
c.showPage()

# ---- page 2 ----
p = Page(c, 88.1)
p.para("今年の年末調整も、昨年に引き続き改正された部分がございます。大きく分けて４点ですが、気を付けて頂きたいのは②～④になります。")
p.para("なお、今回の改正は令和８年12月1日から適用されるため、11月までの毎月の給与計算は従来どおりで、12月の年末調整でまとめて精算いたします。")
p.blank()
p.para("①  基礎控除と給与所得控除のさらなる引き上げ", highlight=True)
p.para("令和８年度の税制改正により、昨年に続いて基礎控除と給与所得控除の控除額がさらに引き上げられました。")
p.para("こちらにつきましては、弊社での年末調整の計算上、考慮して計算を行いますので特に気にしていただくことはございません。")
p.para("簡単な説明としては、本人に税金がかかる年収の壁が、昨年の160万円から178万円になります。(給与所得のケース)")
p.blank()
p.para("②  年収の壁の変更", highlight=True)
p.para("扶養家族(配偶者含む)の扶養に入ることができる壁が、昨年の123万円から136万円に変更になっています。(給与所得のケース)")
p.para("※社会保険の扶養(年収130万円など)の基準は別ですので、ご注意ください。")
p.blank()
p.para("③  特定親族特別控除の対象範囲の変更", highlight=True)
p.para("年齢19歳以上23歳未満の親族(大学生の年齢)のお子様がいらっしゃる従業員様への改正内容です。")
p.para("昨年創設された特定親族特別控除の対象となる年収が引き上げられ、お子様の年収が136万円を超えて197万円までの場合に控除を受けることが可能になっています。(年収159万円までは満額の控除)")
p.para("該当する方は、お子様の今年の年収(見込み)をお知らせください。")
p.blank()
p.para("④  生命保険料控除の拡充", highlight=True)
p.para("23歳未満の扶養親族がいらっしゃる従業員様は、一般の生命保険料控除(新契約分)の上限が4万円から6万円に引き上げられます。")
p.para("該当する方は、生命保険料控除証明書の提出をお忘れなくお願いいたします。")
print("page2 last baseline y from top:", H - p.y)
c.save()
