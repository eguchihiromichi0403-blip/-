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
p.para("※23歳未満の扶養親族がいらっしゃる方は、今年から一般の生命保険料控除の上限が4万円から6万円に引き上げられています。")
p.blank()
p.para("3、社会保険料控除証明書（令和８年分）")
p.para("就職される前に支払っていた国民健康保険や国民年金、または、ご家族の方の社会保険料を支払っている場合も提出をお願い致します。")
p.blank()
p.para("4、地震保険料控除証明書（令和８年分）")
p.blank()
p.para("5、前職の源泉徴収票（途中入社の方）")
p.blank()
p.para("6、住宅ローン残高証明書（住宅ローンがある方のみ）")
c.save()
