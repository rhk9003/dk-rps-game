# -*- coding: utf-8 -*-
"""產出府中店開幕活動 - 獎項總覽 JPG（給店長 / 督導 LINE 用）"""
from PIL import Image, ImageDraw, ImageFont

# ===== 配色 =====
VELVET = "#940201"
SCARLET = "#e41e07"
BLUE = "#1198cd"
DEEP_BLUE = "#0a6d85"
CREAM = "#ece0d8"
CREAM_DARK = "#dbcfc0"
BLACK = "#2D3142"
WHITE = "#ffffff"

FONT_BOLD = "C:/Windows/Fonts/msjhbd.ttc"
FONT_REG = "C:/Windows/Fonts/msjh.ttc"

def F(size, bold=True):
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REG, size)

# ===== 獎項資料（不加 emoji 避免顯示成 □）=====
WIN_PRIZES = [
    ("DK 動態足弓拖鞋",            "$1,280",  "8%",    "限 30 雙 · 每日 1 雙"),
    ("DK 五趾隱形襪",              "$250",    "25%",   "限 60 雙 · 每日 2 雙"),
    ("官網鞋款滿 $2,000 折 $500",  "$500",    "20%",   "LINE 券 · 不限量"),
    ("門市 7 折券",                 "7 折",    "25%",   "LINE 券 · 不限量"),
    ("門市 $500 折價券",            "$500",    "22%",   "LINE 券 · 不限量"),
]
LOSE_PRIZES = [
    ("石墨烯袖套（隱藏款）",       "$680",    "2.5%",  "限 10 件 · 抽完為止"),
    ("DK 涼感底紗襪",              "$180",    "20%",   "限 60 雙 · 每日 2 雙"),
    ("$50 官網抵用金",             "$50",     "20%",   "LINE 券 · 不限量"),
    ("門市 8 折券",                 "8 折",    "27.5%", "LINE 券 · 不限量"),
    ("門市 $100 折價券",            "$100",    "30%",   "LINE 券 · 不限量"),
]

# ===== 畫布 =====
W = 1080
H = 2200   # 加高避免下方備註被切
img = Image.new("RGB", (W, H), CREAM)
d = ImageDraw.Draw(img)

# ===== 頂部橫幅 =====
d.rectangle([0, 0, W, 300], fill=SCARLET)
d.rectangle([0, 300, W, 316], fill=VELVET)
d.text((W//2, 65), "DK 府中店", font=F(42), fill=CREAM, anchor="mm")
d.text((W//2, 150), "開幕活動．來玩猜拳", font=F(76), fill=WHITE, anchor="mm")
d.text((W//2, 235), "和我玩猜拳．三戰二勝．通通有獎", font=F(34, bold=False), fill=CREAM, anchor="mm")

# ===== 共用繪列函式 =====
def draw_section(y, title, prizes, header_color, value_color, prob_color):
    # 區塊標題
    d.rounded_rectangle([50, y, W-50, y+90], radius=20, fill=header_color)
    d.text((W//2, y+45), title, font=F(40), fill=WHITE, anchor="mm")
    y += 110

    # 欄位 header
    col_name, col_val, col_prob = 85, 640, 850
    d.rectangle([50, y, W-50, y+50], fill=header_color)
    d.text((col_name, y+25), "獎項", font=F(26), fill=WHITE, anchor="lm")
    d.text((col_val,  y+25), "面額", font=F(26), fill=WHITE, anchor="lm")
    d.text((col_prob, y+25), "機率", font=F(26), fill=WHITE, anchor="lm")
    y += 55

    for i, (name, value, prob, limit) in enumerate(prizes):
        row_h = 105
        bg = WHITE if i % 2 == 0 else CREAM_DARK
        d.rectangle([50, y, W-50, y+row_h], fill=bg)
        d.text((col_name, y+30), name, font=F(30), fill=BLACK, anchor="lm")
        d.text((col_name, y+70), f"· {limit}", font=F(22, bold=False), fill="#666666", anchor="lm")
        d.text((col_val,  y+row_h//2), value, font=F(28), fill=value_color, anchor="lm")
        d.text((col_prob, y+row_h//2), prob, font=F(32), fill=prob_color, anchor="lm")
        y += row_h
        d.line([(50, y), (W-50, y)], fill=CREAM_DARK, width=1)
    return y

# ===== 贏家池 =====
y = 380
y = draw_section(y, "贏家大獎池（機率合計 100%）", WIN_PRIZES, VELVET, VELVET, SCARLET)

# ===== 輸家池 =====
y += 50
y = draw_section(y, "輸家．參加獎池（機率合計 100%）", LOSE_PRIZES, DEEP_BLUE, DEEP_BLUE, BLUE)

# ===== 底部備註 =====
y += 35
note_lines = [
    "※ 實體獎品：DK 動態足弓拖鞋、DK 五趾隱形襪、石墨烯袖套、DK 涼感底紗襪",
    "※ LINE 券走會員帳號發送，與 POS 連動核銷，客人不需兌換碼",
    "※ 實體獎抽完將自動改發對應的 LINE 券（客人仍有獎，不空手離開）",
]
for line in note_lines:
    d.text((W//2, y), line, font=F(22, bold=False), fill="#444444", anchor="mm")
    y += 34

# ===== 底線 =====
d.rectangle([0, H-55, W, H], fill=VELVET)
d.text((W//2, H-27), "DK 高博士專業氣墊鞋  ·  府中店", font=F(24), fill=WHITE, anchor="mm")

# ===== 存檔 =====
out = "C:/Users/drkao/Desktop/遊戲設計/.claude/tmp/府中店開幕活動_獎項總覽.jpg"
img.save(out, "JPEG", quality=92)
print("OK:", out, img.size)
