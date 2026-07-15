"""
안내 배너(assets/notice.png) 생성 스크립트.

텍스트 없는 배경(SRC) 위에 문구를 코드로 얹는다. 배경은 균일한 단색 파랑이라
좌측을 배경색으로 덮어(잔상·설비 제거) 우측 일러스트로 자연스럽게 페이드시킨 뒤
그 위에 2줄 문구를 그린다. 문구·폰트·색·크기를 바꾸려면 이 파일을 수정 후 실행.

  실행:  python "assets/build/notice_overlay.py"
"""
import os
from PIL import Image, ImageDraw, ImageFont

HERE  = os.path.dirname(os.path.abspath(__file__))
DST   = os.path.join(HERE, "..", "notice.png")           # 출력물
SRC   = r"d:\김용희\'26년 업무\09-1. 참고자료\이메일링 서비스\이메일 배너.png"  # 텍스트 없는 배경
FONTS = os.path.join(os.environ["LOCALAPPDATA"], r"Microsoft\Windows\Fonts")

WHITE = (255, 255, 255)
GOLD  = (255, 205, 0)
BG    = (29, 82, 139)          # 배경 단색 파랑

# 배경 경량화 (1280폭, 흰 배경 평탄화)
bg = Image.open(SRC).convert("RGBA").resize((1280, 239), Image.LANCZOS)
flat = Image.new("RGB", bg.size, (255, 255, 255))
flat.paste(bg, mask=bg.split()[3])
W, H = flat.size

# 좌측 단색 스크림: x<SOLID 완전 불투명, SOLID~FADE 선형 페이드, 이후 투명
SOLID, FADE = 855, 955
mask = Image.new("L", (W, H), 0)
mpx = mask.load()
for x in range(W):
    if x <= SOLID:
        a = 255
    elif x >= FADE:
        a = 0
    else:
        a = round(255 * (FADE - x) / (FADE - SOLID))
    for y in range(H):
        mpx[x, y] = a
flat.paste(Image.new("RGB", (W, H), BG), mask=mask)

# 텍스트 오버레이 (푸른숲체 Bold). 아랫줄은 윗줄보다 1pt 작게
X = 70
FPATH = os.path.join(FONTS, "유한킴벌리_푸른숲체 Bold.ttf")
f1 = ImageFont.truetype(FPATH, 38)
f2 = ImageFont.truetype(FPATH, 37)
line1 = [("개정 법령에 대하여 ", WHITE), ("법규준수평가", GOLD), ("를 진행해주세요!", WHITE)]
line2 = "꾸준함으로 완성되는 빈틈없는 법규준수!"
d = ImageDraw.Draw(flat)
x = X
for text, color in line1:
    d.text((x, 90), text, font=f1, fill=color, anchor="lm")
    x += d.textlength(text, font=f1)
d.text((X, 150), line2, font=f2, fill=WHITE, anchor="lm")

flat.quantize(colors=256, method=Image.MEDIANCUT,
              dither=Image.FLOYDSTEINBERG).save(DST, optimize=True)
print("저장:", DST, f"{os.path.getsize(DST)/1024:.0f}KB")
