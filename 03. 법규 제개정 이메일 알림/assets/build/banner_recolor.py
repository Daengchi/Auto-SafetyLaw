"""
헤더 배너(assets/banner.png) 생성 스크립트.

로고·슬로건 원본은 참고자료 폴더(SRC)에서 읽어 합성한다.
색·문구·크기를 바꾸려면 이 파일을 수정한 뒤 실행하면 banner.png가 갱신된다.

  실행:  python "assets/build/banner_recolor.py"
"""
import os
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
DST  = os.path.join(HERE, "..", "banner.png")            # 출력물
SRC  = r"d:\김용희\'26년 업무\09-1. 참고자료\이메일링 서비스"   # 로고·슬로건 원본

NAVY  = (29, 82, 139)    # 배너·법령헤더와 동일한 남색 (#1D528B)
TITLE = "법규 제개정 알림 서비스"

W, H, PAD, ROW_Y = 1280, 300, 52, 90
BOLD = ImageFont.truetype(r"C:\Windows\Fonts\malgunbd.ttf", 46)


def load(name, h):
    im = Image.open(os.path.join(SRC, name)).convert("RGBA")
    return im.resize((round(im.width * h / im.height), h), Image.LANCZOS)


c = Image.new("RGBA", (W, H), (255, 255, 255, 255))
byd, kkp = load("Beyond the best.png", 46), load("금호석유화학.png", 40)
c.alpha_composite(byd, (PAD, ROW_Y - byd.height // 2))
c.alpha_composite(kkp, (PAD + byd.width + 22, ROW_Y - kkp.height // 2 + 2))
sf = load("안전보건 슬로건.png", 84)
c.alpha_composite(sf, (W - PAD - sf.width, ROW_Y - sf.height // 2))

d = ImageDraw.Draw(c)
d.line([(PAD, 172), (W - PAD, 172)], fill=(232, 232, 232), width=2)
d.rectangle([PAD, 209, PAD + 7, 262], fill=NAVY)          # 제목 옆 라인 → 밝은 남색
d.text((PAD + 24, 236), TITLE, font=BOLD, fill=(51, 51, 51), anchor="lm")
d.rectangle([0, H - 8, W, H], fill=NAVY)                   # 하단 바 → 배너 남색

c.convert("RGB").save(DST)
print("저장:", DST, f"{os.path.getsize(DST)/1024:.0f}KB")
