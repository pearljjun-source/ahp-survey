# -*- coding: utf-8 -*-
"""OG 미리보기 이미지 생성 (1200x630 표준 크기)"""
try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
    from PIL import Image, ImageDraw, ImageFont

# 이미지 생성 (1200x630 - OG 표준)
width, height = 1200, 630
img = Image.new('RGB', (width, height), '#1F4E79')
draw = ImageDraw.Draw(img)

# 배경 그라데이션 효과 (단순 두 톤)
for y in range(height):
    r = int(31 + (46-31) * y/height)
    g = int(78 + (134-78) * y/height)
    b = int(121 + (193-121) * y/height)
    draw.line([(0, y), (width, y)], fill=(r, g, b))

# 하단 강조 바
draw.rectangle([(0, height-8), (width, height)], fill='#27AE60')

# 폰트 (시스템 기본 폰트 사용)
import os
font_paths = [
    "C:/Windows/Fonts/malgunbd.ttf",  # 맑은 고딕 Bold
    "C:/Windows/Fonts/malgun.ttf",     # 맑은 고딕
    "C:/Windows/Fonts/gulim.ttc",      # 굴림
]
font_path = None
for fp in font_paths:
    if os.path.exists(fp):
        font_path = fp
        break

if font_path:
    font_large = ImageFont.truetype(font_path, 42)
    font_medium = ImageFont.truetype(font_path, 28)
    font_small = ImageFont.truetype(font_path, 22)
    font_tiny = ImageFont.truetype(font_path, 18)
else:
    font_large = ImageFont.load_default()
    font_medium = font_large
    font_small = font_large
    font_tiny = font_large

# 상단 라벨
draw.rounded_rectangle([(40, 40), (320, 85)], radius=8, fill='#27AE60')
draw.text((55, 45), "AHP 전문가 설문", fill='white', font=font_small)

# 메인 타이틀
draw.text((40, 120), "충남 지역혁신체제 구축을 위한", fill='white', font=font_large)
draw.text((40, 175), "핵심가치와 정책요인의", fill='white', font=font_large)
draw.text((40, 230), "우선순위 분석", fill='#F1C40F', font=font_large)

# 구분선
draw.line([(40, 300), (500, 300)], fill='rgba(255,255,255,128)', width=2)

# 설명
draw.text((40, 320), "단국대학교 과학기술정책연구센터", fill='#D6E4F0', font=font_medium)
draw.text((40, 365), "조사기간: 2026.05.01 ~ 2026.05.31", fill='#D6E4F0', font=font_small)
draw.text((40, 400), "소요시간: 약 10분", fill='#D6E4F0', font=font_small)

# 오른쪽 아이콘 영역 - 체크리스트 시각화
icon_x = 750
icon_y = 100
items = ["자립성", "연계성", "실용성", "포용성", "지속가능성", "확장성"]
item_colors = ['#3498DB', '#2ECC71', '#F1C40F', '#E74C3C', '#9B59B6', '#E67E22']

for i, (item, color) in enumerate(zip(items, item_colors)):
    y = icon_y + i * 65
    # 바 배경
    draw.rounded_rectangle([(icon_x, y), (icon_x + 380, y + 48)], radius=10, fill='rgba(255,255,255,30)')
    draw.rounded_rectangle([(icon_x, y), (icon_x + 380, y + 48)], radius=10, outline='rgba(255,255,255,60)', width=1)
    # 색상 포인트
    draw.rounded_rectangle([(icon_x + 8, y + 8), (icon_x + 14, y + 40)], radius=3, fill=color)
    # 텍스트
    draw.text((icon_x + 28, y + 10), item, fill='white', font=font_small)
    # 체크마크
    draw.text((icon_x + 330, y + 10), "~", fill=color, font=font_small)

# 하단 URL
draw.text((40, height - 55), "pearljjun-source.github.io/ahp-survey", fill='#D6E4F0', font=font_tiny)

# 저장
img.save('og-image.png', quality=95)
print("OG 이미지 생성 완료: og-image.png")
