import os
import sys
import math
from PIL import Image, ImageDraw, ImageFont

# Add skill core to path
skill_path = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.join(skill_path, '.agent', 'skills', 'slack-gif-creator'))

from core.gif_builder import GIFBuilder
from core.easing import interpolate
from core.frame_composer import create_gradient_background

# Constants
WIDTH, HEIGHT = 480, 480
FPS = 15
TOTAL_FRAMES = 45 # 3 seconds
SLATE = (30, 41, 59)
SKY = (14, 165, 233)
AMBER = (245, 158, 11)
WHITE = (255, 255, 255)

# Font path - Windows specific
font_path = "C:\\Windows\\Fonts\\msjh.ttc" # Microsoft JhengHei
try:
    font_title = ImageFont.truetype(font_path, 40)
    font_sub = ImageFont.truetype(font_path, 24)
except:
    font_title = ImageFont.load_default()
    font_sub = ImageFont.load_default()

builder = GIFBuilder(width=WIDTH, height=HEIGHT, fps=FPS)

for i in range(TOTAL_FRAMES):
    t = i / (TOTAL_FRAMES - 1)
    
    # Create background (slight gradient)
    frame = create_gradient_background(WIDTH, HEIGHT, SLATE, (15, 23, 42))
    draw = ImageDraw.Draw(frame)
    
    # 1. Circle scale in (frames 0-15)
    if i < 15:
        scale_t = i / 14
        radius = interpolate(0, 100, scale_t, 'back_out')
        draw.ellipse([WIDTH/2 - radius, HEIGHT/2 - 100 - radius, WIDTH/2 + radius, HEIGHT/2 - 100 + radius], fill=SKY)
    else:
        # Static circle with small pulse
        pulse = math.sin(t * 10) * 5
        radius = 100 + pulse
        draw.ellipse([WIDTH/2 - radius, HEIGHT/2 - 100 - radius, WIDTH/2 + radius, HEIGHT/2 - 100 + radius], fill=SKY)
    
    # 2. Text fade in (frames 10-30)
    if i >= 10:
        alpha_t = min(1, (i - 10) / 20)
        alpha = int(255 * alpha_t)
        
        # PIL Draw doesn't support alpha directly for text easily on RGB,
        # so we'll just draw it if t > threshold for now, or use a layer.
        title_text = "HR Chatbox"
        sub_text = "進度更新說明!!"
        
        tw, th = draw.textbbox((0, 0), title_text, font=font_title)[2:]
        draw.text(((WIDTH - tw)/2, HEIGHT/2 + 20), title_text, font=font_title, fill=WHITE)
        
        sw, sh = draw.textbbox((0, 0), sub_text, font=font_sub)[2:]
        draw.text(((WIDTH - sw)/2, HEIGHT/2 + 80), sub_text, font=font_sub, fill=SKY)

    # 3. Success indicator pulse (frames 25-45)
    if i >= 25:
        indicator_t = (i - 25) / 20
        pulse = (math.sin(indicator_t * 2 * math.pi * 2) + 1) / 2 # 0 to 1
        ind_radius = 15 + pulse * 10
        draw.ellipse([WIDTH - 80 - ind_radius, 60 - ind_radius, WIDTH - 80 + ind_radius, 60 + ind_radius], 
                     fill=AMBER if i % 4 < 2 else (255, 255, 255))

    builder.add_frame(frame)

output_path = os.path.join(skill_path, 'output', 'update_promo.gif')
os.makedirs(os.path.dirname(output_path), exist_ok=True)
builder.save(output_path, num_colors=64)
print(f"GIF saved to: {output_path}")
