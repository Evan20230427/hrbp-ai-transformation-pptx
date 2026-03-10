import os
import sys
import math
from PIL import Image, ImageDraw, ImageEnhance

class LocalGIFBuilder:
    def __init__(self, width=800, height=450, fps=15):
        self.width = width
        self.height = height
        self.fps = fps
        self.frames = []

    def add_frame(self, img):
        self.frames.append(img.copy())

    def save(self, path):
        self.frames[0].save(
            path,
            save_all=True,
            append_images=self.frames[1:],
            duration=int(1000/self.fps),
            loop=0
        )

# Use the realistic office image as base
base_img_path = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\brain\\cb95dffe-33bd-4e40-a98b-feaff376ea1a\\hrbp_office_collaboration_v10_gif_base_1773159625189.png"
if not os.path.exists(base_img_path):
    # Fallback to empty if not found
    base_img = Image.new('RGB', (800, 450), (15, 23, 42))
else:
    base_img = Image.open(base_img_path).resize((800, 450))

builder = LocalGIFBuilder(800, 450, 15)
TOTAL_FRAMES = 30

for i in range(TOTAL_FRAMES):
    t = i / (TOTAL_FRAMES - 1)
    frame = base_img.copy()
    draw = ImageDraw.Draw(frame, 'RGBA')
    
    # Add animated "Data Energy" nodes over the table/environment
    # Pulsing glow on screens or people
    glow_alpha = int(100 + math.sin(t * 2 * math.pi) * 50)
    # Draw subtle growing connections
    for link in range(3):
        start = (100 + link * 200, 200)
        end = (start[0] + 150 * t, 200 + 50 * math.sin(t*math.pi))
        draw.line([start, end], fill=(59, 130, 246, glow_alpha), width=3)
        draw.ellipse([end[0]-8, end[1]-8, end[0]+8, end[1]+8], fill=(16, 185, 129, glow_alpha))

    # Floating UI elements (simple rectangles)
    ui_y = 100 + math.sin(t * math.pi) * 10
    draw.rectangle([600, ui_y, 750, ui_y + 80], fill=(255, 255, 255, 40), outline=(255, 255, 255, 100))
    draw.text((610, ui_y + 10), "AI ANALYZING...", fill=(255, 255, 255, 200))
    
    builder.add_frame(frame)

output_path = os.path.join(os.getcwd(), 'hrbp_office_dynamic_v10.gif')
builder.save(output_path)
print(f"Realistic GIF saved to: {output_path}")
