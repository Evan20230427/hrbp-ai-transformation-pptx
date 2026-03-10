import os
import sys
import math
from PIL import Image, ImageDraw, ImageFont

# Path config
skill_path = "C:\\Users\\TW-Evan.Chen\\.gemini\\antigravity\\skills\\youtube-downloader" # Using existing skill path structure
if not os.path.exists(skill_path):
    skill_path = os.getcwd()

# Simple GIF Builder (Self-contained)
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

# Constants
WIDTH, HEIGHT = 800, 450
FPS = 15
TOTAL_FRAMES = 30 
SLATE = (15, 23, 42)
BLUE = (37, 99, 235)
EMERALD = (16, 185, 129)
WHITE = (255, 255, 255)

builder = LocalGIFBuilder(WIDTH, HEIGHT, FPS)

for i in range(TOTAL_FRAMES):
    t = i / (TOTAL_FRAMES - 1)
    img = Image.new('RGB', (WIDTH, HEIGHT), SLATE)
    draw = ImageDraw.Draw(img)
    
    # Draw animated connections (lines)
    for line_idx in range(5):
        offset = (t + line_idx * 0.2) % 1.0
        lx = int(offset * WIDTH)
        ly = int(HEIGHT / 2 + math.sin(offset * 2 * math.pi) * 100)
        draw.line([0, ly, lx, ly], fill=(50, 50, 100), width=2)
        draw.ellipse([lx-5, ly-5, lx+5, ly+5], fill=BLUE)

    # Pulsing Center
    pulse = (math.sin(t * 2 * math.pi) + 1) / 2
    r = 50 + int(pulse * 20)
    draw.ellipse([WIDTH/2-r, HEIGHT/2-r, WIDTH/2+r, HEIGHT/2+r], outline=EMERALD, width=3)
    
    # Text placeholder (simple)
    draw.text((WIDTH/2 - 100, HEIGHT/2 + 80), "AI TRANSFORMATION", fill=WHITE)
    
    builder.add_frame(img)

output_path = os.path.join(os.getcwd(), 'hrbp_animated_visual.gif')
builder.save(output_path)
print(f"GIF saved to: {output_path}")
