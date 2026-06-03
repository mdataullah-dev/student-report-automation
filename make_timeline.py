"""Generates a clean 12-month timeline graphic (assets/timeline.png) styled
like the Saad reference report. Static image — no per-student data on it."""
from PIL import Image, ImageDraw, ImageFont

NAVY = (31, 56, 100)
RED = (200, 0, 0)
W, H = 1500, 300
img = Image.new("RGB", (W, H), "white")
d = ImageDraw.Draw(img)

def font(sz):
    for p in (r"C:\Windows\Fonts\arialbd.ttf", r"C:\Windows\Fonts\arial.ttf"):
        try:
            return ImageFont.truetype(p, sz)
        except Exception:
            pass
    return ImageFont.load_default()

f_month = font(22)

# Horizontal axis with arrowhead
y = 190
x0, x1 = 70, W - 70
d.line([(x0, y), (x1, y)], fill=NAVY, width=6)
d.polygon([(x1, y - 16), (x1 + 28, y), (x1, y + 16)], fill=NAVY)

# Start marker (red circle)
d.ellipse([(x0 - 14, y - 14), (x0 + 14, y + 14)], fill=RED)

# 12 month ticks + labels
n = 12
span = (x1 - 40 - x0)
for i in range(n):
    x = x0 + int(span * i / (n - 1))
    d.rectangle([(x - 4, y - 34), (x + 4, y)], fill=NAVY)  # tick bar
    label = "Month %d" % (i + 1)
    tb = d.textbbox((0, 0), label, font=f_month)
    tw = tb[2] - tb[0]
    d.text((x - tw / 2, y - 70), label, fill=(0, 0, 0), font=f_month)

img.save("assets/timeline.png")
print("wrote assets/timeline.png", img.size)
