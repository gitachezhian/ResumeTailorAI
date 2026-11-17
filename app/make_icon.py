"""Generates icon.ico (glossy green aero orb)."""
from PIL import Image, ImageDraw, ImageFilter
S = 256
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
# base orb: radial gradient
orb = Image.new("RGBA", (S, S))
px = orb.load()
cx, cy = S / 2, S * 1.05
for y in range(S):
    for x in range(S):
        d = ((x - cx) ** 2 + (y - cy) ** 2) ** 0.5 / (S * 0.95)
        d = min(1, d)
        c1, c2, c3 = (182, 255, 122), (69, 194, 26), (19, 118, 26)
        if d < 0.45:
            t = d / 0.45; c = [int(c1[i] + (c2[i] - c1[i]) * t) for i in range(3)]
        else:
            t = (d - 0.45) / 0.55; c = [int(c2[i] + (c3[i] - c2[i]) * t) for i in range(3)]
        px[x, y] = (*c, 255)
mask = Image.new("L", (S, S), 0)
ImageDraw.Draw(mask).ellipse((8, 8, S - 8, S - 8), fill=255)
ring = Image.new("L", (S, S), 0)
ImageDraw.Draw(ring).ellipse((2, 2, S - 2, S - 2), fill=255)
img.paste((255, 255, 255, 255), (0, 0), ring)
img.paste(orb, (0, 0), mask)
# gloss
gloss = Image.new("RGBA", (S, S), (0, 0, 0, 0))
g = ImageDraw.Draw(gloss)
for i in range(60):
    a = int(230 * (1 - i / 60))
    g.ellipse((40 + i * 0.2, 18 + i, S - 40 - i * 0.2, 130), outline=None, fill=None)
top = Image.new("L", (S, S), 0)
ImageDraw.Draw(top).ellipse((36, 16, S - 36, 136), fill=255)
grad = Image.new("L", (S, S), 0)
gp = grad.load()
for y in range(S):
    v = max(0, int(235 * (1 - (y - 16) / 120))) if y >= 16 else 0
    for x in range(S):
        gp[x, y] = v
from PIL import ImageChops
alpha = ImageChops.multiply(top, grad)
img.paste((255, 255, 255, 255), (0, 0), alpha)
# document + check glyph
d = ImageDraw.Draw(img)
w = 10
d.rounded_rectangle((84, 70, 172, 190), radius=12, outline=(255, 255, 255, 255), width=w)
d.line((108, 132, 124, 150, 152, 116), fill=(255, 255, 255, 255), width=w + 2, joint="curve")
img.save("icon.ico", sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
print("ok")
