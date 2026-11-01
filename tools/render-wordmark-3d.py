"""Golrang Future Stars wordmark, after the original logo: pale icy letter faces, a thick bright-blue
rim around every letter merging into one block, dark-blue depth below, a blue glow and a neon line behind.

Usage: python3 logo3d.py MASK.png OUT.png   (MASK: white text on black)
"""
import sys
import numpy as np
from scipy import ndimage
from PIL import Image, ImageFilter, ImageDraw

src = Image.open(sys.argv[1]).convert('L')
bb = Image.eval(src, lambda v: 255 if v > 128 else 0).getbbox()
pad = 120
src = src.crop((bb[0] - pad, bb[1] - pad, bb[2] + pad, bb[3] + pad))
# a slight counter-clockwise tilt, as in the original
src = src.rotate(3, resample=Image.BICUBIC, expand=True, fillcolor=0)
m = np.asarray(src).astype(float) / 255
H, W = m.shape
face = m > .5


def norm(v):
    return v / np.maximum(np.sqrt((v ** 2).sum(-1, keepdims=True)), 1e-9)


L = norm(np.array([-.3, -.75, .6]))
HV = norm(L + np.array([0, 0, 1.0]))

# rings around the letters: a bright rim, then the merged block edge
d_out = ndimage.distance_transform_edt(~face)
rim = (d_out <= 13) & ~face
block = d_out <= 13
rim = block & ~face

# depth: the block pushed down a little, darker toward the bottom
depth = np.zeros((H, W), bool)
for i in range(1, 15):
    depth |= np.roll(np.roll(block, i, 0), int(i * .25), 1)
depth &= ~block

img = np.zeros((H, W, 4))
ys = np.linspace(0, 1, H)[:, None] * np.ones((1, W))
dcol = np.array([.10, .27, .78])[None, None] * (1 - .45 * ys[..., None]) + np.array([.03, .08, .3])[None, None] * .45 * ys[..., None]
img[depth, :3] = dcol[depth]; img[depth, 3] = 1

# rim: bright blue, rounded so it catches light on its upper edge
dr = ndimage.distance_transform_edt(block)
rh = np.sin(np.clip(dr / 13, 0, 1) * np.pi / 2) * 10
gy, gx = np.gradient(ndimage.gaussian_filter(rh, 1.2))
rn = norm(np.dstack([-gx, -gy, np.ones_like(rh) * 1.4]))
rl = np.clip(rn @ L, 0, 1)
rcol = np.array([.16, .46, 1.0]) * (.55 + .65 * rl[..., None]) + (np.clip(rn @ HV, 0, 1) ** 24)[..., None] * .55
img[rim, :3] = rcol[rim]; img[rim, 3] = 1

# letter faces: pale ice, a soft top-to-bottom shade per letter and a thin bevel
df = ndimage.distance_transform_edt(face)
fh = np.sin(np.clip(df / 5, 0, 1) * np.pi / 2) * 6
gy, gx = np.gradient(ndimage.gaussian_filter(fh, .8))
fn = norm(np.dstack([-gx, -gy, np.ones_like(fh) * 1.1]))
fl = np.clip(fn @ L, 0, 1)
# per-row vertical shade: lighter at the top of each line of letters
lab, n = ndimage.label(ndimage.binary_dilation(face, iterations=4, structure=np.ones((1, 9))))
shade = np.zeros((H, W))
for sl in ndimage.find_objects(lab):
    if sl is None:
        continue
    y0, y1 = sl[0].start, sl[0].stop
    t = (np.arange(y0, y1) - y0) / max(1, y1 - y0 - 1)
    shade[y0:y1, sl[1]] = t[:, None]
top, bot = np.array([.95, .995, 1.0]), np.array([.66, .86, .97])
fc = top + (bot - top) * shade[..., None]
fc = fc * (.78 + .3 * fl[..., None]) + (np.clip(fn @ HV, 0, 1) ** 30)[..., None] * .35
img[face, :3] = np.clip(fc[face], 0, 1); img[face, 3] = np.clip(m[face] * 1.2, 0, 1)

out = Image.fromarray((np.clip(img, 0, 1) * 255).astype(np.uint8), 'RGBA')

# neon parallelogram line behind, running out past the block on both sides
bx = block.any(0).nonzero()[0]; by = block.any(1).nonzero()[0]
x0, x1, y0, y1 = bx[0], bx[-1], by[0], by[-1]
neon = Image.new('RGBA', out.size, (0, 0, 0, 0))
dn = ImageDraw.Draw(neon)
pts = [(x0 - 90, y0 + (y1 - y0) * .42), (x1 + 20, y0 + (y1 - y0) * .18), (x1 + 90, y0 + (y1 - y0) * .62), (x0 - 20, y0 + (y1 - y0) * .9)]
dn.line(pts + [pts[0]], fill=(110, 210, 255, 255), width=5, joint='curve')
neon_glow = neon.filter(ImageFilter.GaussianBlur(9))

# blue glow around the whole block
a = Image.fromarray((block | depth).astype(np.uint8) * 255).filter(ImageFilter.GaussianBlur(26))
glow = Image.new('RGBA', out.size, (70, 150, 255, 0)); glow.putalpha(a.point(lambda v: int(v * .55)))

base = Image.new('RGBA', out.size, (0, 0, 0, 0))
base.alpha_composite(neon_glow); base.alpha_composite(neon)
base.alpha_composite(glow); base.alpha_composite(out)
base = base.crop(base.getbbox())
base.save(sys.argv[2])
print(base.size)
