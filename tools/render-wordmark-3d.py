"""3D chrome-blue wordmark for Golrang Future Stars from a white-on-black text mask."""
import sys
import numpy as np
from scipy import ndimage
from PIL import Image, ImageFilter, ImageDraw

src = Image.open(sys.argv[1]).convert('L')
# tilt like the original logo
src = src.rotate(6, resample=Image.BICUBIC, expand=True, fillcolor=0)
m = np.asarray(src).astype(float) / 255
H, W = m.shape
pad = 60
m = np.pad(m, pad)
H, W = m.shape
inside = m > .5

def norm(v):
    return v / np.maximum(np.sqrt((v ** 2).sum(-1, keepdims=True)), 1e-9)

# bevelled face: height rises over the first pixels from the edge, then flat
dist = ndimage.distance_transform_edt(inside)
bevel = 9.0
h = np.clip(dist / bevel, 0, 1)
h = np.sin(h * np.pi / 2) * 14
h = ndimage.gaussian_filter(h, 1.0)
gy, gx = np.gradient(h)
n = norm(np.dstack([-gx, -gy, np.ones_like(h) * 1.2]))
L = norm(np.array([-.35, -.7, .62]))
HV = norm(L + np.array([0, 0, 1.0]))
ndl = np.clip(n @ L, 0, 1)
spec = np.clip(n @ HV, 0, 1) ** 28

# chrome-blue face: a vertical sky-to-deep-blue band inside every letter row, reflecting the bevel
ys = np.arange(H)[:, None] / H
band = (np.sin(ys * np.pi * 3 * 2 - .6) + 1) / 2          # brighter at the top of each of the three rows
top = np.array([.93, .99, 1.0]); midc = np.array([.55, .84, 1.0]); low = np.array([.16, .45, .98])
face = np.where(band[..., None] > .5, midc + (top - midc) * ((band[..., None] - .5) * 2), low + (midc - low) * (band[..., None] * 2))
face = face * (.55 + .55 * ndl[..., None]) + spec[..., None] * .9
flat = dist > bevel
face = np.where(flat[..., None], face * 1.04, face)

# extruded depth: stacked copies shifted down and right, getting darker
depth = np.zeros((H, W, 4))
for i in range(16, 0, -1):
    sh = np.roll(np.roll(inside, i, 0), int(i * .35), 1)
    k = i / 16
    c = np.array([.05, .16, .52]) * (1 - .55 * k) + np.array([.02, .05, .2]) * .55 * k
    depth[sh] = np.r_[c, 1]
# thin dark outline around the face
outline = ndimage.binary_dilation(inside, iterations=3) & ~inside

img = depth.copy()
img[outline] = [.03, .08, .3, 1]
img[inside, :3] = np.clip(face[inside], 0, 1)
img[inside, 3] = 1

out = Image.fromarray((np.clip(img, 0, 1) * 255).astype(np.uint8), 'RGBA')
# blue glow behind
a = np.asarray(out)[..., 3]
glow = Image.fromarray(a).filter(ImageFilter.GaussianBlur(22))
g = Image.new('RGBA', out.size, (80, 170, 255, 0)); g.putalpha(glow.point(lambda v: int(v * .45)))
base = Image.new('RGBA', out.size, (0, 0, 0, 0))
base.alpha_composite(g); base.alpha_composite(out)
base = base.crop(base.getbbox())
base.save(sys.argv[2])
print(base.size)
