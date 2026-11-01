"""Render the classic Golrang cloud as a soft 3D clay character (height field + lighting).

Usage: python3 render3d.py OUT_PREFIX
Writes OUT_PREFIX-l1.png .. -l4.png (RGBA). Coordinates follow the SVG drawing (viewBox -30 -120 800 790).
"""
import sys
import numpy as np
from scipy import ndimage
from PIL import Image, ImageFilter, ImageDraw

S = 1.25
X0, Y0, VW, VH = -30, -120, 800, 790
W, H = int(VW * S), int(VH * S)
yy, xx = np.mgrid[0:H, 0:W].astype(np.float64)
# work in drawing units
ux = xx / S + X0
uy = yy / S + Y0


def norm(v):
    n = np.sqrt((v ** 2).sum(-1, keepdims=True))
    return v / np.maximum(n, 1e-9)


def dome(cx, cy, rx, ry, hz, rot=0.0):
    """Ellipsoid cap height; -inf outside."""
    c, s = np.cos(np.radians(rot)), np.sin(np.radians(rot))
    dx, dy = ux - cx, uy - cy
    x2, y2 = dx * c + dy * s, -dx * s + dy * c
    d2 = (x2 / rx) ** 2 + (y2 / ry) ** 2
    h = hz * np.sqrt(np.clip(1 - d2, 0, None))
    return np.where(d2 < 1, h, -1e3), d2


def smax(hs, k):
    hs = np.stack(hs)
    m = hs.max(0)
    return m + k * np.log(np.exp((hs - m) / k).sum(0))


def bezier(p, n=160):
    t = np.linspace(0, 1, n)[:, None]
    p = [np.array(q, float) for q in p]
    if len(p) == 3:
        return (1 - t) ** 2 * p[0] + 2 * (1 - t) * t * p[1] + t ** 2 * p[2]
    return (1 - t) ** 3 * p[0] + 3 * (1 - t) ** 2 * t * p[1] + 3 * (1 - t) * t ** 2 * p[2] + t ** 3 * p[3]


def tube(points, r, taper=None):
    """Height of a round tube along a polyline (drawing units)."""
    pts = np.asarray(points)
    x0, y0 = pts.min(0) - r - 2
    x1, y1 = pts.max(0) + r + 2
    sel = (ux >= x0) & (ux <= x1) & (uy >= y0) & (uy <= y1)
    h = np.full(ux.shape, -1e3)
    px, py = ux[sel], uy[sel]
    best = np.full(px.shape, 1e9)
    rad = np.full(px.shape, float(r))
    for i, (qx, qy) in enumerate(pts):
        d = np.hypot(px - qx, py - qy)
        rr = r if taper is None else r * taper(i / (len(pts) - 1))
        better = d - rr < best - rad
        best = np.where(better, d, best)
        rad = np.where(better, rr, rad)
    hh = np.sqrt(np.clip(rad ** 2 - best ** 2, 0, None))
    h[sel] = np.where(best < rad, hh, -1e3)
    return h


def normals(hf, scale=1.0):
    gy, gx = np.gradient(hf * scale)
    n = np.dstack([-gx, -gy, np.ones_like(hf) * S])
    return norm(n)


L = norm(np.array([-0.42, -0.62, 0.66]))
V = np.array([0, 0, 1.0])
HV = norm(L + V)


def shade(n, albedo, spec_pow, spec_amt, ao=1.0, wrap=0.45, rim_amt=0.35, sss=0.0):
    ndl = (n @ L)
    diff = np.clip((ndl + wrap) / (1 + wrap), 0, 1)
    sky = 0.30 + 0.22 * np.clip(-n[..., 1], 0, 1)          # light from above
    bounce = 0.16 * np.clip(n[..., 1], 0, 1)               # lavender bounce from below
    spec = np.clip(n @ HV, 0, 1) ** spec_pow * spec_amt
    rimdir = norm(np.array([0.75, 0.15, 0.0]))
    rim = (1 - np.clip(n[..., 2], 0, 1)) ** 2.2 * np.clip(n @ rimdir, 0, 1) * rim_amt
    a = np.asarray(albedo, float)
    if isinstance(ao, np.ndarray):
        ao = ao[..., None]
    col = a * (diff[..., None] * np.array([1.0, .98, .95]) * 0.78
               + sky[..., None] * np.array([.80, .83, 1.0])
               + bounce[..., None] * np.array([.78, .70, 1.0])) * ao
    col = col + sss * np.clip(1 - n[..., 2], 0, 1)[..., None] * a * 0.25
    col = col + spec[..., None] + rim[..., None] * np.array([.62, .72, 1.0])
    return col


def layer_alpha(h, edge=1.6):
    return np.clip(h / edge, 0, 1)


def render(level):
    # ---------- body: five lobes of the classic cloud, smooth-unioned so the creases stay soft
    lobes = [(372, 380, 230, 175, 120), (132, 272, 118, 118, 100), (592, 318, 122, 122, 102),
             (236, 468, 138, 104, 96), (474, 500, 152, 98, 96), (336, 222, 194, 190, 150)]
    hs = [dome(cx, cy, rx, ry, hz)[0] for cx, cy, rx, ry, hz in lobes]
    body = smax(hs, 9.0)
    inside = np.max(np.stack(hs), 0) > 0
    body = np.where(inside, np.maximum(body, 0), -1e3)
    bh = np.where(inside, body, 0)
    blur = ndimage.gaussian_filter(bh, 22 * S)
    ao = np.clip(1 - np.clip(blur - bh, 0, None) / 70.0, 0.5, 1)
    nb = normals(bh)
    col = shade(nb, (.985, .985, 1.0), 26, .10, ao, wrap=.55, rim_amt=.45, sss=.6)
    alpha = np.clip(np.where(inside, 1, 0).astype(float), 0, 1)
    # antialias the silhouette from the soft height
    alpha = np.clip(bh / 3.0, 0, 1) * inside
    canvas = np.dstack([col, alpha])

    def over(rgb, a):
        nonlocal canvas
        a = a[..., None]
        canvas[..., :3] = rgb * a + canvas[..., :3] * (1 - a)
        canvas[..., 3] = a[..., 0] + canvas[..., 3] * (1 - a[..., 0])

    def bumped(h_obj):
        """surface = body + object, normals from the sum so objects sit on the cloud"""
        hsum = bh + np.maximum(h_obj, 0)
        return normals(hsum)

    # blush
    for cx, cy in [(200, 252), (420, 242)]:
        g = np.exp(-(((ux - cx) / 26) ** 2 + ((uy - cy) / 15) ** 2))
        canvas[..., :3] = canvas[..., :3] * (1 - .22 * g[..., None]) + np.array([1, .55, .65]) * .22 * g[..., None]

    # ---------- face: brows, mouth and dimple as glossy dark tubes
    ink = (.09, .07, .16)
    strokes = [bezier([(266, 112), (268, 82), (296, 70)], 90), bezier([(308, 64), (348, 60), (364, 102)], 90),
               bezier([(212, 224), (240, 270), (330, 286), (382, 222)], 200), bezier([(378, 194), (382, 218), (410, 222), (446, 202)], 120)]
    for pts in strokes:
        th = tube(pts, 6.2, taper=lambda t: .78 + .22 * np.sin(np.pi * t))
        over(shade(bumped(th), ink, 40, .45, 1, wrap=.2, rim_amt=.1), layer_alpha(th, 1.2))
    # eyes
    for cx, cy in [(292, 142), (334, 140)]:
        eh, _ = dome(cx, cy, 14.5, 23, 16)
        over(shade(bumped(eh), (.06, .05, .12), 60, .9, 1, wrap=.2, rim_amt=.15), layer_alpha(eh, 1.2))
    # glossy red nose, a little tilted, with a soft contact shadow on the face
    nh, nd = dome(288, 204, 66, 50, 44, rot=-6)
    sh = np.exp(-np.clip(nd - 0.8, 0, None) * 3.5) * (nd < 1.9) * (nd > 0.95)
    sh = ndimage.gaussian_filter(((nd > 0.9) & (nd < 1.45)).astype(float) * (uy > 204), 6)
    canvas[..., :3] *= (1 - .18 * sh[..., None])
    over(shade(bumped(nh), (.93, .07, .12), 70, 1.0, 1, wrap=.3, rim_amt=.25, sss=.8), layer_alpha(nh, 1.2))

    # ---------- accessories by level
    if level >= 2:
        # red scarf: a soft fabric tube hugging the body, plus a hanging tail
        mid = bezier([(146, 362), (336, 448), (530, 362)], 240)
        sc = tube(mid, 36, taper=lambda t: .8 + .2 * np.sin(np.pi * t)) * .32
        sc = np.where(inside, sc, -1e3)
        over(shade(bumped(sc + 10), (.86, .09, .13), 18, .18, 1, wrap=.4, rim_amt=.2, sss=.5), layer_alpha(sc, 1.5))
        tail = bezier([(462, 396), (494, 430), (520, 470), (536, 506)], 140)
        tt = tube(tail, 30, taper=lambda t: 1 - .25 * t) * .34
        over(shade(normals(np.maximum(tt, 0) + 20), (.80, .08, .12), 18, .18, 1, wrap=.4, rim_amt=.2), layer_alpha(tt, 1.5))
    if level >= 3:
        # round glasses: thin dark rings with a faint lens tint
        for cx, cy in [(286, 144), (344, 142)]:
            lens = ((ux - cx) ** 2 + (uy - cy) ** 2) < 29 ** 2
            canvas[..., :3] = np.where(lens[..., None], canvas[..., :3] * .9 + np.array([.85, .9, 1]) * .1, canvas[..., :3])
            ang = np.linspace(0, 2 * np.pi, 220)
            ring = np.c_[cx + 31 * np.cos(ang), cy + 31 * np.sin(ang)]
            rh = tube(ring, 3.6)
            over(shade(bumped(rh + 6), (.08, .06, .14), 50, .7, 1, wrap=.2, rim_amt=.15), layer_alpha(rh, 1))
        for pts in [bezier([(312, 140), (315, 134), (318, 140)], 30), bezier([(255, 140), (240, 134), (226, 128)], 40), bezier([(375, 138), (390, 131), (404, 124)], 40)]:
            th = tube(pts, 3.4)
            over(shade(bumped(th + 6), (.08, .06, .14), 50, .7, 1, wrap=.2, rim_amt=.15), layer_alpha(th, 1))
    if level >= 4:
        # golden crown with bevelled edges, tip balls and a red gem
        poly = [(254, 52), (244, -44), (292, -2), (336, -66), (380, -2), (428, -44), (418, 52)]
        img = Image.new('L', (W, H), 0)
        ImageDraw.Draw(img).polygon([((x - X0) * S, (y - Y0) * S) for x, y in poly] + [((336 - X0) * S, (66 - Y0) * S)], fill=255)
        m = np.asarray(img) > 0
        dist = ndimage.distance_transform_edt(m) / S
        ch = np.where(m, np.minimum(dist, 9) * 1.6 + 6, -1e3)
        cn = normals(np.where(m, ch, 0))
        refl = np.clip(-cn[..., 1] * .6 + .5, 0, 1)
        gold = np.array([1.0, .77, .22])
        gcol = gold * (.25 + .85 * np.clip(cn @ L, 0, 1)[..., None]) * (.65 + .55 * refl[..., None]) + (np.clip(cn @ HV, 0, 1) ** 30)[..., None] * .9
        # contact shadow of the crown on the head
        cs = ndimage.gaussian_filter(m.astype(float), 7 * S)
        canvas[..., :3] *= (1 - .25 * np.roll(cs, int(10 * S), 0)[..., None] * inside[..., None])
        over(gcol, layer_alpha(np.where(m, ch, -1e3), 1))
        for cx, cy, r in [(244, -48, 12), (336, -72, 13), (428, -48, 12)]:
            bh2, _ = dome(cx, cy, r, r, r)
            bn = normals(np.maximum(bh2, 0))
            bcol = gold * (.3 + .85 * np.clip(bn @ L, 0, 1)[..., None]) + (np.clip(bn @ HV, 0, 1) ** 40)[..., None] * 1.0
            over(bcol, layer_alpha(bh2, 1))
        gh, _ = dome(336, 22, 13, 15, 12)
        over(shade(normals(np.maximum(gh, 0)), (.9, .06, .12), 80, 1.0, 1, wrap=.2, rim_amt=.2), layer_alpha(gh, 1))

    out = np.clip(canvas, 0, 1)
    im = Image.fromarray((out * 255).astype(np.uint8), 'RGBA')
    # soft ground shadow under the cloud
    shadow = Image.new('L', (W, H), 0)
    ImageDraw.Draw(shadow).ellipse([((150 - X0) * S, (596 - Y0) * S), ((620 - X0) * S, (636 - Y0) * S)], fill=120)
    shadow = shadow.filter(ImageFilter.GaussianBlur(14 * S))
    base = Image.new('RGBA', (W, H), (40, 20, 90, 0))
    base.putalpha(shadow)
    base.alpha_composite(im)
    if level >= 2:
        # three twinkling sparkles
        sp = Image.new('RGBA', (W, H), (0, 0, 0, 0))
        d = ImageDraw.Draw(sp)
        for cx, cy, r in [(52, 64, 30), (710, 76, 22), (744, 482, 18)]:
            pts = []
            for i in range(8):
                a = np.pi / 4 * i - np.pi / 2
                rr = r if i % 2 == 0 else r * .28
                pts.append((((cx + rr * np.cos(a)) - X0) * S, ((cy + rr * np.sin(a)) - Y0) * S))
            d.polygon(pts, fill=(253, 205, 70, 255))
        glow = sp.filter(ImageFilter.GaussianBlur(8))
        base.alpha_composite(glow)
        base.alpha_composite(sp)
    return base


if __name__ == '__main__':
    pre = sys.argv[1]
    levels = [int(a) for a in sys.argv[2:]] or [1, 2, 3, 4]
    for lv in levels:
        im = render(lv)
        im = im.crop(im.getbbox())
        im.save(f'{pre}-l{lv}.png')
        print(lv, im.size)
