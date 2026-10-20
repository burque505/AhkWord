# AhkWord's icon: a W whose strokes are also A, H and K. The W's four strokes are deep teal; the rest of the letters
# (the A's left leg and crossbar, the H's crossbar, the K's arms) warm orange, with thin deep orange edges outlining
# the W's two Vs; a teal bridge joins the W's middle peak. On a light, warm rounded square, so that it shows on light
# and dark backgrounds alike. From 32 pixels up it is the whole mark; at 16 to 24 pixels the W alone (nothing finer
# survives that small).
#
#   python make_icon.py            writes ahkword.ico (16 to 256 pixels), ahkword-256.png (the icon) and
#                                  ahkword-mark.png (the mark without its square: for light pages) here
#   python make_icon.py preview    also preview-sizes.png and preview-enlarged.png, to judge a change
#
# Needs Python 3 with numpy and matplotlib. Only to change the icon: AhkWord itself uses the files it writes.
import os
import struct
import sys
import io
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle

out = os.path.dirname(os.path.abspath(__file__))
preview = len(sys.argv) > 1 and sys.argv[1] == "preview"

INK = "#0E6670"      # deep teal: the strokes that make the W
ACCENT = "#E8833A"   # warm orange, teal's complement: the other strokes
EDGE = "#C25B12"     # deep orange: the thin edges of the W's outer strokes
TILE = "#FBF8F2"     # a light, warm off-white square (deep teal shows on it, on light and dark backgrounds alike)
TILE_EDGE = "#CFC6B5"
ON_TILE = INK
ACCENT_ON_TILE = ACCENT

BOTTOM, TOP = 0.20, 0.80

def mark(ax, ink, accent, detail=True, inset=0.0):
    # in a 0..1 square; inset shrinks it toward the centre (for the tile)
    def P(x, y):
        return (0.5 + (x - 0.5) * (1 - 2 * inset), 0.5 + (y - 0.5) * (1 - 2 * inset))
    width = 0.085 * (1 - 2 * inset)
    thin = width * 0.55
    clip = Rectangle(P(-0.2, BOTTOM), 1.4, (TOP - BOTTOM) * (1 - 2 * inset), transform=ax.transAxes, visible=False)
    ax.add_patch(clip)
    def line(points, w, color, join="miter"):
        # each stroke drawn beyond the top and bottom, then cut level there
        xs, ys = zip(*[P(*p) for p in points])
        art, = ax.plot(xs, ys, color=color, linewidth=w, solid_joinstyle=join, solid_capstyle="butt", transform=ax.transAxes)
        art.set_clip_path(clip)
        return art
    s = ax.figure.dpi * ax.figure.get_size_inches()[0]   # points per unit (of the axes' width)
    lw = width * s * 72 / ax.figure.dpi
    tw = thin * s * 72 / ax.figure.dpi
    def extend(a, b, by=0.15):
        (x1, y1), (x2, y2) = a, b
        dx, dy = x2 - x1, y2 - y1
        n = (dx * dx + dy * dy) ** 0.5
        return (x1 - dx / n * by, y1 - dy / n * by), (x2 + dx / n * by, y2 + dy / n * by)
    if not detail:
        # the small sizes: the W alone, bolder (its strokes the mark's: the A's right leg, the H, the K's upright)
        bold = lw * 1.35
        points = [(0.20, TOP), (0.345, BOTTOM), (0.51, TOP), (0.675, BOTTOM), (0.82, TOP)]
        first = extend(points[1], points[0])[1]
        last = extend(points[-2], points[-1])[1]
        line([first] + points[1:-1] + [last], bold, ink)
        return
    # A: its apex the W's top-left; its right leg the W's first stroke, its left leg outside it
    apex, a_left, a_right = (0.20, TOP + 0.02), (0.07, BOTTOM), (0.33, BOTTOM)
    # H: narrow at the top, wide at the bottom: the W's middle strokes
    h1b, h1t, h2t, h2b = (0.33, BOTTOM), (0.425, TOP), (0.595, TOP), (0.69, BOTTOM)
    # K: its upright the W's right leg, leaning right at the top; its arms thinner
    kb, kt = (0.69, BOTTOM), (0.80, TOP)
    def at(a, b, y):
        return a[0] + (b[0] - a[0]) * (y - a[1]) / (b[1] - a[1])
    # first the accent's strokes (the A's left leg, the crossbars, the K's arms), then the W's over them, unbroken
    line(list(extend(apex, a_left)), lw * 0.85, accent)
    ya, yh, yk = 0.40, 0.52, 0.48
    line([(at(apex, a_left, ya), ya), (at(apex, a_right, ya), ya)], tw * 1.15, accent)
    line([(at(h1b, h1t, yh), yh), (at(h2t, h2b, yh), yh)], lw * 0.8, accent)    # (the H's: heavier)
    j = (at(kb, kt, yk), yk)
    line([j, extend(j, (0.95, TOP), 0.12)[1]], tw * 1.4, accent)    # (the K's arms: a little heavier)
    line([j, extend(j, (0.95, BOTTOM), 0.12)[1]], tw * 1.4, accent)
    # the W's outer strokes (the A's right leg, the K's upright) edged both sides in a very thin deep orange: drawn a
    # little wider in it first, the teal over them, so that the A and the K show whole
    # (each edge a thin line along the stroke's side: the outer ones the whole length; the inner ones, on the side of
    # the V the stroke makes with the H's upright beside it, from the top down to that V's notch, where the two
    # strokes' inner sides meet)
    w = 0.085                      # the strokes' width, in the mark's units
    e = w * 0.12                   # the edges'
    ew = lw * 0.12                 # (in points)
    def offset(a, b, d):
        # the line a-b moved d to its right (toward +x)
        (x1, y1), (x2, y2) = a, b
        dx, dy = x2 - x1, y2 - y1
        n = (dx * dx + dy * dy) ** 0.5
        nx, ny = dy / n, -dx / n
        if nx < 0: nx, ny = -nx, -ny
        return (x1 + nx * d, y1 + ny * d), (x2 + nx * d, y2 + ny * d)
    def meet(p, q):
        # where two lines (each a pair of points) cross
        (x1, y1), (x2, y2) = p; (x3, y3), (x4, y4) = q
        den = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
        t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / den
        return (x1 + t * (x2 - x1), y1 + t * (y2 - y1))
    def strip(a, b, d, upto=None, width=None):
        # an edge along a-b, d to its right; upto: from the top down to that height only
        s, t = offset(a, b, d)
        top, bottom = (s, t) if s[1] > t[1] else (t, s)
        top = extend(bottom, top)[1]
        if upto is not None:
            bottom = (at(top, bottom, upto), upto)
        else:
            bottom = extend(top, bottom)[1]
        line([top, bottom], width or ew, EDGE)
    line(list(extend(h1b, h1t)), lw, ink)
    line(list(extend(h2t, h2b)), lw, ink)
    line(list(extend(apex, a_right)), lw, ink)
    line(list(extend(kb, kt)), lw, ink)
    # the A's right leg: its left edge whole, its right edge down to the notch it makes with the H's left upright
    notch_a = meet(offset(apex, a_right, w / 2), offset(h1b, h1t, -w / 2))
    strip(apex, a_right, -(w / 2 + e / 2))
    strip(apex, a_right, w / 2 + e / 2, upto=notch_a[1])
    # the K's upright: its right edge whole, its left edge down to the notch it makes with the H's right upright
    notch_k = meet(offset(kb, kt, -w / 2), offset(h2t, h2b, w / 2))
    strip(kb, kt, w / 2 + e / 2)
    strip(kb, kt, -(w / 2 + e / 2), upto=notch_k[1])
    # the inner sides of the Vs (the H's uprights' sides that face each other), the whole length
    he, hew = e * 2.5, ew * 2.5    # (the H's edges: thicker)
    strip(h1b, h1t, w / 2 + he / 2, width=hew)
    strip(h2t, h2b, -(w / 2 + he / 2), width=hew)
    # along the bottom of each V (the A's leg with the H's left upright, the H's right upright with the K's): a thin
    # edge just above the baseline, from the V's outside on one side to its outside on the other
    y = BOTTOM + e / 2
    def side(a, b, d):
        s, t = offset(a, b, d)
        return at(s, t, y)
    line([(side(apex, a_right, -(w / 2 + e)), y), (side(h1b, h1t, w / 2 + he), y)], ew, EDGE)
    line([(side(h2t, h2b, -(w / 2 + he)), y), (side(kb, kt, w / 2 + e), y)], ew, EDGE)
    # along the top of each V: across the top of each of its strokes, from edge to edge (the H's uprights' notch sides
    # have no edge: from the stroke's own side there)
    y = TOP - e / 2
    line([(side(apex, a_right, -(w / 2 + e)), y), (side(apex, a_right, w / 2 + e), y)], ew, EDGE)
    # a teal bridge across the top of the H's gap, joining its uprights: the W's middle peak one shape, the W a W and
    # not two Vs; the H's thick top edge runs on across it
    # (only between the legs: from the outside of one leg's orange inner edge to the other's, its top level with the
    # legs' teal tops; a shape of four corners, so that it meets the slanting sides exactly)
    bridge = w * 0.75
    def x_on(a, b, d, y):
        s, t = offset(a, b, d)
        return at(s, t, y)
    y1 = TOP - he    # (level with the teal tops of the legs, below their orange top edges)
    y0 = y1 - bridge
    corners = [(x_on(h1b, h1t, w / 2 + he, y1), y1), (x_on(h2t, h2b, -(w / 2 + he), y1), y1),
               (x_on(h2t, h2b, -(w / 2 + he), y0), y0), (x_on(h1b, h1t, w / 2 + he, y0), y0)]
    xs, ys = zip(*[P(*c) for c in corners])
    ax.fill(xs, ys, color=ink, linewidth=0, transform=ax.transAxes)
    # each leg's own orange top edge
    yt = TOP - he / 2
    line([(side(h1b, h1t, -w / 2), yt), (side(h1b, h1t, w / 2 + he), yt)], hew, EDGE)
    line([(side(h2t, h2b, -(w / 2 + he)), yt), (side(h2t, h2b, w / 2), yt)], hew, EDGE)
    line([(side(kb, kt, -(w / 2 + e)), y), (side(kb, kt, w / 2 + e), y)], ew, EDGE)

def render(px, tile, detail):
    fig = plt.figure(figsize=(px / 100, px / 100), dpi=100)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
    fig.patch.set_alpha(0)
    if tile:
        ax.add_patch(FancyBboxPatch((0.04, 0.04), 0.92, 0.92, boxstyle="round,pad=0,rounding_size=0.2", transform=ax.transAxes, facecolor=TILE, edgecolor=TILE_EDGE, linewidth=max(0.6, px / 128)))
        mark(ax, ON_TILE, ACCENT_ON_TILE, detail, inset=0.10)
    else:
        mark(ax, INK, ACCENT, detail)
    fig.canvas.draw()
    img = np.asarray(fig.canvas.buffer_rgba()).copy()
    plt.close(fig)
    return img

def over(img, bg):
    a = img[..., 3:4] / 255.0
    rgb = img[..., :3] * a + np.array(bg)[None, None, :] * (1 - a)
    return rgb.astype(np.uint8)

SMALL = 24    # (at this size and below: the W alone)

def icon(px):
    return render(px, True, px > SMALL)

def png_bytes(img):
    buffer = io.BytesIO()
    plt.imsave(buffer, img, format="png")
    return buffer.getvalue()

def dib(img):
    # a 32-bit bitmap as an .ico holds one: BITMAPINFOHEADER, then BGRA rows bottom up, then an all-clear AND mask
    h, w = img.shape[:2]
    header = struct.pack("<IiiHHIIiiII", 40, w, h * 2, 1, 32, 0, 0, 0, 0, 0, 0)
    pixels = img[::-1, :, [2, 1, 0, 3]].tobytes()
    mask = b"\0" * (((w + 31) // 32) * 4 * h)
    return header + pixels + mask

def write_ico(path, sizes):
    # small sizes as bitmaps (what every Windows reads), 256 as PNG (as Windows' own icons are)
    images = [(px, png_bytes(icon(px)) if px >= 256 else dib(icon(px))) for px in sizes]
    data = struct.pack("<HHH", 0, 1, len(images))
    at = 6 + 16 * len(images)
    for px, blob in images:
        data += struct.pack("<BBBBHHII", px % 256, px % 256, 0, 0, 1, 32, len(blob), at)
        at += len(blob)
    with open(path, "wb") as f:
        f.write(data + b"".join(blob for _, blob in images))

write_ico(os.path.join(out, "ahkword.ico"), [16, 20, 24, 32, 40, 48, 64, 96, 128, 256])
plt.imsave(os.path.join(out, "ahkword-256.png"), icon(256))
plt.imsave(os.path.join(out, "ahkword-mark.png"), render(512, False, True))
if not preview:
    print("ok")
    sys.exit(0)
sizes = [16, 20, 24, 32, 48, 64, 128, 256]
backs = [(255, 255, 255), (243, 243, 243), (32, 32, 32)]
rows = [("the mark alone", False, True), ("on a rounded square", True, True), ("small sizes: the W alone", True, False)]
gap = 24
W = sum(sizes) + gap * (len(sizes) + 1)
H = (max(sizes) + gap) * len(rows) * len(backs) + gap
sheet = np.full((H, W, 3), 200, np.uint8)
y = gap
for _, tile, detail in rows:
    for bg in backs:
        x = gap
        sheet[y - gap // 2:y + max(sizes) + gap // 2, :] = bg
        for px in sizes:
            img = over(render(px, tile, detail), bg)
            sheet[y + (max(sizes) - px):y + max(sizes), x:x + px] = img
            x += px + gap
        y += max(sizes) + gap
plt.imsave(os.path.join(out, "preview-sizes.png"), sheet)

# enlarged: 16, 24 and 32 pixel versions shown 8x (each pixel a block), and the 256 one
big = []
for tile, detail in [(False, True), (True, True), (True, False)]:
    row = []
    for px in [16, 24, 32]:
        img = over(render(px, tile, detail), (255, 255, 255))
        row.append(np.kron(img, np.ones((8, 8, 1), np.uint8)))
        row.append(np.full((px * 8, 16, 3), 255, np.uint8))
    hmax = max(r.shape[0] for r in row)
    row = [np.vstack([r, np.full((hmax - r.shape[0], r.shape[1], 3), 255, np.uint8)]) for r in row]
    big.append(np.hstack(row))
    big.append(np.full((16, big[-1].shape[1], 3), 255, np.uint8))
plt.imsave(os.path.join(out, "preview-enlarged.png"), np.vstack(big))
print("ok")
