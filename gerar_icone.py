"""
gerar_icone.py — Gera o icone.ico para o Sistema de Ponto
Não precisa de nenhuma lib externa; usa só tkinter + ctypes (já vêm com Python)
Execute: python gerar_icone.py
"""

import struct, zlib, os

# ── Cria um ICO em memória com 4 tamanhos: 16, 32, 48, 256 ──────────────────
# Paleta de cores do Sistema de Ponto
BG      = (30,  30,  46)   # fundo escuro (sidebar)
ACCENT  = (99,  102, 241)  # roxo/índigo
WHITE   = (255, 255, 255)
CLOCK   = (250, 189,  47)  # amarelo — ponteiro do relógio

def make_png(size):
    """Gera um PNG em bytes com o ícone desenhado pixel-a-pixel."""
    import struct, zlib

    w = h = size
    # Canvas RGBA
    img = [[(0, 0, 0, 0)] * w for _ in range(h)]

    def circle(cx, cy, r, color, alpha=255):
        for y in range(h):
            for x in range(w):
                dx, dy = x - cx, y - cy
                dist = (dx*dx + dy*dy) ** 0.5
                if dist <= r:
                    img[y][x] = (*color, alpha)
                elif dist <= r + 1.2:
                    a = int(255 * (r + 1.2 - dist) / 1.2)
                    img[y][x] = (*color, min(alpha, a))

    def rect(x1, y1, x2, y2, color, alpha=255):
        for y in range(max(0,y1), min(h,y2)):
            for x in range(max(0,x1), min(w,x2)):
                img[y][x] = (*color, alpha)

    cx, cy = w // 2, h // 2
    r  = int(w * 0.46)
    rw = int(w * 0.40)

    # Círculo de fundo (roxo)
    circle(cx, cy, r, ACCENT)
    # Círculo interno branco
    circle(cx, cy, rw, WHITE)
    # Face do relógio (branco)
    circle(cx, cy, int(rw * 0.92), BG)

    # Ponteiro hora (curto, branco, vertical-up)
    ph = int(rw * 0.42)
    pw = max(1, int(w * 0.04))
    rect(cx - pw, cy - ph, cx + pw, cy, WHITE)

    # Ponteiro minuto (longo, amarelo, diagonal ~2h)
    import math
    angle = math.radians(-60)  # 2 horas
    pm = int(rw * 0.60)
    ex = cx + int(pm * math.sin(angle))
    ey = cy - int(pm * math.cos(angle))
    # linha de Bresenham
    dx2, dy2 = ex - cx, ey - cy
    steps = max(abs(dx2), abs(dy2))
    for i in range(steps + 1):
        t = i / max(steps, 1)
        px = int(cx + dx2 * t)
        py = int(cy + dy2 * t)
        for oy in range(-max(1,pw-1), max(1,pw)):
            for ox in range(-max(1,pw-1), max(1,pw)):
                nx, ny = px+ox, py+oy
                if 0 <= nx < w and 0 <= ny < h:
                    img[ny][nx] = (*CLOCK, 255)

    # Centro do relógio
    circle(cx, cy, max(2, int(w * 0.06)), CLOCK)

    # ── Codifica como PNG ───────────────────────────────────
    def png_chunk(tag, data):
        c = zlib.crc32(tag + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", c)

    ihdr = struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0)  # 8-bit RGBA
    raw = b""
    for row in img:
        raw += b"\x00"  # filter type None
        for px in row:
            raw += bytes(px)
    idat = zlib.compress(raw, 9)

    png = (b"\x89PNG\r\n\x1a\n"
           + png_chunk(b"IHDR", ihdr)
           + png_chunk(b"IDAT", idat)
           + png_chunk(b"IEND", b""))
    return png


def make_ico(path="icone.ico"):
    sizes = [16, 32, 48, 256]
    pngs  = [make_png(s) for s in sizes]

    n = len(sizes)
    # ICO header: 6 bytes
    # Directory: n * 16 bytes
    # Images: concatenated
    header = struct.pack("<HHH", 0, 1, n)
    offset = 6 + n * 16

    directory = b""
    for i, s in enumerate(sizes):
        png = pngs[i]
        w = h = s if s < 256 else 0  # 256 → 0 no ICO
        directory += struct.pack("<BBBBHHII",
            w, h,      # width, height (0 = 256)
            0,         # color count (0 = truecolor)
            0,         # reserved
            1,         # color planes
            32,        # bits per pixel
            len(png),  # size of image data
            offset,    # offset of image data
        )
        offset += len(png)

    with open(path, "wb") as f:
        f.write(header + directory + b"".join(pngs))

    kb = os.path.getsize(path) // 1024
    print(f"[OK] {path} criado ({kb} KB, {len(sizes)} tamanhos: {sizes})")


if __name__ == "__main__":
    make_ico("icone.ico")
