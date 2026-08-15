# -*- coding: utf-8 -*-
"""dataviz 스킬 validate_palette.js 의 Python 재현 (핵심 게이트만)
- 정상시각 OKLab ΔE×100 ≥ 15 (hard floor)
- CVD(deuteranopia·protanopia, Viénot 1999) ΔE×100 ≥ 8 (target)
- 흰 표면 대비(contrast) 보고 (3:1 미만 = relief rule: 직접 라벨 의무)
"""
import sys, itertools

def hex2rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))

def srgb2lin(c):
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4

def lin2srgb(c):
    c = max(0.0, min(1.0, c))
    return 12.92 * c if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055

def oklab(rgb):
    r, g, b = (srgb2lin(c) for c in rgb)
    l = 0.4122214708*r + 0.5363325363*g + 0.0514459929*b
    m = 0.2119034982*r + 0.6806995451*g + 0.1073969566*b
    s = 0.0883024619*r + 0.2817188376*g + 0.6299787005*b
    l, m, s = l**(1/3), m**(1/3), s**(1/3)
    return (0.2104542553*l + 0.7936177850*m - 0.0040720468*s,
            1.9779984951*l - 2.4285922050*m + 0.4505937099*s,
            0.0259040371*l + 0.7827717662*m - 0.8086757660*s)

def dE(a, b):
    A, B = oklab(a), oklab(b)
    return 100 * sum((x - y) ** 2 for x, y in zip(A, B)) ** 0.5

# Viénot/Brettel 1999 dichromat simulation (linear RGB matrix)
DEUT = [[0.625, 0.375, 0.0], [0.7, 0.3, 0.0], [0.0, 0.3, 0.7]]
PROT = [[0.567, 0.433, 0.0], [0.558, 0.442, 0.0], [0.0, 0.242, 0.758]]

def cvd(rgb, M):
    lin = [srgb2lin(c) for c in rgb]
    out = [sum(M[i][j] * lin[j] for j in range(3)) for i in range(3)]
    return tuple(lin2srgb(c) for c in out)

def luminance(rgb):
    r, g, b = (srgb2lin(c) for c in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b

def contrast(fg, bg):
    L1, L2 = sorted((luminance(fg), luminance(bg)), reverse=True)
    return (L1 + 0.05) / (L2 + 0.05)

def check(names_hexes, surface="#ffffff", pairs="all"):
    cols = [(n, hex2rgb(h), h) for n, h in names_hexes]
    surf = hex2rgb(surface)
    ok = True
    print(f"{'pair':34s} {'normal':>7s} {'deuter':>7s} {'protan':>7s}")
    idx = (itertools.combinations(range(len(cols)), 2) if pairs == "all"
           else [(i, i + 1) for i in range(len(cols) - 1)])
    for i, j in idx:
        (n1, c1, _), (n2, c2, _) = cols[i], cols[j]
        d_n = dE(c1, c2)
        d_d = dE(cvd(c1, DEUT), cvd(c2, DEUT))
        d_p = dE(cvd(c1, PROT), cvd(c2, PROT))
        flag = ""
        if d_n < 15: flag += "  ✗ normal<15"; ok = False
        if min(d_d, d_p) < 8: flag += "  ⚠ CVD<8"
        print(f"{n1+' vs '+n2:34s} {d_n:7.1f} {d_d:7.1f} {d_p:7.1f}{flag}")
    print()
    for n, c, h in cols:
        cr = contrast(c, surf)
        note = "" if cr >= 3 else "  (relief rule: 직접 라벨 의무)"
        print(f"  {n:14s} {h}  contrast {cr:4.2f}:1{note}")
    return ok

if __name__ == "__main__":
    print("═══ 후보 A: 카테고리 3색 (slate blue · terracotta · warm gray) ═══")
    check([("blue", "#34608D"), ("terracotta", "#B4593F"), ("gray", "#A9A29A")])
    print()
    print("═══ 순서형 램프(세대·등급 3단) 인접쌍 ═══")
    check([("light", "#A9C2DC"), ("mid", "#5B84AC"), ("dark", "#2C5379")], pairs="adj")
    print()
    print("═══ 대안 램프(진하게) ═══")
    check([("light", "#93B2D1"), ("mid", "#4E7AA6"), ("dark", "#24476B")], pairs="adj")
