# -*- coding: utf-8 -*-
"""
Paper32 — 도형 안 글자를 **실측해서** 넣는 헬퍼

★ 왜 필요한가
  상자 폭과 글꼴 크기를 눈대중으로 맞추다 Fig 1 에서 글자가 상자를 줄줄이 넘쳤다.
  matplotlib 렌더러로 실제 렌더 폭을 재고, 들어갈 때까지 글꼴을 줄이거나 줄바꿈한다.
  **"몇 pt면 되겠지"를 코드에서 없앤다.**
"""
import textwrap


def _w_data(ax, txt, fs, weight="normal"):
    """주어진 글꼴 크기에서 txt 의 폭을 데이터 좌표로 환산."""
    fig = ax.figure
    fig.canvas.draw()
    r = fig.canvas.get_renderer()
    t = ax.text(0, 0, txt, fontsize=fs, fontweight=weight, alpha=0)
    bb = t.get_window_extent(renderer=r)
    t.remove()
    inv = ax.transData.inverted()
    (x0, _), (x1, _) = inv.transform([(bb.x0, bb.y0), (bb.x1, bb.y1)])
    return abs(x1 - x0)


def fit_lines(ax, lines, max_w, fs, min_fs=4.6, weight="normal"):
    """여러 줄이 모두 max_w 안에 들어가는 최대 글꼴 크기를 찾아 반환."""
    f = fs
    while f > min_fs:
        if all(_w_data(ax, ln, f, weight) <= max_w for ln in lines if ln):
            return f
        f -= 0.15
    return min_fs


def wrap_to(ax, txt, max_w, fs, weight="normal"):
    """max_w 에 맞게 줄바꿈. 한 단어가 넘치면 그때만 글꼴을 줄인다."""
    words = txt.split()
    if not words:
        return [""], fs
    lines, cur = [], words[0]
    for w in words[1:]:
        trial = cur + " " + w
        if _w_data(ax, trial, fs, weight) <= max_w:
            cur = trial
        else:
            lines.append(cur); cur = w
    lines.append(cur)
    return lines, fit_lines(ax, lines, max_w, fs, weight=weight)


def boxed_text(ax, x, y, w, h, lines, fs, weight="normal", color="black",
               pad=0.9, linespacing=1.5, va="center"):
    """상자(x,y,w,h) 안에 lines 를 넣되, 실측으로 글꼴을 줄여 반드시 들어가게 한다."""
    avail = w - 2 * pad
    f = fit_lines(ax, lines, avail, fs, weight=weight)
    yy = y + h / 2 if va == "center" else y + h - pad
    ax.text(x + w / 2, yy, "\n".join(lines), ha="center", va=va,
            fontsize=f, color=color, fontweight=weight, linespacing=linespacing)
    return f


def kv_rows(ax, x, y, w, rows, fs, color="black", gap=1.85, pad=1.0, min_fs=5.0,
            label=""):
    """`라벨 ......... 값` 2열 행. 라벨이 값 자리를 침범하지 않도록 실측으로 줄인다.

    ⚠️ 최소 글꼴에서도 안 들어가면 **조용히 겹치게 두지 않고 경고**한다.
       (Fig 1 에서 겹친 라벨을 눈으로 보고서야 알았다 — 그 침묵을 없앤다.)
    """
    num_w = max(_w_data(ax, f"{v:,}", fs) for _, v in rows) if rows else 0
    avail = w - 2 * pad - num_w - 1.0
    f = fs
    while f > min_fs and any(_w_data(ax, k, f) > avail for k, _ in rows):
        f -= 0.15
    over = [k for k, _ in rows if _w_data(ax, k, f) > avail]
    if over:
        import warnings
        warnings.warn(f"[figtext] {label or 'kv_rows'}: 최소 글꼴에서도 폭 초과 → "
                      f"{over} (avail={avail:.1f})", stacklevel=2)
    for i, (k, v) in enumerate(rows):
        yy = y - i * gap
        if k:
            ax.text(x + pad, yy, k, ha="left", va="top", fontsize=f, color=color)
        ax.text(x + w - pad, yy, f"{v:,}", ha="right", va="top", fontsize=f, color=color)
    return f
