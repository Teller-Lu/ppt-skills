#!/usr/bin/env python3
"""
Aesthetic Quality Gate (Stage 5) — static SVG analyzer.

Two halves (see docs/美学质量门规格.md):
  1. DEFECT checks   — font floor, font-size variety, palette adherence,
                       safe-margin overflow, banned-construct scan.
  2. RICHNESS gate   — the "非朴素门槛": scores icons / structural motif /
                       layered depth; a page that is TOO PLAIN FAILs, exactly
                       like an overflow FAILs. This is what stops the gate
                       from being a mere defect-catcher and makes it ENFORCE
                       the look learned from the real template renders.

Thresholds come from the style's gate.calibration.json (calibrated to the
template), not a generic ruler. Stdlib only — no new dependencies.

Usage:
    python3 aesthetic_checker.py <svg...> --calibration <gate.calibration.json>
    python3 aesthetic_checker.py components/*.svg --calibration styles/academic-defense/gate.calibration.json
    python3 aesthetic_checker.py page.svg --calibration cal.json --role content --json
"""
import argparse
import json
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

SVG_NS = "http://www.w3.org/2000/svg"


# ---------------------------------------------------------------- parsing utils
def _local(tag: str) -> str:
    return tag.split("}", 1)[1] if "}" in tag else tag


def _style_dict(el) -> dict:
    s = el.get("style")
    if not s:
        return {}
    out = {}
    for part in s.split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip()] = v.strip()
    return out


def _prop(el, name):
    """attribute OR inline-style property."""
    v = el.get(name)
    if v is not None:
        return v
    return _style_dict(el).get(name)


def _num(v):
    if v is None:
        return None
    m = re.match(r"^\s*(-?\d+(?:\.\d+)?)", str(v))
    return float(m.group(1)) if m else None


def build_parent_map(root):
    return {c: p for p in root.iter() for c in p}


def ancestor_ids(el, parents):
    ids = []
    cur = el
    while cur is not None:
        i = cur.get("id")
        if i:
            ids.append(i.lower())
        cur = parents.get(cur)
    return ids


DECOR_RE = re.compile(r"(bg|background|decor|atm|atmosphere|ghost|watermark|halo|orb)")


def is_decoration(el, parents):
    own = (el.get("id") or "").lower()
    if DECOR_RE.search(own):
        return True
    return any(DECOR_RE.search(i) for i in ancestor_ids(el, parents))


# ---------------------------------------------------------------- color helpers
HEX_RE = re.compile(r"#([0-9a-fA-F]{6}|[0-9a-fA-F]{3})")


def hex_to_rgb(h):
    h = h.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def color_dist(c1, c2):
    """Redmean weighted RGB distance — good cheap perceptual approximation."""
    r1, g1, b1 = c1
    r2, g2, b2 = c2
    rm = (r1 + r2) / 2
    dr, dg, db = r1 - r2, g1 - g2, b1 - b2
    return ((2 + rm / 256) * dr * dr + 4 * dg * dg + (2 + (255 - rm) / 256) * db * db) ** 0.5


def is_gray(rgb, tol=14):
    return max(rgb) - min(rgb) <= tol


# ---------------------------------------------------------------- gradient/filter index
def index_defs(root):
    grads, filters = set(), set()
    for el in root.iter():
        t = _local(el.tag)
        if t in ("linearGradient", "radialGradient"):
            if el.get("id"):
                grads.add(el.get("id"))
        elif t == "filter":
            if el.get("id"):
                filters.add(el.get("id"))
    return grads, filters


URL_RE = re.compile(r"url\(#([^)]+)\)")


def fill_is_gradient(el, grads):
    f = _prop(el, "fill")
    if not f:
        return False
    m = URL_RE.search(f)
    return bool(m and m.group(1) in grads)


# ================================================================ main analysis
def analyze(path: Path, cal: dict, role: str):
    errors, warnings = [], []
    raw = path.read_text(encoding="utf-8", errors="ignore")
    # Security: SVG files never need a DOCTYPE/entity declaration. Rejecting them
    # neutralizes XXE + billion-laughs without depending on a non-stdlib parser
    # (defusedxml). stdlib ET does not resolve *external* entities by default;
    # this guard closes the *internal* entity-expansion vector too.
    if re.search(r"<!DOCTYPE", raw, re.I) or re.search(r"<!ENTITY", raw, re.I):
        return {"file": path.name, "role": role, "verdict": "FAIL",
                "errors": ["SECURITY: DOCTYPE/ENTITY declaration not allowed in SVG"],
                "warnings": [], "richness": {}}
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as e:
        return {"file": path.name, "role": role, "verdict": "FAIL",
                "errors": [f"XML parse error: {e}"], "warnings": [], "richness": {}}

    parents = build_parent_map(root)
    grads, filters = index_defs(root)

    canvas = cal.get("canvas_px", [1280, 720])
    CW, CH = canvas[0], canvas[1]
    m = cal.get("safe_margin_px", {"left": 70, "right": 70, "top": 48, "bottom": 48})
    floor = cal.get("font_floor_px", 10.7)
    variety_max = cal.get("font_size_variety_max", 9)
    tol = cal.get("palette_tolerance", 26)
    palette = []
    for group in cal.get("palette", {}).values():
        for c in group:
            palette.append(hex_to_rgb(c))

    # ---- collect elements
    texts, shapes, images, icons = [], [], [], []
    font_sizes = set()
    circles, polygons = [], []
    for el in root.iter():
        t = _local(el.tag)
        if t == "text":
            texts.append(el)
            fs = _num(_prop(el, "font-size"))
            if fs:
                font_sizes.add(round(fs, 1))
        elif t in ("rect", "circle", "ellipse", "polygon", "path", "line"):
            shapes.append(el)
            if t == "circle":
                circles.append(el)
            if t == "polygon":
                polygons.append(el)
        elif t == "image":
            images.append(el)
        elif t == "use" and el.get("data-icon"):
            icons.append(el)
    # also count already-embedded icons (post-finalize) by data-icon marker on groups
    embedded_icons = len(re.findall(r'data-icon=', raw))

    # ===================================================== DEFECT CHECKS
    # C1 font floor
    for el in texts:
        fs = _num(_prop(el, "font-size"))
        if fs is not None and fs < floor - 0.05:
            snippet = (el.text or "").strip()[:18]
            errors.append(f"FONT_FLOOR: {fs}px < {floor}px  «{snippet}»")
    # C2 font-size variety
    if len(font_sizes) > variety_max:
        warnings.append(f"FONT_VARIETY: {len(font_sizes)} distinct sizes > {variety_max} "
                        f"({sorted(font_sizes, reverse=True)})")
    # C2b headline floor — every page needs at least one prominent (big) text,
    # else it reads as the "AI-aesthetic" all-small page. (warn)
    headline_min = cal.get("headline_min_px")
    if headline_min and font_sizes and max(font_sizes) < headline_min:
        warnings.append(f"HEADLINE_SMALL: 最大字号 {max(font_sizes)}px < 标题下限 {headline_min}px "
                        f"(无足够显眼的标题/主元素)")
    # C3 palette adherence (skip decorations & grays & gradient refs)
    off_palette = {}
    for el in shapes + texts:
        if is_decoration(el, parents):
            continue
        for attr in ("fill", "stroke"):
            v = _prop(el, attr)
            if not v or v in ("none", "transparent") or v.startswith("url("):
                continue
            mm = HEX_RE.search(v)
            if not mm:
                continue
            rgb = hex_to_rgb(mm.group(0))
            if is_gray(rgb):
                continue
            if palette and min(color_dist(rgb, p) for p in palette) > tol:
                off_palette[mm.group(0).upper()] = off_palette.get(mm.group(0).upper(), 0) + 1
    for col, n in off_palette.items():
        warnings.append(f"PALETTE: {col} off-palette (x{n})")
    # C4 safe-margin overflow (rect/image with explicit geom, non-decoration)
    for el in shapes + images:
        if is_decoration(el, parents):
            continue
        if el.get("data-bleed") == "true":
            continue
        t = _local(el.tag)
        x, y = _num(el.get("x")), _num(el.get("y"))
        w, h = _num(el.get("width")), _num(el.get("height"))
        if None in (x, y, w, h):
            continue
        # Intentional full-bleed: background fills and edge-to-edge bands/fades
        # span the whole width or height. These are not content — exempt them.
        if (x <= 1 and (x + w) >= CW - 1) or (y <= 1 and (y + h) >= CH - 1):
            continue
        if x < m["left"] - 1 or y < m["top"] - 1 or (x + w) > CW - m["right"] + 1 or (y + h) > CH - m["bottom"] + 1:
            errors.append(f"SAFE_MARGIN: <{t}> bbox ({x:.0f},{y:.0f},{w:.0f}x{h:.0f}) crosses safe area")
    # C5 banned constructs (engine-unsafe → would break or flatten on export)
    if re.search(r"<\s*mask\b", raw):
        errors.append("CONSTRUCT: <mask> banned (use stacked rect/gradient overlay)")
    if re.search(r"<\s*g[^>]*\sopacity\s*=", raw):
        errors.append("CONSTRUCT: group opacity banned (set fill/stroke-opacity per child)")
    if "rgba(" in raw:
        errors.append("CONSTRUCT: rgba() banned (use fill-opacity/stroke-opacity)")
    if re.search(r"<\s*image[^>]*\sopacity\s*=", raw):
        errors.append("CONSTRUCT: <image opacity> banned (overlay a rect instead)")
    if "textPath" in raw:
        errors.append("CONSTRUCT: textPath banned (bake curved text to PNG)")
    for el in texts:  # gradient text → flattened to solid on export
        f = _prop(el, "fill") or ""
        mm = URL_RE.search(f)
        if mm and mm.group(1) in grads:
            warnings.append(f"GRADIENT_TEXT: «{(el.text or '').strip()[:18]}» fill is a gradient "
                            f"(flattens to solid on export — use solid or bake)")

    # C6 page-fill — content must not timidly hug the center leaving big empty
    # margins (the "AI-aesthetic" 缩中央留大白 problem). warn-only, est. bbox.
    fw_ratio = cal.get("min_fill_w_ratio")
    fh_ratio = cal.get("min_fill_h_ratio")
    if fw_ratio or fh_ratio:
        xs0, xs1, ys0, ys1 = [], [], [], []
        def _ext(a, b, c, d):
            xs0.append(a); xs1.append(b); ys0.append(c); ys1.append(d)
        for el in shapes + images + icons + texts:
            if is_decoration(el, parents):
                continue
            t = _local(el.tag)
            if t in ("rect", "image", "use"):
                x, y = _num(el.get("x")), _num(el.get("y"))
                w, h = _num(el.get("width")), _num(el.get("height"))
                if None in (x, y, w, h):
                    continue
                if (x <= 1 and x + w >= CW - 1) or (y <= 1 and y + h >= CH - 1):
                    continue  # full-bleed bg
                _ext(x, x + w, y, y + h)
            elif t == "circle":
                cx, cy, r = _num(el.get("cx")), _num(el.get("cy")), _num(el.get("r"))
                if None not in (cx, cy, r):
                    _ext(cx - r, cx + r, cy - r, cy + r)
            elif t == "polygon":
                pts = re.findall(r"(-?\d+\.?\d*)[ ,]+(-?\d+\.?\d*)", el.get("points") or "")
                if pts:
                    px = [float(p[0]) for p in pts]; py = [float(p[1]) for p in pts]
                    _ext(min(px), max(px), min(py), max(py))
            elif t == "text":
                tx, ty = _num(el.get("x")), _num(el.get("y"))
                fs = _num(_prop(el, "font-size")) or 16
                if tx is None or ty is None:
                    continue
                wdt = len("".join(el.itertext()).strip()) * fs * 0.95
                anc = _prop(el, "text-anchor") or "start"
                lo = tx - wdt if anc == "end" else (tx - wdt / 2 if anc == "middle" else tx)
                _ext(lo, lo + wdt, ty - fs, ty)
        if xs0:
            minx, maxx, miny, maxy = min(xs0), max(xs1), min(ys0), max(ys1)
            cw, ch = maxx - minx, maxy - miny
            gap = 80  # only "timid" if narrow/short AND empty on BOTH opposing sides
            l_gap, r_gap = minx - m["left"], (CW - m["right"]) - maxx
            t_gap, b_gap = miny - m["top"], (CH - m["bottom"]) - maxy
            if fw_ratio and cw < fw_ratio * CW and l_gap > gap and r_gap > gap:
                warnings.append(f"PAGE_FILL: 内容横向仅 {cw / CW * 100:.0f}% 且居中(左右各空 "
                                f"{l_gap:.0f}/{r_gap:.0f}px)→ 放大铺开,别缩中央")
            if fh_ratio and ch < fh_ratio * CH and t_gap > gap and b_gap > gap:
                warnings.append(f"PAGE_FILL: 内容纵向仅 {ch / CH * 100:.0f}% 且居中(上下各空 "
                                f"{t_gap:.0f}/{b_gap:.0f}px)→ 放大铺开,别缩中央")

    # ===================================================== RICHNESS GATE (non-plain)
    signals, sig_detail = {}, {}

    # S1 background gradient (a large rect filled with a gradient)
    bg_grad = False
    bg_image = any((_num(im.get("width")) or 0) >= CW * 0.85 and (_num(im.get("height")) or 0) >= CH * 0.85
                   for im in images)
    for el in shapes:
        if _local(el.tag) != "rect":
            continue
        w, h = _num(el.get("width")) or 0, _num(el.get("height")) or 0
        if w >= CW * 0.85 and h >= CH * 0.85 and fill_is_gradient(el, grads):
            bg_grad = True
    signals["bg_visual"] = bg_grad or bg_image
    sig_detail["bg_visual"] = "gradient" if bg_grad else ("image" if bg_image else "MISSING")

    # S2 decoration layer (atmosphere radial circle OR low-opacity ghost stroke)
    decoration = False
    for el in circles + polygons:
        if fill_is_gradient(el, grads) and is_decoration(el, parents):
            decoration = True
        so = _num(_prop(el, "stroke-opacity"))
        if _prop(el, "fill") in ("none", None) and so is not None and so <= 0.35:
            decoration = True
    signals["decoration_layer"] = decoration

    # S3 gradient-filled foreground shapes
    grad_shapes = sum(1 for el in shapes
                      if fill_is_gradient(el, grads) and not is_decoration(el, parents))
    signals["gradient_shapes"] = grad_shapes

    # S4 real icons
    icon_count = max(len(icons), embedded_icons)
    signals["icons"] = icon_count

    # S5 depth (a blur-based shadow/glow filter that is actually referenced)
    has_shadow_filter = any("feGaussianBlur" in raw and fid in raw for fid in filters)
    used_filter = bool(re.search(r'filter\s*=\s*["\']url\(#', raw)) and "feGaussianBlur" in raw
    signals["depth"] = used_filter or has_shadow_filter

    # S6 accent / motif hint (accent bars, hexagons, or multi-circle rings)
    accent_bar = any(_local(el.tag) == "rect" and (_num(el.get("rx")) or 0) > 0
                     and (_num(el.get("width")) or 0) <= 220 and (_num(el.get("height")) or 0) <= 14
                     for el in shapes)
    motif = len(polygons) >= 1 or len(circles) >= 3
    signals["accent_or_motif"] = accent_bar or motif

    # ---- score
    score = (
        (2 if signals["bg_visual"] else 0)
        + (2 if signals["decoration_layer"] else 0)
        + min(grad_shapes, 3)
        + min(icon_count, 3)
        + (1 if signals["depth"] else 0)
        + (1 if signals["accent_or_motif"] else 0)
    )

    # ---- per-role requirements (calibrated to the template's real looks).
    # Only motif/content pages genuinely carry icons in the source deck; toc,
    # timeline, table, references, closing, summary legitimately don't.
    NEEDS_ICON = {"content", "comparison", "hub", "data", "concept", "kpi"}
    if role == "cover":
        required, min_score = ["bg_visual"], 3
    elif role == "section":
        required, min_score = ["bg_visual", "decoration_layer"], 4
    elif role in NEEDS_ICON:
        required, min_score = ["bg_visual", "decoration_layer"], 6
    else:  # toc / timeline / table / references / closing / summary — rich but icon-optional
        required, min_score = ["bg_visual", "decoration_layer"], 4

    missing = [r for r in required if not signals.get(r)]
    if role in NEEDS_ICON and icon_count == 0:
        missing.append("icons (use <use data-icon> — placeholder shapes don't count)")

    too_plain = bool(missing) or score < min_score
    if too_plain:
        errors.append(f"TOO_PLAIN: richness {score}/{min_score} min · missing: "
                      f"{', '.join(missing) if missing else 'raise score'}")

    richness = {
        "score": score, "min": min_score, "signals": signals,
        "detail": sig_detail, "missing": missing,
    }
    verdict = "FAIL" if errors else ("WARN" if warnings else "PASS")
    return {"file": path.name, "role": role, "verdict": verdict,
            "errors": errors, "warnings": warnings, "richness": richness}


def infer_role(name: str) -> str:
    n = name.lower()
    for key, role in (("cover", "cover"), ("section", "section"), ("toc", "toc"),
                      ("seg", "content"), ("compar", "comparison"), ("hub", "hub"),
                      ("logic", "hub"), ("timeline", "timeline"), ("process", "timeline"),
                      ("table", "table"), ("chart", "data"), ("data", "data"),
                      ("kpi", "data"), ("refer", "references"), ("closing", "closing"),
                      ("concept", "content"), ("summary", "content")):
        if key in n:
            return role
    return "content"


def main():
    ap = argparse.ArgumentParser(description="Stage-5 aesthetic quality gate")
    ap.add_argument("svgs", nargs="+")
    ap.add_argument("--calibration", required=True)
    ap.add_argument("--role", default=None, help="override role for all inputs")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    cal = json.loads(Path(args.calibration).read_text(encoding="utf-8"))
    reports = []
    for s in args.svgs:
        p = Path(s)
        role = args.role or infer_role(p.name)
        reports.append(analyze(p, cal, role))

    if args.json:
        print(json.dumps(reports, ensure_ascii=False, indent=2))
    else:
        n_fail = 0
        for r in reports:
            mark = {"PASS": "PASS", "WARN": "WARN", "FAIL": "FAIL"}[r["verdict"]]
            rc = r.get("richness", {})
            print(f"[{mark}] {r['file']}  (role={r['role']}, "
                  f"richness={rc.get('score','?')}/{rc.get('min','?')})")
            for e in r["errors"]:
                print(f"    ERROR  {e}")
            for w in r["warnings"]:
                print(f"    warn   {w}")
            if r["verdict"] == "FAIL":
                n_fail += 1
        print(f"\n{len(reports)} file(s): {n_fail} FAIL, "
              f"{sum(1 for r in reports if r['verdict']=='WARN')} WARN, "
              f"{sum(1 for r in reports if r['verdict']=='PASS')} PASS")
    sys.exit(1 if any(r["verdict"] == "FAIL" for r in reports) else 0)


if __name__ == "__main__":
    main()
