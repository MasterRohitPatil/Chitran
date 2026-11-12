"""
Sidekick OS — Procedural Vector Cyber-Slime Renderer ("Byte")

High-DPI procedural vector renderer (256x256 supersampled, LANCZOS downsampled)
specifically crafted for crisp legibility inside a 44px Windows 11 taskbar and
64px System Tray icon:
  - Big expressive eyes & glossy jelly highlights (zero pixelation, zero seams)
  - Smooth squish -> mini-hop -> land wobble & happy blinks
  - Mode-specific visuals:
      * focused   : Luminous Ice-Cyan Cyber-Slime, big sparkling eyes, happy blinks
      * working   : Warm Amber-Cream slime + flipped-down Amber Cyber-Goggles + sparks
      * alert     : Soft Coral-Pink slime, wide crimson eyes, crisp "!" badge
      * paused    : Calm Mist-Grey melted jelly puddle, peacefully shut eyes (NO Zzz)
      * hydration : Sky-Aqua slime + floating glowing water droplet badge
      * mail      : Mint-Teal slime + floating glowing envelope badge
  - Anti-halo pre-compositing for Windows Tkinter `-transparentcolor` windows
"""

from __future__ import annotations

import math
from typing import Dict, List, Optional, Tuple

from PIL import Image, ImageChops, ImageDraw

from state import VALID_STATES


# ---------------------------------------------------------------------------
# Color palettes per mascot state: (primary_body_rgb, inner_shade_rgb)
# Calibrated in bright pastels so Byte pops against Win11 dark taskbar.
# ---------------------------------------------------------------------------

HOODIE_PALETTES: Dict[str, Tuple[Tuple[int, int, int], Tuple[int, int, int]]] = {
    "focused": ((132, 240, 255), (42, 182, 232)),
    "working": ((255, 220, 128), (235, 148, 24)),
    "alert": ((255, 164, 178), (225, 62, 82)),
    "paused": ((186, 202, 222), (118, 138, 166)),
    "hydration": ((138, 238, 255), (28, 158, 235)),
    "mail": ((142, 246, 208), (24, 182, 134)),
}

OUTLINE_COLOR = (22, 30, 52, 255)
EYE_DARK = (16, 24, 46, 255)
TASKBAR_MATTE_RGB = (24, 28, 38)
SUPERSAMPLE_SIZE = 256
CYCLE_FRAMES = 16

# Pre-rendered steady-state cache: (state, phase, size, facing_right, transparent_bg)
_FRAME_CACHE: Dict[Tuple[str, int, int, bool, Tuple[int, int, int, int]], Image.Image] = {}


def lerp_rgb(
    c1: Tuple[int, int, int],
    c2: Tuple[int, int, int],
    t: float,
) -> Tuple[int, int, int]:
    """Linearly interpolate between two RGB color tuples."""
    t_clamped = max(0.0, min(1.0, float(t)))
    return tuple(
        int(round(a + (b - a) * t_clamped))
        for a, b in zip(c1, c2)
    )  # type: ignore[return-value]


def get_target_hoodie_colors(state: str) -> Tuple[Tuple[int, int, int], Tuple[int, int, int]]:
    """Return (primary_rgb, shade_rgb) target palette for a given mascot state."""
    key = state if state in HOODIE_PALETTES else "focused"
    return HOODIE_PALETTES[key]


def _derive_shade_rgb(primary: Tuple[int, int, int], target_primary: Tuple[int, int, int], target_shade: Tuple[int, int, int]) -> Tuple[int, int, int]:
    """Compute a harmonious inner-shade RGB matching the current body color."""
    if primary == target_primary:
        return target_shade
    # Keep rich saturated tint relative to current primary
    r, g, b = primary
    return (max(18, int(r * 0.56)), max(24, int(g * 0.72)), max(38, int(b * 0.86)))


def _cubic_bezier(
    p0: Tuple[float, float],
    p1: Tuple[float, float],
    p2: Tuple[float, float],
    p3: Tuple[float, float],
    steps: int = 18,
) -> List[Tuple[float, float]]:
    """Evaluate a cubic Bezier curve segment."""
    pts: List[Tuple[float, float]] = []
    for i in range(steps):
        t = i / float(steps)
        u = 1.0 - t
        x = (u ** 3) * p0[0] + 3 * (u ** 2) * t * p1[0] + 3 * u * (t ** 2) * p2[0] + (t ** 3) * p3[0]
        y = (u ** 3) * p0[1] + 3 * (u ** 2) * t * p1[1] + 3 * u * (t ** 2) * p2[1] + (t ** 3) * p3[1]
        pts.append((x, y))
    return pts


def _build_slime_polygon(
    cx: float,
    base_y: float,
    half_w: float,
    height: float,
    crown_taper: float = 0.78,
    lean_x: float = 0.0,
) -> List[Tuple[float, float]]:
    """
    Build a C1-continuous, plush mochi-slime silhouette with a soft flat-ish
    resting bottom and a rounded glossy dome crown.
    """
    top_y = base_y - height
    waist_y = base_y - height * 0.34
    left_x = cx - half_w
    right_x = cx + half_w
    top_x = cx + lean_x

    # Base points: bottom-center -> bottom-right -> right waist -> crown -> left waist -> bottom-left
    b_mid = (cx, base_y)
    r_mid = (right_x, waist_y)
    t_mid = (top_x, top_y)
    l_mid = (left_x, waist_y)

    c_top = half_w * crown_taper * 0.62
    c_side_up = height * 0.36
    c_side_dn = height * 0.24
    c_bot = half_w * 0.82

    pts: List[Tuple[float, float]] = []
    # Bottom-center -> Right waist (plush rounded floor contact)
    pts.extend(_cubic_bezier(b_mid, (cx + c_bot, base_y), (right_x, base_y - c_side_dn * 0.25), r_mid, 20))
    # Right waist -> Crown top
    pts.extend(_cubic_bezier(r_mid, (right_x, waist_y - c_side_up), (top_x + c_top, top_y), t_mid, 22))
    # Crown top -> Left waist
    pts.extend(_cubic_bezier(t_mid, (top_x - c_top, top_y), (left_x, waist_y - c_side_up), l_mid, 22))
    # Left waist -> Bottom-center
    pts.extend(_cubic_bezier(l_mid, (left_x, base_y - c_side_dn * 0.25), (cx - c_bot, base_y), b_mid, 20))
    return pts


def _scale_polygon_around(
    pts: List[Tuple[float, float]],
    cx: float,
    cy: float,
    sx: float,
    sy: float,
) -> List[Tuple[float, float]]:
    return [(cx + (x - cx) * sx, cy + (y - cy) * sy) for (x, y) in pts]


def _draw_arc_stroke(
    draw: ImageDraw.ImageDraw,
    pts: List[Tuple[float, float]],
    fill: Tuple[int, int, int, int],
    width: int,
) -> None:
    """Draw a smooth round-capped polyline stroke."""
    if len(pts) < 2:
        return
    draw.line(pts, fill=fill, width=width, joint="curve")
    r = max(1.0, width / 2.0)
    for px, py in (pts[0], pts[-1]):
        draw.ellipse([px - r, py - r, px + r, py + r], fill=fill)


def _compute_motion(state: str, phase: int) -> Tuple[float, float, float, float, float, bool]:
    """
    Return (sx, sy, lift_y, wobble_x, lean_x, is_blinking) for the given state
    and animation phase (0..15).
    """
    if state == "paused":
        # Gentle slow breathing puddle (NO hops, NO Zzz)
        breathe = math.sin((phase / CYCLE_FRAMES) * 2.0 * math.pi)
        return (1.0 + 0.025 * breathe, 1.0 - 0.04 * breathe, 0.0, 0.0, 0.0, False)

    if state == "alert":
        # Brisk tense alert pulse & horizontal shudder
        pulse = math.sin((phase / 4.0) * 2.0 * math.pi)
        shudder = (-4.5 if (phase % 2 == 0) else 4.5) if (phase < 10) else 0.0
        return (1.0 - 0.04 * pulse, 1.0 + 0.05 * pulse, max(0.0, pulse) * 6.0, shudder, 0.0, False)

    if state == "working":
        # Rhythmic bouncy focus groove
        groove = math.sin((phase / 8.0) * 2.0 * math.pi)
        return (1.0 + 0.055 * groove, 1.0 - 0.065 * groove, max(0.0, -groove) * 7.0, 0.0, 2.5 * groove, False)

    # focused / hydration / mail: playful squish -> mini-hop -> land wobble + happy blink
    # Phases:
    #   0..1  : gentle anticipate squish
    #   2..5  : spring upward mini-hop arc (clearly visible at 44px!)
    #   6..8  : squash on landing + rebound
    #   9..15 : settle into bouncy idle + happy blink on frames 12, 13
    is_blinking = phase in (12, 13)
    if phase in (0, 1):
        t = (phase + 1) / 2.0
        return (1.0 + 0.11 * t, 1.0 - 0.13 * t, 0.0, 0.0, 0.0, False)
    if 2 <= phase <= 5:
        arc = math.sin(((phase - 1.5) / 4.0) * math.pi)  # peaks around phase 3.5
        return (0.91 + 0.04 * (1.0 - arc), 1.09 + 0.03 * arc, 26.0 * arc, 0.0, 0.0, False)
    if phase == 6:
        return (1.13, 0.84, 0.0, 0.0, 0.0, False)
    if phase == 7:
        return (0.95, 1.05, 3.0, 0.0, 0.0, False)
    if phase == 8:
        return (1.03, 0.96, 0.0, 0.0, 0.0, False)

    idle = math.sin(((phase - 8) / 8.0) * math.pi)
    return (1.0 + 0.03 * idle, 1.0 - 0.03 * idle, 0.0, 0.0, 0.0, is_blinking)


def _draw_sparkle_eye(
    draw: ImageDraw.ImageDraw,
    ex: float,
    ey: float,
    rx: float,
    ry: float,
    iris_rgb: Tuple[int, int, int],
    gaze_x: float = 0.0,
    wide_alert: bool = False,
) -> None:
    """Draw a large, high-contrast kawaii eye with colored lower crescent + white shine."""
    draw.ellipse([ex - rx, ey - ry, ex + rx, ey + ry], fill=EYE_DARK)

    # Glowing colored lower iris crescent (keeps dark center contrast while radiating color)
    irx = rx * (0.64 if not wide_alert else 0.72)
    iry = ry * (0.48 if not wide_alert else 0.54)
    icy = ey + ry * 0.24
    draw.chord(
        [ex - irx + gaze_x * 0.4, icy - iry, ex + irx + gaze_x * 0.4, icy + iry],
        start=10,
        end=170,
        fill=(*iris_rgb, 245),
    )

    if wide_alert:
        # Sharp inner pinpoint glint for alert readiness
        pr = max(3.5, rx * 0.24)
        draw.ellipse([ex - pr + gaze_x * 0.3, ey - pr, ex + pr + gaze_x * 0.3, ey + pr], fill=(255, 245, 245, 255))

    # Crisp primary top-left white highlight + secondary bottom-right sparkle
    hr1 = max(4.5, rx * 0.36)
    hx1 = ex - rx * 0.32 + gaze_x * 0.45
    hy1 = ey - ry * 0.36
    draw.ellipse([hx1 - hr1, hy1 - hr1 * 0.9, hx1 + hr1, hy1 + hr1 * 0.9], fill=(255, 255, 255, 255))

    hr2 = max(2.5, rx * 0.17)
    hx2 = ex + rx * 0.30 + gaze_x * 0.35
    hy2 = ey + ry * 0.28
    draw.ellipse([hx2 - hr2, hy2 - hr2, hx2 + hr2, hy2 + hr2], fill=(255, 255, 255, 235))


def _draw_curved_eye(
    draw: ImageDraw.ImageDraw,
    ex: float,
    ey: float,
    span_w: float,
    arch_h: float,
    happy_up: bool = True,
    width: int = 8,
) -> None:
    """Draw a clean ^ (happy blink) or u (peaceful resting) curved eye."""
    sign = -1.0 if happy_up else 1.0
    pts: List[Tuple[float, float]] = []
    steps = 16
    for i in range(steps + 1):
        u = -1.0 + 2.0 * (i / float(steps))
        px = ex + u * span_w
        py = ey + sign * arch_h * (1.0 - u * u) - sign * arch_h * 0.35
        pts.append((px, py))
    _draw_arc_stroke(draw, pts, EYE_DARK, width=width)


def _draw_cyber_goggles(
    draw: ImageDraw.ImageDraw,
    cx: float,
    eye_y: float,
    eye_sep: float,
    gaze_x: float,
    phase: int,
) -> None:
    """Draw sleek flipped-down Amber Cyber-Goggles strictly for WORKING mode."""
    gx = cx + gaze_x * 0.55
    # Strap hugging slime head
    draw.line([gx - eye_sep - 38, eye_y - 4, gx + eye_sep + 38, eye_y - 4], fill=(28, 36, 56, 255), width=12)

    # Outer dark titanium visor bridge & frames
    lens_rx = 32
    lens_ry = 26
    # Nose bridge
    draw.rounded_rectangle([gx - 18, eye_y - 10, gx + 18, eye_y + 10], radius=6, fill=(32, 42, 66, 255))

    scan_shift = ((phase % 8) / 7.0) * 14.0 - 7.0
    for sign in (-1, 1):
        lx = gx + sign * eye_sep
        # Outer metallic rim
        draw.rounded_rectangle(
            [lx - lens_rx - 5, eye_y - lens_ry - 5, lx + lens_rx + 5, eye_y + lens_ry + 5],
            radius=15,
            fill=(26, 34, 56, 255),
            outline=(15, 20, 35, 255),
            width=3,
        )
        # Glowing Amber lens
        draw.rounded_rectangle(
            [lx - lens_rx, eye_y - lens_ry, lx + lens_rx, eye_y + lens_ry],
            radius=12,
            fill=(251, 176, 36, 255),
        )
        # Bright inner upper amber glow
        draw.rounded_rectangle(
            [lx - lens_rx + 4, eye_y - lens_ry + 4, lx + lens_rx - 4, eye_y],
            radius=8,
            fill=(254, 224, 110, 240),
        )
        # Diagonal crisp white cyber reflection glint
        glare_x = lx - 10 + scan_shift * 0.45
        draw.polygon(
            [
                (glare_x - 8, eye_y - lens_ry + 5),
                (glare_x + 4, eye_y - lens_ry + 5),
                (glare_x - 10, eye_y + lens_ry - 5),
                (glare_x - 20, eye_y + lens_ry - 5),
            ],
            fill=(255, 252, 232, 235),
        )


def _draw_alert_badge(draw: ImageDraw.ImageDraw, phase: int) -> None:
    """Draw a bold, unmistakable glowing '!' warning badge at top-right."""
    bob = -3.0 * math.sin((phase / 4.0) * math.pi)
    bx, by = 204.0, 56.0 + bob
    r = 32.0
    # Outer badge ring
    draw.ellipse([bx - r - 4, by - r - 4, bx + r + 4, by + r + 4], fill=(22, 28, 48, 255))
    draw.ellipse([bx - r, by - r, bx + r, by + r], fill=(244, 63, 94, 255))
    draw.ellipse([bx - r + 5, by - r + 5, bx + r - 5, by + r - 5], fill=(254, 102, 126, 255))
    # Bold white exclamation mark
    draw.rounded_rectangle([bx - 5.5, by - 19, bx + 5.5, by + 4], radius=5, fill=(255, 255, 255, 255))
    draw.ellipse([bx - 6.0, by + 10, bx + 6.0, by + 22], fill=(255, 255, 255, 255))


def _draw_hydration_badge(draw: ImageDraw.ImageDraw, phase: int) -> None:
    """Draw a crisp floating glowing sky-aqua water droplet badge."""
    bob = math.sin((phase / CYCLE_FRAMES) * 2.0 * math.pi) * 6.0
    bx, by = 205.0, 66.0 + bob
    pts = [
        (bx, by - 30),
        (bx + 22, by + 3),
        (bx + 21, by + 18),
        (bx + 10, by + 28),
        (bx, by + 30),
        (bx - 10, by + 28),
        (bx - 21, by + 18),
        (bx - 22, by + 3),
    ]
    draw.polygon(pts, fill=(56, 189, 248, 255), outline=OUTLINE_COLOR)
    # Re-stroke clean thick outline
    draw.line(pts + [pts[0]], fill=OUTLINE_COLOR, width=6, joint="curve")
    # Inner bright aqua fill & white shine
    draw.ellipse([bx - 13, by - 2, bx + 13, by + 22], fill=(125, 232, 255, 255))
    draw.ellipse([bx - 9, by + 2, bx - 1, by + 12], fill=(255, 255, 255, 245))


def _draw_mail_badge(draw: ImageDraw.ImageDraw, phase: int) -> None:
    """Draw a crisp glowing envelope icon with an unread dot."""
    bob = math.sin((phase / CYCLE_FRAMES) * 2.0 * math.pi) * 5.0
    mx, my = 198.0, 66.0 + bob
    w, h = 32.0, 23.0
    draw.rounded_rectangle(
        [mx - w, my - h, mx + w, my + h],
        radius=8,
        fill=(242, 252, 255, 255),
        outline=OUTLINE_COLOR,
        width=6,
    )
    # Envelope flap V-line
    _draw_arc_stroke(
        draw,
        [(mx - w + 6, my - h + 6), (mx, my + 4), (mx + w - 6, my - h + 6)],
        fill=(20, 148, 172, 255),
        width=5,
    )
    # Crisp unread notification dot
    draw.ellipse([mx + w - 15, my - h - 7, mx + w + 3, my - h + 11], fill=(244, 63, 94, 255), outline=OUTLINE_COLOR, width=3)


def _draw_working_sparks(draw: ImageDraw.ImageDraw, phase: int) -> None:
    """Draw tiny rising cyber energy bubbles during WORKING mode."""
    t1 = (phase % 12) / 12.0
    t2 = ((phase + 6) % 12) / 12.0
    for bx_base, frac, r_base in ((44.0, t1, 7.5), (210.0, t2, 6.5)):
        by = 120.0 - frac * 68.0
        alpha = int(230 * math.sin(frac * math.pi))
        if alpha > 25:
            r = r_base * (0.7 + 0.3 * math.sin(frac * math.pi))
            draw.ellipse([bx_base - r, by - r, bx_base + r, by + r], fill=(255, 232, 130, alpha), outline=(255, 255, 255, alpha), width=2)


def _apply_taskbar_antihalo(
    img: Image.Image,
    transparent_bg: Tuple[int, int, int, int],
) -> Image.Image:
    """
    For Windows Tkinter `-transparentcolor` windows (which lack partial alpha),
    key near-transparent pixels strictly to `transparent_bg` and pre-composite
    soft edge anti-aliasing onto a dark taskbar matte so zero ragged halos appear.
    """
    if transparent_bg[3] == 0:
        return img

    key_rgba = (transparent_bg[0], transparent_bg[1], transparent_bg[2], 255)
    mr, mg, mb = TASKBAR_MATTE_RGB
    pixels = img.load()
    w, h = img.size
    for y in range(h):
        for x in range(w):
            r, g, b, a = pixels[x, y]
            if a < 105:
                pixels[x, y] = key_rgba
            elif a < 245:
                af = a / 255.0
                inv = 1.0 - af
                pixels[x, y] = (
                    int(r * af + mr * inv),
                    int(g * af + mg * inv),
                    int(b * af + mb * inv),
                    255,
                )
            else:
                pixels[x, y] = (r, g, b, 255)
    return img


def _render_slime_256(
    state: str,
    phase: int,
    primary_rgb: Tuple[int, int, int],
    shade_rgb: Tuple[int, int, int],
    facing_right: bool,
) -> Image.Image:
    """Render the full 256x256 RGBA Cyber-Slime illustration for one frame."""
    canvas = Image.new("RGBA", (SUPERSAMPLE_SIZE, SUPERSAMPLE_SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(canvas, "RGBA")

    sx, sy, lift_y, wobble_x, lean_x, is_blinking = _compute_motion(state, phase)
    cx = 128.0 + wobble_x
    gaze_dir = 1.0 if facing_right else -1.0
    gaze_x = gaze_dir * (5.5 if state != "paused" else 0.0)

    if state == "paused":
        base_y = 238.0
        half_w = 114.0 * sx
        height = 108.0 * sy
        crown_taper = 0.72
    else:
        base_y = 236.0 - lift_y
        half_w = 103.0 * sx
        height = 170.0 * sy
        crown_taper = 0.78

    # Subtle ground shadow directly underneath on taskbar floor (fades cleanly when hopping)
    shadow_alpha = int(max(36, 105 - lift_y * 2.5))
    shadow_w = half_w * max(0.68, 0.94 - lift_y * 0.008)
    draw.ellipse(
        [cx - shadow_w, 232.0, cx + shadow_w, 246.0],
        fill=(12, 16, 26, shadow_alpha),
    )

    body_poly = _build_slime_polygon(
        cx=cx,
        base_y=base_y,
        half_w=half_w,
        height=height,
        crown_taper=crown_taper,
        lean_x=lean_x + gaze_dir * 3.0,
    )

    # 1) Outer base fill
    draw.polygon(body_poly, fill=(*primary_rgb, 255))

    # 2) Inner glossy layer alpha-composited cleanly inside body mask
    mask = Image.new("L", (SUPERSAMPLE_SIZE, SUPERSAMPLE_SIZE), 0)
    ImageDraw.Draw(mask).polygon(body_poly, fill=255)

    shading = Image.new("RGBA", (SUPERSAMPLE_SIZE, SUPERSAMPLE_SIZE), (0, 0, 0, 0))
    sdraw = ImageDraw.Draw(shading, "RGBA")

    # Warm/deep bottom jelly cushion shadow inside body
    sdraw.ellipse(
        [cx - half_w * 1.08, base_y - height * 0.32, cx + half_w * 1.08, base_y + height * 0.34],
        fill=(*shade_rgb, 130),
    )
    # Soft inner rim glow
    inner_rim = _scale_polygon_around(body_poly, cx, base_y - height * 0.45, 0.96, 0.94)
    sdraw.line(inner_rim + [inner_rim[0]], fill=(*shade_rgb, 95), width=12, joint="curve")

    # Luminous translucent top dome sheen + crisp jelly specular highlight pill
    sheen_cy = base_y - height * 0.74
    sdraw.ellipse(
        [cx - half_w * 0.68, base_y - height * 0.95, cx + half_w * 0.68, base_y - height * 0.46],
        fill=(255, 255, 255, 76),
    )
    # Crisp glossy candy/jelly highlight arc top-left
    hl_x = cx - half_w * (0.34 if not facing_right else 0.24)
    hl_y = sheen_cy - height * 0.05
    sdraw.ellipse(
        [hl_x - 24, hl_y - 12, hl_x + 15, hl_y + 8],
        fill=(255, 255, 255, 225),
    )
    sdraw.ellipse(
        [hl_x + 21, hl_y + 4, hl_x + 32, hl_y + 14],
        fill=(255, 255, 255, 200),
    )

    clipped_alpha = ImageChops.multiply(shading.getchannel("A"), mask)
    shading.putalpha(clipped_alpha)
    canvas = Image.alpha_composite(canvas, shading)
    draw = ImageDraw.Draw(canvas, "RGBA")

    # 3) Crisp outer ink stroke around silhouette
    outline_w = 10 if state != "paused" else 9
    draw.line(body_poly + [body_poly[0]], fill=OUTLINE_COLOR, width=outline_w, joint="curve")

    # 4) Expressive Face Layer
    if state == "paused":
        eye_y = base_y - height * 0.44
        eye_sep = 43.0 * sx
        mouth_y = eye_y + 14.0
        # Soft mist-pink cozy cheeks
        for sign in (-1, 1):
            chx = cx + sign * (eye_sep + 26)
            draw.ellipse([chx - 15, eye_y + 5, chx + 15, eye_y + 18], fill=(244, 152, 178, 130))
            _draw_curved_eye(draw, cx + sign * eye_sep, eye_y, span_w=19.0, arch_h=9.0, happy_up=False, width=8)
        # Peaceful little curve mouth
        _draw_arc_stroke(
            draw,
            [(cx - 10, mouth_y), (cx, mouth_y + 6), (cx + 10, mouth_y)],
            fill=EYE_DARK,
            width=6,
        )
    else:
        eye_y = base_y - height * 0.47
        eye_sep = 41.0 * sx
        face_cx = cx + gaze_x
        mouth_y = eye_y + 22.0 * sy

        # Cute coral/pink cheek blushes
        blush_col = (255, 128, 162, 170) if state != "alert" else (255, 96, 122, 195)
        for sign in (-1, 1):
            chx = face_cx + sign * (eye_sep + 24)
            draw.ellipse([chx - 16, eye_y + 16, chx + 16, eye_y + 30], fill=blush_col)

        if state == "working":
            _draw_cyber_goggles(draw, cx, eye_y - 2.0, eye_sep, gaze_x, phase)
            # Focused little smile below goggles
            _draw_arc_stroke(
                draw,
                [(face_cx - 11, mouth_y + 9), (face_cx, mouth_y + 15), (face_cx + 11, mouth_y + 9)],
                fill=EYE_DARK,
                width=6,
            )
            _draw_working_sparks(draw, phase)
        elif is_blinking:
            # Crisp ^ ^ happy eyes + cheerful w-cat/slime smile
            for sign in (-1, 1):
                _draw_curved_eye(draw, face_cx + sign * eye_sep, eye_y + 2, span_w=20.0, arch_h=13.0, happy_up=True, width=9)
            _draw_arc_stroke(
                draw,
                [(face_cx - 13, mouth_y), (face_cx - 5, mouth_y + 7), (face_cx, mouth_y + 2), (face_cx + 5, mouth_y + 7), (face_cx + 13, mouth_y)],
                fill=EYE_DARK,
                width=6,
            )
        else:
            is_alert = state == "alert"
            iris_col = shade_rgb if not is_alert else (244, 63, 94)
            erx = 20.0 if not is_alert else 22.0
            ery = 23.0 * sy if not is_alert else 25.0
            for sign in (-1, 1):
                _draw_sparkle_eye(
                    draw,
                    ex=face_cx + sign * eye_sep,
                    ey=eye_y,
                    rx=erx,
                    ry=ery,
                    iris_rgb=iris_col,
                    gaze_x=gaze_dir * 3.0,
                    wide_alert=is_alert,
                )

            if is_alert:
                # Urgent furrowed brows + open alarmed 'o' mouth
                for sign in (-1, 1):
                    bx_in = face_cx + sign * (eye_sep - 15)
                    bx_out = face_cx + sign * (eye_sep + 18)
                    _draw_arc_stroke(
                        draw,
                        [(bx_out, eye_y - 31), (bx_in, eye_y - 23)],
                        fill=EYE_DARK,
                        width=7,
                    )
                draw.ellipse([face_cx - 8, mouth_y - 3, face_cx + 8, mouth_y + 13], fill=EYE_DARK)
            else:
                # Cute ":3" / happy smile
                _draw_arc_stroke(
                    draw,
                    [(face_cx - 11, mouth_y), (face_cx - 4, mouth_y + 6), (face_cx, mouth_y + 2), (face_cx + 4, mouth_y + 6), (face_cx + 11, mouth_y)],
                    fill=EYE_DARK,
                    width=6,
                )

    # 5) Mode badges (alert / hydration / mail)
    if state == "alert":
        _draw_alert_badge(draw, phase)
    elif state == "hydration":
        _draw_hydration_badge(draw, phase)
    elif state == "mail":
        _draw_mail_badge(draw, phase)

    return canvas


def render_byte_frame(
    state: str = "focused",
    size: int = 44,
    frame_idx: int = 0,
    transparent_bg: Tuple[int, int, int, int] = (0, 0, 0, 0),
    hoodie_rgb: Optional[Tuple[int, int, int]] = None,
    facing_right: bool = False,
) -> Image.Image:
    """
    Render one smooth antialiased Cyber-Slime frame at the requested pixel `size`.

    Preserves the exact public API required by `mascot.py` and `ByteTrayController`:
      - `transparent_bg=(0, 0, 0, 0)` returns clean RGBA (for pystray / previews)
      - opaque `transparent_bg=(1, 2, 3, 255)` triggers anti-halo compositing
        for Windows Tkinter `-transparentcolor` taskbar overlay windows.
    """
    norm_state = state if state in VALID_STATES else "focused"
    phase = int(frame_idx) % CYCLE_FRAMES
    target_primary, target_shade = get_target_hoodie_colors(norm_state)

    use_cache = hoodie_rgb is None or hoodie_rgb == target_primary
    cache_key = (norm_state, phase, int(size), bool(facing_right), transparent_bg)
    if use_cache and cache_key in _FRAME_CACHE:
        return _FRAME_CACHE[cache_key]

    primary_rgb = hoodie_rgb if hoodie_rgb is not None else target_primary
    shade_rgb = _derive_shade_rgb(primary_rgb, target_primary, target_shade)

    hi_res = _render_slime_256(
        state=norm_state,
        phase=phase,
        primary_rgb=primary_rgb,
        shade_rgb=shade_rgb,
        facing_right=facing_right,
    )

    out_size = max(16, int(size))
    resampling = getattr(Image, "Resampling", Image).LANCZOS
    downsampled = hi_res.resize((out_size, out_size), resampling)
    final_img = _apply_taskbar_antihalo(downsampled, transparent_bg)

    if use_cache:
        if len(_FRAME_CACHE) > 512:
            _FRAME_CACHE.clear()
        _FRAME_CACHE[cache_key] = final_img

    return final_img


# Backward-compatible alias if any caller imports render_byte_image
render_byte_image = render_byte_frame
