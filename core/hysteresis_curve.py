"""Generate viscous damper hysteresis curve images."""
import json
import math
import os
import re

from PIL import Image, ImageDraw, ImageFont


MODEL_KEY = '\u4ea7\u54c1\u578b\u53f7'
MAX_FORCE_KEY = '\u6700\u5927\u963b\u5c3c\u529b(kN)'
DESIGN_DISPLACEMENT_KEY = '\u8bbe\u8ba1\u4f4d\u79fb(mm)'
DAMPING_INDEX_KEY = '\u963b\u5c3c\u529b\u6307\u6570(\u03b1)'
CURVE_TEMPLATE_KEY = '_curve_template_name'
_TEMPLATE_CACHE = None


def _to_float(value, default):
    if value is None:
        return default
    text = str(value).strip()
    if not text:
        return default
    match = re.search(r'-?\d+(?:\.\d+)?', text)
    if not match:
        return default
    try:
        return float(match.group(0))
    except ValueError:
        return default


def _font(size, bold=False):
    candidates = [
        'C:/Windows/Fonts/simhei.ttf',
        'C:/Windows/Fonts/msyh.ttc',
        'C:/Windows/Fonts/simsun.ttc',
        'C:/Windows/Fonts/arialbd.ttf' if bold else 'C:/Windows/Fonts/arial.ttf',
    ]
    for path in candidates:
        if path and os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except OSError:
                pass
    return ImageFont.load_default()


def _nice_step(span, target_ticks=8):
    raw = span / max(1, target_ticks)
    exponent = math.floor(math.log10(raw)) if raw > 0 else 0
    fraction = raw / (10 ** exponent)
    if fraction <= 1:
        nice = 1
    elif fraction <= 2:
        nice = 2
    elif fraction <= 5:
        nice = 5
    else:
        nice = 10
    return nice * (10 ** exponent)


def _symmetric_axis_limit(value, padding=1.35, target_ticks=8):
    value = max(abs(value), 1.0) * padding
    step = _nice_step(value * 2, target_ticks)
    return max(step, math.ceil(value / step) * step)


def _smoothstep(t):
    return t * t * (3 - 2 * t)


def _template_path():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(root, 'products', 'viscous_damper', 'curve_templates.json')


def _load_curve_templates():
    global _TEMPLATE_CACHE
    if _TEMPLATE_CACHE is not None:
        return _TEMPLATE_CACHE

    path = _template_path()
    if not os.path.exists(path):
        _TEMPLATE_CACHE = []
        return _TEMPLATE_CACHE

    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        _TEMPLATE_CACHE = data.get('templates', [])
    except (OSError, json.JSONDecodeError):
        _TEMPLATE_CACHE = []
    return _TEMPLATE_CACHE


def choose_curve_template_name(max_force, displacement, used_names=None):
    templates = _load_curve_templates()
    if not templates:
        return None
    used_names = set(used_names or [])

    def score(item):
        item_disp = max(float(item.get('design_displacement') or 1), 1)
        item_force = max(float(item.get('design_force') or 1), 1)
        disp_score = abs(item_disp - displacement) / max(displacement, 1)
        force_score = abs(item_force - max_force) / max(max_force, 1)
        return disp_score * 2.0 + force_score * 0.35

    ranked = sorted(templates, key=score)
    for item in ranked:
        if item.get('name') not in used_names:
            return item.get('name')
    return ranked[0].get('name')


def _choose_curve_template(max_force, displacement, template_name=None):
    templates = _load_curve_templates()
    if not templates:
        return None
    if template_name:
        for item in templates:
            if item.get('name') == template_name:
                return item

    chosen_name = choose_curve_template_name(max_force, displacement)
    for item in templates:
        if item.get('name') == chosen_name:
            return item
    return None


def _scaled_template_points(max_force, displacement, template_name=None):
    template = _choose_curve_template(max_force, displacement, template_name)
    if not template:
        return None

    points = template.get('points') or []
    if not points:
        return None

    design_displacement = max(float(template.get('design_displacement') or 1), 1)
    peak_abs_load = max(float(template.get('peak_abs_load') or 1), 1)
    x_scale = displacement / design_displacement
    y_scale = max_force / peak_abs_load

    return [(float(x) * x_scale, float(y) * y_scale) for x, y in points]


def _curve_points(max_force, displacement, alpha, template_name=None):
    displacement = max(abs(displacement), 1.0)
    max_force = max(abs(max_force), 1.0)
    template_points = _scaled_template_points(max_force, displacement, template_name)
    if template_points:
        return template_points

    points = []
    clearance = min(2.0, displacement * 0.12)

    top_force = max_force * 0.985
    bottom_force = -max_force * 0.985
    displacement_offsets = [0.00, 0.10, -0.08, 0.18, -0.12]

    def add_segment(start, end, count, wiggle=0.0, phase=0.0, x_wiggle=0.0):
        sx, sy = start
        ex, ey = end
        for i in range(count + 1):
            u = i / count
            s = _smoothstep(u)
            x = sx * (1 - s) + ex * s
            y = sy * (1 - s) + ey * s
            if wiggle:
                rough = (
                    0.70 * math.sin(math.pi * u)
                    + 0.22 * math.sin(7 * math.pi * u) * math.cos(phase)
                    + 0.08 * math.sin(17 * math.pi * u) * math.sin(phase + 0.4)
                )
                y += max_force * wiggle * rough
            if x_wiggle:
                x += displacement * x_wiggle * math.sin(math.pi * u + phase)
            points.append((x, y))

    def add_curve(start, control, end, count, wiggle=0.0, phase=0.0):
        sx, sy = start
        cx, cy = control
        ex, ey = end
        for i in range(count + 1):
            u = i / count
            one = 1 - u
            x = one * one * sx + 2 * one * u * cx + u * u * ex
            y = one * one * sy + 2 * one * u * cy + u * u * ey
            if wiggle:
                rough = (
                    0.70 * math.sin(math.pi * u)
                    + 0.22 * math.sin(7 * math.pi * u) * math.cos(phase)
                    + 0.08 * math.sin(17 * math.pi * u) * math.sin(phase + 0.4)
                )
                y += max_force * wiggle * rough
            points.append((x, y))

    # Starting loading section: it stays around zero force and connects
    # horizontally into the outer loop's fixture-clearance step.
    first_right = displacement + displacement_offsets[0]
    add_segment((0.0, 0.0), (first_right - clearance, 0.0), 120, wiggle=0.005)

    def add_outer_cycle(cycle):
        force_gain = 1.0 - 0.004 * cycle + 0.006 * math.sin(cycle * 1.7)
        top = top_force * force_gain
        bottom = bottom_force * force_gain
        phase = cycle * 0.7
        right = displacement + displacement_offsets[cycle % len(displacement_offsets)]
        left = -displacement + 0.8 * displacement_offsets[(cycle + 2) % len(displacement_offsets)]

        # Right side, positive displacement is locked to the design
        # displacement with small realistic control error.
        add_segment((right - clearance, 0.0), (right, 0.0), 28, wiggle=0.005, phase=phase)
        add_segment((right, 0.0), (right, top * 0.38), 18, wiggle=0.0015, phase=phase)
        add_curve((right, top * 0.38), (right * 1.005, top * 0.88), (right * 0.84, top), 74, 0.004, phase)

        # Upper platform.
        add_segment(
            (right * 0.84, top),
            (left * 0.82, top * 0.99),
            170,
            wiggle=0.026,
            x_wiggle=0.003,
            phase=phase + 0.3,
        )

        # Left side with a strict zero-force clearance step.
        add_curve((left * 0.82, top * 0.99), (left * 1.00, top * 0.78), (left + clearance, max_force * 0.16), 78, 0.004, phase)
        add_segment((left + clearance, max_force * 0.16), (left + clearance, 0.0), 18, wiggle=0.0015, phase=phase)
        add_segment((left + clearance, 0.0), (left, 0.0), 28, wiggle=0.005, phase=phase + 0.5)
        add_segment((left, 0.0), (left, bottom * 0.38), 18, wiggle=0.0015, phase=phase)
        add_curve((left, bottom * 0.38), (left * 1.005, bottom * 0.88), (left * 0.84, bottom), 74, 0.004, phase)

        # Lower platform.
        add_segment(
            (left * 0.84, bottom),
            (right * 0.82, bottom * 0.99),
            170,
            wiggle=0.026,
            x_wiggle=0.003,
            phase=phase + 0.9,
        )

        # Right side with a strict zero-force clearance step.
        add_curve((right * 0.82, bottom * 0.99), (right * 1.00, bottom * 0.78), (right - clearance, -max_force * 0.16), 78, 0.004, phase)
        add_segment((right - clearance, -max_force * 0.16), (right - clearance, 0.0), 18, wiggle=0.0015, phase=phase)
        add_segment((right - clearance, 0.0), (right, 0.0), 28, wiggle=0.005, phase=phase + 0.5)

    # Five high-overlap cycles. Max displacement points are identical; only
    # force platforms have small deterministic variation.
    for cycle in range(5):
        add_outer_cycle(cycle)

    # Ending unloading section: horizontal zero-force exit from the final
    # outer loop, with the same small fixture noise as the clearance step.
    last_right = displacement + displacement_offsets[4]
    add_segment((last_right, 0.0), (clearance, 0.0), 120, wiggle=0.005, phase=1.1)

    return points


def generate_hysteresis_curve(mechanical_row, output_path, size=(550, 350)):
    """Generate one hysteresis curve PNG for a mechanical data row."""
    max_force = _to_float(mechanical_row.get(MAX_FORCE_KEY), 300.0)
    displacement = _to_float(mechanical_row.get(DESIGN_DISPLACEMENT_KEY), 30.0)
    alpha = _to_float(mechanical_row.get(DAMPING_INDEX_KEY), 0.3)
    template_name = mechanical_row.get(CURVE_TEMPLATE_KEY)

    width, height = size
    image = Image.new('RGB', size, 'white')
    draw = ImageDraw.Draw(image)

    title_font = _font(13, bold=True)
    label_font = _font(11)
    tick_font = _font(9)

    left, top, right, bottom = 74, 28, width - 22, height - 42
    plot_w = right - left
    plot_h = bottom - top

    x_limit = max(displacement * 1.12, 1.0)
    y_limit = max(max_force * 1.18, 1.0)
    x_step = x_limit / 5
    y_step = y_limit / 5

    def to_px(point):
        x, y = point
        px = left + (x + x_limit) / (2 * x_limit) * plot_w
        py = bottom - (y + y_limit) / (2 * y_limit) * plot_h
        return px, py

    grid_color = (0, 120, 120)
    axis_color = (0, 80, 80)
    text_color = (0, 0, 0)
    blue = (0, 0, 255)

    draw.rectangle([left, top, right, bottom], outline=axis_color, width=1)

    def dashed_line(start, end, dash=4, gap=4):
        x1, y1 = start
        x2, y2 = end
        length = math.hypot(x2 - x1, y2 - y1)
        if length <= 0:
            return
        dx = (x2 - x1) / length
        dy = (y2 - y1) / length
        pos = 0
        while pos < length:
            seg_end = min(pos + dash, length)
            draw.line(
                [
                    (x1 + dx * pos, y1 + dy * pos),
                    (x1 + dx * seg_end, y1 + dy * seg_end),
                ],
                fill=grid_color,
                width=1,
            )
            pos += dash + gap

    x = -math.floor(x_limit / x_step) * x_step
    while x <= x_limit + 1e-9:
        px, _ = to_px((x, 0))
        dashed_line((px, top), (px, bottom))
        label = f'{x:.2f}'
        box = draw.textbbox((0, 0), label, font=tick_font)
        draw.text((px - (box[2] - box[0]) / 2, bottom + 7), label, fill=text_color, font=tick_font)
        x += x_step

    y = -math.floor(y_limit / y_step) * y_step
    while y <= y_limit + 1e-9:
        _, py = to_px((0, y))
        dashed_line((left, py), (right, py))
        label = f'{y:.2f}'
        box = draw.textbbox((0, 0), label, font=tick_font)
        draw.text((left - (box[2] - box[0]) - 6, py - 5), label, fill=text_color, font=tick_font)
        y += y_step

    points = [to_px(p) for p in _curve_points(max_force, displacement, alpha, template_name)]
    draw.line(points, fill=blue, width=1)

    title = '\u963b\u5c3c\u529b\u2014\u4f4d\u79fb\u66f2\u7ebf'
    x_label = '\u4f4d\u79fb(mm)'
    y_label = '\u963b\u5c3c\u529b(kN)'
    title_box = draw.textbbox((0, 0), title, font=title_font)
    draw.text(((width - (title_box[2] - title_box[0])) / 2, 8), title, fill=blue, font=title_font)

    x_box = draw.textbbox((0, 0), x_label, font=label_font)
    draw.text(((width - (x_box[2] - x_box[0])) / 2, height - 23), x_label, fill=text_color, font=label_font)

    y_img = Image.new('RGBA', (90, 18), (255, 255, 255, 0))
    y_draw = ImageDraw.Draw(y_img)
    y_draw.text((0, 0), y_label, fill=text_color, font=label_font)
    y_img = y_img.rotate(90, expand=True)
    image.paste(y_img, (16, top + (plot_h - y_img.height) // 2), y_img)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    image.save(output_path)
    return output_path
