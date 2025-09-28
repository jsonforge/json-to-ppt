# 增强版 gui.py
import io
import json
import logging
import os
import time
import tempfile
import math
import copy
from io import BytesIO
from typing import Any, Dict, List, Optional, Tuple
import platform

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageFilter, ImageOps

# 导入main.py中的函数
from main import build, build_single_slide, get_image_bytes, validate, hex_to_rgb

logger = logging.getLogger("json_to_ppt_gui")
if not logger.handlers:
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)

# 常量定义
DEFAULT_WIDTH = 1280
DEFAULT_HEIGHT = 720
PREVIEW_MAX_WIDTH = 960
PREVIEW_MAX_HEIGHT = 540
DEBOUNCE_MS = 300

# 配色方案
COLOR_BG = "#f8f9fa"
COLOR_SIDEBAR = "#2c3e50"
COLOR_CANVAS_BG = "#ffffff"
COLOR_ERROR = "#e74c3c"
COLOR_SUCCESS = "#27ae60"
COLOR_WARNING = "#f39c12"
COLOR_INFO = "#3498db"
COLOR_ACCENT = "#5b7dea"
COLOR_TEXT = "#2c3e50"
COLOR_TEXT_LIGHT = "#7f8c8d"
COLOR_BORDER = "#dce1e7"

# 图表默认颜色
CHART_COLORS = ["#3B82F6", "#10B981", "#F59E0B", "#8B5CF6", "#EF4444", "#06B6D4", "#EC4899", "#6366F1"]

# 增强的示例JSON
SAMPLE_JSON = json.dumps(
    {
        "version": "1.0",
        "ppt": {
            "size": {"width": DEFAULT_WIDTH, "height": DEFAULT_HEIGHT, "unit": "px"},
            "defaultUnit": "px",
            "theme": {
                "colors": {
                    "primary": "#3B82F6",
                    "secondary": "#10B981",
                    "accent": "#F59E0B",
                    "danger": "#EF4444"
                },
                "fonts": {
                    "heading": "Microsoft YaHei",
                    "body": "Arial"
                }
            },
            "slides": [
                {
                    "id": "slide-1",
                    "title": "欢迎",
                    "background": {
                        "gradient": {
                            "type": "linear",
                            "angle": 45,
                            "stops": [
                                {"color": "#667eea", "position": 0},
                                {"color": "#764ba2", "position": 100}
                            ]
                        }
                    },
                    "transition": {"type": "fade", "duration": 1},
                    "elements": [
                        {
                            "type": "text",
                            "text": "JSON → PPT 设计器",
                            "box": {"x": 640, "y": 200, "w": 600, "h": 100},
                            "style": {"fontSize": 48, "align": "center", "color": "#ffffff", "bold": True},
                            "shadow": {"x": 2, "y": 2, "blur": 4, "color": "#00000040"},
                            "rotation": -2
                        },
                        {
                            "type": "shape",
                            "shapeType": "star",
                            "box": {"x": 100, "y": 100, "w": 100, "h": 100},
                            "fill": "#ffd700",
                            "rotation": 15,
                            "shadow": {"x": 3, "y": 3, "blur": 6, "color": "#00000030"}
                        },
                        {
                            "type": "line",
                            "points": [{"x": 200, "y": 400}, {"x": 1080, "y": 400}],
                            "stroke": "#ffffff",
                            "strokeWidth": 2,
                            "strokeStyle": "dashed"
                        }
                    ],
                },
                {
                    "id": "slide-2",
                    "title": "多元素展示",
                    "background": {"color": "#f7fafc"},
                    "elements": [
                        {
                            "type": "group",
                            "box": {"x": 50, "y": 50, "w": 300, "h": 200},
                            "elements": [
                                {
                                    "type": "shape",
                                    "shapeType": "roundRect",
                                    "box": {"x": 0, "y": 0, "w": 300, "h": 200},
                                    "fill": "#e6f7ff",
                                    "border": {"width": 2, "color": "#1890ff", "style": "solid"}
                                },
                                {
                                    "type": "text",
                                    "text": "组合元素",
                                    "box": {"x": 150, "y": 100, "w": 200, "h": 50},
                                    "style": {"fontSize": 24, "align": "center", "color": "#1890ff"}
                                }
                            ]
                        },
                        {
                            "type": "icon",
                            "icon": {"library": "fontawesome", "name": "star"},
                            "box": {"x": 400, "y": 100, "w": 60, "h": 60},
                            "color": "#ffd700"
                        },
                        {
                            "type": "smartArt",
                            "smartArtType": "process",
                            "box": {"x": 100, "y": 300, "w": 1080, "h": 300},
                            "nodes": [
                                {"text": "开始", "color": "#3B82F6"},
                                {"text": "过程", "color": "#10B981"},
                                {"text": "结束", "color": "#F59E0B"}
                            ]
                        }
                    ]
                }
            ],
        },
    },
    ensure_ascii=False,
    indent=2,
)


class EnhancedChartRenderer:
    """增强的图表渲染器，支持更多图表类型和效果"""

    _font_cache: Dict[str, ImageFont.ImageFont] = {}

    @staticmethod
    def _load_font(size: int, bold: bool = False) -> ImageFont.ImageFont:
        """加载字体，优先支持中文"""
        key = f"{size}-{1 if bold else 0}"
        cached = EnhancedChartRenderer._font_cache.get(key)
        if cached:
            return cached

        candidates = [
            os.path.join(os.environ.get("WINDIR", r"C:\\Windows"), "Fonts", "msyh.ttc"),
            os.path.join(os.environ.get("WINDIR", r"C:\\Windows"), "Fonts", "msyhbd.ttc") if bold else "",
            os.path.join(os.environ.get("WINDIR", r"C:\\Windows"), "Fonts", "simhei.ttf"),
            "/System/Library/Fonts/PingFang.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
            "/Library/Fonts/Microsoft YaHei.ttf",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
            "arial.ttf",
        ]

        for path in candidates:
            if not path:
                continue
            try:
                font = ImageFont.truetype(path, size)
                EnhancedChartRenderer._font_cache[key] = font
                return font
            except Exception:
                continue

        font = ImageFont.load_default()
        EnhancedChartRenderer._font_cache[key] = font
        return font

    @staticmethod
    def render_chart(chart_type: str, data: Dict, width: int, height: int,
                     title: str = "", options: Dict = None) -> Image.Image:
        """渲染图表"""
        img = Image.new('RGBA', (max(120, width), max(80, height)), color='white')
        draw = ImageDraw.Draw(img)

        # 外框
        draw.rectangle([0, 0, width - 1, height - 1], outline='#e0e0e0', width=1)

        # 字体
        title_font = EnhancedChartRenderer._load_font(16, bold=True)
        label_font = EnhancedChartRenderer._load_font(12)

        if title:
            tb = draw.textbbox((0, 0), title, font=title_font)
            draw.text(((width - (tb[2] - tb[0])) / 2, 8), title, fill="#333", font=title_font)

        categories = data.get("categories", []) or []
        series_list = data.get("series", []) or []

        # 根据类型分发
        ct = chart_type or "bar"
        try:
            if ct in ("bar", "barGroup"):
                EnhancedChartRenderer._draw_bar_chart(draw, categories, series_list, width, height, label_font, options)
            elif ct in ("barStacked", "barStacked100"):
                EnhancedChartRenderer._draw_bar_stacked_chart(draw, categories, series_list, width, height, label_font,
                                                              percent=(ct == "barStacked100"))
            elif ct in ("barHorizontal",):
                EnhancedChartRenderer._draw_bar_horizontal_chart(draw, categories, series_list, width, height,
                                                                 label_font)
            elif ct in ("line", "lineSmooth"):
                EnhancedChartRenderer._draw_line_chart(draw, categories, series_list, width, height, label_font,
                                                       smooth=(ct == "lineSmooth"))
            elif ct == "pie":
                EnhancedChartRenderer._draw_pie_chart(draw, categories, series_list, width, height, label_font)
            elif ct == "doughnut":
                EnhancedChartRenderer._draw_doughnut_chart(draw, categories, series_list, width, height, label_font)
            elif ct in ("area", "areaStacked"):
                EnhancedChartRenderer._draw_area_chart(draw, categories, series_list, width, height, label_font,
                                                       stacked=(ct == "areaStacked"))
            elif ct == "scatter":
                EnhancedChartRenderer._draw_scatter_chart(draw, categories, series_list, width, height, label_font)
            elif ct == "bubble":
                EnhancedChartRenderer._draw_bubble_chart(draw, categories, series_list, width, height, label_font)
            elif ct == "radar":
                EnhancedChartRenderer._draw_radar_chart(draw, categories, series_list, width, height, label_font)
            else:
                draw.text((width / 2 - 50, height / 2 - 10), f"{ct.upper()}\nCHART", fill="#666", font=title_font,
                          align="center")
        except Exception as e:
            draw.text((10, height / 2), f"预览失败: {e}", fill="red", font=label_font)

        # 绘制图例
        if options and options.get("legend", True):
            EnhancedChartRenderer._draw_legend(draw, series_list, width, height, label_font)

        return img

    @staticmethod
    def _draw_bar_chart(draw, categories, series_list, width, height, font, options=None):
        """绘制柱状图"""
        legend_h = 40 if options and options.get("legend", True) else 20
        margin_x = max(30, int(width * 0.06))
        margin_y_top = max(30, int(height * 0.12))
        margin_y_bottom = max(38, int(height * 0.18))
        top, bottom = margin_y_top, height - margin_y_bottom
        left, right = margin_x, width - margin_x

        if not categories or not series_list:
            return

        cat_count = len(categories)
        series_count = len(series_list)
        plot_w = max(1, (right - left))
        group_w = plot_w / cat_count

        inner_gap_ratio = 0.2 if series_count > 1 else 0.35
        inner_gap = inner_gap_ratio * group_w
        bar_w = (group_w - inner_gap) / max(1, series_count)

        max_val = max((max(s.get("values", [0])) for s in series_list), default=0)
        if max_val <= 0:
            max_val = 1

        # 绘制网格线
        grid_lines = 5
        for i in range(grid_lines + 1):
            ratio = i / grid_lines
            y = bottom - ratio * (bottom - top)
            val = int(max_val * ratio)
            if 0 < i < grid_lines:
                draw.line([left, y, right, y], fill="#e5e5e5", width=1)
            label = str(val)
            tw, th = draw.textbbox((0, 0), label, font=font)[2:4]
            draw.text((left - tw - 6, y - th / 2), label, fill="#444", font=font)

        # 绘制柱子
        for ci in range(cat_count):
            base_x = left + ci * group_w + inner_gap / 2
            for si, s in enumerate(series_list):
                vals = s.get("values", [])
                if ci >= len(vals):
                    continue
                v = vals[ci]
                color = s.get("color", CHART_COLORS[si % len(CHART_COLORS)])
                h = (v / max_val) * (bottom - top)
                x0 = base_x + si * bar_w
                y0 = bottom - h

                # 支持渐变
                gradient = s.get("gradient")
                if gradient:
                    # 简单的渐变模拟
                    stops = gradient.get("stops", [])
                    if len(stops) >= 2:
                        color = stops[0].get("color", color)

                draw.rectangle([x0, y0, x0 + bar_w * 0.9, bottom], fill=color, outline=color)

                # 数据标签
                if options and options.get("dataLabels", False):
                    draw.text((x0, y0 - 14), str(v), fill="#444", font=font)

        # 分类标签
        for ci, cat in enumerate(categories):
            cx = left + ci * group_w + group_w / 2
            label = str(cat)
            tb = draw.textbbox((0, 0), label, font=font)
            tw = tb[2] - tb[0]
            draw.text((cx - tw / 2, bottom + 6), label, fill="#333", font=font)

    @staticmethod
    def _draw_bar_horizontal_chart(draw, categories, series_list, width, height, font):
        """绘制水平柱状图"""
        margin = 50
        legend_h = 40
        top, bottom = margin, height - margin - legend_h
        left, right = margin + 40, width - margin

        if not categories or not series_list:
            return

        cat_count = len(categories)
        series_count = len(series_list)
        plot_h = bottom - top
        group_h = plot_h / cat_count
        bar_h = group_h / (series_count + 1)

        max_val = max((max(s.get("values", [0])) for s in series_list), default=0) or 1

        # 绘制坐标轴
        draw.line([left, top, left, bottom], fill="#555", width=1)
        draw.line([left, bottom, right, bottom], fill="#555", width=1)

        # 绘制柱子
        for ci, cat in enumerate(categories):
            base_y = top + ci * group_h

            # 分类标签
            draw.text((left - 40, base_y + group_h / 2 - 5), str(cat)[:8], fill="#333", font=font)

            for si, s in enumerate(series_list):
                vals = s.get("values", [])
                if ci >= len(vals):
                    continue
                v = vals[ci]
                color = s.get("color", CHART_COLORS[si % len(CHART_COLORS)])
                w = (v / max_val) * (right - left)
                y0 = base_y + si * bar_h + bar_h * 0.1
                draw.rectangle([left, y0, left + w, y0 + bar_h * 0.8], fill=color, outline=color)

    @staticmethod
    def _draw_doughnut_chart(draw, categories, series_list, width, height, font):
        """绘制环形图"""
        if not series_list or not categories:
            return
        vals = series_list[0].get("values", [])
        total = sum(vals) or 1
        cx, cy = width // 2, height // 2
        outer_radius = min(width, height) // 3
        inner_radius = outer_radius // 2

        start = 0
        for i, v in enumerate(vals[:len(categories)]):
            angle = v / total * 360
            color = series_list[0].get("color") or CHART_COLORS[i % len(CHART_COLORS)]

            # 外圈
            draw.pieslice([cx - outer_radius, cy - outer_radius, cx + outer_radius, cy + outer_radius],
                          start=start, end=start + angle, fill=color, outline="white")
            start += angle

        # 内圈（挖空）
        draw.ellipse([cx - inner_radius, cy - inner_radius, cx + inner_radius, cy + inner_radius],
                     fill="white", outline="white")

    @staticmethod
    def _draw_radar_chart(draw, categories, series_list, width, height, font):
        """绘制雷达图"""
        cx, cy = width // 2, height // 2
        radius = min(width, height) // 3
        n = len(categories)
        if n < 3:
            return

        # 绘制网格
        for r in [radius * 0.2, radius * 0.4, radius * 0.6, radius * 0.8, radius]:
            points = []
            for i in range(n):
                angle = 2 * math.pi * i / n - math.pi / 2
                x = cx + r * math.cos(angle)
                y = cy + r * math.sin(angle)
                points.append((x, y))

            # 绘制多边形网格
            for i in range(n):
                draw.line([points[i], points[(i + 1) % n]], fill="#e0e0e0", width=1)

        # 绘制轴线和标签
        for i, cat in enumerate(categories):
            angle = 2 * math.pi * i / n - math.pi / 2
            x = cx + radius * math.cos(angle)
            y = cy + radius * math.sin(angle)
            draw.line([(cx, cy), (x, y)], fill="#d0d0d0", width=1)

            # 标签
            label_x = cx + (radius + 20) * math.cos(angle)
            label_y = cy + (radius + 20) * math.sin(angle)
            draw.text((label_x - 20, label_y - 5), str(cat)[:8], fill="#333", font=font)

        # 绘制数据
        for si, s in enumerate(series_list):
            vals = s.get("values", [])
            max_val = max(vals) if vals else 1
            color = s.get("color", CHART_COLORS[si % len(CHART_COLORS)])

            points = []
            for i in range(min(n, len(vals))):
                v = vals[i] / max_val
                angle = 2 * math.pi * i / n - math.pi / 2
                x = cx + radius * v * math.cos(angle)
                y = cy + radius * v * math.sin(angle)
                points.append((x, y))

            if len(points) >= 3:
                # 半透明填充
                img_temp = Image.new('RGBA', (width, height), (255, 255, 255, 0))
                draw_temp = ImageDraw.Draw(img_temp)
                r, g, b = hex_to_rgb(color)[:3]
                draw_temp.polygon(points, fill=(r, g, b, 80), outline=color)
                draw.bitmap((0, 0), img_temp, fill=None)

                # 绘制边框
                for i in range(len(points)):
                    draw.line([points[i], points[(i + 1) % len(points)]], fill=color, width=2)

    @staticmethod
    def _draw_bubble_chart(draw, categories, series_list, width, height, font):
        """绘制气泡图"""
        margin = 50
        top, bottom = margin, height - margin - 40
        left, right = margin, width - margin

        if not series_list:
            return

        # 绘制坐标轴
        draw.line([left, bottom, right, bottom], fill="#555", width=1)
        draw.line([left, top, left, bottom], fill="#555", width=1)

        # 收集所有数据点
        all_points = []
        for si, s in enumerate(series_list):
            color = s.get("color", CHART_COLORS[si % len(CHART_COLORS)])
            for val in s.get("values", []):
                if isinstance(val, dict):
                    x = val.get("x", 0)
                    y = val.get("y", 0)
                    size = val.get("size", 10)
                    all_points.append((x, y, size, color))

        if not all_points:
            return

        # 计算范围
        xs = [p[0] for p in all_points]
        ys = [p[1] for p in all_points]
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)

        if max_x == min_x: max_x += 1
        if max_y == min_y: max_y += 1

        # 绘制气泡
        for x, y, size, color in all_points:
            px = left + (x - min_x) / (max_x - min_x) * (right - left)
            py = bottom - (y - min_y) / (max_y - min_y) * (bottom - top)
            r = size / 2

            # 半透明气泡
            img_temp = Image.new('RGBA', (width, height), (255, 255, 255, 0))
            draw_temp = ImageDraw.Draw(img_temp)
            rgb = hex_to_rgb(color)[:3]
            draw_temp.ellipse([px - r, py - r, px + r, py + r], fill=(*rgb, 120), outline=color)
            draw.bitmap((0, 0), img_temp, fill=None)

    @staticmethod
    def _draw_pie_chart(draw, categories, series_list, width, height, font):
        if not series_list or not categories:
            return
        vals = series_list[0].get("values", [])
        total = sum(vals) or 1
        cx, cy = width // 2, height // 2
        radius = min(width, height) // 3
        start = 0
        for i, v in enumerate(vals[:len(categories)]):
            angle = v / total * 360
            color = series_list[0].get("color") or CHART_COLORS[i % len(CHART_COLORS)]
            draw.pieslice([cx - radius, cy - radius, cx + radius, cy + radius], start=start, end=start + angle,
                          fill=color, outline="white")
            mid = start + angle / 2
            lx = cx + int(radius * 0.6 * math.cos(math.radians(mid)))
            ly = cy + int(radius * 0.6 * math.sin(math.radians(mid)))
            pct = f"{(v / total) * 100:.1f}%"
            draw.text((lx - 15, ly - 7), pct, fill="white", font=font)
            start += angle

    @staticmethod
    def _draw_line_chart(draw, categories, series_list, width, height, font, smooth=False):
        margin = 50
        legend_h = 40
        top, bottom = margin, height - margin - legend_h
        left, right = margin, width - margin
        if not categories or not series_list:
            return
        max_v = max((max(s.get("values", [0])) for s in series_list), default=0) or 1

        # 坐标轴
        draw.line([left, bottom, right, bottom], fill="#555", width=1)
        draw.line([left, top, left, bottom], fill="#555", width=1)

        # 网格线
        for i in range(5):
            y = top + (bottom - top) * i / 4
            draw.line([left, y, right, y], fill="#e5e5e5", width=1)

        # X轴标签
        step = (right - left) / (len(categories) - 1 if len(categories) > 1 else 1)
        for i, cat in enumerate(categories):
            x = left + i * step
            draw.text((x - 15, bottom + 5), str(cat)[:6], fill="#333", font=font)

        # 绘制线条
        for si, s in enumerate(series_list):
            color = s.get("color", CHART_COLORS[si % len(CHART_COLORS)])
            pts = []
            for ci, v in enumerate(s.get("values", [])[:len(categories)]):
                x = left + ci * step
                y = bottom - (v / max_v) * (bottom - top)
                pts.append((x, y))

            # 绘制线
            for i in range(len(pts) - 1):
                draw.line([pts[i], pts[i + 1]], fill=color, width=2)

            # 绘制点
            for x, y in pts:
                draw.ellipse([x - 3, y - 3, x + 3, y + 3], fill=color, outline=color)

    @staticmethod
    def _draw_bar_stacked_chart(draw, categories, series_list, width, height, font, percent=False):
        margin = 50
        legend_h = 40
        top, bottom = margin, height - margin - legend_h
        left, right = margin, width - margin
        if not categories or not series_list:
            return

        bar_w = (right - left) / (len(categories) * 1.25)
        gap = bar_w * 0.25

        totals = []
        for ci in range(len(categories)):
            t = 0
            for s in series_list:
                vals = s.get("values", [])
                if ci < len(vals):
                    t += vals[ci]
            totals.append(t or 1)

        max_total = max(totals) or 1

        for ci, cat in enumerate(categories):
            x0 = left + ci * (bar_w + gap)
            y_bottom = bottom

            for si, s in enumerate(series_list):
                vals = s.get("values", [])
                if ci >= len(vals):
                    continue
                v = vals[ci]

                if percent:
                    ratio = v / totals[ci] if totals[ci] else 0
                    h = ratio * (bottom - top)
                else:
                    h = (v / max_total) * (bottom - top)

                color = s.get("color", CHART_COLORS[si % len(CHART_COLORS)])
                y_top = y_bottom - h
                draw.rectangle([x0, y_top, x0 + bar_w, y_bottom], fill=color, outline=color)
                y_bottom = y_top

            # X轴标签
            draw.text((x0, bottom + 5), str(cat)[:6], fill="#333", font=font)

    @staticmethod
    def _draw_area_chart(draw, categories, series_list, width, height, font, stacked=False):
        margin = 50
        legend_h = 40
        top, bottom = margin, height - margin - legend_h
        left, right = margin, width - margin
        if not categories or not series_list:
            return

        if stacked:
            max_v = 0
            for ci in range(len(categories)):
                max_v = max(max_v, sum(
                    (s.get("values", [0])[ci] if ci < len(s.get("values", [])) else 0) for s in series_list))
        else:
            max_v = max((max(s.get("values", [0])) for s in series_list), default=0)
        max_v = max_v or 1

        # 坐标轴
        draw.line([left, bottom, right, bottom], fill="#555", width=1)
        draw.line([left, top, left, bottom], fill="#555", width=1)

        step = (right - left) / (len(categories) - 1 if len(categories) > 1 else 1)
        base = [0] * len(categories)

        for si, s in enumerate(series_list):
            color = s.get("color", CHART_COLORS[si % len(CHART_COLORS)])
            vals = s.get("values", [])
            line_pts = []
            poly = []

            for ci in range(len(categories)):
                v = vals[ci] if ci < len(vals) else 0
                acc = base[ci] + v if stacked else v
                x = left + ci * step
                y = bottom - (acc / max_v) * (bottom - top)
                line_pts.append((x, y))
                poly.append((x, y))

            # 闭合区域
            if stacked:
                for ci in reversed(range(len(categories))):
                    x = left + ci * step
                    y = bottom - (base[ci] / max_v) * (bottom - top)
                    poly.append((x, y))
            else:
                poly.extend([(right, bottom), (left, bottom)])

            # 半透明填充
            img_temp = Image.new('RGBA', (width, height), (255, 255, 255, 0))
            draw_temp = ImageDraw.Draw(img_temp)
            r, g, b = hex_to_rgb(color)[:3]
            draw_temp.polygon(poly, fill=(r, g, b, 80))
            draw.bitmap((0, 0), img_temp, fill=None)

            # 绘制线条
            for i in range(len(line_pts) - 1):
                draw.line([line_pts[i], line_pts[i + 1]], fill=color, width=2)

            if stacked:
                for ci in range(len(categories)):
                    v = vals[ci] if ci < len(vals) else 0
                    base[ci] += v

    @staticmethod
    def _draw_scatter_chart(draw, categories, series_list, width, height, font):
        margin = 50
        legend_h = 40
        top, bottom = margin, height - margin - legend_h
        left, right = margin, width - margin
        if not series_list:
            return

        # 坐标轴
        draw.line([left, bottom, right, bottom], fill="#555", width=1)
        draw.line([left, top, left, bottom], fill="#555", width=1)

        points = []
        for s in series_list:
            for itm in s.get("values", []):
                if isinstance(itm, (list, tuple)) and len(itm) >= 2:
                    points.append((float(itm[0]), float(itm[1])))
                elif isinstance(itm, dict) and 'x' in itm and 'y' in itm:
                    points.append((float(itm['x']), float(itm['y'])))

        if not points:
            return

        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        min_x, max_x = min(xs), max(xs)
        min_y, max_y = min(ys), max(ys)

        if max_x - min_x == 0: max_x += 1
        if max_y - min_y == 0: max_y += 1

        for si, s in enumerate(series_list):
            color = s.get("color", CHART_COLORS[si % len(CHART_COLORS)])
            for itm in s.get("values", []):
                if isinstance(itm, (list, tuple)) and len(itm) >= 2:
                    x_val, y_val = itm[0], itm[1]
                elif isinstance(itm, dict) and 'x' in itm and 'y' in itm:
                    x_val, y_val = itm['x'], itm['y']
                else:
                    continue

                x = left + (float(x_val) - min_x) / (max_x - min_x) * (right - left)
                y = bottom - (float(y_val) - min_y) / (max_y - min_y) * (bottom - top)
                draw.ellipse([x - 4, y - 4, x + 4, y + 4], fill=color, outline=color)

    @staticmethod
    def _draw_legend(draw, series_list, width, height, font):
        if not series_list:
            return
        base_y = height - 30
        x = 10
        for i, s in enumerate(series_list[:8]):
            color = s.get("color", CHART_COLORS[i % len(CHART_COLORS)])
            name = (s.get("name") or f"S{i + 1}")[:10]
            draw.rectangle([x, base_y, x + 14, base_y + 14], fill=color, outline=color)
            draw.text((x + 18, base_y - 1), name, fill="#222", font=font)
            x += 90


class PPTPreview(ttk.Frame):
    """增强的PPT预览组件"""

    def __init__(self, master: tk.Widget):
        super().__init__(master)

        self.preview_container = ttk.Frame(self, style="Preview.TFrame")
        self.preview_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        self.canvas = tk.Canvas(
            self.preview_container,
            bg=COLOR_CANVAS_BG,
            highlightthickness=1,
            highlightbackground=COLOR_BORDER,
            relief=tk.SOLID,
            bd=1
        )
        self.canvas.pack(expand=True)

        self.slide_images: List[ImageTk.PhotoImage] = []
        self.current_slide_index = 0
        self.slides: List[Dict[str, Any]] = []
        self.slide_size = (DEFAULT_WIDTH, DEFAULT_HEIGHT)
        self.current_meta: Optional[Dict[str, Any]] = None
        self.scale = 1.0
        self.chart_renderer = EnhancedChartRenderer()
        self.theme = {}  # 主题配置

    def set_meta(self, meta: Dict[str, Any]):
        """设置PPT元数据"""
        ppt_cfg = meta.get("ppt", {})
        size_cfg = ppt_cfg.get("size", {})
        width = float(size_cfg.get("width", DEFAULT_WIDTH))
        height = float(size_cfg.get("height", DEFAULT_HEIGHT))
        self.slide_size = (width, height)
        self.slides = ppt_cfg.get("slides", [])
        self.current_slide_index = max(0, min(self.current_slide_index, len(self.slides) - 1))
        self.current_meta = meta
        self.theme = ppt_cfg.get("theme", {})

        self.scale = min(PREVIEW_MAX_WIDTH / width, PREVIEW_MAX_HEIGHT / height, 1.0)

        logger.info(
            "预览设置: 幻灯片数=%s, 当前索引=%s, 尺寸=%s, 缩放=%s",
            len(self.slides),
            self.current_slide_index,
            self.slide_size,
            self.scale,
        )
        self.render()

    def resolve_color(self, color: str) -> str:
        """解析颜色值，支持主题变量"""
        if not color:
            return "#000000"

        if color.startswith('$'):
            var_name = color[1:]
            if self.theme and 'colors' in self.theme:
                return self.theme['colors'].get(var_name, color)

        return color

    @staticmethod
    def _resolve_box_px(box: Dict[str, Any], slide_w: float, slide_h: float, default_unit: str):
        unit = box.get("unit", default_unit or "px")

        def to_px(value: float, total: float) -> float:
            if value is None:
                return 0.0
            if unit == "percent":
                return total * float(value) / 100.0
            return float(value)

        width = to_px(box.get("w", slide_w), slide_w)
        height = to_px(box.get("h", slide_h), slide_h)
        return (
            to_px(box.get("x", 0), slide_w),
            to_px(box.get("y", 0), slide_h),
            width,
            height,
        )

    def create_gradient_image(self, width: int, height: int, gradient_cfg: Dict[str, Any]) -> Image.Image:
        """创建渐变图像"""
        img = Image.new('RGBA', (width, height), (255, 255, 255, 255))
        draw = ImageDraw.Draw(img)

        gradient_type = gradient_cfg.get("type", "linear")
        stops = gradient_cfg.get("stops", [])

        if not stops or len(stops) < 2:
            return img

        if gradient_type == "linear":
            angle = gradient_cfg.get("angle", 0)

            # 简化实现：水平或垂直渐变
            for y in range(height):
                ratio = y / height

                # 找到对应的颜色区间
                color = None
                for i in range(len(stops) - 1):
                    stop1 = stops[i]
                    stop2 = stops[i + 1]
                    pos1 = stop1.get("position", 0) / 100
                    pos2 = stop2.get("position", 100) / 100

                    if pos1 <= ratio <= pos2:
                        # 在两个停止点之间插值
                        local_ratio = (ratio - pos1) / (pos2 - pos1) if pos2 > pos1 else 0
                        color1 = hex_to_rgb(self.resolve_color(stop1.get("color", "#000000")))[:3]
                        color2 = hex_to_rgb(self.resolve_color(stop2.get("color", "#ffffff")))[:3]

                        r = int(color1[0] + (color2[0] - color1[0]) * local_ratio)
                        g = int(color1[1] + (color2[1] - color1[1]) * local_ratio)
                        b = int(color1[2] + (color2[2] - color1[2]) * local_ratio)
                        color = (r, g, b)
                        break

                if color:
                    draw.line([(0, y), (width, y)], fill=color)

        elif gradient_type == "radial":
            # 径向渐变（简化实现）
            cx, cy = width // 2, height // 2
            max_radius = math.sqrt(cx ** 2 + cy ** 2)

            for r in range(int(max_radius)):
                ratio = r / max_radius

                # 找到对应的颜色
                color = None
                for i in range(len(stops) - 1):
                    stop1 = stops[i]
                    stop2 = stops[i + 1]
                    pos1 = stop1.get("position", 0) / 100
                    pos2 = stop2.get("position", 100) / 100

                    if pos1 <= ratio <= pos2:
                        local_ratio = (ratio - pos1) / (pos2 - pos1) if pos2 > pos1 else 0
                        color1 = hex_to_rgb(self.resolve_color(stop1.get("color", "#000000")))[:3]
                        color2 = hex_to_rgb(self.resolve_color(stop2.get("color", "#ffffff")))[:3]

                        red = int(color1[0] + (color2[0] - color1[0]) * local_ratio)
                        green = int(color1[1] + (color2[1] - color1[1]) * local_ratio)
                        blue = int(color1[2] + (color2[2] - color1[2]) * local_ratio)
                        color = (red, green, blue)
                        break

                if color:
                    draw.ellipse([cx - r, cy - r, cx + r, cy + r], outline=color)

        return img

    def apply_shadow(self, img: Image.Image, shadow_cfg: Dict[str, Any]) -> Image.Image:
        """应用阴影效果"""
        if not shadow_cfg:
            return img

        x_offset = int(shadow_cfg.get("x", 2) * self.scale)
        y_offset = int(shadow_cfg.get("y", 2) * self.scale)
        blur = shadow_cfg.get("blur", 4)
        color = hex_to_rgb(self.resolve_color(shadow_cfg.get("color", "#00000040")))[:4]

        # 创建阴影层
        shadow = Image.new('RGBA', img.size, (0, 0, 0, 0))
        shadow.paste((color[0], color[1], color[2], color[3] if len(color) > 3 else 128),
                     [0, 0, img.width, img.height])

        # 应用模糊
        if blur > 0:
            shadow = shadow.filter(ImageFilter.GaussianBlur(radius=blur))

        # 合成
        final = Image.new('RGBA',
                          (img.width + abs(x_offset), img.height + abs(y_offset)),
                          (0, 0, 0, 0))
        final.paste(shadow, (max(0, x_offset), max(0, y_offset)))
        final.paste(img, (max(0, -x_offset), max(0, -y_offset)), img)

        return final

    def render_background(self, slide_cfg: Dict[str, Any]):
        """渲染背景，支持渐变"""
        self.canvas.delete("background")
        sw, sh = self.slide_size
        scaled_w = int(sw * self.scale)
        scaled_h = int(sh * self.scale)

        bg_cfg = slide_cfg.get("background", {}) if slide_cfg else {}

        # 优先渐变，其次纯色
        gradient = bg_cfg.get("gradient")
        if gradient:
            gradient_img = self.create_gradient_image(scaled_w, scaled_h, gradient)
            photo = ImageTk.PhotoImage(gradient_img)
            self.slide_images.append(photo)
            self.canvas.create_image(0, 0, anchor="nw", image=photo, tags="background")
        else:
            color = self.resolve_color(bg_cfg.get("color", "#ffffff"))
            self.canvas.create_rectangle(
                0, 0, scaled_w, scaled_h,
                fill=color, outline="", tags="background"
            )

        # 背景图片
        img_cfg = bg_cfg.get("image") if isinstance(bg_cfg, dict) else None
        if img_cfg and isinstance(img_cfg, dict):
            content = get_image_bytes(img_cfg.get("src"), logger)
            if content:
                try:
                    image = Image.open(BytesIO(content))
                    image = image.resize((scaled_w, scaled_h), Image.Resampling.LANCZOS)

                    # 应用透明度
                    if "opacity" in img_cfg:
                        opacity = img_cfg["opacity"]
                        if 0 <= opacity < 1:
                            image = image.convert("RGBA")
                            alpha = image.split()[-1]
                            alpha = alpha.point(lambda p: p * opacity)
                            image.putalpha(alpha)

                    # 应用滤镜
                    filter_type = img_cfg.get("filter")
                    if filter_type == "blur":
                        blur_val = img_cfg.get("blur", 5)
                        image = image.filter(ImageFilter.GaussianBlur(radius=blur_val))
                    elif filter_type == "grayscale":
                        image = ImageOps.grayscale(image)

                    photo = ImageTk.PhotoImage(image=image)
                    self.slide_images.append(photo)
                    self.canvas.create_image(0, 0, anchor="nw", image=photo, tags="background")
                except Exception as exc:
                    logger.warning("背景图片绘制失败: %s", exc)

    def draw_text(self, elem: Dict[str, Any], default_unit: str):
        """绘制文本，支持阴影和旋转"""
        slide_w, slide_h = self.slide_size
        x, y, width, height = self._resolve_box_px(elem.get("box", {}), slide_w, slide_h, default_unit)

        x *= self.scale
        y *= self.scale
        width *= self.scale
        height *= self.scale

        text = elem.get("text", "")
        style = elem.get("style", {})

        # 处理段落
        paragraphs = elem.get("paragraphs")
        if paragraphs:
            texts = []
            for para in paragraphs:
                para_text = para.get("text", "")
                list_type = para.get("listType")
                if list_type == "bullet":
                    bullet_char = para.get("bulletChar", "•")
                    para_text = f"{bullet_char} {para_text}"
                elif list_type == "number":
                    number = para.get("number", len(texts) + 1)
                    prefix = para.get("numberPrefix", "")
                    para_text = f"{prefix}{number}. {para_text}"
                texts.append(para_text)
            text = "\n".join(texts)

        font_size = int(style.get("fontSize", 32) * self.scale * 0.75)
        font_family = style.get("fontFamily", "Microsoft YaHei")

        # 支持主题字体
        if font_family.startswith('$'):
            if self.theme and 'fonts' in self.theme:
                if font_family == '$heading':
                    font_family = self.theme['fonts'].get('heading', 'Arial')
                elif font_family == '$body':
                    font_family = self.theme['fonts'].get('body', 'Arial')

        fill = self.resolve_color(style.get("color", "#000000"))
        align = style.get("align", "left")
        bold = style.get("bold", False)
        italic = style.get("italic", False)

        font_style = []
        if bold:
            font_style.append("bold")
        if italic:
            font_style.append("italic")
        font_tuple = (font_family, max(8, font_size), " ".join(font_style) if font_style else "normal")

        # 背景色
        if elem.get("fill"):
            self.canvas.create_rectangle(
                x, y, x + width, y + height,
                fill=self.resolve_color(elem["fill"]),
                outline="",
                tags="element"
            )

        # 边框
        if elem.get("border"):
            self.draw_border(x, y, width, height, elem["border"])

        if align == "center":
            anchor = "n"
            text_x = x + width / 2
        elif align == "right":
            anchor = "ne"
            text_x = x + width
        else:
            anchor = "nw"
            text_x = x

        # 文本阴影
        shadow = elem.get("shadow") or style.get("textShadow")
        if shadow:
            shadow_x = shadow.get("x", 2) * self.scale
            shadow_y = shadow.get("y", 2) * self.scale
            shadow_color = self.resolve_color(shadow.get("color", "#00000040"))
            self.canvas.create_text(
                text_x + shadow_x, y + shadow_y,
                text=text,
                font=font_tuple,
                fill=shadow_color,
                anchor=anchor,
                width=width,
                tags="element",
                justify={"center": "center", "right": "right"}.get(align, "left")
            )

        # 主文本
        text_item = self.canvas.create_text(
            text_x, y,
            text=text,
            font=font_tuple,
            fill=fill,
            anchor=anchor,
            width=width,
            tags="element",
            justify={"center": "center", "right": "right"}.get(align, "left")
        )

        # 旋转
        if "rotation" in elem:
            self.apply_rotation(text_item, elem["rotation"], x + width / 2, y + height / 2)

    def draw_border(self, x: float, y: float, width: float, height: float, border_cfg: Dict[str, Any]):
        """绘制边框"""
        if not border_cfg:
            return

        border_width = border_cfg.get("width", 1)
        border_color = self.resolve_color(border_cfg.get("color", "#000000"))
        border_style = border_cfg.get("style", "solid")

        dash_pattern = None
        if border_style == "dashed":
            dash_pattern = (5, 2)
        elif border_style == "dotted":
            dash_pattern = (2, 2)

        self.canvas.create_rectangle(
            x, y, x + width, y + height,
            outline=border_color,
            width=border_width,
            dash=dash_pattern,
            fill="",
            tags="element"
        )

    def apply_rotation(self, item_id, angle: float, cx: float, cy: float):
        """应用旋转（简化实现）"""
        # Canvas不直接支持旋转，这里仅作标记
        pass

    def draw_shape(self, elem: Dict[str, Any], default_unit: str):
        """绘制形状，支持更多类型"""
        slide_w, slide_h = self.slide_size
        x, y, width, height = self._resolve_box_px(elem.get("box", {}), slide_w, slide_h, default_unit)

        x *= self.scale
        y *= self.scale
        width *= self.scale
        height *= self.scale

        shape_type = elem.get("shapeType", "rect")
        fill = self.resolve_color(elem.get("fill", "#d1d5db"))

        # 渐变填充
        gradient = elem.get("gradient")
        if gradient:
            gradient_img = self.create_gradient_image(int(width), int(height), gradient)
            photo = ImageTk.PhotoImage(gradient_img)
            self.slide_images.append(photo)
            self.canvas.create_image(x, y, anchor="nw", image=photo, tags="element")

            # 边框
            if elem.get("border"):
                self.draw_border(x, y, width, height, elem["border"])
            return

        # 阴影
        shadow = elem.get("shadow")
        if shadow:
            shadow_x = shadow.get("x", 2) * self.scale
            shadow_y = shadow.get("y", 2) * self.scale
            shadow_blur = shadow.get("blur", 4)
            shadow_color = self.resolve_color(shadow.get("color", "#00000030"))

            # 绘制阴影形状
            self._draw_shape_primitive(
                x + shadow_x, y + shadow_y, width, height,
                shape_type, shadow_color, None, "shadow"
            )

        # 主形状
        shape_id = self._draw_shape_primitive(x, y, width, height, shape_type, fill, elem.get("border"), "element")

        # 旋转
        if "rotation" in elem and shape_id:
            self.apply_rotation(shape_id, elem["rotation"], x + width / 2, y + height / 2)

    def _draw_shape_primitive(self, x: float, y: float, width: float, height: float,
                              shape_type: str, fill: str, border: Optional[Dict], tags: str):
        """绘制基本形状"""
        if shape_type in ["ellipse", "circle"]:
            if shape_type == "circle":
                size = min(width, height)
                width = height = size
            return self.canvas.create_oval(
                x, y, x + width, y + height,
                fill=fill, outline="", tags=tags
            )
        elif shape_type == "triangle":
            points = [x + width / 2, y, x, y + height, x + width, y + height]
            return self.canvas.create_polygon(points, fill=fill, outline="", tags=tags)
        elif shape_type == "star" or shape_type.startswith("star"):
            # 五角星
            cx, cy = x + width / 2, y + height / 2
            outer_r = min(width, height) / 2
            inner_r = outer_r * 0.4
            points = []
            n = 5  # 默认五角星
            if shape_type == "star6":
                n = 6
            elif shape_type == "star8":
                n = 8

            for i in range(n * 2):
                angle = math.pi * i / n - math.pi / 2
                r = outer_r if i % 2 == 0 else inner_r
                px = cx + r * math.cos(angle)
                py = cy + r * math.sin(angle)
                points.extend([px, py])

            return self.canvas.create_polygon(points, fill=fill, outline="", tags=tags)
        elif shape_type == "hexagon":
            cx, cy = x + width / 2, y + height / 2
            r = min(width, height) / 2
            points = []
            for i in range(6):
                angle = math.pi * i / 3
                px = cx + r * math.cos(angle)
                py = cy + r * math.sin(angle)
                points.extend([px, py])
            return self.canvas.create_polygon(points, fill=fill, outline="", tags=tags)
        elif shape_type == "diamond":
            points = [x + width / 2, y, x + width, y + height / 2, x + width / 2, y + height, x, y + height / 2]
            return self.canvas.create_polygon(points, fill=fill, outline="", tags=tags)
        elif shape_type in ["arrow", "arrowRight"]:
            points = [x, y + height * 0.3, x + width * 0.6, y + height * 0.3,
                      x + width * 0.6, y, x + width, y + height / 2,
                      x + width * 0.6, y + height, x + width * 0.6, y + height * 0.7,
                      x, y + height * 0.7]
            return self.canvas.create_polygon(points, fill=fill, outline="", tags=tags)
        elif shape_type == "heart":
            # 简化的心形
            points = []
            for t in range(0, 360, 10):
                rad = math.radians(t)
                px = 16 * math.sin(rad) ** 3
                py = -(13 * math.cos(rad) - 5 * math.cos(2 * rad) - 2 * math.cos(3 * rad) - math.cos(4 * rad))
                points.extend([x + width / 2 + px * width / 32, y + height / 2 + py * height / 32])
            return self.canvas.create_polygon(points, fill=fill, outline="", tags=tags, smooth=True)
        elif shape_type == "plus":
            # 十字
            w3 = width / 3
            h3 = height / 3
            points = [
                x + w3, y, x + 2 * w3, y, x + 2 * w3, y + h3,
                x + width, y + h3, x + width, y + 2 * h3,
                x + 2 * w3, y + 2 * h3, x + 2 * w3, y + height,
                x + w3, y + height, x + w3, y + 2 * h3,
                x, y + 2 * h3, x, y + h3, x + w3, y + h3
            ]
            return self.canvas.create_polygon(points, fill=fill, outline="", tags=tags)
        else:
            # 默认矩形
            shape_id = self.canvas.create_rectangle(
                x, y, x + width, y + height,
                fill=fill, outline="", tags=tags
            )

            # 边框
            if border:
                self.draw_border(x, y, width, height, border)

            return shape_id

    def draw_line(self, elem: Dict[str, Any], default_unit: str):
        """绘制线条元素"""
        slide_w, slide_h = self.slide_size
        points = elem.get("points", [])

        if len(points) < 2:
            return

        stroke = self.resolve_color(elem.get("stroke", "#000000"))
        stroke_width = elem.get("strokeWidth", 1) * self.scale
        stroke_style = elem.get("strokeStyle", "solid")

        dash_pattern = None
        if stroke_style == "dashed":
            dash_pattern = (5, 2)
        elif stroke_style == "dotted":
            dash_pattern = (2, 2)

        # 转换点坐标
        converted_points = []
        for point in points:
            px = point.get("x", 0) * self.scale
            py = point.get("y", 0) * self.scale
            converted_points.extend([px, py])

        # 绘制线条
        if elem.get("curved"):
            self.canvas.create_line(
                converted_points,
                fill=stroke,
                width=stroke_width,
                dash=dash_pattern,
                smooth=True,
                tags="element"
            )
        else:
            self.canvas.create_line(
                converted_points,
                fill=stroke,
                width=stroke_width,
                dash=dash_pattern,
                tags="element"
            )

        # 箭头
        if len(points) >= 2:
            start_arrow = elem.get("startArrow", "none")
            end_arrow = elem.get("endArrow", "none")

            if end_arrow == "arrow":
                # 绘制箭头
                p1 = points[-2]
                p2 = points[-1]
                self._draw_arrow_head(
                    p1.get("x", 0) * self.scale,
                    p1.get("y", 0) * self.scale,
                    p2.get("x", 0) * self.scale,
                    p2.get("y", 0) * self.scale,
                    stroke
                )

    def _draw_arrow_head(self, x1: float, y1: float, x2: float, y2: float, color: str):
        """绘制箭头头部"""
        angle = math.atan2(y2 - y1, x2 - x1)
        arrow_length = 10 * self.scale
        arrow_angle = math.pi / 6

        # 计算箭头两个点
        x3 = x2 - arrow_length * math.cos(angle - arrow_angle)
        y3 = y2 - arrow_length * math.sin(angle - arrow_angle)
        x4 = x2 - arrow_length * math.cos(angle + arrow_angle)
        y4 = y2 - arrow_length * math.sin(angle + arrow_angle)

        self.canvas.create_polygon(
            [x2, y2, x3, y3, x4, y4],
            fill=color,
            outline=color,
            tags="element"
        )

    def draw_icon(self, elem: Dict[str, Any], default_unit: str):
        """绘制图标元素"""
        slide_w, slide_h = self.slide_size
        x, y, width, height = self._resolve_box_px(elem.get("box", {}), slide_w, slide_h, default_unit)

        x *= self.scale
        y *= self.scale
        width *= self.scale
        height *= self.scale

        # 使用形状模拟图标
        color = self.resolve_color(elem.get("color", "#000000"))
        gradient = elem.get("gradient")

        if gradient:
            gradient_img = self.create_gradient_image(int(width), int(height), gradient)
            photo = ImageTk.PhotoImage(gradient_img)
            self.slide_images.append(photo)
            self.canvas.create_image(x, y, anchor="nw", image=photo, tags="element")
        else:
            # 绘制一个简单的星形作为图标
            cx, cy = x + width / 2, y + height / 2
            r = min(width, height) / 2
            points = []
            for i in range(10):
                angle = math.pi * i / 5 - math.pi / 2
                radius = r if i % 2 == 0 else r * 0.5
                px = cx + radius * math.cos(angle)
                py = cy + radius * math.sin(angle)
                points.extend([px, py])

            self.canvas.create_polygon(points, fill=color, outline="", tags="element")

    def draw_group(self, elem: Dict[str, Any], default_unit: str):
        """绘制组元素"""
        elements = elem.get("elements", [])

        # 递归绘制子元素
        for sub_elem in elements:
            self.draw_element(sub_elem, default_unit)

    def draw_video(self, elem: Dict[str, Any], default_unit: str):
        """绘制视频占位符"""
        slide_w, slide_h = self.slide_size
        x, y, width, height = self._resolve_box_px(elem.get("box", {}), slide_w, slide_h, default_unit)

        x *= self.scale
        y *= self.scale
        width *= self.scale
        height *= self.scale

        # 绘制视频占位符
        self.canvas.create_rectangle(
            x, y, x + width, y + height,
            fill="#000000",
            outline=COLOR_BORDER,
            width=2,
            tags="element"
        )

        # 播放按钮
        cx, cy = x + width / 2, y + height / 2
        r = min(width, height) / 8
        self.canvas.create_oval(
            cx - r, cy - r, cx + r, cy + r,
            fill="#ffffff",
            outline="",
            tags="element"
        )

        # 三角形播放图标
        points = [
            cx - r * 0.3, cy - r * 0.5,
            cx - r * 0.3, cy + r * 0.5,
            cx + r * 0.5, cy
        ]
        self.canvas.create_polygon(points, fill="#000000", outline="", tags="element")

    def draw_smartart(self, elem: Dict[str, Any], default_unit: str):
        """绘制SmartArt"""
        slide_w, slide_h = self.slide_size
        x, y, width, height = self._resolve_box_px(elem.get("box", {}), slide_w, slide_h, default_unit)

        x *= self.scale
        y *= self.scale
        width *= self.scale
        height *= self.scale

        smartart_type = elem.get("smartArtType", "list")
        nodes = elem.get("nodes", [])

        if not nodes:
            return

        if smartart_type in ["list", "process"]:
            # 列表或流程图
            node_width = width * 0.8 / len(nodes)
            node_height = height * 0.6

            for i, node in enumerate(nodes):
                node_x = x + i * (width / len(nodes)) + width * 0.1 / len(nodes)
                node_y = y + height * 0.2

                # 节点颜色
                node_color = self.resolve_color(node.get("color", "#3B82F6"))

                # 绘制节点
                self.canvas.create_rectangle(
                    node_x, node_y,
                    node_x + node_width, node_y + node_height,
                    fill=node_color,
                    outline="",
                    tags="element"
                )

                # 节点文本
                text = node.get("text", f"Item {i + 1}")
                self.canvas.create_text(
                    node_x + node_width / 2,
                    node_y + node_height / 2,
                    text=text,
                    font=("Microsoft YaHei", max(10, int(14 * self.scale))),
                    fill="white",
                    anchor="center",
                    width=node_width * 0.9,
                    tags="element"
                )

                # 连接线
                if i < len(nodes) - 1:
                    self.canvas.create_line(
                        node_x + node_width, node_y + node_height / 2,
                        node_x + node_width + width * 0.1 / len(nodes), node_y + node_height / 2,
                        fill="#666666",
                        width=2,
                        arrow=tk.LAST,
                        tags="element"
                    )

        elif smartart_type == "cycle":
            # 循环图
            cx, cy = x + width / 2, y + height / 2
            radius = min(width, height) / 3

            for i, node in enumerate(nodes):
                angle = 2 * math.pi * i / len(nodes) - math.pi / 2
                node_x = cx + radius * math.cos(angle)
                node_y = cy + radius * math.sin(angle)

                node_color = self.resolve_color(node.get("color", CHART_COLORS[i % len(CHART_COLORS)]))

                # 绘制节点
                node_r = min(width, height) / 10
                self.canvas.create_oval(
                    node_x - node_r, node_y - node_r,
                    node_x + node_r, node_y + node_r,
                    fill=node_color,
                    outline="",
                    tags="element"
                )

                # 节点文本
                text = node.get("text", f"Item {i + 1}")
                self.canvas.create_text(
                    node_x, node_y,
                    text=text[:5],
                    font=("Microsoft YaHei", max(9, int(11 * self.scale))),
                    fill="white",
                    anchor="center",
                    tags="element"
                )

    def draw_image(self, elem: Dict[str, Any], default_unit: str):
        """绘制图片，支持滤镜"""
        slide_w, slide_h = self.slide_size
        x, y, width, height = self._resolve_box_px(elem.get("box", {}), slide_w, slide_h, default_unit)

        x *= self.scale
        y *= self.scale
        width *= self.scale
        height *= self.scale

        content = get_image_bytes(elem.get("source"), logger)
        if not content:
            # 占位符
            self.canvas.create_rectangle(
                x, y, x + width, y + height,
                outline=COLOR_BORDER, dash=(4, 2), width=2, tags="element"
            )
            self.canvas.create_text(
                x + width / 2, y + height / 2,
                text="🖼️ 图片",
                font=("Microsoft YaHei", max(10, int(14 * self.scale))),
                fill=COLOR_TEXT_LIGHT,
                anchor="center",
                tags="element",
            )
            return

        try:
            img = Image.open(BytesIO(content))
            img = img.resize((int(width), int(height)), Image.Resampling.LANCZOS)

            # 应用滤镜
            filter_type = elem.get("filter")
            if filter_type == "grayscale":
                img = ImageOps.grayscale(img)
            elif filter_type == "sepia":
                # 简单的棕褐色效果
                img = ImageOps.grayscale(img)
                img = ImageOps.colorize(img, '#704214', '#f0e68c')
            elif filter_type == "blur":
                filter_value = elem.get("filterValue", 5)
                img = img.filter(ImageFilter.GaussianBlur(radius=filter_value))

            photo = ImageTk.PhotoImage(image=img)
            self.slide_images.append(photo)
            self.canvas.create_image(x, y, anchor="nw", image=photo, tags="element")

            # 边框
            if elem.get("border"):
                self.draw_border(x, y, width, height, elem["border"])

        except Exception as exc:
            logger.warning("图片渲染失败: %s", exc)

    def draw_chart(self, elem: Dict[str, Any], default_unit: str):
        """渲染图表预览"""
        slide_w, slide_h = self.slide_size
        x, y, width, height = self._resolve_box_px(elem.get("box", {}), slide_w, slide_h, default_unit)

        x *= self.scale
        y *= self.scale
        width *= self.scale
        height *= self.scale

        chart_type = elem.get("chartType", "bar")
        title = elem.get("title", "")
        raw_data = elem.get("data", {}) or {}
        options = elem.get("chartOptions", {})

        # 兼容别名 -> 内部实际渲染类型
        alias_map = {
            "lineSmooth": "line",
            "barStacked": "bar",
            "barStacked100": "bar",
            "area": "line",          # 使用折线 + fill 模拟
            "areaStacked": "area",   # 交给内部逻辑（如果支持），否则退化
        }
        mapped_type = alias_map.get(chart_type, chart_type)
        chart_type = mapped_type

        # 深拷贝并解析颜色（支持主题变量 $primary 等），防止 PIL 出现 unknown color specifier
        try:
            data = copy.deepcopy(raw_data)
            series_list = data.get("series") or []
            for s in series_list:
                # 解析直接颜色
                if isinstance(s, dict) and s.get("color"):
                    s["color"] = self.resolve_color(s.get("color"))
                # 解析渐变颜色 stops
                gradient = s.get("gradient") if isinstance(s, dict) else None
                if isinstance(gradient, dict):
                    stops = gradient.get("stops")
                    if isinstance(stops, list):
                        for stop in stops:
                            if isinstance(stop, dict) and stop.get("color"):
                                stop["color"] = self.resolve_color(stop.get("color"))
        except Exception as color_exc:
            logger.debug(f"chart color resolve skipped: {color_exc}")
            data = raw_data  # 退回原始数据（失败时仍尝试渲染）

        try:
            # 使用增强的图表渲染器
            chart_img = self.chart_renderer.render_chart(
                chart_type, data, int(width), int(height), title, options
            )

            photo = ImageTk.PhotoImage(chart_img)
            self.slide_images.append(photo)
            self.canvas.create_image(x, y, anchor="nw", image=photo, tags="element")

        except Exception as e:
            logger.error(f"图表渲染失败: {e}")
            # 显示占位符
            self.canvas.create_rectangle(
                x, y, x + width, y + height,
                fill="#f8f9fb", outline=COLOR_ACCENT, width=2, tags="element"
            )

            icon = {
                "bar": "📊", "line": "📈", "pie": "🥧", "area": "📉",
                "doughnut": "🍩", "radar": "🎯", "bubble": "🫧", "scatter": "📍"
            }.get(chart_type, "📊")

            self.canvas.create_text(
                x + width / 2, y + height / 2,
                text=f"{icon}\n{chart_type.title()} Chart\n(预览)",
                font=("Microsoft YaHei", max(10, int(14 * self.scale))),
                fill=COLOR_TEXT_LIGHT,
                anchor="center",
                tags="element",
            )

    def draw_table(self, elem: Dict[str, Any], default_unit: str):
        """绘制表格，支持更多样式"""
        slide_w, slide_h = self.slide_size
        x, y, width, height = self._resolve_box_px(elem.get("box", {}), slide_w, slide_h, default_unit)

        x *= self.scale
        y *= self.scale
        width *= self.scale
        height *= self.scale

        table_cfg = elem.get("table", {})
        header = table_cfg.get("header") or []
        rows = table_cfg.get("rows") or []

        rows_count = len(rows) + (1 if header else 0)
        cols_count = len(header) if header else (len(rows[0]) if rows else 0)

        if rows_count == 0 or cols_count == 0:
            return

        cell_width = width / cols_count
        cell_height = height / rows_count

        all_rows: List[List[Any]] = []
        if header:
            all_rows.append(header)
        all_rows.extend(rows)

        # 表格样式
        style_cfg = table_cfg.get("style", {})
        banded_rows = table_cfg.get("bandedRows", False)
        banded_cols = table_cfg.get("bandedColumns", False)

        for i, row in enumerate(all_rows):
            for j in range(cols_count):
                cell_x0 = x + j * cell_width
                cell_y0 = y + i * cell_height
                cell_x1 = cell_x0 + cell_width
                cell_y1 = cell_y0 + cell_height

                # 单元格背景色
                if i == 0 and header:
                    # 表头
                    bg_color = "#f3f4f6"
                    if style_cfg.get("header", {}).get("fill"):
                        bg_color = self.resolve_color(style_cfg["header"]["fill"])
                    font = ("Microsoft YaHei", max(10, int(12 * self.scale)), "bold")
                    text_color = self.resolve_color(style_cfg.get("header", {}).get("color", COLOR_TEXT))
                else:
                    # 表体
                    bg_color = "white"
                    if banded_rows and (i - (1 if header else 0)) % 2 == 1:
                        bg_color = "#f9fafb"
                    if banded_cols and j % 2 == 1:
                        bg_color = "#f9fafb"
                    if style_cfg.get("body", {}).get("fill"):
                        bg_color = self.resolve_color(style_cfg["body"]["fill"])
                    font = ("Microsoft YaHei", max(9, int(11 * self.scale)))
                    text_color = self.resolve_color(style_cfg.get("body", {}).get("color", COLOR_TEXT_LIGHT))

                self.canvas.create_rectangle(
                    cell_x0, cell_y0, cell_x1, cell_y1,
                    fill=bg_color, outline=COLOR_BORDER, width=1, tags="element"
                )

                text = str(row[j]) if j < len(row) else ""
                self.canvas.create_text(
                    (cell_x0 + cell_x1) / 2,
                    (cell_y0 + cell_y1) / 2,
                    text=text,
                    font=font,
                    fill=text_color,
                    anchor="center",
                    width=cell_width - 10,
                    tags="element",
                )

    def draw_element(self, elem: Dict[str, Any], default_unit: str):
        """绘制元素的统一入口"""
        elem_type = elem.get("type")
        try:
            if elem_type == "text":
                self.draw_text(elem, default_unit)
            elif elem_type == "image":
                self.draw_image(elem, default_unit)
            elif elem_type == "shape":
                self.draw_shape(elem, default_unit)
            elif elem_type == "chart":
                self.draw_chart(elem, default_unit)
            elif elem_type == "table":
                self.draw_table(elem, default_unit)
            elif elem_type == "line":
                self.draw_line(elem, default_unit)
            elif elem_type == "icon":
                self.draw_icon(elem, default_unit)
            elif elem_type == "group":
                self.draw_group(elem, default_unit)
            elif elem_type == "video":
                self.draw_video(elem, default_unit)
            elif elem_type == "smartArt":
                self.draw_smartart(elem, default_unit)
        except Exception as e:
            logger.error(f"元素渲染失败: type={elem_type}, error={e}")

    def render_slide(self, slide_cfg: Dict[str, Any]):
        """渲染单个幻灯片"""
        self.canvas.delete("all")
        self.slide_images.clear()

        if not slide_cfg:
            self.canvas.create_text(
                PREVIEW_MAX_WIDTH / 2,
                PREVIEW_MAX_HEIGHT / 2,
                text="没有幻灯片",
                font=("Microsoft YaHei", 16),
                fill=COLOR_TEXT_LIGHT,
            )
            return

        slide_w, slide_h = self.slide_size
        scaled_w = slide_w * self.scale
        scaled_h = slide_h * self.scale

        self.canvas.config(width=scaled_w, height=scaled_h)

        # 渲染背景
        self.render_background(slide_cfg)

        default_unit = self.current_meta.get("ppt", {}).get("defaultUnit", "px") if self.current_meta else "px"

        # 获取所有元素并按zIndex排序
        elements = slide_cfg.get("elements", [])
        sorted_elements = sorted(elements, key=lambda e: e.get("zIndex", 0))

        # 渲染所有元素
        for elem in sorted_elements:
            self.draw_element(elem, default_unit)

        # 显示过渡效果信息（如果有）
        transition = slide_cfg.get("transition")
        if transition and transition.get("type") != "none":
            transition_type = transition.get("type", "fade")
            duration = transition.get("duration", 1)
            # 在角落显示过渡效果标记
            self.canvas.create_text(
                scaled_w - 10, scaled_h - 10,
                text=f"🎬 {transition_type} ({duration}s)",
                font=("Microsoft YaHei", 9),
                fill=COLOR_TEXT_LIGHT,
                anchor="se",
                tags="transition_info"
            )

    def render(self):
        """渲染当前幻灯片"""
        if not self.slides:
            self.canvas.delete("all")
            self.canvas.config(width=PREVIEW_MAX_WIDTH, height=PREVIEW_MAX_HEIGHT / 2)
            self.canvas.create_text(
                PREVIEW_MAX_WIDTH / 2,
                PREVIEW_MAX_HEIGHT / 4,
                text="📝 添加幻灯片以查看预览",
                font=("Microsoft YaHei", 18),
                fill=COLOR_TEXT_LIGHT,
            )
            return

        slide_cfg = self.slides[self.current_slide_index]
        self.render_slide(slide_cfg)

    def next_slide(self):
        """切换到下一张幻灯片"""
        if self.current_slide_index < len(self.slides) - 1:
            self.current_slide_index += 1
            logger.info("切换到下一页: 索引=%s", self.current_slide_index)
            self.render()

    def prev_slide(self):
        """切换到上一张幻灯片"""
        if self.current_slide_index > 0:
            self.current_slide_index -= 1
            logger.info("切换到上一页: 索引=%s", self.current_slide_index)
            self.render()


class JSONPPTApp:
    """增强的JSON到PPT设计器应用"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("✨ JSON → PPT Designer (增强版)")
        self.root.geometry("1600x900")
        self.root.minsize(1200, 700)

        # 设置应用图标（如果可能）
        try:
            self.root.iconbitmap(default='icon.ico')
        except:
            pass

        self.setup_styles()

        self.preview = None
        self.editor: Optional[ScrolledText] = None
        self.status_var = tk.StringVar()
        self.slide_info = tk.StringVar()
        self.render_job = None
        self.current_meta: Optional[Dict[str, Any]] = None
        self.error_line = None  # 记录错误行号

        self.setup_layout()
        self.setup_keyboard_shortcuts()
        self.load_sample()

    def setup_styles(self):
        """设置UI样式"""
        style = ttk.Style()
        style.theme_use('clam')

        # 基础样式
        style.configure(".", background=COLOR_BG, foreground=COLOR_TEXT)
        style.configure("Title.TLabel", font=("Microsoft YaHei", 12, "bold"))
        style.configure("Preview.TFrame", background=COLOR_BG)
        style.configure("Toolbar.TFrame", background=COLOR_SIDEBAR, relief="flat")
        style.configure("Status.TFrame", background=COLOR_SIDEBAR)
        style.configure("Status.TLabel", background=COLOR_SIDEBAR, foreground="white")

        # 按钮样式
        style.configure("Toolbar.TButton",
                        background="#34495e",
                        foreground="white",
                        borderwidth=0,
                        focuscolor='none')
        style.map("Toolbar.TButton",
                  background=[('active', '#4a5f7a')])

    def setup_layout(self):
        """设置UI布局"""
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True)

        self.create_toolbar(main_container)

        # 创建主分割面板
        paned = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        # 左侧编辑器
        editor_frame = ttk.Frame(paned)
        paned.add(editor_frame, weight=1)

        editor_header = ttk.Frame(editor_frame)
        editor_header.pack(fill=tk.X, padx=10, pady=(10, 5))
        ttk.Label(editor_header, text="📝 JSON 编辑器", style="Title.TLabel").pack(side=tk.LEFT)

        # 编辑器工具栏
        editor_tools = ttk.Frame(editor_header)
        editor_tools.pack(side=tk.RIGHT)
        ttk.Button(editor_tools, text="🔍 查找", command=self.find_in_editor, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(editor_tools, text="↩️ 撤销", command=self.undo_edit, width=8).pack(side=tk.LEFT, padx=2)

        editor_container = ttk.Frame(editor_frame)
        editor_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 行号显示
        line_number_frame = ttk.Frame(editor_container)
        line_number_frame.pack(side=tk.LEFT, fill=tk.Y)

        self.line_numbers = tk.Text(
            line_number_frame,
            width=4,
            padx=3,
            takefocus=0,
            border=0,
            state='disabled',
            wrap='none',
            background='#34495e',
            foreground='#95a5a6'
        )
        self.line_numbers.pack(fill=tk.Y)

        # JSON编辑器
        self.editor = ScrolledText(
            editor_container,
            wrap=tk.NONE,
            font=("Consolas", 11),
            bg="#2d2d30",
            fg="#d4d4d4",
            insertbackground="#ffffff",
            selectbackground="#264f78",
            selectforeground="#ffffff",
            relief=tk.FLAT,
            borderwidth=0,
            undo=True,
            maxundo=-1
        )
        self.editor.pack(fill=tk.BOTH, expand=True)
        self.editor.bind("<<Modified>>", self.on_editor_modified)
        self.editor.bind("<KeyRelease>", self.update_line_numbers)
        self.editor.bind("<MouseWheel>", self.sync_scroll)

        # 右侧预览
        preview_frame = ttk.Frame(paned)
        paned.add(preview_frame, weight=1)

        preview_header = ttk.Frame(preview_frame)
        preview_header.pack(fill=tk.X, padx=10, pady=(10, 5))
        ttk.Label(preview_header, text="👁️ 实时预览", style="Title.TLabel").pack(side=tk.LEFT)

        # 导航控制
        nav_frame = ttk.Frame(preview_header)
        nav_frame.pack(side=tk.RIGHT)

        ttk.Button(nav_frame, text="⏮️", command=self.first_slide, width=3).pack(side=tk.LEFT, padx=1)
        ttk.Button(nav_frame, text="◀", command=self.prev_slide, width=3).pack(side=tk.LEFT, padx=1)
        self.slide_label = ttk.Label(nav_frame, textvariable=self.slide_info, font=("Microsoft YaHei", 10, "bold"))
        self.slide_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(nav_frame, text="▶", command=self.next_slide, width=3).pack(side=tk.LEFT, padx=1)
        ttk.Button(nav_frame, text="⏭️", command=self.last_slide, width=3).pack(side=tk.LEFT, padx=1)

        # 缩放控制
        ttk.Separator(nav_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Label(nav_frame, text="🔍").pack(side=tk.LEFT)
        self.zoom_var = tk.StringVar(value="适应")
        zoom_combo = ttk.Combobox(nav_frame, textvariable=self.zoom_var, width=8, state="readonly")
        zoom_combo['values'] = ["50%", "75%", "100%", "125%", "150%", "适应"]
        zoom_combo.pack(side=tk.LEFT, padx=2)
        zoom_combo.bind("<<ComboboxSelected>>", self.on_zoom_changed)

        self.preview = PPTPreview(preview_frame)
        self.preview.pack(fill=tk.BOTH, expand=True)

        self.create_statusbar()

    def create_toolbar(self, parent):
        """创建工具栏"""
        toolbar = ttk.Frame(parent, style="Toolbar.TFrame", height=50)
        toolbar.pack(fill=tk.X)
        toolbar.pack_propagate(False)

        btn_frame = ttk.Frame(toolbar, style="Toolbar.TFrame")
        btn_frame.pack(side=tk.LEFT, padx=10, pady=8)

        # 文件操作
        ttk.Button(btn_frame, text="📂 打开", command=self.open_file, style="Toolbar.TButton").pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="💾 保存", command=self.save_file, style="Toolbar.TButton").pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="📄 新建", command=self.new_file, style="Toolbar.TButton").pack(side=tk.LEFT, padx=3)

        ttk.Separator(btn_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        # 编辑操作
        ttk.Button(btn_frame, text="🎯 格式化", command=self.format_json, style="Toolbar.TButton").pack(side=tk.LEFT,
                                                                                                       padx=3)
        ttk.Button(btn_frame, text="✅ 验证", command=self.validate_json, style="Toolbar.TButton").pack(side=tk.LEFT,
                                                                                                       padx=3)
        ttk.Button(btn_frame, text="🎨 主题", command=self.show_theme_editor, style="Toolbar.TButton").pack(side=tk.LEFT,
                                                                                                           padx=3)

        ttk.Separator(btn_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        # 导出操作
        ttk.Button(btn_frame, text="📤 导出PPT", command=self.export_ppt, style="Toolbar.TButton").pack(side=tk.LEFT,
                                                                                                       padx=3)
        ttk.Button(btn_frame, text="🖼️ 导出图片", command=self.export_images, style="Toolbar.TButton").pack(
            side=tk.LEFT, padx=3)

        # 右侧快捷操作
        quick_frame = ttk.Frame(toolbar, style="Toolbar.TFrame")
        quick_frame.pack(side=tk.RIGHT, padx=10, pady=8)

        ttk.Button(quick_frame, text="📚 模板", command=self.show_templates, style="Toolbar.TButton").pack(side=tk.LEFT,
                                                                                                          padx=3)
        ttk.Button(quick_frame, text="❓ 帮助", command=self.show_help, style="Toolbar.TButton").pack(side=tk.LEFT,
                                                                                                     padx=3)

    def create_statusbar(self):
        """创建状态栏"""
        statusbar = ttk.Frame(self.root, style="Status.TFrame", height=30)
        statusbar.pack(fill=tk.X, side=tk.BOTTOM)
        statusbar.pack_propagate(False)

        self.status_label = ttk.Label(statusbar, textvariable=self.status_var, style="Status.TLabel")
        self.status_label.pack(side=tk.LEFT, padx=15, pady=5)

        # 编码信息
        encoding_label = ttk.Label(statusbar, text="UTF-8", style="Status.TLabel")
        encoding_label.pack(side=tk.RIGHT, padx=5, pady=5)

        ttk.Separator(statusbar, orient=tk.VERTICAL).pack(side=tk.RIGHT, fill=tk.Y, padx=5)

        # 位置信息
        self.position_var = tk.StringVar(value="行 1, 列 1")
        position_label = ttk.Label(statusbar, textvariable=self.position_var, style="Status.TLabel")
        position_label.pack(side=tk.RIGHT, padx=10, pady=5)

        version_label = ttk.Label(statusbar, text="v2.0.0", style="Status.TLabel")
        version_label.pack(side=tk.RIGHT, padx=15, pady=5)

    def setup_keyboard_shortcuts(self):
        """设置键盘快捷键"""
        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-n>", lambda e: self.new_file())
        self.root.bind("<Control-e>", lambda e: self.export_ppt())
        self.root.bind("<Control-f>", lambda e: self.find_in_editor())
        self.root.bind("<F5>", lambda e: self.refresh_preview())
        self.root.bind("<Control-Left>", lambda e: self.prev_slide())
        self.root.bind("<Control-Right>", lambda e: self.next_slide())

        # 编辑器位置更新
        self.editor.bind("<KeyRelease>", self.update_cursor_position)
        self.editor.bind("<ButtonRelease-1>", self.update_cursor_position)

    def update_cursor_position(self, event=None):
        """更新光标位置显示"""
        try:
            position = self.editor.index(tk.INSERT)
            line, col = position.split('.')
            self.position_var.set(f"行 {line}, 列 {int(col) + 1}")
        except:
            pass

    def update_line_numbers(self, event=None):
        """更新行号显示"""
        try:
            # 获取总行数
            lines = self.editor.get("1.0", "end-1c").split("\n")
            line_numbers_text = "\n".join(str(i) for i in range(1, len(lines) + 1))

            # 更新行号
            self.line_numbers.config(state='normal')
            self.line_numbers.delete('1.0', 'end')
            self.line_numbers.insert('1.0', line_numbers_text)
            self.line_numbers.config(state='disabled')

            # 同步滚动
            self.sync_scroll()
        except:
            pass

    def sync_scroll(self, event=None):
        """同步编辑器和行号的滚动"""
        try:
            self.line_numbers.yview_moveto(self.editor.yview()[0])
        except:
            pass

    def load_sample(self):
        """加载示例JSON"""
        sample_path = os.path.join(os.path.dirname(__file__), "sample.json")
        if os.path.exists(sample_path):
            try:
                with open(sample_path, "r", encoding="utf-8") as fp:
                    sample_text = fp.read()
            except:
                sample_text = SAMPLE_JSON
        else:
            sample_text = SAMPLE_JSON

        self.editor.insert("1.0", sample_text)
        self.editor.edit_modified(False)
        self.update_line_numbers()
        self.schedule_render()

    def on_editor_modified(self, event=None):
        """编辑器内容修改事件"""
        if self.editor.edit_modified():
            self.editor.edit_modified(False)
            self.schedule_render()

    def schedule_render(self):
        """安排渲染（防抖）"""
        if self.render_job:
            self.root.after_cancel(self.render_job)
        self.render_job = self.root.after(DEBOUNCE_MS, self.refresh_preview)

    def refresh_preview(self):
        """刷新预览"""
        if not self.editor:
            return

        content = self.editor.get("1.0", tk.END).strip()
        if not content:
            return

        try:
            meta = json.loads(content)
            validate(meta)
            self.current_meta = meta
            self.preview.set_meta(meta)
            self.update_slide_info()
            self.set_status("✅ JSON 有效", COLOR_SUCCESS)
            self.clear_error_highlight()
        except json.JSONDecodeError as exc:
            self.show_error(f"JSON 语法错误: 第{exc.lineno}行 - {exc.msg}")
            self.highlight_error_line(exc.lineno)
        except ValueError as exc:
            self.show_error(f"验证失败: {exc}")
        except Exception as exc:
            self.show_error(f"错误: {exc}")

    def highlight_error_line(self, line_no: int):
        """高亮错误行"""
        try:
            self.clear_error_highlight()
            self.error_line = line_no
            start = f"{line_no}.0"
            end = f"{line_no}.end"
            self.editor.tag_add("error", start, end)
            self.editor.tag_config("error", background="#ffcccc")
            self.editor.see(start)
        except:
            pass

    def clear_error_highlight(self):
        """清除错误高亮"""
        try:
            self.editor.tag_remove("error", "1.0", tk.END)
        except:
            pass

    def show_error(self, message: str):
        """显示错误信息"""
        self.preview.canvas.delete("all")
        self.preview.canvas.config(width=PREVIEW_MAX_WIDTH, height=PREVIEW_MAX_HEIGHT / 2)

        # 错误图标和消息
        error_text = f"❌ {message}"
        self.preview.canvas.create_text(
            PREVIEW_MAX_WIDTH / 2,
            PREVIEW_MAX_HEIGHT / 4 - 20,
            text=error_text,
            font=("Microsoft YaHei", 14),
            fill=COLOR_ERROR,
            width=PREVIEW_MAX_WIDTH - 40,
            tags="error"
        )

        # 提示信息
        self.preview.canvas.create_text(
            PREVIEW_MAX_WIDTH / 2,
            PREVIEW_MAX_HEIGHT / 4 + 30,
            text="💡 提示：检查JSON格式，确保所有引号、逗号和括号正确",
            font=("Microsoft YaHei", 11),
            fill=COLOR_TEXT_LIGHT,
            width=PREVIEW_MAX_WIDTH - 60,
            tags="hint"
        )

        self.set_status(message, COLOR_ERROR)
        self.current_meta = None
        self.preview.slides = []
        self.slide_info.set("")

    def set_status(self, message: str, color: str = COLOR_INFO):
        """设置状态栏消息"""
        self.status_var.set(message)
        # 可以添加颜色变化效果

    def new_file(self):
        """新建文件"""
        if self.editor.edit_modified():
            result = messagebox.askyesnocancel("新建文件", "当前文件已修改，是否保存？")
            if result is None:
                return
            elif result:
                self.save_file()

        self.editor.delete("1.0", tk.END)
        # 插入基础模板
        template = {
            "version": "1.0",
            "ppt": {
                "size": {"width": 1280, "height": 720, "unit": "px"},
                "defaultUnit": "px",
                "slides": []
            }
        }
        self.editor.insert("1.0", json.dumps(template, ensure_ascii=False, indent=2))
        self.editor.edit_modified(False)
        self.update_line_numbers()
        self.schedule_render()

    def format_json(self):
        """格式化JSON"""
        try:
            content = self.editor.get("1.0", tk.END)
            parsed = json.loads(content)
            formatted = json.dumps(parsed, ensure_ascii=False, indent=2)

            # 保存当前光标位置
            cursor_pos = self.editor.index(tk.INSERT)

            self.editor.delete("1.0", tk.END)
            self.editor.insert("1.0", formatted)

            # 尝试恢复光标位置
            try:
                self.editor.mark_set(tk.INSERT, cursor_pos)
            except:
                pass

            self.update_line_numbers()
            self.set_status("✨ JSON 已格式化", COLOR_SUCCESS)
        except Exception as exc:
            self.set_status(f"格式化失败: {exc}", COLOR_ERROR)

    def validate_json(self):
        """验证JSON"""
        try:
            content = self.editor.get("1.0", tk.END)
            meta = json.loads(content)
            validate(meta)

            # 统计信息
            slides_count = len(meta.get("ppt", {}).get("slides", []))
            elements_count = sum(len(s.get("elements", [])) for s in meta.get("ppt", {}).get("slides", []))

            message = f"JSON 验证通过！\n\n" \
                      f"📊 统计信息：\n" \
                      f"• 幻灯片数量：{slides_count}\n" \
                      f"• 元素总数：{elements_count}\n" \
                      f"• 文件大小：{len(content)} 字符"

            self.set_status("✅ JSON 验证通过", COLOR_SUCCESS)
            messagebox.showinfo("验证成功", message)
        except Exception as exc:
            self.set_status(f"验证失败: {exc}", COLOR_ERROR)
            messagebox.showerror("验证失败", str(exc))

    def open_file(self):
        """打开文件"""
        file_path = filedialog.askopenfilename(
            title="选择JSON文件",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8") as fp:
                data = fp.read()

            # 先验证JSON格式
            json.loads(data)

            self.editor.delete("1.0", tk.END)
            self.editor.insert("1.0", data)
            self.editor.edit_modified(False)
            self.update_line_numbers()
            self.set_status(f"✅ 已加载: {os.path.basename(file_path)}", COLOR_SUCCESS)

            # 保存当前文件路径
            self.current_file_path = file_path
            self.root.title(f"✨ JSON → PPT Designer - {os.path.basename(file_path)}")

        except json.JSONDecodeError as e:
            messagebox.showerror("打开文件失败", f"JSON格式错误：{e}")
        except Exception as exc:
            messagebox.showerror("打开文件失败", str(exc))

    def save_file(self):
        """保存文件"""
        # 如果有当前文件路径，直接保存
        if hasattr(self, 'current_file_path') and self.current_file_path:
            file_path = self.current_file_path
        else:
            file_path = filedialog.asksaveasfilename(
                title="保存JSON文件",
                defaultextension=".json",
                filetypes=[("JSON文件", "*.json")]
            )

        if not file_path:
            return

        try:
            content = self.editor.get("1.0", tk.END)
            # 验证JSON格式
            json.loads(content)

            with open(file_path, "w", encoding="utf-8") as fp:
                fp.write(content.strip())

            self.editor.edit_modified(False)
            self.set_status(f"✅ 已保存: {os.path.basename(file_path)}", COLOR_SUCCESS)
            self.current_file_path = file_path
            self.root.title(f"✨ JSON → PPT Designer - {os.path.basename(file_path)}")

        except json.JSONDecodeError:
            messagebox.showerror("保存失败", "JSON格式错误，请先修复错误后再保存")
        except Exception as exc:
            messagebox.showerror("保存失败", str(exc))

    def export_ppt(self):
        """导出PPT文件"""
        if not self.current_meta:
            messagebox.showerror("导出PPT", "请先修复JSON错误。")
            return

        file_path = filedialog.asksaveasfilename(
            title="导出PPT文件",
            defaultextension=".pptx",
            filetypes=[("PowerPoint文件", "*.pptx")]
        )
        if not file_path:
            return

        try:
            self.set_status("⏳ 正在导出PPT...", COLOR_INFO)
            self.root.update()

            prs, slide_count = build(self.current_meta, logger)
            prs.save(file_path)

            file_size = os.path.getsize(file_path)
            file_size_mb = file_size / (1024 * 1024)

            self.set_status(f"✅ 导出成功: {os.path.basename(file_path)} ({file_size_mb:.2f} MB)", COLOR_SUCCESS)

            result = messagebox.askyesno(
                "导出成功",
                f"PPT已成功导出！\n\n"
                f"📊 文件信息：\n"
                f"• 幻灯片数量：{slide_count} 页\n"
                f"• 文件大小：{file_size_mb:.2f} MB\n"
                f"• 保存位置：{file_path}\n\n"
                f"是否立即打开文件？"
            )

            if result:
                self.open_file_in_system(file_path)

        except Exception as exc:
            logger.exception("导出失败")
            messagebox.showerror("导出失败", f"导出过程中发生错误：\n{str(exc)}")
            self.set_status("❌ 导出失败", COLOR_ERROR)

    def export_images(self):
        """导出为图片"""
        if not self.current_meta:
            messagebox.showerror("导出图片", "请先修复JSON错误。")
            return

        folder_path = filedialog.askdirectory(title="选择导出文件夹")
        if not folder_path:
            return

        try:
            self.set_status("⏳ 正在导出图片...", COLOR_INFO)
            slides = self.current_meta.get("ppt", {}).get("slides", [])

            for i, slide in enumerate(slides):
                # 临时切换到该幻灯片并渲染
                self.preview.current_slide_index = i
                self.preview.render_slide(slide)

                # 获取canvas内容并保存为图片
                ps = self.preview.canvas.postscript(colormode='color')
                img = Image.open(io.BytesIO(ps.encode('utf-8').encode('latin-1')))

                # 保存图片
                img_path = os.path.join(folder_path, f"slide_{i + 1:03d}.png")
                img.save(img_path, "PNG")

            self.set_status(f"✅ 已导出 {len(slides)} 张图片到: {folder_path}", COLOR_SUCCESS)
            messagebox.showinfo("导出成功", f"已成功导出 {len(slides)} 张图片！")

            # 恢复原来的幻灯片
            self.preview.render()

        except Exception as e:
            logger.exception("导出图片失败")
            messagebox.showerror("导出失败", f"导出图片时发生错误：\n{str(e)}")

    def open_file_in_system(self, file_path: str):
        """在系统中打开文件"""
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                import subprocess
                subprocess.call(["open", file_path])
            else:  # Linux
                import subprocess
                subprocess.call(["xdg-open", file_path])
        except Exception as e:
            logger.error(f"无法打开文件: {e}")

    def show_theme_editor(self):
        """显示主题编辑器"""
        theme_window = tk.Toplevel(self.root)
        theme_window.title("🎨 主题编辑器")
        theme_window.geometry("600x400")

        # 简单的主题编辑界面
        ttk.Label(theme_window, text="主题编辑器", font=("Microsoft YaHei", 16, "bold")).pack(pady=10)

        # 颜色设置
        colors_frame = ttk.LabelFrame(theme_window, text="颜色设置", padding=10)
        colors_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 获取当前主题
        current_theme = self.current_meta.get("ppt", {}).get("theme", {}) if self.current_meta else {}
        current_colors = current_theme.get("colors", {})

        color_vars = {}
        for i, (key, default_color) in enumerate([
            ("primary", "#3B82F6"),
            ("secondary", "#10B981"),
            ("accent", "#F59E0B"),
            ("danger", "#EF4444")
        ]):
            row_frame = ttk.Frame(colors_frame)
            row_frame.grid(row=i, column=0, sticky="ew", pady=5)

            ttk.Label(row_frame, text=f"{key}:").pack(side=tk.LEFT, padx=5)

            color_var = tk.StringVar(value=current_colors.get(key, default_color))
            color_vars[key] = color_var

            color_entry = ttk.Entry(row_frame, textvariable=color_var, width=10)
            color_entry.pack(side=tk.LEFT, padx=5)

            color_button = tk.Button(row_frame, text="选择", width=6,
                                     command=lambda k=key, v=color_var: self.choose_color(k, v))
            color_button.pack(side=tk.LEFT, padx=5)

            # 颜色预览
            preview_label = tk.Label(row_frame, width=10, bg=color_var.get())
            preview_label.pack(side=tk.LEFT, padx=5)
            color_var.trace("w", lambda *args, l=preview_label, v=color_var: l.config(bg=v.get()))

        # 字体设置
        fonts_frame = ttk.LabelFrame(theme_window, text="字体设置", padding=10)
        fonts_frame.pack(fill=tk.X, padx=20, pady=10)

        current_fonts = current_theme.get("fonts", {})

        ttk.Label(fonts_frame, text="标题字体:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        heading_font_var = tk.StringVar(value=current_fonts.get("heading", "Microsoft YaHei"))
        ttk.Entry(fonts_frame, textvariable=heading_font_var, width=30).grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(fonts_frame, text="正文字体:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        body_font_var = tk.StringVar(value=current_fonts.get("body", "Arial"))
        ttk.Entry(fonts_frame, textvariable=body_font_var, width=30).grid(row=1, column=1, padx=5, pady=5)

        # 应用按钮
        def apply_theme():
            try:
                # 更新主题到当前JSON
                if self.current_meta:
                    if "theme" not in self.current_meta.get("ppt", {}):
                        self.current_meta["ppt"]["theme"] = {}

                    self.current_meta["ppt"]["theme"]["colors"] = {
                        k: v.get() for k, v in color_vars.items()
                    }
                    self.current_meta["ppt"]["theme"]["fonts"] = {
                        "heading": heading_font_var.get(),
                        "body": body_font_var.get()
                    }

                    # 更新编辑器
                    self.editor.delete("1.0", tk.END)
                    self.editor.insert("1.0", json.dumps(self.current_meta, ensure_ascii=False, indent=2))

                    self.set_status("✅ 主题已应用", COLOR_SUCCESS)
                    theme_window.destroy()
            except Exception as e:
                messagebox.showerror("应用失败", str(e))

        ttk.Button(theme_window, text="应用主题", command=apply_theme).pack(pady=20)

    def choose_color(self, key: str, var: tk.StringVar):
        """选择颜色"""
        from tkinter import colorchooser
        color = colorchooser.askcolor(initialcolor=var.get())
        if color[1]:
            var.set(color[1])

    def show_templates(self):
        """显示模板库"""
        template_window = tk.Toplevel(self.root)
        template_window.title("📚 模板库")
        template_window.geometry("800x600")

        # 模板列表
        templates = [
            {
                "name": "商务演示",
                "description": "专业的商务演示模板，包含封面、目录、内容页和结尾",
                "preview": "🏢",
                "data": self._get_business_template()
            },
            {
                "name": "教育课件",
                "description": "适合教学使用的课件模板，包含标题页、知识点和练习",
                "preview": "📚",
                "data": self._get_education_template()
            },
            {
                "name": "产品介绍",
                "description": "产品展示模板，包含产品特性、对比和价格",
                "preview": "📱",
                "data": self._get_product_template()
            },
            {
                "name": "数据报告",
                "description": "数据分析报告模板，包含各种图表展示",
                "preview": "📊",
                "data": self._get_data_template()
            }
        ]

        # 创建模板网格
        for i, template in enumerate(templates):
            frame = ttk.Frame(template_window, relief=tk.RAISED, borderwidth=1)
            frame.grid(row=i // 2, column=i % 2, padx=10, pady=10, sticky="nsew")

            # 模板预览
            preview_label = tk.Label(frame, text=template["preview"], font=("Arial", 48))
            preview_label.pack(pady=10)

            # 模板名称
            name_label = ttk.Label(frame, text=template["name"], font=("Microsoft YaHei", 14, "bold"))
            name_label.pack()

            # 模板描述
            desc_label = ttk.Label(frame, text=template["description"], wraplength=300)
            desc_label.pack(pady=5)

            # 使用按钮
            use_button = ttk.Button(frame, text="使用此模板",
                                    command=lambda t=template: self.use_template(t, template_window))
            use_button.pack(pady=10)

        # 配置网格权重
        template_window.grid_rowconfigure(0, weight=1)
        template_window.grid_rowconfigure(1, weight=1)
        template_window.grid_columnconfigure(0, weight=1)
        template_window.grid_columnconfigure(1, weight=1)

    def use_template(self, template: dict, window: tk.Toplevel):
        """使用模板"""
        if self.editor.edit_modified():
            result = messagebox.askyesno("使用模板", "当前文件已修改，使用模板将覆盖当前内容。是否继续？")
            if not result:
                return

        self.editor.delete("1.0", tk.END)
        self.editor.insert("1.0", json.dumps(template["data"], ensure_ascii=False, indent=2))
        self.editor.edit_modified(False)
        self.update_line_numbers()
        self.schedule_render()

        window.destroy()
        self.set_status(f"✅ 已加载模板: {template['name']}", COLOR_SUCCESS)

    def _get_business_template(self) -> dict:
        """获取商务模板"""
        return {
            "version": "1.0",
            "ppt": {
                "size": {"width": 1280, "height": 720, "unit": "px"},
                "defaultUnit": "px",
                "theme": {
                    "colors": {
                        "primary": "#1e40af",
                        "secondary": "#64748b",
                        "accent": "#f59e0b"
                    },
                    "fonts": {
                        "heading": "Microsoft YaHei",
                        "body": "Arial"
                    }
                },
                "slides": [
                    {
                        "id": "cover",
                        "background": {
                            "gradient": {
                                "type": "linear",
                                "angle": 135,
                                "stops": [
                                    {"color": "#1e40af", "position": 0},
                                    {"color": "#3730a3", "position": 100}
                                ]
                            }
                        },
                        "elements": [
                            {
                                "type": "text",
                                "text": "商务演示标题",
                                "box": {"x": 640, "y": 300, "w": 800, "h": 100},
                                "style": {"fontSize": 56, "color": "#ffffff", "align": "center", "bold": True}
                            },
                            {
                                "type": "text",
                                "text": "副标题文字",
                                "box": {"x": 640, "y": 400, "w": 600, "h": 60},
                                "style": {"fontSize": 28, "color": "#e0e7ff", "align": "center"}
                            }
                        ]
                    }
                ]
            }
        }

    def _get_education_template(self) -> dict:
        """获取教育模板"""
        return {
            "version": "1.0",
            "ppt": {
                "size": {"width": 1280, "height": 720, "unit": "px"},
                "defaultUnit": "px",
                "theme": {
                    "colors": {
                        "primary": "#059669",
                        "secondary": "#34d399",
                        "accent": "#fbbf24"
                    }
                },
                "slides": [
                    {
                        "id": "title",
                        "background": {"color": "#ecfdf5"},
                        "elements": [
                            {
                                "type": "text",
                                "text": "课程标题",
                                "box": {"x": 640, "y": 300, "w": 800, "h": 100},
                                "style": {"fontSize": 48, "color": "$primary", "align": "center", "bold": True}
                            }
                        ]
                    }
                ]
            }
        }

    def _get_product_template(self) -> dict:
        """获取产品模板"""
        return {
            "version": "1.0",
            "ppt": {
                "size": {"width": 1280, "height": 720, "unit": "px"},
                "defaultUnit": "px",
                "slides": [
                    {
                        "id": "product",
                        "background": {"color": "#ffffff"},
                        "elements": [
                            {
                                "type": "text",
                                "text": "产品名称",
                                "box": {"x": 100, "y": 100, "w": 600, "h": 80},
                                "style": {"fontSize": 42, "color": "#1f2937", "bold": True}
                            }
                        ]
                    }
                ]
            }
        }

    def _get_data_template(self) -> dict:
        """获取数据模板"""
        return {
            "version": "1.0",
            "ppt": {
                "size": {"width": 1280, "height": 720, "unit": "px"},
                "defaultUnit": "px",
                "slides": [
                    {
                        "id": "data",
                        "background": {"color": "#f9fafb"},
                        "elements": [
                            {
                                "type": "text",
                                "text": "数据分析报告",
                                "box": {"x": 640, "y": 50, "w": 800, "h": 80},
                                "style": {"fontSize": 36, "color": "#111827", "align": "center", "bold": True}
                            },
                            {
                                "type": "chart",
                                "chartType": "bar",
                                "box": {"x": 100, "y": 150, "w": 500, "h": 400},
                                "data": {
                                    "categories": ["Q1", "Q2", "Q3", "Q4"],
                                    "series": [
                                        {"name": "销售额", "values": [100, 150, 120, 180]}
                                    ]
                                }
                            }
                        ]
                    }
                ]
            }
        }

    def show_help(self):
        """显示帮助"""
        help_window = tk.Toplevel(self.root)
        help_window.title("❓ 帮助")
        help_window.geometry("700x500")

        # 创建notebook
        notebook = ttk.Notebook(help_window)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 快捷键页面
        shortcuts_frame = ttk.Frame(notebook)
        notebook.add(shortcuts_frame, text="快捷键")

        shortcuts_text = ScrolledText(shortcuts_frame, wrap=tk.WORD, height=20)
        shortcuts_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        shortcuts_text.insert("1.0", """
键盘快捷键：

文件操作：
  Ctrl+N    新建文件
  Ctrl+O    打开文件
  Ctrl+S    保存文件
  Ctrl+E    导出PPT

编辑操作：
  Ctrl+F    查找
  Ctrl+Z    撤销
  Ctrl+Y    重做
  F5        刷新预览

导航操作：
  Ctrl+←    上一页
  Ctrl+→    下一页
  Home      第一页
  End       最后一页

其他：
  Ctrl+H    显示帮助
  Esc       关闭对话框
        """)
        shortcuts_text.config(state='disabled')

        # 元素类型页面
        elements_frame = ttk.Frame(notebook)
        notebook.add(elements_frame, text="元素类型")

        elements_text = ScrolledText(elements_frame, wrap=tk.WORD, height=20)
        elements_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        elements_text.insert("1.0", """
支持的元素类型：

📝 text - 文本元素
   支持样式、段落、列表等

🖼️ image - 图片元素
   支持本地文件、URL、base64

📊 chart - 图表元素
   柱状图、折线图、饼图、雷达图等

🔷 shape - 形状元素
   矩形、圆形、星形、箭头等30+种

📋 table - 表格元素
   支持样式、条纹行列

➖ line - 线条元素
   支持箭头、曲线

⭐ icon - 图标元素
   支持多个图标库

👥 group - 组元素
   组合多个元素

🎬 video - 视频元素
   视频占位符

🎯 smartArt - 智能图形
   流程图、循环图等
        """)
        elements_text.config(state='disabled')

        # 关于页面
        about_frame = ttk.Frame(notebook)
        notebook.add(about_frame, text="关于")

        about_text = ScrolledText(about_frame, wrap=tk.WORD, height=20)
        about_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        about_text.insert("1.0", """
JSON → PPT Designer v2.0.0

一个强大的JSON到PowerPoint转换工具，支持：
• 实时预览
• 丰富的元素类型
• 主题系统
• 渐变和阴影效果
• 多种图表类型
• 智能图形

作者：AI Assistant
许可：MIT License

感谢使用！
        """)
        about_text.config(state='disabled')

    def find_in_editor(self):
        """在编辑器中查找"""
        find_window = tk.Toplevel(self.root)
        find_window.title("查找")
        find_window.geometry("400x100")

        ttk.Label(find_window, text="查找内容:").grid(row=0, column=0, padx=5, pady=5)

        find_var = tk.StringVar()
        find_entry = ttk.Entry(find_window, textvariable=find_var, width=30)
        find_entry.grid(row=0, column=1, padx=5, pady=5)
        find_entry.focus()

        def do_find():
            search_text = find_var.get()
            if not search_text:
                return

            # 清除之前的高亮
            self.editor.tag_remove("found", "1.0", tk.END)

            # 搜索文本
            start = "1.0"
            while True:
                pos = self.editor.search(search_text, start, tk.END)
                if not pos:
                    break
                end = f"{pos}+{len(search_text)}c"
                self.editor.tag_add("found", pos, end)
                start = end

            # 配置高亮样式
            self.editor.tag_config("found", background="#ffff00")

            # 跳转到第一个匹配
            first = self.editor.search(search_text, "1.0", tk.END)
            if first:
                self.editor.see(first)

        ttk.Button(find_window, text="查找", command=do_find).grid(row=0, column=2, padx=5, pady=5)

        # 绑定回车键
        find_entry.bind("<Return>", lambda e: do_find())

    def undo_edit(self):
        """撤销编辑"""
        try:
            self.editor.edit_undo()
        except:
            pass

    def next_slide(self):
        """下一页"""
        self.preview.next_slide()
        self.update_slide_info()

    def prev_slide(self):
        """上一页"""
        self.preview.prev_slide()
        self.update_slide_info()

    def first_slide(self):
        """第一页"""
        if self.preview.slides:
            self.preview.current_slide_index = 0
            self.preview.render()
            self.update_slide_info()

    def last_slide(self):
        """最后一页"""
        if self.preview.slides:
            self.preview.current_slide_index = len(self.preview.slides) - 1
            self.preview.render()
            self.update_slide_info()

    def on_zoom_changed(self, event=None):
        """缩放变化"""
        zoom_value = self.zoom_var.get()
        if zoom_value == "适应":
            # 恢复自适应缩放
            if self.preview.current_meta:
                self.preview.set_meta(self.preview.current_meta)
        else:
            # 设置固定缩放
            try:
                zoom_percent = int(zoom_value.rstrip('%'))
                self.preview.scale = zoom_percent / 100.0
                self.preview.render()
            except:
                pass

    def update_slide_info(self):
        """更新幻灯片信息"""
        total_slides = len(self.preview.slides)
        if total_slides:
            current = self.preview.current_slide_index + 1
            slide_id = self.preview.slides[self.preview.current_slide_index].get("id", "")
            if slide_id:
                self.slide_info.set(f"{current} / {total_slides} [{slide_id}]")
            else:
                self.slide_info.set(f"{current} / {total_slides}")
        else:
            self.slide_info.set("0 / 0")


def main():
    """主函数"""
    root = tk.Tk()

    # 设置DPI感知（Windows）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except:
        pass

    # 设置应用图标
    try:
        if platform.system() == "Windows":
            root.iconbitmap(default='icon.ico')
    except:
        pass

    app = JSONPPTApp(root)

    # 居中窗口
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')

    root.mainloop()


if __name__ == "__main__":
    main()