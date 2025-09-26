import json
import base64
import io
import os
import uuid
import tempfile
from typing import Any, Dict
import http.client
import mimetypes
from codecs import encode

try:
    import requests
except ImportError:
    requests = None

try:
    from pptx import Presentation
    from pptx.util import Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
except ImportError:
    Presentation = None

COZE_UPLOAD_URL = "https://api.coze.cn/v1/files/upload"


def hex_to_rgb(color: str):
    c = color.strip().lstrip('#') if color else ''
    if len(c) == 3:
        c = ''.join(ch * 2 for ch in c)
    if len(c) != 6:
        return (0, 0, 0)
    return (int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))


def unit_to_emu(value: float, total: int, unit: str) -> int:
    if unit == "percent":
        px = total * value / 100.0
    else:
        px = value
    return int(px * 9525)


def resolve_box(box: Dict[str, Any], slide_w: int, slide_h: int, default_unit: str):
    unit = box.get("unit", default_unit)
    x = unit_to_emu(box.get("x", 0), slide_w, unit)
    y = unit_to_emu(box.get("y", 0), slide_h, unit)
    w = unit_to_emu(box.get("w", slide_w), slide_w, unit)
    h = unit_to_emu(box.get("h", slide_h), slide_h, unit)
    return x, y, w, h


def ppt_color(fill_obj, color_hex: str):
    if not color_hex:
        return
    r, g, b = hex_to_rgb(color_hex)
    fill_obj.solid()
    fill_obj.fore_color.rgb = RGBColor(r, g, b)


def map_align(a: str):
    if a == "center":
        return PP_ALIGN.CENTER
    if a == "right":
        return PP_ALIGN.RIGHT
    return PP_ALIGN.LEFT


def map_chart_type(t: str):
    return {
        "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "pie": XL_CHART_TYPE.PIE
    }.get(t, XL_CHART_TYPE.COLUMN_CLUSTERED)


def pick_shape(shape_type: str):
    return {
        "rect": MSO_SHAPE.RECTANGLE,
        "roundRect": MSO_SHAPE.ROUNDED_RECTANGLE,
        "ellipse": MSO_SHAPE.OVAL
    }.get(shape_type, MSO_SHAPE.RECTANGLE)


def get_image_bytes(source: str, logger):
    if not source:
        return None
    if source.startswith("base64:"):
        b64 = source[len("base64:"):]
        try:
            return base64.b64decode(b64)
        except Exception as e:
            logger and logger.warning(f"base64 decode failed: {e}")
            return None
    elif source.startswith("url:"):
        if requests is None:
            return None
        url = source[len("url:"):]
        try:
            resp = requests.get(url, timeout=10)
            if resp.status_code == 200:
                return resp.content
        except Exception as e:
            logger and logger.warning(f"download image failed: {e}")
        return None
    elif source.startswith("file:"):
        path = source[len("file:"):]
        try:
            with open(path, "rb") as f:
                return f.read()
        except Exception as e:
            logger and logger.warning(f"read file image failed: {e}")
            return None
    return None


def add_background(slide, bg_cfg: Dict[str, Any], slide_w_emu: int, slide_h_emu: int, logger):
    if not bg_cfg:
        return
    color = bg_cfg.get("color")
    if color:
        ppt_color(slide.background.fill, color)
    img = bg_cfg.get("image")
    if img and isinstance(img, dict):
        content = get_image_bytes(img.get("src"), logger)
        if content:
            from io import BytesIO
            slide.shapes.add_picture(BytesIO(content), 0, 0, width=slide_w_emu, height=slide_h_emu)


def add_text(slide, elem, slide_w, slide_h, default_unit):
    box = elem.get("box", {})
    x, y, w, h = resolve_box(box, slide_w, slide_h, default_unit)
    shape = slide.shapes.add_textbox(x, y, w, h)
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = elem.get("text", "")
    style = elem.get("style", {})
    p.alignment = map_align(style.get("align", "left"))
    run = p.runs[0]
    if "fontSize" in style:
        run.font.size = Pt(style["fontSize"])
    if style.get("bold"): run.font.bold = True
    if style.get("italic"): run.font.italic = True
    if style.get("color"):
        r, g, b = hex_to_rgb(style["color"])
        run.font.color.rgb = RGBColor(r, g, b)
    if "fontFamily" in style:
        run.font.name = style["fontFamily"]


def add_image(slide, elem, slide_w, slide_h, default_unit, logger):
    box = elem.get("box", {})
    x, y, w, h = resolve_box(box, slide_w, slide_h, default_unit)
    content = get_image_bytes(elem.get("source"), logger)
    if not content:
        return
    from io import BytesIO
    slide.shapes.add_picture(BytesIO(content), x, y, width=w, height=h)


def add_shape(slide, elem, slide_w, slide_h, default_unit):
    box = elem.get("box", {})
    x, y, w, h = resolve_box(box, slide_w, slide_h, default_unit)
    shape = slide.shapes.add_shape(pick_shape(elem.get("shapeType", "rect")), x, y, w, h)
    fill_color = elem.get("fill")
    if fill_color:
        ppt_color(shape.fill, fill_color)


def add_chart(slide, elem, slide_w, slide_h, default_unit):
    box = elem.get("box", {})
    x, y, w, h = resolve_box(box, slide_w, slide_h, default_unit)
    chart_type = map_chart_type(elem.get("chartType", "bar"))
    data_cfg = elem.get("data", {})
    categories = data_cfg.get("categories", [])
    series_cfg = data_cfg.get("series", [])
    cdata = CategoryChartData()
    cdata.categories = categories
    for s in series_cfg:
        cdata.add_series(s.get("name", "Series"), s.get("values", []))
    chart = slide.shapes.add_chart(chart_type, x, y, w, h, cdata).chart
    title = elem.get("title")
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = title


def validate(meta):
    if "ppt" not in meta:
        raise ValueError("missing ppt root")
    if "slides" not in meta["ppt"]:
        raise ValueError("missing ppt.slides")
    if not isinstance(meta["ppt"]["slides"], list):
        raise ValueError("ppt.slides must be list")


def build(meta, logger):
    if Presentation is None:
        raise RuntimeError("python-pptx not installed")
    prs = Presentation()
    ppt_cfg = meta["ppt"]
    size_cfg = ppt_cfg.get("size", {"width": 1280, "height": 720, "unit": "px"})
    sw = size_cfg.get("width", 1280)
    sh = size_cfg.get("height", 720)
    prs.slide_width = unit_to_emu(sw, sw, "px")
    prs.slide_height = unit_to_emu(sh, sh, "px")
    default_unit = ppt_cfg.get("defaultUnit", "px")
    slides_cfg = ppt_cfg["slides"]
    for s in slides_cfg:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_background(slide, s.get("background"), prs.slide_width, prs.slide_height, logger)
        for elem in s.get("elements", []):
            t = elem.get("type")
            try:
                if t == "text":
                    add_text(slide, elem, sw, sh, default_unit)
                elif t == "image":
                    add_image(slide, elem, sw, sh, default_unit, logger)
                elif t == "chart":
                    add_chart(slide, elem, sw, sh, default_unit)
                elif t == "shape":
                    add_shape(slide, elem, sw, sh, default_unit)
            except Exception as e:
                logger and logger.warning(f"element failed id={elem.get('id')} err={e}")
    return prs, len(slides_cfg)


def save_to_temp(prs) -> str:
    tmp_dir = tempfile.gettempdir()
    name = f"ppt_{uuid.uuid4().hex}.pptx"
    path = os.path.abspath(os.path.join(tmp_dir, name))
    prs.save(path)
    return path


def upload_to_coze(file_path: str, token: str, logger, debug: bool=False) -> Dict[str, Any]:
    import requests

    url = "https://api.coze.cn/v1/files/upload"

    payload = {}
    files = [
        ('file',
         (file_path, open(file_path, 'rb'), 'image/png'))
    ]
    headers = {
        'Authorization': 'Bearer cztei_h4a8BkwkIlIIEWCwEeYQt0gD8LClA3jppdKdewRXar2kJGBLbzGK78iv468MYZa7o'
    }

    response = requests.request("POST", url, headers=headers, data=payload, files=files)

    return response.json()['data']


def handler(args):  # Args[Input]
    logger = getattr(args, 'logger', None)
    # 支持 mets 或 meta
    meta_str = getattr(args.input, 'mets', None) or getattr(args.input, 'meta', None)
    token = 'Bearer cztei_loBK9vTkHLpwzulhcf2WaXeMbmsEf3ZFXu0lEnNDxqhwAb5QEVUTcj9z6apwIh5Pe'
    debug = bool(getattr(args.input, 'debug_upload', False))
    extra_upload_path = getattr(args.input, 'upload_file_path', None)  # 可选：额外文件(图片等)上传
    if not meta_str:
        return {"message": "missing mets/meta", "page_count": 0, "ppt_base64": "", "file_path": "", "file_id": "", "upload_error": "no meta json"}
    try:
        meta = json.loads(meta_str)
        validate(meta)
        prs, count = build(meta, logger)
        file_path = save_to_temp(prs)
        bio = io.BytesIO(); prs.save(bio)
        upload_info = upload_to_coze(file_path, token, logger, debug=debug) if token else {"id":"","error":"missing token","http_status":None}
        return {
            "message": "OK" ,
            "page_count": count,
            "file_path": file_path,
            "file_id": upload_info.get('id',''),
            "upload_error": upload_info.get('error',''),
            "upload_http_status": upload_info.get('http_status'),
            "upload_raw": upload_info.get('raw') if debug else None
        }
    except Exception as e:
        if logger:
            logger.error(f"failed: {e}")
        return {"message": f"failed: {e}", "page_count": 0, "ppt_base64": "", "file_path": "", "file_id": "", "upload_error": str(e)}


if __name__ == "__main__":
    # 简单本地测试入口（需安装 python-pptx）
    class DummyInput: pass
    class DummyArgs:
        def __init__(self, meta, token=None):
            self.input = DummyInput()
            self.input.meta = meta
            self.input.coze_token = token
            self.logger = None

    # 配置logger
    import logging
    logging.basicConfig(level=logging.DEBUG)
    logger = logging.getLogger(__name__)


    sample = json.dumps({
    "version": "1.0",
    "ppt": {
        "size": {
            "width": 1280,
            "height": 720,
            "unit": "px"
        },
        "defaultUnit": "px",
        "slides": [
            {
                "id": "cover",
                "title": "封面",
                "background": {
                    "color": "#002B36"
                },
                "elements": [
                    {
                        "type": "text",
                        "id": "title_main",
                        "box": {
                            "x": 80,
                            "y": 180,
                            "w": 1120,
                            "h": 180
                        },
                        "text": "数据分析报告 2025",
                        "style": {
                            "fontSize": 72,
                            "color": "#FFFFFF",
                            "align": "center",
                            "bold": True,
                            "fontFamily": "Arial"
                        }
                    },
                    {
                        "type": "text",
                        "id": "title_sub",
                        "box": {
                            "x": 80,
                            "y": 360,
                            "w": 1120,
                            "h": 80
                        },
                        "text": "自动生成 · 内部评审版",
                        "style": {
                            "fontSize": 36,
                            "color": "#93A1A1",
                            "align": "center"
                        }
                    }
                ]
            },
            {
                "id": "agenda",
                "title": "目录",
                "background": {
                    "color": "#FFFFFF"
                },
                "elements": [
                    {
                        "type": "text",
                        "id": "agenda_title",
                        "box": {
                            "x": 80,
                            "y": 40,
                            "w": 800,
                            "h": 80
                        },
                        "text": "目录",
                        "style": {
                            "fontSize": 54,
                            "color": "#222222",
                            "bold": True
                        }
                    },
                    {
                        "type": "text",
                        "id": "agenda_items",
                        "box": {
                            "x": 100,
                            "y": 140,
                            "w": 1000,
                            "h": 440
                        },
                        "text": "1. 总览\n2. 指标趋势\n3. 区域对比\n4. 产品构成\n5. 后续计划",
                        "style": {
                            "fontSize": 32,
                            "color": "#333333",
                            "fontFamily": "Microsoft YaHei"
                        }
                    }
                ]
            },
            {
                "id": "overview",
                "title": "总览",
                "background": {
                    "color": "#F5F7FA"
                },
                "elements": [
                    {
                        "type": "shape",
                        "id": "overview_panel",
                        "shapeType": "roundRect",
                        "box": {
                            "x": 60,
                            "y": 80,
                            "w": 1160,
                            "h": 520
                        },
                        "fill": "#FFFFFF"
                    },
                    {
                        "type": "text",
                        "id": "kpi_title",
                        "box": {
                            "x": 90,
                            "y": 100,
                            "w": 800,
                            "h": 60
                        },
                        "text": "核心 KPI 概览",
                        "style": {
                            "fontSize": 40,
                            "color": "#222222",
                            "bold": True
                        }
                    },
                    {
                        "type": "text",
                        "id": "kpi_list",
                        "box": {
                            "x": 100,
                            "y": 180,
                            "w": 500,
                            "h": 360
                        },
                        "text": "收入：¥ 12,345,678\n同比：+18.4%\n活跃用户：238,900\n转化率：12.6%\n留存率：43.2%",
                        "style": {
                            "fontSize": 30,
                            "color": "#2A3F54"
                        }
                    },
                    {
                        "type": "image",
                        "id": "logo_img",
                        "box": {
                            "x": 980,
                            "y": 100,
                            "w": 180,
                            "h": 180
                        },
                        "source": "base64:iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIW2P4////FwAJ+wP9iT6n2QAAAABJRU5ErkJggg=="
                    }
                ]
            },
            {
                "id": "trend",
                "title": "指标趋势",
                "background": {
                    "color": "#FFFFFF"
                },
                "elements": [
                    {
                        "type": "text",
                        "id": "trend_title",
                        "box": {
                            "x": 70,
                            "y": 40,
                            "w": 800,
                            "h": 60
                        },
                        "text": "月度收入趋势 (万元)",
                        "style": {
                            "fontSize": 38,
                            "color": "#111111",
                            "bold": True
                        }
                    },
                    {
                        "type": "chart",
                        "id": "trend_chart",
                        "chartType": "line",
                        "box": {
                            "x": 80,
                            "y": 120,
                            "w": 960,
                            "h": 400
                        },
                        "title": "2025 YTD 收入趋势",
                        "data": {
                            "categories": [
                                "1月",
                                "2月",
                                "3月",
                                "4月",
                                "5月",
                                "6月",
                                "7月",
                                "8月"
                            ],
                            "series": [
                                {
                                    "name": "收入",
                                    "values": [
                                        820,
                                        860,
                                        910,
                                        980,
                                        1050,
                                        1110,
                                        1180,
                                        1250
                                    ]
                                },
                                {
                                    "name": "目标",
                                    "values": [
                                        800,
                                        840,
                                        900,
                                        960,
                                        1020,
                                        1080,
                                        1140,
                                        1200
                                    ]
                                }
                            ]
                        }
                    }
                ]
            },
            {
                "id": "compare",
                "title": "区域对比",
                "background": {
                    "color": "#F8FAFC"
                },
                "elements": [
                    {
                        "type": "text",
                        "id": "compare_title",
                        "box": {
                            "x": 80,
                            "y": 50,
                            "w": 900,
                            "h": 60
                        },
                        "text": "区域收入对比 (Q2)",
                        "style": {
                            "fontSize": 36,
                            "color": "#222222",
                            "bold": True
                        }
                    },
                    {
                        "type": "chart",
                        "id": "compare_chart",
                        "chartType": "bar",
                        "box": {
                            "x": 80,
                            "y": 140,
                            "w": 520,
                            "h": 420
                        },
                        "title": "区域柱状图",
                        "data": {
                            "categories": [
                                "华北",
                                "华东",
                                "华南",
                                "西南",
                                "海外"
                            ],
                            "series": [
                                {
                                    "name": "收入",
                                    "values": [
                                        320,
                                        410,
                                        380,
                                        260,
                                        190
                                    ]
                                },
                                {
                                    "name": "去年同期",
                                    "values": [
                                        290,
                                        360,
                                        340,
                                        230,
                                        150
                                    ]
                                }
                            ]
                        }
                    },
                    {
                        "type": "chart",
                        "id": "pie_chart",
                        "chartType": "pie",
                        "box": {
                            "x": 660,
                            "y": 160,
                            "w": 480,
                            "h": 380
                        },
                        "title": "区域占比",
                        "data": {
                            "categories": [
                                "华北",
                                "华东",
                                "华南",
                                "西南",
                                "海外"
                            ],
                            "series": [
                                {
                                    "name": "占比",
                                    "values": [
                                        26,
                                        33,
                                        24,
                                        10,
                                        7
                                    ]
                                }
                            ]
                        }
                    }
                ]
            },
            {
                "id": "composition",
                "title": "产品构成",
                "background": {
                    "image": {
                        "src": "url:https://via.placeholder.com/1280x720/EEF2F7/CCCCCC?text=BG"
                    }
                },
                "elements": [
                    {
                        "type": "text",
                        "id": "comp_title",
                        "box": {
                            "x": 70,
                            "y": 60,
                            "w": 900,
                            "h": 60
                        },
                        "text": "产品线收入构成",
                        "style": {
                            "fontSize": 40,
                            "color": "#1F2937",
                            "bold": True
                        }
                    },
                    {
                        "type": "shape",
                        "id": "comp_panel",
                        "shapeType": "rect",
                        "box": {
                            "x": 60,
                            "y": 140,
                            "w": 1160,
                            "h": 460
                        },
                        "fill": "#FFFFFF"
                    },
                    {
                        "type": "text",
                        "id": "comp_desc",
                        "box": {
                            "x": 90,
                            "y": 160,
                            "w": 1040,
                            "h": 120
                        },
                        "text": "A 系列：45%\nB 系列：28%\nC 系列：15%\n增值及其他：12%",
                        "style": {
                            "fontSize": 30,
                            "color": "#374151"
                        }
                    },
                    {
                        "type": "image",
                        "id": "comp_image",
                        "box": {
                            "x": 800,
                            "y": 300,
                            "w": 360,
                            "h": 240,
                            "unit": "px"
                        },
                        "source": "url:https://via.placeholder.com/360x240.png?text=Product+Mix"
                    }
                ]
            },
            {
                "id": "plan",
                "title": "后续计划",
                "background": {
                    "color": "#1E293B"
                },
                "elements": [
                    {
                        "type": "text",
                        "id": "plan_title",
                        "box": {
                            "x": 80,
                            "y": 60,
                            "w": 1000,
                            "h": 80
                        },
                        "text": "下阶段优化方向",
                        "style": {
                            "fontSize": 46,
                            "color": "#FFFFFF",
                            "bold": True
                        }
                    },
                    {
                        "type": "text",
                        "id": "plan_items",
                        "box": {
                            "x": 100,
                            "y": 170,
                            "w": 1080,
                            "h": 420
                        },
                        "text": "1. 强化渠道精细化投放，提高 ROI\n2. 完善新用户引导路径，提升首周留存\n3. 优化大客户报价体系与折扣策略\n4. 推进数据中台指标对齐与实时化\n5. 构建风险预警 Dashboard",
                        "style": {
                            "fontSize": 32,
                            "color": "#E2E8F0",
                            "fontFamily": "Microsoft YaHei"
                        }
                    }
                ]
            }
        ]
    }
})
    print(handler(DummyArgs(sample)))
