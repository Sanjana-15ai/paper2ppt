# Paper2PPT - makes a PowerPoint from a research paper
# Run with:  uvicorn app:app --reload

import base64
import copy
import io
import json
import os
import re
from collections import Counter
import tempfile
import time
import math
import random
import traceback
import zlib

import fitz  # pymupdf, used to read pdf files
import requests
from PIL import Image, ImageDraw
from docx import Document
from fastapi import Body, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, Response
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.util import Inches, Pt
from typing import Optional

app = FastAPI()
MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-5-5")  # "Pro" model


# ---------------- step 1: read the paper ----------------

def read_paper(filename, data):
    name = filename.lower()
    if name.endswith(".pdf"):
        pdf = fitz.open(stream=data, filetype="pdf")
        text = ""
        for page in pdf:
            text += page.get_text() + "\n"
        return text
    elif name.endswith(".docx"):
        doc = Document(io.BytesIO(data))
        return "\n".join(p.text for p in doc.paragraphs)
    elif name.endswith(".txt") or name.endswith(".md"):
        return data.decode("utf-8", errors="ignore")
    else:
        raise HTTPException(400, "Please upload a PDF, DOCX or TXT file")


# ---------------- step 2: get slide content from the AI ----------------

PROMPT = """Read the material below (a research paper, a document, or just a topic to build from your own knowledge) and make slide content for it.
Tone: {tone}. Write all slide text in {language}.
Reply with ONLY json in this format:
{{"title": "...", "subtitle": "...",
  "palette": {{"primary": "RRGGBB", "secondary": "RRGGBB", "accent": "RRGGBB", "light": "RRGGBB"}},
  "slides": [{{"title": "...", "bullets": ["...", "..."], "notes": "...", "image_query": "...", "layout": "bullets", "takeaway": "one-sentence main message of the slide",
    "items": [], "stats": [], "chart": null, "diagram": null}}], "title_image_query": "...", "references": []}}

Make the deck varied and visual, like a designer made it. Pick a "layout" for each slide:
- "cards": 3-4 parallel ideas. "items": [{{"heading": max 4 words, "text": max 18 words}}]
- "steps": a process or method in 3-5 stages. Same "items" format.
- "stats": ONLY when the paper reports real numbers. "stats": [{{"value": "96.4%", "label": "accuracy"}}] (2-4 stats, exact numbers from the paper)
- "diagram": ONLY for a slide showing a system architecture or workflow: set "layout": "diagram" and "diagram": {{"nodes": ["Input", "Preprocessing", "Model", "Output"]}} (3-6 short stage names).
- "bullets": everything else.
For a research paper, follow this order where the paper has it: Abstract, Introduction, Problem Statement, Methodology, Architecture/Workflow, Results, Discussion, Conclusion, References.
Use at least 3 different layouts, never repeat one more than twice in a row, and use "bullets" (these get a photo) for about half of the slides.
"takeaway" = one sentence, max 16 words, stating that slide's main message.
"references" = up to 6 short entries copied from the paper's OWN reference list; empty list if it has none. Never invent references.
Always ALSO fill "bullets" for every slide (other designs use them). Leave "items" and "stats" empty when not used.

For "chart": use null, EXCEPT on slides where the paper reports real numbers (accuracy, comparisons,
results). Then use {{"type": "bar" or "line" or "pie", "title": "...", "categories": ["A", "B"], "values": [1, 2]}}.
Use the paper's exact numbers only. NEVER invent numbers.

- exactly {n} slides (problem, background, method, results, conclusion, future work)
- 4 to 5 bullets per slide, 10 to 16 words each. Be concrete: include specific facts, numbers, names and reasons from the material. Slides should feel full but readable.
- notes = short speaker script (2-3 sentences)
- image_query = 2-3 concrete words for a stock photo that matches THAT slide's content.
  Every slide must have a different query (good: 'tomato leaf disease', 'smartphone farmer';
  bad: 'research', 'technology', 'science')
- palette colours should match the topic. primary must be dark (white text goes on it),
  light must be almost white. No '#' in the hex codes.

PAPER:
{text}"""


def provider():
    # which AI to use, first one with a key/setting wins. Gemini, Groq and Ollama can be used for free.
    if os.getenv("GEMINI_API_KEY"):
        return "gemini"
    if os.getenv("GROQ_API_KEY"):
        return "groq"
    if os.getenv("OLLAMA_MODEL"):
        return "ollama"
    if os.getenv("ANTHROPIC_API_KEY"):
        return "claude"
    return None


def has_ai():
    return provider() is not None


def openai_style(url, key, name, prompt):
    # Groq and Ollama both speak the same "chat completions" format
    headers = {"content-type": "application/json"}
    if key:
        headers["Authorization"] = "Bearer " + key
    r = requests.post(url, headers=headers, timeout=600, json={
        "model": name, "messages": [{"role": "user", "content": prompt}],
        "response_format": {"type": "json_object"}})
    if r.status_code != 200:
        raise Exception("AI error %s: %s" % (r.status_code, r.text[:300]))
    return r.json()["choices"][0]["message"]["content"]


GEMINI_URL = "https://generativelanguage.googleapis.com/v1beta/models"
_gemini_ok = {}  # remembers which model name worked


def gemini_models(key):
    # asks Google which models this key can use (names change often)
    r = requests.get(GEMINI_URL, params={"key": key, "pageSize": 200}, timeout=30)
    names = []
    for m in r.json().get("models", []):
        n = m["name"].replace("models/", "")
        bad = any(x in n for x in ("image", "tts", "audio", "live", "vision", "embedding"))
        if "generateContent" in m.get("supportedGenerationMethods", []) and "flash" in n and not bad:
            names.append(n)
    return sorted(names, reverse=True)  # newest version first


def gemini_text(prompt, model):
    key = os.getenv("GEMINI_API_KEY")
    first = _gemini_ok.get(model) or os.getenv("GEMINI_MODEL_PRO" if model == "pro" else "GEMINI_MODEL_LITE") or "gemini-3.8-flash"
    tried, queue, listed, last, tries = set(), [first], False, "", {}
    # level 0: ask for JSON + switch "thinking" off (faster). If Gemini refuses a setting (error 400),
    # level 1 drops the thinking setting, level 2 drops the JSON setting too. We read the JSON ourselves anyway.
    level = _gemini_ok.get("level", 0)
    while queue:
        name = queue.pop(0)
        if name in tried:
            continue
        tried.add(name)
        config = {}
        if level < 2:
            config["responseMimeType"] = "application/json"
        if level < 1:
            config["thinkingConfig"] = {"thinkingBudget": 0}
        try:
            r = requests.post("%s/%s:generateContent" % (GEMINI_URL, name), params={"key": key}, timeout=120,
                              json={"contents": [{"parts": [{"text": prompt}]}], "generationConfig": config})
        except requests.exceptions.RequestException as e:
            r = None  # too slow or no connection: try another (smaller) model
            last = "Gemini did not answer in time (%s)" % type(e).__name__
        if r is not None:
            if r.status_code == 200:
                _gemini_ok[model] = name
                _gemini_ok["level"] = level
                parts = r.json()["candidates"][0]["content"]["parts"]
                return "".join(p.get("text", "") for p in parts if not p.get("thought"))
            last = "Gemini error %s: %s" % (r.status_code, r.text[:300])
            if r.status_code == 400 and level < 2:
                level += 1  # a setting was refused: retry with fewer settings
                tried.discard(name)
                queue.insert(0, name)
                continue
            if r.status_code in (429, 500, 502, 503, 504):
                tries[name] = tries.get(name, 0) + 1
                if tries[name] <= 2:  # busy: wait a little and retry
                    time.sleep(3 * tries[name])
                    tried.discard(name)
                    queue.insert(0, name)
                    continue
            elif r.status_code != 404:  # wrong key etc. - no point trying other names
                raise Exception("Gemini error %s: %s" % (r.status_code, r.text[:300]))
        if not listed:  # look up the other models this key can use, smallest/fastest first
            listed = True
            found = gemini_models(key)
            found.sort(key=lambda n: ("lite" not in n) if (model == "lite" or r is None) else ("lite" in n))
            queue += found[:4]
    raise Exception("Gemini is busy or unavailable right now, please try again in a minute. (" + last + ")")


def call_llm(prompt, model="pro"):
    # Plain HTTPS with `requests` (no AI packages needed). model = "lite" or "pro"
    p = provider()
    if p == "gemini":
        return gemini_text(prompt, model)
    if p == "groq":
        name = os.getenv("GROQ_MODEL_PRO", "llama-3.3-70b-versatile") if model == "pro" else os.getenv("GROQ_MODEL_LITE", "llama-3.1-8b-instant")
        return openai_style("https://api.groq.com/openai/v1/chat/completions", os.getenv("GROQ_API_KEY"), name, prompt)
    if p == "ollama":  # runs on your own computer, completely free and offline
        url = os.getenv("OLLAMA_URL", "http://localhost:11434") + "/v1/chat/completions"
        return openai_style(url, None, os.getenv("OLLAMA_MODEL"), prompt)

    name = MODEL if model == "pro" else os.getenv("CLAUDE_MODEL_LITE", "claude-haiku-5-5")
    r = requests.post(
        "https://api.anthropic.com/v1/messages", timeout=180,
        headers={"x-api-key": os.getenv("ANTHROPIC_API_KEY", ""), "anthropic-version": "2023-06-01",
                 "content-type": "application/json"},
        json={"model": name, "max_tokens": 7000, "messages": [{"role": "user", "content": prompt}]})
    if r.status_code != 200:
        try:
            msg = r.json()["error"]["message"]
        except Exception:
            msg = r.text[:300]
        raise Exception("Claude error %s: %s" % (r.status_code, msg))
    return r.json()["content"][0]["text"]


def ask_ai(text, n, tone="professional", language="English", model="pro"):
    limit = 15000 if provider() in ("groq", "ollama") else 60000  # free/local models take shorter input
    answer = call_llm(PROMPT.format(n=n, tone=tone, language=language, text=text[:limit]), model)
    match = re.search(r"\{.*\}", answer, re.S)  # keep only the {...} part
    return json.loads(match.group(0))


def keywords(chunk):
    # picks the 2 most common longer words of a text, used as a photo search
    skip = {"which", "their", "these", "those", "there", "using", "should", "would", "about", "between", "because", "through"}
    words = [w for w in re.findall(r"[a-z]{6,}", chunk.lower()) if w not in skip]
    return " ".join(w for w, _ in Counter(words).most_common(2)) or "research"


def simple_plan(text, n):
    # backup when there is no api key - just splits the paper into parts
    paras = [p.strip() for p in re.split(r"\n\s*\n", text) if len(p.strip()) > 80]
    step = max(1, len(paras) // n)
    slides = []
    for i in range(n):
        chunk = " ".join(paras[i * step:(i + 1) * step])[:600]
        sentences = re.split(r"(?<=[.!?])\s+", chunk)[:4]
        slides.append({
            "title": "Key Point " + str(i + 1),
            "bullets": [s[:110] for s in sentences if s],
            "notes": chunk[:300],
            "image_query": keywords(chunk),
        })
    title = text.strip().split("\n")[0][:90] or "Presentation"
    return {
        "title": title,
        "subtitle": "Auto-generated summary",
        "palette": {"primary": "1F3A5F", "secondary": "3E6FA3", "accent": "F2A541", "light": "F5F7FA"},
        "slides": slides,
    }


# ---------------- step 3: images ----------------

UA = {"User-Agent": "SlideForge/1.0 (student project)"}


def commons_image(query, used):
    # free, no key needed. Wikimedia Commons photos: licences vary, so credit them in your report
    r = requests.get("https://commons.wikimedia.org/w/api.php", timeout=15, headers=UA, params={
        "action": "query", "generator": "search", "gsrsearch": "filetype:bitmap " + query,
        "gsrnamespace": 6, "gsrlimit": 12, "prop": "imageinfo", "iiprop": "url|mime",
        "iiurlwidth": 1200, "format": "json"})
    pages = sorted(r.json().get("query", {}).get("pages", {}).values(), key=lambda p: p.get("index", 0))
    for p in pages:
        info = (p.get("imageinfo") or [{}])[0]
        if info.get("mime") != "image/jpeg" or ("commons", p["pageid"]) in used:
            continue
        img = download_image(info.get("thumburl") or info.get("url"))
        if img:
            used.add(("commons", p["pageid"]))
            return img
    return None


def download_image(url):
    # download and check it is really an image (apis sometimes send an error page)
    r = requests.get(url, timeout=15, headers=UA)
    if r.status_code != 200:
        return None
    try:
        Image.open(io.BytesIO(r.content)).verify()
    except Exception:
        return None
    return io.BytesIO(r.content)


def get_image(query, used):
    # tries pixabay first, then pexels. skips pictures already used on other slides
    if not query:
        return None
    try:
        key = os.getenv("PIXABAY_API_KEY")
        if key:
            r = requests.get("https://pixabay.com/api/", timeout=10, params={
                "key": key, "q": query, "image_type": "photo",
                "orientation": "horizontal", "per_page": 10, "safesearch": "true"})
            for hit in r.json().get("hits", []):
                if ("pixabay", hit["id"]) in used:
                    continue
                img = download_image(hit["largeImageURL"])
                if img:
                    used.add(("pixabay", hit["id"]))
                    return img

        key = os.getenv("PEXELS_API_KEY")
        if key:
            r = requests.get("https://api.pexels.com/v1/search", timeout=10,
                             params={"query": query, "per_page": 10, "orientation": "landscape"},
                             headers={"Authorization": key})
            for photo in r.json().get("photos", []):
                if ("pexels", photo["id"]) in used:
                    continue
                img = download_image(photo["src"]["large"])
                if img:
                    used.add(("pexels", photo["id"]))
                    return img

        img = commons_image(query, used)  # no key (or nothing found): use free Wikimedia photos
        if img:
            return img
    except Exception:
        pass  # no internet or blocked: fall through to the built-in images
    return library_image(query, used)


def valid_diagram(d):
    nodes = d.get("nodes") if isinstance(d, dict) else None
    return isinstance(nodes, list) and 3 <= len(nodes) <= 6 and all(isinstance(x, str) and x.strip() for x in nodes)


def valid_chart(c):
    if not isinstance(c, dict):
        return False
    cats, vals = c.get("categories"), c.get("values")
    if not isinstance(cats, list) or not isinstance(vals, list):
        return False
    if len(cats) < 2 or len(cats) != len(vals):
        return False
    return all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals)


def add_chart(slide, chart, x, y, w, h, color, dark=False):
    data = CategoryChartData()
    data.categories = [str(c) for c in chart["categories"]]
    data.add_series(chart.get("title") or "Result", chart["values"])
    kinds = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "line": XL_CHART_TYPE.LINE_MARKERS,
             "pie": XL_CHART_TYPE.PIE}
    kind = kinds.get(chart.get("type"), XL_CHART_TYPE.COLUMN_CLUSTERED)
    ch = slide.shapes.add_chart(kind, x, y, w, h, data).chart
    ch.font.size = Pt(12)
    if dark:
        ch.font.color.rgb = RGBColor(255, 255, 255)
    ch.has_title = True
    ch.chart_title.text_frame.text = chart.get("title") or ""
    ch.has_legend = (kind == XL_CHART_TYPE.PIE)
    plot = ch.plots[0]
    plot.has_data_labels = True
    series = plot.series[0]
    if kind == XL_CHART_TYPE.COLUMN_CLUSTERED:
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = color
    elif kind == XL_CHART_TYPE.LINE_MARKERS:
        series.format.line.color.rgb = color


def cover_pic(slide, img, x, y, w, h):
    # fills the box with the photo and crops the extra part
    img.seek(0)
    pic = slide.shapes.add_picture(img, x, y)
    iw, ih = pic.image.size
    if iw / ih > w / h:
        pic.crop_left = pic.crop_right = (1 - (w / h) / (iw / ih)) / 2
    else:
        pic.crop_top = pic.crop_bottom = (1 - (iw / ih) / (w / h)) / 2
    pic.left, pic.top, pic.width, pic.height = x, y, w, h


def put_visual(slide, img, chart, x, y, w, color, dark=False, h=None):
    # a chart (real numbers from the paper) wins over a stock photo
    if chart:
        add_chart(slide, chart, x, y, w, int(w * 0.8), color, dark)
    elif img:
        cover_pic(slide, img, x, y, w, h or int(w * 0.7))


# ---------------- step 4: make the pptx ----------------

def hex_to_rgb(h):
    return RGBColor.from_string(h.replace("#", "").upper())


def remove_old_slides(prs):
    # templates often come with sample slides, delete them first
    ids = prs.slides._sldIdLst
    for s in list(ids):
        prs.part.drop_rel(s.rId)
        ids.remove(s)


def add_bullets(text_frame, bullets, size=None, color=None):
    for i, b in enumerate(bullets):
        para = text_frame.paragraphs[0] if i == 0 else text_frame.add_paragraph()
        para.text = b
        para.space_after = Pt(8)
        if size:
            para.font.size = Pt(size)
        if color:
            para.font.color.rgb = color


def make_with_layouts(plan, template_bytes, images):
    prs = Presentation(io.BytesIO(template_bytes))
    remove_old_slides(prs)

    title_layout = prs.slide_layouts[0]
    content_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else title_layout
    for lay in prs.slide_layouts:  # look for the standard "Title Slide" and "Title and Content" layouts by name
        idxs = [ph.placeholder_format.idx for ph in lay.placeholders]
        if "title slide" in lay.name.lower():
            title_layout = lay
        if "title and content" in lay.name.lower() and 0 in idxs and 1 in idxs:
            content_layout = lay

    # title slide
    slide = prs.slides.add_slide(title_layout)
    if slide.shapes.title is not None:
        slide.shapes.title.text = plan["title"]
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 1:
            ph.text = plan.get("subtitle", "")

    # content slides
    for data, img in zip(plan["slides"], images):
        slide = prs.slides.add_slide(content_layout)
        if slide.shapes.title is not None:
            slide.shapes.title.text = data["title"]

        body = None
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == 1:
                body = ph
        if body is not None:
            chars = sum(len(b) for b in data["bullets"])
            if img or valid_chart(data.get("chart")):
                chars = int(chars * 1.8)  # narrower text column next to the picture
            add_bullets(body.text_frame, data["bullets"],
                        24 if chars < 300 else 22 if chars < 420 else 20 if chars < 540 else 18 if chars < 680 else 16)
            l, t, w_ = body.left, body.top, body.width  # set position AND size together (setting only one breaks placeholders)
            body.left, body.top, body.width = l, t, w_
            body.height = max(Inches(2), prs.slide_height - Inches(1.5) - t)  # leave room for the takeaway bar
            chart = data.get("chart") if valid_chart(data.get("chart")) else None
            if img or chart:
                # make the text box narrower and put the picture on the right
                left, top, height = body.left, body.top, body.height
                body.left, body.top = left, top
                body.width, body.height = int(prs.slide_width * 0.55), height
                put_visual(slide, img, chart, int(prs.slide_width * 0.62), body.top,
                           int(prs.slide_width * 0.33), RGBColor(0x2E, 0x5C, 0x8A))
        if body is not None and data.get("takeaway"):
            bar = slide.shapes.add_shape(1, body.left, prs.slide_height - Inches(1.3), prs.slide_width - 2 * body.left, Inches(0.7))
            bar.fill.solid()
            bar.fill.fore_color.theme_color = MSO_THEME_COLOR.ACCENT_1  # the template's own accent colour
            bar.line.fill.background()
            bar.text_frame.word_wrap = True
            bar.text_frame.text = "Key takeaway: " + data["takeaway"]
            para = bar.text_frame.paragraphs[0]
            para.font.size, para.font.bold, para.font.color.rgb = Pt(16), True, RGBColor(255, 255, 255)
        slide.notes_slide.notes_text_frame.text = data.get("notes", "")
    prs.sf_mode = "layouts"
    return prs


# Some templates draw their design (bars, logo) on the sample slides instead of the
# layouts. For those we copy the sample slide's design and reuse its fonts/colours.

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def is_decor(shape):
    # pictures and shapes without text (bars, logos) make up the template look
    try:
        if shape.is_placeholder:
            return False
        if shape.shape_type == 13:  # picture
            return True
        return shape.shape_type == 1 and not shape.text_frame.text.strip()
    except Exception:
        return False


def text_boxes(slide):
    boxes = []
    for sh in slide.shapes:
        if sh.has_text_frame and not sh.is_placeholder and sh.text_frame.text.strip():
            boxes.append(sh)
    return boxes


def get_style(shape):
    style = {"name": None, "bold": None, "color": None}
    if shape is None:
        return style
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.text.strip():
                style["name"] = run.font.name
                style["bold"] = run.font.bold
                try:
                    style["color"] = run.font.color.rgb
                except Exception:
                    pass
                return style
    return style


def copy_slide_design(prs, src, keep_header=False):
    new = prs.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):  # remove the empty default placeholders
        sh._element.getparent().remove(sh._element)
    for sh in src.shapes:
        keep = is_decor(sh)
        if keep_header and sh in text_boxes(src):
            one_line = "\n" not in sh.text_frame.text.strip() and len(sh.text_frame.text) < 80
            if one_line and sh.top is not None and sh.top < prs.slide_height * 0.35:
                keep = True
        if not keep:
            continue
        el = copy.deepcopy(sh._element)
        for blip in el.iter("{%s}blip" % A_NS):
            rid = blip.get("{%s}embed" % R_NS)
            if rid:
                part = src.part.related_part(rid)
                blip.set("{%s}embed" % R_NS, new.part.relate_to(part, RT.IMAGE))
        new.shapes._spTree.append(el)
    return new


def write_text(slide, x, y, w, h, paragraphs, size, style, bold=False, align=None, default_color=None, space=8):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, t in enumerate(paragraphs):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.text = t
        para.space_after = Pt(space)
        if align is not None:
            para.alignment = align
        para.font.size = Pt(size)
        para.font.bold = bold
        if style["name"]:
            para.font.name = style["name"]
        color = style["color"] or default_color
        if color is not None:
            para.font.color.rgb = color
    return tb


def make_with_template(plan, template_bytes, images):
    prs = Presentation(io.BytesIO(template_bytes))
    samples = list(prs.slides)
    if not samples:
        return make_with_layouts(plan, template_bytes, images)
    title_model = samples[0]
    content_model = samples[-1]
    if not any(is_decor(sh) for sh in content_model.shapes):
        return make_with_layouts(plan, template_bytes, images)

    # work out how the template writes headings and normal text
    boxes = text_boxes(content_model)
    heading = min(boxes, key=lambda b: len(b.text_frame.text)) if boxes else None
    body = max(boxes, key=lambda b: len(b.text_frame.text)) if boxes else None
    head_style, body_style = get_style(heading), get_style(body)
    W, H = prs.slide_width, prs.slide_height
    left = Inches(0.8)
    head_top = heading.top if heading is not None and heading.top else Inches(1.4)

    # title slide
    slide = copy_slide_design(prs, title_model, keep_header=True)
    write_text(slide, Inches(1), int(H * 0.36), W - Inches(2), Inches(2), [plan["title"]], 30,
               head_style, True, PP_ALIGN.CENTER, RGBColor(0x1F, 0x1F, 0x1F))
    write_text(slide, Inches(1), int(H * 0.62), W - Inches(2), Inches(1), [plan.get("subtitle", "")], 18,
               body_style, False, PP_ALIGN.CENTER, RGBColor(0x22, 0x22, 0x22))

    # content slides
    for data, img in zip(plan["slides"], images):
        slide = copy_slide_design(prs, content_model)
        size = 28 if len(data["title"]) < 45 else 22
        write_text(slide, left, head_top, W - 2 * left, Inches(0.9), [data["title"]], size,
                   head_style, True, PP_ALIGN.CENTER, RGBColor(0x1F, 0x1F, 0x1F))
        body_top = head_top + Inches(1.0)
        chart = data.get("chart") if valid_chart(data.get("chart")) else None
        body_w = int(W * 0.58) - left if (img or chart) else W - 2 * left
        chars = sum(len(b) for b in data["bullets"])
        if img or chart:
            chars = int(chars * 1.4)  # narrower text column
        size = 26 if chars < 300 else 24 if chars < 420 else 22 if chars < 520 else 20 if chars < 640 else 18  # fill the slide, shrink if long
        n_b = len(data["bullets"])
        if n_b > 5:  # long lists (agenda, references) need smaller text
            size = min(size, 20 if n_b <= 6 else 18)
        write_text(slide, left, body_top, body_w, H - body_top - Inches(1.5),
                   ["\u2022  " + b for b in data["bullets"]], size, body_style, False, None,
                   RGBColor(0x22, 0x22, 0x22), 14 if n_b <= 5 else 6)
        put_visual(slide, img, chart, int(W * 0.62), body_top, int(W * 0.33),
                   head_style["color"] or RGBColor(0x2E, 0x5C, 0x8A), False, Inches(3.6))
        if data.get("takeaway"):  # key message bar at the bottom, in the template's heading colour
            bar = slide.shapes.add_shape(1, left, H - Inches(1.3), W - 2 * left, Inches(0.7))
            bar.fill.solid()
            bar.fill.fore_color.rgb = head_style["color"] or RGBColor(0x1F, 0x3A, 0x5F)
            bar.line.fill.background()
            write_text(slide, left + Inches(0.2), H - Inches(1.2), W - 2 * left - Inches(0.4), Inches(0.5),
                       ["Key takeaway: " + data["takeaway"]], 16, dict(body_style, color=None), True, None,
                       RGBColor(255, 255, 255))
        slide.notes_slide.notes_text_frame.text = data.get("notes", "")

    # remove the original sample slides (they are the first ones in the list)
    ids = prs.slides._sldIdLst
    for sld in list(ids)[:len(samples)]:
        prs.part.drop_rel(sld.rId)
        ids.remove(sld)
    prs.sf_mode = "copied"
    return prs


_T = [  # id, primary, secondary, accent, light, style, font
    ("ocean",    "0B3C5D", "328CC1", "D9B310", "F4F8FB", "bar",    "Calibri"),
    ("sunset",   "7A2E3A", "E4572E", "F3A712", "FFF8F0", "side",   "Georgia"),
    ("midnight", "14213D", "2A4A7F", "FCA311", "E5E5E5", "dark",   "Arial"),
    ("forest",   "1B4332", "2D6A4F", "F4A261", "F1F8F4", "band",   "Georgia"),
    ("royal",    "2B2D6E", "5B5FC7", "FFB703", "F5F5FF", "corner", "Calibri"),
    ("coral",    "8C2F39", "FF6B6B", "FFD166", "FFF5F5", "corner", "Trebuchet MS"),
    ("lavender", "4A3F8F", "8E7CC3", "F6BD60", "F7F3FF", "side",   "Calibri"),
    ("slate",    "263238", "546E7A", "26C6DA", "F4F7F8", "band",   "Arial"),
    ("mint",     "0F4C5C", "2A9D8F", "E9C46A", "F1FAF7", "bar",    "Calibri"),
    ("crimson",  "6A040F", "9D0208", "FFBA08", "FFF6F0", "dark",   "Georgia"),
    ("noir",     "111111", "2B2B2B", "D4AF37", "F5F5F5", "dark",   "Georgia"),
    ("sky",      "0077B6", "00B4D8", "FFB703", "F0F9FF", "band",   "Calibri"),
    ("rose",     "7B2D5B", "D6659B", "FFD6A5", "FFF5FA", "corner", "Georgia"),
    ("sand",     "5C4033", "A47551", "E0A458", "FBF6EE", "side",   "Cambria"),
    ("teal",     "004E64", "00A5CF", "FFB30F", "EAF6F9", "bar",    "Arial"),
    ("plum",     "3C1642", "6A3D7C", "F7B801", "F8F1FB", "dark",   "Trebuchet MS"),
    ("meadow",   "386641", "6A994E", "F2CC8F", "F4F9EC", "corner", "Calibri"),
    ("steel",    "1F2A44", "4F6D9A", "F2A65A", "F3F6FB", "band",   "Arial"),
    ("wine",     "590D22", "C9184A", "FFB3C1", "FFF0F3", "side",   "Georgia"),
    ("aqua",     "0B132B", "1C2541", "5BC0BE", "E8F1F2", "dark",   "Arial"),
]
THEMES = {t[0]: dict(zip(("primary", "secondary", "accent", "light", "style", "font"), t[1:])) for t in _T}


def make_from_scratch(plan, images, theme="auto"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    W, H = prs.slide_width, prs.slide_height

    t = THEMES.get(theme)
    if t is None:  # "auto" = colours chosen by the AI for this paper
        p = plan["palette"]
        t = {"primary": p["primary"], "secondary": p["secondary"], "accent": p["accent"],
             "light": p["light"], "style": "bar", "font": "Calibri"}
    dark, mid = hex_to_rgb(t["primary"]), hex_to_rgb(t["secondary"])
    accent, light = hex_to_rgb(t["accent"]), hex_to_rgb(t["light"])
    style, font = t["style"], t["font"]
    white, ink = RGBColor(255, 255, 255), RGBColor(0x22, 0x22, 0x22)
    is_dark = style == "dark"
    m = Inches(0.9) if style == "side" else Inches(0.7)   # page margin
    card_fill = mid if is_dark else white
    card_text = white if is_dark else ink
    head_color = accent if is_dark else dark
    big_color = accent if is_dark else mid

    def box(slide, x, y, w, h, color, kind=1, round_=False):
        shape = slide.shapes.add_shape(5 if round_ else kind, x, y, w, h)  # 1 rect, 5 rounded, 9 circle
        if round_:
            shape.adjustments[0] = 0.06
        shape.fill.solid()
        shape.fill.fore_color.rgb = color
        shape.line.fill.background()
        return shape

    def label(slide, x, y, w, h, text, size, color, bold=False, align=None):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.text = text
        para = tf.paragraphs[0]
        para.font.size, para.font.bold, para.font.color.rgb = Pt(size), bold, color
        para.font.name = font
        if align is not None:
            para.alignment = align
        return tf

    def bullet_list(slide, x, y, w, h, bullets, size, color):
        tf = label(slide, x, y, w, h, "", size, color)
        add_bullets(tf, ["\u2022  " + b for b in bullets], size, color)
        for para in tf.paragraphs:
            para.font.name = font

    def cover(slide, img, x, y, w, h):
        # fills the box with the picture and crops the extra part (like a photo background)
        img.seek(0)
        pic = slide.shapes.add_picture(img, x, y)
        iw, ih = pic.image.size
        if iw / ih > w / h:
            pic.crop_left = pic.crop_right = (1 - (w / h) / (iw / ih)) / 2
        else:
            pic.crop_top = pic.crop_bottom = (1 - (iw / ih) / (w / h)) / 2
        pic.left, pic.top, pic.width, pic.height = x, y, w, h

    def new_slide(title, x=None, width=None):
        slide = prs.slides.add_slide(blank)
        x = m if x is None else x
        width = W - 2 * m if width is None else width
        box(slide, 0, 0, W, H, dark if is_dark else light)
        tcolor, ty = head_color, Inches(0.4)
        if style == "side":
            box(slide, 0, 0, Inches(0.35), H, dark)
        if style == "band":  # coloured header band across the top
            box(slide, 0, 0, W, Inches(1.35), dark)
            box(slide, 0, Inches(1.35), W, Inches(0.06), accent)
            tcolor, ty = white, Inches(0.3)
        if style == "corner":  # decorative circles in the top-right corner
            box(slide, W - Inches(2.0), -Inches(1.5), Inches(3.5), Inches(3.5), accent, 9)
            box(slide, W - Inches(2.6), -Inches(1.9), Inches(2.4), Inches(2.4), mid, 9)
            width = min(width, W - 2 * m - Inches(2.2))
        label(slide, x, ty, width, Inches(1), title, 32 if len(title) < 42 else 26, tcolor, True)
        if style != "band":
            box(slide, x + Inches(0.08), Inches(1.4), Inches(1.2), Inches(0.07), accent)
        return slide

    # ---- title slide: text panel + photo ----
    slide = prs.slides.add_slide(blank)
    box(slide, 0, 0, W, H, dark)
    tw = int(W * 0.58) - Inches(1.3)
    title_img = plan.get("_title_img")
    if title_img:
        cover(slide, title_img, int(W * 0.58), 0, W - int(W * 0.58), H)
    else:
        box(slide, W - Inches(5), Inches(3.2), Inches(8), Inches(8), mid, 9)
        tw = Inches(9.5)
    box(slide, Inches(0.8), Inches(1.5), Inches(1.2), Inches(0.08), accent)
    label(slide, Inches(0.8), Inches(1.8), tw, Inches(3), plan["title"], 36 if len(plan["title"]) < 70 else 30, white, True)
    label(slide, Inches(0.8), Inches(5.0), tw, Inches(1.2), plan.get("subtitle", ""), 18, light)

    # ---- layouts ----
    def cards(slide, items, num=True):
        n = len(items)
        gap = Inches(0.3)
        cw = int((W - 2 * m - gap * (n - 1)) / n)
        for i, it in enumerate(items):
            x = m + i * (cw + gap)
            box(slide, x, Inches(1.9), cw, Inches(4.3), card_fill, 1, True)
            box(slide, x + Inches(0.3), Inches(2.15), Inches(0.6), Inches(0.07), accent)
            label(slide, x + Inches(0.2), Inches(2.3), cw - Inches(0.4), Inches(0.8), "%02d" % (i + 1), 30, big_color, True)
            label(slide, x + Inches(0.2), Inches(3.1), cw - Inches(0.4), Inches(1), it["heading"], 20 if n < 4 else 18, head_color, True)
            label(slide, x + Inches(0.2), Inches(4.0), cw - Inches(0.4), Inches(2.1), it.get("text", ""), 18 if n < 4 else 16, card_text)

    def steps(slide, items):
        n = len(items)
        cw = int((W - 2 * m) / n)
        box(slide, m + cw // 2, Inches(2.75), cw * (n - 1), Inches(0.05), accent)  # line through the circles
        for i, it in enumerate(items):
            cx = m + i * cw
            circle = box(slide, cx + cw // 2 - Inches(0.45), Inches(2.3), Inches(0.9), Inches(0.9), mid if not is_dark else accent, 9)
            label(slide, cx + cw // 2 - Inches(0.45), Inches(2.47), Inches(0.9), Inches(0.6), str(i + 1), 24,
                  white if not is_dark else dark, True, PP_ALIGN.CENTER)
            label(slide, cx + Inches(0.1), Inches(3.5), cw - Inches(0.2), Inches(0.9), it["heading"], 19, head_color, True, PP_ALIGN.CENTER)
            label(slide, cx + Inches(0.1), Inches(4.4), cw - Inches(0.2), Inches(2.4), it.get("text", ""), 15, card_text if is_dark else ink, False, PP_ALIGN.CENTER)

    def stats(slide, items, bullets):
        n = len(items)
        gap = Inches(0.3)
        cw = int((W - 2 * m - gap * (n - 1)) / n)
        for i, it in enumerate(items):
            x = m + i * (cw + gap)
            box(slide, x, Inches(1.9), cw, Inches(2.6), card_fill, 1, True)
            label(slide, x, Inches(2.2), cw, Inches(1.2), str(it["value"]), 48 if n < 4 else 40, big_color, True, PP_ALIGN.CENTER)
            label(slide, x + Inches(0.2), Inches(3.5), cw - Inches(0.4), Inches(0.9), it.get("label", ""), 16, card_text, False, PP_ALIGN.CENTER)
        if bullets:
            bullet_list(slide, m, Inches(4.8), W - 2 * m, Inches(1.5), bullets[:3], 18, white if is_dark else ink)

    def chart_slide(slide, chart, bullets):
        cw = int(W * 0.55)
        box(slide, m, Inches(1.9), cw, Inches(4.3), card_fill, 1, True)
        add_chart(slide, chart, m + Inches(0.2), Inches(2.0), cw - Inches(0.4), Inches(4.1), mid if not is_dark else accent, is_dark)
        x2 = m + cw + Inches(0.4)
        bullet_list(slide, x2, Inches(2.0), W - x2 - m, Inches(4.1), bullets[:4], 16, white if is_dark else ink)

    def takeaway(slide, text, x=None, w=None):
        if not text:
            return
        x = m if x is None else x
        w = W - 2 * m if w is None else w
        box(slide, x, Inches(6.35), w, Inches(0.75), accent if is_dark else mid, 1, True)
        label(slide, x + Inches(0.25), Inches(6.5), w - Inches(0.5), Inches(0.5), "Key takeaway: " + text, 16,
              dark if is_dark else white, True)

    def flow(slide, nodes, bullets):
        # architecture / workflow drawn as real, editable PowerPoint shapes
        n = len(nodes)
        gap = Inches(0.6)
        bw = int((W - 2 * m - gap * (n - 1)) / n)
        fill = accent if is_dark else mid
        for i, name in enumerate(nodes):
            x = m + i * (bw + gap)
            box(slide, x, Inches(2.2), bw, Inches(1.6), fill, 1, True)
            label(slide, x + Inches(0.05), Inches(2.65), bw - Inches(0.1), Inches(0.9), name,
                  18 if n <= 4 else 15, dark if is_dark else white, True, PP_ALIGN.CENTER)
            if i < n - 1:
                box(slide, x + bw + Inches(0.1), Inches(2.75), gap - Inches(0.2), Inches(0.5), accent, 33)  # arrow
        if bullets:
            bullet_list(slide, m, Inches(4.3), W - 2 * m, Inches(2.0), bullets[:3], 20, white if is_dark else ink)

    for i, (data, img) in enumerate(zip(plan["slides"], images)):
        bullets = data.get("bullets") or []
        chart = data.get("chart") if valid_chart(data.get("chart")) else None
        items = [x for x in (data.get("items") or []) if isinstance(x, dict) and x.get("heading")][:5]
        nums = [x for x in (data.get("stats") or []) if isinstance(x, dict) and x.get("value")][:4]
        layout = data.get("layout")
        text_c = white if is_dark else ink
        tk_x = tk_w = None

        if chart:
            slide = new_slide(data["title"])
            chart_slide(slide, chart, bullets)
        elif valid_diagram(data.get("diagram")):
            slide = new_slide(data["title"])
            flow(slide, [str(x) for x in data["diagram"]["nodes"]], bullets)
        elif layout == "stats" and len(nums) >= 2:
            slide = new_slide(data["title"])
            stats(slide, nums, bullets)
        elif layout == "cards" and len(items) >= 3:
            slide = new_slide(data["title"])
            cards(slide, items[:4])
        elif layout == "steps" and len(items) >= 3:
            slide = new_slide(data["title"])
            steps(slide, items)
        elif img:  # half-page photo, switching sides to keep it interesting
            on_right = i % 2 == 0
            pw = int(W * 0.42)
            tx = m if on_right else pw + Inches(0.5)
            slide = new_slide(data["title"], tx, W - pw - m - Inches(0.5))
            cover(slide, img, W - pw if on_right else 0, 0, pw, H)
            tk_x, tk_w = tx, W - pw - m - Inches(0.5)
            bullet_list(slide, tx, Inches(1.9), W - pw - m - Inches(0.5), Inches(4.3), bullets, 20, text_c)
        else:
            slide = new_slide(data["title"])
            box(slide, m, Inches(1.9), W - 2 * m, Inches(4.3), card_fill, 1, True)
            bullet_list(slide, m + Inches(0.4), Inches(2.2), W - 2 * m - Inches(0.8), Inches(3.9), bullets, 20, card_text)
        takeaway(slide, data.get("takeaway"), tk_x, tk_w)
        slide.notes_slide.notes_text_frame.text = data.get("notes", "")
    return prs


# ---------------- web routes ----------------

@app.get("/", response_class=HTMLResponse)
def home():
    path = os.path.join(os.path.dirname(__file__), "index.html")
    with open(path, encoding="utf-8") as f:
        return f.read()


@app.post("/generate")
async def generate(paper: UploadFile = File(...),
                   template: Optional[UploadFile] = File(None),
                   slides: Optional[int] = Form(8),
                   use_images: bool = Form(True),
                   theme: str = Form("auto")):

    text = read_paper(paper.filename, await paper.read())
    if len(text.strip()) < 200:
        raise HTTPException(400, "Could not read enough text from the file (scanned pdf?)")
    slides = max(3, min(slides or 8, 20))

    # get slide content
    try:
        if has_ai():
            plan = ask_ai(text, slides)
        else:
            plan = simple_plan(text, slides)
    except Exception as e:
        raise HTTPException(500, "AI step failed: " + str(e))

    # get images
    used = set()
    if use_images:
        plan["_title_img"] = get_image(plan.get("title_image_query") or keywords(plan["title"]), used)
    images = []
    for s in plan["slides"]:
        images.append(get_image(s.get("image_query"), used) if use_images and not valid_chart(s.get("chart")) else None)

    # build the ppt
    template_bytes = None
    if template is not None and template.filename:
        template_bytes = await template.read()
    try:
        if template_bytes:
            prs = make_with_template(plan, template_bytes, images)
        else:
            prs = make_from_scratch(plan, images, theme)
        out = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        prs.save(out.name)
    except Exception as e:
        traceback.print_exc()  # full details show in the terminal
        raise HTTPException(500, "Building the PPT failed: " + repr(e))
    mode = "ai" if has_ai() else "offline"
    return FileResponse(out.name, filename="presentation.pptx", headers={"X-Mode": mode},
                        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")


# ---------------- SlideForge API: outline -> edit -> build ----------------

def pick_image(sd, query, skip, use_images, used):
    if sd.get("photo") == "none":
        return None
    data = sd.get("user_image")
    if data and "," in data:  # a picture the user chose from the Media page
        try:
            raw = base64.b64decode(data.split(",", 1)[1])
            Image.open(io.BytesIO(raw)).verify()
            return io.BytesIO(raw)
        except Exception:
            pass
    return get_image(query, used) if use_images and not skip else None


def build_deck(plan, template_bytes, theme, use_images):
    plan.setdefault("palette", {"primary": "1F3A5F", "secondary": "3E6FA3", "accent": "F2A541", "light": "F5F7FA"})
    plan.setdefault("title", "Presentation")
    if not isinstance(plan.get("slides"), list) or not plan["slides"]:
        raise HTTPException(400, "The outline has no slides")
    titles = [str(x.get("title", "")) for x in plan["slides"]]
    if len(titles) >= 5 and not any(t.lower() in ("agenda", "contents", "outline", "table of contents") for t in titles):
        items = ["%d.  %s" % (i + 1, t) for i, t in enumerate(titles[:8])]  # agenda built from the slide titles
        plan["slides"].insert(0, {"title": "Agenda", "bullets": items, "notes": "Overview of the talk.", "photo": "none"})
    refs = [str(r)[:150] for r in (plan.get("references") or []) if str(r).strip()][:6]
    if refs and not any(t.lower().startswith("reference") for t in titles):  # only references that are in the paper
        plan["slides"].append({"title": "References", "bullets": refs, "notes": "", "photo": "none"})
    used = set()
    plan["_title_img"] = get_image(plan.get("title_image_query") or keywords(plan["title"]), used) if use_images else None
    images = []
    for sd in plan["slides"]:
        sd["bullets"] = [str(b) for b in (sd.get("bullets") or [])][:5]
        skip = valid_chart(sd.get("chart")) or valid_diagram(sd.get("diagram"))
        query = sd.get("image_query") or keywords(sd.get("title", ""))
        images.append(pick_image(sd, query, skip, use_images, used))
    if template_bytes:
        prs = make_with_template(plan, template_bytes, images)
    else:
        prs = make_from_scratch(plan, images, theme)
    out = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(out.name)
    return out.name, sum(1 for i in images if i), getattr(prs, "sf_mode", "none")


@app.post("/api/outline")
async def api_outline(paper: Optional[UploadFile] = File(None), topic: str = Form(""),
                      slides: int = Form(8), tone: str = Form("professional"),
                      language: str = Form("English"), model: str = Form("pro")):
    text = ""
    if paper is not None and paper.filename:
        text = read_paper(paper.filename, await paper.read())
    if topic.strip():
        text = "TOPIC / INSTRUCTIONS: " + topic.strip() + "\n\n" + text
    if len(text.strip()) < 3:
        raise HTTPException(400, "Type a topic or upload a paper first")
    slides = max(3, min(slides or 8, 20))
    try:
        return ask_ai(text, slides, tone, language, model) if has_ai() else simple_plan(text, slides)
    except Exception as e:
        raise HTTPException(500, "AI step failed: " + str(e))


@app.post("/api/build")
async def api_build(plan: str = Form(...), template: Optional[UploadFile] = File(None),
                    theme: str = Form("auto"), use_images: bool = Form(True)):
    tpl = await template.read() if template is not None and template.filename else None
    try:
        path, n_img, tpl_mode = build_deck(json.loads(plan), tpl, theme, use_images)
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        raise HTTPException(500, "Building the PPT failed: " + repr(e))
    return FileResponse(path, filename="presentation.pptx", headers={"X-Mode": "ai" if has_ai() else "offline", "X-Images": str(n_img), "X-Template": tpl_mode},
                        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")


@app.post("/api/rewrite")
async def api_rewrite(body: dict = Body(...)):
    bullets, mode = body.get("bullets", []), body.get("mode", "shorter")
    if not has_ai():
        return {"bullets": [" ".join(b.split()[:8]) for b in bullets] if mode == "shorter" else bullets}
    how = {"shorter": "make them shorter and punchier", "expand": "add useful detail",
           "formal": "use a formal tone", "casual": "use a friendly casual tone"}.get(mode, "improve them")
    prompt = ('Rewrite these slide bullets: %s. Max 5 bullets, max 12 words each. '
              'Reply with ONLY JSON: {"bullets": ["..."]}\nSlide title: %s\nBullets: %s'
              % (how, body.get("title", ""), json.dumps(bullets)))
    try:
        return json.loads(re.search(r"\{.*\}", call_llm(prompt, "lite"), re.S).group(0))
    except Exception as e:
        raise HTTPException(500, "Rewrite failed: " + str(e))


# ---------------- free photo library for the Media page ----------------
from urllib.parse import urlparse

STOCK_HOSTS = ("upload.wikimedia.org", "pixabay.com", "images.pexels.com")  # only these may be downloaded


def search_stock(q, n=10):
    out = []
    key = os.getenv("PIXABAY_API_KEY")
    if key:
        r = requests.get("https://pixabay.com/api/", timeout=15, params={
            "key": key, "q": q, "image_type": "photo", "orientation": "horizontal",
            "per_page": max(3, min(n, 50)), "safesearch": "true"})
        for hit in r.json().get("hits", [])[:n]:
            out.append({"thumb": hit["webformatURL"], "full": hit["largeImageURL"], "title": hit.get("tags", q)})
        if out:
            return out
    r = requests.get("https://commons.wikimedia.org/w/api.php", timeout=20, headers=UA, params={
        "action": "query", "generator": "search", "gsrsearch": "filetype:bitmap " + q,
        "gsrnamespace": 6, "gsrlimit": n + 8, "prop": "imageinfo", "iiprop": "url|mime",
        "iiurlwidth": 800, "format": "json"})
    pages = sorted(r.json().get("query", {}).get("pages", {}).values(), key=lambda p: p.get("index", 0))
    for p in pages:
        info = (p.get("imageinfo") or [{}])[0]
        if info.get("mime") == "image/jpeg" and info.get("thumburl"):
            title = re.sub(r"\.\w+$", "", p.get("title", "").replace("File:", ""))
            out.append({"thumb": info["thumburl"], "full": info["thumburl"], "title": title[:60]})
    return out[:n]


@app.get("/api/stock")
def api_stock(q: str, n: int = 10):
    try:
        return {"images": search_stock(q, max(1, min(n, 20)))}
    except Exception as e:
        raise HTTPException(500, "Photo search failed: " + str(e))


@app.get("/api/stock/image")
def api_stock_image(url: str):
    host = urlparse(url).hostname or ""
    if urlparse(url).scheme != "https" or not any(host == h or host.endswith("." + h) for h in STOCK_HOSTS):
        raise HTTPException(400, "That picture source is not allowed")
    try:
        r = requests.get(url, timeout=20, headers=UA)
        im = Image.open(io.BytesIO(r.content)).convert("RGB")
        im.thumbnail((1000, 1000))
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=75)
        return {"data": "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()}
    except Exception as e:
        raise HTTPException(500, "Could not download the picture: " + str(e))


# ---------------- 100 built-in images (drawn by the app, work offline) ----------------
LIB_CATS = {  # id: keywords that match a slide's topic to this category
    "technology": "computer data network software ai learning model algorithm digital cyber tech neural",
    "science": "lab chemical experiment research science physics biology",
    "nature": "plant leaf forest tree environment climate nature ecology",
    "medicine": "health medical patient disease hospital diagnosis clinical drug",
    "agriculture": "crop farm agricultur soil harvest farmer",
    "space": "space astronaut satellite orbit planet mission astronomy",
    "education": "education student school teach university class",
    "business": "business market finance economy cost revenue management",
    "city": "city urban building architecture infrastructure",
    "energy": "energy solar wind power electric battery renewable",
}
PALETTES = ["0B3C5D 328CC1 F2C14E", "1B4332 2D6A4F F4A261", "2B2D6E 5B5FC7 FFB703", "7A2E3A E4572E F3A712", "263238 546E7A 26C6DA",
            "4A3F8F 8E7CC3 F6BD60", "0F4C5C 2A9D8F E9C46A", "5C4033 A47551 E0A458", "004E64 00A5CF FFB30F", "590D22 C9184A FFB3C1"]
WHITE = (255, 255, 255)


def mix(a, b, t):
    return tuple(int(a[i] * (1 - t) + b[i] * t) for i in range(3))


def draw_scene(cat, n, W=1600, H=1000):
    dk, md, ac = [tuple(int(c[i:i + 2], 16) for i in (0, 2, 4)) for c in PALETTES[(n - 1) % 10].split()]
    r = random.Random(zlib.crc32(("%s%d" % (cat, n)).encode()))
    im = Image.new("RGB", (W, H))
    d = ImageDraw.Draw(im)
    end = (5, 5, 20) if cat == "space" else md
    for y in range(H):  # background gradient
        d.line([(0, y), (W, y)], fill=mix(dk, end, y / H * (0.9 if cat != "space" else 1)))
    cx, cy = W // 2, H // 2
    if cat == "technology":
        for _ in range(26):  # circuit traces
            x, y = r.randrange(0, W, 80), r.randrange(0, H, 80)
            pts = [(x, y)]
            for _ in range(3):
                x += r.choice([-1, 1, 0]) * r.randrange(80, 320, 80)
                y += r.choice([-1, 1, 0]) * r.randrange(80, 240, 80)
                pts.append((x, y))
            d.line(pts, fill=mix(md, WHITE, .35), width=6)
            d.ellipse([pts[-1][0] - 14, pts[-1][1] - 14, pts[-1][0] + 14, pts[-1][1] + 14], fill=ac)
        for i in range(-3, 4):  # chip pins
            d.rectangle([cx + i * 60 - 12, cy - 250, cx + i * 60 + 12, cy - 190], fill=ac)
            d.rectangle([cx + i * 60 - 12, cy + 190, cx + i * 60 + 12, cy + 250], fill=ac)
            d.rectangle([cx - 250, cy + i * 60 - 12, cx - 190, cy + i * 60 + 12], fill=ac)
            d.rectangle([cx + 190, cy + i * 60 - 12, cx + 250, cy + i * 60 + 12], fill=ac)
        d.rounded_rectangle([cx - 190, cy - 190, cx + 190, cy + 190], 36, fill=dk, outline=ac, width=10)
        d.rounded_rectangle([cx - 110, cy - 110, cx + 110, cy + 110], 20, fill=md)
    elif cat == "science":
        y0 = cy + 60
        d.polygon([(cx - 80, y0 - 300), (cx + 80, y0 - 300), (cx + 80, y0 - 120), (cx + 300, y0 + 250), (cx - 300, y0 + 250), (cx - 80, y0 - 120)], fill=WHITE)
        d.polygon([(cx - 175, y0 + 40), (cx + 175, y0 + 40), (cx + 300, y0 + 250), (cx - 300, y0 + 250)], fill=ac)
        d.rounded_rectangle([cx - 110, y0 - 330, cx + 110, y0 - 295], 12, fill=md)
        for _ in range(9):
            bx, by, br = cx + r.randrange(-150, 150), y0 + r.randrange(-330, 200), r.randrange(10, 28)
            d.ellipse([bx - br, by - br, bx + br, by + br], outline=WHITE, width=5)
        for tx, col in ((cx - 520, md), (cx + 440, ac)):  # test tubes
            d.rounded_rectangle([tx, y0 - 120, tx + 80, y0 + 250], 40, fill=WHITE)
            d.rounded_rectangle([tx + 8, y0 + 60, tx + 72, y0 + 242], 34, fill=col)
    elif cat == "nature" or cat == "agriculture":
        sx = r.randrange(int(W * .15), int(W * .85))
        d.ellipse([sx - 110, 120, sx + 110, 340], fill=ac)
        base = H * (0.52 if cat == "agriculture" else 0.5)
        if cat == "nature":
            for layer, col in enumerate((mix(md, WHITE, .3), md, mix(dk, (0, 0, 0), .2))):
                pts, x, b = [(0, H)], 0, base + layer * 110
                while x <= W + 200:
                    pts.append((x, b - r.randrange(60, 260) * (1 - layer * .2)))
                    x += r.randrange(160, 340)
                d.polygon(pts + [(W, H)], fill=col)
            for _ in range(16):
                x, y, h = r.randrange(0, W), r.randrange(int(H * .75), H - 40), r.randrange(110, 220)
                d.rectangle([x - 8, y, x + 8, y + 40], fill=(70, 45, 25))
                d.polygon([(x, y - h), (x - h * .32, y), (x + h * .32, y)], fill=mix(dk, (0, 90, 40), .5))
        else:
            d.polygon([(0, H)] + [(x, base + 40 * math.sin(x / 200)) for x in range(0, W + 1, 20)] + [(W, H)], fill=mix(dk, (0, 0, 0), .2))
            for i in range(14):  # field rows
                d.line([(W / 2, base + 20), (i * W / 13, H)], fill=mix(md, ac, .5), width=14)
            for _ in range(10):  # wheat
                x, y = r.randrange(60, W - 60), r.randrange(int(H * .68), H - 60)
                d.line([(x, y), (x, y - 190)], fill=ac, width=7)
                for k in range(5):
                    d.ellipse([x - 22, y - 190 - k * 26, x + 22, y - 160 - k * 26], fill=ac)
    elif cat == "medicine":
        d.rounded_rectangle([cx - 210, cy - 290, cx + 210, cy + 130], 70, fill=WHITE)
        d.rounded_rectangle([cx - 50, cy - 230, cx + 50, cy + 70], 20, fill=ac)
        d.rounded_rectangle([cx - 150, cy - 130, cx + 150, cy - 30], 20, fill=ac)
        y0, pts, x = H - 220, [], -40
        while x < W:
            pts += [(x, y0), (x + 90, y0), (x + 120, y0 - 100), (x + 160, y0 + 110), (x + 200, y0 - 40), (x + 230, y0), (x + 300, y0)]
            x += 300
        d.line(pts, fill=WHITE, width=12)
    elif cat == "space":
        for _ in range(260):
            x, y, s_ = r.randrange(W), r.randrange(H), r.choice([2, 2, 3, 4])
            d.ellipse([x - s_, y - s_, x + s_, y + s_], fill=mix((120, 120, 160), WHITE, r.random()))
        px, py = int(W * r.uniform(.35, .65)), cy
        d.arc([px - 420, py - 90, px + 420, py + 90], 180, 360, fill=ac, width=22)
        d.ellipse([px - 230, py - 230, px + 230, py + 230], fill=md)
        d.ellipse([px - 230, py - 230, px + 80, py + 80], fill=mix(md, WHITE, .18))
        d.arc([px - 420, py - 90, px + 420, py + 90], 0, 180, fill=ac, width=22)
        d.ellipse([W - 330, 140, W - 190, 280], fill=(225, 225, 230))
    elif cat == "education":
        y0 = cy + 90
        d.polygon([(cx, y0 - 120), (cx - 380, y0 - 60), (cx - 380, y0 + 220), (cx, y0 + 160)], fill=WHITE)
        d.polygon([(cx, y0 - 120), (cx + 380, y0 - 60), (cx + 380, y0 + 220), (cx, y0 + 160)], fill=mix(WHITE, md, .15))
        for i in range(5):
            for sgn in (-1, 1):
                d.line([(cx + sgn * 40, y0 - 60 + i * 46), (cx + sgn * 330, y0 - 8 + i * 46)], fill=md, width=8)
        d.polygon([(cx - 130, y0 - 330), (cx - 130, y0 - 250), (cx, y0 - 205), (cx + 130, y0 - 250), (cx + 130, y0 - 330)], fill=md)
        d.polygon([(cx, y0 - 450), (cx - 250, y0 - 350), (cx, y0 - 250), (cx + 250, y0 - 350)], fill=ac)
        d.line([(cx + 190, y0 - 335), (cx + 190, y0 - 230)], fill=ac, width=8)
    elif cat == "business":
        base, hs = H - 170, []
        for i in range(7):
            h = 140 + i * 70 + r.randrange(-30, 40)
            x = 240 + i * 170
            d.rounded_rectangle([x, base - h, x + 130, base], 14, fill=ac if i == 6 else mix(md, WHITE, .25))
            hs.append((x + 65, base - h - 50))
        d.line(hs, fill=WHITE, width=12)
        for px, py in hs:
            d.ellipse([px - 16, py - 16, px + 16, py + 16], fill=WHITE)
        d.line([(180, base + 6), (W - 180, base + 6)], fill=WHITE, width=8)
    elif cat == "city":
        d.ellipse([W - 360, 110, W - 160, 310], fill=ac)
        x = 0
        while x < W:
            w, h = r.randrange(110, 230), r.randrange(260, 700)
            col = mix(dk, md, r.random() * .6)
            d.rectangle([x, H - h, x + w, H], fill=col)
            for wy in range(H - h + 30, H - 40, 60):
                for wx in range(x + 22, x + w - 30, 46):
                    d.rectangle([wx, wy, wx + 22, wy + 30], fill=ac if r.random() < .45 else mix(col, WHITE, .18))
            x += w + r.randrange(0, 25)
    elif cat == "energy":
        d.ellipse([140, 110, 330, 300], fill=ac)
        d.polygon([(0, H)] + [(x, H * .8 + 30 * math.sin(x / 250)) for x in range(0, W + 1, 20)] + [(W, H)], fill=mix(dk, (0, 0, 0), .25))
        for tx, th in ((W * .45, 560), (W * .7, 440), (W * .9, 340)):
            hx, hy = tx, H * .8 - th
            d.polygon([(tx - 10, H * .82), (tx + 10, H * .82), (tx + 4, hy), (tx - 4, hy)], fill=WHITE)
            a0 = r.random() * 120
            for k in range(3):
                ang = math.radians(a0 + k * 120)
                d.line([(hx, hy), (hx + th * .42 * math.cos(ang), hy + th * .42 * math.sin(ang))], fill=WHITE, width=14)
            d.ellipse([hx - 16, hy - 16, hx + 16, hy + 16], fill=ac)
        pts = [(120, H - 150), (520, H - 150), (600, H - 370), (200, H - 370)]  # solar panel
        d.polygon(pts, fill=mix(md, (0, 0, 40), .4), outline=WHITE)
        for k in range(1, 5):
            t = k / 5
            d.line([(120 + 80 * t, H - 150 - 220 * t), (520 + 80 * t, H - 150 - 220 * t)], fill=WHITE, width=3)
            d.line([(120 + 400 * t, H - 150), (200 + 400 * t, H - 370)], fill=WHITE, width=3)
    return im.resize((1280, 800), Image.LANCZOS)


# 9 different icons per category (image 1 is the big scene above), so no two pictures look alike
MOTIFS = {
    "technology": "chip network gear bulb shield cloud globe magnifier document",
    "science": "flask atom dna drop magnifier bulb gear linechart planet",
    "nature": "mountain tree leaf sun cloud drop moon star globe",
    "medicine": "heart dna pill drop shield building flask magnifier document",
    "agriculture": "wheat tree sun cloud drop leaf building gear barchart",
    "space": "planet rocket star moon sun globe atom target network",
    "education": "cap book bulb document globe magnifier target people star",
    "business": "barchart linechart pie target people building gear bulb globe",
    "city": "building shield cloud sun network gear bolt people globe",
    "energy": "turbine solar battery bolt sun gear drop leaf bulb",
}


def motif(name, d, cx, cy, k, ac):
    w = WHITE
    c = lambda x, y, r, **kw: d.ellipse([x - r, y - r, x + r, y + r], **kw)
    rr = lambda x0, y0, x1, y1, rad, **kw: d.rounded_rectangle([cx + x0 * k, cy + y0 * k, cx + x1 * k, cy + y1 * k], int(rad), **kw)
    poly = lambda pts, **kw: d.polygon([(cx + x * k, cy + y * k) for x, y in pts], **kw)
    ln = lambda pts, **kw: d.line([(cx + x * k, cy + y * k) for x, y in pts], **kw)
    if name == "chip":
        for i in (-.5, 0, .5):
            for x0, y0, x1, y1 in ((i - .06, -1, i + .06, -.8), (i - .06, .8, i + .06, 1), (-1, i - .06, -.8, i + .06), (.8, i - .06, 1, i + .06)):
                rr(x0, y0, x1, y1, 4, fill=ac)
        rr(-.8, -.8, .8, .8, 30, fill=w)
        rr(-.4, -.4, .4, .4, 14, fill=ac)
    elif name == "network":
        pts = [(0, 0), (-1, -.6), (1, -.5), (-.8, .7), (.9, .6)]
        for p in pts[1:]:
            ln([pts[0], p], fill=w, width=8)
        for i, (x, y) in enumerate(pts):
            c(cx + x * k, cy + y * k, k * (.26 if i == 0 else .17), fill=ac if i == 0 else w)
    elif name == "gear":
        for i in range(8):
            a = i * math.pi / 4
            px, py = cx + .85 * k * math.cos(a), cy + .85 * k * math.sin(a)
            d.polygon([(px + dx * math.cos(a) - dy * math.sin(a), py + dx * math.sin(a) + dy * math.cos(a))
                       for dx, dy in ((-.18 * k, -.17 * k), (.18 * k, -.17 * k), (.18 * k, .17 * k), (-.18 * k, .17 * k))], fill=w)
        c(cx, cy, .72 * k, fill=w)
        c(cx, cy, .3 * k, fill=ac)
    elif name == "bulb":
        c(cx, cy - .2 * k, .65 * k, fill=w)
        poly([(-.3, .35), (.3, .35), (.22, .8), (-.22, .8)], fill=w)
        ln([(-.2, -.05), (0, -.3), (.2, -.05)], fill=ac, width=10)
        for a in (-60, -30, 0, 30, 60):
            t = math.radians(a - 90)
            d.line([(cx + .85 * k * math.cos(t), cy - .2 * k + .85 * k * math.sin(t)), (cx + 1.05 * k * math.cos(t), cy - .2 * k + 1.05 * k * math.sin(t))], fill=ac, width=10)
    elif name == "shield":
        poly([(-.7, -.7), (.7, -.7), (.7, .1), (0, .9), (-.7, .1)], fill=w)
        ln([(-.3, 0), (-.05, .25), (.35, -.25)], fill=ac, width=22)
    elif name == "cloud":
        c(cx - .35 * k, cy, .4 * k, fill=w)
        c(cx + .1 * k, cy - .25 * k, .5 * k, fill=w)
        c(cx + .5 * k, cy + .05 * k, .35 * k, fill=w)
        rr(-.75, 0, .85, .4, 40, fill=w)
    elif name == "globe":
        c(cx, cy, .8 * k, outline=w, width=12)
        d.ellipse([cx - .35 * k, cy - .8 * k, cx + .35 * k, cy + .8 * k], outline=w, width=8)
        for y in (-.4, 0, .4):
            ln([(-.78 * math.sqrt(1 - y * y / .64) if y else -.8, y), (.78 * math.sqrt(1 - y * y / .64) if y else .8, y)], fill=w, width=8)
        c(cx + .45 * k, cy - .35 * k, .13 * k, fill=ac)
    elif name == "magnifier":
        c(cx - .15 * k, cy - .15 * k, .55 * k, outline=w, width=24)
        ln([(.25, .25), (.8, .8)], fill=ac, width=34)
    elif name == "document":
        poly([(-.55, -.8), (.2, -.8), (.55, -.45), (.55, .8), (-.55, .8)], fill=w)
        for i in range(4):
            ln([(-.35, -.2 + i * .28), (.35, -.2 + i * .28)], fill=ac, width=12)
    elif name == "flask":
        poly([(-.2, -.8), (.2, -.8), (.2, -.2), (.7, .7), (-.7, .7), (-.2, -.2)], fill=w)
        poly([(-.46, .2), (.46, .2), (.7, .7), (-.7, .7)], fill=ac)
    elif name == "atom":
        d.ellipse([cx - .95 * k, cy - .35 * k, cx + .95 * k, cy + .35 * k], outline=w, width=10)
        d.ellipse([cx - .35 * k, cy - .95 * k, cx + .35 * k, cy + .95 * k], outline=w, width=10)
        c(cx, cy, .22 * k, fill=ac)
        c(cx + .95 * k, cy, .1 * k, fill=ac)
        c(cx, cy - .95 * k, .1 * k, fill=ac)
    elif name == "dna":
        for i in range(25):
            t = i / 24
            y, x = -.9 + 1.8 * t, .5 * math.sin(t * 4 * math.pi)
            ln([(x, y), (-x, y)], fill=mix(w, ac, .5), width=6)
            c(cx + x * k, cy + y * k, .07 * k, fill=w)
            c(cx - x * k, cy + y * k, .07 * k, fill=ac)
    elif name == "drop":
        poly([(0, -.9), (-.55, .1), (.55, .1)], fill=w)
        c(cx, cy + .15 * k, .55 * k, fill=w)
        d.ellipse([cx - .3 * k, cy + .05 * k, cx - .1 * k, cy + .3 * k], fill=ac)
    elif name in ("linechart", "barchart"):
        ln([(-.9, -.8), (-.9, .8), (.9, .8)], fill=w, width=10)
        vals = [.3, .1, .35, -.3, -.1, -.6]
        if name == "linechart":
            pts = [(-.7 + i * .28, v + .1) for i, v in enumerate(vals)]
            ln(pts, fill=w, width=12)
            for x, y in pts:
                c(cx + x * k, cy + y * k, .07 * k, fill=ac)
        else:
            for i, v in enumerate((.4, .2, .55, .8)):
                rr(-.7 + i * .43, .8 - v * 1.5, -.4 + i * .43, .8, 8, fill=ac if i == 3 else w)
    elif name == "planet":
        c(cx, cy, .6 * k, fill=w)
        d.ellipse([cx - k, cy - .25 * k, cx + k, cy + .25 * k], outline=ac, width=14)
    elif name == "leaf":
        a = math.radians(-35)
        pts = [(-.85 + 1.7 * t / 20, s_ * .45 * math.sin(math.pi * t / 20)) for s_ in (-1, 1) for t in (range(21) if s_ < 0 else range(20, -1, -1))]
        d.polygon([(cx + (x * math.cos(a) - y * math.sin(a)) * k, cy + (x * math.sin(a) + y * math.cos(a)) * k) for x, y in pts], fill=w)
        d.line([(cx - .85 * k * math.cos(a), cy - .85 * k * math.sin(a)), (cx + .8 * k * math.cos(a), cy + .8 * k * math.sin(a))], fill=ac, width=10)
    elif name == "mountain":
        poly([(-1, .7), (-.3, -.6), (.1, 0), (.5, -.9), (1.1, .7)], fill=w)
        c(cx - .75 * k, cy - .6 * k, .17 * k, fill=ac)
    elif name == "tree":
        d.rectangle([cx - .08 * k, cy + .2 * k, cx + .08 * k, cy + .8 * k], fill=w)
        c(cx, cy - .2 * k, .55 * k, fill=w)
        c(cx - .38 * k, cy, .36 * k, fill=w)
        c(cx + .38 * k, cy, .36 * k, fill=w)
    elif name == "sun":
        c(cx, cy, .45 * k, fill=ac)
        for i in range(12):
            t = i * math.pi / 6
            d.line([(cx + .62 * k * math.cos(t), cy + .62 * k * math.sin(t)), (cx + .9 * k * math.cos(t), cy + .9 * k * math.sin(t))], fill=w, width=14)
    elif name == "moon":
        c(cx, cy, .7 * k, fill=w)
        for x, y, r_ in ((-.25, -.2, .17), (.2, .15, .22), (-.1, .4, .1)):
            c(cx + x * k, cy + y * k, r_ * k, fill=mix(w, ac, .35))
    elif name == "heart":
        c(cx - .35 * k, cy - .2 * k, .45 * k, fill=w)
        c(cx + .35 * k, cy - .2 * k, .45 * k, fill=w)
        poly([(-.78, -.05), (.78, -.05), (0, .85)], fill=w)
        ln([(-.6, .05), (-.25, .05), (-.1, -.25), (.1, .3), (.25, .05), (.6, .05)], fill=ac, width=10)
    elif name == "pill":
        rr(-.9, -.35, .9, .35, .35 * k, fill=w)
        rr(0, -.35, .9, .35, .35 * k, fill=ac)
        d.rectangle([cx, cy - .35 * k, cx + .5 * k, cy + .35 * k], fill=ac)
    elif name == "building":
        rr(-.55, -.9, .35, .8, 8, fill=w)
        rr(.35, -.2, .85, .8, 8, fill=mix(w, ac, .3))
        for y in (-.7, -.4, -.1, .2, .5):
            for x in (-.4, -.1):
                d.rectangle([cx + x * k, cy + y * k, cx + (x + .15) * k, cy + (y + .15) * k], fill=ac)
    elif name == "wheat":
        for off in (-.5, 0, .5):
            ln([(off, .85), (off, -.5)], fill=w, width=8)
            for j in range(5):
                c(cx + (off - .1) * k, cy + (-.5 + j * .14) * k, .06 * k, fill=ac)
                c(cx + (off + .1) * k, cy + (-.43 + j * .14) * k, .06 * k, fill=ac)
            c(cx + off * k, cy - .62 * k, .06 * k, fill=ac)
    elif name == "rocket":
        d.ellipse([cx - .3 * k, cy - .9 * k, cx + .3 * k, cy + .5 * k], fill=w)
        c(cx, cy - .3 * k, .14 * k, fill=ac)
        poly([(-.3, .1), (-.7, .6), (-.3, .45)], fill=ac)
        poly([(.3, .1), (.7, .6), (.3, .45)], fill=ac)
        poly([(-.15, .5), (.15, .5), (0, .95)], fill=mix(ac, w, .4))
    elif name == "star":
        pts = [(math.sin(i * math.pi / 5) * (.9 if i % 2 == 0 else .4), -math.cos(i * math.pi / 5) * (.9 if i % 2 == 0 else .4)) for i in range(10)]
        poly(pts, fill=w)
    elif name == "target":
        for r_, col in ((.85, w), (.62, ac), (.4, w), (.18, ac)):
            c(cx, cy, r_ * k, fill=col)
        ln([(0, 0), (.9, -.9)], fill=w, width=10)
    elif name == "book":
        poly([(0, -.4), (-.9, -.25), (-.9, .6), (0, .45)], fill=w)
        poly([(0, -.4), (.9, -.25), (.9, .6), (0, .45)], fill=mix(w, ac, .25))
        for i in range(4):
            ln([(-.75, -.1 + i * .16), (-.1, 0 + i * .15)], fill=ac, width=7)
    elif name == "cap":
        poly([(-.5, 0), (-.5, .35), (0, .6), (.5, .35), (.5, 0)], fill=mix(w, ac, .3))
        poly([(0, -.7), (-.95, -.3), (0, .1), (.95, -.3)], fill=w)
        ln([(.75, -.25), (.75, .3)], fill=ac, width=8)
    elif name == "turbine":
        poly([(-.05, .9), (.05, .9), (.03, -.3), (-.03, -.3)], fill=w)
        for i in range(3):
            t = math.radians(-90 + i * 120)
            d.line([(cx, cy - .3 * k), (cx + .8 * k * math.cos(t), cy - .3 * k + .8 * k * math.sin(t))], fill=w, width=16)
        c(cx, cy - .3 * k, .09 * k, fill=ac)
    elif name == "solar":
        poly([(-.8, .3), (.8, .3), (.55, -.5), (-.55, -.5)], fill=mix(w, ac, .15))
        for i in range(1, 4):
            ln([(-.8 + .27 * i * 2 / 1.5 * .75, .3), (-.55 + .275 * i * .75 * 1.0, -.5)], fill=ac, width=4)
            ln([(-.7 + i * 0.0, .3 - i * .2), (.7 - 0, .3 - i * .2)], fill=ac, width=4)
        ln([(0, .3), (0, .85)], fill=w, width=12)
        c(cx + .85 * k, cy - .7 * k, .2 * k, fill=ac)
    elif name == "battery":
        rr(-.85, -.4, .7, .4, 22, outline=w, width=14)
        rr(.7, -.15, .85, .15, 6, fill=w)
        for i in range(3):
            rr(-.7 + i * .43, -.27, -.35 + i * .43, .27, 8, fill=ac if i < 2 else mix(w, ac, .5))
    elif name == "bolt":
        poly([(.1, -.9), (-.5, .1), (-.05, .1), (-.15, .9), (.5, -.15), (.05, -.15)], fill=w)
    elif name == "pie":
        d.pieslice([cx - .75 * k, cy - .75 * k, cx + .75 * k, cy + .75 * k], 0, 250, fill=w)
        d.pieslice([cx - .65 * k, cy - .65 * k, cx + .85 * k, cy + .85 * k], 250, 360, fill=ac)
    elif name == "people":
        for dx, col, sc in ((-.65, w, .8), (0, ac, 1), (.65, w, .8)):
            c(cx + dx * k, cy - .35 * k * sc, .22 * k * sc, fill=col)
            d.pieslice([cx + dx * k - .36 * k * sc, cy - .05 * k * sc, cx + dx * k + .36 * k * sc, cy + .67 * k * sc], 180, 360, fill=col)


def draw_library(cat, n):
    if n == 1:  # the big scene for this category
        return draw_scene(cat, n)
    W, H = 1600, 1000
    dk, md, ac = [tuple(int(c[i:i + 2], 16) for i in (0, 2, 4)) for c in PALETTES[(n - 1) % 10].split()]
    r = random.Random(zlib.crc32(("%s%d" % (cat, n)).encode()))
    im = Image.new("RGB", (W, H))
    d = ImageDraw.Draw(im)
    style = (n - 2) % 4
    if style == 1:  # diagonal gradient
        for i in range(0, W + H, 2):
            d.line([(i, 0), (i - H, H)], fill=mix(dk, md, i / (W + H)), width=3)
    elif style == 2:  # dark with big rings
        d.rectangle([0, 0, W, H], fill=dk)
        for k_ in range(7):
            d.ellipse([W / 2 - 150 - k_ * 120, H / 2 - 150 - k_ * 120, W / 2 + 150 + k_ * 120, H / 2 + 150 + k_ * 120], outline=mix(dk, md, .35 + k_ * .05), width=6)
    elif style == 3:  # light with dots and a dark panel behind the icon
        d.rectangle([0, 0, W, H], fill=mix(WHITE, md, .18))
        for x in range(60, W, 90):
            for y in range(60, H, 90):
                d.ellipse([x - 4, y - 4, x + 4, y + 4], fill=mix(WHITE, md, .4))
    else:  # vertical gradient
        for y in range(H):
            d.line([(0, y), (W, y)], fill=mix(dk, md, y / H))
    for _ in range(6):  # a few soft decorative circles
        x, y, rad = r.randrange(W), r.randrange(H), r.randrange(40, 130)
        d.ellipse([x - rad, y - rad, x + rad, y + rad], outline=mix(md, WHITE, .35), width=5)
    cx, cy, k = W // 2, H // 2, 330
    if style == 3:
        d.rounded_rectangle([cx - 400, cy - 400, cx + 400, cy + 400], 120, fill=dk)
        d.ellipse([cx - 330, cy - 330, cx + 330, cy + 330], fill=mix(dk, md, .35))
    motif(MOTIFS[cat].split()[n - 2], d, cx, cy, k, ac)
    return im.resize((1280, 800), Image.LANCZOS)


_lib_cache = {}


def lib_jpeg(cat, n):
    if (cat, n) not in _lib_cache:
        buf = io.BytesIO()
        draw_library(cat, n).save(buf, "JPEG", quality=82)
        _lib_cache[(cat, n)] = buf.getvalue()
    return _lib_cache[(cat, n)]


def library_image(query, used):
    # picks the built-in category that matches the slide topic, a different picture for every slide
    q = (query or "").lower()
    best, score = "technology", 0
    for cat, words in LIB_CATS.items():
        hits = sum(1 for w in words.split() if w in q)
        if hits > score:
            best, score = cat, hits
    for cat in [best] + [c for c in LIB_CATS if c != best]:
        for n in range(1, 11):
            if ("lib", cat, n) not in used:
                used.add(("lib", cat, n))
                return io.BytesIO(lib_jpeg(cat, n))
    return None


@app.get("/library/{cat}/{n}.jpg")
def library_file(cat: str, n: int):
    if cat not in LIB_CATS or not 1 <= n <= 10:
        raise HTTPException(404, "No such image")
    return Response(content=lib_jpeg(cat, n), media_type="image/jpeg", headers={"Cache-Control": "max-age=86400"})