import io
import os
import json
import datetime
from flask import Flask, request, jsonify, render_template, send_file
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from style_extractor import extract_styles, apply_styles_to_slide
from deep_scanner import deep_scan

app = Flask(__name__)

# ── Configuration ──────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
ALLOWED_EXTENSIONS = {"pptx"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


# ── Helpers ────────────────────────────────────────────────────────────────────
def allowed_file(filename: str) -> bool:
    """Return True if the file has an allowed extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def placeholder_type_name(ph_type) -> str:
    """Convert a placeholder type enum to a human-readable string."""
    try:
        return ph_type.name  # e.g. 'TITLE', 'BODY', 'PICTURE', …
    except AttributeError:
        return str(ph_type)


def scan_layout(pptx_path: str) -> list[dict]:
    """
    Scan every slide layout in the given .pptx file and return a list of
    layout descriptors.

    Each descriptor has the shape:
    {
        "layout_index": int,
        "layout_name":  str,
        "placeholders": [
            {"idx": int, "name": str, "type": str},
            ...
        ]
    }
    """
    prs = Presentation(pptx_path)
    layouts = []

    for index, layout in enumerate(prs.slide_layouts):
        placeholders = []
        for ph in layout.placeholders:
            placeholders.append(
                {
                    "idx": ph.placeholder_format.idx,
                    "name": ph.name,
                    "type": placeholder_type_name(ph.placeholder_format.type),
                }
            )

        layouts.append(
            {
                "layout_index": index,
                "layout_name": layout.name,
                "placeholders": placeholders,
            }
        )

    return layouts


# ── Routes ─────────────────────────────────────────────────────────────────────
def _layout_json_path(pptx_filename: str) -> str:
    """Return the path to the sidecar layout JSON for a given PPTX filename."""
    stem = os.path.splitext(secure_filename(pptx_filename))[0]
    return os.path.join(app.config["UPLOAD_FOLDER"], f"{stem}.layout.json")


def _styles_json_path(pptx_filename: str) -> str:
    """Return the path to the deep-format styles JSON for a given PPTX filename."""
    stem = os.path.splitext(secure_filename(pptx_filename))[0]
    return os.path.join(app.config["UPLOAD_FOLDER"], f"{stem}.styles.json")


def _profile_json_path(pptx_filename: str) -> str:
    """Return the path to the deep-scan profile JSON for a given PPTX filename."""
    stem = os.path.splitext(secure_filename(pptx_filename))[0]
    return os.path.join(app.config["UPLOAD_FOLDER"], f"{stem}.profile.json")


def _structure_json_path(pptx_filename: str) -> str:
    """Return the path to the deep-scan structure JSON for a given PPTX filename."""
    stem = os.path.splitext(secure_filename(pptx_filename))[0]
    return os.path.join(app.config["UPLOAD_FOLDER"], f"{stem}.structure.json")


# ── Built-in profile paths ─────────────────────────────────────────────────────
MASTER_PROFILE_PATH   = os.path.join(BASE_DIR, "master_profile.json")
MASTER_STRUCTURE_PATH = os.path.join(BASE_DIR, "master_structure.json")
MASTER_SLIDE_PATH     = os.path.join(BASE_DIR, "master_slide.pptx")

# Alignment string → python-pptx enum
_ALIGN_MAP = {
    "LEFT":    PP_ALIGN.LEFT,
    "CENTER":  PP_ALIGN.CENTER,
    "RIGHT":   PP_ALIGN.RIGHT,
    "JUSTIFY": PP_ALIGN.JUSTIFY,
}


def _hex_to_rgb(hex_color: str) -> RGBColor:
    """Convert '#RRGGBB' to RGBColor."""
    h = hex_color.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def load_full_profile() -> dict:
    """
    Merge master_profile.json + master_structure.json into a single dict.
    Uses utf-8-sig to handle optional BOM.
    """
    with open(MASTER_PROFILE_PATH, "r", encoding="utf-8-sig") as f:
        profile = json.load(f)
    with open(MASTER_STRUCTURE_PATH, "r", encoding="utf-8-sig") as f:
        structure = json.load(f)
    all_layouts = profile.get("layouts", []) + structure.get("additional_layouts", [])
    all_layouts.sort(key=lambda l: l["layout_index"])
    return {
        "template_identity": profile.get("template_identity"),
        "canvas_size":       profile.get("canvas_size"),
        "color_palette":     profile.get("color_palette"),
        "fixed_elements":    profile.get("fixed_elements"),
        "common_elements":   structure.get("common_elements"),
        "layouts":           all_layouts,
    }


@app.route("/")
def index():
    """Serve the main upload page."""
    return render_template("index.html")


# ── /list_masters ──────────────────────────────────────────────────────────────
@app.route("/list_masters", methods=["GET"])
def list_masters():
    """Return all saved master layout records (newest first)."""
    records = []
    for fname in os.listdir(app.config["UPLOAD_FOLDER"]):
        if not fname.endswith(".layout.json"):
            continue
        fpath = os.path.join(app.config["UPLOAD_FOLDER"], fname)
        try:
            with open(fpath, "r", encoding="utf-8") as f:
                data = json.load(f)
            pptx_file = data.get("filename", "")
            pptx_exists = os.path.isfile(
                os.path.join(app.config["UPLOAD_FOLDER"], pptx_file)
            )
            records.append({
                "filename":      pptx_file,
                "layout_file":   fname,
                "saved_at":      data.get("saved_at", ""),
                "total_layouts": data.get("total_layouts", 0),
                "pptx_exists":   pptx_exists,
            })
        except Exception:
            continue
    records.sort(key=lambda r: r["saved_at"], reverse=True)
    return jsonify({"masters": records})


# ── /load_layout/<filename> ───────────────────────────────────────────────────
@app.route("/load_layout/<path:filename>", methods=["GET"])
def load_layout(filename):
    """Load and return a previously saved layout JSON."""
    fpath = _layout_json_path(filename)
    if not os.path.isfile(fpath):
        return jsonify({"error": f"No saved layout found for '{filename}'."}), 404
    with open(fpath, "r", encoding="utf-8") as f:
        data = json.load(f)
    return jsonify(data)


# ── /update_layout/<filename> ─────────────────────────────────────────────────
@app.route("/update_layout/<path:filename>", methods=["POST"])
def update_layout(filename):
    """
    Overwrite the saved layout JSON for *filename*.

    Expected body: { "layouts": [ ... ] }
    """
    body = request.get_json(silent=True)
    if not body or "layouts" not in body:
        return jsonify({"error": "Request must contain a 'layouts' array."}), 400

    fpath = _layout_json_path(filename)
    if not os.path.isfile(fpath):
        return jsonify({"error": f"No saved layout found for '{filename}'."}), 404

    with open(fpath, "r", encoding="utf-8") as f:
        existing = json.load(f)

    existing["layouts"]       = body["layouts"]
    existing["total_layouts"] = len(body["layouts"])
    existing["saved_at"]      = datetime.datetime.now().isoformat(timespec="seconds")

    with open(fpath, "w", encoding="utf-8") as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    return jsonify({
        "message":       "Layout updated successfully.",
        "filename":      existing["filename"],
        "total_layouts": existing["total_layouts"],
        "saved_at":      existing["saved_at"],
        "layouts":       existing["layouts"],
    })


@app.route("/upload_master", methods=["POST"])
def upload_master():
    """
    Accept a .pptx file upload, scan its slide layouts, and return the
    layout structure as JSON.

    Expected request: multipart/form-data with a field named 'file'.
    """
    if "file" not in request.files:
        return jsonify({"error": "No file field in request."}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only .pptx files are supported."}), 415

    filename = secure_filename(file.filename)
    save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(save_path)

    try:
        layout_data = scan_layout(save_path)
    except Exception as exc:
        return jsonify({"error": f"Failed to parse PowerPoint file: {exc}"}), 500

    # ── Auto-save layout JSON beside the PPTX ─────────────────────────────────
    layout_record = {
        "filename": filename,
        "saved_at": datetime.datetime.now().isoformat(timespec="seconds"),
        "total_layouts": len(layout_data),
        "layouts": layout_data,
    }
    layout_json_path = _layout_json_path(filename)
    with open(layout_json_path, "w", encoding="utf-8") as f:
        json.dump(layout_record, f, ensure_ascii=False, indent=2)

    # ── Auto-extract and save deep styles JSON ────────────────────────────────
    try:
        styles_data = extract_styles(save_path)
        styles_data["filename"] = filename
        styles_data["saved_at"] = layout_record["saved_at"]
        with open(_styles_json_path(filename), "w", encoding="utf-8") as f:
            json.dump(styles_data, f, ensure_ascii=False, indent=2)
        layout_record["styles_saved"] = True
    except Exception as exc:
        layout_record["styles_saved"] = False
        layout_record["styles_error"] = str(exc)

    # ── Auto-run deep scan (profile + structure) ──────────────────────────────
    try:
        ds = deep_scan(save_path)
        ds["master_profile"]["filename"] = filename
        ds["master_profile"]["saved_at"] = layout_record["saved_at"]
        ds["master_structure"]["filename"] = filename
        ds["master_structure"]["saved_at"] = layout_record["saved_at"]
        with open(_profile_json_path(filename), "w", encoding="utf-8") as f:
            json.dump(ds["master_profile"], f, ensure_ascii=False, indent=2)
        with open(_structure_json_path(filename), "w", encoding="utf-8") as f:
            json.dump(ds["master_structure"], f, ensure_ascii=False, indent=2)
        layout_record["deep_scan_saved"] = True
        layout_record["deep_scan"] = ds
    except Exception as exc:
        layout_record["deep_scan_saved"] = False
        layout_record["deep_scan_error"] = str(exc)

    return jsonify(layout_record)


# ── /extract_styles/<filename> ───────────────────────────────────────────────
@app.route("/extract_styles/<path:filename>", methods=["GET"])
def extract_styles_route(filename):
    """
    (Re-)extract deep formatting from the stored master PPTX and return it.
    Also overwrites the sidecar .styles.json on disk.
    """
    pptx_path = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(filename))
    if not os.path.isfile(pptx_path):
        return jsonify({"error": f"Master PPTX '{filename}' not found."}), 404
    try:
        data = extract_styles(pptx_path)
        data["filename"] = filename
        data["saved_at"] = datetime.datetime.now().isoformat(timespec="seconds")
        with open(_styles_json_path(filename), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return jsonify(data)
    except Exception as exc:
        return jsonify({"error": f"Extraction failed: {exc}"}), 500


# ── /style_info/<filename> ────────────────────────────────────────────────────
@app.route("/style_info/<path:filename>", methods=["GET"])
def style_info(filename):
    """Return the previously saved deep-format styles JSON."""
    fpath = _styles_json_path(filename)
    if not os.path.isfile(fpath):
        return jsonify({"error": f"No styles found for '{filename}'. Upload the file first."}), 404
    with open(fpath, "r", encoding="utf-8") as f:
        data = json.load(f)
    return jsonify(data)


# ── /deep_scan/<filename> ─────────────────────────────────────────────────────
@app.route("/deep_scan/<path:filename>", methods=["GET"])
def deep_scan_route(filename):
    """
    (Re-)run deep scan on the stored master PPTX and return (+ save)
    master_profile + master_structure JSONs.
    """
    pptx_path = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(filename))
    if not os.path.isfile(pptx_path):
        return jsonify({"error": f"Master PPTX '{filename}' not found."}), 404
    try:
        ds = deep_scan(pptx_path)
        now = datetime.datetime.now().isoformat(timespec="seconds")
        ds["master_profile"]["filename"] = filename
        ds["master_profile"]["saved_at"] = now
        ds["master_structure"]["filename"] = filename
        ds["master_structure"]["saved_at"] = now
        with open(_profile_json_path(filename), "w", encoding="utf-8") as f:
            json.dump(ds["master_profile"], f, ensure_ascii=False, indent=2)
        with open(_structure_json_path(filename), "w", encoding="utf-8") as f:
            json.dump(ds["master_structure"], f, ensure_ascii=False, indent=2)
        return jsonify(ds)
    except Exception as exc:
        return jsonify({"error": f"Deep scan failed: {exc}"}), 500


# ── /scan_profile/<filename> ──────────────────────────────────────────────────
@app.route("/scan_profile/<path:filename>", methods=["GET"])
def scan_profile(filename):
    """Return the previously saved deep-scan profile JSON."""
    fpath = _profile_json_path(filename)
    if not os.path.isfile(fpath):
        return jsonify({"error": f"No profile found for '{filename}'. Upload and scan first."}), 404
    with open(fpath, "r", encoding="utf-8") as f:
        data = json.load(f)
    return jsonify(data)


# ── /scan_structure/<filename> ────────────────────────────────────────────────
@app.route("/scan_structure/<path:filename>", methods=["GET"])
def scan_structure(filename):
    """Return the previously saved deep-scan structure JSON."""
    fpath = _structure_json_path(filename)
    if not os.path.isfile(fpath):
        return jsonify({"error": f"No structure found for '{filename}'. Upload and scan first."}), 404
    with open(fpath, "r", encoding="utf-8") as f:
        data = json.load(f)
    return jsonify(data)


# ── /layout_schema ────────────────────────────────────────────────────────────
@app.route("/layout_schema", methods=["GET"])
def layout_schema():
    """Return the merged profile+structure layout schema (used by the frontend AI prompt)."""
    try:
        return jsonify(load_full_profile())
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# ── Profile-based PPTX builder ───────────────────────────────────────────────
def _apply_font_to_run(run, font_def: dict) -> None:
    """Apply font overrides from master_profile.json to a run.

    Color is applied only when:
    - color_from_xml is True (set explicitly in the placeholder XML), OR
    - the font_def has no 'color_from_xml' key (hand-crafted profile entries)
    This preserves the master PPTX theme-color chain for placeholders whose
    colours are only inherited via the cascade.
    """
    if not font_def:
        return
    if font_def.get("name"):
        run.font.name = font_def["name"]
    size = font_def.get("size")
    if size is not None:
        try:
            run.font.size = Pt(float(size))
        except Exception:
            pass
    bold = font_def.get("bold")
    if bold is not None:
        run.font.bold = bool(bold)
    color = font_def.get("color")
    # Only apply color if:
    #   1. It is a valid '#RRGGBB' hex string
    #   2. It was found directly in the placeholder XML (color_from_xml=True)
    #      OR the entry has no such flag (backward-compat with hand-crafted profiles)
    color_from_xml = font_def.get("color_from_xml", True)
    if (color and isinstance(color, str)
            and color.startswith("#") and len(color) == 7
            and color_from_xml):
        try:
            run.font.color.rgb = _hex_to_rgb(color)
        except Exception:
            pass


def _apply_para_format(para, font_def: dict, para_def: dict) -> None:
    """Apply paragraph-level formatting from the profile."""
    align_str = font_def.get("align")
    if align_str:
        para.alignment = _ALIGN_MAP.get(align_str.upper(), PP_ALIGN.LEFT)
    ls = para_def.get("line_spacing")
    if ls:
        para.line_spacing = float(ls)


def create_pptx_from_profile(json_data: dict) -> io.BytesIO:
    """
    Build a PPTX using master_slide.pptx as the template — writing directly
    into native layout placeholders (not textboxes).

    This approach preserves ALL slide master properties:
    - Backgrounds, themes, decorative shapes
    - Placeholder positions and sizes
    - Default fonts and colors

    Font overrides from master_profile.json are applied on top of the
    native formatting.
    """
    if not os.path.isfile(MASTER_SLIDE_PATH):
        raise FileNotFoundError(
            f"master_slide.pptx not found at {MASTER_SLIDE_PATH}. "
            "Place the file in the project root directory."
        )

    # Build a lookup: layout_index → { ph_idx → {font, paragraph} }
    full_profile = load_full_profile()
    profile_map: dict = {}
    for l in full_profile.get("layouts", []):
        ph_map = {}
        for p in l.get("placeholders", []):
            ph_map[p["idx"]] = {
                "font": p.get("font", {}),
                "paragraph": p.get("paragraph", {}),
            }
        profile_map[l["layout_index"]] = ph_map

    prs = Presentation(MASTER_SLIDE_PATH)

    for slide_data in json_data.get("slides", []):
        layout_index = int(slide_data.get("layout_index", 0))
        if layout_index >= len(prs.slide_layouts):
            layout_index = 0

        layout = prs.slide_layouts[layout_index]
        slide  = prs.slides.add_slide(layout)

        # Build content map: ph_idx → {content, type}
        content_map: dict = {}
        if "title" in slide_data:
            content_map[0] = {"content": slide_data["title"], "type": "text"}
        for ph in slide_data.get("placeholders", []):
            idx = int(ph.get("id", ph.get("idx", 0)))
            content_map[idx] = {
                "content": ph.get("content", ""),
                "type":    ph.get("type", "text"),
            }

        # Profile font/paragraph overrides for this layout
        ph_styles = profile_map.get(layout_index, {})

        # Write into native placeholders
        for ph_idx, info in content_map.items():
            # Find the placeholder on the slide
            target = None
            try:
                target = slide.placeholders[ph_idx]
            except KeyError:
                for ph in slide.placeholders:
                    if ph.placeholder_format.idx == ph_idx:
                        target = ph
                        break

            if target is None:
                continue  # no such placeholder on this layout

            style    = ph_styles.get(ph_idx, {})
            font_def = style.get("font", {})
            para_def = style.get("paragraph", {})
            content  = info["content"]
            c_type   = info["type"]

            tf = target.text_frame
            tf.clear()

            if c_type == "list" and isinstance(content, list):
                items = content
            elif c_type == "list":
                items = [str(content)]
            else:
                items = str(content).split("\n") if content else [""]

            for i, item_text in enumerate(items):
                para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                _apply_para_format(para, font_def, para_def)
                run = para.add_run()
                run.text = str(item_text)
                _apply_font_to_run(run, font_def)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ── Core generator ────────────────────────────────────────────────────────────
def create_pptx_from_json(json_data: dict, master_path: str) -> io.BytesIO:
    """
    Build a new PowerPoint file from *json_data* using *master_path* as the
    slide-master template.

    Expected json_data shape
    ------------------------
    {
      "presentation_name": "Optional title",   # optional
      "slides": [
        {
          "layout_index": 0,
          "title": "Slide title",               # optional shortcut
          "placeholders": [
            { "id": 1, "content": "plain text",        "type": "text" },
            { "id": 2, "content": ["item1","item2"],   "type": "list" }
          ]
        },
        ...
      ]
    }

    Returns a BytesIO object positioned at offset 0 ready for streaming.
    """
    prs = Presentation(master_path)

    # ── Load deep scan profile for font overrides (if available) ──────────────
    _ds_profile_map: dict = {}
    _ds_profile_path = _profile_json_path(os.path.basename(master_path))
    if os.path.isfile(_ds_profile_path):
        try:
            with open(_ds_profile_path, "r", encoding="utf-8") as _pf:
                _ds_data = json.load(_pf)
            # Also load structure if exists
            _ds_structure_path = _structure_json_path(os.path.basename(master_path))
            if os.path.isfile(_ds_structure_path):
                with open(_ds_structure_path, "r", encoding="utf-8") as _sf:
                    _ds_struct = json.load(_sf)
                _all_ds_layouts = _ds_data.get("layouts", []) + _ds_struct.get("additional_layouts", [])
            else:
                _all_ds_layouts = _ds_data.get("layouts", [])
            for _l in _all_ds_layouts:
                _ph_map = {}
                for _p in _l.get("placeholders", []):
                    _ph_map[_p["idx"]] = {
                        "font": _p.get("font", {}),
                        "paragraph": _p.get("paragraph", {}),
                    }
                _ds_profile_map[_l["layout_index"]] = _ph_map
        except Exception:
            pass

    for slide_data in json_data.get("slides", []):
        layout_index = int(slide_data.get("layout_index", 0))

        # Guard against out-of-range indices
        if layout_index >= len(prs.slide_layouts):
            layout_index = 0
        layout = prs.slide_layouts[layout_index]

        slide = prs.slides.add_slide(layout)

        # ── Apply deep styles from saved .styles.json (if available) ──────────
        styles_path = _styles_json_path(os.path.basename(master_path))
        if os.path.isfile(styles_path):
            try:
                with open(styles_path, "r", encoding="utf-8") as _sf:
                    _styles = json.load(_sf)
                _layout_styles = next(
                    (l for l in _styles.get("layouts", [])
                     if l["layout_index"] == layout_index),
                    None,
                )
                if _layout_styles:
                    apply_styles_to_slide(slide, _layout_styles)
            except Exception:
                pass  # style application is best-effort

        # ── Title shortcut ─────────────────────────────────────────────────────
        # Apply font overrides (same as placeholder loop) so the title also
        # inherits the deep-scan profile's font name / size / bold.
        title_text = slide_data.get("title")
        if title_text and slide.shapes.title is not None:
            _t_ph_styles = _ds_profile_map.get(layout_index, {})
            _t_font = _t_ph_styles.get(0, {}).get("font", {})
            _t_para = _t_ph_styles.get(0, {}).get("paragraph", {})
            tf_title = slide.shapes.title.text_frame
            tf_title.clear()
            para_title = tf_title.paragraphs[0]
            _apply_para_format(para_title, _t_font, _t_para)
            run_title = para_title.add_run()
            run_title.text = str(title_text)
            _apply_font_to_run(run_title, _t_font)

        # ── Placeholders ───────────────────────────────────────────────────────
        for ph_data in slide_data.get("placeholders", []):
            ph_id      = int(ph_data.get("id", 0))
            content    = ph_data.get("content", "")
            ph_type    = ph_data.get("type", "text")   # "text" | "list"

            # Locate the placeholder by idx; skip gracefully if not found
            target = None
            try:
                target = slide.placeholders[ph_id]
            except KeyError:
                # Fall back: search by idx through all placeholders
                for ph in slide.placeholders:
                    if ph.placeholder_format.idx == ph_id:
                        target = ph
                        break

            if target is None:
                continue  # placeholder not present in this layout — skip

            tf = target.text_frame
            tf.clear()  # remove default factory text

            # Get font override from deep scan profile
            _ds_ph_styles = _ds_profile_map.get(layout_index, {})
            _ds_style = _ds_ph_styles.get(ph_id, {})
            _ds_font = _ds_style.get("font", {})
            _ds_para = _ds_style.get("paragraph", {})

            if ph_type == "list" and isinstance(content, list):
                for i, item in enumerate(content):
                    para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    _apply_para_format(para, _ds_font, _ds_para)
                    run = para.add_run()
                    run.text = str(item)
                    _apply_font_to_run(run, _ds_font)
            else:
                # Plain text (may contain newlines)
                lines = str(content).split("\n")
                for i, line in enumerate(lines):
                    para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    _apply_para_format(para, _ds_font, _ds_para)
                    run = para.add_run()
                    run.text = str(line)
                    _apply_font_to_run(run, _ds_font)

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ── /generate route ────────────────────────────────────────────────────────────
@app.route("/generate", methods=["POST"])
def generate():
    """
    Receive slide JSON and generate a PPTX for download.

    Accepted request bodies
    -----------------------
    Profile mode (no master PPTX needed):
        { "mode": "profile", "slides": [...], "presentation_name": "..." }

    Master mode (uploaded PPTX required):
        { "filename": "master.pptx", "slides": [...] }
    """
    body = request.get_json(silent=True)
    if not body:
        return jsonify({"error": "Request body must be JSON."}), 400

    if "slides" not in body or not isinstance(body["slides"], list) or len(body["slides"]) == 0:
        return jsonify({"error": "'slides' must be a non-empty array."}), 400

    json_data = {
        "presentation_name": body.get("presentation_name", "Presentation"),
        "slides": body["slides"],
    }

    mode = body.get("mode", "master")
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    # ── Profile mode: build from master_profile.json (no upload needed) ───────
    if mode == "profile" or not body.get("filename"):
        try:
            buf = create_pptx_from_profile(json_data)
        except Exception as exc:
            return jsonify({"error": f"Failed to generate PowerPoint: {exc}"}), 500
        out_name = f"presentation_{timestamp}.pptx"

    # ── Master mode: use an uploaded .pptx as template ─────────────────────────
    else:
        filename = secure_filename(body.get("filename", ""))
        if not filename:
            return jsonify({"error": "'filename' field is required in master mode."}), 400
        master_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        if not os.path.isfile(master_path):
            return jsonify({"error": f"Master file '{filename}' not found. Please re-upload."}), 404
        try:
            buf = create_pptx_from_json(json_data, master_path)
        except Exception as exc:
            return jsonify({"error": f"Failed to generate PowerPoint: {exc}"}), 500
        out_name = f"{os.path.splitext(filename)[0]}_generated_{timestamp}.pptx"

    return send_file(
        buf,
        as_attachment=True,
        download_name=out_name,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    app.run(debug=True, host="0.0.0.0", port=5000)
