"""Scan all master slides in master_slide/ and write pre-built profiles to builtin_profiles/."""
import json
import os
from deep_scanner import deep_scan

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_SLIDE_DIR = os.path.join(BASE_DIR, "master_slide")
BUILTIN_DIR = os.path.join(BASE_DIR, "builtin_profiles")
os.makedirs(BUILTIN_DIR, exist_ok=True)

for fname in sorted(os.listdir(MASTER_SLIDE_DIR)):
    if not fname.lower().endswith(".pptx"):
        continue
    path = os.path.join(MASTER_SLIDE_DIR, fname)
    stem = os.path.splitext(fname)[0]
    print(f"Scanning {fname}...")
    try:
        ds = deep_scan(path)
    except Exception as exc:
        print(f"  ERROR: {exc}")
        continue

    profile = ds["master_profile"]
    structure = ds["master_structure"]

    with open(os.path.join(BUILTIN_DIR, f"{stem}.profile.json"), "w", encoding="utf-8") as f:
        json.dump(profile, f, ensure_ascii=False, indent=2)

    with open(os.path.join(BUILTIN_DIR, f"{stem}.structure.json"), "w", encoding="utf-8") as f:
        json.dump(structure, f, ensure_ascii=False, indent=2)

    # Merged layout list (for AI prompt / layout_schema endpoint)
    merged = profile.get("layouts", []) + structure.get("additional_layouts", [])
    merged.sort(key=lambda l: l["layout_index"])

    # Compact summary suitable for the frontend selector
    summary = {
        "id": stem,
        "name": profile.get("template_identity", stem),
        "pptx": fname,
        "canvas_size": profile.get("canvas_size"),
        "color_palette": profile.get("color_palette", {}),
        "theme_colors": profile.get("theme_colors", {}),
        "theme_fonts": profile.get("theme_fonts", {}),
        "total_layouts": len(merged),
        "layouts": merged,
    }
    with open(os.path.join(BUILTIN_DIR, f"{stem}.json"), "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print(f"  -> {len(merged)} layouts saved to builtin_profiles/{stem}.json")

print("Done.")
