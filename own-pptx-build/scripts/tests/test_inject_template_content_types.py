"""inject_template.py: 注入後の [Content_Types].xml が、ppt/media/ にある全拡張子の Default を持つこと。

生成側の pptx に .jpg の画像があると、注入後に `Default Extension="jpg"` が無いまま残り、
python-pptx が開けない・PowerPoint が修復を要求する（2026-10-19 豊橋デッキで実測）。
"""

import re
import subprocess
import sys
import zipfile
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.util import Inches

SKILL = Path(__file__).resolve().parents[2]
SCRIPT = SKILL / "scripts" / "inject_template.py"
TEMPLATE = SKILL / "style-guide" / "template_v3.pptx"


def make_gen(path: Path, image_name: str, fmt: str) -> Path:
    img = path.parent / image_name
    Image.new("RGB", (20, 10), (200, 100, 50)).save(img, fmt)
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_picture(str(img), Inches(1), Inches(1), Inches(2), Inches(1))
    prs.save(path)
    return path


def inject(gen: Path, out: Path) -> None:
    r = subprocess.run(
        [sys.executable, str(SCRIPT), str(gen), str(TEMPLATE), str(out)],
        capture_output=True,
        text=True,
    )
    assert r.returncode == 0, r.stderr + r.stdout


def default_exts(pptx: Path) -> set[str]:
    xml = zipfile.ZipFile(pptx).read("[Content_Types].xml").decode("utf-8")
    return set(re.findall(r'<Default Extension="([^"]+)"', xml))


def media_exts(pptx: Path) -> set[str]:
    return {
        n.rsplit(".", 1)[1].lower()
        for n in zipfile.ZipFile(pptx).namelist()
        if n.startswith("ppt/media/") and "." in n
    }


def test_jpg_media_gets_default_content_type(tmp_path):
    gen = make_gen(tmp_path / "gen.pptx", "photo.jpg", "JPEG")
    out = tmp_path / "out.pptx"
    inject(gen, out)
    assert "jpg" in media_exts(out), "前提: 注入後に .jpg の media が残っている"
    missing = media_exts(out) - {e.lower() for e in default_exts(out)}
    assert not missing, f"Default content type が無い media 拡張子: {missing}"


def test_png_media_still_has_default(tmp_path):
    gen = make_gen(tmp_path / "gen.pptx", "pic.png", "PNG")
    out = tmp_path / "out.pptx"
    inject(gen, out)
    missing = media_exts(out) - {e.lower() for e in default_exts(out)}
    assert not missing, missing
