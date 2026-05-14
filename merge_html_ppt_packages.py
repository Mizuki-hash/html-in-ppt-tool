from __future__ import annotations

import argparse
import contextlib
import re
import shutil
import time
import uuid
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from html_to_ppt_packager import (
    MANIFEST_NAME,
    OLD_PPT_NAME,
    OUTPUT_DIR,
    TEMPLATE_DIR,
    WEB_DIR_NAME,
    copytree_clean,
    package_zip,
    read_text,
    safe_name,
    update_manifest,
    update_start_script,
    write_setup_script,
    write_text,
)


COLLECTION_DIR = Path(__file__).resolve().parent / "collections"
PORT = 8765
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SLIDE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
WEBEXT_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/webextension"

ET.register_namespace("", REL_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)


@dataclass
class SourcePackage:
    root: Path
    ppt: Path
    web: Path
    title: str
    slug: str


def log(message: str) -> None:
    print(message, flush=True)


def resolve_source(path: Path) -> SourcePackage:
    path = path.expanduser().resolve()
    if path.is_file() and path.suffix.lower() == ".pptx":
        root = path.parent
        ppt = path
    elif path.is_dir():
        root = path
        candidates = sorted(
            [p for p in root.glob("*.pptx") if p.name != OLD_PPT_NAME],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not candidates:
            candidates = sorted(root.glob("*.pptx"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not candidates:
            raise RuntimeError(f"No PPTX found in: {root}")
        ppt = candidates[0]
    else:
        raise RuntimeError(f"Path is not a generated folder or PPTX: {path}")

    web = root / WEB_DIR_NAME
    if not (web / "index.html").exists():
        raise RuntimeError(f"Cannot find {WEB_DIR_NAME}\\index.html beside: {ppt}")
    title = ppt.stem.replace("_interactive_in_ppt", "")
    return SourcePackage(root=root, ppt=ppt, web=web, title=title, slug=safe_name(title))


def copy_web_payload(src_web: Path, dst_page: Path) -> None:
    if dst_page.exists():
        shutil.rmtree(dst_page)
    dst_page.mkdir(parents=True)
    excluded_names = {
        "server_https.py",
        "start_https_server.bat",
        "setup_powerpoint_addin.ps1",
        "localhost_selfsigned.crt",
        "localhost_selfsigned.key",
        "soi_dynamic_webpage.png",
        MANIFEST_NAME,
    }
    excluded_patterns = ("manifest_",)
    for item in src_web.iterdir():
        if item.name in excluded_names or item.name.startswith(excluded_patterns):
            continue
        target = dst_page / item.name
        if item.is_dir():
            shutil.copytree(item, target)
        elif item.is_file():
            shutil.copy2(item, target)


def rel_target_to_zip_path(slide_rels_name: str, target: str) -> str:
    base = Path(slide_rels_name).parent.parent
    resolved = (base / target).as_posix()
    while "/../" in resolved:
        resolved = re.sub(r"[^/]+/\.\./", "", resolved, count=1)
    return resolved


def read_first_slide_parts(ppt: Path) -> tuple[str, bytes, str]:
    with zipfile.ZipFile(ppt, "r") as z:
        slide_xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
        rels_name = "ppt/slides/_rels/slide1.xml.rels"
        rels_xml = z.read(rels_name)
        rels_root = ET.fromstring(rels_xml)
        image_target = None
        for rel in rels_root:
            if rel.attrib.get("Type") == IMAGE_REL_TYPE:
                image_target = rel.attrib.get("Target")
                break
        if not image_target:
            raise RuntimeError(f"No fallback image relationship found in {ppt}")
        image_path = rel_target_to_zip_path(rels_name, image_target)
        image_bytes = z.read(image_path)
    return slide_xml, image_bytes, image_path


def make_slide_rels(slide_index: int, page_url: str) -> bytes:
    root = ET.Element(f"{{{REL_NS}}}Relationships")
    ET.SubElement(root, f"{{{REL_NS}}}Relationship", {
        "Id": "rId1",
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        "Target": "../slideLayouts/slideLayout1.xml",
    })
    ET.SubElement(root, f"{{{REL_NS}}}Relationship", {
        "Id": "rId2",
        "Type": IMAGE_REL_TYPE,
        "Target": f"../media/html_fallback_{slide_index}.png",
    })
    ET.SubElement(root, f"{{{REL_NS}}}Relationship", {
        "Id": "rId3",
        "Type": HYPERLINK_REL_TYPE,
        "Target": page_url,
        "TargetMode": "External",
    })
    ET.SubElement(root, f"{{{REL_NS}}}Relationship", {
        "Id": "rId4",
        "Type": WEBEXT_REL_TYPE,
        "Target": f"../webextensions/webextension{slide_index}.xml",
    })
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def make_webextension_xml(addin_id: str, version: str) -> bytes:
    outer_id = "{" + str(uuid.uuid4()).upper() + "}"
    xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" id="{outer_id}">
  <we:reference id="{addin_id}" version="{version}" store="developer" storeType="Registry"/>
  <we:alternateReferences/>
  <we:properties/>
  <we:bindings/>
  <we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</we:webextension>'''
    return xml.encode("utf-8")


def rewrite_presentation_xml(data: bytes, slide_count: int) -> bytes:
    root = ET.fromstring(data)
    sld_list = root.find(f"{{{P_NS}}}sldIdLst")
    if sld_list is None:
        raise RuntimeError("presentation.xml has no sldIdLst")
    for child in list(sld_list):
        sld_list.remove(child)
    for i in range(1, slide_count + 1):
        item = ET.SubElement(sld_list, f"{{{P_NS}}}sldId")
        item.set("id", str(255 + i))
        item.set(f"{{{R_NS}}}id", f"rIdSlide{i}")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def rewrite_presentation_rels(data: bytes, slide_count: int) -> bytes:
    root = ET.fromstring(data)
    for rel in list(root):
        if rel.attrib.get("Type") == SLIDE_REL_TYPE:
            root.remove(rel)
    for i in range(1, slide_count + 1):
        ET.SubElement(root, f"{{{REL_NS}}}Relationship", {
            "Id": f"rIdSlide{i}",
            "Type": SLIDE_REL_TYPE,
            "Target": f"slides/slide{i}.xml",
        })
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def rewrite_content_types(data: bytes, slide_count: int) -> bytes:
    root = ET.fromstring(data)
    for item in list(root):
        if item.tag == f"{{{CT_NS}}}Override" and re.fullmatch(r"/ppt/slides/slide\d+\.xml", item.attrib.get("PartName", "")):
            root.remove(item)
    for i in range(1, slide_count + 1):
        ET.SubElement(root, f"{{{CT_NS}}}Override", {
            "PartName": f"/ppt/slides/slide{i}.xml",
            "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
        })
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def rewrite_zip(zip_path: Path, replacements: dict[str, bytes], removals: set[str]) -> None:
    tmp = zip_path.with_suffix(".tmp")
    with zipfile.ZipFile(zip_path, "r") as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        seen = set()
        for item in src.infolist():
            if item.filename in removals:
                continue
            if item.filename in replacements:
                dst.writestr(item.filename, replacements[item.filename])
                seen.add(item.filename)
            else:
                dst.writestr(item, src.read(item.filename))
        for name, data in replacements.items():
            if name not in seen:
                dst.writestr(name, data)
    tmp.replace(zip_path)


def write_collection_readme(package_dir: Path, sources: list[SourcePackage]) -> None:
    text = f"""关闭所有 PowerPoint 后，双击 START_HERE.bat。
这个 PPT 包含 {len(sources)} 个 HTML 交互页；演示时保持服务器窗口运行。
需要 Microsoft PowerPoint 桌面版 + Python 3；WPS 不支持交互。
"""
    write_text(package_dir / "README_使用说明.txt", text)


def merge_packages(paths: list[Path]) -> Path:
    if len(paths) < 2:
        raise RuntimeError("Please provide at least two generated folders or PPTX files.")
    sources = [resolve_source(path) for path in paths]

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    package_name = f"htmlppt_collection_{stamp}"
    package_dir = COLLECTION_DIR / package_name
    web_dir = package_dir / WEB_DIR_NAME
    pages_dir = web_dir / "pages"
    ppt_name = "merged_html_interactive_in_ppt.pptx"
    ppt_path = package_dir / ppt_name

    log(f"Creating collection: {package_dir}")
    COLLECTION_DIR.mkdir(exist_ok=True)
    copytree_clean(TEMPLATE_DIR, package_dir)

    old_ppt = package_dir / OLD_PPT_NAME
    if old_ppt.exists():
        old_ppt.rename(ppt_path)

    for old_name in ("pd_soi_fbe_simulator_v3.html", "pd_soi_fbe_ppt_content.html", "pd_soi_fbe_visual_config.json", MANIFEST_NAME):
        with contextlib.suppress(FileNotFoundError):
            (web_dir / old_name).unlink()
    with contextlib.suppress(FileNotFoundError):
        (package_dir / "01_setup_current_user.bat").unlink()

    replacements: dict[str, bytes] = {}
    removals: set[str] = set()
    with zipfile.ZipFile(ppt_path, "r") as z:
        replacements["ppt/presentation.xml"] = rewrite_presentation_xml(z.read("ppt/presentation.xml"), len(sources))
        replacements["ppt/_rels/presentation.xml.rels"] = rewrite_presentation_rels(z.read("ppt/_rels/presentation.xml.rels"), len(sources))
        replacements["[Content_Types].xml"] = rewrite_content_types(z.read("[Content_Types].xml"), len(sources))
        for item in z.infolist():
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", item.filename) and item.filename != "ppt/slides/slide1.xml":
                removals.add(item.filename)
            if re.fullmatch(r"ppt/slides/_rels/slide\d+\.xml\.rels", item.filename) and item.filename != "ppt/slides/_rels/slide1.xml.rels":
                removals.add(item.filename)
            if re.fullmatch(r"ppt/webextensions/webextension\d+\.xml", item.filename):
                removals.add(item.filename)

    manifest_lines = []
    for i, src in enumerate(sources, start=1):
        page_slug = f"{i:02d}_{src.slug}"
        page_dir = pages_dir / page_slug
        copy_web_payload(src.web, page_dir)

        addin_id = str(uuid.uuid4())
        version = f"1.0.{int(time.time()) % 9000 + i}.0"
        page_url = f"https://127.0.0.1:{PORT}/pages/{page_slug}/index.html?v={stamp}_{i}"

        manifest_name = f"manifest_{i:02d}_{src.slug}.xml"
        manifest_path = web_dir / manifest_name
        shutil.copy2(TEMPLATE_DIR / WEB_DIR_NAME / MANIFEST_NAME, manifest_path)
        update_manifest(manifest_path, addin_id, version, f"pages/{page_slug}/index.html?v={stamp}_{i}")
        # update_manifest always writes the standard index URL, so patch the path after it.
        text = read_text(manifest_path)
        text = re.sub(r'https://127\.0\.0\.1:8765/index\.html\?v=[^"]+', page_url, text, count=1)
        write_text(manifest_path, text)

        slide_xml, image_bytes, _ = read_first_slide_parts(src.ppt)
        replacements[f"ppt/slides/slide{i}.xml"] = slide_xml.encode("utf-8")
        replacements[f"ppt/slides/_rels/slide{i}.xml.rels"] = make_slide_rels(i, page_url)
        replacements[f"ppt/media/html_fallback_{i}.png"] = image_bytes
        replacements[f"ppt/webextensions/webextension{i}.xml"] = make_webextension_xml(addin_id, version)
        manifest_lines.append((addin_id, manifest_path))

    rewrite_zip(ppt_path, replacements, removals)
    write_setup_script(web_dir)
    update_start_script(package_dir / "02_start_server_and_open_ppt.bat", ppt_name)
    write_collection_readme(package_dir, sources)
    return package_dir


def main() -> int:
    parser = argparse.ArgumentParser(description="Merge generated HTML-in-PPT folders into one multi-slide PPT folder.")
    parser.add_argument("paths", nargs="+", help="Generated package folders or generated .pptx files")
    args = parser.parse_args()
    try:
        package_dir = merge_packages([Path(p) for p in args.paths])
    except Exception as exc:
        log("")
        log(f"ERROR: {exc}")
        return 1
    log("")
    log("Done.")
    log(f"Collection folder: {package_dir}")
    log("Open the folder and run START_HERE.bat.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
