#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
批量优化现有PPT图片：压缩原图 + 生成缩略图 + 更新manifest
"""
import json, os, sys
from pathlib import Path
from PIL import Image

SCRIPT_DIR = Path(__file__).resolve().parent
PPTS_DIR = SCRIPT_DIR / "ppts"
MANIFEST_FILE = PPTS_DIR / "manifest.json"

def load_manifest():
    if MANIFEST_FILE.exists():
        return json.loads(MANIFEST_FILE.read_text(encoding="utf-8"))
    return {}

def save_manifest(m):
    MANIFEST_FILE.write_text(json.dumps(m, ensure_ascii=False, indent=2), encoding="utf-8")

def optimize_all():
    manifest = load_manifest()
    total_files = 0
    total_original = 0
    total_optimized = 0

    for cat, stages in manifest.items():
        for stage, items in stages.items():
            for item in items:
                slides = item.get("slides", [])
                if not slides:
                    continue
                # 获取图片目录
                img_dir = SCRIPT_DIR / Path(slides[0]).parent
                if not img_dir.exists():
                    continue
                thumb_dir = img_dir / "_thumbs"
                thumb_dir.mkdir(exist_ok=True)
                # 清理旧缩略图
                for old in thumb_dir.glob("thumb_*.jpg"):
                    old.unlink()

                thumbs = []
                for slide_path in slides:
                    src = SCRIPT_DIR / slide_path
                    if not src.exists():
                        continue
                    total_files += 1
                    orig_size = src.stat().st_size
                    total_original += orig_size

                    # 压缩原图
                    img = Image.open(src)
                    img = img.convert("RGB")
                    img_resized = img.resize((1280, 720), Image.LANCZOS)
                    img_resized.save(src, "PNG", optimize=True)
                    opt_size = src.stat().st_size
                    total_optimized += opt_size

                    # 生成缩略图
                    basename = Path(slide_path).stem  # slide_0001
                    thumb = img.resize((120, 67), Image.LANCZOS)
                    thumb_name = basename.replace("slide_", "thumb_") + ".jpg"
                    thumb_path = thumb_dir / thumb_name
                    thumb.save(thumb_path, "JPEG", quality=60, optimize=True)

                    rel_thumb = str(thumb_path.relative_to(SCRIPT_DIR)).replace("\\", "/")
                    thumbs.append(rel_thumb)

                    print(f"  [{cat}/{stage}] {basename}: {orig_size/1024:.0f}KB→{opt_size/1024:.0f}KB, thumb={thumb_path.stat().st_size/1024:.0f}KB")

                # 更新 manifest
                item["thumbs"] = thumbs

    save_manifest(manifest)
    print(f"\n{'='*50}")
    print(f"处理完成: {total_files} 张图片")
    print(f"原始大小: {total_original/1024/1024:.1f} MB")
    print(f"优化后:   {total_optimized/1024/1024:.1f} MB")
    print(f"节省:     {(total_original-total_optimized)/1024/1024:.1f} MB ({(total_original-total_optimized)/total_original*100:.0f}%)")
    print(f"manifest.json 已更新 thumbs 字段")

if __name__ == "__main__":
    optimize_all()
