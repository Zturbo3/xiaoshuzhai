#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
小书斋编程学习平台 - 课件管理工具 (静态部署版)
用法:
  python ppt_manager.py add         <ppt文件路径> <分类> <阶段> [显示名称]
  python ppt_manager.py batch_add   <文件夹路径> <分类> <阶段> [显示名称前缀]
  python ppt_manager.py delete      <分类> <阶段> <序号或名称>
  python ppt_manager.py list        [分类]
  python ppt_manager.py user       add    <用户名> <密码> [--admin]
  python ppt_manager.py user       delete <用户名>
  python ppt_manager.py user       list
  python ppt_manager.py push        [提交说明]
  python ppt_manager.py push_data   [提交说明]   # 仅推送图片到数据仓库
  python ppt_manager.py init_data                # 首次初始化数据仓库（同步所有现
有图片）
  python ppt_manager.py menu                 # 交互式菜单

分类: scratch | python | cpp
阶段: scratch -> A | B
      python  -> 初级 | 中级 | 高级
      cpp     -> 基础 | 进阶 | 竞赛算法

示例:
  python ppt_manager.py add "D:\\课件\\第6讲 链表.pptx" scratch A
  python ppt_manager.py add "D:\\课件\\变量与类型.pptx" python 初级 "第1讲 变量"
  python ppt_manager.py batch_add "D:\\课件\\scratch_A" scratch A "第"
  python ppt_manager.py delete scratch A 3
  python ppt_manager.py user add 张三 123456
  python ppt_manager.py push "添加第6讲课件"
"""

import sys
import os
import json
import shutil
import subprocess
import base64
from pathlib import Path
from datetime import datetime, timezone
from email.utils import formatdate

from PIL import Image as PILImage

# ── 路径配置 ──
SCRIPT_DIR = Path(__file__).resolve().parent
PPTS_DIR = SCRIPT_DIR / "ppts"
MANIFEST_FILE = PPTS_DIR / "manifest.json"
ACCOUNTS_FILE = SCRIPT_DIR / "data" / "accounts.json"
DATA_REPO_DIR = SCRIPT_DIR / ".data-repo"  # 图片数据仓库（工作区内）

# ── 分类与阶段映射 ──
STAGES = {
    "scratch": ["A", "B"],
    "python": ["初级", "中级", "高级"],
    "cpp": ["基础", "进阶", "竞赛算法"],
}

def load_manifest():
    if MANIFEST_FILE.exists():
        return json.loads(MANIFEST_FILE.read_text(encoding="utf-8"))
    return {cat: {s: [] for s in stages} for cat, stages in STAGES.items()}

def save_manifest(manifest):
    PPTS_DIR.mkdir(parents=True, exist_ok=True)
    MANIFEST_FILE.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

def load_accounts():
    if ACCOUNTS_FILE.exists():
        return json.loads(ACCOUNTS_FILE.read_text(encoding="utf-8"))
    return []

def save_accounts(accounts):
    ACCOUNTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    ACCOUNTS_FILE.write_text(
        json.dumps(accounts, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

def validate_cat_stage(cat, stage):
    if cat not in STAGES:
        print(f"错误: 未知分类 '{cat}'，可选: {', '.join(STAGES.keys())}")
        sys.exit(1)
    if stage not in STAGES[cat]:
        print(f"错误: 分类 '{cat}' 的阶段 '{stage}' 无效，可选: {', '.join(STAGES[cat])}")
        sys.exit(1)

# ── PPT 转 PNG (PowerPoint COM) ──
def convert_ppt_to_images(ppt_path, output_dir):
    """使用 PowerPoint COM 将 PPT 转为 PNG 图片"""
    ppt_path = Path(ppt_path).resolve()
    if not ppt_path.exists():
        print(f"错误: 文件不存在: {ppt_path}")
        sys.exit(1)

    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    # 清空旧图片
    for old in output_dir.glob("slide_*.png"):
        old.unlink()

    print(f"正在转换: {ppt_path.name}")
    print(f"输出目录: {output_dir}")

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("错误: 需要安装 pywin32")
        print("运行: pip install pywin32")
        sys.exit(1)

    pythoncom.CoInitialize()
    ppt_app = None
    try:
        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        # PowerPoint 需要可见才能可靠工作，但设为最小化
        try:
            ppt_app.Visible = True
            ppt_app.WindowState = 2  # ppWindowMinimized
        except Exception:
            pass

        pres = ppt_app.Presentations.Open(
            str(ppt_path),
            ReadOnly=True,
            WithWindow=False
        )

        total = pres.Slides.Count
        print(f"共 {total} 页幻灯片，正在导出...")

        for i, slide in enumerate(pres.Slides, 1):
            tmp_path = output_dir / f"slide_{i:04d}.png"
            slide.Export(str(tmp_path), "PNG", 1920, 1080)
            print(f"  [{i}/{total}] 导出完成", end="\r")

        pres.Close()
        print(f"\n转换完成! 共导出 {total} 张图片")

        # ── 压缩原图 + 生成缩略图 ──
        print("正在压缩图片并生成缩略图...")
        thumb_dir = output_dir / "_thumbs"
        thumb_dir.mkdir(exist_ok=True)
        # 清空旧缩略图
        for old in thumb_dir.glob("thumb_*.jpg"):
            old.unlink()

        for i in range(1, total + 1):
            src = output_dir / f"slide_{i:04d}.png"
            # 压缩原图: 1280x720, 质量85%
            img = PILImage.open(src)
            img = img.convert("RGB")
            img_resized = img.resize((1280, 720), PILImage.LANCZOS)
            img_resized.save(src, "PNG", optimize=True)
            # 生成缩略图: 120x67, 质量60%, JPEG
            thumb = img.resize((120, 67), PILImage.LANCZOS)
            thumb_path = thumb_dir / f"thumb_{i:04d}.jpg"
            thumb.save(thumb_path, "JPEG", quality=60, optimize=True)
            print(f"  [{i}/{total}] 压缩完成", end="\r")

        print(f"\n压缩完成! 原图: 1280x720 | 缩略图: 120x67")

        # 返回相对路径列表
        rel_dir = output_dir.relative_to(SCRIPT_DIR)
        thumb_rel_dir = thumb_dir.relative_to(SCRIPT_DIR)
        images = sorted(output_dir.glob("slide_*.png"))
        thumbs = sorted(thumb_dir.glob("thumb_*.jpg"))
        return (
            [str(rel_dir / img.name).replace("\\", "/") for img in images],
            [str(thumb_rel_dir / t.name).replace("\\", "/") for t in thumbs],
            total
        )

    except Exception as e:
        print(f"\n转换失败: {e}")
        sys.exit(1)
    finally:
        if ppt_app:
            try:
                ppt_app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

# ── 添加课件 ──
def cmd_add(args):
    if len(args) < 3:
        print("用法: python ppt_manager.py add <ppt文件路径> <分类> <阶段> [显示名称]")
        print("示例: python ppt_manager.py add \"D:\\第6讲.pptx\" scratch A")
        sys.exit(1)

    ppt_path = args[0]
    cat = args[1]
    stage = args[2]
    display_name = args[3] if len(args) > 3 else Path(ppt_path).stem

    validate_cat_stage(cat, stage)

    # 目标目录
    dest_dir = PPTS_DIR / cat / stage
    dest_dir.mkdir(parents=True, exist_ok=True)

    # 复制 PPT 文件到项目目录
    ppt_filename = Path(ppt_path).name
    dest_ppt = dest_dir / ppt_filename
    if Path(ppt_path).resolve() != dest_ppt.resolve():
        shutil.copy2(ppt_path, dest_ppt)
        print(f"已复制: {dest_ppt}")

    # 转换图片
    stem = Path(ppt_filename).stem
    img_dir = dest_dir / f"{stem}_images"
    slides, thumbs, slide_count = convert_ppt_to_images(dest_ppt, img_dir)

    # 更新 manifest
    manifest = load_manifest()
    rel_ppt = f"ppts/{cat}/{stage}/{ppt_filename}"
    rel_img_dir = f"ppts/{cat}/{stage}/{stem}_images"

    # 检查是否已存在同名课件
    existing = manifest[cat][stage]
    for i, item in enumerate(existing):
        if item.get("name") == display_name:
            # 更新已有项
            existing[i] = {
                "name": display_name,
                "file": rel_ppt,
                "size": dest_ppt.stat().st_size,
                "uploadedAt": formatdate(timeval=None, localtime=False),
                "slideCount": slide_count,
                "slides": slides,
                "thumbs": thumbs
            }
            print(f"已更新课件: {display_name}")
            break
    else:
        # 添加新项
        existing.append({
            "name": display_name,
            "file": rel_ppt,
            "size": dest_ppt.stat().st_size,
            "uploadedAt": formatdate(timeval=None, localtime=False),
            "slideCount": slide_count,
            "slides": slides,
            "thumbs": thumbs
        })
        print(f"已添加课件: {display_name}")

    save_manifest(manifest)
    print(f"manifest.json 已更新")
    print(f"\n下一步: 运行 python ppt_manager.py push 上传到 GitHub")

# ── 批量添加课件 ──
def cmd_batch_add(args):
    if len(args) < 3:
        print("用法: python ppt_manager.py batch_add <文件夹路径> <分类> <阶段> [显示名称前缀]")
        print("示例: python ppt_manager.py batch_add \"D:\\课件\\scratch_A\" scratch A \"第\"")
        print("说明: 将批量处理文件夹内所有 .pptx 和 .ppt 文件")
        sys.exit(1)

    folder_path = Path(args[0])
    cat = args[1]
    stage = args[2]
    name_prefix = args[3] if len(args) > 3 else ""

    validate_cat_stage(cat, stage)

    if not folder_path.exists():
        print(f"错误: 文件夹不存在: {folder_path}")
        sys.exit(1)

    # 查找所有PPT文件
    ppt_files = sorted(folder_path.glob("*.pptx")) + sorted(folder_path.glob("*.ppt"))

    if not ppt_files:
        print(f"错误: 文件夹内没有找到 .pptx 或 .ppt 文件")
        sys.exit(1)

    print(f"\n找到 {len(ppt_files)} 个PPT文件:")
    for i, f in enumerate(ppt_files, 1):
        print(f"  {i}. {f.name}")
    print()

    # 确认
    confirm = input("确认批量添加这些课件? (y/n): ").strip().lower()
    if confirm != "y":
        print("已取消")
        sys.exit(0)

    # 批量处理
    success_count = 0
    fail_count = 0

    for i, ppt_file in enumerate(ppt_files, 1):
        print(f"\n{'='*60}")
        print(f"[{i}/{len(ppt_files)}] 处理: {ppt_file.name}")
        print(f"{'='*60}")

        # 生成显示名称
        display_name = f"{name_prefix}{Path(ppt_file).stem}" if name_prefix else Path(ppt_file).stem

        try:
            # 调用cmd_add的逻辑
            ppt_path = str(ppt_file)

            # 目标目录
            dest_dir = PPTS_DIR / cat / stage
            dest_dir.mkdir(parents=True, exist_ok=True)

            # 复制PPT文件
            ppt_filename = Path(ppt_path).name
            dest_ppt = dest_dir / ppt_filename
            if Path(ppt_path).resolve() != dest_ppt.resolve():
                shutil.copy2(ppt_path, dest_ppt)
                print(f"已复制: {dest_ppt}")

            # 转换图片
            stem = Path(ppt_filename).stem
            img_dir = dest_dir / f"{stem}_images"
            slides, thumbs, slide_count = convert_ppt_to_images(dest_ppt, img_dir)

            # 更新manifest
            manifest = load_manifest()
            rel_ppt = f"ppts/{cat}/{stage}/{ppt_filename}"
            rel_img_dir = f"ppts/{cat}/{stage}/{stem}_images"

            # 检查是否已存在同名课件
            existing = manifest[cat][stage]
            for j, item in enumerate(existing):
                if item.get("name") == display_name:
                    # 更新已有项
                    existing[j] = {
                        "name": display_name,
                        "file": rel_ppt,
                        "size": dest_ppt.stat().st_size,
                        "uploadedAt": formatdate(timeval=None, localtime=False),
                        "slideCount": slide_count,
                        "slides": slides,
                        "thumbs": thumbs
                    }
                    print(f"已更新课件: {display_name}")
                    break
            else:
                # 添加新项
                existing.append({
                    "name": display_name,
                    "file": rel_ppt,
                    "size": dest_ppt.stat().st_size,
                    "uploadedAt": formatdate(timeval=None, localtime=False),
                    "slideCount": slide_count,
                    "slides": slides,
                    "thumbs": thumbs
                })
                print(f"已添加课件: {display_name}")

            save_manifest(manifest)
            success_count += 1

        except Exception as e:
            print(f"处理失败: {e}")
            fail_count += 1

    print(f"\n{'='*60}")
    print(f"批量添加完成!")
    print(f"  成功: {success_count} 个")
    print(f"  失败: {fail_count} 个")
    print(f"{'='*60}")
    print(f"\n下一步: 运行 python ppt_manager.py push 上传到 GitHub")

# ── 删除课件 ──
def cmd_delete(args):
    if len(args) < 3:
        print("用法: python ppt_manager.py delete <分类> <阶段> <序号或名称>")
        sys.exit(1)

    cat = args[0]
    stage = args[1]
    target = args[2]

    validate_cat_stage(cat, stage)

    manifest = load_manifest()
    items = manifest[cat][stage]

    # 查找目标
    idx = None
    try:
        idx = int(target) - 1  # 1-based to 0-based
        if idx < 0 or idx >= len(items):
            print(f"错误: 序号超出范围 (1-{len(items)})")
            sys.exit(1)
    except ValueError:
        # 按名称查找
        for i, item in enumerate(items):
            if item.get("name") == target:
                idx = i
                break
        if idx is None:
            print(f"错误: 未找到名为 '{target}' 的课件")
            sys.exit(1)

    item = items[idx]
    print(f"将删除: {item['name']} ({item.get('slideCount', '?')} 页)")

    # 删除图片目录
    img_dir = SCRIPT_DIR / item.get("slides", [""])[0] if item.get("slides") else None
    if img_dir and img_dir.parent.exists():
        # slides[0] = "ppts/scratch/A/xxx_images/slide_0001.png"
        # img_dir.parent = "ppts/scratch/A/xxx_images"
        if img_dir.parent.name.endswith("_images"):
            shutil.rmtree(img_dir.parent, ignore_errors=True)
            print(f"已删除图片目录: {img_dir.parent.name}")

    # 删除 PPT 文件
    ppt_file = SCRIPT_DIR / item.get("file", "")
    if ppt_file.exists():
        ppt_file.unlink()
        print(f"已删除: {ppt_file.name}")

    # 从 manifest 移除
    items.pop(idx)
    save_manifest(manifest)
    print(f"manifest.json 已更新")
    print(f"\n下一步: 运行 python ppt_manager.py push 同步到 GitHub")

# ── 列出课件 ──
def cmd_list(args):
    cat = args[0] if args else None
    manifest = load_manifest()

    cats = [cat] if cat else STAGES.keys()
    for c in cats:
        if c not in manifest:
            continue
        print(f"\n{'='*50}")
        print(f"  {c.upper()}")
        print(f"{'='*50}")
        for stage, items in manifest[c].items():
            print(f"\n  [{stage}] ({len(items)} 个课件)")
            for i, item in enumerate(items, 1):
                count = item.get("slideCount", len(item.get("slides", [])))
                print(f"    {i}. {item['name']}  ({count} 页)")

# ── 用户管理 ──
def cmd_user(args):
    if not args:
        print("用法: python ppt_manager.py user <add|delete|list> ...")
        sys.exit(1)

    action = args[0]
    accounts = load_accounts()

    if action == "add":
        if len(args) < 3:
            print("用法: python ppt_manager.py user add <用户名> <密码> [--admin]")
            sys.exit(1)
        username = args[1]
        password = args[2]
        is_admin = "--admin" in args or len(args) > 3 and args[3] == "--admin"

        # 检查是否已存在
        if any(a["user"] == username for a in accounts):
            print(f"错误: 用户 '{username}' 已存在")
            sys.exit(1)

        encoded_pass = base64.b64encode(password.encode()).decode()
        # 第一个用户自动管理员
        if len(accounts) == 0:
            is_admin = True

        accounts.append({
            "user": username,
            "pass": encoded_pass,
            "admin": is_admin
        })
        save_accounts(accounts)
        role = "管理员" if is_admin else "普通用户"
        print(f"已添加{role}: {username}")
        print(f"\n下一步: 运行 python ppt_manager.py push 同步到 GitHub")

    elif action == "delete":
        if len(args) < 2:
            print("用法: python ppt_manager.py user delete <用户名>")
            sys.exit(1)
        username = args[1]
        original_len = len(accounts)
        accounts = [a for a in accounts if a["user"] != username]
        if len(accounts) == original_len:
            print(f"错误: 未找到用户 '{username}'")
            sys.exit(1)
        save_accounts(accounts)
        print(f"已删除用户: {username}")
        print(f"\n下一步: 运行 python ppt_manager.py push 同步到 GitHub")

    elif action == "list":
        if not accounts:
            print("暂无用户")
        else:
            print(f"\n{'='*40}")
            print(f"  用户列表 ({len(accounts)} 人)")
            print(f"{'='*40}")
            for a in accounts:
                role = "管理员" if a.get("admin") else "普通用户"
                print(f"  {a['user']:12s}  [{role}]")
    else:
        print(f"未知操作: {action}")
        print("可用操作: add, delete, list")
        sys.exit(1)

# ── 推送到 GitHub ──
def cmd_push(args):
    msg = " ".join(args) if args else f"更新于 {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    print("=" * 50)
    print("  推送到 GitHub")
    print("=" * 50)

    main_ok = False
    data_ok = False

    # 1. 推送主仓库 (代码 + manifest + accounts)
    main_ok = push_main_repo(msg)

    # 2. 同步并推送数据仓库 (PPT 图片)
    if DATA_REPO_DIR.exists() and (DATA_REPO_DIR / ".git").exists():
        data_ok = push_data_repo(msg)
    else:
        print(f"\n[!] 数据仓库不存在: {DATA_REPO_DIR}")
        print("   请先运行: git clone https://github.com/Zturbo3/xiaoshuzhai-data.git")

    print(f"\n{'='*50}")
    if main_ok and data_ok:
        print("  推送成功! GitHub Pages 将在 1-2 分钟内更新。")
    elif main_ok and not data_ok:
        print("  [!] 主仓库推送成功，但数据仓库推送失败!")
        print("  课件列表已更新，但部分PPT图片可能无法显示。")
        print("  请重新运行 push 命令重试数据仓库推送。")
    elif not main_ok and data_ok:
        print("  [!] 数据仓库推送成功，但主仓库推送失败!")
        print("  PPT图片已上传，但课件列表未更新，网页看不到新课件。")
        print("  请重新运行 push 命令重试主仓库推送。")
    else:
        print("  [!] 推送失败! 网络可能不稳定，请稍后重试。")
        print("  手动重试命令:")
        print(f"    cd \"{SCRIPT_DIR}\" && git push origin main")
        print(f"    cd \"{DATA_REPO_DIR}\" && git push origin main")
    print(f"  网站: https://zturbo3.github.io/xiaoshuzhai/")
    print(f"{'='*50}")


def push_main_repo(msg):
    """推送代码仓库（不含图片）"""
    print("\n[1/2] 推送主仓库 (代码+配置)...")

    result = subprocess.run(
        ["git", "add", "-A"],
        cwd=SCRIPT_DIR,
        capture_output=True,
        encoding="utf-8",
        errors="replace"
    )
    if result.returncode != 0:
        print(f"  git add 失败: {result.stderr}")
        return False

    result = subprocess.run(
        ["git", "diff", "--cached", "--stat"],
        cwd=SCRIPT_DIR,
        capture_output=True,
        encoding="utf-8",
        errors="replace"
    )
    if not result.stdout.strip():
        print("  (主仓库没有需要提交的变更，直接推送远程同步...)")
        try:
            push_result = subprocess.run(["git", "push", "origin", "main"], cwd=SCRIPT_DIR,
                           capture_output=True, encoding="utf-8", errors="replace", timeout=120)
            if push_result.returncode == 0:
                print("  主仓库推送成功! (main 分支)")
                return True
            else:
                print(f"  git push 失败: {push_result.stderr.strip()}")
                return False
        except subprocess.TimeoutExpired:
            print("  git push 超时 (120秒)，网络可能不稳定")
            return False

    print(result.stdout)

    result = subprocess.run(
        ["git", "commit", "-m", msg],
        cwd=SCRIPT_DIR,
        capture_output=True,
        encoding="utf-8",
        errors="replace"
    )
    if result.returncode != 0:
        print(f"  git commit 失败: {result.stderr}")
        return False

    # 推送 (尝试 main 再尝试 master)
    for branch in ["main", "master"]:
        try:
            result = subprocess.run(
                ["git", "push", "origin", branch],
                cwd=SCRIPT_DIR,
                capture_output=True,
                encoding="utf-8",
                errors="replace",
                timeout=120
            )
            if result.returncode == 0:
                print(f"  主仓库推送成功! ({branch} 分支)")
                return True
            else:
                print(f"  git push ({branch}) 失败: {result.stderr.strip()}")
        except subprocess.TimeoutExpired:
            print(f"  git push ({branch}) 超时 (120秒)，网络可能不稳定")
        except Exception as e:
            print(f"  git push ({branch}) 异常: {e}")

    print("  [!] 主仓库推送失败! 课件列表未更新到网站。")
    return False


def push_data_repo(msg):
    """同步图片到数据仓库并推送"""
    print("\n[2/2] 推送数据仓库 (PPT图片)...")

    # 用 robocopy 同步图片目录（仅复制新增/变更的文件）
    data_ppts = DATA_REPO_DIR / "ppts"
    data_ppts.mkdir(parents=True, exist_ok=True)

    # 同步所有 *_images 目录
    image_dirs = list(PPTS_DIR.glob("*/*/*_images"))
    synced = 0
    for img_dir in image_dirs:
        # 计算相对路径: ppts/scratch/A/xxx_images
        rel_dir = img_dir.relative_to(PPTS_DIR)
        dest = data_ppts / rel_dir
        dest.parent.mkdir(parents=True, exist_ok=True)

        # 使用 robocopy 同步 (Windows)
        result = subprocess.run(
            ["robocopy", str(img_dir), str(dest), "/MIR", "/NJH", "/NJS", "/NP", "/NDL", "/NC", "/NS"],
            capture_output=True,
            encoding="utf-8",
            errors="replace"
        )
        # robocopy 返回码: 0-7 都是成功
        if result.returncode < 8:
            synced += 1

    print(f"  已同步 {synced}/{len(image_dirs)} 个图片目录")

    # 复制 manifest.json 到数据仓库（作为备份参考）
    manifest_dest = data_ppts / "manifest.json"
    shutil.copy2(MANIFEST_FILE, manifest_dest)

    # git add & commit & push
    result = subprocess.run(
        ["git", "add", "-A"],
        cwd=DATA_REPO_DIR,
        capture_output=True,
        encoding="utf-8",
        errors="replace"
    )
    if result.returncode != 0:
        print(f"  数据仓库 git add 失败: {result.stderr}")
        return False

    result = subprocess.run(
        ["git", "diff", "--cached", "--stat"],
        cwd=DATA_REPO_DIR,
        capture_output=True,
        encoding="utf-8",
        errors="replace"
    )
    if not result.stdout.strip():
        print("  (数据仓库没有需要提交的变更，直接推送远程同步...)")
        try:
            push_result = subprocess.run(["git", "push", "origin", "main"], cwd=DATA_REPO_DIR,
                           capture_output=True, encoding="utf-8", errors="replace", timeout=120)
            if push_result.returncode == 0:
                print("  数据仓库推送成功! (main 分支)")
                return True
            else:
                print(f"  数据仓库 git push 失败: {push_result.stderr.strip()}")
                return False
        except subprocess.TimeoutExpired:
            print("  数据仓库 git push 超时 (120秒)")
            return False

    print(result.stdout)

    result = subprocess.run(
        ["git", "commit", "-m", msg],
        cwd=DATA_REPO_DIR,
        capture_output=True,
        encoding="utf-8",
        errors="replace"
    )
    if result.returncode != 0:
        print(f"  数据仓库 git commit 失败: {result.stderr}")
        return False

    for branch in ["main", "master"]:
        try:
            result = subprocess.run(
                ["git", "push", "origin", branch],
                cwd=DATA_REPO_DIR,
                capture_output=True,
                encoding="utf-8",
                errors="replace",
                timeout=300  # 图片可能很大
            )
            if result.returncode == 0:
                print(f"  数据仓库推送成功! ({branch} 分支)")
                return True
            else:
                print(f"  数据仓库 git push ({branch}) 失败: {result.stderr.strip()}")
        except subprocess.TimeoutExpired:
            print(f"  数据仓库 git push ({branch}) 超时 (300秒)")
        except Exception as e:
            print(f"  数据仓库 git push ({branch}) 异常: {e}")

    print("  [!] 数据仓库推送失败! PPT图片未上传。")
    return False

# ── 初始化数据仓库 ──
def cmd_init_data(args):
    """首次初始化：同步所有现有图片到数据仓库"""
    if not DATA_REPO_DIR.exists():
        print(f"错误: 数据仓库目录不存在: {DATA_REPO_DIR}")
        print(f"请先运行: git clone https://github.com/Zturbo3/xiaoshuzhai-data.git")
        print(f"克隆到: {DATA_REPO_DIR.parent}")
        sys.exit(1)

    if not (DATA_REPO_DIR / ".git").exists():
        print(f"错误: 不是有效的 git 仓库: {DATA_REPO_DIR}")
        sys.exit(1)

    print("=" * 50)
    print("  初始化数据仓库 (首次同步)")
    print("=" * 50)

    # 同步所有图片
    data_ppts = DATA_REPO_DIR / "ppts"
    data_ppts.mkdir(parents=True, exist_ok=True)

    image_dirs = list(PPTS_DIR.glob("*/*/*_images"))
    print(f"共找到 {len(image_dirs)} 个图片目录")

    for i, img_dir in enumerate(image_dirs, 1):
        rel_dir = img_dir.relative_to(PPTS_DIR)
        dest = data_ppts / rel_dir
        dest.parent.mkdir(parents=True, exist_ok=True)

        result = subprocess.run(
            ["robocopy", str(img_dir), str(dest), "/MIR", "/NJH", "/NJS", "/NP", "/NDL", "/NC", "/NS"],
            capture_output=True,
            encoding="utf-8",
            errors="replace"
        )
        print(f"  [{i}/{len(image_dirs)}] {rel_dir}")

    # 复制 manifest.json
    manifest_dest = data_ppts / "manifest.json"
    shutil.copy2(MANIFEST_FILE, manifest_dest)
    print(f"\n已同步所有图片到数据仓库")

    # 提交并推送
    result = subprocess.run(
        ["git", "add", "-A"],
        cwd=DATA_REPO_DIR,
        capture_output=True,
        encoding="utf-8",
        errors="replace"
    )
    if result.returncode == 0:
        result = subprocess.run(
            ["git", "commit", "-m", "初始化：同步所有PPT图片"],
            cwd=DATA_REPO_DIR,
            capture_output=True,
            encoding="utf-8",
            errors="replace"
        )
        print("已提交到数据仓库")
    else:
        print(f"git add 失败: {result.stderr}")

    print("\n准备推送到 GitHub...")
    confirm = input("确认推送? (y/n): ").strip().lower()
    if confirm != "y":
        print("已取消。稍后可运行: python ppt_manager.py push_data")
        return

    result = subprocess.run(
        ["git", "push", "origin", "main"],
        cwd=DATA_REPO_DIR,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        timeout=300
    )
    if result.returncode == 0:
        print("数据仓库推送成功!")
        print(f"CDN地址: https://cdn.jsdelivr.net/gh/Zturbo3/xiaoshuzhai-data@main/ppts/")
    else:
        print(f"推送失败: {result.stderr}")


# ── 仅推送数据仓库 ──
def cmd_push_data(args):
    """仅推送图片到数据仓库（不推送主仓库）"""
    msg = " ".join(args) if args else f"更新图片于 {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    if not DATA_REPO_DIR.exists() or not (DATA_REPO_DIR / ".git").exists():
        print(f"错误: 数据仓库不存在: {DATA_REPO_DIR}")
        print(f"请先运行: python ppt_manager.py init_data")
        sys.exit(1)

    push_data_repo(msg)


# ── 交互式菜单 ──
def cmd_menu():
    while True:
        print("\n" + "=" * 50)
        print("    小书斋编程学习平台 - 课件管理工具")
        print("=" * 50)
        print("  1. 添加课件 (PPT转图片并上传)")
        print("  2. 批量添加课件 (文件夹批量处理)")
        print("  3. 删除课件")
        print("  4. 查看所有课件")
        print("  5. 添加用户")
        print("  6. 删除用户")
        print("  7. 查看用户列表")
        print("  8. 推送到 GitHub")
        print("  9. 退出")
        print("-" * 50)
        choice = input("请选择操作 (1-9): ").strip()

        def _run(fn, args):
            """包装执行，捕获异常防止窗口直接关闭"""
            try:
                fn(args)
            except SystemExit as e:
                if e.code != 0:
                    print("\n[!] 操作失败，请查看上方错误信息")
            except Exception as e:
                print(f"\n[!] 操作失败: {e}")

        if choice == "1":
            print("\n--- 添加课件 ---")
            print("分类: scratch | python | cpp")
            print("阶段: scratch->A,B  python->初级,中级,高级  cpp->基础,进阶,竞赛算法")
            ppt_path = input("PPT文件路径 (可拖入文件): ").strip().strip('"')
            cat = input("输入分类 (scratch/python/cpp): ").strip()
            stage = input("输入阶段: ").strip()
            name = input("输入显示名称 (回车使用文件名): ").strip()
            _run(cmd_add, [ppt_path, cat, stage] + ([name] if name else []))
            input("\n按回车继续...")

        elif choice == "2":
            print("\n--- 批量添加课件 ---")
            print("说明: 将批量处理指定文件夹内所有 .pptx 和 .ppt 文件")
            print("分类: scratch | python | cpp")
            print("阶段: scratch->A,B  python->初级,中级,高级  cpp->基础,进阶,竞赛算法")
            folder_path = input("文件夹路径 (可拖入文件夹): ").strip().strip('"')
            cat = input("输入分类 (scratch/python/cpp): ").strip()
            stage = input("输入阶段: ").strip()
            name_prefix = input("输入显示名称前缀 (回车不使用前缀): ").strip()
            _run(cmd_batch_add, [folder_path, cat, stage] + ([name_prefix] if name_prefix else []))
            input("\n按回车继续...")

        elif choice == "3":
            print("\n--- 删除课件 ---")
            cat = input("输入分类 (scratch/python/cpp): ").strip()
            stage = input("输入阶段: ").strip()
            target = input("输入序号或名称: ").strip()
            _run(cmd_delete, [cat, stage, target])
            input("\n按回车继续...")

        elif choice == "4":
            _run(cmd_list, [])
            input("\n按回车继续...")

        elif choice == "5":
            print("\n--- 添加用户 ---")
            username = input("输入用户名: ").strip()
            password = input("输入密码: ").strip()
            is_admin = input("设为管理员? (y/n): ").strip().lower() == "y"
            _run(cmd_user, ["add", username, password] + (["--admin"] if is_admin else []))
            input("\n按回车继续...")

        elif choice == "6":
            print("\n--- 删除用户 ---")
            username = input("输入用户名: ").strip()
            _run(cmd_user, ["delete", username])
            input("\n按回车继续...")

        elif choice == "7":
            _run(cmd_user, ["list"])
            input("\n按回车继续...")

        elif choice == "8":
            print("\n--- 推送到 GitHub ---")
            msg = input("输入提交说明 (回车使用默认): ").strip()
            _run(cmd_push, [msg] if msg else [])
            input("\n按回车继续...")

        elif choice == "9":
            print("再见!")
            break
        else:
            print("无效选择，请重新输入。")


# ── 主入口 ──
def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    cmd = sys.argv[1]
    args = sys.argv[2:]

    if cmd == "add":
        cmd_add(args)
    elif cmd == "batch_add":
        cmd_batch_add(args)
    elif cmd == "delete":
        cmd_delete(args)
    elif cmd == "list":
        cmd_list(args)
    elif cmd == "user":
        cmd_user(args)
    elif cmd == "push":
        cmd_push(args)
    elif cmd == "push_data":
        cmd_push_data(args)
    elif cmd == "init_data":
        cmd_init_data(args)
    elif cmd == "menu":
        cmd_menu()
    elif cmd == "help" or cmd == "--help" or cmd == "-h":
        print(__doc__)
    else:
        print(f"未知命令: {cmd}")
        print(__doc__)
        sys.exit(1)

if __name__ == "__main__":
    main()
