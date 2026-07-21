#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
长舟编程学习平台 - 腾讯云COS图片上传工具

功能：
  1. 批量上传本地 ppts/ 下的所有图片到腾讯云COS
  2. 增量上传（跳过已存在的文件）
  3. 上传单个课件目录
  4. 删除COS上的指定文件

用法：
  python cos_uploader.py sync          # 全量同步（增量上传所有图片）
  python cos_uploader.py upload <目录>  # 上传指定目录（如 ppts/python/中级/第19课-相声报菜名_images）
  python cos_uploader.py delete <路径>  # 删除COS上的文件或目录
  python cos_uploader.py list           # 列出COS上的文件统计
"""

import sys
import os
import json
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_FILE = SCRIPT_DIR / "cos_config.json"
PPTS_DIR = SCRIPT_DIR / "ppts"

# 图片扩展名
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".webp"}


def load_config():
    if not CONFIG_FILE.exists():
        print("错误: cos_config.json 不存在")
        print("请先运行配置向导: python cos_uploader.py setup")
        sys.exit(1)
    config = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    if not config.get("secret_id") or not config.get("secret_key"):
        print("错误: cos_config.json 中 secret_id 或 secret_key 为空")
        print("请编辑 cos_config.json 填入你的腾讯云API密钥")
        sys.exit(1)
    if not config.get("bucket"):
        print("错误: cos_config.json 中 bucket 为空")
        print("请编辑 cos_config.json 填入你的存储桶名称（如 changzhou-ppt-1250000000）")
        sys.exit(1)
    return config


def get_cos_client(config):
    from qcloud_cos import CosConfig, CosS3Client

    cos_config = CosConfig(
        Region=config["region"],
        SecretId=config["secret_id"],
        SecretKey=config["secret_key"],
        Scheme="https"
    )
    return CosS3Client(cos_config)


def get_cdn_base(config):
    """获取CDN基础URL"""
    if config.get("cdn_base"):
        return config["cdn_base"].rstrip("/") + "/"
    # 默认域名格式: https://<bucket>.cos.<region>.myqcloud.com/
    return f"https://{config['bucket']}.cos.{config['region']}.myqcloud.com/"


def upload_file(client, bucket, local_path, cos_key, skip_existing=True):
    """上传单个文件到COS，返回是否实际上传"""
    try:
        # 检查是否已存在（通过HEAD请求）
        if skip_existing:
            try:
                client.head_object(Bucket=bucket, Key=cos_key)
                return False  # 已存在，跳过
            except Exception:
                pass  # 不存在，继续上传

        client.upload_file(
            Bucket=bucket,
            Key=cos_key,
            LocalFilePath=str(local_path),
            EnableMD5=True
        )
        return True
    except Exception as e:
        print(f"  [!] 上传失败: {cos_key} -> {e}")
        return False


def upload_dir(client, bucket, local_dir, cos_prefix=""):
    """上传整个目录到COS"""
    local_dir = Path(local_dir)
    if not local_dir.exists():
        print(f"目录不存在: {local_dir}")
        return 0, 0

    uploaded = 0
    skipped = 0
    files = [f for f in local_dir.rglob("*") if f.is_file() and f.suffix.lower() in IMAGE_EXTS]

    for i, fpath in enumerate(files, 1):
        # 计算COS中的key
        rel_path = fpath.relative_to(local_dir)
        cos_key = cos_prefix + "/" + str(rel_path).replace("\\", "/") if cos_prefix else str(rel_path).replace("\\", "/")

        # 如果是从ppts目录上传，保持 ppts/... 的路径结构
        if not cos_prefix:
            try:
                rel_to_ppts = fpath.relative_to(PPTS_DIR)
                cos_key = "ppts/" + str(rel_to_ppts).replace("\\", "/")
            except ValueError:
                pass

        result = upload_file(client, bucket, fpath, cos_key)
        if result:
            uploaded += 1
            if uploaded % 50 == 0:
                print(f"  已上传 {uploaded} 个文件...")
        else:
            skipped += 1

    return uploaded, skipped


def cmd_sync(args):
    """全量同步所有图片到COS"""
    config = load_config()
    client = get_cos_client(config)

    print("=" * 60)
    print("长舟编程学习平台 - COS全量图片同步")
    print("=" * 60)
    print(f"存储桶: {config['bucket']}")
    print(f"地域:   {config['region']}")
    print(f"CDN地址: {get_cdn_base(config)}")
    print()

    # 找到所有图片目录
    image_dirs = list(PPTS_DIR.glob("*/*/*_images"))
    if not image_dirs:
        print("未找到任何图片目录")
        return

    print(f"发现 {len(image_dirs)} 个图片目录\n")

    total_uploaded = 0
    total_skipped = 0

    for i, img_dir in enumerate(image_dirs, 1):
        print(f"[{i}/{len(image_dirs)}] {img_dir.relative_to(PPTS_DIR)}")

        # 计算COS key前缀: ppts/<cat>/<stage>/<stem>_images
        rel = img_dir.relative_to(PPTS_DIR)
        cos_prefix = "ppts/" + str(rel).replace("\\", "/")

        uploaded, skipped = upload_dir(client, config["bucket"], img_dir, cos_prefix)
        total_uploaded += uploaded
        total_skipped += skipped
        print(f"  上传 {uploaded} 个, 跳过 {skipped} 个")

    print()
    print("=" * 60)
    print(f"同步完成! 新上传: {total_uploaded}, 跳过(已存在): {total_skipped}")
    print(f"CDN地址: {get_cdn_base(config)}")
    print("=" * 60)


def cmd_upload(args):
    """上传指定目录"""
    if len(args) < 1:
        print("用法: python cos_uploader.py upload <目录路径>")
        sys.exit(1)

    config = load_config()
    client = get_cos_client(config)
    target_dir = Path(args[0])

    print(f"上传目录: {target_dir}")
    uploaded, skipped = upload_dir(client, config["bucket"], target_dir)
    print(f"\n完成! 上传 {uploaded} 个, 跳过 {skipped} 个")


def cmd_delete(args):
    """删除COS上的文件或目录"""
    if len(args) < 1:
        print("用法: python cos_uploader.py delete <COS路径前缀>")
        print("示例: python cos_uploader.py delete ppts/python/中级/第19课-相声报菜名_images")
        sys.exit(1)

    config = load_config()
    client = get_cos_client(config)
    prefix = args[0].replace("\\", "/").strip("/")

    print(f"删除COS路径前缀: {prefix}/")

    # 列出所有匹配的文件
    deleted = 0
    marker = ""
    while True:
        response = client.list_objects(
            Bucket=config["bucket"],
            Prefix=prefix + "/",
            Marker=marker,
            MaxKeys=1000
        )
        contents = response.get("Contents", [])
        for item in contents:
            key = item["Key"]
            client.delete_object(Bucket=config["bucket"], Key=key)
            deleted += 1
            if deleted % 50 == 0:
                print(f"  已删除 {deleted} 个文件...")

        if response.get("IsTruncated") == "true":
            marker = response.get("NextMarker", "")
        else:
            break

    print(f"\n完成! 删除 {deleted} 个文件")


def cmd_list(args):
    """列出COS上的文件统计"""
    config = load_config()
    client = get_cos_client(config)

    print(f"存储桶: {config['bucket']}")
    print(f"CDN地址: {get_cdn_base(config)}")
    print()

    # 统计各分类的文件数
    categories = {}
    total = 0
    marker = ""
    while True:
        response = client.list_objects(
            Bucket=config["bucket"],
            Prefix="ppts/",
            Marker=marker,
            MaxKeys=1000
        )
        contents = response.get("Contents", [])
        for item in contents:
            total += 1
            # 提取分类: ppts/<cat>/<stage>/...
            parts = item["Key"].split("/")
            if len(parts) >= 3:
                cat_stage = f"{parts[1]}/{parts[2]}"
                categories[cat_stage] = categories.get(cat_stage, 0) + 1

        if response.get("IsTruncated") == "true":
            marker = response.get("NextMarker", "")
        else:
            break

    print(f"总文件数: {total}")
    print()
    for cat_stage in sorted(categories.keys()):
        print(f"  {cat_stage}: {categories[cat_stage]} 个文件")


def cmd_setup(args):
    """配置向导"""
    print("=" * 60)
    print("腾讯云COS配置向导")
    print("=" * 60)
    print()
    print("请按以下步骤操作：")
    print()
    print("1. 登录腾讯云控制台: https://console.cloud.tencent.com/cos")
    print("2. 开通COS服务（如未开通）")
    print("3. 创建存储桶:")
    print("   - 名称: 自定义（如 changzhou-ppt）")
    print("   - 地域: 建议选广州(ap-guangzhou)或离你最近的")
    print("   - 访问权限: 选择「公有读私有写」")
    print("   - 记下完整的存储桶名称（如 changzhou-ppt-1250000000）")
    print()
    print("4. 获取API密钥:")
    print("   - 访问: https://console.cloud.tencent.com/cam/capi")
    print("   - 新建密钥，复制 SecretId 和 SecretKey")
    print()
    print("5. 将获取的信息填入 cos_config.json:")
    print("   {")
    print('     "secret_id":  "你的SecretId",')
    print('     "secret_key": "你的SecretKey",')
    print('     "region":     "ap-guangzhou",')
    print('     "bucket":     "changzhou-ppt-1250000000",')
    print('     "cdn_base":   ""  (留空用默认域名)')
    print("   }")
    print()
    print("配置完成后运行: python cos_uploader.py sync  批量上传所有图片")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    cmd = sys.argv[1]
    args = sys.argv[2:]

    if cmd == "sync":
        cmd_sync(args)
    elif cmd == "upload":
        cmd_upload(args)
    elif cmd == "delete":
        cmd_delete(args)
    elif cmd == "list":
        cmd_list(args)
    elif cmd == "setup":
        cmd_setup(args)
    else:
        print(f"未知命令: {cmd}")
        print(__doc__)


if __name__ == "__main__":
    main()
