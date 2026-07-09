#!/usr/bin/env python3
"""更新 manifest.json，为每个 PPT 添加 slides 图片路径列表"""
import json
from pathlib import Path

PPTS_DIR = Path(__file__).parent / 'ppts'
manifest = json.loads((PPTS_DIR / 'manifest.json').read_text(encoding='utf-8'))

for cat, stages in manifest.items():
    for stage, files in stages.items():
        for f in files:
            file_rel = f.get('file', '')
            if not file_rel:
                continue
            stem = Path(file_rel).stem
            img_dir = Path(file_rel).parent / f'{stem}_images'
            img_abs = PPTS_DIR.parent / img_dir
            if img_abs.exists():
                imgs = sorted(img_abs.glob('slide_*.png'))
                f['slides'] = [str(img_dir / img.name).replace('\\', '/') for img in imgs]
                f['slideCount'] = len(imgs)
                print(f'{file_rel}: {len(imgs)} slides')
            else:
                f['slides'] = []
                f['slideCount'] = 0
                print(f'{file_rel}: NO IMAGES')

(PPTS_DIR / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
print('manifest.json updated!')
