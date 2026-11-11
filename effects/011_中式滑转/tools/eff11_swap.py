"""效果11 · 媒体替换法（已验证可用，不会触发 PowerPoint 修复）

复制模板 → 只替换 ppt/media/image1~3.jpg → 输出。
页数固定为模板的 4 页（3 张图 + 封面）；增减页会破坏包结构，请另存多页版本。

用法:
    python eff11_swap.py --template 效果11.pptx --out 输出.pptx 图1 图2 图3
图片顺序按模板每页实际引用的媒体：
    image1.jpg → P1/P2    image2.jpg → P3    image3.jpg → P4
"""
import os, re, sys, zipfile, argparse


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--template', required=True)
    ap.add_argument('--out', required=True)
    ap.add_argument('images', nargs='+')
    a = ap.parse_args()

    tpl = a.template
    with zipfile.ZipFile(tpl) as z:
        items = {n: z.read(n) for n in z.namelist()}

    # 统计每页引用的媒体，按出现顺序收集（去重保序）
    media_order = []
    for n in sorted(items):
        m = re.match(r'ppt/slides/slide(\d+)\.xml$', n)
        if not m:
            continue
        rels_name = 'ppt/slides/_rels/slide%s.xml.rels' % m.group(1)
        if rels_name not in items:
            continue
        rels = items[rels_name].decode('utf-8')
        for t in re.findall(r'Target="\.\./media/([^"]+)"', rels):
            if t not in media_order:
                media_order.append(t)

    if len(a.images) < len(media_order):
        print('注意：模板需要 %d 张图，实际给了 %d 张，后几页保持原图' % (len(media_order), len(a.images)))

    for i, name in enumerate(media_order):
        if i >= len(a.images):
            break
        src = a.images[i]
        with open(src, 'rb') as f:
            items['ppt/media/' + name] = f.read()
        print('  替换 ppt/media/%s ← %s' % (name, os.path.basename(src)))

    with zipfile.ZipFile(a.out, 'w', zipfile.ZIP_DEFLATED) as z:
        for k, v in items.items():
            z.writestr(k, v)
    print('输出', a.out, round(os.path.getsize(a.out) / 1024), 'KB（包结构未改动）')


if __name__ == '__main__':
    main()
