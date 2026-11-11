import zipfile, argparse, os
ap=argparse.ArgumentParser(); ap.add_argument('--template',required=True); ap.add_argument('--out',required=True); ap.add_argument('images',nargs='+'); a=ap.parse_args()
z=zipfile.ZipFile(a.template); items={n:z.read(n) for n in z.namelist()}
media=sorted(n for n in items if n.startswith('ppt/media/'))
for i,m in enumerate(media):
    items[m]=open(a.images[i%len(a.images)],'rb').read()
with zipfile.ZipFile(a.out,'w',zipfile.ZIP_DEFLATED) as zo:
    for k,v in items.items(): zo.writestr(k,v)
print('输出',a.out,round(os.path.getsize(a.out)/1024),'KB | 替换',len(media),'个媒体')
