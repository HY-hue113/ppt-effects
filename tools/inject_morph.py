import zipfile
from lxml import etree

f = r'<USER_HOME>\<OUTPUT_DIR>\效果10.pptx'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main'

TMPL = (
    '<mc:AlternateContent xmlns:mc="' + MC + '" xmlns:p="' + P + '" '
    'xmlns:p14="' + P14 + '" xmlns:p159="' + P159 + '">'
    '<mc:Choice Requires="p159">'
    '<p:transition spd="slow" p14:dur="1800"><p159:morph option="byObject"/></p:transition>'
    '</mc:Choice>'
    '<mc:Fallback>'
    '<p:transition spd="slow" p14:dur="1800"><p:fade/></p:transition>'
    '</mc:Fallback>'
    '</mc:AlternateContent>'
)

with zipfile.ZipFile(f) as z:
    items = {n: z.read(n) for n in z.namelist()}

fix = 0
for n in ('ppt/slides/slide2.xml', 'ppt/slides/slide3.xml'):
    root = etree.fromstring(items[n])
    if root.find('{%s}transition' % P) is not None:
        print(n, '已有 transition')
        continue
    root.append(etree.fromstring(TMPL.encode()))
    items[n] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    fix += 1

with zipfile.ZipFile(f, 'w', zipfile.ZIP_DEFLATED) as zo:
    for k, d in items.items():
        zo.writestr(k, d)

print('injected', fix)
with zipfile.ZipFile(f) as z:
    for n in ('ppt/slides/slide1.xml', 'ppt/slides/slide2.xml', 'ppt/slides/slide3.xml'):
        s = z.read(n).decode()
        print(n, 'morph:', 'morph' in s, '| dur:', 'p14:dur="1800"' in s)
