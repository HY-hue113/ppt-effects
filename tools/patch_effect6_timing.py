import zipfile, os
from lxml import etree

SRC = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\preview\效果6.pptx'
DST = r'<USER_HOME>\Documents\<DSH_HOME>\default-workspace\preview\效果6.pptx'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
ns = {'p': P}

with zipfile.ZipFile(SRC) as z:
    items = {n: z.read(n) for n in z.namelist()}

root = etree.fromstring(items['ppt/slides/slide1.xml'])
ctns = {c.get('id'): c for c in root.iter('{%s}cTn' % P)}

# 1) 整个旋转组的启动：从“等待单击”改为自动开始
g = ctns['14']
st = g.find('p:stCondLst', ns)
c0 = st.find('p:cond', ns)
old = c0.get('delay')
c0.set('delay', '0')
g.set('nodeType', 'afterEffect')

# 2) 第一个陀螺旋：clickEffect -> afterEffect，延迟归零
first = ctns['16']
first.set('nodeType', 'afterEffect')
st2 = first.find('p:stCondLst', ns)
if st2 is not None:
    c = st2.find('p:cond', ns)
    if c is not None:
        c.set('delay', '0')

items['ppt/slides/slide1.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
    for n, d in items.items():
        z.writestr(n, d)

print('patched; group delay was', old, '| first spin nodeType -> afterEffect')
