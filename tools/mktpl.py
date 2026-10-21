import zipfile, re, sys, base64
src, out = sys.argv[1], sys.argv[2]
z = zipfile.ZipFile(src)
files = {n: z.read(n) for n in z.namelist()}
keep = {}
def t(n): return files[n].decode('utf8')
# workbook
wb = t('xl/workbook.xml')
wb = re.sub(r'<sheets>.*?</sheets>', '<sheets><sheet name="Propuesta" sheetId="1" r:id="rId1"/><sheet name="_Lenovo" sheetId="2" r:id="rId2"/></sheets>', wb, flags=re.S)
wb = re.sub(r'<definedNames>.*?</definedNames>', '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">{{PRINT}}</definedName></definedNames>', wb, flags=re.S)
wb = re.sub(r'<calcPr[^>]*/>', '<calcPr calcId="191029" fullCalcOnLoad="1"/>', wb)
wb = re.sub(r'activeTab="\d+"', 'activeTab="0"', wb)
wb = re.sub(r'<mc:AlternateContent.*?</mc:AlternateContent>', '', wb, flags=re.S)
keep['xl/workbook.xml'] = wb
keep['xl/_rels/workbook.xml.rels'] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
  '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
  '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
  '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
  '<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
  '</Relationships>')
def tplsheet(x, extra_tail=False):
    x = re.sub(r'<dimension ref="[^"]*"/>', '<dimension ref="{{DIM}}"/>', x)
    x = re.sub(r'<sheetData>.*?</sheetData>', '<sheetData>{{DATA}}</sheetData>', x, flags=re.S)
    x = re.sub(r'<mergeCells count="\d+">.*?</mergeCells>', '{{MERGES}}', x, flags=re.S)
    x = re.sub(r'<selection [^>]*/>', '<selection activeCell="B2" sqref="B2"/>', x)
    return x
s1 = tplsheet(t('xl/worksheets/sheet2.xml'))
s1 = s1.replace('Hernán Gavilanes', '{{FIRMA}}')  # firma editable
s1 = s1.replace('(09)9 321-5391', '{{CEL}}').replace('gerencia2@coresolutions.com.ec', '{{EMAIL}}')
assert '{{CEL}}' in s1 and '{{EMAIL}}' in s1
assert '{{FIRMA}}' in s1
s2 = tplsheet(t('xl/worksheets/sheet4.xml')).replace('tabSelected="1" ','')
s1 = s1  # tab selected stays on propuesta
keep['xl/worksheets/sheet1.xml'] = s1
keep['xl/worksheets/sheet2.xml'] = s2
keep['xl/worksheets/_rels/sheet1.xml.rels'] = t('xl/worksheets/_rels/sheet2.xml.rels').replace('printerSettings2.bin','printerSettings1.bin').replace('vmlDrawing2.vml','vmlDrawing1.vml')
keep['xl/printerSettings/printerSettings1.bin'] = files['xl/printerSettings/printerSettings2.bin']
keep['xl/drawings/vmlDrawing1.vml'] = files['xl/drawings/vmlDrawing2.vml']
keep['xl/drawings/_rels/vmlDrawing1.vml.rels'] = files['xl/drawings/_rels/vmlDrawing2.vml.rels']
for n in ['xl/theme/theme1.xml','xl/styles.xml','docProps/core.xml','docProps/app.xml','_rels/.rels']:
    keep[n] = files[n]
print(t('xl/drawings/_rels/vmlDrawing2.vml.rels'))
for n in re.findall(r'media/image\d+\.png', t('xl/drawings/_rels/vmlDrawing2.vml.rels')):
    keep['xl/'+n] = files['xl/'+n]
keep['xl/sharedStrings.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>'
app = t('docProps/app.xml')
app = re.sub(r'<HeadingPairs>.*?</TitlesOfParts>', '', app, flags=re.S)
keep['docProps/app.xml'] = app
ct = t('[Content_Types].xml')
ct = re.sub(r'<Override PartName="/xl/(worksheets/sheet[34]\.xml|calcChain\.xml)"[^>]*/>', '', ct)
ct = re.sub(r'<Override PartName="/customXml[^>]*/>', '', ct)
keep['[Content_Types].xml'] = ct
print(ct)
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as o:
    for n, d in keep.items():
        o.writestr(n, d if isinstance(d, bytes) else d.encode('utf8'))
print('ok', sum(len(v) for v in keep.values()))
