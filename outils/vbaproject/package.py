"""Transforme un .xlsx écrit par openpyxl en .xlsm avec un vbaProject.bin donné."""
import re, zipfile


def make_xlsm(xlsx_in, vba_bin, out, sheet_codenames, wb_codename='ThisWorkbook'):
    """sheet_codenames : {nom d'onglet: codeName}"""
    zin = zipfile.ZipFile(xlsx_in)
    names = zin.namelist()
    wbxml = zin.read('xl/workbook.xml').decode('utf-8')
    rels = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    # onglet -> fichier xml
    rid_target = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid_target.update({a: b for b, a in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    sheet_file = {}
    for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wbxml):
        nm = m.group(1).replace('&amp;', '&').replace('&apos;', "'").replace('&quot;', '"')
        t = rid_target[m.group(2)]
        t = t.lstrip('/')
        if not t.startswith('xl/'):
            t = 'xl/' + t
        sheet_file[nm] = t
    # codeName du classeur
    if '<workbookPr' in wbxml:
        if 'codeName=' not in wbxml.split('<workbookPr', 1)[1].split('>', 1)[0]:
            wbxml = wbxml.replace('<workbookPr', '<workbookPr codeName="%s"' % wb_codename, 1)
    else:
        wbxml = re.sub(r'(<workbook[^>]*>)', r'\1<workbookPr codeName="%s"/>' % wb_codename, wbxml, 1)
    rels = rels.replace('</Relationships>',
                        '<Relationship Id="rIdVBA1" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
    ct = zin.read('[Content_Types].xml').decode('utf-8')
    ct = ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
                    'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
    if 'Extension="bin"' not in ct:
        ct = ct.replace('<Default ', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ', 1)

    zout = zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED)
    for n in names:
        data = zin.read(n)
        if n == 'xl/workbook.xml':
            data = wbxml.encode('utf-8')
        elif n == 'xl/_rels/workbook.xml.rels':
            data = rels.encode('utf-8')
        elif n == '[Content_Types].xml':
            data = ct.encode('utf-8')
        else:
            for nm, f in sheet_file.items():
                if f == n and nm in sheet_codenames:
                    x = data.decode('utf-8')
                    cn = sheet_codenames[nm]
                    if re.search(r'<sheetPr\b', x):
                        if 'codeName=' in x.split('<sheetPr', 1)[1].split('>', 1)[0]:
                            x = re.sub(r'(<sheetPr\b[^>]*?)codeName="[^"]*"', r'\1codeName="%s"' % cn, x, 1)
                        else:
                            x = x.replace('<sheetPr', '<sheetPr codeName="%s"' % cn, 1)
                    else:
                        x = re.sub(r'(<worksheet\b[^>]*>)', r'\1<sheetPr codeName="%s"/>' % cn, x, 1)
                    data = x.encode('utf-8')
        zout.writestr(n, data)
    zout.writestr('xl/vbaProject.bin', open(vba_bin, 'rb').read())
    zout.close()
