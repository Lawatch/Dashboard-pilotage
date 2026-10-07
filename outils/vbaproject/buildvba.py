"""Construit un vbaProject.bin (source seule, recompilé par Excel à l'ouverture).

modules : liste de (nom, 'doc' | 'std', code_sans_attributs, base) ; base = 'wb' | 'ws' pour les modules document.
On réutilise PROJECTINFORMATION + PROJECTREFERENCES du projet d'origine (dir stream).
"""
import json, os, struct, subprocess, tempfile, uuid
from ovba import compress
from dirparse import records

HERE = os.path.dirname(os.path.abspath(__file__))
WRITER = os.path.join(HERE, 'writecfb.js')
CP = 'cp1252'

DOC_BASE = {'wb': '0{00020819-0000-0000-C000-000000000046}',
            'ws': '0{00020820-0000-0000-C000-000000000046}'}


def _rec(rid, data):
    return struct.pack('<HI', rid, len(data)) + data


def _attrs(name, kind, base):
    if kind == 'std':
        return 'Attribute VB_Name = "%s"\r\n' % name
    return ('Attribute VB_Name = "%s"\r\n'
            'Attribute VB_Base = "%s"\r\n'
            'Attribute VB_GlobalNameSpace = False\r\n'
            'Attribute VB_Creatable = False\r\n'
            'Attribute VB_PredeclaredId = True\r\n'
            'Attribute VB_Exposed = True\r\n'
            'Attribute VB_TemplateDerived = False\r\n'
            'Attribute VB_Customizable = True\r\n') % (name, DOC_BASE[base])


def build(modules, orig_dir, orig_project_text, out_path):
    # 1) dir stream : en-tête d'origine jusqu'à PROJECTMODULES
    head = b''
    for rid, i, b in records(orig_dir):
        if rid == 0x000F:
            head = orig_dir[:i]
            break
    assert head
    body = _rec(0x000F, struct.pack('<H', len(modules))) + _rec(0x0013, struct.pack('<H', 0xFFFF))
    for name, kind, code, base in modules:
        nb = name.encode(CP); nu = name.encode('utf-16-le')
        body += _rec(0x0019, nb) + _rec(0x0047, nu)
        body += _rec(0x001A, nb) + _rec(0x0032, nu)
        body += _rec(0x001C, b'') + _rec(0x0048, b'')
        body += _rec(0x0031, struct.pack('<I', 0))
        body += _rec(0x001E, struct.pack('<I', 0))
        body += _rec(0x002C, struct.pack('<H', 0xFFFF))
        body += _rec(0x0021 if kind == 'std' else 0x0022, b'')
        body += _rec(0x002B, b'')
    body += _rec(0x0010, b'')
    dir_stream = head + body

    # 2) PROJECT stream
    lines = orig_project_text.replace('\r\n', '\n').split('\n')
    keep = [l for l in lines if l.startswith(('Name=', 'HelpContextID=', 'VersionCompatible32=', 'CMG=', 'DPB=', 'GC='))]
    proj = [l for l in lines if l.startswith('ID=')]
    for name, kind, _, _ in modules:
        proj.append(('Document=%s/&H00000000' % name) if kind == 'doc' else ('Module=%s' % name))
    proj += keep
    # ID, CMG, DPB, GC repris du projet d'origine (ils sont liés entre eux)
    proj += ['',
             '[Host Extender Info]', '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000', '',
             '[Workspace]']
    proj += ['%s=0, 0, 0, 0, C' % m[0] for m in modules]
    project_stream = ('\r\n'.join(proj) + '\r\n').encode(CP)

    # 3) PROJECTwm
    wm = b''
    for name, _, _, _ in modules:
        wm += name.encode(CP) + b'\x00' + name.encode('utf-16-le') + b'\x00\x00'
    wm += b'\x00\x00'

    tmp = tempfile.mkdtemp()
    items = []

    def add(path, data):
        f = os.path.join(tmp, str(len(items)))
        open(f, 'wb').write(data)
        items.append([path, f])

    add('/PROJECT', project_stream)
    add('/PROJECTwm', wm)
    add('/VBA/_VBA_PROJECT', bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]))
    add('/VBA/dir', compress(dir_stream))
    for name, kind, code, base in modules:
        src = _attrs(name, kind, base) + code.replace('\r\n', '\n').replace('\n', '\r\n')
        if not src.endswith('\r\n'):
            src += '\r\n'
        add('/VBA/' + name, compress(src.encode(CP)))
    man = os.path.join(tmp, 'manifest.json')
    json.dump(items, open(man, 'w'))
    subprocess.check_call(['node', WRITER, man, os.path.abspath(out_path)], cwd=HERE)
    return out_path
