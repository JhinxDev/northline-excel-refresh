"""Remove workstation metadata from an Excel Power Query release copy.

MS-QDEFF framing: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/22557f6d-7c29-4554-8fe4-7b7a54ac7a2b
This never lowers Power Query privacy settings. Re-test the resulting copy in Excel.
"""
from pathlib import Path
import argparse, base64, io, re, struct, zipfile
import xml.etree.ElementTree as ET

PLACEHOLDER='SET_YOUR_FIXTURES_FOLDER'

def unpack_stream(data, count):
    assert data[:4]==b'\0'*4, 'Unsupported stream version'
    pos=4;fields=[]
    for _ in range(count):
        n=struct.unpack_from('<I',data,pos)[0];pos+=4
        assert pos+n<=len(data),'Invalid stream length'
        fields.append(data[pos:pos+n]);pos+=n
    assert pos==len(data),'Unexpected trailing stream bytes'
    return fields

def pack_stream(fields):
    return b'\0'*4+b''.join(struct.pack('<I',len(v))+v for v in fields)

def sanitize(source, target):
    with zipfile.ZipFile(source) as z:
        parts={n:z.read(n) for n in z.namelist()}
    mashup=next(n for n in parts if n.startswith('customXml/item') and b'D\x00a\x00t\x00a\x00M\x00a\x00s\x00h\x00u\x00p' in parts[n])
    root=ET.fromstring(parts[mashup]);fields=unpack_stream(base64.b64decode(root.text),4)
    with zipfile.ZipFile(io.BytesIO(fields[0])) as z:
        inner={n:z.read(n) for n in z.namelist()}
    code=inner['Formulas/Section1.m'].decode('utf-8-sig')
    code,count=re.subn(r'shared DataRoot = .*?;(?=\r?\n)',f'shared DataRoot = "{PLACEHOLDER}";',code,count=1,flags=re.S)
    assert count==1,'Missing DataRoot query'
    inner['Formulas/Section1.m']=code.encode('utf-8')
    stream=io.BytesIO()
    with zipfile.ZipFile(stream,'w',zipfile.ZIP_DEFLATED) as z:
        for n,v in inner.items():z.writestr(n,v)
    fields[0]=stream.getvalue()
    # Drop opaque user/machine-bound permission signatures, not privacy policy.
    fields[3]=b''
    root.attrib.pop('sqmid',None)
    root.text=base64.b64encode(pack_stream(fields)).decode('ascii')
    ET.register_namespace('','http://schemas.microsoft.com/DataMashup')
    parts[mashup]=ET.tostring(root,encoding='utf-16',xml_declaration=True)
    core=parts['docProps/core.xml'].decode('utf-8')
    for tag in ['dc:creator','cp:lastModifiedBy']:
        core=re.sub(fr'<{tag}>.*?</{tag}>',f'<{tag}>JhinxDev</{tag}>',core)
    core=core.replace('v1.0.0','v1.1.0')
    parts['docProps/core.xml']=core.encode('utf-8')
    workbook=parts['xl/workbook.xml'].decode('utf-8')
    workbook=re.sub(r'<[\w]+:absPath\b[^>]*/>','',workbook)
    parts['xl/workbook.xml']=workbook.encode('utf-8')
    # Printer DEVMODE blobs can expose device-specific information.
    for n in list(parts):
        if n.startswith('xl/printerSettings/'):
            del parts[n]
        elif n.endswith('.rels'):
            rel=parts[n].decode('utf-8')
            rel=re.sub(r'<Relationship\b[^>]*Type="[^"]*/printerSettings"[^>]*/>','',rel)
            parts[n]=rel.encode('utf-8')
        elif n.startswith('xl/worksheets/') and n.endswith('.xml'):
            s=parts[n].decode('utf-8')
            s=re.sub(r'(<pageSetup\b[^>]*?) r:id="[^"]*"',r'\1',s)
            parts[n]=s.encode('utf-8')
    target.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED) as z:
        for n,v in parts.items():z.writestr(n,v)

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('source',type=Path);p.add_argument('target',type=Path)
    a=p.parse_args()
    if a.source.resolve()==a.target.resolve():p.error('Use a separate release copy')
    sanitize(a.source,a.target)
