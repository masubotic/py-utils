"""
Skrypt do laczenia plikow PPTX w jeden plik.
Uzycie: python merge_pptx.py <folder> [plik_wyjsciowy.pptx]
Wymagania: pip install lxml
"""
import sys
import re
import zipfile
from pathlib import Path
from lxml import etree

PRS_NS   = 'http://schemas.openxmlformats.org/presentationml/2006/main'
REL_NS   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT_NS    = 'http://schemas.openxmlformats.org/package/2006/content-types'
SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
NOTES_CT = 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml'

REMOVE_FILES = {'ppt/revisionInfo.xml', 'ppt/changesInfos/changesInfo1.xml'}
REMOVE_TYPES = {
    'http://schemas.microsoft.com/office/2015/10/relationships/revisionInfo',
    'http://schemas.microsoft.com/office/2016/11/relationships/changesInfo',
}

EMPTY_RELS = (
    b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>'
)

CT_MIME = {
    'jpg': 'image/jpeg', 'jpeg': 'image/jpeg',
    'gif': 'image/gif',  'bmp': 'image/bmp',
    'tiff': 'image/tiff', 'wmf': 'image/x-wmf',
    'emf': 'image/x-emf', 'svg': 'image/svg+xml',
}


def fix_decl(data):
    data = data.replace(b"<?xml version='1.0' encoding='UTF-8' standalone='yes'?>",
                        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    data = re.sub(rb'(<\?xml[^?]*\?>)(\r?\n)?', rb'\1\r\n', data, count=1)
    return data


def serialize(root):
    return fix_decl(etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True))


def remove_revision(data_dict):
    for f in REMOVE_FILES:
        data_dict.pop(f, None)
    ct_data = data_dict.get('[Content_Types].xml')
    if ct_data:
        root = etree.fromstring(ct_data)
        for el in list(root):
            if el.get('PartName', '').lstrip('/') in REMOVE_FILES:
                root.remove(el)
        data_dict['[Content_Types].xml'] = serialize(root)
    for name in list(data_dict.keys()):
        if not name.endswith('.rels'):
            continue
        root = etree.fromstring(data_dict[name])
        if any(el.get('Type') in REMOVE_TYPES for el in root):
            for el in list(root):
                if el.get('Type') in REMOVE_TYPES:
                    root.remove(el)
            data_dict[name] = serialize(root)


def get_max_rid(root):
    ids = [int(re.search(r'\d+', el.get('Id', '')).group())
           for el in root if re.search(r'\d+', el.get('Id', ''))]
    return max(ids, default=0)


def merge_pptx(input_files, output_file):
    sources = []
    for f in input_files:
        with zipfile.ZipFile(f) as z:
            infos = [i for i in z.infolist() if not i.is_dir()]
            sources.append({
                'path': f,
                'infos': infos,
                'data': {i.filename: z.read(i.filename) for i in infos},
            })

    base = sources[0]
    bd   = base['data']
    remove_revision(bd)

    prs_root      = etree.fromstring(bd['ppt/presentation.xml'])
    prs_rels_root = etree.fromstring(bd['ppt/_rels/presentation.xml.rels'])
    ct_root       = etree.fromstring(bd['[Content_Types].xml'])

    sldIdLst    = prs_root.find(f'{{{PRS_NS}}}sldIdLst')
    next_id     = max((int(e.get('id')) for e in sldIdLst.findall(f'{{{PRS_NS}}}sldId')), default=255) + 1
    nrid        = get_max_rid(prs_rels_root) + 1
    existing_ct = {el.get('PartName') for el in ct_root if el.get('PartName')}

    slide_counter = sum(1 for k in bd if re.match(r'ppt/slides/slide\d+\.xml$', k))
    notes_counter = sum(1 for k in bd if re.match(r'ppt/notesSlides/notesSlide\d+\.xml$', k))

    new_data  = {}
    new_flags = {}

    for src_idx, src in enumerate(sources[1:], start=1):
        print(f"Dodaje: {Path(src['path']).name}")
        sd    = src['data']
        names = set(sd.keys())

        # --- media ---
        media_map = {}
        for n in sorted(n for n in names if n.startswith('ppt/media/')):
            stem, ext = Path(n).stem, Path(n).suffix
            new_fn = f"{stem}_s{src_idx}{ext}"
            arc = 'ppt/media/' + new_fn
            new_data[arc]  = sd[n]
            new_flags[arc] = (zipfile.ZIP_STORED, 0)
            media_map[Path(n).name] = new_fn

        # Uzupelnij brakujace Default extensions w Content_Types (np. jpg, jpeg, gif)
        known_exts = {el.get('Extension', '') for el in ct_root if el.get('Extension')}
        for n in (n for n in names if n.startswith('ppt/media/')):
            ext = Path(n).suffix.lstrip('.').lower()
            if ext and ext not in known_exts:
                known_exts.add(ext)
                e = etree.SubElement(ct_root, f'{{{CT_NS}}}Default')
                e.set('Extension', ext)
                e.set('ContentType', CT_MIME.get(ext, f'image/{ext}'))

        # --- embeddings ---
        embed_map = {}
        for n in sorted(n for n in names if n.startswith('ppt/embeddings/')):
            stem, ext = Path(n).stem, Path(n).suffix
            new_fn = f"{stem}_s{src_idx}{ext}"
            arc = 'ppt/embeddings/' + new_fn
            new_data[arc]  = sd[n]
            new_flags[arc] = (zipfile.ZIP_DEFLATED, 6)
            embed_map[Path(n).name] = new_fn

        # --- notes slides: buduj mape stary->nowy numer ---
        src_notes = sorted(
            [n for n in names if re.match(r'ppt/notesSlides/notesSlide\d+\.xml$', n)],
            key=lambda x: int(re.search(r'\d+', Path(x).name).group()),
        )
        notes_map = {}
        for np_ in src_notes:
            notes_counter += 1
            notes_map[int(re.search(r'\d+', Path(np_).name).group())] = notes_counter

        # --- slides: buduj mape stary->nowy numer ---
        src_slides = sorted(
            [n for n in names if re.match(r'ppt/slides/slide\d+\.xml$', n)],
            key=lambda x: int(re.search(r'\d+', Path(x).name).group()),
        )
        slide_num_map = {}
        for sp in src_slides:
            slide_counter += 1
            slide_num_map[int(re.search(r'\d+', Path(sp).name).group())] = slide_counter

        # --- kopiuj notes slides z uaktualnionymi rels ---
        for old_n, new_n in notes_map.items():
            arc = f'ppt/notesSlides/notesSlide{new_n}.xml'
            new_data[arc]  = sd[f'ppt/notesSlides/notesSlide{old_n}.xml']
            new_flags[arc] = (zipfile.ZIP_DEFLATED, 6)

            rels_src = f'ppt/notesSlides/_rels/notesSlide{old_n}.xml.rels'
            rels_arc = f'ppt/notesSlides/_rels/notesSlide{new_n}.xml.rels'
            if rels_src in sd:
                txt = sd[rels_src].decode('utf-8')
                txt = re.sub(
                    r'Target="\.\./slides/slide(\d+)\.xml"',
                    lambda m: f'Target="../slides/slide{slide_num_map.get(int(m.group(1)), int(m.group(1)))}.xml"',
                    txt,
                )
                new_data[rels_arc] = txt.encode('utf-8')
            else:
                new_data[rels_arc] = EMPTY_RELS
            new_flags[rels_arc] = (zipfile.ZIP_DEFLATED, 6)

            part = f'/ppt/notesSlides/notesSlide{new_n}.xml'
            if part not in existing_ct:
                existing_ct.add(part)
                e = etree.SubElement(ct_root, f'{{{CT_NS}}}Override')
                e.set('PartName', part)
                e.set('ContentType', NOTES_CT)

        # --- kopiuj slides z uaktualnionymi rels, rejestruj w presentation ---
        rename_map = {**media_map, **embed_map}
        for sp in src_slides:
            old_s = int(re.search(r'\d+', Path(sp).name).group())
            new_s = slide_num_map[old_s]
            new_name = f"slide{new_s}.xml"

            xml = sd[sp].decode('utf-8')
            for old, new in rename_map.items():
                xml = xml.replace(old, new)
            arc = f'ppt/slides/{new_name}'
            new_data[arc]  = xml.encode('utf-8')
            new_flags[arc] = (zipfile.ZIP_DEFLATED, 6)

            rels_src = f'ppt/slides/_rels/slide{old_s}.xml.rels'
            rels_arc = f'ppt/slides/_rels/{new_name}.rels'
            if rels_src in sd:
                txt = sd[rels_src].decode('utf-8')
                for old, new in rename_map.items():
                    txt = txt.replace(old, new)
                txt = re.sub(
                    r'Target="\.\./notesSlides/notesSlide(\d+)\.xml"',
                    lambda m: f'Target="../notesSlides/notesSlide{notes_map.get(int(m.group(1)), int(m.group(1)))}.xml"',
                    txt,
                )
                new_data[rels_arc] = txt.encode('utf-8')
            else:
                new_data[rels_arc] = EMPTY_RELS
            new_flags[rels_arc] = (zipfile.ZIP_DEFLATED, 6)

            rid = f"rId{nrid}"
            nrid += 1
            etree.SubElement(
                prs_rels_root, 'Relationship',
                Id=rid, Type=f'{REL_NS}/slide', Target=f'slides/{new_name}',
            )
            el = etree.SubElement(sldIdLst, f'{{{PRS_NS}}}sldId')
            el.set('id', str(next_id))
            el.set(f'{{{REL_NS}}}id', rid)
            next_id += 1

            part = f'/ppt/slides/{new_name}'
            if part not in existing_ct:
                existing_ct.add(part)
                e = etree.SubElement(ct_root, f'{{{CT_NS}}}Override')
                e.set('PartName', part)
                e.set('ContentType', SLIDE_CT)

    bd['ppt/presentation.xml']            = serialize(prs_root)
    bd['ppt/_rels/presentation.xml.rels'] = serialize(prs_rels_root)
    bd['[Content_Types].xml']             = serialize(ct_root)

    out = Path(output_file)
    with zipfile.ZipFile(out, 'w') as zout:
        written = set()
        for item in base['infos']:
            name = item.filename
            if name in REMOVE_FILES or name in written or name not in bd:
                continue
            info = zipfile.ZipInfo(name)
            info.compress_type = item.compress_type
            info.flag_bits     = item.flag_bits
            info.create_system = item.create_system
            info.external_attr = item.external_attr
            zout.writestr(info, bd[name])
            written.add(name)
        for arc_name, data in new_data.items():
            if arc_name in written:
                continue
            compress, flags = new_flags.get(arc_name, (zipfile.ZIP_DEFLATED, 6))
            info = zipfile.ZipInfo(arc_name)
            info.compress_type = compress
            info.flag_bits     = flags
            info.create_system = 0      # Windows
            info.external_attr = 0x20  # FILE_ATTRIBUTE_ARCHIVE
            zout.writestr(info, data)
            written.add(arc_name)

    print(f"\nGotowe: {out} ({slide_counter} slajdow, {notes_counter} notesSlides)")


def main():
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("Uzycie: python merge_pptx.py <folder> [plik_wyjsciowy.pptx]")
        sys.exit(1)
    folder = Path(sys.argv[1])
    output = sys.argv[2] if len(sys.argv) == 3 else "merged.pptx"
    if not folder.is_dir():
        print(f"Blad: '{folder}' nie jest folderem.")
        sys.exit(1)
    output_abs = str(Path(output).resolve())
    files = sorted(
        [f for f in folder.glob("*.pptx") if str(f.resolve()) != output_abs],
        key=lambda x: x.name,
    )
    if not files:
        print(f"Blad: brak plikow PPTX w folderze '{folder}'.")
        sys.exit(1)
    print(f"Znalezione pliki ({len(files)}):")
    for f in files:
        print(f"  {f.name}")
    print()
    merge_pptx([str(f) for f in files], output)


if __name__ == "__main__":
    main()
