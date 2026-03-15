from __future__ import annotations

import copy
import json
import re
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
import xml.etree.ElementTree as ET


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"

ET.register_namespace("a", A_NS)
ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("", PKG_NS)
ET.register_namespace("", CT_NS)
ET.register_namespace("", APP_NS)
ET.register_namespace("vt", VT_NS)
ET.register_namespace("p14", P14_NS)


PUNCT_RE = re.compile(r"[，。！？；：、,.!?;:\"'“”‘’（）()《》〈〉【】〔〕［］—…·/\\-]+")


@dataclass
class SlideRef:
    order_index: int
    slide_id: int
    rel_id: str
    slide_num: int
    texts: list[str]


def qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def read_slide_texts(slide_path: Path) -> list[str]:
    root = ET.parse(slide_path).getroot()
    return [t.text.strip() for t in root.findall(f".//{qn(A_NS, 't')}") if t.text and t.text.strip()]


def collapsed_text(texts: list[str]) -> str:
    return "".join(part.replace(" ", "") for part in texts)


def normalize_line(text: str) -> str:
    text = PUNCT_RE.sub("  ", text.strip())
    text = re.sub(r"\s+", " ", text)
    return text.replace(" ", "  ").strip()


def replace_text_box(sp: ET.Element, lines: list[str]) -> None:
    tx_body = sp.find(qn(P_NS, "txBody"))
    if tx_body is None:
        return

    paragraphs = tx_body.findall(qn(A_NS, "p"))
    template_p = paragraphs[0] if paragraphs else ET.Element(qn(A_NS, "p"))
    template_r = template_p.find(qn(A_NS, "r"))
    template_rpr = copy.deepcopy(template_r.find(qn(A_NS, "rPr"))) if template_r is not None else None
    template_end = copy.deepcopy(template_p.find(qn(A_NS, "endParaRPr")))

    for p in paragraphs:
        tx_body.remove(p)

    for line in lines if lines else [""]:
        p = ET.Element(qn(A_NS, "p"))
        if line:
            r = ET.SubElement(p, qn(A_NS, "r"))
            if template_rpr is not None:
                r.append(copy.deepcopy(template_rpr))
            else:
                ET.SubElement(r, qn(A_NS, "rPr"))
            t = ET.SubElement(r, qn(A_NS, "t"))
            t.text = line
        if template_end is not None:
            p.append(copy.deepcopy(template_end))
        tx_body.append(p)


def metadata_line(song: dict) -> str:
    return f"词：{song.get('lyricist', '').strip()}  曲：{song.get('composer', '').strip()}  出处：{song.get('source', '').strip()}"


def update_lyric_slide(slide_path: Path, song: dict, lines: list[str]) -> None:
    tree = ET.parse(slide_path)
    root = tree.getroot()
    title = song["title"].strip()
    details = metadata_line(song)
    for sp in root.findall(f".//{qn(P_NS, 'sp')}"):
        ph = sp.find(f"./{qn(P_NS, 'nvSpPr')}/{qn(P_NS, 'nvPr')}/{qn(P_NS, 'ph')}")
        if ph is None:
            continue
        ph_type = ph.attrib.get("type")
        ph_idx = ph.attrib.get("idx")
        if ph_type == "title":
            replace_text_box(sp, [title])
        elif ph_idx == "1":
            replace_text_box(sp, lines)
        elif ph_idx == "10":
            replace_text_box(sp, [details])
    tree.write(slide_path, encoding="UTF-8", xml_declaration=True)


def update_separator_slide(slide_path: Path) -> None:
    tree = ET.parse(slide_path)
    root = tree.getroot()
    for sp in root.findall(f".//{qn(P_NS, 'sp')}"):
        tx_body = sp.find(qn(P_NS, "txBody"))
        if tx_body is None:
            continue
        texts = [t.text for t in tx_body.findall(f".//{qn(A_NS, 't')}") if t.text]
        joined = "".join(texts).replace(" ", "")
        if joined == "新城":
            replace_text_box(sp, ["新城"])
    tree.write(slide_path, encoding="UTF-8", xml_declaration=True)


def add_override(content_types: Path, slide_num: int) -> None:
    tree = ET.parse(content_types)
    root = tree.getroot()
    part_name = f"/ppt/slides/slide{slide_num}.xml"
    if not any(el.attrib.get("PartName") == part_name for el in root.findall(qn(CT_NS, "Override"))):
        root.append(
            ET.Element(
                qn(CT_NS, "Override"),
                {
                    "PartName": part_name,
                    "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
                },
            )
        )
    tree.write(content_types, encoding="UTF-8", xml_declaration=True)


def update_app_xml(app_xml: Path, start_idx: int, end_idx: int, new_total_slides: int, new_titles: list[str]) -> None:
    tree = ET.parse(app_xml)
    root = tree.getroot()

    slides = root.find(qn(APP_NS, "Slides"))
    old_total_slides = int(slides.text) if slides is not None and slides.text is not None else new_total_slides
    if slides is not None:
        slides.text = str(new_total_slides)

    heading_pairs = root.find(f"./{qn(APP_NS, 'HeadingPairs')}/{qn(VT_NS, 'vector')}")
    if heading_pairs is not None:
        variants = heading_pairs.findall(qn(VT_NS, "variant"))
        if len(variants) >= 6:
            count_el = variants[5].find(qn(VT_NS, "i4"))
            if count_el is not None:
                count_el.text = str(new_total_slides)

    titles_vector = root.find(f"./{qn(APP_NS, 'TitlesOfParts')}/{qn(VT_NS, 'vector')}")
    if titles_vector is not None:
        existing = titles_vector.findall(qn(VT_NS, "lpstr"))
        non_slide_count = len(existing) - old_total_slides
        slide_titles = existing[non_slide_count:]
        fixed_prefix = existing[:non_slide_count] + slide_titles[:start_idx]
        fixed_suffix = slide_titles[end_idx + 1 :]
        for child in list(titles_vector):
            titles_vector.remove(child)
        inserted = [ET.Element(qn(VT_NS, "lpstr")) for _ in new_titles]
        for el, title in zip(inserted, new_titles):
            el.text = title
        for el in fixed_prefix + inserted + fixed_suffix:
            titles_vector.append(el)
        titles_vector.attrib["size"] = str(len(list(titles_vector)))

    tree.write(app_xml, encoding="UTF-8", xml_declaration=True)


def load_song_blocks(song_json: Path) -> list[dict]:
    data = json.loads(song_json.read_text())
    songs = data["songs"] if isinstance(data, dict) and "songs" in data else data
    normalized = []
    for song in songs:
        blocks = []
        for block in song["blocks"]:
            label = block.get("label", "")
            lines = [normalize_line(line) for line in block["lines"] if line.strip()]
            blocks.append({"label": label, "lines": lines})
        normalized.append(
            {
                "title": song["title"].strip(),
                "lyricist": song.get("lyricist", "").strip(),
                "composer": song.get("composer", "").strip(),
                "source": song.get("source", "").strip(),
                "blocks": blocks,
            }
        )
    return normalized


def parse_slides(presentation_xml: Path, rels_xml: Path, slides_dir: Path) -> list[SlideRef]:
    rel_root = ET.parse(rels_xml).getroot()
    rel_map = {
        rel.attrib["Id"]: int(Path(rel.attrib["Target"]).stem.replace("slide", ""))
        for rel in rel_root.findall(qn(PKG_NS, "Relationship"))
        if rel.attrib.get("Type", "").endswith("/slide")
    }

    pres_root = ET.parse(presentation_xml).getroot()
    slides = []
    for idx, el in enumerate(pres_root.find(qn(P_NS, "sldIdLst")).findall(qn(P_NS, "sldId"))):
        rel_id = el.attrib[qn(R_NS, "id")]
        slide_num = rel_map[rel_id]
        texts = read_slide_texts(slides_dir / f"slide{slide_num}.xml")
        slides.append(SlideRef(idx, int(el.attrib["id"]), rel_id, slide_num, texts))
    return slides


def find_song_section(slides: list[SlideRef]) -> tuple[int, int, int]:
    creed_end_idx = next(i for i, slide in enumerate(slides) if "我信永生" in collapsed_text(slide.texts))
    lord_prayer_idx = next(i for i, slide in enumerate(slides) if "主禱文" in collapsed_text(slide.texts))
    separator_before_lord_prayer = max(i for i in range(creed_end_idx + 1, lord_prayer_idx) if collapsed_text(slides[i].texts) == "新城")
    separator_indices = [i for i in range(creed_end_idx + 1, separator_before_lord_prayer) if collapsed_text(slides[i].texts) == "新城"]
    if not separator_indices:
        raise ValueError("Could not identify the existing song section.")
    first_song_separator = separator_indices[0]
    lyric_template_idx = next(i for i in range(first_song_separator + 1, separator_before_lord_prayer) if collapsed_text(slides[i].texts) not in {"", "新城"})
    return first_song_separator, separator_before_lord_prayer, lyric_template_idx


def rebuild_presentation_list(
    presentation_xml: Path,
    start_idx: int,
    end_idx: int,
    new_entries: list[tuple[int, str]],
) -> int:
    tree = ET.parse(presentation_xml)
    root = tree.getroot()
    sld_id_lst = root.find(qn(P_NS, "sldIdLst"))
    assert sld_id_lst is not None
    children = list(sld_id_lst)
    kept = children[:start_idx] + children[end_idx + 1 :]
    for child in list(sld_id_lst):
        sld_id_lst.remove(child)
    for child in kept[:start_idx]:
        sld_id_lst.append(child)
    for slide_id, rel_id in new_entries:
        sld_id_lst.append(ET.Element(qn(P_NS, "sldId"), {"id": str(slide_id), qn(R_NS, "id"): rel_id}))
    for child in kept[start_idx:]:
        sld_id_lst.append(child)

    section_list = root.find(f"./{qn(P_NS, 'extLst')}/{qn(P_NS, 'ext')}/{qn(P14_NS, 'sectionLst')}/{qn(P14_NS, 'section')}/{qn(P14_NS, 'sldIdLst')}")
    if section_list is not None:
        sec_children = list(section_list)
        sec_kept = sec_children[:start_idx] + sec_children[end_idx + 1 :]
        for child in list(section_list):
            section_list.remove(child)
        for child in sec_kept[:start_idx]:
            section_list.append(child)
        for slide_id, _ in new_entries:
            section_list.append(ET.Element(qn(P14_NS, "sldId"), {"id": str(slide_id)}))
        for child in sec_kept[start_idx:]:
            section_list.append(child)

    tree.write(presentation_xml, encoding="UTF-8", xml_declaration=True)
    return len(list(sld_id_lst))


def update_relationships(rels_xml: Path, new_rel_entries: list[tuple[str, int]]) -> None:
    tree = ET.parse(rels_xml)
    root = tree.getroot()
    existing = {rel.attrib["Id"] for rel in root.findall(qn(PKG_NS, "Relationship"))}
    for rel_id, slide_num in new_rel_entries:
        if rel_id in existing:
            continue
        root.append(
            ET.Element(
                qn(PKG_NS, "Relationship"),
                {
                    "Id": rel_id,
                    "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                    "Target": f"slides/slide{slide_num}.xml",
                },
            )
        )
    tree.write(rels_xml, encoding="UTF-8", xml_declaration=True)


def build_insert_sequence(songs: list[dict]) -> list[tuple[str, dict | None, list[str]]]:
    sequence: list[tuple[str, dict | None, list[str]]] = [("separator", None, ["新城"])]
    for song_index, song in enumerate(songs):
        for block in song["blocks"]:
            sequence.append(("lyric", song, block["lines"]))
        if song_index != len(songs) - 1:
            sequence.append(("separator", None, ["新城"]))
    sequence.append(("separator", None, ["新城"]))
    return sequence


def write_output_pptx(source_pptx: Path, output_pptx: Path, songs_json: Path) -> None:
    songs = load_song_blocks(songs_json)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        with zipfile.ZipFile(source_pptx) as zf:
            zf.extractall(tmp)

        ppt_dir = tmp / "ppt"
        slides_dir = ppt_dir / "slides"
        slide_rels_dir = slides_dir / "_rels"
        presentation_xml = ppt_dir / "presentation.xml"
        rels_xml = ppt_dir / "_rels" / "presentation.xml.rels"

        slides = parse_slides(presentation_xml, rels_xml, slides_dir)
        start_idx, end_idx, lyric_template_idx = find_song_section(slides)
        separator_template_idx = start_idx

        lyric_template_num = slides[lyric_template_idx].slide_num
        separator_template_num = slides[separator_template_idx].slide_num
        lyric_template_slide = slides_dir / f"slide{lyric_template_num}.xml"
        lyric_template_rels = slide_rels_dir / f"slide{lyric_template_num}.xml.rels"
        separator_template_slide = slides_dir / f"slide{separator_template_num}.xml"
        separator_template_rels = slide_rels_dir / f"slide{separator_template_num}.xml.rels"

        existing_slide_nums = [slide.slide_num for slide in slides]
        next_slide_num = max(existing_slide_nums) + 1
        next_slide_id = max(slide.slide_id for slide in slides) + 1
        rel_root = ET.parse(rels_xml).getroot()
        next_rel_num = max(
            int(rel.attrib["Id"].replace("rId", ""))
            for rel in rel_root.findall(qn(PKG_NS, "Relationship"))
            if rel.attrib.get("Id", "").startswith("rId")
        ) + 1

        sequence = build_insert_sequence(songs)
        rel_entries: list[tuple[str, int]] = []
        pres_entries: list[tuple[int, str]] = []
        new_titles: list[str] = []

        for kind, song, lines in sequence:
            slide_num = next_slide_num
            slide_id = next_slide_id
            rel_id = f"rId{next_rel_num}"
            next_slide_num += 1
            next_slide_id += 1
            next_rel_num += 1

            target_slide = slides_dir / f"slide{slide_num}.xml"
            target_rels = slide_rels_dir / f"slide{slide_num}.xml.rels"
            if kind == "separator":
                shutil.copy2(separator_template_slide, target_slide)
                shutil.copy2(separator_template_rels, target_rels)
                update_separator_slide(target_slide)
                new_titles.append("PowerPoint Presentation")
            else:
                shutil.copy2(lyric_template_slide, target_slide)
                shutil.copy2(lyric_template_rels, target_rels)
                assert song is not None
                update_lyric_slide(target_slide, song, lines)
                new_titles.append(song["title"])

            add_override(tmp / "[Content_Types].xml", slide_num)
            rel_entries.append((rel_id, slide_num))
            pres_entries.append((slide_id, rel_id))

        update_relationships(rels_xml, rel_entries)
        total_slides = rebuild_presentation_list(presentation_xml, start_idx, end_idx, pres_entries)
        update_app_xml(tmp / "docProps" / "app.xml", start_idx, end_idx, total_slides, new_titles)

        with zipfile.ZipFile(output_pptx, "w", zipfile.ZIP_DEFLATED) as zf:
            for path in sorted(tmp.rglob("*")):
                if path.is_file():
                    zf.write(path, path.relative_to(tmp))


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Replace the worship-song section of a service PPTX.")
    parser.add_argument("source_pptx", type=Path)
    parser.add_argument("songs_json", type=Path)
    parser.add_argument("output_pptx", type=Path)
    args = parser.parse_args()

    write_output_pptx(args.source_pptx, args.output_pptx, args.songs_json)
    print(args.output_pptx)


if __name__ == "__main__":
    main()
