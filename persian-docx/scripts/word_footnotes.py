"""Add a native English term footnote immediately after a selected Word run."""
from docx.opc.constants import CONTENT_TYPE, RELATIONSHIP_TYPE
from docx.opc.part import XmlPart
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def _el(name, **attrs):
    node = OxmlElement("w:" + name)
    for key, value in attrs.items():
        node.set(qn("w:" + key), str(value))
    return node


def _set_page_numbering(props):
    for name, value, successors in (
        ("numFmt", "decimal", ("w:numStart", "w:numRestart", "w:footnote")),
        ("numStart", "1", ("w:numRestart", "w:footnote")),
        ("numRestart", "eachPage", ("w:footnote",)),
    ):
        node = props.find(qn("w:" + name))
        if node is None:
            node = _el(name)
            following = {qn(tag) for tag in successors}
            next_node = next((child for child in props if child.tag in following), None)
            if next_node is None:
                props.append(node)
            else:
                next_node.addprevious(node)
        node.set(qn("w:val"), value)


def _restart_footnotes_each_page(main):
    settings = main.settings.element
    props = settings.find(qn("w:footnotePr"))
    if props is None:
        props = _el("footnotePr")
        settings.insert_element_before(props, "w:endnotePr", "w:compat", "w:docVars",
                                       "w:rsids", "m:mathPr", "w:themeFontLang", "w:clrSchemeMapping")
    _set_page_numbering(props)
    # Word reads per-page restarts from section properties, even with a document default.
    for section in main.document.sections:
        section_props = section._sectPr
        override = section_props.find(qn("w:footnotePr"))
        if override is None:
            override = _el("footnotePr")
            first_setting = next((child for child in section_props if child.tag not in
                                  (qn("w:headerReference"), qn("w:footerReference"))), None)
            if first_setting is None:
                section_props.append(override)
            else:
                first_setting.addprevious(override)
        _set_page_numbering(override)


def add_english_footnote(anchor_run, english, *, restart_each_page=True):
    """Insert an English term note; leave template numbering alone when opted out."""
    if not english.strip() or any(ch in english for ch in "\r\n\t"):
        raise ValueError("Provide a nonempty single-line English term")
    main = anchor_run.part
    if restart_each_page:
        _restart_footnotes_each_page(main)
    try:
        part = main.part_related_by(RELATIONSHIP_TYPE.FOOTNOTES)
        if not isinstance(part, XmlPart):
            part = XmlPart.load(part.partname, part.content_type, part.blob, part.package)
            for rel in main.rels.values():
                if rel.reltype == RELATIONSHIP_TYPE.FOOTNOTES:
                    rel._target = part
                    break
    except KeyError:
        root = _el("footnotes")
        for ident, kind in [(-1, "separator"), (0, "continuationSeparator")]:
            sep = _el("footnote", id=ident, type=kind)
            p = _el("p")
            r = _el("r")
            r.append(_el(kind))
            p.append(r)
            sep.append(p)
            root.append(sep)
        name = main.package.next_partname("/word/footnotes%d.xml")
        part = XmlPart(name, CONTENT_TYPE.WML_FOOTNOTES, root, main.package)
        main.relate_to(part, RELATIONSHIP_TYPE.FOOTNOTES)

    ident = 1 + max([0] + [int(n.get(qn("w:id"))) for n in part.element])
    note = _el("footnote", id=ident)
    p = _el("p")
    pp = _el("pPr")
    pp.append(_el("bidi", val=0))
    pp.append(_el("jc", val="left"))
    p.append(pp)
    for is_number in (True, False):
        r = _el("r")
        rp = _el("rPr")
        rp.append(_el("rFonts", ascii="Arial", hAnsi="Arial", cs="Arial"))
        rp.append(_el("sz", val=19))
        rp.append(_el("szCs", val=19))
        if is_number:
            rp.append(_el("vertAlign", val="superscript"))
        rp.append(_el("rtl", val=0))
        rp.append(_el("lang", val="en-US"))
        r.append(rp)
        if is_number:
            r.append(_el("footnoteRef"))
        else:
            t = _el("t")
            t.set(qn("xml:space"), "preserve")
            t.text = "  " + english.strip()
            r.append(t)
        p.append(r)
    note.append(p)
    part.element.append(note)

    marker = _el("r")
    rp = _el("rPr")
    rp.append(_el("vertAlign", val="superscript"))
    marker.append(rp)
    marker.append(_el("footnoteReference", id=ident))
    anchor_run._r.addnext(marker)
    return ident
