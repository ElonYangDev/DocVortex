"""DOCX 原生分页语义回归：显式分页符、渲染分页、pageBreakBefore 及其关闭态。"""

from __future__ import annotations

from io import BytesIO
from typing import Any, Callable, List
from zipfile import ZIP_DEFLATED, ZipFile

from _span_test_utils import visible_content
from docx import Document
from lxml import etree

from docvortex.analyzers.native import DocxModel
from docvortex.schema import BlockType

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _text_run(paragraph: etree._Element, text: str) -> etree._Element:
    """在段落末尾追加一个纯文本 run。"""

    run = etree.SubElement(paragraph, f"{{{_W}}}r")
    node = etree.SubElement(run, f"{{{_W}}}t")
    node.text = text
    node.set(_XML_SPACE, "preserve")
    return run


def _lrbp_run(paragraph: etree._Element) -> etree._Element:
    """在段落末尾追加一个 lastRenderedPageBreak（最后一次 Word 渲染记录的分页）。"""

    run = etree.SubElement(paragraph, f"{{{_W}}}r")
    etree.SubElement(run, f"{{{_W}}}lastRenderedPageBreak")
    return run


def _page_break_run(paragraph: etree._Element) -> etree._Element:
    """在段落末尾追加一个显式分页符 ``w:br w:type="page"``。"""

    run = etree.SubElement(paragraph, f"{{{_W}}}r")
    br = etree.SubElement(run, f"{{{_W}}}br")
    br.set(f"{{{_W}}}type", "page")
    return br


def _page_break_before(paragraph: etree._Element, enabled: bool | None) -> None:
    """设置段首分页标志；``None`` 保持未设置，``False`` 写出显式关闭态（WPS 常见）。"""

    if enabled is None:
        return
    p_pr = etree.Element(f"{{{_W}}}pPr")
    flag = etree.SubElement(p_pr, f"{{{_W}}}pageBreakBefore")
    if not enabled:
        flag.set(f"{{{_W}}}val", "0")
    paragraph.insert(0, p_pr)


def _build_page(document: etree._Element, build: Callable[[etree._Element], None]) -> None:
    """在正文 sectPr 之前创建一个由 build 定制的段落。"""

    p = document.makeelement(f"{{{_W}}}p", {})
    build(p)
    sect_pr = document.find(f"{{{_W}}}sectPr")
    if sect_pr is not None:
        document.insert(list(document).index(sect_pr), p)
    else:
        document.append(p)


def _build_docx(body: Callable[[etree._Element], None], *, normal_page_break_before: bool = False) -> bytes:
    """基于 python-docx 默认模板构造自定义正文的最小 DOCX。"""

    document = Document()
    buffer = BytesIO()
    document.save(buffer)
    with ZipFile(BytesIO(buffer.getvalue())) as source:
        members = {name: source.read(name) for name in source.namelist()}

    root = etree.fromstring(members["word/document.xml"])
    body_element = root.find(f"{{{_W}}}body")
    assert body_element is not None
    for child in list(body_element):
        if etree.QName(child.tag).localname == "p":
            body_element.remove(child)
    body(body_element)
    members["word/document.xml"] = etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )
    if normal_page_break_before:
        styles = etree.fromstring(members["word/styles.xml"])
        normal = styles.find(f".//{{{_W}}}style[@{{{_W}}}styleId='Normal']")
        assert normal is not None
        p_pr = normal.find(f"{{{_W}}}pPr")
        if p_pr is None:
            p_pr = etree.Element(f"{{{_W}}}pPr")
            normal.insert(0, p_pr)
        p_pr.append(etree.SubElement(p_pr, f"{{{_W}}}pageBreakBefore"))
        members["word/styles.xml"] = etree.tostring(
            styles,
            xml_declaration=True,
            encoding="UTF-8",
            standalone=True,
        )

    output = BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as target:
        for name, payload in members.items():
            target.writestr(name, payload)
    return output.getvalue()


def _pages(data: bytes) -> List[list[dict[str, Any]]]:
    return DocxModel().predict(BytesIO(data))


def _page_text(page: list[dict[str, Any]]) -> str:
    return "".join(visible_content(block.get("content", "")) for block in page)


def _assert_pages(pages: List[list[dict[str, Any]]], expected_texts: list[str]) -> None:
    assert len(pages) == len(expected_texts), [(_page_text(page), page) for page in pages]
    for page, expected in zip(pages, expected_texts):
        assert _page_text(page) == expected


def test_docx_rendered_page_break_splits_paragraph_across_pages() -> None:
    """验证 lastRenderedPageBreak 按可见字符偏移切分段落，跨页产出两个文本块。"""

    def build(body: etree._Element) -> None:
        _build_page(body, lambda p: (_text_run(p, "alpha"), _lrbp_run(p), _text_run(p, "beta")))

    pages = _pages(_build_docx(build))
    _assert_pages(pages, ["alpha", "beta"])
    for page in pages:
        assert [block["type"] for block in page] == [BlockType.TEXT]


def test_docx_explicit_page_break_splits_paragraph_across_pages() -> None:
    """验证 w:br w:type="page" 与渲染分页符同样按偏移切分段落内容。"""

    def build(body: etree._Element) -> None:
        _build_page(body, lambda p: (_text_run(p, "alpha"), _page_break_run(p), _text_run(p, "beta")))

    pages = _pages(_build_docx(build))
    _assert_pages(pages, ["alpha", "beta"])


def test_docx_trailing_page_break_starts_next_page() -> None:
    """验证段尾显式分页符让下一页承接下一段落，且同页不重复产出。"""

    def build(body: etree._Element) -> None:
        _build_page(body, lambda p: (_text_run(p, "alpha"), _page_break_run(p)))
        _build_page(body, lambda p: (_text_run(p, "beta"),))

    pages = _pages(_build_docx(build))
    _assert_pages(pages, ["alpha", "beta"])


def test_docx_page_break_before_starts_new_page() -> None:
    """验证 pPr/w:pageBreakBefore 让段落从新一页开始。"""

    def build(body: etree._Element) -> None:
        _build_page(body, lambda p: (_text_run(p, "alpha"),))
        _build_page(body, lambda p: (_page_break_before(p, True), _text_run(p, "beta")))

    pages = _pages(_build_docx(build))
    _assert_pages(pages, ["alpha", "beta"])


def test_docx_page_break_before_explicit_off_state_is_ignored() -> None:
    """验证 WPS 常见 ``w:pageBreakBefore w:val="0"`` 显式关闭态不触发段首分页。"""

    def build(body: etree._Element) -> None:
        _build_page(body, lambda p: (_text_run(p, "alpha"),))
        _build_page(body, lambda p: (_page_break_before(p, False), _text_run(p, "beta")))

    pages = _pages(_build_docx(build))
    _assert_pages(pages, ["alphabeta"])
    assert [block["type"] for block in pages[0]] == [BlockType.TEXT, BlockType.TEXT]


def test_docx_page_break_before_inherited_from_style_chain() -> None:
    """验证 Normal 样式声明 pageBreakBefore 时未直接设置的段落也段首分页。"""

    def build(body: etree._Element) -> None:
        _build_page(body, lambda p: (_text_run(p, "alpha"),))
        _build_page(body, lambda p: (_text_run(p, "beta"),))

    pages = _pages(_build_docx(build, normal_page_break_before=True))
    _assert_pages(pages, ["alpha", "beta"])


def test_docx_consecutive_page_breaks_collapse_into_single_page() -> None:
    """验证同一边界处的显式分页符 + 下一段渲染分页符折叠为一页，不产生空页。"""

    def build(body: etree._Element) -> None:
        _build_page(body, lambda p: (_text_run(p, "alpha"), _page_break_run(p)))
        _build_page(body, lambda p: (_lrbp_run(p), _text_run(p, "beta")))

    pages = _pages(_build_docx(build))
    _assert_pages(pages, ["alpha", "beta"])


def test_docx_doc_without_markers_keeps_single_page() -> None:
    """防误触发：没有任何分页标记的文档页数保持一页。"""

    def build(body: etree._Element) -> None:
        _build_page(body, lambda p: (_text_run(p, "alpha"),))
        _build_page(body, lambda p: (_text_run(p, "beta"),))

    pages = _pages(_build_docx(build))
    assert len(pages) == 1
    assert _page_text(pages[0]) == "alphabeta"
