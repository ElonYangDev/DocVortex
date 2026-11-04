"""DOCX 表格跨页回归：单元格内 lastRenderedPageBreak 的行级切页与 demo 页数金标。"""

from __future__ import annotations

import re
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any, Callable, List, Tuple
from zipfile import ZIP_DEFLATED, ZipFile

from _span_test_utils import visible_content
from docx import Document
from lxml import etree

from docvortex.analyzers.native import DocxModel
from docvortex.schema import BlockType

from test_docx_page_breaks import _W, _build_docx, _lrbp_run, _page_break_run, _text_run

_MSOFFICE_DIR = Path(__file__).parents[2] / "demo" / "ms_office_docs"


def _cell(parent: etree._Element, text: str, *, lrbp: bool = False) -> None:
    """在行内追加一个单元格，可选在首 run 携带 lastRenderedPageBreak。"""

    tc = etree.SubElement(parent, f"{{{_W}}}tc")
    p = etree.SubElement(tc, f"{{{_W}}}p")
    if lrbp:
        _lrbp_run(p)
    _text_run(p, text)


def _add_row(parent: etree._Element, cells: list[str], lrbp_cell: int | None = None) -> None:
    """追加一行；lrbp_cell 指定在该列单元格首 run 注入渲染分页标记。"""

    tr = etree.SubElement(parent, f"{{{_W}}}tr")
    for index, text in enumerate(cells):
        _cell(tr, text, lrbp=index == lrbp_cell)


def _table(parent: etree._Element, rows: List[Tuple[list[str], int | None]]) -> None:
    """在正文 sectPr 前插入一个两列无嵌套表格。"""

    tbl = etree.Element(f"{{{_W}}}tbl")
    grid = etree.SubElement(tbl, f"{{{_W}}}tblGrid")
    for _ in list(rows)[0][0]:
        etree.SubElement(grid, f"{{{_W}}}gridCol")
    for cells, lrbp_cell in rows:
        _add_row(tbl, cells, lrbp_cell)
    sect_pr = parent.find(f"{{{_W}}}sectPr")
    if sect_pr is not None:
        parent.insert(list(parent).index(sect_pr), tbl)
    else:
        parent.append(tbl)


def _pages(data: bytes) -> List[list[dict[str, Any]]]:
    return DocxModel().predict(BytesIO(data))


def _page_text(page: list[dict[str, Any]]) -> str:
    """拼接页内块可见文字；TABLE 块（内容是 HTML）按标签剥除后取文本。"""

    from bs4 import BeautifulSoup

    parts: list[str] = []
    for block in page:
        content = block.get("content", "")
        if block.get("type") == BlockType.TABLE:
            parts.append(BeautifulSoup(str(content), "html.parser").get_text(strip=False))
        else:
            parts.append(visible_content(content))
    return "".join(parts)


def _tables(page: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [block for block in page if block.get("type") == BlockType.TABLE]


def test_table_row_lrbp_splits_table_across_pages() -> None:
    """验证行内渲染分页标记把表格按行边界切为两块并跨页产出。"""

    def build(body: etree._Element) -> None:
        _table(body, [(["alpha", "beta"], None), (["gamma", "delta"], 0), (["epsilon", "zeta"], None)])

    pages = _pages(_build_docx(build))
    assert len(pages) == 2
    assert _page_text(pages[0]) == "alphabeta"
    assert _page_text(pages[1]) == "gammadeltaepsilonzeta"
    assert [len(_tables(page)) for page in pages] == [1, 1]


def test_table_two_boundary_rows_produce_three_table_chunks() -> None:
    """验证同一表格两个续页行切出三块，连续换页不产生空页。"""

    def build(body: etree._Element) -> None:
        _table(body, [(["a1", "a2"], None), (["b1", "b2"], 0), (["c1", "c2"], 0)])

    pages = _pages(_build_docx(build))
    assert len(pages) == 3
    assert [_page_text(page) for page in pages] == ["a1a2", "b1b2", "c1c2"]
    assert all(len(_tables(page)) == 1 for page in pages)


def test_table_first_row_lrbp_starts_new_page() -> None:
    """验证表格首行含渲染分页标记时整表从新一页开始。"""

    def build(body: etree._Element) -> None:
        _body_page = body.makeelement(f"{{{_W}}}p", {})
        sect_pr = body.find(f"{{{_W}}}sectPr")
        if sect_pr is not None:
            body.insert(list(body).index(sect_pr), _body_page)
        else:
            body.append(_body_page)
        _text_run(_body_page, "alpha")
        _table(body, [(["a1", "a2"], 0), (["b1", "b2"], None)])

    pages = _pages(_build_docx(build))
    assert [_page_text(page) for page in pages] == ["alpha", "a1a2b1b2"]
    assert len(_tables(pages[1])) == 1


def test_table_first_row_lrbp_folds_with_preceding_explicit_break() -> None:
    """验证段尾显式分页符后的表格首行 lrbp 折叠为同一边界，不产生空页。"""

    def build(body: etree._Element) -> None:
        _body_page = body.makeelement(f"{{{_W}}}p", {})
        sect_pr = body.find(f"{{{_W}}}sectPr")
        if sect_pr is not None:
            body.insert(list(body).index(sect_pr), _body_page)
        else:
            body.append(_body_page)
        _text_run(_body_page, "alpha")
        _page_break_run(_body_page)
        _table(body, [(["a1", "a2"], 0), (["b1", "b2"], None)])

    pages = _pages(_build_docx(build))
    assert [_page_text(page) for page in pages] == ["alpha", "a1a2b1b2"]


def test_table_without_markers_keeps_single_block() -> None:
    """防误触发：无渲染分页标记的表格保持整块单页。"""

    def build(body: etree._Element) -> None:
        _table(body, [(["alpha", "beta"], None), (["gamma", "delta"], None)])

    pages = _pages(_build_docx(build))
    assert len(pages) == 1
    assert len(_tables(pages[0])) == 1


def test_demo_docx_analyzed_pages_match_word_saved_count_anchor() -> None:
    """验证 demo 实排页数金标（``docProps/app.xml`` 的 ``<Pages>``）已锁定。

    Word 保存时把渲染页数写入 ``docProps/app.xml``，本测试先核对样例文件内
    该金标自洽，再断言恢复页数不超过金标且等于当前已知恢复值。恢复值 13
    与 15 页实排的残差为长表单行跨页续排（行级切分不细分行内续排）。
    """

    sample = _MSOFFICE_DIR / "docx_01.docx"
    data = sample.read_bytes()
    with ZipFile(BytesIO(data)) as bundle:
        app_xml = bundle.read("docProps/app.xml").decode("utf-8")
    saved_pages = int(re.search(r"<Pages>(\d+)</Pages>", app_xml).group(1))
    assert saved_pages == 15

    pages = _pages(data)
    assert len(pages) == 13
    assert len(pages) <= saved_pages
    assert all(page for page in pages)
