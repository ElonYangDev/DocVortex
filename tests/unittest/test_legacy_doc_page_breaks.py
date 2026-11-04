"""DOC（Word 97–2003）原生解析的显式分页符（0x0C 段落终止符）语义回归。"""

from __future__ import annotations

from io import BytesIO

from _legacy_doc_test_utils import build_doc
from _span_test_utils import inline
from docvortex.analyzers.native import DocModel
from docvortex.schema import BlockType


def _pages(text: str):
    return DocModel().predict(BytesIO(build_doc(text)))


def test_doc_page_break_splits_section_across_pages() -> None:
    """验证段内 0x0C 分页符把段落切到下一页，段前内容保留在上一页。"""

    pages = _pages("alpha\rpara2\x0cpara3\r")

    assert pages == [
        [
            {"type": BlockType.TEXT, "content": inline("alpha")},
            {"type": BlockType.TEXT, "content": inline("para2")},
        ],
        [{"type": BlockType.TEXT, "content": inline("para3")}],
    ]


def test_doc_leading_page_break_does_not_add_empty_page() -> None:
    """验证段首 0x0C 不产生额外空页：首段直接落在 section 第一页。"""

    pages = _pages("\x0cFirst\r")

    assert pages == [[{"type": BlockType.TEXT, "content": inline("First")}]]


def test_doc_trailing_page_break_does_not_add_empty_page() -> None:
    """验证段尾 0x0C（后接空段落）不产生尾随空页。"""

    pages = _pages("First\x0c\r")

    assert pages == [[{"type": BlockType.TEXT, "content": inline("First")}]]


def test_doc_consecutive_page_breaks_collapse_into_single_split() -> None:
    """验证连续多个 0x0C 分页符折叠为一页，不产生多余空页。"""

    pages = _pages("A\r\x0c\x0cB\r")

    assert pages == [
        [{"type": BlockType.TEXT, "content": inline("A")}],
        [{"type": BlockType.TEXT, "content": inline("B")}],
    ]
