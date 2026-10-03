from xml.etree import ElementTree as ET

from verified_docx_mcp.mutations import _preserve_table_properties
from verified_docx_mcp.projection import W_NS


def test_preserves_each_tables_properties_without_aliasing():
    def table(size):
        return ET.fromstring(f'<w:tbl xmlns:w="{W_NS}"><w:tblPr><w:tblBorders>'
                             f'<w:top w:val="single" w:sz="{size}"/>'
                             '</w:tblBorders></w:tblPr></w:tbl>')
    old = [table(4), table(12)]
    new = [table(1), table(1), table(1)]
    _preserve_table_properties(old, new)
    for i in range(2):
        assert ET.tostring(new[i]) == ET.tostring(old[i])
        assert new[i][0] is not old[i][0]
    assert new[2][0][0][0].get(f'{{{W_NS}}}sz') == '1'


def test_preserves_deliberately_borderless_table():
    old = ET.Element(f'{{{W_NS}}}tbl')
    new = ET.fromstring(f'<w:tbl xmlns:w="{W_NS}"><w:tblPr><w:tblBorders/></w:tblPr></w:tbl>')
    _preserve_table_properties([old], [new])
    assert new.find(f'{{{W_NS}}}tblPr') is None


def test_reordered_and_removed_tables_keep_their_own_properties():
    def table(label, size):
        return ET.fromstring(
            f'<w:tbl xmlns:w="{W_NS}"><w:tblPr><w:tblBorders>'
            f'<w:top w:val="single" w:sz="{size}"/></w:tblBorders></w:tblPr>'
            f'<w:tr><w:tc><w:p><w:r><w:t>{label}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>')

    def top_size(tbl):
        return tbl.find(f'{{{W_NS}}}tblPr')[0][0].get(f'{{{W_NS}}}sz')

    old = [table("A", 4), table("B", 12), table("C", 20)]
    reordered = [table("C", 1), table("A", 1), table("B", 1)]
    _preserve_table_properties(old, reordered)
    assert [top_size(t) for t in reordered] == ["20", "4", "12"]
    dropped = [table("B", 1), table("C", 1)]
    _preserve_table_properties(old, dropped)
    assert [top_size(t) for t in dropped] == ["12", "20"]
    inserted = [table("N", 1), table("A", 1)]
    _preserve_table_properties(old, inserted)
    assert top_size(inserted[1]) == "4"
    assert top_size(inserted[0]) in {"12", "20"}  # leftover pairing, never A's
