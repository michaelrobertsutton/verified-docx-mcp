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
