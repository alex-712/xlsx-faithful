"""核心功能测试。用 openpyxl 造模板,只是为了不依赖外部文件。

⚠️ 真正的保真度测试在 test_fidelity.py,那个必须用真实业务文件。
"""

import zipfile
from pathlib import Path

import pytest

from xlsx_faithful import XlsxFaithfulError, list_parts, write_cells
from xlsx_faithful.core import _col_to_num, _split_ref

openpyxl = pytest.importorskip("openpyxl")


@pytest.fixture
def template(tmp_path: Path) -> Path:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "标题"
    ws["B3"] = "原值"
    ws["D5"] = 0
    ws["A10"] = "锚点"
    p = tmp_path / "t.xlsx"
    wb.save(p)
    return p


def _read(path: Path, ref: str):
    wb = openpyxl.load_workbook(path)
    return wb.active[ref].value


# ---------- 辅助函数 ----------

def test_split_ref():
    assert _split_ref("B3") == ("B", 3)
    assert _split_ref("aa12") == ("AA", 12)
    with pytest.raises(XlsxFaithfulError):
        _split_ref("3B")


def test_col_to_num():
    assert _col_to_num("A") == 1
    assert _col_to_num("Z") == 26
    assert _col_to_num("AA") == 27


# ---------- 写值 ----------

def test_overwrite_existing_cell(template, tmp_path):
    out = tmp_path / "out.xlsx"
    write_cells(template, out, {"B3": "新值"})
    assert _read(out, "B3") == "新值"


def test_write_types(template, tmp_path):
    out = tmp_path / "out.xlsx"
    write_cells(src=template, dst=out, cells={
        "B3": "文本", "D5": 1200, "E5": 3.5, "F5": True, "A1": None,
    })
    assert _read(out, "B3") == "文本"
    assert _read(out, "D5") == 1200
    assert _read(out, "E5") == 3.5
    assert _read(out, "F5") is True
    assert _read(out, "A1") is None


def test_create_cell_in_existing_row(template, tmp_path):
    """B3 所在行已存在,往同一行写个新单元格。"""
    out = tmp_path / "out.xlsx"
    write_cells(template, out, {"F3": "新单元格"})
    assert _read(out, "F3") == "新单元格"
    assert _read(out, "B3") == "原值"  # 原有的没被动


def test_create_new_row(template, tmp_path):
    out = tmp_path / "out.xlsx"
    write_cells(template, out, {"C99": "很远的一行"})
    assert _read(out, "C99") == "很远的一行"
    assert _read(out, "A10") == "锚点"


def test_cell_order_preserved(template, tmp_path):
    """列顺序错了 Excel 会报文件损坏,这里直接查 XML。"""
    out = tmp_path / "out.xlsx"
    write_cells(template, out, {"Z3": "z", "C3": "c", "A3": "a"})

    with zipfile.ZipFile(out) as z:
        xml = z.read("xl/worksheets/sheet1.xml").decode()

    row3 = xml.split('<row r="3"')[1].split("</row>")[0]
    positions = [row3.find(f'r="{ref}"') for ref in ("A3", "B3", "C3", "Z3")]
    assert all(p >= 0 for p in positions), "单元格缺失"
    assert positions == sorted(positions), "列顺序错乱,Excel 会拒绝打开"


def test_string_uses_inline_not_shared_pool(template, tmp_path):
    """字符串走 inlineStr,不动 sharedStrings。"""
    out = tmp_path / "out.xlsx"
    write_cells(template, out, {"B3": "inline 测试"})

    with zipfile.ZipFile(template) as a, zipfile.ZipFile(out) as b:
        sheet = b.read("xl/worksheets/sheet1.xml").decode()
        name = "xl/sharedStrings.xml"
        has_pool = name in a.namelist()
        shared_before = a.read(name) if has_pool else None
        shared_after = b.read(name) if name in b.namelist() else None

    assert 't="inlineStr"' in sheet
    assert shared_before == shared_after, "共享池被改动了"


def test_leading_space_preserved(template, tmp_path):
    out = tmp_path / "out.xlsx"
    write_cells(template, out, {"B3": "  带空格  "})
    assert _read(out, "B3") == "  带空格  "


# ---------- 保真度(基础版) ----------

def test_all_parts_preserved(template, tmp_path):
    out = tmp_path / "out.xlsx"
    write_cells(template, out, {"B3": "x"})
    assert list_parts(template) == list_parts(out)


def test_untouched_parts_are_byte_identical(template, tmp_path):
    """除目标工作表外,每个零件都必须逐字节相同。"""
    out = tmp_path / "out.xlsx"
    write_cells(template, out, {"B3": "x"})

    target = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(template) as a, zipfile.ZipFile(out) as b:
        for name in a.namelist():
            if name == target:
                continue
            assert a.read(name) == b.read(name), f"{name} 被改动了"


# ---------- 错误处理 ----------

def test_missing_template(tmp_path):
    with pytest.raises(XlsxFaithfulError, match="不存在"):
        write_cells(tmp_path / "nope.xlsx", tmp_path / "o.xlsx", {"A1": 1})


def test_missing_sheet(template, tmp_path):
    with pytest.raises(XlsxFaithfulError, match="找不到"):
        write_cells(template, tmp_path / "o.xlsx", {"A1": 1},
                    sheet="xl/worksheets/sheet99.xml")


def test_bad_ref(template, tmp_path):
    with pytest.raises(XlsxFaithfulError):
        write_cells(template, tmp_path / "o.xlsx", {"不是引用": 1})


# ---------- 工作表 XML 层面的保真(防命名空间前缀被改写)----------

def test_only_sheetdata_region_rewritten(template, tmp_path):
    """除 <sheetData> 外,工作表 XML 的其余字节必须完全不变。

    回归测试:早期版本用 ElementTree 重新序列化整个工作表,
    把 xmlns:r 改成了 xmlns:ns1,导致 <drawing r:id> 变成 <drawing ns1:id>,
    很多解析器直接打不开文件。
    """
    from xlsx_faithful.core import _SHEETDATA_RE

    out = tmp_path / "out.xlsx"
    write_cells(template, out, {"B3": "x"})

    name = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(template) as a, zipfile.ZipFile(out) as b:
        xa, xb = a.read(name), b.read(name)

    strip = lambda x: _SHEETDATA_RE.sub(b"@@SHEETDATA@@", x)
    assert strip(xa) == strip(xb), "sheetData 之外的部分被改动了"


def test_namespace_prefixes_unchanged(template, tmp_path):
    """根元素上的命名空间声明必须原样保留。"""
    out = tmp_path / "out.xlsx"
    write_cells(template, out, {"B3": "x"})

    name = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(template) as a, zipfile.ZipFile(out) as b:
        head_a = a.read(name)[:400]
        head_b = b.read(name)[:400]

    import re as _re
    ns = lambda x: sorted(_re.findall(rb'xmlns:?[a-z0-9]*="[^"]+"', x))
    assert ns(head_a) == ns(head_b), "命名空间声明被改写了"
