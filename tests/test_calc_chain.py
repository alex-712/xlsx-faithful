"""覆盖公式单元格时,calcChain 必须跟着走。

xl/calcChain.xml 记录着「哪些格子有公式、按什么顺序算」。写值会把
单元格里的 <f> 清掉(不清的话 Excel 重算时会把写进去的值盖掉),
但如果 calcChain 里那条记录还在,就成了指向「已经没有公式的格子」的
悬空引用 —— Excel 打开时可能提示「发现不可读取的内容」并自动修复。

真实业务模板几乎都带这个零件,所以这不是边缘情况。
"""

import re
import zipfile
from pathlib import Path

import pytest

from xlsx_faithful import list_parts, write_cells

FIXTURE = Path(__file__).parent / "fixtures" / "real_template.xlsx"

pytestmark = pytest.mark.skipif(
    not FIXTURE.is_file(),
    reason="缺少 tests/fixtures/real_template.xlsx,详见该目录 README",
)

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _calc_chain_refs(path) -> set[str]:
    with zipfile.ZipFile(path) as z:
        if "xl/calcChain.xml" not in z.namelist():
            return set()
        return set(re.findall(r'<c r="([A-Z]+\d+)"', z.read("xl/calcChain.xml").decode()))


def _formula_cells(path) -> set[str]:
    from xml.etree import ElementTree as ET
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    return {c.get("r") for c in root.iter(f"{NS}c") if c.find(f"{NS}f") is not None}


def test_fixture_actually_has_formulas():
    """前提:fixture 必须真的带公式和 calcChain,否则下面的测试什么都证明不了。"""
    assert _formula_cells(FIXTURE), "fixture 里没有公式单元格,测试无效"
    assert _calc_chain_refs(FIXTURE), "fixture 里没有 calcChain,测试无效"


def test_overwriting_formula_cell_drops_its_calc_chain_entry(tmp_path):
    victim = sorted(_formula_cells(FIXTURE))[0]
    out = tmp_path / "out.xlsx"
    write_cells(FIXTURE, out, {victim: "覆盖公式"})

    assert victim not in _formula_cells(out), "公式没被清掉"
    assert victim not in _calc_chain_refs(out), (
        f"{victim} 的公式已删,但 calcChain 里还留着记录(悬空引用)"
    )


def test_untouched_formulas_keep_their_calc_chain_entries(tmp_path):
    """只删被覆盖的那条,别把整个 calcChain 清空。"""
    formulas = sorted(_formula_cells(FIXTURE))
    assert len(formulas) >= 2, "fixture 需要至少两个公式格才能验证这一条"
    victim, survivor = formulas[0], formulas[1]

    out = tmp_path / "out.xlsx"
    write_cells(FIXTURE, out, {victim: "覆盖"})

    assert survivor in _calc_chain_refs(out), f"{survivor} 没被碰,它的 calcChain 记录不该消失"


def test_calc_chain_part_dropped_when_it_becomes_empty(tmp_path):
    """所有公式都被覆盖后,空的 calcChain 该整个删掉。

    Excel 自己就是这么处理的:没有公式时不写这个零件。留一个空的
    <calcChain/> 反而可能被判为损坏。

    这是「零件一个不少」这条不变量唯一有据可查的例外,所以顺带断言
    [Content_Types].xml 和 workbook 关联里的登记也一起清掉了 —— 否则
    就成了指向不存在零件的悬空关联,换了个地方犯同样的错。
    """
    out = tmp_path / "out.xlsx"
    write_cells(FIXTURE, out, {ref: "覆盖" for ref in _formula_cells(FIXTURE)})

    parts = set(list_parts(out))
    assert "xl/calcChain.xml" not in parts, "calcChain 已空,应该整个删掉"

    with zipfile.ZipFile(out) as z:
        assert "/xl/calcChain.xml" not in z.read("[Content_Types].xml").decode()
        assert "calcChain.xml" not in z.read("xl/_rels/workbook.xml.rels").decode()

    # 其余零件一个都不能少
    expected = set(list_parts(FIXTURE)) - {"xl/calcChain.xml"}
    assert expected <= parts, f"误删了 {sorted(expected - parts)}"


def test_no_formula_touched_keeps_calc_chain_intact(tmp_path):
    """没碰公式格的话,calcChain 应当原封不动。"""
    out = tmp_path / "out.xlsx"
    write_cells(FIXTURE, out, {"A1": "只改普通格"})

    assert _calc_chain_refs(out) == _calc_chain_refs(FIXTURE)
    with zipfile.ZipFile(FIXTURE) as a, zipfile.ZipFile(out) as b:
        assert a.read("xl/calcChain.xml") == b.read("xl/calcChain.xml"), "应当字节级不变"
