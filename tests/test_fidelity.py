"""保真度对比测试:xlsx-faithful vs openpyxl。

⚠️ 这是这个库存在的意义所在,必须用**真实业务文件**跑。
自造的简化模板里本来就没什么可丢的,会骗过你。

把带表单控件 / 图形 / 打印设置的真实模板(脱敏后)放到
tests/fixtures/real_template.xlsx,这些测试会自动启用。
"""

import shutil
import zipfile
from pathlib import Path

import pytest

from xlsx_faithful import list_parts, write_cells

FIXTURE = Path(__file__).parent / "fixtures" / "real_template.xlsx"

pytestmark = pytest.mark.skipif(
    not FIXTURE.is_file(),
    reason="缺少 tests/fixtures/real_template.xlsx,详见该目录 README",
)


def test_faithful_keeps_every_part(tmp_path):
    out = tmp_path / "out.xlsx"
    write_cells(FIXTURE, out, {"A1": "保真度测试"})

    before, after = list_parts(FIXTURE), list_parts(out)
    lost = set(before) - set(after)
    assert not lost, f"丢失了 {len(lost)} 个零件: {sorted(lost)}"


def test_faithful_size_not_shrunk(tmp_path):
    """体积骤降是零件丢失的信号。"""
    out = tmp_path / "out.xlsx"
    write_cells(FIXTURE, out, {"A1": "x"})

    src_size, out_size = FIXTURE.stat().st_size, out.stat().st_size
    assert out_size > src_size * 0.9, (
        f"体积从 {src_size} 掉到 {out_size},很可能丢了零件"
    )


@pytest.mark.parametrize("kind", ["ctrlProps", "drawings", "printerSettings"])
def test_critical_parts_survive(tmp_path, kind):
    """表单控件 / 图形 / 打印设置,这三类是 openpyxl 最常丢的。"""
    if not any(kind in n for n in list_parts(FIXTURE)):
        pytest.skip(f"模板里本来就没有 {kind}")

    out = tmp_path / "out.xlsx"
    write_cells(FIXTURE, out, {"A1": "x"})
    assert any(kind in n for n in list_parts(out)), f"{kind} 丢了"


def test_compare_against_openpyxl(tmp_path):
    """把差距量化出来。这个数字就是 README 里那张表的来源。"""
    openpyxl = pytest.importorskip("openpyxl")

    faithful_out = tmp_path / "faithful.xlsx"
    write_cells(FIXTURE, faithful_out, {"A1": "x"})

    openpyxl_out = tmp_path / "openpyxl.xlsx"
    shutil.copy(FIXTURE, openpyxl_out)
    wb = openpyxl.load_workbook(openpyxl_out)
    wb.active["A1"] = "x"
    wb.save(openpyxl_out)

    original = set(list_parts(FIXTURE))
    lost_by_openpyxl = original - set(list_parts(openpyxl_out))
    lost_by_faithful = original - set(list_parts(faithful_out))

    print(
        f"\n零件数    原始 {len(original)}"
        f" | openpyxl {len(original) - len(lost_by_openpyxl)}"
        f" | faithful {len(original) - len(lost_by_faithful)}"
    )
    print(
        f"文件大小  原始 {FIXTURE.stat().st_size}"
        f" | openpyxl {openpyxl_out.stat().st_size}"
        f" | faithful {faithful_out.stat().st_size}"
    )
    if lost_by_openpyxl:
        print(f"openpyxl 丢失: {sorted(lost_by_openpyxl)}")

    assert not lost_by_faithful


def test_written_value_is_readable(tmp_path):
    """保真的同时,值得真写进去了。"""
    openpyxl = pytest.importorskip("openpyxl")
    out = tmp_path / "out.xlsx"
    write_cells(FIXTURE, out, {"A1": "写入校验"})

    wb = openpyxl.load_workbook(out)
    assert wb.active["A1"].value == "写入校验"


def test_output_opens_as_valid_zip(tmp_path):
    out = tmp_path / "out.xlsx"
    write_cells(FIXTURE, out, {"A1": "x"})
    with zipfile.ZipFile(out) as z:
        assert z.testzip() is None, "输出的 zip 结构损坏"
