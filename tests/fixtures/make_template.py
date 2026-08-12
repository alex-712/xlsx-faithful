"""生成用于保真度测试的样本模板。

为什么要专门造一个:保真度测试必须用**内容丰富**的文件。
拿一个只有几个单元格的空模板去测,里面本来就没什么可丢的,
测出来一切正常,然后你就上当了。

这个脚本造出来的文件包含 openpyxl 会丢的东西(图形对象及其关联关系),
以及一批它能保住的(数据验证、条件格式、页眉页脚、冻结窗格),
所以既能证明问题,也能说明"不是全都丢,只是丢你不会去检查的那部分"。

用法:
    pip install xlsxwriter
    python tests/fixtures/make_template.py
"""

from pathlib import Path

import xlsxwriter

OUT = Path(__file__).parent / "real_template.xlsx"


def build(path: Path) -> None:
    wb = xlsxwriter.Workbook(str(path))
    ws = wb.add_worksheet("会議記録")

    title = wb.add_format({
        "bold": True, "align": "center", "valign": "vcenter",
        "border": 1, "bg_color": "#DCE6F1", "font_size": 14,
    })
    cell = wb.add_format({"border": 1})

    # 合并单元格 + 边框网格
    ws.merge_range("A1:F1", "会議記録テンプレート", title)
    for r in range(2, 12):
        for c in range(6):
            ws.write_blank(r, c, None, cell)

    # 数据验证(下拉)
    ws.data_validation("B3", {
        "validate": "list",
        "source": ["有償", "無償", "保留"],
    })

    # 条件格式
    ws.conditional_format("C3:C10", {
        "type": "cell", "criteria": ">", "value": 1000,
        "format": wb.add_format({"bg_color": "#FFC7CE"}),
    })

    # 图形对象 —— openpyxl 会把这个丢掉。
    # 刻意放在表格正下方而不是右侧:截图对比时必须落在可见区域内,
    # 否则"丢了什么"根本看不出来。
    ws.insert_textbox("A14", "注意事項:本テンプレートは社外秘", {
        "width": 320, "height": 70,
        "fill": {"color": "#FFFF99"},
        "border": {"color": "#BF8F00", "width": 1.5},
        "font": {"size": 12, "bold": True},
        "align": {"vertical": "middle", "horizontal": "center"},
    })

    # 公式 —— 为的是让包里长出 xl/calcChain.xml。
    # 覆盖一个有公式的单元格时,公式会被清掉,但 calcChain 里那条记录
    # 如果不跟着删,就成了指向「已经没有公式的格子」的悬空引用,
    # Excel 打开时可能提示「发现不可读取的内容」并自动修复。
    ws.write_formula("C11", "=SUM(C3:C10)")
    ws.write_formula("D11", "=C11*2")

    # 自动换行 + 显式行高 —— 排版相关功能要有东西可测。
    wrapped = wb.add_format({"border": 1, "text_wrap": True, "valign": "top"})
    ws.write("A12", "折り返して表示する長めの注記テキスト", wrapped)
    ws.set_row(11, 40)

    # 打印设置
    ws.set_landscape()
    ws.set_paper(9)  # A4
    ws.print_area("A1:F12")
    ws.repeat_rows(0)
    ws.set_header("&L社内資料&R&P/&N")
    ws.set_footer("&C機密")
    ws.fit_to_pages(1, 0)

    # 视图
    ws.freeze_panes(2, 0)
    ws.autofilter("A2:F2")

    wb.close()
    _inject_calc_chain(path)


def _inject_calc_chain(path: Path) -> None:
    """补一个 xl/calcChain.xml。

    xlsxwriter 不产这个零件(它是可选的,Excel 打开时会自己重建),
    但真实业务模板里几乎都有 —— 只要用 Excel 存过带公式的文件就会带上它。
    保真库必须处理它:覆盖公式格时若不同步删掉对应记录,就留下指向
    「已经没有公式的格子」的悬空引用。没有这个零件就测不出那个 bug。
    """
    import re
    import shutil
    import tempfile
    import zipfile

    calc_chain = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<c r="C11" i="1" l="1"/><c r="D11" i="1"/></calcChain>'
    )
    ct_override = (
        '<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>'
    )
    rel = (
        '<Relationship Id="rIdCalcChain" Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>'
    )

    with zipfile.ZipFile(path) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]

    tmp = Path(tempfile.mkdtemp()) / "with_calcchain.xlsx"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in items:
            if info.filename == "[Content_Types].xml":
                data = data.replace(b"</Types>", ct_override.encode() + b"</Types>")
            elif info.filename == "xl/_rels/workbook.xml.rels":
                data = re.sub(rb"</Relationships>", rel.encode() + b"</Relationships>", data)
            zout.writestr(info, data)
        zout.writestr("xl/calcChain.xml", calc_chain)

    shutil.move(str(tmp), str(path))


if __name__ == "__main__":
    build(OUT)
    import zipfile

    parts = sorted(zipfile.ZipFile(OUT).namelist())
    print(f"✅ {OUT}")
    print(f"   {len(parts)} 个零件 / {OUT.stat().st_size} bytes")
    for p in parts:
        print("   ", p)
