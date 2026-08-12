"""往 xlsx 模板里写单元格,同时保住模板里的一切。

两层保真:

1. **包层面** —— xlsx 就是个 zip。只重写存单元格数据的那个 XML,
   其余所有内部零件字节级原样复制。

2. **XML 层面** —— 即使在目标工作表里,也只重写 ``<sheetData>`` 这一段。
   根元素的命名空间声明、``<drawing>``、``<pageSetup>``、``<dataValidations>``
   这些统统按原始字节保留。

第 2 层是吃过亏才加的:早期版本用 ElementTree 重新序列化整个工作表,
结果命名空间前缀从 ``r:`` 变成了 ``ns1:``。XML 语义上等价,但很多解析器
硬编码期望 ``r:``,文件直接打不开。

一个宣称保真的库,不能自己重新生成任何不必要的东西。
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from typing import Any, Mapping
from xml.etree import ElementTree as ET

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_REF_RE = re.compile(r"^([A-Za-z]+)(\d+)$")
# 匹配 <sheetData/> 或 <sheetData ...>...</sheetData>
_SHEETDATA_RE = re.compile(
    rb"<sheetData\s*/>|<sheetData(?:\s[^>]*)?>.*?</sheetData>", re.S
)


class XlsxFaithfulError(Exception):
    """本库抛出的所有异常的基类。"""


# ---------------------------------------------------------------- 引用解析


def _split_ref(ref: str) -> tuple[str, int]:
    """'B3' -> ('B', 3)"""
    m = _REF_RE.match(ref.strip())
    if not m:
        raise XlsxFaithfulError(f"非法的单元格引用: {ref!r}(期望形如 'B3')")
    return m.group(1).upper(), int(m.group(2))


def _col_to_num(col: str) -> int:
    """'A' -> 1, 'Z' -> 26, 'AA' -> 27"""
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def _q(tag: str) -> str:
    return f"{{{NS}}}{tag}"


# ---------------------------------------------------------------- 命名空间


def _register_namespaces(xml: bytes) -> None:
    """把原始 XML 里出现的前缀全部注册回去。

    不做这一步,ElementTree 序列化时会自己编 ns0/ns1 这类前缀。
    """
    try:
        for _, (prefix, uri) in ET.iterparse(io.BytesIO(xml), events=["start-ns"]):
            ET.register_namespace(prefix, uri)
    except ET.ParseError:
        pass  # 交给后面的正式解析去报错
    ET.register_namespace("", NS)


# ---------------------------------------------------------------- 行 / 单元格


def _find_row(sheet_data: ET.Element, rownum: int) -> ET.Element | None:
    for row in sheet_data.findall(_q("row")):
        r = row.get("r")
        if r and int(r) == rownum:
            return row
    return None


def _insert_row(sheet_data: ET.Element, rownum: int) -> ET.Element:
    """按行号顺序插入一个新的 <row>。"""
    row = ET.Element(_q("row"), {"r": str(rownum)})
    idx = len(list(sheet_data))
    for i, existing in enumerate(sheet_data):
        r = existing.get("r")
        if r and int(r) > rownum:
            idx = i
            break
    sheet_data.insert(idx, row)
    return row


def _find_cell(row: ET.Element, ref: str) -> ET.Element | None:
    for cell in row.findall(_q("c")):
        if cell.get("r") == ref:
            return cell
    return None


def _insert_cell(row: ET.Element, ref: str) -> ET.Element:
    """按列顺序插入一个新的 <c>。顺序错了 Excel 会报文件损坏。"""
    col, _ = _split_ref(ref)
    target = _col_to_num(col)
    cell = ET.Element(_q("c"), {"r": ref})
    idx = len(list(row))
    for i, existing in enumerate(row):
        r = existing.get("r")
        if not r:
            continue
        c_col, _ = _split_ref(r)
        if _col_to_num(c_col) > target:
            idx = i
            break
    row.insert(idx, cell)
    return cell


def _had_formula(cell: ET.Element) -> bool:
    return cell.find(_q("f")) is not None


def _set_value(cell: ET.Element, value: Any) -> None:
    """写值。字符串走 inline string,绕开 sharedStrings。

    绕开共享池的原因:改共享池要查重、追加、更新计数、修正所有索引,
    任何一步错了都会破坏其他单元格的引用。inline string 完全不碰它。
    """
    # 清掉原有的 <v> / <is> / <f>。覆盖公式单元格时公式必须一起去掉,
    # 否则 Excel 重算时会把我们写的值盖掉。
    for child in list(cell):
        cell.remove(child)

    if value is None:
        cell.attrib.pop("t", None)
        return

    if isinstance(value, bool):
        cell.set("t", "b")
        ET.SubElement(cell, _q("v")).text = "1" if value else "0"
        return

    if isinstance(value, (int, float)):
        cell.attrib.pop("t", None)  # 数字类型不带 t 属性
        ET.SubElement(cell, _q("v")).text = (
            repr(value) if isinstance(value, float) else str(value)
        )
        return

    text = str(value)
    cell.set("t", "inlineStr")
    t_el = ET.SubElement(ET.SubElement(cell, _q("is")), _q("t"))
    t_el.text = text
    if text != text.strip():
        t_el.set(XML_SPACE, "preserve")


# ---------------------------------------------------------------- 片段改写


def _apply_cells(sheet_xml: bytes, cells: Mapping[str, Any]) -> tuple[bytes, set[str]]:
    """只重写 <sheetData> 这一段,工作表其余字节原样保留。

    返回改写后的 XML,以及「原本有公式、被这次写值清掉了」的单元格引用集合 ——
    调用方要拿它去同步 calcChain,否则留下悬空引用。
    """
    cleared_formulas: set[str] = set()

    m = _SHEETDATA_RE.search(sheet_xml)
    if not m:
        raise XlsxFaithfulError("工作表 XML 里找不到 <sheetData>")

    _register_namespaces(sheet_xml)

    fragment = m.group(0)
    if fragment.endswith(b"/>"):  # 空表:<sheetData/>
        fragment = b"<sheetData></sheetData>"

    # 包一层带默认命名空间的根,片段才能被正确解析
    wrapper = b'<wrap xmlns="' + NS.encode() + b'">' + fragment + b"</wrap>"
    sheet_data = ET.fromstring(wrapper).find(_q("sheetData"))
    if sheet_data is None:
        raise XlsxFaithfulError("解析 <sheetData> 失败")

    for ref, value in cells.items():
        _, rownum = _split_ref(ref)
        ref = ref.strip().upper()

        # 不能用 `or`:ElementTree 里没有子元素的 Element 是 falsy,
        # 一个空的 <row> 会被判成假,导致重复插入同号行。
        row = _find_row(sheet_data, rownum)
        if row is None:
            row = _insert_row(sheet_data, rownum)

        cell = _find_cell(row, ref)
        if cell is None:
            cell = _insert_cell(row, ref)

        if _had_formula(cell):
            cleared_formulas.add(ref)
        _set_value(cell, value)

    new_fragment = ET.tostring(sheet_data, encoding="utf-8")
    # ET 会在片段根上补一个 xmlns 声明,原文里没有,去掉
    new_fragment = new_fragment.replace(b' xmlns="' + NS.encode() + b'"', b"", 1)

    return sheet_xml[: m.start()] + new_fragment + sheet_xml[m.end() :], cleared_formulas


# ---------------------------------------------------------------- calcChain

_CALC_CHAIN = "xl/calcChain.xml"
_CT = "[Content_Types].xml"
_WB_RELS = "xl/_rels/workbook.xml.rels"


def _prune_calc_chain(chain_xml: bytes, cleared: set[str]) -> bytes | None:
    """从 calcChain 里删掉指定单元格的记录。全删空了返回 None。

    calcChain 记的是「哪些格子有公式、按什么顺序算」。写值会把 <f> 清掉,
    这里的记录不跟着删就成了指向「已经没有公式的格子」的悬空引用,
    Excel 打开时可能提示「发现不可读取的内容」。

    用正则而不是解析重写:和整体思路一致 —— 能不重新生成就不重新生成。
    """
    remaining = chain_xml
    for ref in cleared:
        remaining = re.sub(
            rb'<c r="' + re.escape(ref.encode()) + rb'"[^>]*/>', b"", remaining
        )
    if not re.search(rb"<c\s", remaining):
        return None
    return remaining


def _drop_part_registrations(data: bytes, filename: str) -> bytes:
    """把 [Content_Types].xml / workbook.xml.rels 里对某个零件的登记去掉。

    删了零件却留着登记,就是把悬空引用换个地方犯一遍。
    """
    if filename == _CT:
        return re.sub(rb'<Override PartName="/xl/calcChain\.xml"[^>]*/>', b"", data)
    if filename == _WB_RELS:
        return re.sub(rb'<Relationship[^>]*Target="calcChain\.xml"[^>]*/>', b"", data)
    return data


# ---------------------------------------------------------------- 公开 API


def list_parts(path: str | Path) -> list[str]:
    """列出 xlsx 内部的所有零件。用来验证保真度。"""
    with zipfile.ZipFile(path) as z:
        return sorted(z.namelist())


def write_cells(
    src: str | Path,
    dst: str | Path,
    cells: Mapping[str, Any],
    sheet: str = "xl/worksheets/sheet1.xml",
) -> None:
    """把 cells 写进 src 模板,输出到 dst,其余内容原样保留。

    参数:
        src:   模板文件路径
        dst:   输出文件路径
        cells: {"B3": "值", "C7": 1200} 形式的映射。
               str 走 inline string,int/float 作数字,bool 作布尔,None 清空。
        sheet: 目标工作表在包内的路径,默认第一张表。

    示例:
        >>> write_cells("template.xlsx", "out.xlsx", {"B3": "会議記録", "D5": 1200})
    """
    src, dst = Path(src), Path(dst)
    if not src.is_file():
        raise XlsxFaithfulError(f"模板文件不存在: {src}")

    with zipfile.ZipFile(src) as zin:
        if sheet not in zin.namelist():
            raise XlsxFaithfulError(
                f"包内找不到 {sheet}。可用的工作表: "
                f"{[n for n in zin.namelist() if n.startswith('xl/worksheets/')]}"
            )

        new_sheet, cleared_formulas = _apply_cells(zin.read(sheet), cells)

        # 覆盖了公式格就得同步 calcChain,否则留下悬空引用。
        # drop_calc_chain 为真时,这个零件连同它在 Content_Types / workbook
        # 关联里的登记一起去掉 —— 这是「零件一个不少」唯一的例外,
        # 因为 Excel 自己在没有公式时也不写这个零件,留个空的反而可能被判损坏。
        new_calc_chain: bytes | None = None
        drop_calc_chain = False
        if cleared_formulas and _CALC_CHAIN in zin.namelist():
            new_calc_chain = _prune_calc_chain(zin.read(_CALC_CHAIN), cleared_formulas)
            drop_calc_chain = new_calc_chain is None

        dst.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if drop_calc_chain and item.filename == _CALC_CHAIN:
                    continue

                # 传 ZipInfo 而不是文件名,保留原始时间戳和压缩方式
                if item.filename == sheet:
                    data = new_sheet
                elif item.filename == _CALC_CHAIN and new_calc_chain is not None:
                    data = new_calc_chain
                else:
                    data = zin.read(item.filename)
                    if drop_calc_chain:
                        data = _drop_part_registrations(data, item.filename)
                zout.writestr(item, data)
