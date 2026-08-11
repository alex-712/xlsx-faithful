"""最小示例:往模板里填数据并验证保真度。"""

from xlsx_faithful import list_parts, write_cells

SRC, DST = "template.xlsx", "out.xlsx"

write_cells(SRC, DST, {
    "B3": "会議記録",
    "D5": 1200,
    "E5": 3.5,
    "F5": True,
    "G5": None,      # 清空
})

before, after = list_parts(SRC), list_parts(DST)
lost = set(before) - set(after)
print(f"零件 {len(before)} → {len(after)}")
print("✅ 一个不少" if not lost else f"❌ 丢了: {sorted(lost)}")
