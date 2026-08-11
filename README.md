# xlsx-faithful

**openpyxl 会毁掉你的 Excel 模板,这个库不会。**

客户给的模板,填完数据原样还回去。

```python
from xlsx_faithful import write_cells

write_cells("客户模板.xlsx", "交付.xlsx", {"B3": "有償", "C3": 1200})
```

零依赖,只用标准库。Python 3.9+。

---

## 问题:它丢东西,而且不告诉你

仓库自带的样本模板跑一次 `pytest -k compare_against_openpyxl -s`:

```
零件数    原始 12  | openpyxl 9    | xlsx-faithful 12
文件大小  原始 7439 | openpyxl 5713 | xlsx-faithful 7446

openpyxl 丢失:
  xl/drawings/drawing1.xml             图形对象
  xl/worksheets/_rels/sheet1.xml.rels  图形的关联关系
  xl/sharedStrings.xml
```

模板里那个黄色文本框,保存之后彻底不存在了。

**但请注意它没丢什么**:数据验证在、条件格式在、页眉页脚在、冻结窗格在。所以你打开文件一眼看去毫无异常。**它丢的恰好是你不会去检查的那部分。**

更麻烦的是,很多轻量预览器(包括 VSCode 的 Excel 插件)压根不渲染图形对象 —— 你连"它没了"都发现不了。

真实业务模板上更糟。一份建筑业会议记录表(55 列 × 103 行、180 个合并单元格、带表单控件),openpyxl 保存后 **23 个零件只剩 10 个,30,615 字节掉到 16,438**。

自己复现:

```bash
python tests/fixtures/make_template.py     # 生成样本
pytest -k compare_against_openpyxl -s      # 跑对比
```

## 为什么

`openpyxl` 保存 xlsx 的方式,本质是**读进来、按自己的对象模型重新生成一遍**。它的模型覆盖不到的零件,重新生成时就不存在了。

而 xlsx 其实就是个 zip:

```
[Content_Types].xml
xl/workbook.xml
xl/worksheets/sheet1.xml     ← 单元格数据只在这里
xl/styles.xml
xl/drawings/drawing1.xml     ← 图形
xl/ctrlProps/ctrlProp1.xml   ← 表单控件
xl/printerSettings/...       ← 打印设置
...
```

填数据只需要动 `sheet1.xml`。**其余二十几个文件根本不需要碰。**

所以这个库的做法是:只重写那一个 XML,**其余全部字节级原样复制**。

可靠性来自一个很朴素的道理:

> 它根本不理解那些控件和图形,所以它不可能弄丢。

## 安装

```bash
pip install xlsx-faithful
```

## 用法

### 基本

```python
from xlsx_faithful import write_cells

write_cells(
    "template.xlsx",
    "out.xlsx",
    {
        "B3": "会議記録",     # 字符串
        "D5": 1200,          # 整数
        "E5": 3.14,          # 浮点
        "F5": True,          # 布尔
        "G5": None,          # 清空
    },
)
```

### 指定工作表

```python
write_cells("t.xlsx", "out.xlsx", {"A1": "x"}, sheet="xl/worksheets/sheet2.xml")
```

### 验证保真度

```python
from xlsx_faithful import list_parts

before = list_parts("template.xlsx")
after  = list_parts("out.xlsx")
assert before == after          # 零件一个不少
print(set(before) - set(after)) # 空集
```

**建议把这个断言加进你自己的测试。** 保真度是这个库唯一的卖点,值得每次验证。

## 一个实现细节:绕开 sharedStrings

xlsx 里的字符串默认存在共享池 `sharedStrings.xml`,单元格通过索引引用:

```xml
<c r="B3" t="s"><v>42</v></c>          <!-- 引用共享池第 42 项 -->
```

按这条路走,每写一个字符串都要查重、追加、更新计数、修正索引,错一步就会破坏其他单元格的引用。

本库改用 xlsx 规范允许的 inline string,值直接写在单元格里:

```xml
<c r="B3" t="inlineStr"><is><t>值</t></is></c>
```

**完全不碰共享池。** 代价是文件略大,对业务数据量来说可以忽略。

## 限制(诚实告知)

| 限制 | 说明 |
|---|---|
| 只重写目标工作表的 XML | 该表的 XML 会被重新序列化。**其他所有零件不受影响**,这是核心保证 |
| 不更新 `<dimension>` | 写入超出原范围的单元格时,dimension 不会扩展。Excel 通常能容忍 |
| 不重算公式 | 覆盖公式单元格时会删掉 `<f>`,只留你写的值 |
| 不处理共享公式 | 目标单元格若是共享公式的一部分,建议先在模板里拆开 |
| 不新建工作表 | 只往已有的表里写 |

**如果你需要的是创建新文件、做数据分析、画图表,请用 openpyxl。** 那些场景它做得很好。

这个库只解决一件事:**客户给的模板必须原样返回**。

## 谁需要它

典型场景是**格式不能动的交付**:

| 场景 | |
|---|---|
| 政务 / 金融 / 建筑 / 制造 | 必须用甲方给的表格,改一点就被打回 |
| 日本企业 | 表格文化重,模板要求严 |
| 企业内部系统 | 导出到财务、法务指定的固定格式 |
| **AI / Agent 应用** | 见下一节 |

### ❌ 不适用,请用 openpyxl

创建新文件 · 读数据做分析 · 生成图表 · 需要重算公式。

**这些场景 openpyxl 做得很好,这个库一件都不做。** 它只解决"模板必须原样返回"这一件事。

## Agent 场景

agent 从非结构化内容里提取字段、再填进固定模板,是当下很高频的一类需求。这个场景下模板丢失更致命 —— **通常是批量跑的,丢一次就是丢一批**,而且没人会逐个打开检查。

推荐的分工:

```
agent  →  只输出 {单元格: 值} 的映射
代码   →  负责把映射写进文件
```

```python
cells = agent_extract(会议记录)      # agent 只做判断
write_cells(模板, 输出, cells)        # 代码负责落盘
```

这样结果不对时,你能一眼分清是**填错了**(模型判断问题)还是**写坏了**(文件操作问题)。这两类问题排查方向完全不同,混在一起会非常难查。

**别让 agent 直接生成二进制格式。** 把不确定的部分和确定的部分分开,是和 LLM 打交道的通用做法。

## 开发

```bash
git clone https://github.com/alex-712/xlsx-faithful
cd xlsx-faithful
pip install -e ".[dev]"
pytest
```

### ⚠️ 关于测试样本

保真度测试的样本**必须内容丰富**,否则测了等于没测。

我第一次验证的时候,用自己造的一个 4KB 空模板,几个单元格加一点格式。结果一切正常,于是我以为没问题。**当然正常,那个文件里本来就没什么可丢的。**

`tests/fixtures/make_template.py` 生成的样本刻意包含了图形对象、数据验证、条件格式、打印设置和冻结窗格,所以能真的测出差异。

如果你手上有带表单控件的真实业务模板(脱敏后),放进 `tests/fixtures/real_template.xlsx` 覆盖掉默认样本,能测得更彻底 —— 表单控件是 openpyxl 丢得最狠的一类,但它没法用代码生成。

## License

MIT
