# xmap — Generic Excel schema-map tool

一个不绑定任何具体任务的通用工具，把「理解表格结构」和「读取表格数据」彻底分离：
**模型只写“读表规格”(spec) 并看校验摘要；引擎用 Excel COM 把整张表读出来。**
表再大、结构再千变万化，模型每张表只付出“写几百字节 spec + 看几行摘要”的成本。

## 为什么这样设计（省 token 的核心）

| 做法 | 效果 |
|---|---|
| 只探测结构（表名/表头/几行样例），不 dump 全表 | 模型永远看不到整张表 |
| 模型输出的是“列映射规格”，不是数据 | 模型工作量 = 几百字节 JSON，与表大小无关 |
| 引擎逐格读全表（Excel COM API） | 读取发生在模型上下文之外，零 token |
| verify 只产出一份小摘要 | 模型靠摘要判断对错，不回读明细 |
| 表头指纹复用 spec | 同族表再次出现时模型零参与 |

## 命令

```
# 看有哪些表
xmap.ps1 sheets <file.xlsx>

# 看某张表的结构（表头 + 前几行样例）
xmap.ps1 headers <file.xlsx> -Sheet <sheet名> [-Rows 5]

# 表头指纹（用于 spec 复用：相同指纹 → 复用同一份 spec）
xmap.ps1 hash <file.xlsx> -Sheet <sheet名>

# 把一份 spec 按指纹存入 specs/ 目录（下次同指纹表自动复用）
xmap.ps1 save-spec <file.xlsx> -Sheet <sheet名> -Spec <spec.json>

# 提取整张表 → 标准化记录(JSONL)。-Spec 可省略：缺省按指纹从 specs/ 自动匹配
xmap.ps1 extract <file.xlsx> -Sheet <sheet名> [-Spec <spec.json>] -Out <records.jsonl>

# 校验：记录数 + 各数值列求和 + 可选与表内“合计”行对比
xmap.ps1 verify <file.xlsx> -Sheet <sheet名> -Spec <spec.json> -Records <records.jsonl> [-Report report.txt]

# 跨表聚合：多份 JSONL 按字段分组求和
xmap.ps1 aggregate -Records a.jsonl,b.jsonl -GroupBy name,category -Sum qty,amount -Out agg.jsonl
```

## Spec 格式

```json
{
  "headerRows": 2,            // 表头占几行（列名查找范围）
  "dataStart": 3,             // 数据从第几行开始
  "cols": {                   // 语义字段 → 定位方式（表头文字 / 列号 / 列字母）
    "name": "商品名称",
    "category": "类别",
    "qty": "数量",
    "amount": "金额"
  },
  "skipIfNameEmpty": true,    // 跳过名称为空的行（小计/合并行）
  "numeric": { "qty": 1, "amount": 1 },       // 校验时求和的字段
  "checkTotal": { "col": "金额", "nameCol": "商品名称", "label": "合计" }  // 可选：与表内合计行对比
}
```

**字段定位规则**（`cols` 里的值）：
- 是数字 → 当作列号
- 是 1~3 位字母 → 当作列字母（如 `A`、`F`）
- 其他 → 当作表头文字，在 `headerRows` 行内精确匹配

## 复用工作流（spec 库 + 指纹）

1. 第一张表：`hash` 拿指纹 → 写 `spec` → `extract` + `verify` 跑通
2. `save-spec` 按指纹存进 `specs/<指纹>.json`
3. 之后同指纹表直接 `extract`（自动匹配 spec），**无需模型重新理解**

## 输出

`extract` → JSONL（每行一条记录，字段名 = spec 的 `cols` 键，数值用原始 `Value2`）。
`verify` → 记录数 + 各数值列求和 + 可选 checkTotal。
`aggregate` → 按 `-GroupBy` 字段分组的 JSONL，对 `-Sum` 字段求和。

## 依赖

- Windows + 安装了 Microsoft Excel（使用 Excel COM）
- PowerShell（`pwsh` 或 Windows PowerShell）
- `.ps1` 需为 UTF-8 with BOM（否则中文表头会乱码）

## 注意

- 合并单元格：表头跨行列时，用 `headers` 先看清，再在 spec 里指定正确列。
- 同一列可能有多个表头名，spec 里写实际出现的那个。
- `verify` 的 `checkTotal` 只有在该表存在可识别的“合计”行时才有效；没有就省略。
- 包内 `specs/` 仅含一份通用示例 `example_spec.json`；真实 spec 请按你的表用 `save-spec` 沉淀。
