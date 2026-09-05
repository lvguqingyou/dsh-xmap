# dsh-xmap

DSH（DeepSeek Harness）工具插件：把通用 Excel 读表工具 [xmap](xmap/README.md) 封装为模型可调用的 7 个工具。

**核心思路**：模型永远不 dump 整张表。模型只写一份几百字节的「列映射 spec」并查看小摘要；引擎通过 Excel COM 在模型上下文之外逐格读取整张表。表再大、版式再杂，每张表只付出「写一份 spec + 看几行摘要」的 token 成本；同版式的表用「表头指纹」自动复用 spec，模型零参与。

## 工具一览

| 工具 | 作用 |
|---|---|
| `xmap_sheets` | 列出工作簿的所有工作表（名称/行数/列数） |
| `xmap_headers` | 预览某张表的表头与前几行，用于理解结构 |
| `xmap_hash` | 表头指纹（MD5）——同指纹即同版式，用于 spec 复用 |
| `xmap_extract` | 按 spec 把整张表提取为标准化 JSONL（缺 spec 时按指纹自动匹配） |
| `xmap_verify` | 校验：记录数 + 各数值列求和（可选与表内合计行对比） |
| `xmap_save_spec` | 把验证过的 spec 存入 spec 库（按指纹命名） |
| `xmap_aggregate` | 多份 JSONL 按字段分组求和 |

## 环境要求

- **Windows** + 已安装 **Microsoft Excel**（引擎使用 Excel COM）。
- **DeepSeek Harness**：在当前产品代（DSH `0.1.3-alpha.1` 系列 / cordis `4.0.x`）上开发与验证。API 在旧版/未来版本宿主上可能有差异。
- PowerShell：优先 `pwsh`，没有时自动回退 Windows PowerShell 5.1。
- 工具调用沿用会话沙箱策略；跨工作区写文件时按宿主惯例提供 `sandbox_permissions` + `justification`。

## 安装

从 npm（发布后）或本地目录均可：

```sh
# 方式一：npm 包（发布后可用）
dsh plugin --profile web add dsh-xmap

# 方式二：从源码目录
git clone https://github.com/lvguqingyou/dsh-xmap.git   # HTTPS 不通时可用 SSH：git@github.com:lvguqingyou/dsh-xmap.git
dsh plugin --profile web add <克隆下来的 dsh-xmap 目录>
```

装完**重启 dsh web**（新增 bundle 不热生效），之后工具即出现在模型工具表中。

## 用法示例

```
# 1) 看有哪些表
xmap_sheets file="C:\data\book.xlsx"
# 2) 看某张表结构
xmap_headers file="C:\data\book.xlsx" sheet="Sheet1" rows=5
# 3) 提取 → JSONL
xmap_extract file="C:\data\book.xlsx" sheet="Sheet1" out="C:\data\records.jsonl"
# 4) 校验
xmap_verify file="C:\data\book.xlsx" sheet="Sheet1" spec="C:\spec.json" records="C:\data\records.jsonl"
# 5) 跨表聚合
xmap_aggregate records="C:\a.jsonl,C:\b.jsonl" groupBy="name" sum="amount" out="agg.jsonl"
```

Spec 格式与完整命令说明见 [xmap/README.md](xmap/README.md)（含 `headerRows` / `dataStart` / `cols` / `numeric` / `skipIfNameEmpty` / `checkTotal`）。

## 配置（cordis.yml 的 `config:`）

| 键 | 含义 | 默认 |
|---|---|---|
| `specDir` | spec 库目录；留空用包内自带 `xmap/specs` | 包内 specs |
| `timeoutMs` | 单次 Excel COM 前台超时（毫秒） | 300000 |

## 本地开发

```sh
node scripts/smoke.mjs   # 桩 ctx 冒烟：7 工具注册 / 命令构造 / 结果映射（不碰 Excel、不碰 GUI）
```

从源码目录安装（`link:` 方式）时，Node 按真实路径解析模块，需要把 DSH 模块库的包链接到本目录（Windows junction / 符号链接），缺失会报找不到 `@deepseek-ai/...`：

```
node_modules/@deepseek-ai/<pkg>  ->  %USERPROFILE%\.dsh\profiles\node_modules\@deepseek-ai\<pkg>
```

需要的包：`schemastery`、`dsh-tools`、`dsh-sandbox`（`cordis` 由宿主导入）。改代码后无需重装，但需重启 dsh web 生效。

## 安全提示

插件通过宿主 `ctx.shell` 执行本包内的 `xmap/xmap.ps1`；请从可信来源安装，并在安装前检查来源（第三方插件在你的机器上以你的权限运行）。
