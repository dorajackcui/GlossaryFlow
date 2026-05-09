# GlossaryFlow

一个用于游戏/本地化术语表处理的小助手：从需求 Excel 拆出主干/后缀术语，用过往 TB 预填，检查未翻译和重复冲突，回填需求 Excel，再把本批次完成的术语写回过往 TB。

日常使用只需要运行这个入口：

```bash
python term_assistant.py
```

它会显示：

```text
1. 开始新批次：拆分需求 Excel
2. 用过往 TB 预填本批次术语表
3. 我已经翻译完了：检查问题
4. 检查通过了：回填需求 Excel
5. 将本批次已完成术语写回过往 TB
q. 退出
```

跟着提示输入文件路径即可。

## 你平时怎么看文件

小助手会把每次任务的文件都放在**需求 Excel 所在文件夹**下面的 `outputs/` 里。

如果你的需求文件是：

```text
/项目A/需求.xlsx
```

输出会在：

```text
/项目A/outputs/20260508_213000_需求/
  01_本批次术语表.xlsx
  01_本批次术语表_汇总.xlsx
  02_本批次术语表_预填.xlsx
  04_需求_已回填.xlsx
```

你主要看这几个文件：

- `01_本批次术语表_汇总.xlsx`：把主干术语表和后缀术语表放在一个 sheet 里的版本，方便筛选和编辑。
- `02_本批次术语表_预填.xlsx`：需要你翻译的文件。
- `04_需求_已回填.xlsx`：最终交付文件。
- 过往 TB 会在第五步被原地更新，不会另存新文件。

## 最常见操作

### 第一步：拆分需求 Excel

运行：

```bash
python term_assistant.py
```

选择：

```text
1
```

然后输入：

```text
需求 Excel 路径
```

拆分时会自动识别需求 Excel 里的 `target`：

- 如果某一行 `source` 和 `target` 都有内容，会拆成已经配对的主干/后缀术语。
- 如果某一行只有 `source`，也会生成术语，但 `target` 留空，后面再翻译。
- 如果同一个术语重复出现，会优先保留第一个非空 `target`。

小助手会生成：

```text
01_本批次术语表.xlsx
01_本批次术语表_汇总.xlsx
```

两个文件内容等价：`01_本批次术语表.xlsx` 分成 `主干术语表`、`后缀术语表` 两个 sheet；`01_本批次术语表_汇总.xlsx` 只有一个 `汇总术语表` sheet，用 `term_type` 区分主干和后缀。

### 第二步：用过往 TB 预填

再次运行：

```bash
python term_assistant.py
```

选择：

```text
2
```

然后输入：

```text
本批次术语表路径或汇总术语表路径
过往 TB 路径
```

第一行可以填第一步生成的 `01_本批次术语表.xlsx`，也可以填 `01_本批次术语表_汇总.xlsx`。

小助手会生成：

```text
02_本批次术语表_预填.xlsx
```

接下来打开 `02_本批次术语表_预填.xlsx`。

看 `status` 列：

- `matched`：过往 TB 已经有翻译，不用管。
- `needs_translation`：需要你补 `target`。

补完后建议另存为：

```text
02_本批次术语表_已翻译.xlsx
```

### 第三步：翻译完后检查

再次运行：

```bash
python term_assistant.py
```

选择：

```text
3
```

然后输入：

```text
已翻译术语表路径
过往 TB 路径
```

`已翻译术语表路径` 可以是两个 sheet 的术语表，也可以是只有 `汇总术语表` 的汇总文件。

小助手会直接更新你输入的已翻译术语表：

- 每行末尾会多出 `issue_type` 和 `issue_detail` 两列。
- 如果这两列之前已经存在，会先清掉旧内容，再写入本次检查结果。
- 如果这两列为空：这一行没有检查出问题。
- 如果这两列有内容：直接修正同一行的术语，再重新选择 `3` 检查。
- 如果 `target` 为空，会写入 `missing_target`，需要先补翻译。
- `source_conflict` 的 `issue_detail` 会说明是本批次内重复，还是与过往术语表冲突。

检查规则：

- 一个 `source` 只能对应一个 `target`。
- 一个 `target` 不能被多个 `source` 共用。
- 检查范围是：本批次术语表 + 过往 TB。
- 空 `target` 会作为 `missing_target` 写到对应行，不参与冲突检查。

### 第四步：检查通过后回填

再次运行：

```bash
python term_assistant.py
```

选择：

```text
4
```

然后输入：

```text
需求 Excel 路径
已翻译术语表路径
```

`已翻译术语表路径` 可以继续使用两个 sheet 的术语表，也可以使用汇总文件。

小助手会生成：

```text
04_需求_已回填.xlsx
```

这个就是最终交付文件。交付确认后，可以继续第五步，把本批次已完成术语写回过往 TB。

### 第五步：写回过往 TB

再次运行：

```bash
python term_assistant.py
```

选择：

```text
5
```

然后输入：

```text
已翻译术语表路径
过往 TB 路径
```

`已翻译术语表路径` 支持两种格式：`主干术语表` + `后缀术语表` 分开的版本，或只有 `汇总术语表` 的合并版本。

小助手会直接原地更新你输入的过往 TB：

- `target` 为空的术语不会写回。
- 过往 TB 已有同一 `source` 且 `target` 相同：跳过。
- 过往 TB 已有同一 `source` 但 `target` 为空：补上。
- 过往 TB 没有该 `source`：追加到对应术语表。
- 过往 TB 已有同一 `source` 但 `target` 不同：整次拒绝写入，需要先回到第三步检查并处理冲突。

## 文件格式要求

### 需求 Excel

至少有 `source` 列，`target` 列可以有也可以没有：

| source | target |
| --- | --- |
| 命运刻面·笼中梦 | Coupes du destin : Rêve en cage |
| 夜息垂芒 | |

如果有 `target`，第一步会把它一起拆进术语表；如果没有或为空，第一步会生成待翻译术语。回填时如果需求 Excel 没有 `target` 列，脚本会自动新增。

### 过往 TB

可以使用两个 sheet 的格式：

- `主干术语表`
- `后缀术语表`

每个 sheet 至少包含：

| source | target |
| --- | --- |
| 命运刻面 | Coupes du destin |
| ·笼中梦 |  : Rêve en cage |

也可以使用单 sheet 汇总格式，sheet 名为 `汇总术语表`：

| term_type | source | target |
| --- | --- | --- |
| 主干术语表 | 命运刻面 | Coupes du destin |
| 后缀术语表 | ·笼中梦 |  : Rêve en cage |

## 高级用法：直接命令行

如果你想跳过小助手，也可以直接跑：

```bash
python term_workflow.py split 需求.xlsx --output 本批次术语表.xlsx --summary-output 本批次术语表_汇总.xlsx
python term_workflow.py prefill 本批次术语表.xlsx 过往TB.xlsx --output 本批次术语表_预填.xlsx
python term_workflow.py check 本批次术语表_已翻译.xlsx 过往TB.xlsx
python term_workflow.py fill 需求.xlsx 本批次术语表_已翻译.xlsx --output 需求_已回填.xlsx
python term_workflow.py update-tb 本批次术语表_已翻译.xlsx 过往TB.xlsx
```

`prefill`、`check`、`fill`、`update-tb` 的术语表参数都可以传两个 sheet 格式，也可以传 `汇总术语表` 格式。

## 安装依赖

如果当前环境没有 `openpyxl`：

```bash
pip install -r requirements.txt
```

## 运行测试

```bash
python -m unittest discover -s tests
```
