# GlossaryFlow

一个用于游戏/本地化术语表处理的小助手：从需求 Excel 拆出主干/后缀术语，用过往 TB 预填，检查重复冲突，再回填需求 Excel。

日常使用只需要运行这个入口：

```bash
python term_assistant.py
```

它会显示：

```text
1. 开始新批次：拆分需求 Excel + 用过往 TB 预填
2. 我已经翻译完了：检查冲突
3. 检查通过了：回填需求 Excel
4. 用 test.xlsx 跑一遍演示
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
  02_本批次术语表_预填.xlsx
  03_检查报告.xlsx
  04_需求_已回填.xlsx
```

你主要看这几个文件：

- `02_本批次术语表_预填.xlsx`：需要你翻译的文件。
- `03_检查报告.xlsx`：检查是否有重复或冲突。
- `04_需求_已回填.xlsx`：最终交付文件。

## 最常见操作

### 第一次：开始新批次

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
过往 TB 路径
```

小助手会生成：

```text
01_本批次术语表.xlsx
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

### 第二次：翻译完后检查

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
已翻译术语表路径
过往 TB 路径
```

小助手会生成：

```text
03_检查报告.xlsx
```

打开 `03_检查报告.xlsx`：

- 如果只有表头，没有数据行：检查通过。
- 如果有数据行：先修正术语表，再重新选择 `2` 检查。

检查规则：

- 一个 `source` 只能对应一个 `target`。
- 一个 `target` 不能被多个 `source` 共用。
- 检查范围是：本批次术语表 + 过往 TB。
- 空 `target` 不参与冲突检查。

### 第三次：检查通过后回填

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
需求 Excel 路径
已翻译术语表路径
```

小助手会生成：

```text
04_需求_已回填.xlsx
```

这个就是最终文件。

## 文件格式要求

### 需求 Excel

至少有一列：

| source |
| --- |
| 命运刻面·笼中梦 |
| 夜息垂芒 |

可以没有 `target` 列。回填时如果没有，脚本会自动新增。

### 过往 TB

固定包含两个 sheet：

- `主干术语表`
- `后缀术语表`

每个 sheet 至少包含：

| source | target |
| --- | --- |
| 命运刻面 | Coupes du destin |
| ·笼中梦 |  : Rêve en cage |

## 试跑演示

当前目录里有样例文件：

- `test.xlsx`
- `test_glossaries.xlsx`

运行：

```bash
python term_assistant.py
```

选择：

```text
4
```

它会自动跑完整流程，并把结果放进 `outputs/`。

## 高级用法：直接命令行

如果你想跳过小助手，也可以直接跑：

```bash
python term_workflow.py split 需求.xlsx --output 本批次术语表.xlsx
python term_workflow.py prefill 本批次术语表.xlsx 过往TB.xlsx --output 本批次术语表_预填.xlsx
python term_workflow.py check 本批次术语表_已翻译.xlsx 过往TB.xlsx --output 检查报告.xlsx
python term_workflow.py fill 需求.xlsx 本批次术语表_已翻译.xlsx --output 需求_已回填.xlsx
```

## 安装依赖

如果当前环境没有 `openpyxl`：

```bash
pip install -r requirements.txt
```

## 运行测试

```bash
python -m unittest discover -s tests
```
