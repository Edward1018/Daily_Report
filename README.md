# 教育出版学科销量日报工具（单文件版）

一个**单文件、自包含**的 Python 工具，读取 Quick BI 学科销量明细表（Excel），按「大盘 / 学科 / 体系 / 年级 / 版本 / 区域 / 机构」多维度清洗、汇总、对标目标，并生成「结构化数据 / 纯文本 / HTML / 智能体提示词」四种输出。

> 本仓库只放 **一个文件** `daily_report.py`（配置/映射/逻辑全部内嵌），不区分 codex 版、workbuddy 版两个文件夹——同一份代码，**两个大模型工具共用**，仅在「下载数据」这一步用到各自不同的浏览器插件（见下文）。

---

## 一、你要上传哪个文件夹？

**只上传 `Daily_Report/` 这一个文件夹即可**，里面包含：

| 文件 | 作用 |
|---|---|
| `daily_report.py` | 唯一核心文件，单文件自包含，无任何子包依赖 |
| `README.md` | 本说明（含双平台对话调用方式） |
| `outputs/`（可选，可不上传） | 本地试运行产物，建议加入 `.gitignore` |

不需要 `edupub-sales-daily-toolkit` 那套多文件 package，已被本单文件版取代。

---

## 二、环境依赖（与平台无关，二者相同）

工具本身是纯 Python，**除了 Python 标准库，只需两个第三方库**：

```bash
pip install pandas openpyxl
```

> 已验证在 pandas 3.0.x 下可运行。如需图表导出 HTML，本工具内置渲染，不另装包。

---

## 三、两种大模型工具如何「对话调用」本文件

本工具是单文件 Python 脚本。**Codex 和 WorkBuddy 都是能执行 shell 命令的 LLM 智能体**，因此「调用」的本质对两者完全一致：

1. **取脚本**：从 GitHub 拉取单文件（一行 `curl` 即可，无需 clone 整个仓库、无需 `pip install` 本工具）；
2. **取数据**：从 Quick BI 导出明细表 Excel —— **这一步两者用的浏览器插件不同**（见下）；
3. **跑脚本**：`python daily_report.py --source 数据.xlsx --current 26暑`。

### ⚠️ 唯一的平台差异：下载数据用的浏览器插件不同

| 环节 | Codex | WorkBuddy |
|---|---|---|
| **下载 Quick BI 数据** | 用 **OpenAI Codex 应用内浏览器插件**（`agent.browsers.*`，自动登录并导出 Excel） | 用 **Playwright MCP / 内置浏览器技能**（按 `--download workbuddy` 生成的操作清单自动登录导出） |
| 运行脚本 | 直接 shell 执行 | 直接 shell 执行（Craft 模式） |
| 输出落盘/推送 | 同上 | 可直接接腾讯文档 / 飞书连接器 |

> 一句话：**脚本和参数完全一样，只有「怎么把 Quick BI 的 Excel 弄到本地」这一步，Codex 用它的浏览器、WorkBuddy 用 Playwright**。

---

### 方式 A：在 WorkBuddy 里对话调用

把下面这句话发给 WorkBuddy（已连接 Quick BI 浏览器能力时）：

```
请帮我从 GitHub 拉取 Edward1018/Daily_Report 的 daily_report.py，
用 Playwright（workbuddy 下载方式）从 Quick BI 导出最新学科销量明细表，
然后运行脚本生成本期（26暑）日报，输出 HTML 和文本。
```

WorkBuddy 会：先用 `--download workbuddy` 生成 Quick BI 操作清单并自动导出 Excel → 再 `curl` 拉取脚本 → 执行 `python daily_report.py --source <导出文件> --current 26暑 --render html,text`。

### 方式 B：在 Codex 里对话调用

把下面这句话发给 Codex：

```
请从 GitHub 拉取 Edward1018/Daily_Report 的 daily_report.py，
用 codex 内置浏览器插件从 Quick BI 导出最新学科销量明细表，
然后运行脚本生成本期（26暑）日报，输出 HTML 和文本。
```

Codex 会：用其应用内浏览器（`agent.browsers`）登录 Quick BI 导出 Excel → `curl` 拉取脚本 → 执行 `python daily_report.py --source <导出文件> --current 26暑 --render html,text`。

### 极简「一行取脚本」命令（两个平台通用）

```bash
curl -L -o daily_report.py https://raw.githubusercontent.com/Edward1018/Daily_Report/main/daily_report.py
```

---

## 四、命令行参数

```text
--source, -s         Quick BI 明细表 Excel 路径（本地文件模式）
--download {codex,workbuddy}   仅生成对应平台的浏览器下载操作清单，不分析数据
--report-url          Quick BI 报表页地址（配合 --download）
--save-dir            下载 Excel 保存目录（配合 --download）
--current, -c         本期学期（默认 26暑）
--prior, -p           上期学期（默认 26春，用于环比）
--yoy YOY             同比学期（默认 25暑）
--focus, -f           聚焦分组，如 小学数学 / 语文 / 英语
--output-dir, -o      输出目录（默认 ./outputs）
--render, -r          text | html | agent_prompt | json（可逗号分隔，如 html,text）
```

### 示例

```bash
# 1) 用本地已有的 Quick BI 数据，生成本期 26暑 的 HTML + 文本日报
python daily_report.py -s 平台数据明细.xlsx -c 26暑 -r html,text

# 2) 只看「语文」分组
python daily_report.py -s 平台数据明细.xlsx -c 26暑 -f 语文 -r text

# 3) 只生成 WorkBuddy 的 Quick BI 下载操作清单（不分析）
python daily_report.py --download workbuddy --report-url "https://<quickbi>/..." --save-dir ./downloads

# 4) 只生成 Codex 的下载操作清单
python daily_report.py --download codex --report-url "https://<quickbi>/..." --save-dir ./downloads
```

---

## 五、目标配置与学科口径（已内嵌，无需改代码）

目标值与学科合并口径写死在 `daily_report.py` 顶部 `TARGETS` / `BIZ_MAP` 中，对照看板目标：

- **大盘** = 合计目标；
- **小学数学** = 小数同步 + 小数培优（二者之和对照小学数学目标）；
- **语文** = 小学语文 + 初中语文；
- **英语** = 小学英语 + 初中英语；
- 学期对应清楚：**看哪一学期的销量，就对照哪一学期的目标**（25春/25暑/25秋/26春/26暑…）。

> 注：原始 Quick BI 明细表的 `学科` 列只有「小学数学」粒度（用 `体系` 区分思维突破/能力提高等四大体系），「小数同步/小数培优」是**业务合并口径**，原始宽表只给「小学数学」合计，无法在表内拆出同步/培优——本工具按看板口径做合计对标，已在代码注释中说明。

---

## 六、智能体动态生成（不死板）

`--render agent_prompt` 会输出一份「数据实况 + 分析任务」的提示词，交给 LLM（或本工具的文本/HTML 兜底渲染）根据**当期真实数据**撰写「整体情况 → 关键问题 → 行动建议」，而非套固定模板。

---

## 七、常见问题

- **Q：必须 clone 整个仓库吗？** 不用。单文件 `curl` 一行即可取用（见第三节）。
- **Q：没有 Quick BI 浏览器权限怎么办？** 手动从 Quick BI 导出 Excel，用 `-s` 直接喂给脚本即可，下载环节可跳过。
- **Q：pandas 版本？** 已验证 3.0.x；2.x 同样可用。
