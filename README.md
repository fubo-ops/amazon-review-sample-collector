# Amazon 评论样本采集器

## 一眼看懂

- 这是什么：按 ASIN 采集亚马逊竞品的评论样本。
- 什么时候用：要做竞品分析、想先拿一批评论原文时。
- 给它什么：一批 ASIN 和登录好的浏览器会话。
- 它给你什么：评论合并表、数据文件和一份汇总说明。

## 项目简介

`amazon-review-sample-collector` 是可独立运行、也可作为 Codex Skill 使用的 Amazon US 竞品评论样本采集器。默认结合 Amazon 多关键词查询、卖家精灵反查词和评论高频词二次扩量，按 ASIN 与评论 ID 去重后输出 Excel、CSV 和 JSON。

仓库只做采集、去重和校验：结果是多关键词召回的评论样本，不代表商品的全部评论，也不能拿来推算全量评论的精确占比；不导出任何浏览器凭据，也不做语义分析。

## 核心功能

- 📦 **批量 ASIN**：一次采集多个竞品 ASIN
- 🔎 **Rufus 查询**：从浏览器评论窗口取样
- 🧩 **混合扩量**：卖家精灵词加评论高频词
- 🛟 **降级继续**：卖家精灵不可用仍扩量
- 💾 **会话与断点**：专用资料目录可续跑
- 📋 **八列输出**：CSV、Excel 与汇总说明
- ✅ **输出校验**：字段、ID 与行数核对

## 技术栈

| 层级 | 技术 | 版本 |
| --- | --- | --- |
| 语言 | Python | 3.12.13 |
| 浏览器自动化 | Playwright（Chromium / 本机 Chrome） | >=1.55,<2 |
| Excel 生成与回读 | openpyxl | >=3.1,<4 |
| 文本解析与输出 | Python 标准库（HTMLParser、asyncio、csv、json） | — |
| 打包 | setuptools + `amazon-review-collector` 命令 | — |
| 测试 | Python `unittest` | — |
| CI | GitHub Actions（Windows / macOS / Ubuntu） | — |

版本取自 `.python-version`、`requirements.txt` 和 `pyproject.toml`。setuptools 只声明 `>=68`，CI 的 Python 矩阵写在工作流里，两者都无可比版本号，记「—」。

## 快速开始

需要 Python 3.10–3.13（本机 `.python-version` 为 3.12.13）；首次采集需要能登录 Amazon 与卖家精灵的浏览器会话。以下命令在仓库根目录执行；Windows PowerShell 激活环境用 `.venv\Scripts\Activate.ps1`。

1. 克隆仓库，建虚拟环境并安装依赖与 Chromium

   ```bash
   git clone https://github.com/fubo-ops/amazon-review-sample-collector.git
   cd amazon-review-sample-collector
   python -m venv .venv
   source .venv/bin/activate
   python -m pip install -r requirements.txt
   python -m playwright install chromium
   ```

2. 跑环境检查，确认 Python 版本与依赖就绪

   ```bash
   python scripts/amazon_reviews.py doctor
   ```

3. 首次登录：打开 Amazon 与卖家精灵标签页，登录完成后回终端按 Enter

   ```bash
   python scripts/amazon_reviews.py init-profile
   ```

4. 批量采集，把占位 ASIN 换成真实的 10 位 ASIN

   ```bash
   python scripts/amazon_reviews.py rufus --asins <ASIN_A> <ASIN_B> --output outputs/run-001
   ```

5. 中断后复用同一输出目录续跑，或切换扩量模式

   ```bash
   python scripts/amazon_reviews.py rufus --asins <ASIN_A> --output outputs/run-001
   python scripts/amazon_reviews.py rufus --expansion second-pass --asins <ASIN_A> --output outputs/run-002
   ```

6. 校验输出目录，只有退出码为 0 才算通过

   ```bash
   python scripts/amazon_reviews.py validate --output outputs/run-001
   ```

`doctor` 只检查 Python 版本和依赖，不替代浏览器启动与登录验证。加 `--headed` 可看到采集过程；只有明确要重新开始时才加 `--restart`。会话默认保存在用户目录下的 `.amazon-review-collector/amazon-profile/`，不要并发启动多个共用同一资料目录的进程。

## 项目结构

只列两层，完整文件清单见文末《详细说明》。

```text
amazon-review-sample-collector/
├── .github/                            # CI 工作流
├── agents/                             # Codex Skill 元数据
├── references/                         # 输出契约与故障处理
├── scripts/                            # 采集脚本
│   ├── __init__.py                     # 包入口
│   └── amazon_reviews.py               # 命令行与采集实现
├── tests/                              # 离线测试
├── .python-version                     # Python 版本固定
├── README.md                           # 项目说明
├── SKILL.md                            # Skill 入口与流程
├── pyproject.toml                      # 项目元数据与命令入口
└── requirements.txt                    # Python 依赖
```

## 架构图

三种扩量模式共用同一条去重与出表链路，断点保存在输出目录里：

```text
┌────────────────────────────────────────────┐
│   ASIN 列表 + 命令行参数                    │
└──────────────────────┬─────────────────────┘
                       ▼
┌────────────────────────────────────────────┐
│   专用浏览器会话（init-profile 登录）       │
└──────────────────────┬─────────────────────┘
                       ▼
┌────────────────────────────────────────────┐
│   Amazon Rufus 基础查询取样                 │
└──────────────────────┬─────────────────────┘
                       ▼
┌────────────────────────────────────────────┐
│   扩量模式：hybrid / second-pass / baseline │
│   hybrid = 卖家精灵反查词 → 评论高频词      │
│   卖家精灵不可用时自动降级继续评论词扩量    │
└──────────────────────┬─────────────────────┘
                       ▼
┌────────────────────────────────────────────┐
│   按 ASIN 与评论 ID 去重                    │
└──────────────────────┬─────────────────────┘
                       ▼
┌────────────────────────────────────────────┐
│   Excel / CSV / JSON / 汇总说明 / 采集日志  │
└──────────────────────┬─────────────────────┘
                       ▼
┌────────────────────────────────────────────┐
│   validate：字段、评论 ID 与行数一致性校验  │
└────────────────────────────────────────────┘

【中断恢复】
采集过程把进度写入输出目录的 checkpoint.json，
复用同一 --output 目录即从断点续跑；全部完成后删除断点。
```

生成文件只保留约定的八列评论字段与汇总信息，不导出浏览器凭据；`outputs/`、登录会话和断点都不提交到 Git。

## 开发计划

### Phase 1: 采集链路

- ✅ 批量 ASIN 与 Rufus 评论取样
- ✅ hybrid / second-pass / baseline 三种扩量模式
- ✅ 卖家精灵不可用时降级继续

### Phase 2: 会话与恢复

- ✅ 专用浏览器资料目录保存登录会话
- ✅ 同一输出目录中断续跑
- ✅ `checkpoint.json` 保留与清理

### Phase 3: 输出与校验

- ✅ 八列评论表（CSV / Excel）
- ✅ 汇总、采集日志与采集说明输出
- ✅ 输出校验（字段、评论 ID 与行数）

### Phase 4: 工程化与验证

- ✅ GitHub Actions（三系统 × Python 3.10/3.13）
- ✅ 离线测试覆盖解析、去重、模式选择、断点与校验
- ⬜ 真实业务环境验收记录
- ⬜ 独立发布路线图

## 文档

- [Skill 入口与执行流程](SKILL.md)
- [输出字段与数据边界](references/output-schema.md)
- [依赖、登录、会话和访问故障处理](references/troubleshooting.md)
- [离线测试](tests/test_amazon_reviews.py)
- [CI 配置](.github/workflows/test.yml)

## 许可证

待补充。仓库当前未提供 `LICENSE` 文件，使用和分发授权范围需由维护者确认。

---

原 README 里的环境准备、登录步骤、完整命令与模式说明、输出契约、数据边界和当前状态等细节，已全部迁移到 [docs/详细说明.md](docs/详细说明.md)。
