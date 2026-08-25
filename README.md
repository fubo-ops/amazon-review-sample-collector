# Amazon Review Sample Collector

一个可独立运行、也可作为 Codex Skill 使用的 Amazon 竞品评论样本采集器。输出为去重后的多关键词评论样本，不代表产品全部评论。

## 支持环境

- Windows、macOS、Linux
- Python 3.10–3.13
- 托管 Chromium 或本机 Google Chrome

仓库通过 GitHub Actions 在三个操作系统及 Python 最低、最高支持版本上测试。Excel 和 CSV 的字段、顺序及内容规则在各平台保持一致。

## 安装

```bash
git clone https://github.com/fubo-ops/amazon-review-sample-collector.git
cd amazon-review-sample-collector
python -m venv .venv
```

Windows PowerShell：

```powershell
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m playwright install chromium
```

macOS / Linux：

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
python -m playwright install chromium
```

检查运行环境：

```bash
python scripts/amazon_reviews.py doctor
```

## 首次登录

```bash
python scripts/amazon_reviews.py init-profile
```

在打开的浏览器中完成 Amazon 登录，然后回到终端按 Enter。会话保存在用户目录，不放入仓库或结果目录。

## 批量采集

```bash
python scripts/amazon_reviews.py rufus --asins ASIN_A ASIN_B --output outputs/run-001
```

需要查看浏览器过程时添加 `--headed`。中断后使用相同输出目录即可续跑；明确需要重新开始时添加 `--restart`。

## 验证

```bash
python scripts/amazon_reviews.py validate --output outputs/run-001
```

完成结果包含：

```text
ASIN_raw.json
评论合并.csv
评论合并.xlsx
汇总.json
采集日志.log
采集说明.md
```

Excel 和 CSV 固定为八列：

```text
下载ASIN, 评论ID, 评论标题, 评论内容, 评论人, 评论人链接, 评分, 评论时间
```

## 技术脱敏

- 不在结果文件中保存本地路径、会话目录、请求地址、请求参数、状态码或原始响应。
- 不读取、导出或打印 Cookie、令牌及浏览器凭据。
- JSON 中的评论字段与 Excel/CSV 保持一致。
- 完整运行结束后清除临时断点；被中断时仅保存恢复所需的最小状态。
- `.gitignore` 排除虚拟环境、输出目录、浏览器资料和缓存。

字段说明见 `references/output-schema.md`，故障处理见 `references/troubleshooting.md`。
