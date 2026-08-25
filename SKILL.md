---
name: amazon-review-sample-collector
description: Use when collecting and validating Amazon competitor-review samples for one or more ASINs, especially requests mentioning 批量抓取竞品评论, review VOC, persistent Amazon session, resume, or Excel output.
---

# Amazon Review Sample Collector

Run the bundled CLI. Do not manually transcribe, invent, or describe the result as every review.

## Setup

From this Skill directory:

```bash
python -m venv .venv
python -m pip install -r requirements.txt
python -m playwright install chromium
python scripts/amazon_reviews.py doctor
python scripts/amazon_reviews.py init-profile
```

Use `.venv\Scripts\python.exe` on Windows or `.venv/bin/python` on macOS/Linux when the virtual environment is not activated. Wait for the user to finish Amazon sign-in during `init-profile`.

## Collect

```bash
python scripts/amazon_reviews.py rufus --asins ASIN_A ASIN_B --output OUTPUT_DIR
```

Use `--headed` for visible monitoring. Reusing `OUTPUT_DIR` resumes an interrupted run. Use `--restart` only when the user explicitly requests a fresh run.

## Validate

```bash
python scripts/amazon_reviews.py validate --output OUTPUT_DIR
```

Only report validation success when the command exits 0. Report the saved user-facing reason when collection is blocked.

## Output contract

The Excel and CSV columns remain:

`下载ASIN, 评论ID, 评论标题, 评论内容, 评论人, 评论人链接, 评分, 评论时间`

Generated public files exclude local paths, session details, request addresses, raw responses, request parameters, and browser credentials. Report ASIN counts, total unique Review IDs, absolute output paths, validation result, and this limitation:

> The dataset is a multi-keyword Amazon review sample, not every review for the product.

Read `references/output-schema.md` for fields and `references/troubleshooting.md` only when setup or collection is blocked.
