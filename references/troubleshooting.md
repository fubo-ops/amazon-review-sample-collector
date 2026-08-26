# 故障处理

## 依赖缺失

```bash
python -m pip install -r requirements.txt
python -m playwright install chromium
python scripts/amazon_reviews.py doctor
```

确认安装依赖与运行命令使用同一个 Python 解释器。

## 需要重新登录

```bash
python scripts/amazon_reviews.py init-profile
```

完成 Amazon 和卖家精灵登录后，使用原输出目录继续运行。若卖家精灵暂时不可用，混合模式会自动使用评论高频词二次扩量。

## 会话被占用

关闭仍在运行的采集浏览器或采集进程，再重新执行。不要同时启动两次同一会话。

## 找不到浏览器

```bash
python -m playwright install chromium
```

也可以安装 Google Chrome，并在采集命令中添加 `--browser chrome`。

## 访问受限或需要人工验证

稍后使用原输出目录继续运行；必要时降低为 `--asin-workers 1 --concurrency 2`。程序只保存恢复所需的最小状态。

## 验证失败

```bash
python scripts/amazon_reviews.py validate --output OUTPUT_DIR
```

根据返回的缺失文件、重复评论ID或行数不一致信息检查结果目录。
