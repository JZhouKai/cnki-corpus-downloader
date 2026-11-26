# CNKI Corpus Downloader

面向科研语料整理的知网批量下载工具，包含以下功能：

1. **DOI 流水线**：已知 DOI 的中文文献，按 `link.cnki.net/doi/{DOI}` 解析到详情页，再下载 PDF/CAJ。
2. **标题检索流水线**：没有 DOI 的中文文献，在官方知网按标题检索，核对标题/作者/期刊/年份后再下载。
3. **OCR 标注**（可选）：检测 PDF 是否含文本层，把纯图片 PDF 标记为需要 OCR，不做 OCR 处理。

> **免责声明**
>
> 本项目仅用于已获得授权的机构网络/账号下的科研数据整理。请遵守中国知网、学校图书馆的数据库授权协议和著作权规定，不要批量传播或再分发下载的文献。项目不包含任何验证码自动破解功能；遇到滑块或登录失效时需要人工处理。

---

## 环境要求

- Python 3.10+
- 可正常访问知网的校园网/机构账号
- Windows / macOS / Linux
- Playwright 使用的 Chromium（安装命令见下）

## 安装

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
playwright install chromium
```

## 准备数据

### 1. 语料 Excel

DOI 脚本默认读取项目根目录的 `未下载_含DOI.xlsx`（可修改脚本顶部的 `EXCEL_FILE` 常量）；标题检索脚本默认读取 `未下载_无DOI_1.xlsx`（可用 `--excel` 指定任意路径）。两个脚本都读取 Sheet 0。

需要包含以下列（列名可识别中英文别名）：

| 列   | 说明                                               |
| ---- | -------------------------------------------------- |
| 序号 | 用于生成文件名和进度键，建议唯一                   |
| 标题 | 中文标题                                           |
| 作者 | 用于检索结果核对                                   |
| 期刊 | 用于核对和加分                                     |
| 年份 | 用于核对和加分                                     |
| DOI  | 有 DOI 的行走 DOI 流水线；为空的行由标题流水线处理 |

### 2. 知网登录 Cookie

用 Cookie-Editor / EditThisCookie 等浏览器扩展登录知网后导出 JSON，保存为：

```
data/cnki_cookies.json
```

格式参考 `cnki_cookies.example.json`（示例中的值全部是占位符，不能直接使用）。

注意：

- 至少要包含 `.cnki.net` 下的登录 Cookie，例如 `Ecp_LoginStuts`、`LID`、`c_m_LinID`、`Ecp_session`、`fingerprint`；
- 如果扩展导出了 `kns.cnki.net`、`bar.cnki.net` 的 Cookie，也一并保留；
- **不要**从控制台 `document.cookie` 复制，会漏掉 HttpOnly Cookie；
- `data/` 已在 `.gitignore` 中，不会把 Cookie 提交到 Git。

## 使用

### A. 有 DOI 的文章：DOI 流水线

```bash
python -X utf8 -u download_corpus_pdfs.py --phase pipeline --chunk-size 20 --captcha-wait 30 --login-wait 300
```

- 每批解析 20 篇链接后立即下载，避免下载链接过期；
- 已下载/已有有效 PDF 的自动跳过；
- DOI 未注册/已失效、文章页没有 PDF/CAJ 按钮：记为**永久失败**，下次运行直接跳过；
- 登录失效、网络异常等**临时失败**：下次运行自动重试；
- 遇到滑块最多等 30 秒；登录失效最多等 300 秒；
- 想一直等人工处理，用 `--captcha-wait 0`；
- 想重新尝试已标记的永久失败，加 `--retry-failed`。

#### 断点与失败记录

每次运行结束（包括 Ctrl+C 中断）都会更新：

- `data/unresolved_articles.csv`：永久失败清单（DOI、标题、状态、原因、时间）；
- `data/download_manifest.csv`：全量清单，状态为 已下载 / 永久失败 / 待处理；
- `data/resume_checkpoint.json`：下一行序号、标题、DOI 和剩余数量。

下次运行会自动从第一个「待处理」条目继续，不会重复解析已下载或已标记失败的文章。临时失败（登录、网络）不会写入永久失败清单，下一轮仍会重试。

其他阶段：

```bash
# 只提取下载链接
python -X utf8 -u download_corpus_pdfs.py --phase extract --max-articles 100

# 查看统计
python -X utf8 -u download_corpus_pdfs.py --phase stats
```

### B. 没有 DOI 的文章：标题检索流水线

先小批量检索、不下载，确认匹配效果：

```bash
python -X utf8 -u download_corpus_by_title.py --phase search --max-articles 20 --captcha-wait 120
```

再正式检索 + 下载：

```bash
python -X utf8 -u download_corpus_by_title.py --phase all --max-articles 500 --captcha-wait 120 --login-wait 300 --retry-failed --no-annotate
```

常用参数：

| 参数                          | 说明                               |
| ----------------------------- | ---------------------------------- |
| `--phase search`            | 只检索、核对、写报告，不下载       |
| `--phase download`          | 复用已保存的详情页链接下载         |
| `--phase all`               | 检索 + 下载（默认）                |
| `--max-articles N`          | 本次最多处理 N 篇，分批跑时使用    |
| `--start-no` / `--end-no` | 只处理指定序号范围                 |
| `--search-limit 10`         | 每篇比对前 10 条检索结果           |
| `--min-score 0.60`          | 标题相似度阈值                     |
| `--captcha-wait`            | 滑块等待秒数，`0` 表示不限时     |
| `--login-wait`              | 重新登录等待秒数，`0` 表示不限时 |
| `--retry-failed`            | 重试之前失败的记录                 |
| `--no-annotate`             | 下载后不检测 PDF 文本层            |
| `--no-caj-fallback`         | 没有 PDF 时不回退 CAJ              |

下载优先级：**PDF 优先**；详情页没有 PDF 按钮时自动尝试 CAJ。CAJ 会保存为 `.caj`。

### C. OCR 标注

下载时默认会检测 PDF 是否含可提取文本层；也可以单独扫描已有 PDF：

```bash
python -X utf8 -u download_corpus_by_title.py --phase annotate
# 重新检测已有标注
python -X utf8 -u download_corpus_by_title.py --phase annotate --recheck
```

输出 `data/pdf_text_check.csv`：

| 字段               | 说明                    |
| ------------------ | ----------------------- |
| `pdf_type=text`  | 含文本层                |
| `pdf_type=image` | 图片型，`needs_ocr=Y` |
| `pdf_type=caj`   | CAJ 文件，无法直接检测  |
| `pdf_type=error` | 解析失败                |

检测只看前 3 页并提前退出，不会修改 PDF，也不会执行 OCR。

## 匹配规则（标题流水线）

1. 标题归一化：只保留中英文、数字，忽略标点/空格/大小写；
2. 相似度：互为子串记 `1.0`，否则使用 `difflib` 比率；
3. 候选标题分低于 `0.60` 直接淘汰；
4. 作者明确不一致直接淘汰；
5. 排名加分：作者 `+0.15`、年份 `+0.08`、期刊 `+0.05`；
6. 标题完全一致即接受；否则必须作者或年份匹配；
7. 不满足的记为 `no_match`，报告里会记录最佳分数、三项校验结果和 runner-up，便于人工复核。

## 输出文件

运行后 `data/` 下会产生：

```
data/
├── PDFs/                          下载的 PDF/CAJ
├── download_progress.pkl          DOI 流水线进度
├── download_urls.json             DOI 流水线 URL 缓存
├── unresolved_articles.csv        DOI 流水线永久失败清单
├── download_manifest.csv          DOI 流水线全量清单（已下载/永久失败/待处理）
├── resume_checkpoint.json         DOI 流水线断点（下一条处理位置）
├── download_*.log                 运行日志
├── title_search_progress.json     标题流水线进度快照
├── title_search_progress.jsonl    标题流水线逐篇日志
├── title_search_report.csv        标题流水线逐篇报告
├── pdf_text_check.csv             PDF 文本层/OCR 标注
├── browser_profile/               DOI 流水线浏览器配置（含登录态）
└── browser_profile_title/         标题流水线浏览器配置（含登录态）
```

`data/` 目录、PDF、日志、Cookie、浏览器配置都已在 `.gitignore` 中排除。

## 注意事项

- **不要同时运行两套流水线**：它们共用同一个机构账号，容易互相踢登录并显著增加验证码；
- 知网有下载频率风控，建议 `--delay 3` 或更高、每批 200–500 篇；临时失败下一轮重试，DOI 失效/无下载链接会被标记并跳过；
- 滑块和登录必须人工处理；本项目不会自动破解验证码；
- 标题流水线依赖官方知网 KNS8 页面结构。如果知网改版，可能需要调整脚本顶部的 `RESULT_ROW_SELECTORS` / `TITLE_LINK_SELECTORS`；
- CAJ 是知网私有格式，需要 CAJViewer 等工具查看；纯图片 PDF 需要 OCR 才能得到文本；
- 下载的文献仅供本人研究使用，请勿再分发。

## 常见问题

**Playwright 启动失败 / 找不到浏览器**

```bash
playwright install chromium
```

**详情页没有 PDF 按钮**

确认账号有下载权限、Cookie 未过期；可以先用浏览器手动打开一篇确认。无全文时会被记为 `no_pdf`，仍可尝试 CAJ（默认已开启兜底）。

**想重新下载之前跳过的文章**

运行 `--retry-failed` 会清空永久失败和下载失败记录，重新解析。注意 DOI 失效的文章仍会再次失败并重新标记。标题流水线也支持同样的参数。

**检索结果解析不到**

先运行 `--phase search --max-articles 3`，如果日志提示“未提取到标题/链接”，查看 `data/title_search_debug.png`，再调整脚本顶部的选择器。

**登录/滑块反复出现**

降低下载速度、分批运行，避免与其他客户端同时使用同一账号；需要人工处理时用 `--captcha-wait 0`。

## License

MIT License，见 [LICENSE](LICENSE)。
