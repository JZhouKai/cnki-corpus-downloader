#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
无 DOI 中文文章：知网标题检索 + 元数据核对 + PDF 下载

与 download_corpus_pdfs.py 完全独立：
- 使用独立的进度文件、日志、浏览器配置，不会影响正在运行的 DOI 任务；
- 仅处理 Excel Sheet 0「语料」中标题含中文且 DOI 为空的行；
- 在官方知网 https://kns.cnki.net 的 KNS8 检索页按标题检索；
- 按标题相似度 + 作者/期刊/年份核对候选；
- 打开详情页点击 #pdfDown 下载，文件保存到 data/PDFs/。

用法：
  # 先小批量试检索，不下载
  python -X utf8 -u download_corpus_by_title.py --phase search --max-articles 20 --captcha-wait 0

  # 检索 + 下载
  python -X utf8 -u download_corpus_by_title.py --phase all --max-articles 100 --captcha-wait 30

  # 大批量分批跑
  python -X utf8 -u download_corpus_by_title.py --phase all --max-articles 500 --captcha-wait 30

  # 重试上次失败的记录
  python -X utf8 -u download_corpus_by_title.py --phase all --retry-failed --captcha-wait 30

  # 扫描已有 PDF，标注文本型/图片型（需 OCR）
  python -X utf8 -u download_corpus_by_title.py --phase annotate
  python -X utf8 -u download_corpus_by_title.py --phase annotate --recheck

输出：
  data/PDFs/                          PDF/CAJ 文件（与 DOI 任务共用，文件名按序号唯一）
  data/title_search_progress.json     进度快照
  data/title_search_progress.jsonl    每篇追加的进度日志
  data/title_search_report.csv        逐篇结果报告（UTF-8 BOM）
  data/pdf_text_check.csv             PDF 文本层检测/OCR 标注
  data/title_search.log               运行日志
  data/browser_profile_title/         独立浏览器配置与登录态

下载优先级：
  优先 #pdfDown；没有 PDF 按钮时自动回退 #cajDown（--no-caj-fallback 可关闭）。
  CAJ 文件会保存为 .caj，并在标注里记为 pdf_type=caj。

OCR 标注：
  每篇下载完成后会自动检测 PDF 文本层，写入 data/pdf_text_check.csv：
    pdf_type=text     可提取文本
    pdf_type=image    图片型，needs_ocr=Y
    pdf_type=caj      CAJ 文件，无法直接检测
    pdf_type=error    解析失败
"""

import argparse
import csv
import difflib
import json
import logging
import os
import random
import re
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import openpyxl

# ===== 路径与常量 =====
SCRIPT_DIR = Path(__file__).resolve().parent
EXCEL_FILE = SCRIPT_DIR / "未下载_无DOI_1.xlsx"
OUTPUT_DIR = SCRIPT_DIR / "data"
PDF_DIR = OUTPUT_DIR / "PDFs"
PROGRESS_FILE = OUTPUT_DIR / "title_search_progress.json"
PROGRESS_JOURNAL = OUTPUT_DIR / "title_search_progress.jsonl"
REPORT_FILE = OUTPUT_DIR / "title_search_report.csv"
LOG_FILE = OUTPUT_DIR / "title_search.log"
DEBUG_SCREENSHOT = OUTPUT_DIR / "title_search_debug.png"
PROFILE_DIR = OUTPUT_DIR / "browser_profile_title"
COOKIES_FILE = OUTPUT_DIR / "cnki_cookies.json"
COOKIES_OUT_FILE = OUTPUT_DIR / "cnki_cookies_title.json"
ANNOTATION_FILE = OUTPUT_DIR / "pdf_text_check.csv"
ANNOTATION_FIELDS = [
    "file", "pages", "text_chars", "pdf_type", "needs_ocr", "error", "checked_at",
]

SEARCH_URL = "https://kns.cnki.net/kns8s/defaultresult/index"
CNKI_BASE = "https://kns.cnki.net/"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

RESULT_ROW_SELECTORS = (
    ".result-table-list tbody tr",
    "table.result-table-list tbody tr",
    "#gridTable tbody tr",
)
TITLE_LINK_SELECTORS = (
    "td.name a.fz14",
    "td.name a",
    "a.fz14",
)

REPORT_FIELDS = [
    "no", "title", "status", "matched_title", "score",
    "author_check", "journal_check", "year_check",
    "file", "message", "url", "time",
]

TERMINAL_STATUSES = {"downloaded", "caj", "skipped_existing"}
RETRYABLE_STATUSES = {
    "no_match", "no_pdf", "captcha", "login_required",
    "detail_mismatch", "detail_404", "search_failed", "failed",
}

# 日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("title_search")


# ===== 基础工具 =====
def ensure_dir(directory):
    Path(directory).mkdir(parents=True, exist_ok=True)


def now_iso():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def is_chinese(text):
    return bool(re.search(r"[\u4e00-\u9fff]", str(text)))


def sanitize_filename(name):
    name = re.sub(r'[<>:"/\\|?*]', "", str(name))
    name = re.sub(r"\s+", "_", name)[:70].strip("_")
    return name if name else "unnamed"


def make_filename(seq_no, title):
    try:
        no = int(seq_no)
    except (TypeError, ValueError):
        no = 0
    return f"{no:06d}_{sanitize_filename(title)}.pdf"


def atomic_write_json(path, data):
    ensure_dir(Path(path).parent)
    tmp = str(path) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def resolve_href(href, base=CNKI_BASE):
    """把检索结果里的相对链接解析为绝对链接。"""
    href = str(href or "").strip()
    if not href or href.lower().startswith("javascript:"):
        return ""
    if href.startswith("//"):
        return "https:" + href
    return urljoin(base, href)


# ===== Excel 读取 =====
def read_corpus_excel(filepath, sheet_index=0, only_chinese=True):
    """读取 Excel，返回中文文章列表（包含有/无 DOI 的行）。"""
    filepath = str(filepath)
    if not os.path.exists(filepath):
        log.error(f"文件不存在: {filepath}")
        sys.exit(1)

    log.info(f"读取文件: {filepath}")
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    if sheet_index < len(wb.sheetnames):
        ws = wb[wb.sheetnames[sheet_index]]
    else:
        ws = wb.active

    log.info(f"使用 Sheet[{sheet_index}]: {ws.title}")
    header = [cell.value for cell in ws[1]]
    log.info(f"列: {header}")

    def col_index(names, default):
        for name in names:
            if name in header:
                return header.index(name)
        return default

    i_no = col_index(["序号", "编号"], 0)
    i_title = col_index(["标题", "论文", "title"], 1)
    i_authors = col_index(["作者", "authors"], 2)
    i_journal = col_index(["期刊", "journal"], 3)
    i_year = col_index(["年份", "year"], 4)
    i_doi = col_index(["DOI", "doi"], 5)

    articles = []
    seen = set()
    dup_count = 0
    total_rows = ws.max_row - 1

    for row_idx, row in enumerate(
            ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True), 2):
        def cell(idx):
            return row[idx] if idx < len(row) else None

        title = str(cell(i_title) or "").strip()
        if not title:
            continue
        if only_chinese and not is_chinese(title):
            continue

        doi_val = cell(i_doi)
        doi = str(doi_val).strip() if doi_val else ""
        seq_val = cell(i_no)
        try:
            no = int(seq_val)
        except (TypeError, ValueError):
            no = len(articles) + 1

        key = str(no)
        if key in seen:
            key = f"{no}:{row_idx}"
            dup_count += 1
        seen.add(str(no))

        articles.append({
            "no": no,
            "key": key,
            "row": row_idx,
            "title": title,
            "authors": str(cell(i_authors) or "").strip(),
            "journal": str(cell(i_journal) or "").strip(),
            "year": str(cell(i_year) or "").strip(),
            "doi": doi or None,
            "has_doi": bool(doi),
        })

    wb.close()

    log.info(f"共 {total_rows} 行数据，中文文章 {len(articles)} 篇")
    doi_count = sum(1 for a in articles if a["has_doi"])
    log.info(f"其中有 DOI: {doi_count} 篇，无 DOI: {len(articles) - doi_count} 篇")
    if dup_count:
        log.warning(f"发现 {dup_count} 个重复序号，已用 '序号:行号' 作为进度键")
    return articles


# ===== 文件检查 =====
def is_valid_pdf(path):
    path = str(path)
    if not os.path.exists(path) or os.path.getsize(path) < 10240:
        return False
    with open(path, "rb") as f:
        return f.read(4) == b"%PDF"


def is_valid_caj(path):
    path = str(path)
    if not os.path.exists(path) or os.path.getsize(path) < 1024:
        return False
    with open(path, "rb") as f:
        return f.read(3) == b"CAJ"


def find_existing_file(article, scan_prefix=False):
    """返回已存在的有效 PDF/CAJ 路径，没有则 None。

    scan_prefix=True 时才按序号前缀扫描目录，用于处理 Excel 标题被改过的情况；
    默认只做精确路径检查，避免几万篇文章反复遍历整个 PDF 目录。
    """
    save_path = PDF_DIR / make_filename(article["no"], article["title"])
    if is_valid_pdf(save_path):
        return save_path
    caj_path = save_path.with_suffix(".caj")
    if is_valid_caj(caj_path):
        return caj_path

    if not scan_prefix:
        return None

    # 标题可能在 Excel 中做过修改，按序号前缀兜底查找
    prefix = f"{int(article['no']):06d}_"
    try:
        for f in sorted(PDF_DIR.glob(prefix + "*")):
            if f.suffix.lower() == ".pdf" and is_valid_pdf(f):
                return f
            if f.suffix.lower() == ".caj" and is_valid_caj(f):
                return f
    except OSError:
        pass
    return None


# ===== PDF 文本层检测 / OCR 标注 =====
_ANNOTATION_CACHE = None
_ANNOTATION_DIRTY = 0


def check_pdf_text(path):
    """检测 PDF 是否含可提取文本层；返回标注信息。"""
    try:
        from pypdf import PdfReader
    except ImportError:
        return {"pages": 0, "text_chars": 0, "pdf_type": "error",
                "needs_ocr": "unknown", "error": "pypdf未安装"}

    try:
        # 抑制 pypdf 的重复字典等 warning
        logging.getLogger("pypdf").setLevel(logging.ERROR)
        reader = PdfReader(str(path))
        pages = len(reader.pages)
        chars = 0
        # 只看前 3 页并提前退出：有文本层的文章首页通常就能判定
        for idx, pg in enumerate(reader.pages):
            if idx >= 3:
                break
            try:
                chars += len((pg.extract_text() or "").strip())
            except Exception:
                pass
            if chars >= 50:
                break
        pdf_type = "text" if chars >= 50 else "image"
        return {
            "pages": pages,
            "text_chars": chars,
            "pdf_type": pdf_type,
            "needs_ocr": "N" if pdf_type == "text" else "Y",
            "error": "",
        }
    except Exception as e:
        return {"pages": 0, "text_chars": 0, "pdf_type": "error",
                "needs_ocr": "unknown", "error": str(e)[:160]}


def get_annotation_cache():
    global _ANNOTATION_CACHE
    if _ANNOTATION_CACHE is not None:
        return _ANNOTATION_CACHE

    cache = {}
    if ANNOTATION_FILE.exists():
        try:
            with open(ANNOTATION_FILE, "r", encoding="utf-8-sig", newline="") as f:
                for row in csv.DictReader(f):
                    if row.get("file"):
                        cache[row["file"]] = row
        except Exception as e:
            log.warning(f"读取标注文件失败: {e}")
    _ANNOTATION_CACHE = cache
    return cache


def save_annotation_cache(cache=None):
    global _ANNOTATION_DIRTY
    cache = cache if cache is not None else get_annotation_cache()
    ensure_dir(OUTPUT_DIR)
    tmp = str(ANNOTATION_FILE) + ".tmp"
    with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f, fieldnames=ANNOTATION_FIELDS, lineterminator="\n")
        writer.writeheader()
        for name in sorted(cache):
            writer.writerow({k: cache[name].get(k, "") for k in ANNOTATION_FIELDS})
    os.replace(tmp, ANNOTATION_FILE)
    _ANNOTATION_DIRTY = 0


def annotate_pdf_file(path, cache=None, persist=False):
    """检测单个文件并写入缓存；每 50 个自动落盘。"""
    global _ANNOTATION_DIRTY
    path = Path(path)
    if cache is None:
        cache = get_annotation_cache()

    name = path.name
    if path.suffix.lower() == ".caj":
        info = {"pages": "", "text_chars": "", "pdf_type": "caj",
                "needs_ocr": "unknown", "error": ""}
    else:
        info = check_pdf_text(path)

    info["file"] = name
    info["checked_at"] = now_iso()
    cache[name] = info
    _ANNOTATION_DIRTY += 1

    if persist or _ANNOTATION_DIRTY >= 50:
        save_annotation_cache(cache)
    return info


def annotate_all(args):
    """扫描 data/PDFs，标注文本型/图片型/CAJ/异常，写出 pdf_text_check.csv。"""
    ensure_dir(PDF_DIR)
    ensure_dir(OUTPUT_DIR)
    cache = get_annotation_cache()

    files = sorted(PDF_DIR.glob("*.pdf")) + sorted(PDF_DIR.glob("*.caj"))
    if not files:
        log.info("PDF 目录为空，没有可标注的文件")
        return 0

    log.info(f"待标注文件: {len(files)} 个；已有标注: {len(cache)} 条")
    counters = Counter()
    processed = 0
    skipped = 0

    for i, path in enumerate(files, 1):
        name = path.name
        if not getattr(args, "recheck", False) and name in cache:
            counters[cache[name].get("pdf_type", "unknown")] += 1
            skipped += 1
            continue

        info = annotate_pdf_file(path, cache, persist=False)
        counters[info.get("pdf_type", "unknown")] += 1
        processed += 1

        if i % 100 == 0:
            save_annotation_cache(cache)
            log.info(f"标注进度: {i}/{len(files)}")

    save_annotation_cache(cache)

    log.info("=" * 60)
    log.info("PDF 文本层检测完成:")
    log.info(f"  本次新检测: {processed} 个；跳过已有: {skipped} 个")
    log.info(f"  文本型(text): {counters.get('text', 0)}")
    log.info(f"  图片型(需OCR): {counters.get('image', 0)}")
    log.info(f"  CAJ: {counters.get('caj', 0)}")
    log.info(f"  异常: {counters.get('error', 0)}")
    log.info(f"  标注文件: {ANNOTATION_FILE}")
    log.info("=" * 60)
    return 0


# ===== 进度 =====
def empty_progress():
    return {"version": 1, "updated": now_iso(), "records": {}}


def load_progress():
    if not PROGRESS_FILE.exists():
        return empty_progress()
    try:
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict) or not isinstance(data.get("records"), dict):
            raise ValueError("进度结构不正确")
        return data
    except Exception as e:
        log.warning(f"进度文件损坏，忽略: {e}")
        return empty_progress()


def merge_journal_into_progress(progress):
    if not PROGRESS_JOURNAL.exists():
        return progress
    merged = 0
    try:
        with open(PROGRESS_JOURNAL, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except json.JSONDecodeError:
                    continue
                key = str(rec.get("key") or rec.get("no", ""))
                if not key:
                    continue
                old = progress["records"].get(key)
                if old and str(old.get("updated", "")) > str(rec.get("updated", "")):
                    continue
                progress["records"][key] = rec
                merged += 1
    except OSError as e:
        log.warning(f"读取进度日志失败: {e}")
    if merged:
        log.info(f"从进度日志恢复了 {merged} 条记录")
    return progress


def append_progress_journal(record):
    ensure_dir(OUTPUT_DIR)
    try:
        with open(PROGRESS_JOURNAL, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
            f.flush()
    except OSError as e:
        log.warning(f"写进度日志失败: {e}")


def save_progress(progress, compact=False):
    progress["version"] = 1
    progress["updated"] = now_iso()
    atomic_write_json(PROGRESS_FILE, progress)
    if compact:
        try:
            with open(PROGRESS_JOURNAL, "w", encoding="utf-8"):
                pass
        except OSError:
            pass


def record_progress(progress, article, status, **kwargs):
    rec = {
        "no": article["no"],
        "key": article["key"],
        "title": article["title"],
        "status": status,
        "updated": now_iso(),
    }
    rec.update(kwargs)
    progress["records"][article["key"]] = rec
    append_progress_journal(rec)
    return rec


# ===== 报告 =====
def init_report(reset=False):
    ensure_dir(OUTPUT_DIR)
    need_header = reset or not REPORT_FILE.exists() or REPORT_FILE.stat().st_size == 0
    if need_header:
        with open(REPORT_FILE, "w", encoding="utf-8-sig", newline="") as f:
            csv.DictWriter(f, fieldnames=REPORT_FIELDS).writeheader()


def append_report_row(row):
    with open(REPORT_FILE, "a", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=REPORT_FIELDS, lineterminator="\n")
        writer.writerow(row)
        f.flush()


def format_check(value):
    if value is True:
        return "Y"
    if value is False:
        return "N"
    return "-"


def build_report_row(article, status, best=None, score=0.0, checks=None,
                     file="", message="", url="", runner_up=None):
    checks = checks or {}
    message = message or ""
    if runner_up:
        message += (f"; runner_up={str(runner_up.get('title', ''))[:40]}"
                    f"(score={float(runner_up.get('ts', 0.0)):.2f})")
    return {
        "no": str(article.get("no", "")),
        "title": article.get("title", ""),
        "status": status,
        "matched_title": (best or {}).get("title", ""),
        "score": f"{float(score):.3f}",
        "author_check": format_check(checks.get("author")),
        "journal_check": format_check(checks.get("journal")),
        "year_check": format_check(checks.get("year")),
        "file": file or "",
        "message": message,
        "url": url or "",
        "time": now_iso(),
    }


# ===== 标题/元数据匹配 =====
def normalize_title(s):
    return re.sub(r"[^\w\u4e00-\u9fff]", "", str(s)).lower()


def title_score(query, candidate):
    a, b = normalize_title(query), normalize_title(candidate)
    if not a or not b:
        return 0.0
    if a in b or b in a:
        return 1.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def parse_author_names(s):
    parts = re.split(r"[;；,，、/\s]+", str(s))
    return [n for n in (normalize_title(x) for x in parts) if n]


def extract_year(s):
    m = re.search(r"(19|20)\d{2}", str(s))
    return m.group(0) if m else None


def verify_candidate(article, row):
    """比对 Excel（标题/作者/期刊/年份）与检索结果，返回 (标题分, 校验字典)。"""
    ts = title_score(article.get("title", ""), row.get("title", ""))
    checks = {"author": None, "journal": None, "year": None}

    if article.get("authors") and row.get("author"):
        names = parse_author_names(article["authors"])
        result_author = normalize_title(row["author"])
        if names:
            checks["author"] = any(n in result_author for n in names)

    if article.get("journal") and row.get("source"):
        a = normalize_title(article["journal"])
        b = normalize_title(row["source"])
        if a and b:
            checks["journal"] = (a in b) or (b in a)

    ay = extract_year(article.get("year"))
    ry = extract_year(row.get("date"))
    if ay and ry:
        checks["year"] = (ay == ry)

    return ts, checks


def pick_best(article, rows, min_score):
    """返回 (best, best_ts, best_checks, runner_up)。"""
    evaluated = []
    best = None
    best_rank = -1.0
    best_ts = 0.0
    best_checks = None

    for row in rows:
        ts, checks = verify_candidate(article, row)
        best_ts = max(best_ts, ts)
        if ts < min_score:
            continue
        if checks.get("author") is False:
            continue

        rank = ts
        if checks.get("author"):
            rank += 0.15
        if checks.get("year"):
            rank += 0.08
        if checks.get("journal"):
            rank += 0.05

        evaluated.append((rank, ts, row, checks))
        if rank > best_rank:
            best_rank = rank
            best = row
            best_checks = checks

    evaluated.sort(key=lambda x: x[0], reverse=True)
    runner_up = None
    for rank, ts, row, checks in evaluated:
        if best is not None and row is best:
            continue
        runner_up = {"title": row.get("title", ""), "ts": ts}
        break

    return best, best_ts, best_checks, runner_up


def acceptance(article, best, ts, checks):
    if not best:
        return False, f"no_match(ts={ts:.2f})"
    checks = checks or {}
    if checks.get("author") is False:
        return False, "author_mismatch"

    title_exact = normalize_title(article.get("title", "")) == normalize_title(best.get("title", ""))
    strong_aux = checks.get("author") is True or checks.get("year") is True

    if title_exact:
        return True, "exact_title"
    if strong_aux:
        return True, f"title+aux(ts={ts:.2f})"
    return False, (
        f"verify_failed(ts={ts:.2f},author={checks.get('author')},"
        f"journal={checks.get('journal')},year={checks.get('year')})"
    )


# ===== Cookie =====
def build_browser_cookies():
    cookies = []
    cookie_str = os.environ.get("CNKI_COOKIES", "").strip()
    if cookie_str:
        for item in cookie_str.split(";"):
            item = item.strip()
            if "=" in item:
                name, value = item.split("=", 1)
                cookies.append({
                    "name": name.strip(), "value": value.strip(),
                    "domain": ".cnki.net", "path": "/",
                })
        return cookies

    # 优先使用较新的 Cookie 文件；--write-cookies 的输出也可直接复用
    cookie_file = COOKIES_FILE
    try:
        if COOKIES_OUT_FILE.exists() and (
                not COOKIES_FILE.exists()
                or COOKIES_OUT_FILE.stat().st_mtime > COOKIES_FILE.stat().st_mtime):
            cookie_file = COOKIES_OUT_FILE
    except OSError:
        pass

    if cookie_file.exists():
        try:
            with open(cookie_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            for c in data:
                if c.get("name") and c.get("value") is not None:
                    cookies.append({
                        "name": c["name"],
                        "value": str(c["value"]),
                        "domain": c.get("domain") or ".cnki.net",
                        "path": c.get("path") or "/",
                    })
            log.info(f"从 {cookie_file.name} 读取 Cookie")
        except Exception as e:
            log.warning(f"加载Cookie文件失败: {e}")
    return cookies


def inject_cookies_into_context(context):
    cookies = build_browser_cookies()
    if cookies:
        context.add_cookies(cookies)
        log.info(f"已注入 {len(cookies)} 个Cookie")
    return len(cookies)


def maybe_save_cookies(context, args):
    """仅在 --write-cookies 时写入独立文件，绝不改动共享 Cookie。"""
    if not getattr(args, "write_cookies", False):
        return
    try:
        cookies = [c for c in context.cookies()
                   if "cnki.net" in (c.get("domain") or "")]
        if not cookies:
            return
        data = [{
            "name": c["name"],
            "value": c["value"],
            "domain": c.get("domain") or ".cnki.net",
            "path": c.get("path") or "/",
            "expires": None if not c.get("expires") or c["expires"] < 0 else c["expires"],
        } for c in cookies]
        atomic_write_json(COOKIES_OUT_FILE, data)
        log.info(f"已写入独立Cookie文件: {COOKIES_OUT_FILE}（{len(data)} 个）")
    except Exception as e:
        log.warning(f"保存Cookie失败: {e}")


# ===== 浏览器状态与等待 =====
def page_state(page):
    """返回 verify / login / ok / closed"""
    try:
        url = (page.url or "").lower()
        title = page.title() or ""
    except Exception:
        return "closed"

    if (any(k in url for k in ("/verify", "/bar/verify", "captcha"))
            or any(k in title for k in ("安全验证", "拼图校验", "验证码"))):
        return "verify"

    # 只有明确的登录域名/路径才判登录，不能搜 URL 任意位置含 login
    try:
        parsed = urlparse(url)
        host = (parsed.netloc or "").lower()
        path = (parsed.path or "").lower()
    except Exception:
        host, path = "", ""
    if (host.endswith("login.cnki.net") or "passport" in host
            or path.startswith("/login") or path.startswith("/user/login")):
        return "login"

    # 真正的内容标记优先，避免文章标题里含「登录」时被误判
    try:
        if page.locator("#pdfDown, #cajDown, #txt_search, "
                        ".result-table-list, #gridTable").count() > 0:
            return "ok"
        if page.locator("#TextBoxUserName").count() > 0:
            return "login"
    except Exception:
        pass

    if "登录" in title and "拼图" not in title and "安全验证" not in title:
        return "login"
    return "ok"


def looks_like_verification_text(text):
    keys = ("安全验证", "拼图校验", "拖动滑块", "滑块验证", "请完成验证",
            "人机验证", "验证码", "腾讯验证")
    return any(k in text for k in keys)


def wait_for_search_ready(page, args, hard_timeout=30):
    """等待搜索框就绪；期间允许人工通过滑块/重新登录。"""
    start = time.time()
    captcha_deadline = None
    login_deadline = None
    noticed_verify = False
    noticed_login = False
    reloaded = False

    while True:
        if page.is_closed():
            return False

        state = page_state(page)
        if state == "ok":
            try:
                if page.locator("#txt_search").count() > 0:
                    return True
            except Exception:
                return False
            # 登录/验证完成后可能落在别的页面，回检索页一次
            if (noticed_login or noticed_verify) and not reloaded:
                try:
                    page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=60000)
                    reloaded = True
                except Exception:
                    return False
        elif state == "verify":
            if not noticed_verify:
                desc = "不限时" if args.captcha_wait <= 0 else f"最长等待 {args.captcha_wait} 秒"
                log.info(f"  → 检索页出现安全验证，请在浏览器中完成滑块（{desc}）")
                noticed_verify = True
                captcha_deadline = (None if args.captcha_wait <= 0
                                    else time.time() + args.captcha_wait)
            if captcha_deadline is not None and time.time() > captcha_deadline:
                log.warning("  ! 安全验证等待超时")
                return False
        elif state == "login":
            if not noticed_login:
                desc = "不限时" if args.login_wait <= 0 else f"最长等待 {args.login_wait} 秒"
                log.info(f"  → 登录已失效，请在浏览器中登录（{desc}）")
                noticed_login = True
                login_deadline = (None if args.login_wait <= 0
                                  else time.time() + args.login_wait)
            if login_deadline is not None and time.time() > login_deadline:
                log.warning("  ! 登录等待超时")
                return False

        if not noticed_verify and not noticed_login and time.time() - start > hard_timeout:
            return False

        try:
            page.wait_for_timeout(500)
        except Exception:
            return False


def get_detail_title(page):
    try:
        title = (page.title() or "").strip()
    except Exception:
        return None
    title = re.sub(r"\s*[-_|]\s*(中国知网|CNKI)\s*$", "", title, flags=re.I).strip()

    bad = (not title or title in ("中国知网", "CNKI")
           or any(k in title for k in ("安全验证", "拼图校验", "登录", "404")))
    if not bad:
        return title

    # 兜底：meta citation_title / h1
    for sel, attr in (('meta[name="citation_title"]', "content"), ("h1", None)):
        try:
            loc = page.locator(sel).first
            if loc.count() > 0:
                value = loc.get_attribute(attr) if attr else loc.inner_text()
                value = (value or "").strip()
                if value and not any(k in value for k in ("安全验证", "登录")):
                    return value
        except Exception:
            pass
    return None


def wait_for_detail_ready(page, args, hard_timeout=45):
    """等待详情页可读；期间允许人工通过滑块/重新登录。

    发现验证/登录后按 --captcha-wait / --login-wait 等待（0 = 不限时），
    不再被 hard_timeout 提前打断。
    """
    start = time.time()
    captcha_deadline = None
    login_deadline = None
    noticed_verify = False
    noticed_login = False

    while True:
        if page.is_closed():
            return False

        state = page_state(page)
        if state == "ok":
            try:
                if page.locator("#pdfDown, #cajDown").count() > 0:
                    return True
                if get_detail_title(page):
                    return True
            except Exception:
                return False
            if time.time() - start > 15:
                return True
        elif state == "verify":
            if not noticed_verify:
                desc = "不限时" if args.captcha_wait <= 0 else f"最长等待 {args.captcha_wait} 秒"
                log.info(f"  → 详情页出现安全验证，请在浏览器中完成滑块（{desc}）")
                noticed_verify = True
                captcha_deadline = (None if args.captcha_wait <= 0
                                    else time.time() + args.captcha_wait)
            if captcha_deadline is not None and time.time() > captcha_deadline:
                return False
        elif state == "login":
            if not noticed_login:
                desc = "不限时" if args.login_wait <= 0 else f"最长等待 {args.login_wait} 秒"
                log.info(f"  → 登录已失效，请在浏览器中登录（{desc}）")
                noticed_login = True
                login_deadline = (None if args.login_wait <= 0
                                  else time.time() + args.login_wait)
            if login_deadline is not None and time.time() > login_deadline:
                return False
        else:
            return False

        # 尚未出现验证/登录时，只用 hard_timeout 兜底
        if (not noticed_verify and not noticed_login
                and time.time() - start > hard_timeout):
            return False

        try:
            page.wait_for_timeout(500)
        except Exception:
            return False


def wait_for_download(context, downloads, start_index, timeout, login_timeout=None):
    """等待 context 产生新的下载事件；期间处理滑块/登录页。"""
    deadline = None if timeout <= 0 else time.time() + timeout
    noticed_verify = False
    noticed_login = False

    while deadline is None or time.time() < deadline:
        if len(downloads) > start_index:
            return downloads[start_index]

        pages = context.pages
        if not noticed_verify and any(
                page_state(p) == "verify" for p in pages):
            desc = "不限时" if timeout <= 0 else f"最长等待 {timeout} 秒"
            log.info(f"  → 请在浏览器窗口中拖动滑块完成验证（{desc}）")
            noticed_verify = True

        if not noticed_login and any(page_state(p) == "login" for p in pages):
            login_timeout = 300 if login_timeout is None else login_timeout
            deadline = None if login_timeout <= 0 else time.time() + login_timeout
            desc = "不限时" if login_timeout <= 0 else f"最长等待 {login_timeout} 秒"
            log.info(f"  → 登录已失效，请在浏览器中登录，登录后会自动继续（{desc}）")
            noticed_login = True

        if pages:
            try:
                pages[0].wait_for_timeout(300)
                continue
            except Exception:
                pass
        time.sleep(0.3)

    return downloads[start_index] if len(downloads) > start_index else None


def save_browser_download(download, save_path):
    """保存 Playwright 下载，并校验文件头。"""
    save_path = str(save_path)
    tmp_path = save_path + ".part"
    try:
        download.save_as(tmp_path)
    except Exception as e:
        return "error", f"保存失败: {e}"

    try:
        with open(tmp_path, "rb") as f:
            head = f.read(4)
        size = os.path.getsize(tmp_path)

        if head == b"%PDF":
            os.replace(tmp_path, save_path)
            return "ok", f"PDF saved ({size} bytes)"

        if head[:3] == b"CAJ":
            caj_path = save_path[:-4] + ".caj" if save_path.lower().endswith(".pdf") else save_path + ".caj"
            os.replace(tmp_path, caj_path)
            return "caj", f"CAJ saved ({size} bytes)"
    finally:
        if os.path.exists(tmp_path) and not os.path.exists(save_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass

    return "invalid", "下载内容不是PDF/CAJ"


def save_bytes_validated(body, save_path):
    save_path = str(save_path)
    tmp_path = save_path + ".part"
    with open(tmp_path, "wb") as f:
        f.write(body)
    if body[:4] == b"%PDF":
        os.replace(tmp_path, save_path)
        return "ok", f"PDF saved ({len(body)} bytes)"
    if body[:3] == b"CAJ":
        caj_path = save_path[:-4] + ".caj" if save_path.lower().endswith(".pdf") else save_path + ".caj"
        os.replace(tmp_path, caj_path)
        return "caj", f"CAJ saved ({len(body)} bytes)"
    try:
        os.remove(tmp_path)
    except OSError:
        pass
    return "invalid", "下载内容不是PDF/CAJ"


def close_extra_pages(context, keep_page):
    for extra in list(context.pages):
        if extra is not keep_page:
            try:
                extra.close()
            except Exception:
                pass


def capture_debug(page):
    try:
        page.screenshot(path=str(DEBUG_SCREENSHOT), full_page=False)
        log.warning(f"已保存调试截图: {DEBUG_SCREENSHOT}")
    except Exception:
        pass
    try:
        log.warning(f"当前页面: {page.url} | 标题: {page.title()}")
    except Exception:
        pass


# ===== 检索 =====
def extract_result_rows(page, limit):
    js = r"""
    (els, arg) => {
      const limit = arg.limit, sel = arg.sel;
      return els.slice(0, limit).map((e, i) => {
        const a = e.querySelector(sel);
        if (!a) return null;
        const txt = s => {
          const n = e.querySelector(s);
          return n ? n.innerText.replace(/\s+/g, ' ').trim() : '';
        };
        return {
          title: (a.innerText || '').trim(),
          href: a.getAttribute('href') || a.getAttribute('data-href') || '',
          author: txt('td.author'),
          source: txt('td.source'),
          date: txt('td.date'),
          db: txt('td.data'),
          rank: i + 1
        };
      }).filter(x => x && x.title && x.href);
    }
    """
    for row_sel in RESULT_ROW_SELECTORS:
        try:
            if page.locator(row_sel).count() == 0:
                continue
        except Exception:
            continue
        for link_sel in TITLE_LINK_SELECTORS:
            try:
                rows = page.eval_on_selector_all(
                    row_sel, js, {"limit": int(limit), "sel": link_sel})
                if rows:
                    return rows
            except Exception:
                continue
    return []


def wait_for_results(page, args, timeout=20):
    """返回 rows / empty / retry / captcha / login / timeout / closed"""
    deadline = time.time() + timeout
    ticks = 0

    while time.time() < deadline:
        try:
            for sel in RESULT_ROW_SELECTORS:
                if page.locator(sel).count() > 0:
                    return "rows"
        except Exception:
            return "closed"

        state = page_state(page)
        if state == "verify":
            log.info("  → 检索页需要安全验证")
            if wait_for_search_ready(page, args):
                return "retry"
            return "captcha"
        if state == "login":
            log.info("  → 检索页需要重新登录")
            if wait_for_search_ready(page, args):
                return "retry"
            return "login"

        ticks += 1
        if ticks % 4 == 0:
            try:
                text = page.locator("body").inner_text(timeout=2000)
                if any(k in text for k in ("没有找到", "未找到相关", "检索结果为空", "没有检索到")):
                    return "empty"
            except Exception:
                pass

        try:
            page.wait_for_timeout(300)
        except Exception:
            return "closed"

    return "timeout"


def search_rows(page, query, args):
    """按标题检索，返回 (rows, error)。error 为 None/captcha/login/search_failed"""
    for attempt in (1, 2):
        try:
            if page.is_closed():
                return [], "search_failed"

            on_search_page = ("kns8s" in (page.url or "")
                              and page.locator("#txt_search").count() > 0)
            if not on_search_page:
                page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=60000)

            if not wait_for_search_ready(page, args):
                state = page_state(page)
                if state == "verify":
                    return [], "captcha"
                if state == "login":
                    return [], "login"
                return [], "search_failed"

            page.wait_for_selector("#txt_search", timeout=15000)
            page.wait_for_timeout(800)

            box = page.locator("#txt_search").first
            box.click(timeout=5000)
            box.fill("")
            box.fill(query)
            box.press("Enter")

            outcome = wait_for_results(page, args, timeout=20)

            # Enter 没触发时，尝试点击检索按钮
            if outcome == "timeout":
                clicked = False
                for sel in ("#btn_search", ".search-btn",
                            "input[type='button'][value='检索']",
                            "button:has-text('检索')", "a:has-text('检索')"):
                    try:
                        if page.locator(sel).count() > 0:
                            page.locator(sel).first.click(timeout=3000)
                            clicked = True
                            break
                    except Exception:
                        continue
                if clicked:
                    outcome = wait_for_results(page, args, timeout=20)

            if outcome == "rows":
                page.wait_for_timeout(800)
                rows = extract_result_rows(page, args.search_limit)
                if rows:
                    return rows, None
                try:
                    text = page.locator("body").inner_text(timeout=2000)
                except Exception:
                    text = ""
                if any(k in text for k in ("没有找到", "未找到相关", "检索结果为空", "没有检索到")):
                    return [], None
                log.warning("  检索结果行存在，但未提取到标题/链接")
                capture_debug(page)
                return [], "search_failed"

            if outcome == "empty":
                return [], None
            if outcome == "retry":
                continue
            if outcome in ("captcha", "login"):
                return [], outcome
            if outcome == "closed":
                return [], "search_failed"

            log.warning(f"  检索第 {attempt} 次超时")

        except Exception as e:
            log.warning(f"  检索第 {attempt} 次异常: {e}")

        # 第二次尝试前强制回到检索页
        try:
            page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass

    capture_debug(page)
    return [], "search_failed"


# ===== 详情页与下载 =====
def open_detail_page(context, url, args):
    page = context.new_page()
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60000)
    except Exception as e:
        log.warning(f"  打开详情页失败: {e}")
        try:
            page.close()
        except Exception:
            pass
        return None, "failed"

    if not wait_for_detail_ready(page, args):
        state = page_state(page)
        try:
            page.close()
        except Exception:
            pass
        if state == "verify":
            return None, "captcha"
        if state == "login":
            return None, "login_required"
        return None, "failed"
    return page, None


def verify_detail_page(article, page, min_score):
    detail_title = get_detail_title(page)
    if not detail_title:
        try:
            has_button = page.locator("#pdfDown, #cajDown").count() > 0
        except Exception:
            has_button = False
        if has_button:
            return True, "button_only", 0.0, ""
        return False, "detail_title_missing", 0.0, ""

    score = title_score(article.get("title", ""), detail_title)
    if normalize_title(article.get("title", "")) == normalize_title(detail_title):
        return True, "detail_exact_title", score, detail_title
    if score >= min_score:
        return True, f"detail_title(ts={score:.2f})", score, detail_title
    return False, f"detail_mismatch(ts={score:.2f})", score, detail_title


def try_request_download(context, url, referer, save_path):
    try:
        resp = context.request.get(
            url,
            headers={"Referer": referer or SEARCH_URL, "User-Agent": UA},
            timeout=90000,
            max_redirects=20,
        )
        body = resp.body()
        if body[:4] == b"%PDF" or body[:3] == b"CAJ":
            return save_bytes_validated(body, save_path)
        text = body.decode("utf-8", "ignore")
        if "/verify" in resp.url or looks_like_verification_text(text):
            return "captcha", "需要人工验证"
        return "failed", f"HTTP响应不是文件（status={resp.status}）"
    except Exception as e:
        return "failed", f"请求下载失败: {e}"


def download_from_detail(context, page, downloads, save_path, args):
    """返回 (status, message, url)。status 为 downloaded/caj/skipped_existing/no_pdf/captcha/login_required/failed"""
    save_path = str(save_path)
    if is_valid_pdf(save_path):
        return "skipped_existing", "PDF已存在", page.url
    caj_path = save_path[:-4] + ".caj"
    if is_valid_caj(caj_path):
        return "skipped_existing", "CAJ已存在", page.url

    # 详情页按钮偶尔比标题晚渲染，短暂等待
    try:
        page.wait_for_selector("#pdfDown, #cajDown", timeout=8000)
    except Exception:
        pass

    pdf_loc = page.locator("#pdfDown")
    caj_loc = page.locator("#cajDown")

    target = None
    if pdf_loc.count() > 0:
        target = pdf_loc.first
    elif (not getattr(args, "no_caj_fallback", False)
          and caj_loc.count() > 0):
        target = caj_loc.first
        log.info("  PDF 不可用，改用 CAJ")
    else:
        return "no_pdf", "详情页没有可用的PDF/CAJ下载按钮", page.url

    try:
        href = target.get_attribute("href") or ""
    except Exception:
        href = ""

    before = len(downloads)
    try:
        target.click(timeout=15000)
    except Exception as e:
        log.warning(f"  点击下载按钮失败: {e}")

    download = wait_for_download(
        context, downloads, before, args.captcha_wait, args.login_wait
    )

    if download is not None:
        status, msg = save_browser_download(download, save_path)
        mapping = {"ok": "downloaded", "caj": "caj",
                   "invalid": "failed", "error": "failed"}
        return mapping.get(status, status), msg, page.url

    states = [page_state(p) for p in context.pages]
    if "verify" in states:
        return "captcha", "等待验证超时", page.url
    if "login" in states:
        return "login_required", "登录失效", page.url

    # 兜底：共享 Cookie 的请求通道
    if href:
        abs_url = resolve_href(href, page.url or CNKI_BASE)
        status, msg = try_request_download(
            context, abs_url, page.url, save_path)
        if status in ("ok", "caj", "captcha"):
            mapping = {"ok": "downloaded", "caj": "caj", "captcha": "captcha"}
            return mapping[status], msg, abs_url
        log.warning(f"  请求通道兜底失败: {msg}")

    return "failed", "未捕获下载事件", page.url


# ===== 单篇处理 =====
def finish_download(context, search_page, downloads, article, best, ts,
                    checks, runner_up, detail_url, args, progress):
    detail_page, derr = open_detail_page(context, detail_url, args)
    if derr:
        rec = record_progress(
            progress, article, derr,
            matched_title=best.get("title", ""),
            score=round(ts, 3), checks=checks, detail_url=detail_url,
            message=f"打开详情页失败: {derr}",
        )
        return rec, build_report_row(
            article, derr, best=best, score=ts, checks=checks,
            message=rec["message"], url=detail_url, runner_up=runner_up)

    try:
        ok, reason, dscore, dtitle = verify_detail_page(
            article, detail_page, args.min_score)
        if not ok:
            status = "detail_mismatch"
            rec = record_progress(
                progress, article, status,
                matched_title=best.get("title", ""),
                score=round(dscore or ts, 3), checks=checks,
                detail_url=detail_url, message=reason,
            )
            return rec, build_report_row(
                article, status, best=best, score=dscore or ts, checks=checks,
                message=reason, url=detail_url, runner_up=runner_up)

        save_path = PDF_DIR / make_filename(article["no"], article["title"])
        status, msg, url = download_from_detail(
            context, detail_page, downloads, save_path, args)

        file_name = ""
        pdf_info = {}
        existing = find_existing_file(article)
        if existing:
            file_name = existing.name
            if (status in ("downloaded", "caj")
                    and not getattr(args, "no_annotate", False)):
                info = annotate_pdf_file(existing)
                pdf_info = {
                    "pdf_type": info.get("pdf_type", ""),
                    "needs_ocr": info.get("needs_ocr", ""),
                }
                if info.get("needs_ocr") == "Y":
                    log.info(f"  ⚠ 图片型PDF，需OCR（可提取文本 "
                             f"{info.get('text_chars', 0)} 字符）")

        rec = record_progress(
            progress, article, status,
            matched_title=best.get("title", ""),
            score=round(ts, 3), checks=checks,
            detail_url=detail_url, file=file_name, message=msg,
            **pdf_info,
        )
        return rec, build_report_row(
            article, status, best=best, score=ts, checks=checks,
            file=file_name, message=msg, url=url, runner_up=runner_up)
    finally:
        close_extra_pages(context, search_page)


def process_article_all(context, search_page, downloads, article, args, progress):
    existing = find_existing_file(article)
    if existing:
        rec = record_progress(
            progress, article, "skipped_existing",
            file=existing.name, message="already_exists")
        return rec, None

    rows, err = search_rows(search_page, article["title"], args)
    if err:
        status = {"captcha": "captcha", "login": "login_required"}.get(
            err, "search_failed")
        rec = record_progress(progress, article, status, message=err)
        return rec, build_report_row(article, status, message=err)

    if not rows:
        rec = record_progress(progress, article, "no_match", message="无检索结果")
        return rec, build_report_row(article, "no_match", message="无检索结果")

    best, ts, checks, runner_up = pick_best(article, rows, args.min_score)
    ok, reason = acceptance(article, best, ts, checks)
    if not ok:
        rec = record_progress(
            progress, article, "no_match",
            matched_title=(best or {}).get("title", ""),
            score=round(ts, 3), checks=checks, message=reason,
        )
        return rec, build_report_row(
            article, "no_match", best=best, score=ts, checks=checks,
            message=reason, runner_up=runner_up)

    detail_url = resolve_href(best.get("href", ""), CNKI_BASE)
    rec = record_progress(
        progress, article, "matched",
        matched_title=best.get("title", ""),
        score=round(ts, 3), checks=checks,
        detail_url=detail_url, message=reason,
    )

    if args.dry_run or args.phase == "search":
        suffix = " (dry-run)" if args.dry_run else ""
        return rec, build_report_row(
            article, "matched", best=best, score=ts, checks=checks,
            message=reason + suffix, url=detail_url, runner_up=runner_up)

    return finish_download(
        context, search_page, downloads, article, best, ts,
        checks, runner_up, detail_url, args, progress)


def process_article_download(context, search_page, downloads, article, args, progress):
    rec = progress["records"].get(article["key"]) or {}
    detail_url = rec.get("detail_url")

    if not detail_url or rec.get("status") not in ("matched", "downloaded", "caj"):
        # 没有可复用的详情页链接，回退到重新检索
        return process_article_all(
            context, search_page, downloads, article, args, progress)

    existing = find_existing_file(article)
    if existing:
        rec = record_progress(
            progress, article, "skipped_existing",
            file=existing.name, message="already_exists")
        return rec, None

    best = {"title": rec.get("matched_title", "")}
    ts = float(rec.get("score") or 0.0)
    checks = rec.get("checks") or {}
    result, row = finish_download(
        context, search_page, downloads, article, best, ts,
        checks, None, detail_url, args, progress)

    # 详情页链接过期/错配时，重新检索一次
    if result.get("status") in ("detail_mismatch", "failed", "detail_404"):
        log.info("  详情页链接失效，重新检索")
        return process_article_all(
            context, search_page, downloads, article, args, progress)
    return result, row


# ===== 待处理筛选 =====
def select_pending_articles(articles, progress, args):
    pending = []
    skipped = Counter()
    records = progress["records"]

    for article in articles:
        if article["has_doi"]:
            continue
        if args.start_no is not None and article["no"] < args.start_no:
            continue
        if args.end_no is not None and article["no"] > args.end_no:
            continue

        rec = records.get(article["key"])
        status = rec.get("status") if rec else None

        existing = find_existing_file(
            article, scan_prefix=status in TERMINAL_STATUSES)
        if existing:
            old = rec or {}
            if old.get("status") not in TERMINAL_STATUSES:
                rec = {
                    "no": article["no"],
                    "key": article["key"],
                    "title": article["title"],
                    "status": "skipped_existing",
                    "file": existing.name,
                    "message": "already_exists",
                    "updated": now_iso(),
                }
                records[article["key"]] = rec
                append_progress_journal(rec)
            continue

        if status in TERMINAL_STATUSES:
            # 进度说已完成，但文件不存在/损坏 → 重新处理
            records.pop(article["key"], None)
            pending.append(article)
            continue

        if status in RETRYABLE_STATUSES and not args.retry_failed:
            skipped[status] += 1
            continue

        if status == "matched" and args.phase == "search":
            continue

        pending.append(article)

    if args.max_articles:
        pending = pending[:args.max_articles]
    return pending, skipped


# ===== 主流程 =====
def run(args):
    ensure_dir(OUTPUT_DIR)
    ensure_dir(PDF_DIR)

    if args.phase == "annotate":
        return annotate_all(args)

    if args.dry_run:
        args.phase = "search"

    articles = read_corpus_excel(args.excel, 0, only_chinese=True)
    no_doi = [a for a in articles if not a["has_doi"]]
    log.info(f"待处理范围：中文无DOI文章 {len(no_doi)} 篇")

    progress = load_progress()
    progress = merge_journal_into_progress(progress)

    if args.reset:
        log.info("已清空进度（--reset）")
        progress = empty_progress()
        try:
            with open(PROGRESS_JOURNAL, "w", encoding="utf-8"):
                pass
        except OSError:
            pass
        init_report(reset=True)
    else:
        init_report(reset=False)

    if args.retry_failed:
        removed = 0
        for key in list(progress["records"]):
            if progress["records"][key].get("status") in RETRYABLE_STATUSES:
                del progress["records"][key]
                removed += 1
        log.info(f"--retry-failed：已清除 {removed} 条可重试失败记录")

    pending, skipped = select_pending_articles(no_doi, progress, args)
    save_progress(progress)

    if skipped:
        detail = "，".join(f"{k} {v}" for k, v in skipped.most_common())
        log.info(f"已跳过历史失败 {sum(skipped.values())} 篇（{detail}）；"
                 f"--retry-failed 可重试")

    log.info(f"本次待处理: {len(pending)} 篇 | 阶段: {args.phase}")
    if not pending:
        log.info("没有待处理的文章，退出")
        return 0

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        log.error("未安装 playwright，请运行: pip install playwright && playwright install chromium")
        return 2

    ensure_dir(PROFILE_DIR)
    stats = Counter()

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(args.profile or PROFILE_DIR),
            headless=args.headless,
            viewport={"width": 1400, "height": 950},
            accept_downloads=True,
            args=["--start-maximized"],
        )

        log.info(f"浏览器配置: {args.profile or PROFILE_DIR}")
        inject_cookies_into_context(context)

        downloads = []
        context.on("download", lambda d: downloads.append(d))
        search_page = context.pages[0] if context.pages else context.new_page()

        try:
            try:
                search_page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=60000)
            except Exception as e:
                log.warning(f"打开检索页失败: {e}")

            if not wait_for_search_ready(search_page, args):
                log.error("无法进入知网检索页，请检查登录/Cookie 后重试")
                return 2

            for i, article in enumerate(pending, 1):
                if search_page.is_closed():
                    search_page = context.new_page()
                    try:
                        search_page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=60000)
                    except Exception:
                        pass

                log.info(f"[{i}/{len(pending)}] no={article['no']} "
                         f"标题: {article['title'][:50]}")

                try:
                    if args.phase == "download":
                        rec, row = process_article_download(
                            context, search_page, downloads, article, args, progress)
                    else:
                        rec, row = process_article_all(
                            context, search_page, downloads, article, args, progress)
                except Exception as e:
                    log.exception(f"处理异常: {e}")
                    rec = record_progress(
                        progress, article, "failed", message=f"unhandled: {e}")
                    row = build_report_row(
                        article, "failed", message=f"unhandled: {e}")

                status = (rec or {}).get("status", "failed")
                stats[status] += 1
                if row:
                    append_report_row(row)

                if status == "captcha" and args.captcha_cooldown > 0:
                    log.info(f"  冷却 {args.captcha_cooldown} 秒后继续")
                    time.sleep(args.captcha_cooldown)

                if i % 20 == 0:
                    save_progress(progress)
                    log.info(f"进度已保存（{i}/{len(pending)}）")

                time.sleep(args.delay + random.uniform(0, 0.8))

        except KeyboardInterrupt:
            log.warning("用户中断，正在保存进度...")
        finally:
            save_progress(progress, compact=True)
            maybe_save_cookies(context, args)
            try:
                save_annotation_cache()
            except Exception:
                pass
            try:
                context.close()
            except Exception:
                pass

    log.info("=" * 60)
    log.info("本次运行完成:")
    for status, count in sorted(stats.items(), key=lambda x: (-x[1], x[0])):
        log.info(f"  {status}: {count}")
    log.info(f"进度: {PROGRESS_FILE}")
    log.info(f"报告: {REPORT_FILE}")
    log.info("=" * 60)
    return 0


def build_arg_parser():
    parser = argparse.ArgumentParser(
        description="无DOI中文文章：知网标题检索 + 元数据核对 + PDF下载",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 小批量试检索（不下载）
  python -X utf8 -u download_corpus_by_title.py --phase search --max-articles 20 --captcha-wait 0

  # 检索 + 下载 100 篇
  python -X utf8 -u download_corpus_by_title.py --phase all --max-articles 100 --captcha-wait 30

  # 大批量分批
  python -X utf8 -u download_corpus_by_title.py --phase all --max-articles 500 --captcha-wait 30

  # 重试失败项
  python -X utf8 -u download_corpus_by_title.py --phase all --retry-failed --captcha-wait 30
        """,
    )
    parser.add_argument("--phase", "-p",
                        choices=["search", "download", "all", "annotate"],
                        default="all",
                        help="search=只检索校验; download=用已保存详情页下载; "
                             "all=检索+下载（默认）; annotate=只检测PDF文本层并标注OCR")
    parser.add_argument("--max-articles", "-m", type=int, default=None,
                        help="本次最多处理多少篇")
    parser.add_argument("--start-no", type=int, default=None,
                        help="Excel 序号下限（含）")
    parser.add_argument("--end-no", type=int, default=None,
                        help="Excel 序号上限（含）")
    parser.add_argument("--search-limit", type=int, default=10,
                        help="每篇比对前 N 条检索结果（默认10）")
    parser.add_argument("--min-score", type=float, default=0.60,
                        help="标题相似度阈值（默认0.60）")
    parser.add_argument("--delay", type=float, default=2.0,
                        help="每篇之间基础延迟秒数（默认2.0）")
    parser.add_argument("--captcha-wait", type=int, default=30,
                        help="等待人工通过滑块的秒数（默认30，0=不限时）")
    parser.add_argument("--login-wait", type=int, default=300,
                        help="等待人工重新登录的秒数（默认300，0=不限时）")
    parser.add_argument("--captcha-cooldown", type=int, default=20,
                        help="出现验证码后冷却秒数（默认20）")
    parser.add_argument("--retry-failed", action="store_true",
                        help="重试之前失败的记录")
    parser.add_argument("--reset", action="store_true",
                        help="清空进度和报告（不删除PDF）")
    parser.add_argument("--dry-run", action="store_true",
                        help="只检索校验，不下载（等价 --phase search）")
    parser.add_argument("--headless", action="store_true",
                        help="无头模式（无法人工过滑块，不推荐）")
    parser.add_argument("--profile", type=str, default=None,
                        help="Playwright 配置目录（默认 data/browser_profile_title）")
    parser.add_argument("--excel", type=str, default=str(EXCEL_FILE),
                        help="语料 Excel 路径")
    parser.add_argument("--write-cookies", action="store_true",
                        help="把浏览器 Cookie 写入独立文件 data/cnki_cookies_title.json")
    parser.add_argument("--no-caj-fallback", action="store_true",
                        help="没有 PDF 时不回退到 CAJ（默认会优先PDF、缺失时用CAJ）")
    parser.add_argument("--recheck", action="store_true",
                        help="annotate 阶段重新检测已有标注的文件")
    parser.add_argument("--no-annotate", action="store_true",
                        help="下载后不检测PDF文本层（不影响下载本身）")
    return parser


def main():
    parser = build_arg_parser()
    args = parser.parse_args()

    if args.dry_run and args.phase == "download":
        parser.error("--dry-run 不能与 --phase download 同时使用")

    ensure_dir(OUTPUT_DIR)
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
    log.addHandler(fh)

    try:
        rc = run(args)
    except KeyboardInterrupt:
        log.warning("用户中断")
        rc = 130
    return rc


if __name__ == "__main__":
    sys.exit(main())
