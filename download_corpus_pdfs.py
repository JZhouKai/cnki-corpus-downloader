#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CNKI 语料 PDF 下载工具
从 20261108-语料整理表.xlsx 中提取中文文章，通过知网下载 PDF

用法：
  # 第一步：提取下载链接（从DOI解析获取知网下载URL）
  python download_corpus_pdfs.py --phase extract --max-articles 100

  # 第二步：下载PDF文件（推荐，浏览器版，遇到滑块时人工拖一次）
  python download_corpus_pdfs.py --phase download-browser --max-articles 100

  # 大批量推荐：分批提取链接并立即下载，避免 token 过期
  python download_corpus_pdfs.py --phase pipeline --chunk-size 20

  # 或者：纯 requests 下载（无法通过知网滑块验证，通常只适合无验证场景）
  python download_corpus_pdfs.py --phase download --max-articles 100

  # 完整流程：提取 + 下载（分步执行）
  python download_corpus_pdfs.py --phase all --max-articles 10

  # 只处理有DOI的文章
  python download_corpus_pdfs.py --phase extract --only-doi

  # 为非DOI文章通过CrossRef查找DOI
  python download_corpus_pdfs.py --phase crossref --max-articles 50
"""

import os
import re
import sys
import csv
import json
import time
import logging
import argparse
import hashlib
import pickle
from urllib.parse import quote, unquote, urlparse, parse_qs
from datetime import datetime
from pathlib import Path

import openpyxl
import requests

# ===== 配置 =====
EXCEL_FILE = "未下载_含DOI.xlsx"
SHEET_INDEX = 0  # Sheet 0 = 语料（中文）, Sheet 1 = 3739篇（英文）
OUTPUT_DIR = "data"
COOKIES_FILE = os.path.join(OUTPUT_DIR, "cnki_cookies.json")
PROGRESS_FILE = os.path.join(OUTPUT_DIR, "download_progress.pkl")
DOWNLOAD_URLS_FILE = os.path.join(OUTPUT_DIR, "download_urls.json")
CROSSREF_CACHE = os.path.join(OUTPUT_DIR, "crossref_cache.json")
BROWSER_PROFILE_DIR = os.path.join(OUTPUT_DIR, "browser_profile")

# 请求头
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,"
              "image/avif,image/webp,image/apng,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
}

# 下载延迟（秒）
REQUEST_DELAY = 2.0

# 永久性解析失败（后续运行直接跳过，--retry-failed 可重试）
PERMANENT_RESOLVE_STATUSES = {"doi_not_found", "no_pdf"}
UNRESOLVED_CSV = os.path.join(OUTPUT_DIR, "unresolved_articles.csv")
ARTICLE_URL_RE = re.compile(r"kns\.cnki\.net/kcms2?/(?:article|detail)/", re.I)
DOWNLOAD_MARKERS = ("pdfDown", "cajDown", "btn-dlpdf", "btn-dlcaj",
                    "bar.cnki.net/bar/download/order")

STATUS_LABELS = {
    "doi_not_found": "DOI未注册或已失效",
    "no_pdf": "文章页无PDF/CAJ下载链接",
    "login_required": "需要登录认证",
    "not_article": "未进入知网文章页",
    "resolve_failed": "DOI跳转异常",
    "error": "HTTP请求失败",
    "exception": "解析请求异常",
    "legacy_failed": "历史解析失败（旧版记录）",
    "download_failed": "历史下载失败（未保存PDF）",
}

# ===== 日志配置 =====
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)


# ===== 工具函数 =====
def ensure_dir(directory):
    """确保目录存在"""
    Path(directory).mkdir(parents=True, exist_ok=True)


def sanitize_filename(name):
    """清理文件名, 与social_find脚本一致: 70字符, 空格转下划线"""
    name = re.sub(r'[<>:"/\\|?*]', '', str(name))
    name = re.sub(r'\s+', '_', name)[:70].strip('_')
    return name if name else "unnamed"


def make_filename(seq_no, title):
    """生成文件名: {序号:06d}_{清理后标题}.pdf"""
    return f"{int(seq_no):06d}_{sanitize_filename(title)}.pdf"


def apply_cookies_to_session(session, cookie_str=None):
    """将Cookie应用到requests Session.
    优先级: cookie_str > CNKI_COOKIES环境变量 > data/cnki_cookies.json
    """
    cookie_str = cookie_str or os.environ.get("CNKI_COOKIES", "")

    if cookie_str:
        for item in cookie_str.split(";"):
            item = item.strip()
            if "=" in item:
                key, value = item.split("=", 1)
                session.cookies.set(key.strip(), value.strip())
        return True

    # 尝试从JSON文件加载（支持含中文的cookie值，自动URL编码）
    if os.path.exists(COOKIES_FILE):
        try:
            with open(COOKIES_FILE, "r", encoding="utf-8") as f:
                cookies_list = json.load(f)
            for c in cookies_list:
                name = c.get("name", "")
                value = c.get("value", "")
                domain = c.get("domain", "")
                if name and value:
                    session.cookies.set(name, value, domain=domain)
            log.info(f"从 {COOKIES_FILE} 加载了 {len(cookies_list)} 个Cookie")
            return True
        except Exception as e:
            log.warning(f"加载Cookie文件失败: {e}")
    return False


def save_progress(progress):
    """保存下载进度"""
    ensure_dir(OUTPUT_DIR)
    with open(PROGRESS_FILE, "wb") as f:
        pickle.dump(progress, f)


def load_progress():
    """加载下载进度，并兼容旧版本缺失的字段"""
    progress = {}
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "rb") as f:
            progress = pickle.load(f)
    progress.setdefault("downloaded", set())
    progress.setdefault("failed", set())
    progress.setdefault("no_pdf", set())  # 旧字段，保留兼容，不再使用
    progress.setdefault("urls", [])
    if not isinstance(progress.get("unresolved"), dict):
        progress["unresolved"] = {}
    return progress


def make_failure_record(article, status, detail, page_url):
    """构造一条无法解析的记录（写入 progress["unresolved"] 或 CSV）"""
    doi = article.get("doi") or article.get("title", "")[:40]
    return {
        "no": article.get("no"),
        "id": doi,
        "doi": article.get("doi"),
        "title": article.get("title", ""),
        "status": status,
        "reason": detail or "",
        "page_url": page_url or "",
        "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def write_unresolved_csv(unresolved):
    """把永久未解决的文章写成可读清单（utf-8-sig，Excel 可直接打开）"""
    ensure_dir(OUTPUT_DIR)
    rows = sorted(
        unresolved.values(),
        key=lambda r: (r.get("no") is None, r.get("no") or 0),
    )
    tmp = UNRESOLVED_CSV + ".tmp"
    with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["序号", "DOI", "标题", "状态", "状态说明", "原因", "页面URL", "记录时间"])
        for r in rows:
            status = r.get("status", "")
            writer.writerow([
                r.get("no", ""), r.get("doi", ""), r.get("title", ""),
                status, STATUS_LABELS.get(status, status),
                r.get("reason", ""), r.get("page_url", ""), r.get("ts", ""),
            ])
    os.replace(tmp, UNRESOLVED_CSV)


def write_download_manifest(articles, progress, pdf_dir):
    """按 Excel 顺序生成全量清单和断点文件，便于核对/恢复"""
    unresolved = progress.get("unresolved", {})
    downloaded = progress.get("downloaded", set())
    rows = []
    seen = set()
    pending_count = 0
    next_article = None

    for a in articles:
        doi = (a.get("doi") or "").strip()
        if doi:
            if doi in seen:
                continue
            seen.add(doi)
        save_path = os.path.join(pdf_dir, make_filename(a["no"], a["title"]))
        if is_valid_pdf(save_path) or (doi and doi in downloaded):
            state, note = "已下载", os.path.basename(save_path)
        elif not doi:
            state, note = "无DOI（跳过）", ""
        elif doi in unresolved:
            rec = unresolved[doi]
            state = "永久失败"
            note = STATUS_LABELS.get(rec.get("status"), rec.get("status", ""))
        else:
            state, note = "待处理", ""
            pending_count += 1
            if next_article is None:
                next_article = a
        rows.append([a.get("no"), a.get("title"), a.get("journal"),
                     a.get("year"), doi, state, note])

    manifest_path = os.path.join(OUTPUT_DIR, "download_manifest.csv")
    tmp = manifest_path + ".tmp"
    with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["序号", "标题", "期刊", "年份", "DOI", "状态", "说明/文件"])
        writer.writerows(rows)
    os.replace(tmp, manifest_path)

    checkpoint = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "excel": EXCEL_FILE,
        "total_rows": len(rows),
        "downloaded_pdfs": len(downloaded),
        "marked_permanent_failed": len(unresolved),
        "remaining_pending": pending_count,
        "next_article": ({"no": next_article.get("no"),
                          "title": next_article.get("title"),
                          "doi": next_article.get("doi")}
                         if next_article else None),
    }
    checkpoint_path = os.path.join(OUTPUT_DIR, "resume_checkpoint.json")
    with open(checkpoint_path, "w", encoding="utf-8") as f:
        json.dump(checkpoint, f, ensure_ascii=False, indent=2)


def sanitize_url_records(records):
    """丢弃旧版本写入的临时失败记录（如 login_required），避免被永久跳过"""
    kept, dropped = [], 0
    for r in records:
        status = r.get("status")
        if status and status != "urls_extracted":
            dropped += 1
            continue
        kept.append(r)
    return kept, dropped


def is_chinese(text):
    """检查文本是否包含中文字符"""
    return bool(re.search(r"[\u4e00-\u9fff]", str(text)))


# ===== Excel 读取 =====
def read_corpus_excel(filepath, sheet_index=0, only_chinese=True):
    """
    读取语料 Excel 文件
    返回: list[dict] 每篇文章的信息
    """
    if not os.path.exists(filepath):
        log.error(f"文件不存在: {filepath}")
        sys.exit(1)

    log.info(f"读取文件: {filepath}")
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)

    # 按索引选择 sheet（0=论文，1=3739篇）
    if sheet_index < len(wb.sheetnames):
        ws = wb[wb.sheetnames[sheet_index]]
    else:
        ws = wb.active

    sheet_name = ws.title
    log.info(f"使用 Sheet[{sheet_index}]: {sheet_name}")

    # 读取表头
    header = [cell.value for cell in ws[1]]
    log.info(f"列: {header}")

    # 确定列索引
    col_map = {}
    for i, col_name in enumerate(header):
        col_map[col_name] = i

    articles = []
    total_rows = ws.max_row - 1  # 减去表头
    chinese_count = 0

    for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True), 2):
        # 检查是否有中文字符
        seq_no = row[col_map.get("序号", 0)]
        title = str(row[col_map.get("标题", col_map.get("论文", col_map.get("title", 1)))]) if row[col_map.get("标题", col_map.get("论文", col_map.get("title", 1)))] else ""
        authors = str(row[col_map.get("作者", col_map.get("authors", 2))]) if row[col_map.get("作者", col_map.get("authors", 2))] else ""
        journal = str(row[col_map.get("期刊", col_map.get("journal", 3))]) if row[col_map.get("期刊", col_map.get("journal", 3))] else ""
        year = str(row[col_map.get("年份", col_map.get("year", 4))]) if row[col_map.get("年份", col_map.get("year", 4))] else ""
        doi = str(row[col_map.get("DOI", col_map.get("doi", 5))]).strip() if row[col_map.get("DOI", col_map.get("doi", 5))] else ""

        if only_chinese and not is_chinese(title):
            continue

        chinese_count += 1

        article = {
            "no": int(seq_no) if seq_no else chinese_count,
            "row": row_idx,
            "title": title,
            "authors": authors,
            "journal": journal,
            "year": year,
            "doi": doi if doi else None,
            "has_doi": bool(doi),
        }
        articles.append(article)

    wb.close()

    log.info(f"共 {total_rows} 行数据，中文文章 {chinese_count} 篇")
    doi_count = sum(1 for a in articles if a["has_doi"])
    log.info(f"其中有 DOI: {doi_count} 篇，无 DOI: {chinese_count - doi_count} 篇")

    return articles


# ===== DOI → CNKI 摘要页 → PDF下载链接 =====
def _url_host(url):
    try:
        return (urlparse(url or "").netloc or "").lower()
    except Exception:
        return ""


def classify_resolution(final_url, status_code, html):
    """根据最终 URL 和页面内容判定解析结果，返回 (status, detail)。

    注意：不能用页面里是否出现「登录」判断 —— doi.cnki.net 首页也有「登录」。
    """
    if status_code != 200:
        return "error", f"HTTP {status_code}"

    url = final_url or ""
    host = _url_host(url)
    path = (urlparse(url).path or "/").lower()
    text = html or ""
    low = text.lower()

    # 知网页面：有下载按钮即文章页；无按钮但有登录表单则是登录墙
    if host.endswith("kns.cnki.net"):
        if any(m in text for m in DOWNLOAD_MARKERS):
            return "ok", ""
        if re.search(r'id=["\']textboxusername["\']', low):
            return "login_required", "页面为知网登录表单"
        if ARTICLE_URL_RE.search(url):
            return "ok", ""
        return "not_article", f"未进入知网文章页: {url}"

    # 失效/未注册 DOI 会被 link.cnki.net 重定向到 DOI 注册中心首页
    if host.endswith("doi.cnki.net"):
        if path.rstrip("/") in ("", "/home", "/home/index"):
            return "doi_not_found", "DOI未注册或已失效（跳转到 doi.cnki.net 首页）"
        return "resolve_failed", f"跳转到 doi.cnki.net 非文章页: {url}"

    if (host.endswith("login.cnki.net") or "passport" in host
            or path.startswith("/login") or path.startswith("/user/login")
            or re.search(r'id=["\']textboxusername["\']', low)):
        return "login_required", f"跳转到登录页: {host or url}"

    return "resolve_failed", f"未识别的跳转地址: {url}"


def resolve_doi_to_cnki(session, doi):
    """通过 link.cnki.net/doi/{doi} 解析 DOI 到知网文章页。

    返回 (status, article_url, html_content, detail)。
    仅 status == "ok" 时 html_content 非空。
    """
    doi_url = f"https://link.cnki.net/doi/{doi}"
    try:
        resp = session.get(
            doi_url,
            timeout=20,
            allow_redirects=True,
        )
    except Exception as e:
        return ("exception", "", "", str(e))

    status, detail = classify_resolution(resp.url, resp.status_code, resp.text)
    html = resp.text if status == "ok" else ""
    return (status, resp.url, html, detail)


def extract_download_urls(html_content):
    """
    从知网详情页提取CAJ/PDF下载链接
    返回: dict {caj: [urls], pdf: [urls]}
    """
    result = {"caj": [], "pdf": []}

    # 提取 PDF 下载按钮链接
    pdf_patterns = [
        r'btn-dlpdf.*?href="([^"]+)"',
        r'id="pdfDown".*?href="([^"]+)"',
        r'name="pdfDown".*?href="([^"]+)"',
    ]
    for pattern in pdf_patterns:
        matches = re.findall(pattern, html_content, re.DOTALL)
        for url in matches:
            if url not in result["pdf"]:
                result["pdf"].append(url)

    # 提取 CAJ 下载按钮链接
    caj_patterns = [
        r'btn-dlcaj.*?href="([^"]+)"',
        r'id="cajDown".*?href="([^"]+)"',
        r'name="cajDown".*?href="([^"]+)"',
    ]
    for pattern in caj_patterns:
        matches = re.findall(pattern, html_content, re.DOTALL)
        for url in matches:
            if url not in result["caj"]:
                result["caj"].append(url)

    # 如果上面没找到，尝试通用模式
    if not result["pdf"] and not result["caj"]:
        bar_pattern = r'bar\.cnki\.net/bar/download/order\?id=[^"\'\\\s<>]+'
        matches = re.findall(bar_pattern, html_content)
        # 去重
        unique_urls = list(set(f"https://{m}" if not m.startswith("http") else m for m in matches))
        # 通常第一个是CAJ，第二个是PDF
        for url in unique_urls:
            result["caj"].append(url)
            result["pdf"].append(url)

    return result


def get_article_title_from_page(html_content):
    """从详情页提取文章标题"""
    match = re.search(r"<title>(.*?)</title>", html_content, re.DOTALL)
    if match:
        title = match.group(1).strip()
        # 去掉 " - 中国知网" 后缀
        title = re.sub(r"\s*-\s*中国知网\s*$", "", title)
        return title
    return None


# ===== 从 CNKI 下载 PDF =====
def download_pdf(session, download_url, save_path, max_retries=3):
    """
    从 bar.cnki.net 下载 PDF
    注意：需要校园网IP或登录Cookie
    """
    for attempt in range(max_retries):
        try:
            resp = session.get(
                download_url,
                timeout=60,
                allow_redirects=True,
                headers={
                    "Referer": "https://kns.cnki.net/",
                    "User-Agent": HEADERS["User-Agent"],
                }
            )

            if resp.status_code != 200:
                return ("error", f"HTTP {resp.status_code}")

            content_type = resp.headers.get("Content-Type", "")
            content = resp.content

            # 检查是否真正的PDF
            if content[:4] == b"%PDF":
                with open(save_path, "wb") as f:
                    f.write(content)
                return ("ok", f"PDF saved ({len(content)} bytes)")

            # 检查是否登录页面
            text = content.decode("utf-8", errors="ignore")

            # 知网下载验证页（腾讯滑块），requests 无法通过，重试无意义
            if "/bar/verify/" in resp.url or "TJCaptcha" in text or "拼图校验" in text:
                return ("captcha", "遇到滑块验证，请改用 --phase download-browser")

            if "登录" in text and ("login" in text.lower() or "账户" in text):
                return ("login_required", "需要登录或学校IP认证")

            # 其他情况 - 可能是CAJ文件
            if content[:3] == b"CAJ":
                with open(save_path.replace(".pdf", ".caj"), "wb") as f:
                    f.write(content)
                return ("caj_saved", f"CAJ saved ({len(content)} bytes)")

            # 未知响应
            return ("unknown", f"Content-Type: {content_type}, size: {len(content)}")

        except requests.exceptions.Timeout:
            if attempt < max_retries - 1:
                time.sleep(3)
                continue
            return ("timeout", "下载超时")
        except Exception as e:
            return ("exception", str(e))

    return ("failed", "重试次数耗尽")


# ===== CrossRef API 查找DOI =====
def search_doi_via_crossref(session, title, authors="", year=""):
    """
    通过 CrossRef API 根据标题查找 DOI
    """
    url = "https://api.crossref.org/works"
    params = {
        "query.title": title,
        "rows": 3,
    }
    if authors:
        params["query.author"] = authors[:50]  # 限制长度

    try:
        resp = session.get(url, params=params, timeout=15)
        if resp.status_code != 200:
            return None

        data = resp.json()
        items = data.get("message", {}).get("items", [])
        if not items:
            return None

        # 找最佳匹配
        for item in items:
            item_title = item.get("title", [""])[0] if item.get("title") else ""
            item_doi = item.get("DOI", "")

            if item_doi and item_title:
                # 计算标题相似度（简单判断）
                if title[:20].lower() in item_title.lower() or item_title[:20].lower() in title.lower():
                    return item_doi

        # 没有精确匹配，返回第一个
        first_doi = items[0].get("DOI", "")
        return first_doi if first_doi else None

    except Exception as e:
        log.debug(f"CrossRef 查询失败: {e}")
        return None


# ===== 主处理逻辑 =====
def phase_extract(args):
    """
    第一阶段：提取下载链接
    从Excel读取文章 → 解析DOI → 获取知网详情页 → 提取PDF下载链接
    """
    log.info("=" * 60)
    log.info("第一阶段：提取下载链接")
    log.info("=" * 60)

    ensure_dir(OUTPUT_DIR)
    articles = read_corpus_excel(EXCEL_FILE, SHEET_INDEX, only_chinese=True)
    all_articles = articles
    log.info(f"共 {len(articles)} 篇中文文章")

    # 过滤
    if args.only_doi:
        articles = [a for a in articles if a["has_doi"]]
        log.info(f"仅处理有DOI的文章: {len(articles)} 篇")
    elif args.only_no_doi:
        articles = [a for a in articles if not a["has_doi"]]
        log.info(f"仅处理无DOI的文章: {len(articles)} 篇")

    # 加载进度
    progress = load_progress()
    if getattr(args, "retry_failed", False):
        n_unresolved = len(progress["unresolved"])
        progress["unresolved"].clear()
        progress["failed"].clear()
        progress["no_pdf"].clear()
        save_progress(progress)
        log.info(f"--retry-failed：已清空 {n_unresolved} 条永久未解决记录，本次重新解析")

    downloaded_urls = progress.get("urls", [])
    if not downloaded_urls and os.path.exists(DOWNLOAD_URLS_FILE):
        with open(DOWNLOAD_URLS_FILE, "r", encoding="utf-8") as f:
            downloaded_urls = json.load(f)

    # 清理旧版本写入的临时失败记录（如 login_required），避免被永久跳过
    downloaded_urls, dropped = sanitize_url_records(downloaded_urls)
    if dropped:
        log.info(f"已清理 {dropped} 条旧的临时失败记录，将重新解析")
        with open(DOWNLOAD_URLS_FILE, "w", encoding="utf-8") as f:
            json.dump(downloaded_urls, f, ensure_ascii=False, indent=2)
        progress["urls"] = downloaded_urls
        save_progress(progress)

    # 先剔除已处理过的文章、永久未解决的，再按 max_articles 截断，
    # 这样多次运行 --max-articles N 才能持续推进到后面的文章
    processed_ids = {u.get("id") for u in downloaded_urls if u.get("id")}
    unresolved = progress["unresolved"]
    before = len(articles)
    articles = [a for a in articles
                if (a.get("doi") or a["title"][:40]) not in processed_ids
                and a.get("doi") not in unresolved]
    skipped_unresolved = before - len(articles)
    log.info(f"待处理新文章: {len(articles)} 篇（已跳过已处理/永久未解决 {skipped_unresolved} 篇）")

    if args.max_articles and args.max_articles < len(articles):
        articles = articles[:args.max_articles]
        log.info(f"限制处理前 {args.max_articles} 篇")

    session = requests.Session()
    session.headers.update(HEADERS)
    apply_cookies_to_session(session)

    stats = {"ok": 0, "doi_not_found": 0, "no_pdf": 0, "no_url": 0,
             "login_required": 0, "transient": 0, "error": 0, "skipped": 0}

    for i, article in enumerate(articles):
        # 跳过已处理的
        article_id = article.get("doi") or article["title"][:40]
        if any(u.get("id") == article_id for u in downloaded_urls):
            stats["skipped"] += 1
            continue

        if i % 10 == 0 and i > 0:
            write_unresolved_csv(unresolved)
            log.info(f"进度: {i}/{len(articles)} | OK={stats['ok']} "
                     f"DOI失效={stats['doi_not_found']} NoPDF={stats['no_pdf']} "
                     f"临时失败={stats['transient']} Skipped={stats['skipped']}")

        title_short = article["title"][:50]

        if article["has_doi"]:
            # 有DOI：直接解析
            log.info(f"[{i+1}/{len(articles)}] DOI: {article['doi']}")
            log.info(f"  标题: {title_short}")

            status, page_url, html, detail = resolve_doi_to_cnki(session, article["doi"])

            if status == "ok" and html:
                urls = extract_download_urls(html)
                page_title = get_article_title_from_page(html)

                record = {
                    "no": article["no"],
                    "id": article["doi"],
                    "title": article["title"],
                    "doi": article["doi"],
                    "page_title": page_title,
                    "page_url": page_url,
                    "pdf_urls": urls.get("pdf", []),
                    "caj_urls": urls.get("caj", []),
                    "has_pdf": len(urls.get("pdf", [])) > 0,
                    "has_caj": len(urls.get("caj", [])) > 0,
                    "status": "urls_extracted",
                }

                if urls.get("pdf") or urls.get("caj"):
                    downloaded_urls.append(record)
                    stats["ok"] += 1
                    log.info(f"  ✓ 找到PDF下载链接")
                else:
                    fail = make_failure_record(
                        article, "no_pdf", "文章页无 PDF/CAJ 下载链接", page_url)
                    unresolved[fail["id"]] = fail
                    stats["no_pdf"] += 1
                    stats["no_url"] += 1
                    log.warning("  ✗ 文章页无 PDF/CAJ 下载链接，已永久跳过")

            elif status == "ok":
                stats["error"] += 1
                log.warning("  ! 文章页内容为空，本次跳过（下次重试）")

            elif status in PERMANENT_RESOLVE_STATUSES:
                fail = make_failure_record(article, status, detail, page_url)
                unresolved[fail["id"]] = fail
                stats[status] = stats.get(status, 0) + 1
                log.warning(f"  ✗ {STATUS_LABELS.get(status, status)}，已永久跳过")

            else:
                stats["transient"] += 1
                if status == "login_required":
                    stats["login_required"] += 1
                log.warning(f"  ! {STATUS_LABELS.get(status, status)}"
                            f"（{detail}），本次跳过（下次重试）")

        else:
            # 无DOI：记录日志，后续处理
            stats["no_url"] += 1
            log.info(f"[{i+1}/{len(articles)}] 无DOI: {title_short}")

        # 保存进度
        progress["urls"] = downloaded_urls
        save_progress(progress)

        # 延迟
        time.sleep(REQUEST_DELAY)

    # 保存最终结果
    with open(DOWNLOAD_URLS_FILE, "w", encoding="utf-8") as f:
        json.dump(downloaded_urls, f, ensure_ascii=False, indent=2)
    progress["urls"] = downloaded_urls
    save_progress(progress)
    write_unresolved_csv(unresolved)
    try:
        write_download_manifest(
            all_articles, progress, os.path.join(OUTPUT_DIR, "PDFs"))
    except Exception as e:
        log.warning(f"生成下载清单失败: {e}")

    log.info("=" * 60)
    log.info(f"提取完成:")
    log.info(f"  成功提取URL: {stats['ok']}")
    log.info(f"  永久跳过(DOI失效): {stats['doi_not_found']}")
    log.info(f"  永久跳过(无下载链接): {stats['no_pdf']}")
    log.info(f"  临时失败(下次重试): {stats['transient']}")
    log.info(f"  需要登录: {stats['login_required']}")
    log.info(f"  解析错误: {stats['error']}")
    log.info(f"  已跳过: {stats['skipped']}")
    log.info(f"  总记录: {len(downloaded_urls)}")
    log.info(f"  累计永久未解决: {len(unresolved)}")
    log.info(f"  URL文件: {DOWNLOAD_URLS_FILE}")
    log.info(f"  未解决清单: {UNRESOLVED_CSV}")
    log.info("=" * 60)


def phase_crossref(args):
    """
    通过 CrossRef 为无DOI的文章查找DOI
    """
    log.info("=" * 60)
    log.info("CrossRef DOI 查找阶段")
    log.info("=" * 60)

    ensure_dir(OUTPUT_DIR)
    articles = read_corpus_excel(EXCEL_FILE, SHEET_INDEX, only_chinese=True)

    # 只处理无DOI的
    no_doi_articles = [a for a in articles if not a["has_doi"]]
    log.info(f"无DOI文章: {len(no_doi_articles)} 篇")

    if args.max_articles and args.max_articles < len(no_doi_articles):
        no_doi_articles = no_doi_articles[:args.max_articles]
        log.info(f"限制处理前 {args.max_articles} 篇")

    # 加载缓存
    crossref_cache = {}
    if os.path.exists(CROSSREF_CACHE):
        with open(CROSSREF_CACHE, "r", encoding="utf-8") as f:
            crossref_cache = json.load(f)

    session = requests.Session()
    session.headers.update({
        "User-Agent": "CNKI-Download-Tool/1.0 (mailto:example@example.com)",
    })

    found = 0
    not_found = 0

    for i, article in enumerate(no_doi_articles):
        title = article["title"]

        # 检查缓存
        if title in crossref_cache:
            if crossref_cache[title]:
                found += 1
            else:
                not_found += 1
            continue

        if i % 5 == 0 and i > 0:
            log.info(f"进度: {i}/{len(no_doi_articles)} | 找到: {found}, 未找到: {not_found}")

        log.info(f"[{i+1}/{len(no_doi_articles)}] 搜索: {title[:50]}")

        doi = search_doi_via_crossref(session, title, article["authors"])
        crossref_cache[title] = doi

        if doi:
            found += 1
            log.info(f"  ✓ 找到 DOI: {doi}")
        else:
            not_found += 1
            log.debug(f"  ✗ 未找到")

        # 保存缓存
        with open(CROSSREF_CACHE, "w", encoding="utf-8") as f:
            json.dump(crossref_cache, f, ensure_ascii=False, indent=2)

        time.sleep(0.5)  # CrossRef API 速率限制

    log.info("=" * 60)
    log.info(f"CrossRef 查找完成:")
    log.info(f"  找到DOI: {found}")
    log.info(f"  未找到: {not_found}")
    log.info("=" * 60)

    # 输出找到的DOI文章
    output_file = os.path.join(OUTPUT_DIR, "crossref_found_dois.json")
    found_articles = []
    for a in no_doi_articles:
        doi = crossref_cache.get(a["title"])
        if doi:
            found_articles.append({
                "title": a["title"],
                "authors": a["authors"],
                "journal": a["journal"],
                "doi": doi,
            })

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(found_articles, f, ensure_ascii=False, indent=2)
    log.info(f"找到DOI的文章已保存到: {output_file}")


def phase_download(args):
    """
    第二阶段：下载PDF
    使用第一阶段提取的URL下载PDF文件（需要校园网IP）
    """
    log.info("=" * 60)
    log.info("第二阶段：下载PDF")
    log.info("=" * 60)

    ensure_dir(OUTPUT_DIR)

    # 加载URL列表
    if not os.path.exists(DOWNLOAD_URLS_FILE):
        log.error(f"URL文件不存在: {DOWNLOAD_URLS_FILE}")
        log.error("请先运行 --phase extract 提取下载链接")
        return

    with open(DOWNLOAD_URLS_FILE, "r", encoding="utf-8") as f:
        url_records = json.load(f)

    log.info(f"共 {len(url_records)} 条下载记录")

    # 过滤出有下载URL的记录
    to_download = [r for r in url_records if r.get("pdf_urls") or r.get("caj_urls")]
    log.info(f"有下载链接: {len(to_download)} 篇")

    if args.max_articles and args.max_articles < len(to_download):
        to_download = to_download[:args.max_articles]
        log.info(f"限制下载前 {args.max_articles} 篇")

    # 加载下载进度
    progress = load_progress()
    if getattr(args, "retry_failed", False):
        progress["failed"].clear()
        log.info("已清空失败记录，本次将重试此前失败的文章")

    session = requests.Session()
    session.headers.update(HEADERS)
    apply_cookies_to_session(session)

    pdf_dir = os.path.join(OUTPUT_DIR, "PDFs")
    ensure_dir(pdf_dir)

    stats = {"ok": 0, "login": 0, "failed": 0, "skipped": 0, "caj": 0}

    for i, record in enumerate(to_download):
        title = record.get("title", "未知")
        article_id = record.get("id", title[:40])

        # 跳过已下载的
        if article_id in progress.get("downloaded", set()):
            stats["skipped"] += 1
            continue

        # 跳过已失败的
        if article_id in progress.get("failed", set()):
            stats["failed"] += 1
            continue

        # 选择下载URL（优先PDF，其次CAJ）
        dl_urls = record.get("pdf_urls", []) or record.get("caj_urls", [])
        if not dl_urls:
            continue

        dl_url = dl_urls[0]
        fname = make_filename(record.get("no", i+1), title)
        save_path = os.path.join(pdf_dir, fname)

        log.info(f"[{i+1}/{len(to_download)}] 下载: {title[:50]}")

        result_status, result_msg = download_pdf(session, dl_url, save_path)

        if result_status == "ok":
            stats["ok"] += 1
            progress["downloaded"].add(article_id)
            log.info(f"  ✓ {result_msg}")
        elif result_status == "caj_saved":
            stats["caj"] += 1
            progress["downloaded"].add(article_id)
            log.info(f"  △ {result_msg}")
        elif result_status == "login_required":
            stats["login"] += 1
            progress["failed"].add(article_id)
            log.warning(f"  ! {result_msg}")
        else:
            stats["failed"] += 1
            progress["failed"].add(article_id)
            log.warning(f"  ✗ {result_msg}")

        # 保存进度
        save_progress(progress)

        time.sleep(REQUEST_DELAY)

    log.info("=" * 60)
    log.info(f"下载完成:")
    log.info(f"  PDF下载成功: {stats['ok']}")
    log.info(f"  CAJ下载成功: {stats['caj']}")
    log.info(f"  需要登录: {stats['login']}")
    log.info(f"  下载失败: {stats['failed']}")
    log.info(f"  已跳过: {stats['skipped']}")
    log.info("=" * 60)


# ===== Playwright 半自动下载（可人工通过滑块） =====
def build_browser_cookies():
    """构造 Playwright 格式的 Cookie 列表，来源同 apply_cookies_to_session"""
    cookies = []
    cookie_str = os.environ.get("CNKI_COOKIES", "").strip()
    if cookie_str:
        for item in cookie_str.split(";"):
            item = item.strip()
            if "=" in item:
                name, value = item.split("=", 1)
                cookies.append({"name": name.strip(), "value": value.strip(),
                                "domain": ".cnki.net", "path": "/"})
        return cookies

    if os.path.exists(COOKIES_FILE):
        try:
            with open(COOKIES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for c in data:
                if c.get("name") and c.get("value") is not None:
                    cookies.append({
                        "name": c["name"],
                        "value": str(c["value"]),
                        "domain": c.get("domain") or ".cnki.net",
                        "path": c.get("path") or "/",
                    })
        except Exception as e:
            log.warning(f"加载Cookie文件失败: {e}")
    return cookies


def is_valid_pdf(path):
    """检查文件是否存在且是真正的PDF"""
    if not os.path.exists(path) or os.path.getsize(path) < 10240:
        return False
    with open(path, "rb") as f:
        return f.read(4) == b"%PDF"


def page_needs_login(page):
    """粗略判断当前页面是否是登录页"""
    try:
        url = (page.url or "").lower()
        if "login" in url or "passport" in url:
            return True
        title = page.title() or ""
        if "登录" in title and "拼图" not in title:
            return True
        if page.locator("#TextBoxUserName").count() > 0:
            return True
    except Exception:
        pass
    return False


def wait_for_pdf_button(page, timeout):
    """登录失效后等待用户在浏览器里完成登录，直到详情页重新出现PDF按钮"""
    deadline = None if timeout <= 0 else time.time() + timeout
    while deadline is None or time.time() < deadline:
        try:
            if page.locator("#pdfDown").count() > 0:
                return True
        except Exception:
            return False
        try:
            page.wait_for_timeout(500)
        except Exception:
            time.sleep(0.5)
    return False


def save_browser_cookies(context):
    """把浏览器中的知网Cookie回写到 data/cnki_cookies.json"""
    try:
        cookies = [c for c in context.cookies()
                   if "cnki.net" in (c.get("domain") or "")]
        if not cookies:
            return False
        data = [{
            "name": c["name"],
            "value": c["value"],
            "domain": c.get("domain") or ".cnki.net",
            "path": c.get("path") or "/",
            "expires": None if not c.get("expires") or c["expires"] < 0 else c["expires"],
        } for c in cookies]
        tmp = COOKIES_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, COOKIES_FILE)
        log.info(f"已更新Cookie文件: {COOKIES_FILE}（{len(data)} 个）")
        return True
    except Exception as e:
        log.warning(f"保存Cookie失败: {e}")
        return False


def inject_cookies_into_context(context):
    """每次启动都把 cnki_cookies.json 注入浏览器，便于更新登录态"""
    cookies = build_browser_cookies()
    if cookies:
        context.add_cookies(cookies)
        log.info(f"已注入 {len(cookies)} 个Cookie")
    return len(cookies)


def wait_for_download(context, downloads, start_index, timeout, login_timeout=None):
    """轮询等待 context 产生新的下载事件（期间保持浏览器事件循环）
    timeout <= 0 时表示一直等待。
    若中途检测到登录页，则把等待时间延长为 login_timeout（默认 300 秒）。
    """
    deadline = None if timeout <= 0 else time.time() + timeout
    noticed_verify = False
    noticed_login = False
    while deadline is None or time.time() < deadline:
        if len(downloads) > start_index:
            return downloads[start_index]

        pages = context.pages
        if not noticed_verify and any("/bar/verify/" in (p.url or "") for p in pages):
            wait_desc = "不限时" if timeout <= 0 else f"最长等待 {timeout} 秒"
            log.info(f"  → 请在浏览器窗口中拖动滑块完成验证，通过后会自动下载"
                     f"（{wait_desc}）")
            noticed_verify = True

        if not noticed_login and any(page_needs_login(p) for p in pages):
            login_timeout = 300 if login_timeout is None else login_timeout
            deadline = None if login_timeout <= 0 else time.time() + login_timeout
            wait_desc = "不限时" if login_timeout <= 0 else f"最长等待 {login_timeout} 秒"
            log.info(f"  → 登录已失效，请在浏览器中登录，登录后会自动继续（{wait_desc}）")
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
    """保存 Playwright 下载，并检查文件头"""
    tmp_path = save_path + ".part"
    try:
        download.save_as(tmp_path)
    except Exception as e:
        return ("error", f"保存失败: {e}")

    with open(tmp_path, "rb") as f:
        head = f.read(4)
    size = os.path.getsize(tmp_path)

    if head == b"%PDF":
        os.replace(tmp_path, save_path)
        return ("ok", f"PDF saved ({size} bytes)")

    if head[:3] == b"CAJ":
        caj_path = save_path[:-4] + ".caj" if save_path.lower().endswith(".pdf") else save_path + ".caj"
        os.replace(tmp_path, caj_path)
        return ("caj", f"CAJ saved ({size} bytes)")

    try:
        os.remove(tmp_path)
    except OSError:
        pass
    return ("invalid", f"下载内容不是PDF/CAJ（{size} bytes）")


def close_extra_pages(context, page):
    """关闭验证页/弹窗，保留主页面"""
    for extra in list(context.pages):
        if extra is not page:
            try:
                extra.close()
            except Exception:
                pass


def download_records_browser(context, page, downloads, records, args, progress, stats):
    """在已打开的浏览器上下文中逐条下载记录，更新 progress/stats"""
    pdf_dir = os.path.join(OUTPUT_DIR, "PDFs")
    ensure_dir(pdf_dir)
    timeout_desc = "不限时" if args.captcha_wait <= 0 else f"{args.captcha_wait} 秒"

    for i, record in enumerate(records):
        title = record.get("title", "未知")
        article_id = record.get("id", title[:40])
        dl_urls = record.get("pdf_urls", []) or record.get("caj_urls", [])
        if not dl_urls:
            continue

        save_path = os.path.join(pdf_dir, make_filename(record.get("no", i + 1), title))

        # 文件已存在且有效 → 直接记为已下载（修复进度与文件不一致的情况）
        if is_valid_pdf(save_path):
            stats["skipped"] += 1
            progress["downloaded"].add(article_id)
            progress["failed"].discard(article_id)
            save_progress(progress)
            log.info(f"[{i+1}/{len(records)}] 已存在，跳过: {title[:50]}")
            continue

        if article_id in progress.get("downloaded", set()):
            stats["skipped"] += 1
            log.info(f"[{i+1}/{len(records)}] 进度已记录，跳过: {title[:50]}")
            continue

        log.info(f"[{i+1}/{len(records)}] 下载: {title[:50]}")

        # 优先从详情页点击 PDF 按钮；页面失效时回退到 DOI 解析
        page_url = record.get("page_url")
        try:
            page.goto(page_url, wait_until="domcontentloaded", timeout=60000)

            # 登录失效：暂停等待用户在浏览器里完成登录
            if page.locator("#pdfDown").count() == 0 and page_needs_login(page):
                login_desc = "不限时" if args.login_wait <= 0 else f"最长等待 {args.login_wait} 秒"
                log.warning(f"  ! 登录已失效，请在浏览器中登录（{login_desc}）")
                if wait_for_pdf_button(page, args.login_wait):
                    log.info("  ✓ 已重新登录，继续下载")
                    save_browser_cookies(context)
                else:
                    log.warning("  ! 等待登录超时")

            if page.locator("#pdfDown").count() == 0 and record.get("doi"):
                page.goto(f"https://link.cnki.net/doi/{record['doi']}",
                          wait_until="domcontentloaded", timeout=60000)
            if page.locator("#pdfDown").count() == 0:
                log.warning("  ! 页面上未找到PDF下载按钮")
        except Exception as e:
            log.warning(f"  ! 打开详情页失败: {e}")

        before = len(downloads)
        clicked = False
        try:
            if page.locator("#pdfDown").count() > 0:
                page.locator("#pdfDown").first.click(timeout=15000)
                clicked = True
            else:
                page.goto(dl_urls[0], wait_until="domcontentloaded", timeout=60000)
                clicked = True
        except Exception as e:
            log.warning(f"  ! 点击下载按钮失败: {e}")

        download = (wait_for_download(context, downloads, before,
                                      args.captcha_wait, args.login_wait)
                    if clicked else None)

        if download is None:
            stats["failed"] += 1
            progress["failed"].add(article_id)
            save_progress(progress)
            log.warning(f"  ✗ 等待{timeout_desc}未捕获下载，记为失败（下次运行会自动重试）")
            close_extra_pages(context, page)
            continue

        result_status, result_msg = save_browser_download(download, save_path)
        if result_status == "ok":
            stats["ok"] += 1
            progress["downloaded"].add(article_id)
            progress["failed"].discard(article_id)
            log.info(f"  ✓ {result_msg}")
        elif result_status == "caj":
            stats["caj"] += 1
            progress["downloaded"].add(article_id)
            progress["failed"].discard(article_id)
            log.info(f"  △ {result_msg}")
        else:
            stats["failed"] += 1
            progress["failed"].add(article_id)
            log.warning(f"  ✗ {result_msg}")
        save_progress(progress)

        close_extra_pages(context, page)
        time.sleep(REQUEST_DELAY)

    return stats


def phase_download_browser(args):
    """
    第二阶段（浏览器版）：用 Playwright 打开知网页面并点击 PDF 下载。
    遇到腾讯滑块时在可见浏览器中人工拖动一次，下载会自动继续；
    登录态与验证态保存在 data/browser_profile 中，下次可直接复用。
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        log.error("未安装 playwright，请先运行: pip install playwright && playwright install chromium")
        return

    log.info("=" * 60)
    log.info("第二阶段：浏览器半自动下载PDF")
    log.info("=" * 60)

    if not os.path.exists(DOWNLOAD_URLS_FILE):
        log.error(f"URL文件不存在，请先运行 --phase extract")
        return

    with open(DOWNLOAD_URLS_FILE, "r", encoding="utf-8") as f:
        records = json.load(f)

    log.info(f"共 {len(records)} 条下载记录")
    to_download = [r for r in records if r.get("pdf_urls") or r.get("caj_urls")]
    log.info(f"有下载链接: {len(to_download)} 篇")
    if args.max_articles and args.max_articles < len(to_download):
        to_download = to_download[:args.max_articles]
        log.info(f"限制下载前 {args.max_articles} 篇")

    progress = load_progress()
    for key in ("downloaded", "failed", "no_pdf"):
        progress.setdefault(key, set())
    # 浏览器阶段允许重试历史失败项
    for r in to_download:
        progress["failed"].discard(r.get("id"))

    profile_dir = args.browser_profile or BROWSER_PROFILE_DIR
    profile_is_new = not os.path.isdir(profile_dir) or not os.listdir(profile_dir)
    ensure_dir(profile_dir)

    stats = {"ok": 0, "caj": 0, "failed": 0, "skipped": 0}

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            headless=False,
            viewport={"width": 1280, "height": 900},
            accept_downloads=True,
            args=["--start-maximized"],
        )

        log.info(("新建" if profile_is_new else "复用") + f"浏览器配置: {profile_dir}")
        inject_cookies_into_context(context)

        downloads = []
        context.on("download", lambda d: downloads.append(d))
        page = context.pages[0] if context.pages else context.new_page()

        try:
            download_records_browser(context, page, downloads, to_download, args, progress, stats)
        finally:
            save_browser_cookies(context)
            context.close()

    log.info("=" * 60)
    log.info("下载完成:")
    log.info(f"  PDF下载成功: {stats['ok']}")
    log.info(f"  CAJ下载成功: {stats['caj']}")
    log.info(f"  下载失败: {stats['failed']}")
    log.info(f"  已跳过: {stats['skipped']}")
    log.info(f"  浏览器配置: {profile_dir}")
    log.info("=" * 60)


def extract_records_for_articles(session, articles, offset=0, total=None):
    """
    为一批文章解析知网下载链接。
    返回 (records, failures)：records 为可下载记录，failures 为解析失败记录。
    用于 pipeline 阶段：解析完一小批后立即下载，避免 bar.cnki.net token 过期。
    """
    records = []
    failures = []
    total = total if total is not None else len(articles)

    for i, article in enumerate(articles):
        idx = offset + i + 1
        title = article["title"]

        if not article.get("has_doi"):
            log.info(f"[{idx}/{total}] 无DOI，跳过: {title[:50]}")
            continue

        log.info(f"[{idx}/{total}] DOI: {article['doi']}")
        log.info(f"  标题: {title[:50]}")

        status, page_url, html, detail = resolve_doi_to_cnki(session, article["doi"])
        if status == "ok" and not html:
            status, detail = "error", "响应内容为空"

        if status != "ok":
            failures.append(make_failure_record(article, status, detail, page_url))
            if status in PERMANENT_RESOLVE_STATUSES:
                log.warning(f"  ✗ {STATUS_LABELS.get(status, status)}，已永久跳过")
            else:
                log.warning(f"  ! {STATUS_LABELS.get(status, status)}"
                            f"（{detail}），本次跳过（下次重试）")
            time.sleep(REQUEST_DELAY)
            continue

        urls = extract_download_urls(html)
        if not (urls.get("pdf") or urls.get("caj")):
            failures.append(make_failure_record(
                article, "no_pdf", "文章页无 PDF/CAJ 下载链接", page_url))
            log.warning("  ✗ 文章页无 PDF/CAJ 下载链接，已永久跳过")
            time.sleep(REQUEST_DELAY)
            continue

        records.append({
            "no": article["no"],
            "id": article["doi"],
            "title": title,
            "doi": article["doi"],
            "page_title": get_article_title_from_page(html),
            "page_url": page_url,
            "pdf_urls": urls.get("pdf", []),
            "caj_urls": urls.get("caj", []),
            "has_pdf": len(urls.get("pdf", [])) > 0,
            "has_caj": len(urls.get("caj", [])) > 0,
            "status": "urls_extracted",
        })
        log.info("  ✓ 找到下载链接")
        time.sleep(REQUEST_DELAY)

    return records, failures


def phase_pipeline(args):
    """
    边提取边下载：每批解析 chunk-size 篇文章的链接后立即下载。
    - DOI 失效、文章页无下载链接 → 永久标记，下次运行直接跳过；
    - 登录失效/网络异常等临时失败 → 下次运行重试；
    - 已有有效PDF的文章不会重复解析；
    - 避免一次性提取全部URL导致 token 过期。
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        log.error("未安装 playwright，请先运行: pip install playwright && playwright install chromium")
        return

    log.info("=" * 60)
    log.info("增量流水线：提取 + 立即下载")
    log.info("=" * 60)

    chunk_size = max(1, getattr(args, "chunk_size", 20))
    articles = read_corpus_excel(EXCEL_FILE, SHEET_INDEX, only_chinese=True)
    all_articles = articles

    if args.only_no_doi:
        log.warning("无DOI文章无法通过知网DOI解析，pipeline 阶段仅处理有DOI文章")
        return
    if args.only_doi:
        articles = [a for a in articles if a["has_doi"]]

    progress = load_progress()
    if getattr(args, "retry_failed", False):
        n_unresolved = len(progress["unresolved"])
        n_failed = len(progress["failed"])
        progress["unresolved"].clear()
        progress["failed"].clear()
        progress["no_pdf"].clear()
        save_progress(progress)
        log.info(f"--retry-failed：已清空 {n_unresolved} 条永久未解决、"
                 f"{n_failed} 条下载失败记录，本次重新解析")

    # 已有有效PDF的不再解析，避免重复劳动
    pdf_dir = os.path.join(OUTPUT_DIR, "PDFs")
    ensure_dir(pdf_dir)

    # 加载并清理旧的 URL 记录（旧版本可能把 login_required 也写进去）
    existing = []
    if os.path.exists(DOWNLOAD_URLS_FILE):
        with open(DOWNLOAD_URLS_FILE, "r", encoding="utf-8") as f:
            existing = json.load(f)
    existing, dropped = sanitize_url_records(existing)
    if dropped:
        with open(DOWNLOAD_URLS_FILE, "w", encoding="utf-8") as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)
        log.info(f"已清理 {dropped} 条旧的临时失败记录，将重新解析")
    url_index = {r.get("id"): i for i, r in enumerate(existing)}

    unresolved = progress["unresolved"]
    pending = []
    seen = set()
    for a in articles:
        doi = a.get("doi")
        if not doi or doi in seen:
            continue
        seen.add(doi)
        if doi in progress["downloaded"]:
            continue
        save_path = os.path.join(pdf_dir, make_filename(a["no"], a["title"]))
        if is_valid_pdf(save_path):
            progress["downloaded"].add(doi)
            unresolved.pop(doi, None)
            continue
        if doi in unresolved:
            continue
        pending.append(a)

    save_progress(progress)

    if args.max_articles:
        pending = pending[:args.max_articles]
    if not pending:
        log.info(f"没有待处理的文章，退出（累计永久未解决 {len(unresolved)} 篇）")
        return

    log.info(f"待处理: {len(pending)} 篇，每批 {chunk_size} 篇"
             f"（累计永久未解决 {len(unresolved)} 篇）")

    session = requests.Session()
    session.headers.update(HEADERS)
    apply_cookies_to_session(session)

    profile_dir = args.browser_profile or BROWSER_PROFILE_DIR
    profile_is_new = not os.path.isdir(profile_dir) or not os.listdir(profile_dir)
    ensure_dir(profile_dir)

    stats = {"ok": 0, "caj": 0, "failed": 0, "skipped": 0}
    resolve_stats = {"ok": 0, "doi_not_found": 0, "no_pdf": 0, "transient": 0}
    batches = (len(pending) + chunk_size - 1) // chunk_size

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=profile_dir,
            headless=False,
            viewport={"width": 1280, "height": 900},
            accept_downloads=True,
            args=["--start-maximized"],
        )

        log.info(("新建" if profile_is_new else "复用") + f"浏览器配置: {profile_dir}")
        inject_cookies_into_context(context)

        downloads = []
        context.on("download", lambda d: downloads.append(d))
        page = context.pages[0] if context.pages else context.new_page()

        try:
            for batch_no, start in enumerate(range(0, len(pending), chunk_size), 1):
                chunk = pending[start:start + chunk_size]
                log.info("-" * 60)
                log.info(f"批次 {batch_no}/{batches}: 解析 {len(chunk)} 篇文章链接")
                log.info("-" * 60)

                records, failures = extract_records_for_articles(
                    session, chunk, offset=start, total=len(pending)
                )

                # 永久失败落盘；临时失败只记日志，下次重试
                permanent = [f for f in failures
                             if f["status"] in PERMANENT_RESOLVE_STATUSES]
                transient = len(failures) - len(permanent)
                for f in permanent:
                    unresolved[f["id"]] = f
                resolve_stats["ok"] += len(records)
                resolve_stats["transient"] += transient
                for f in permanent:
                    resolve_stats[f["status"]] = resolve_stats.get(f["status"], 0) + 1

                # 合并/更新 URL 记录并落盘，便于中断后复查
                for rec in records:
                    idx = url_index.get(rec["id"])
                    if idx is None:
                        url_index[rec["id"]] = len(existing)
                        existing.append(rec)
                    else:
                        existing[idx] = rec
                with open(DOWNLOAD_URLS_FILE, "w", encoding="utf-8") as f:
                    json.dump(existing, f, ensure_ascii=False, indent=2)
                progress["urls"] = existing
                if permanent:
                    write_unresolved_csv(unresolved)
                save_progress(progress)

                log.info(
                    f"批次 {batch_no}/{batches}: 解析成功 {len(records)} 篇；"
                    f"永久跳过 {len(permanent)} 篇"
                    f"（DOI失效 {sum(1 for f in permanent if f['status'] == 'doi_not_found')}，"
                    f"无下载链接 {sum(1 for f in permanent if f['status'] == 'no_pdf')}）；"
                    f"临时失败 {transient} 篇（下次重试）")
                download_records_browser(context, page, downloads, records, args, progress, stats)

                save_browser_cookies(context)
                log.info(f"批次 {batch_no}/{batches} 完成 | 累计成功 {stats['ok']} | "
                         f"失败 {stats['failed']} | 跳过 {stats['skipped']}")
        finally:
            # Ctrl+C 时也保存状态并关闭浏览器，避免进度/清单丢失
            try:
                write_unresolved_csv(unresolved)
            except Exception as e:
                log.warning(f"写入未解决清单失败: {e}")
            try:
                save_progress(progress)
            except Exception as e:
                log.warning(f"保存进度失败: {e}")
            try:
                write_download_manifest(all_articles, progress, pdf_dir)
            except Exception as e:
                log.warning(f"生成下载清单失败: {e}")
            try:
                save_browser_cookies(context)
            except Exception as e:
                log.warning(f"保存Cookie失败: {e}")
            try:
                context.close()
            except Exception:
                pass

    log.info("=" * 60)
    log.info("流水线完成:")
    log.info(f"  PDF下载成功: {stats['ok']}")
    log.info(f"  CAJ下载成功: {stats['caj']}")
    log.info(f"  下载失败: {stats['failed']}")
    log.info(f"  已跳过: {stats['skipped']}")
    log.info(f"  本次解析成功: {resolve_stats['ok']}")
    log.info(f"  本次永久跳过: {resolve_stats['doi_not_found'] + resolve_stats['no_pdf']} 篇"
             f"（DOI失效 {resolve_stats['doi_not_found']}，"
             f"无下载链接 {resolve_stats['no_pdf']}）")
    log.info(f"  本次临时失败（下次重试）: {resolve_stats['transient']} 篇")
    log.info(f"  累计永久未解决: {len(unresolved)} 篇")
    log.info(f"  未解决清单: {UNRESOLVED_CSV}")
    log.info(f"  浏览器配置: {profile_dir}")
    log.info("=" * 60)


def phase_export_aria2(args):
    """
    导出aria2c批量下载脚本
    aria2c是一个支持多线程的下载工具，可以大幅提升下载速度
    """
    log.info("=" * 60)
    log.info("导出aria2c批量下载脚本")
    log.info("=" * 60)

    if not os.path.exists(DOWNLOAD_URLS_FILE):
        log.error(f"URL文件不存在，请先运行 --phase extract")
        return

    with open(DOWNLOAD_URLS_FILE, "r", encoding="utf-8") as f:
        records = json.load(f)

    aria2_file = os.path.join(OUTPUT_DIR, "aria2_download.txt")
    bat_file = os.path.join(OUTPUT_DIR, "aria2_download.bat")

    # 过滤出有下载链接的记录
    to_download = [r for r in records if r.get("pdf_urls") or r.get("caj_urls")]
    log.info(f"有下载链接: {len(to_download)} 篇")

    with open(aria2_file, "w", encoding="utf-8") as f:
        f.write("# aria2c 批量下载配置文件\n")
        f.write("# 使用方法: aria2c -c -i aria2_download.txt\n")
        f.write(f"# 共 {len(to_download)} 个文件\n")
        f.write(f"# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

        for record in to_download:
            title = record.get("title", "未知")
            dl_urls = record.get("pdf_urls", []) or record.get("caj_urls", [])
            if dl_urls:
                fname = make_filename(record.get("no", 0), title)
                # aria2c格式: URL\n  out=文件名\n
                f.write(f"{dl_urls[0]}\n")
                f.write(f"  out={fname}\n")

    with open(bat_file, "w", encoding="utf-8") as f:
        f.write("@echo off\n")
        f.write(f"REM CNKI 批量下载脚本\n")
        f.write(f"REM 共 {len(to_download)} 个文件\n")
        f.write(f"REM 双击运行此文件（需要先安装 aria2c）\n\n")
        f.write(f"aria2c -c -j 3 -s 3 --continue=true -i aria2_download.txt\n")
        f.write(f"pause\n")

    log.info(f"aria2c配置文件: {aria2_file}")
    log.info(f"Windows批处理: {bat_file}")
    log.info(f"\n使用方法:")
    log.info(f"  1. 安装 aria2c (https://github.com/aria2/aria2/releases)")
    log.info(f"  2. 双击 {bat_file} 开始下载")
    log.info(f"  或运行: aria2c -c -j 3 -s 3 --continue=true -i {aria2_file}")


def phase_stats(args):
    """统计 Excel 数据概况"""
    log.info("=" * 60)
    log.info("数据统计")
    log.info("=" * 60)

    articles = read_corpus_excel(EXCEL_FILE, SHEET_INDEX, only_chinese=True)
    log.info(f"中文文章总数: {len(articles)}")

    doi_count = sum(1 for a in articles if a["has_doi"])
    no_doi_count = len(articles) - doi_count
    log.info(f"有 DOI: {doi_count} ({doi_count/len(articles)*100:.1f}%)")
    log.info(f"无 DOI: {no_doi_count} ({no_doi_count/len(articles)*100:.1f}%)")

    # 按年份统计
    years = {}
    for a in articles:
        y = a.get("year", "未知")
        years[y] = years.get(y, 0) + 1
    log.info("\n按年份分布（前20）:")
    for y in sorted(years.keys(), reverse=True)[:20]:
        log.info(f"  {y}: {years[y]} 篇")

    # 按期刊统计
    journals = {}
    for a in articles:
        j = a.get("journal", "未知")
        if j and j != "None" and j != "":
            journals[j] = journals.get(j, 0) + 1
    top_journals = sorted(journals.items(), key=lambda x: -x[1])[:20]
    log.info("\n按期刊分布（前20）:")
    for j, c in top_journals:
        log.info(f"  {j[:40]}: {c} 篇")

    log.info("=" * 60)


def phase_generate_report(args):
    """
    生成下载URL报告和批量下载脚本
    """
    log.info("=" * 60)
    log.info("生成报告")
    log.info("=" * 60)

    if not os.path.exists(DOWNLOAD_URLS_FILE):
        log.error(f"URL文件不存在，请先运行 --phase extract")
        return

    with open(DOWNLOAD_URLS_FILE, "r", encoding="utf-8") as f:
        records = json.load(f)

    report_file = os.path.join(OUTPUT_DIR, "download_report.txt")
    batch_script = os.path.join(OUTPUT_DIR, "batch_download_urls.txt")

    with open(report_file, "w", encoding="utf-8") as rf, \
         open(batch_script, "w", encoding="utf-8") as bf:

        rf.write("=" * 80 + "\n")
        rf.write(f"CNKI 下载报告\n")
        rf.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        rf.write(f"总记录数: {len(records)}\n")
        rf.write("=" * 80 + "\n\n")

        bf.write("# CNKI 批量下载URL列表\n")
        bf.write("# 每行格式: 文章标题 | 下载URL\n")
        bf.write(f"# 生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

        ok_count = sum(1 for r in records if r.get("has_pdf") or r.get("has_caj"))
        login_count = sum(1 for r in records if r.get("status") == "login_required")

        rf.write(f"有下载链接: {ok_count}\n")
        rf.write(f"需登录: {login_count}\n\n")

        for record in records:
            title = record.get("title", "未知")
            doi = record.get("doi", "N/A")
            status = record.get("status", "unknown")

            rf.write(f"【{status}】{title}\n")
            rf.write(f"  DOI: {doi}\n")

            if record.get("pdf_urls"):
                rf.write(f"  PDF: {record['pdf_urls'][0]}\n")
                for url in record["pdf_urls"]:
                    bf.write(f"{title}\t{url}\tPDF\n")
            elif record.get("caj_urls"):
                rf.write(f"  CAJ: {record['caj_urls'][0]}\n")
                for url in record["caj_urls"]:
                    bf.write(f"{title}\t{url}\tCAJ\n")
            else:
                rf.write(f"  无下载链接\n")

            rf.write("\n")

    log.info(f"报告已生成: {report_file}")
    log.info(f"批量URL: {batch_script}")


# ===== 命令行入口 =====
def main():
    parser = argparse.ArgumentParser(
        description="CNKI 语料PDF下载工具 - 从Excel读取中文文章并通过知网下载",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  # 查看数据统计
  python download_corpus_pdfs.py --phase stats

  # 提取前100篇文章的下载链接
  python download_corpus_pdfs.py --phase extract --max-articles 100

  # 提取所有有DOI文章的下载链接（19530篇）
  python download_corpus_pdfs.py --phase extract --only-doi

  # 为非DOI文章查找DOI
  python download_corpus_pdfs.py --phase crossref --max-articles 200

  # 浏览器半自动下载PDF（弹出浏览器，遇到滑块时人工拖动一次）
  python download_corpus_pdfs.py --phase download-browser --max-articles 50

  # 大批量：分批提取+立即下载（失败项下次运行自动重试）
  python download_corpus_pdfs.py --phase pipeline --chunk-size 20

  # 一直等待人工过滑块（不设超时）
  python download_corpus_pdfs.py --phase pipeline --chunk-size 20 --captcha-wait 0

  # 请求式下载（知网出现滑块验证时会直接失败）
  python download_corpus_pdfs.py --phase download --max-articles 50

  # 使用浏览器Cookie下载（从浏览器复制Cookie后粘贴）
  python download_corpus_pdfs.py --phase download --cookies "SID=xxx; SID_kns_new=xxx"

  # 导出aria2c批量下载脚本（可用aria2c加速下载）
  python download_corpus_pdfs.py --phase export-aria2

  # 生成下载报告
  python download_corpus_pdfs.py --phase report
        """
    )

    parser.add_argument(
        "--phase", "-p",
        choices=["extract", "download", "download-browser", "pipeline", "crossref",
                 "stats", "report", "export-aria2", "all"],
        default="stats",
        help="执行阶段: extract=提取URL, download=requests下载PDF, "
             "download-browser=浏览器半自动下载（可人工过滑块）, "
             "pipeline=分批提取+立即下载（推荐大批量）, "
             "crossref=查找DOI, stats=统计, report=生成报告, "
             "export-aria2=导出aria2c脚本, all=全部"
    )
    parser.add_argument(
        "--max-articles", "-m",
        type=int,
        default=None,
        help="最大处理文章数（用于测试小批量；pipeline 中为本次处理的新文章数）"
    )
    parser.add_argument(
        "--chunk-size",
        type=int,
        default=20,
        help="pipeline 每批解析+下载的文章数（默认20）"
    )
    parser.add_argument(
        "--only-doi",
        action="store_true",
        help="仅处理有DOI的文章"
    )
    parser.add_argument(
        "--only-no-doi",
        action="store_true",
        help="仅处理无DOI的文章"
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=2.0,
        help="请求延迟秒数（默认2.0秒）"
    )
    parser.add_argument(
        "--cookies",
        type=str,
        default=None,
        help="知网登录Cookie字符串（从浏览器复制），如 'SID=xxx; SID_kns_new=xxx'"
    )
    parser.add_argument(
        "--captcha-wait",
        type=int,
        default=180,
        help="浏览器下载时等待人工完成滑块验证的秒数（默认180，0=不限时）"
    )
    parser.add_argument(
        "--login-wait",
        type=int,
        default=300,
        help="检测到登录失效时等待人工重新登录的秒数（默认300，0=不限时）"
    )
    parser.add_argument(
        "--browser-profile",
        type=str,
        default=None,
        help="Playwright 持久化浏览器配置目录（默认 data/browser_profile）"
    )
    parser.add_argument(
        "--retry-failed",
        action="store_true",
        help="清空失败/永久未解决进度，重试此前失败的文章（extract、download、pipeline 阶段生效）"
    )

    args = parser.parse_args()

    # 设置全局延迟
    global REQUEST_DELAY  # noqa: used before global in arg parsing
    REQUEST_DELAY = args.delay

    # 确保输出目录存在
    ensure_dir(OUTPUT_DIR)

    # 加载Cookie（如果提供）
    if args.cookies:
        log.info("使用提供的Cookie进行认证")
        # 设置到默认session（用于后续请求）
        os.environ["CNKI_COOKIES"] = args.cookies

    # 文件处理日志
    fh = logging.FileHandler(
        os.path.join(OUTPUT_DIR, f"download_{args.phase.replace('-', '_')}.log"),
        encoding="utf-8",
    )
    fh.setLevel(logging.DEBUG)
    log.addHandler(fh)

    if args.phase == "extract":
        phase_extract(args)
    elif args.phase == "download":
        phase_download(args)
    elif args.phase == "download-browser":
        phase_download_browser(args)
    elif args.phase == "pipeline":
        phase_pipeline(args)
    elif args.phase == "crossref":
        phase_crossref(args)
    elif args.phase == "stats":
        phase_stats(args)
    elif args.phase == "report":
        phase_generate_report(args)
    elif args.phase == "export-aria2":
        phase_export_aria2(args)
    elif args.phase == "all":
        log.info("执行完整流程...")
        phase_extract(args)
        phase_download(args)
        phase_generate_report(args)


if __name__ == "__main__":
    main()