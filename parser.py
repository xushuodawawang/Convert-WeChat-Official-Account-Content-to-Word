from __future__ import annotations

import html
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Iterable
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup, NavigableString, Tag


REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
    "Cache-Control": "no-cache",
    "Pragma": "no-cache",
}

HEADING_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6"}
PARAGRAPH_TAGS = {"p", "blockquote", "pre"}
LIST_TAGS = {"ul", "ol"}
IMAGE_ATTRS = ("data-src", "data-original", "src")
UNWANTED_TAGS = {"script", "style", "noscript", "iframe", "form", "button", "svg"}
UNWANTED_KEYWORDS = (
    "recommend",
    "related",
    "comment",
    "reward",
    "qr_code",
    "footer",
    "original_area",
    "js_tags",
    "profile_card",
    "page_bottom",
    "weapp_display_element",
)


class ArticleParserError(Exception):
    pass


class ArticleFetchError(ArticleParserError):
    pass


class ArticleParseError(ArticleParserError):
    pass


@dataclass
class TextRun:
    text: str
    bold: bool = False
    italic: bool = False


@dataclass
class ContentBlock:
    type: str
    text: str = ""
    level: int = 0
    runs: list[TextRun] = field(default_factory=list)
    image_url: str | None = None
    image_bytes: bytes | None = None
    image_name: str | None = None
    alt_text: str = ""


@dataclass
class ArticleData:
    url: str
    title: str
    author: str = "未知"
    publisher: str = "未知"
    publish_time: str = "未知"
    blocks: list[ContentBlock] = field(default_factory=list)
    raw_text: str = ""
    source: str = "requests"


@dataclass
class ParseResult:
    url: str
    success: bool
    article: ArticleData | None = None
    error: str = ""
    logs: list[str] = field(default_factory=list)


class WechatArticleParser:
    def __init__(self, keep_images: bool = True, use_playwright_fallback: bool = False, request_timeout: int = 20) -> None:
        self.keep_images = keep_images
        self.use_playwright_fallback = use_playwright_fallback
        self.request_timeout = request_timeout
        self.session = requests.Session()
        self.session.headers.update(REQUEST_HEADERS)

    def fetch_article(self, url: str) -> ParseResult:
        logs = [f"开始处理: {url}"]
        try:
            html_text = self._fetch_with_requests(url, logs)
            article = self._parse_article(url, html_text, source="requests", logs=logs)
            logs.append("requests 抓取与解析成功")
            return ParseResult(url=url, success=True, article=article, logs=logs)
        except ArticleParserError as exc:
            logs.append(f"requests 失败: {exc}")
            if not self.use_playwright_fallback:
                return ParseResult(url=url, success=False, error=str(exc), logs=logs)
        except Exception as exc:  # pragma: no cover
            logs.append(f"requests 异常: {exc}")
            if not self.use_playwright_fallback:
                return ParseResult(url=url, success=False, error=f"抓取异常: {exc}", logs=logs)

        try:
            html_text = self._fetch_with_playwright(url, logs)
            article = self._parse_article(url, html_text, source="playwright", logs=logs)
            logs.append("Playwright 回退成功")
            return ParseResult(url=url, success=True, article=article, logs=logs)
        except Exception as exc:
            logs.append(f"Playwright 回退失败: {exc}")
            return ParseResult(url=url, success=False, error=f"requests 和 Playwright 均失败: {exc}", logs=logs)

    def _fetch_with_requests(self, url: str, logs: list[str]) -> str:
        try:
            response = self.session.get(url, timeout=(10, self.request_timeout), allow_redirects=True)
        except requests.Timeout as exc:
            raise ArticleFetchError("请求超时，请稍后重试") from exc
        except requests.RequestException as exc:
            raise ArticleFetchError(f"页面无法访问: {exc}") from exc

        if response.status_code >= 400:
            raise ArticleFetchError(f"页面访问失败，状态码: {response.status_code}")
        if response.apparent_encoding:
            response.encoding = response.apparent_encoding
        logs.append(f"requests 状态码: {response.status_code}")
        return response.text

    def _fetch_with_playwright(self, url: str, logs: list[str]) -> str:
        try:
            from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
            from playwright.sync_api import sync_playwright
        except ImportError as exc:
            raise ArticleFetchError("未安装 Playwright。请先安装 playwright 并执行 python -m playwright install chromium。") from exc

        try:
            with sync_playwright() as playwright:
                browser = playwright.chromium.launch(headless=True)
                page = browser.new_page(user_agent=REQUEST_HEADERS["User-Agent"])
                page.set_default_timeout(self.request_timeout * 1000)
                page.goto(url, wait_until="networkidle")
                page.wait_for_timeout(1000)
                content = page.content()
                browser.close()
                logs.append("Playwright 已完成页面渲染")
                return content
        except PlaywrightTimeoutError as exc:
            raise ArticleFetchError("Playwright 抓取超时") from exc
        except Exception as exc:
            raise ArticleFetchError(f"Playwright 抓取失败: {exc}") from exc

    def _parse_article(self, url: str, html_text: str, source: str, logs: list[str]) -> ArticleData:
        soup = BeautifulSoup(html_text, "lxml")
        content_root = soup.find(id="js_content") or soup.select_one(".rich_media_content") or soup.select_one("#img-content")
        if not content_root:
            raise ArticleParseError("未找到正文区域，可能不是标准公众号文章页")

        self._clean_content_root(content_root)
        title = self._extract_title(soup, html_text)
        author = self._extract_author(soup, html_text)
        publisher = self._extract_publisher(soup, html_text)
        publish_time = self._extract_publish_time(soup, html_text)
        blocks = self._extract_blocks(content_root, url, logs)
        raw_text = "
".join(block.text.strip() for block in blocks if block.type in {"heading", "paragraph", "quote"} and block.text.strip()).strip()
        if not raw_text:
            raise ArticleParseError("正文提取为空，页面结构可能已变化")

        logs.append(f"正文块数量: {len(blocks)}")
        return ArticleData(url=url, title=title or "未命名文章", author=author or "未知", publisher=publisher or "未知", publish_time=publish_time or "未知", blocks=blocks, raw_text=raw_text, source=source)

    def _clean_content_root(self, root: Tag) -> None:
        for tag in root.find_all(UNWANTED_TAGS):
            tag.decompose()
        for tag in list(root.find_all(True)):
            classes = " ".join(tag.get("class", [])) if tag.get("class") else ""
            marker = f"{tag.get('id', '')} {classes}".lower()
            style = str(tag.get("style", "")).lower()
            if any(keyword in marker for keyword in UNWANTED_KEYWORDS):
                tag.decompose()
                continue
            if "display:none" in style or "visibility:hidden" in style:
                tag.decompose()

    def _extract_title(self, soup: BeautifulSoup, html_text: str) -> str:
        for selector in ["#activity-name", "h1.rich_media_title", "meta[property='og:title']", "meta[name='twitter:title']", "title"]:
            node = soup.select_one(selector)
            if not node:
                continue
            value = node.get("content", "") if node.name == "meta" else node.get_text(" ", strip=True)
            if value.strip():
                return self._normalize_text(value)
        return self._search_regex(html_text, [r"vars+msg_titles*=s*'([^']+)'", r'"title"s*:s*"([^"]+)"'])

    def _extract_author(self, soup: BeautifulSoup, html_text: str) -> str:
        for selector in ["#js_author_name", "#author_name", "meta[property='article:author']"]:
            node = soup.select_one(selector)
            if not node:
                continue
            value = node.get("content", "") if node.name == "meta" else node.get_text(" ", strip=True)
            normalized = self._normalize_text(value)
            if normalized and normalized != "微信号":
                return normalized
        return self._search_regex(html_text, [r"vars+authors*=s*htmlDecode("([^"]*)")", r"vars+authors*=s*'([^']*)'", r'"author"s*:s*"([^"]+)"'], default="未知")

    def _extract_publisher(self, soup: BeautifulSoup, html_text: str) -> str:
        for selector in ["#js_name", "meta[name='profile_nickname']", "meta[property='og:site_name']"]:
            node = soup.select_one(selector)
            if not node:
                continue
            value = node.get("content", "") if node.name == "meta" else node.get_text(" ", strip=True)
            normalized = self._normalize_text(value)
            if normalized:
                return normalized
        return self._search_regex(html_text, [r"vars+nicknames*=s*htmlDecode("([^"]*)")", r'"nickname"s*:s*"([^"]+)"'], default="未知")

    def _extract_publish_time(self, soup: BeautifulSoup, html_text: str) -> str:
        for selector in ["#publish_time", "em#publish_time", "meta[property='article:published_time']"]:
            node = soup.select_one(selector)
            if not node:
                continue
            value = node.get("content", "") if node.name == "meta" else node.get_text(" ", strip=True)
            normalized = self._normalize_text(value)
            if normalized:
                return normalized
        timestamp = self._search_regex(html_text, [r"cts*=s*['"]?(d{10})", r'"publish_time"s*:s*"([^"]+)"'])
        if timestamp.isdigit() and len(timestamp) == 10:
            try:
                return datetime.fromtimestamp(int(timestamp)).strftime("%Y-%m-%d %H:%M:%S")
            except (OSError, ValueError):
                return "未知"
        return timestamp or "未知"

    def _extract_blocks(self, root: Tag, base_url: str, logs: list[str]) -> list[ContentBlock]:
        blocks: list[ContentBlock] = []
        self._walk(root, blocks, base_url, logs)
        return self._dedupe_blocks(blocks)

    def _walk(self, parent: Tag, blocks: list[ContentBlock], base_url: str, logs: list[str]) -> None:
        for child in parent.children:
            if isinstance(child, NavigableString) or not isinstance(child, Tag):
                continue
            if child.name in HEADING_TAGS:
                runs = self._extract_runs(child)
                text = self._runs_text(runs)
                if text:
                    blocks.append(ContentBlock(type="heading", text=text, level=int(child.name[1]), runs=runs))
                self._append_images(child, blocks, base_url, logs)
                continue
            if child.name in PARAGRAPH_TAGS:
                self._append_paragraph(child, blocks)
                self._append_images(child, blocks, base_url, logs)
                continue
            if child.name in LIST_TAGS:
                ordered = child.name == "ol"
                for idx, item in enumerate(child.find_all("li", recursive=False), start=1):
                    self._append_paragraph(item, blocks, prefix=f"{idx}. " if ordered else "• ")
                    self._append_images(item, blocks, base_url, logs)
                continue
            if child.name == "img":
                self._append_images(child, blocks, base_url, logs)
                continue
            direct_runs = self._extract_direct_runs(child)
            direct_text = self._runs_text(direct_runs)
            if direct_text:
                blocks.append(ContentBlock(type="paragraph", text=direct_text, runs=direct_runs))
            self._walk(child, blocks, base_url, logs)

    def _append_paragraph(self, node: Tag, blocks: list[ContentBlock], prefix: str = "") -> None:
        runs = self._extract_runs(node)
        if prefix:
            runs = [TextRun(prefix)] + runs
        text = self._runs_text(runs)
        if text:
            blocks.append(ContentBlock(type="paragraph", text=text, runs=runs))

    def _append_images(self, node: Tag, blocks: list[ContentBlock], base_url: str, logs: list[str]) -> None:
        if not self.keep_images:
            return
        image_tags: Iterable[Tag] = [node] if node.name == "img" else node.find_all("img", recursive=False)
        for image_tag in image_tags:
            image_url = self._extract_image_url(image_tag, base_url)
            if not image_url:
                continue
            image_bytes, image_name = self._download_image(image_url, logs)
            blocks.append(ContentBlock(type="image", image_url=image_url, image_bytes=image_bytes, image_name=image_name, alt_text=self._normalize_text(image_tag.get("alt", ""))))

    def _download_image(self, image_url: str, logs: list[str]) -> tuple[bytes | None, str | None]:
        try:
            response = self.session.get(image_url, timeout=(10, self.request_timeout), stream=True)
            response.raise_for_status()
            content = response.content
            if not content:
                return None, None
            filename = image_url.rstrip("/").rsplit("/", 1)[-1] or "image.jpg"
            if "." not in filename:
                filename += ".jpg"
            logs.append(f"图片下载成功: {filename}")
            return content, filename
        except requests.RequestException as exc:
            logs.append(f"图片下载失败: {image_url} ({exc})")
            return None, None

    def _extract_image_url(self, image_tag: Tag, base_url: str) -> str | None:
        for attr in IMAGE_ATTRS:
            value = image_tag.get(attr)
            if not value:
                continue
            candidate = str(value).strip()
            if not candidate or candidate.startswith("data:"):
                continue
            if candidate.startswith("//"):
                return f"https:{candidate}"
            return urljoin(base_url, candidate)
        return None

    def _extract_runs(self, node: Tag | NavigableString, bold: bool = False, italic: bool = False) -> list[TextRun]:
        if isinstance(node, NavigableString):
            text = self._normalize_inline(str(node))
            return [TextRun(text=text, bold=bold, italic=italic)] if text else []
        if not isinstance(node, Tag):
            return []
        if node.name == "img":
            return []
        if node.name == "br":
            return [TextRun(text="
", bold=bold, italic=italic)]
        style = str(node.get("style", "")).lower()
        next_bold = bold or node.name in {"b", "strong"} or "font-weight:bold" in style or "font-weight:700" in style
        next_italic = italic or node.name in {"i", "em"} or "font-style:italic" in style
        runs: list[TextRun] = []
        for child in node.children:
            runs.extend(self._extract_runs(child, bold=next_bold, italic=next_italic))
        return self._merge_runs(runs)

    def _extract_direct_runs(self, node: Tag) -> list[TextRun]:
        runs: list[TextRun] = []
        for child in node.children:
            if isinstance(child, NavigableString):
                runs.extend(self._extract_runs(child))
            elif isinstance(child, Tag) and child.name not in HEADING_TAGS | PARAGRAPH_TAGS | LIST_TAGS and child.name != "img":
                runs.extend(self._extract_runs(child))
        return self._merge_runs(runs)

    def _merge_runs(self, runs: list[TextRun]) -> list[TextRun]:
        merged: list[TextRun] = []
        for run in runs:
            if not run.text:
                continue
            if merged and merged[-1].bold == run.bold and merged[-1].italic == run.italic:
                merged[-1].text += run.text
            else:
                merged.append(TextRun(text=run.text, bold=run.bold, italic=run.italic))
        return [run for run in merged if run.text and run.text.strip("
 ")]

    def _runs_text(self, runs: list[TextRun]) -> str:
        return self._normalize_text("".join(run.text for run in runs))

    def _normalize_inline(self, text: str) -> str:
        cleaned = html.unescape(text).replace(" ", " ").replace("​", "")
        return re.sub(r"s+", " ", cleaned)

    def _normalize_text(self, text: str) -> str:
        cleaned = html.unescape(text).replace(" ", " ").replace("​", "")
        cleaned = cleaned.replace("
", "
")
        cleaned = re.sub(r"[ 	]+", " ", cleaned)
        cleaned = re.sub(r"
{3,}", "

", cleaned)
        return cleaned.strip()

    def _dedupe_blocks(self, blocks: list[ContentBlock]) -> list[ContentBlock]:
        cleaned: list[ContentBlock] = []
        previous = ""
        for block in blocks:
            if block.type in {"heading", "paragraph", "quote"} and not block.text.strip():
                continue
            signature = f"{block.type}:{block.text}:{block.image_url}"
            if signature == previous:
                continue
            cleaned.append(block)
            previous = signature
        return cleaned

    def _search_regex(self, text: str, patterns: list[str], default: str = "") -> str:
        for pattern in patterns:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                return self._normalize_text(match.group(1))
        return default
