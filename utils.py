from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path
from typing import Iterable
from urllib.parse import urlsplit, urlunsplit


WECHAT_HOSTS = {"mp.weixin.qq.com", "weixin.qq.com"}


def normalize_wechat_url(raw_url: str) -> str:
    value = (raw_url or "").strip()
    if not value:
        raise ValueError("链接不能为空")

    if value.startswith("//"):
        value = f"https:{value}"
    elif not re.match(r"^https?://", value, flags=re.IGNORECASE):
        value = f"https://{value}"

    parsed = urlsplit(value)
    if parsed.scheme not in {"http", "https"}:
        raise ValueError("仅支持 http 或 https 链接")

    host = parsed.netloc.lower()
    if host not in WECHAT_HOSTS and not host.endswith(".weixin.qq.com"):
        raise ValueError("仅支持微信公众号公开文章链接")

    if not parsed.path:
        raise ValueError("链接缺少路径")

    return urlunsplit((parsed.scheme, parsed.netloc, parsed.path, parsed.query, ""))


def parse_multiline_urls(text: str) -> tuple[list[str], list[str]]:
    valid_urls: list[str] = []
    invalid_lines: list[str] = []
    seen: set[str] = set()

    for raw_line in (text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        try:
            normalized = normalize_wechat_url(line)
        except ValueError:
            invalid_lines.append(line)
            continue

        if normalized not in seen:
            valid_urls.append(normalized)
            seen.add(normalized)

    return valid_urls, invalid_lines


def sanitize_filename(filename: str, fallback: str = "wechat_article") -> str:
    value = (filename or "").strip() or fallback
    value = re.sub(r"[\\/:*?\"<>|]+", "_", value)
    value = re.sub(r"\s+", " ", value).strip(" .")
    value = value[:120].strip()
    return value or fallback


def ensure_directory(path_text: str) -> Path:
    directory = Path(path_text).expanduser().resolve()
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def resolve_line_spacing(mode: str, custom_value: float | int | None = None) -> float:
    normalized_mode = (mode or "single").lower().strip()
    preset_map = {
        "single": 1.0,
        "1.5": 1.5,
        "double": 2.0,
    }
    if normalized_mode == "custom":
        if custom_value is None:
            raise ValueError("自定义行距不能为空")
        value = float(custom_value)
        if value <= 0:
            raise ValueError("自定义行距必须大于 0")
        return value

    if normalized_mode not in preset_map:
        raise ValueError(f"不支持的行距模式: {mode}")
    return preset_map[normalized_mode]


def build_zip_bytes(files: Iterable[tuple[str, bytes]]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
        for filename, content in files:
            archive.writestr(filename, content)
    buffer.seek(0)
    return buffer.read()


def article_preview_text(article: object, max_blocks: int = 5, max_chars: int = 500) -> str:
    blocks = getattr(article, "blocks", []) or []
    segments: list[str] = []
    current_length = 0

    for block in blocks:
        block_type = getattr(block, "type", "")
        if block_type not in {"heading", "paragraph", "quote"}:
            continue

        text = (getattr(block, "text", "") or "").strip()
        if not text:
            continue

        segments.append(text)
        current_length += len(text)
        if len(segments) >= max_blocks or current_length >= max_chars:
            break

    preview = "\n\n".join(segments)
    if len(preview) > max_chars:
        preview = preview[: max_chars - 3].rstrip() + "..."
    return preview or "未提取到可预览的正文内容。"


def join_logs(logs: Iterable[str], max_lines: int = 20) -> str:
    collected = [line for line in logs if line]
    if len(collected) > max_lines:
        collected = collected[-max_lines:]
    return "\n".join(collected)
