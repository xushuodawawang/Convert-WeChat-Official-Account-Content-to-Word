from __future__ import annotations

from pathlib import Path

import streamlit as st

from docx_exporter import DocxExportError, ExportOptions, export_article_to_docx_bytes
from parser import ParseResult, WechatArticleParser
from utils import (
    article_preview_text,
    build_zip_bytes,
    ensure_directory,
    join_logs,
    parse_multiline_urls,
    resolve_line_spacing,
    sanitize_filename,
)


st.set_page_config(
    page_title="公众号文章导出 Word 工具",
    page_icon="📄",
    layout="wide",
)


def init_state() -> None:
    st.session_state.setdefault("parse_results", [])
    st.session_state.setdefault("export_artifacts", [])
    st.session_state.setdefault("status_rows", [])
    st.session_state.setdefault("global_logs", [])


def build_status_rows(results: list[ParseResult]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for result in results:
        rows.append(
            {
                "链接": result.url,
                "抓取状态": "抓取成功" if result.success else "抓取失败",
                "导出状态": "待导出" if result.success else "-",
                "说明": result.error if not result.success else "已完成正文解析，可预览后导出",
            }
        )
    return rows


def update_export_status(
    rows: list[dict[str, str]],
    url: str,
    export_status: str,
    message: str,
) -> None:
    for row in rows:
        if row["链接"] == url:
            row["导出状态"] = export_status
            row["说明"] = message
            return


def save_artifact_if_needed(directory_text: str, filename: str, data: bytes) -> str | None:
    if not directory_text.strip():
        return None
    directory = ensure_directory(directory_text)
    file_path = directory / filename
    file_path.write_bytes(data)
    return str(file_path)


def export_articles(
    successful_results: list[ParseResult],
    base_filename: str,
    export_dir: str,
    options: ExportOptions,
) -> tuple[list[dict[str, object]], list[str]]:
    artifacts: list[dict[str, object]] = []
    logs: list[str] = []

    total = len(successful_results)
    for index, result in enumerate(successful_results, start=1):
        article = result.article
        if article is None:
            continue

        if total == 1 and base_filename.strip():
            stem = sanitize_filename(base_filename, fallback=article.title)
        elif base_filename.strip():
            stem = sanitize_filename(f"{base_filename}_{index:02d}_{article.title}")
        else:
            stem = sanitize_filename(article.title or f"article_{index:02d}")

        output_name = f"{stem}.docx"
        try:
            file_bytes = export_article_to_docx_bytes(article, options)
            saved_path = save_artifact_if_needed(export_dir, output_name, file_bytes)
            artifacts.append(
                {
                    "url": result.url,
                    "title": article.title,
                    "file_name": output_name,
                    "bytes": file_bytes,
                    "saved_path": saved_path,
                }
            )
            if saved_path:
                logs.append(f"{output_name} 已保存到 {saved_path}")
            else:
                logs.append(f"{output_name} 已生成，可直接下载")
        except (DocxExportError, OSError, ValueError) as exc:
            artifacts.append(
                {
                    "url": result.url,
                    "title": article.title,
                    "file_name": output_name,
                    "bytes": None,
                    "saved_path": None,
                    "error": str(exc),
                }
            )
            logs.append(f"{output_name} 导出失败: {exc}")

    return artifacts, logs


def render_preview_section(results: list[ParseResult]) -> None:
    st.subheader("抓取预览")
    st.caption("成功抓取后会先展示标题和部分正文，确认无误后再导出为 Word。")

    for index, result in enumerate(results, start=1):
        if not result.success or result.article is None:
            continue

        article = result.article
        with st.expander(f"{index}. {article.title}", expanded=index == 1):
            st.markdown(
                "\n".join(
                    [
                        f"**原文链接**: {article.url}",
                        f"**作者**: {article.author or '未知'}",
                        f"**公众号**: {article.publisher or '未知'}",
                        f"**发布时间**: {article.publish_time or '未知'}",
                        f"**抓取方式**: {article.source}",
                    ]
                )
            )
            st.text_area(
                "正文预览",
                value=article_preview_text(article),
                height=220,
                key=f"preview_{index}_{article.url}",
                disabled=True,
            )
            if result.logs:
                st.code(join_logs(result.logs, max_lines=15), language="text")


def render_export_results(artifacts: list[dict[str, object]], export_dir: str) -> None:
    if not artifacts:
        return

    st.subheader("导出结果")
    success_artifacts = [item for item in artifacts if item.get("bytes")]
    failed_artifacts = [item for item in artifacts if not item.get("bytes")]

    if export_dir.strip():
        st.info(f"文件已尝试保存到目录: `{Path(export_dir).expanduser()}`")

    for item in success_artifacts:
        st.success(f"{item['file_name']} 导出成功")
        if item.get("saved_path"):
            st.caption(f"本地保存路径: `{item['saved_path']}`")

    for item in failed_artifacts:
        st.error(f"{item['file_name']} 导出失败: {item.get('error', '未知错误')}")

    if len(success_artifacts) == 1:
        artifact = success_artifacts[0]
        st.download_button(
            label=f"下载 {artifact['file_name']}",
            data=artifact["bytes"],
            file_name=artifact["file_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        return

    if len(success_artifacts) > 1:
        zip_name = "wechat_articles_export.zip"
        zip_bytes = build_zip_bytes(
            (str(item["file_name"]), item["bytes"]) for item in success_artifacts if item.get("bytes")
        )
        st.download_button(
            label=f"下载打包 ZIP ({len(success_artifacts)} 个文件)",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
        )


def main() -> None:
    init_state()

    st.title("微信公众号文章抓取与 Word 导出工具")
    st.caption("仅处理你主动提供的公开微信公众号文章链接，不提供任何绕过登录、付费或权限限制的能力。")

    with st.form("wechat_export_form"):
        st.markdown("### 抓取设置")
        urls_text = st.text_area(
            "公众号文章链接",
            height=180,
            placeholder="每行填写一个微信公众号文章链接，例如：\nhttps://mp.weixin.qq.com/s/xxxxxx",
        )

        input_col1, input_col2 = st.columns(2)
        export_dir = input_col1.text_input(
            "导出目录",
            value="",
            help="留空时仅提供浏览器下载；填写本地目录后会在本地额外保存 docx 文件。",
        )
        file_name = input_col2.text_input(
            "导出文件名",
            value="",
            help="单篇文章时作为完整文件名；批量时会作为文件名前缀，未填写则默认使用文章标题。",
        )

        style_col1, style_col2, style_col3, style_col4 = st.columns(4)
        line_spacing_mode = style_col1.selectbox("正文行距", ["single", "1.5", "double", "custom"], index=1)
        custom_spacing = style_col2.number_input(
            "自定义行距",
            min_value=0.8,
            max_value=5.0,
            value=1.8,
            step=0.1,
            disabled=line_spacing_mode != "custom",
        )
        font_name = style_col3.text_input("字体", value="宋体")
        font_size = style_col4.number_input("字号 (pt)", min_value=8, max_value=28, value=12, step=1)

        option_col1, option_col2 = st.columns(2)
        keep_images = option_col1.toggle("保留正文图片", value=True)
        use_playwright = option_col2.toggle("requests 失败时尝试 Playwright", value=False)

        submit_fetch = st.form_submit_button("开始抓取并导出", type="primary")

    if submit_fetch:
        valid_urls, invalid_lines = parse_multiline_urls(urls_text)
        st.session_state["export_artifacts"] = []
        st.session_state["global_logs"] = []

        if invalid_lines:
            st.warning("以下内容不是合法的公众号文章链接，已跳过:\n\n" + "\n".join(invalid_lines))

        if not valid_urls:
            st.session_state["parse_results"] = []
            st.session_state["status_rows"] = []
            st.error("没有可处理的合法链接。请检查输入内容后重试。")
        else:
            parser = WechatArticleParser(
                keep_images=keep_images,
                use_playwright_fallback=use_playwright,
            )
            results: list[ParseResult] = []
            progress = st.progress(0.0)
            with st.spinner("正在抓取文章内容，请稍候..."):
                for index, url in enumerate(valid_urls, start=1):
                    results.append(parser.fetch_article(url))
                    progress.progress(index / len(valid_urls))

            st.session_state["parse_results"] = results
            st.session_state["status_rows"] = build_status_rows(results)
            st.session_state["global_logs"] = [
                f"本次共处理 {len(valid_urls)} 个链接，成功 {sum(1 for item in results if item.success)} 个。"
            ]

    results = st.session_state.get("parse_results", [])
    status_rows = st.session_state.get("status_rows", [])

    failed_results = [result for result in results if not result.success]
    if failed_results:
        with st.expander("失败详情", expanded=False):
            for item in failed_results:
                st.error(f"{item.url}\n{item.error}")
                if item.logs:
                    st.code(join_logs(item.logs, max_lines=15), language="text")

    successful_results = [result for result in results if result.success and result.article is not None]
    if successful_results:
        render_preview_section(successful_results)

        st.markdown("### 确认导出")
        st.caption("如果上面的预览内容正常，可以点击下面的按钮生成 Word 文件。")
        if st.button("确认并生成 Word 文件", type="primary"):
            try:
                line_spacing = resolve_line_spacing(line_spacing_mode, custom_spacing)
            except ValueError as exc:
                st.error(str(exc))
            else:
                export_options = ExportOptions(
                    font_name=font_name.strip() or "宋体",
                    font_size=int(font_size),
                    line_spacing=line_spacing,
                    keep_images=keep_images,
                    link_position="start",
                )

                try:
                    artifacts, export_logs = export_articles(
                        successful_results=successful_results,
                        base_filename=file_name,
                        export_dir=export_dir,
                        options=export_options,
                    )
                except Exception as exc:  # pragma: no cover
                    st.session_state["export_artifacts"] = []
                    st.error(f"批量导出失败: {exc}")
                else:
                    st.session_state["export_artifacts"] = artifacts
                    st.session_state["global_logs"] = st.session_state.get("global_logs", []) + export_logs

                    for artifact in artifacts:
                        if artifact.get("bytes"):
                            update_export_status(
                                status_rows,
                                url=str(artifact["url"]),
                                export_status="导出成功",
                                message=f"已生成 {artifact['file_name']}",
                            )
                        else:
                            update_export_status(
                                status_rows,
                                url=str(artifact["url"]),
                                export_status="导出失败",
                                message=str(artifact.get("error", "Word 导出失败")),
                            )
                    st.session_state["status_rows"] = status_rows

    st.subheader("处理状态")
    if status_rows:
        st.table(status_rows)
    else:
        st.info("尚未开始处理。")

    artifacts = st.session_state.get("export_artifacts", [])
    if artifacts:
        render_export_results(artifacts, export_dir)

    global_logs = st.session_state.get("global_logs", [])
    if global_logs:
        st.subheader("日志信息")
        st.code(join_logs(global_logs, max_lines=50), language="text")


if __name__ == "__main__":
    main()
