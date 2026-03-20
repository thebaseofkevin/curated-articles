# 依赖安装：
# pip install python-docx odfpy
# 可选：安装 LibreOffice（命令 `soffice`）以启用 PDF 转换

import argparse
import subprocess
import os
import shutil
import tempfile
import re

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from odf.opendocument import load
from odf.text import P
from odf import teletype
from odf.style import Style


class PaperFormatter:
    def __init__(self, input_file):
        self.input_file = os.path.abspath(input_file)
        base = os.path.splitext(self.input_file)[0]
        self.output_docx = base + ".docx"
        self.output_pdf = base + ".pdf"
        self.output_md = base + ".md"
        self.doc = Document()

        # 页面样式和封面
        self._set_style()
        self._add_cover_page()
        self._add_page_number()  # 从正文开始加页码

    def _set_style(self):
        """设置默认排版样式"""
        style = self.doc.styles['Normal']
        style.font.name = "Times New Roman"
        style._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')
        style.font.size = Pt(12)
        # 同时把同样的边距应用到所有当前的节
        for section in self.doc.sections:
            section.top_margin = Cm(3)
            section.bottom_margin = Cm(2.8)
            section.left_margin = Cm(2.7)
            section.right_margin = Cm(2.6)

    def _add_cover_page(self):
        """添加封面（居中标题、作者、日期）"""
        # 空行
        self.doc.add_paragraph()
        self.doc.add_paragraph()
        self.doc.add_paragraph()
        
        # 封面标题
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run("合集")
        run.font.size = Pt(30)
        run.bold = True
        run.font.name = "Times New Roman"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')

        # 空行
        self.doc.add_paragraph()
        self.doc.add_paragraph()
        self.doc.add_paragraph()

        # 作者信息
        p2 = self.doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run2 = p2.add_run("作者：无名")
        run2.font.size = Pt(16)
        run2.font.name = "Times New Roman"
        run2._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')

        # 日期
        p3 = self.doc.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run3 = p3.add_run("2026-03-12")
        run3.font.size = Pt(14)
        run3.font.name = "Times New Roman"
        run3._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')

        # 插入分页符，正文从第二页开始
        self.doc.add_page_break()

    def _add_page_number(self):
        """正文页脚添加页码"""
        # 从第二节开始，确保封面不显示页码
        section = self.doc.sections[-1]
        section.start_type = 0  # 新节开始
        # 应用正规的公文页边距
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
        section.footer_distance = Cm(2.0)
        footer = section.footer
        paragraph = footer.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 前部短横线
        paragraph.add_run("—")
        run = paragraph.add_run()
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')
        instrText = OxmlElement('w:instrText')
        instrText.text = "PAGE"
        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        # 后部短横线
        paragraph.add_run("—")

    def read_odt(self):
        doc = load(self.input_file)
        all_paragraphs = []
        unique_paragraphs = []
        seen = set()

        for p in list(doc.getElementsByType(P)):
            text = teletype.extractText(p)
            normalized = self._normalize_internal_spaces(text.strip())
            if not normalized:
                continue
            normalized = self._normalize_punctuation(normalized)
            all_paragraphs.append(normalized)

            if normalized in seen:
                if p.parentNode is not None:
                    p.parentNode.removeChild(p)
                continue

            seen.add(normalized)
            unique_paragraphs.append(normalized)
            self._set_paragraph_text(p, normalized)

        doc.save(self.input_file)
        return all_paragraphs, unique_paragraphs

    @staticmethod
    def _deduplicate_paragraphs(paragraphs):
        """按首次出现顺序去重段落。"""
        seen = set()
        unique = []
        for text in paragraphs:
            if text not in seen:
                seen.add(text)
                unique.append(text)
        return unique

    def _normalize_punctuation(self, text):
        """统一为中文标点"""
        # 保护序号点号（如 2. 、一. ），避免被替换成中文句号
        list_dot_token = "__LIST_DOT__"
        text = re.sub(
            r"(^|[\s（(【\[])([0-9]+|[一二三四五六七八九十]+)\.(?=\s)",
            r"\1\2" + list_dot_token,
            text
        )
        # 历史误替换修复：数字之间的中文句号视为小数点
        text = re.sub(r"(?<=\d)。(?=\d)", ".", text)
        # 先保护小数点，避免 2.6 被替换为 2。6
        decimal_token = "__DECIMAL_DOT__"
        text = re.sub(r"(?<=\d)\.(?=\d)", decimal_token, text)

        # 常见英文标点转中文标点
        basic_map = str.maketrans({
            ",": "，",
            ".": "。",
            "?": "？",
            "!": "！",
            ":": "：",
            ";": "；",
            "(": "（",
            ")": "）",
            "[": "【",
            "]": "】",
            "<": "《",
            ">": "》",
        })
        text = text.translate(basic_map)

        # 省略号与破折号统一
        text = re.sub(r"\.{3,}", "……", text)
        text = re.sub(r"—{2,}|-{2,}", "——", text)

        # 引号按出现顺序成对替换
        text = self._replace_paired_quotes(text, '"', "“", "”")
        text = self._replace_paired_quotes(text, "'", "‘", "’")
        # 删除指定引号字符
        text = text.replace("“", "").replace("”", "").replace("‘", "")
        # 连续相同标点只保留一个
        text = self._dedupe_repeated_punctuation(text)
        # 还原序号点号
        text = text.replace(list_dot_token, ".")
        # 还原小数点
        text = text.replace(decimal_token, ".")
        return text

    @staticmethod
    def _dedupe_repeated_punctuation(text):
        """删除连续重复标点，只保留一个。"""
        return re.sub(r"([，。！？：；、,.?!:;…—\-])\1+", r"\1", text)

    @staticmethod
    def _normalize_internal_spaces(text):
        """清理段落内部多余空格，保留英文单词间必要空格。"""
        text = text.replace("\u3000", " ").replace("\xa0", " ")
        text = re.sub(r"\s+", " ", text)
        # 中文字符之间不保留空格
        text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", text)
        # 去掉中文标点前后的多余空格
        text = re.sub(r"\s+([，。！？：；、）》】」』])", r"\1", text)
        text = re.sub(r"([（《【「『])\s+", r"\1", text)
        return text

    @staticmethod
    def _replace_paired_quotes(text, src, left, right):
        result = []
        open_quote = True
        for ch in text:
            if ch == src:
                result.append(left if open_quote else right)
                open_quote = not open_quote
            else:
                result.append(ch)
        return "".join(result)

    @staticmethod
    def _set_paragraph_text(paragraph, text):
        """替换段落文本，避免删除 Text 节点触发 odfpy 断言。"""
        first_text_node = None
        for child in list(paragraph.childNodes):
            if child.__class__.__name__ == "Text":
                if first_text_node is None:
                    first_text_node = child
                else:
                    child.data = ""
            else:
                paragraph.removeChild(child)

        if first_text_node is not None:
            first_text_node.data = text
        else:
            teletype.addTextToElement(paragraph, text)

    def add_content(self, text_list):
        for text in text_list:
            p = self.doc.add_paragraph(text, style='List Number')
            p_format = p.paragraph_format
            p_format.line_spacing = 1.5
            p_format.first_line_indent = Cm(0.75)
            p_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # 设置为两边对齐
            p_format.space_after = Pt(6)

            for run in p.runs:
                run.font.name = "Times New Roman"
                run._element.rPr.rFonts.set(qn('w:eastAsia'), 'SimSun')

    def export_pdf(self):
        soffice = shutil.which("soffice")
        if not soffice:
            print("[!] 未检测到 LibreOffice")
            return
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_docx = os.path.join(tmpdir, os.path.basename(self.output_docx))
            shutil.copy(self.output_docx, temp_docx)
            try:
                subprocess.run([
                    soffice,
                    "--headless",
                    "--convert-to", "pdf:writer_pdf_Export",
                    "--outdir", tmpdir,
                    temp_docx
                ], check=True)
                temp_pdf = os.path.join(
                    tmpdir,
                    os.path.splitext(os.path.basename(self.output_docx))[0] + ".pdf"
                )
                if os.path.exists(temp_pdf):
                    shutil.copy(temp_pdf, self.output_pdf)
                    print(f"[√] PDF 生成成功: {self.output_pdf}")
                else:
                    print("[!] PDF 未生成")
            except Exception as e:
                print("[!] 文件转换失败:", e)

    @staticmethod
    def _node_local_name(node):
        qname = getattr(node, "qname", None)
        if isinstance(qname, tuple) and len(qname) == 2:
            return qname[1]
        return None

    @staticmethod
    def _escape_markdown_text(text):
        if not text:
            return ""
        text = text.replace("\\", "\\\\")
        text = text.replace("`", "\\`")
        text = text.replace("*", "\\*")
        text = text.replace("_", "\\_")
        text = text.replace("[", "\\[")
        text = text.replace("]", "\\]")
        text = text.replace("|", "\\|")
        return text

    @staticmethod
    def _build_style_map(doc):
        style_map = {}
        for container in (doc.styles, doc.automaticstyles):
            for style in container.getElementsByType(Style):
                name = style.getAttribute("name")
                if name:
                    style_map[name] = style
        return style_map

    def _style_text_properties(self, style_name, style_map, seen=None):
        if not style_name:
            return {}
        if seen is None:
            seen = set()
        if style_name in seen:
            return {}
        seen.add(style_name)
        style = style_map.get(style_name)
        if not style:
            return {}

        properties = {}
        parent = style.getAttribute("parentstylename")
        if parent:
            properties.update(self._style_text_properties(parent, style_map, seen))
        for child in style.childNodes:
            if self._node_local_name(child) == "text-properties":
                for (_, attr_name), value in child.attributes.items():
                    properties[attr_name] = value
        return properties

    def _style_lineage(self, style_name, style_map):
        names = []
        seen = set()
        current = style_name
        while current and current not in seen:
            seen.add(current)
            names.append(current)
            style = style_map.get(current)
            if not style:
                break
            current = style.getAttribute("parentstylename")
        return names

    def _heading_level_from_style(self, style_name, style_map):
        lineage = self._style_lineage(style_name, style_map)
        for name in lineage:
            normalized = name.lower().replace("_20_", "_")
            match = re.search(r"heading[_ ]?(\d+)", normalized)
            if match:
                level = int(match.group(1))
                return max(1, min(6, level))
            if normalized == "heading":
                return 1
        return None

    def _inline_style_flags(self, style_name, style_map):
        props = self._style_text_properties(style_name, style_map)

        bold = any(props.get(key) == "bold" for key in (
            "font-weight", "font-weight-asian", "font-weight-complex"
        ))
        italic = any(props.get(key) in ("italic", "oblique") for key in (
            "font-style", "font-style-asian", "font-style-complex"
        ))
        underline = any(
            key.startswith("text-underline") and str(value).lower() not in ("none", "false")
            for key, value in props.items()
        )
        strike = any(
            key.startswith("text-line-through") and str(value).lower() not in ("none", "false")
            for key, value in props.items()
        )
        return {
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "strike": strike,
        }

    @staticmethod
    def _apply_inline_styles(text, flags):
        if not text.strip():
            return text
        if flags["bold"] and flags["italic"]:
            text = f"***{text}***"
        elif flags["bold"]:
            text = f"**{text}**"
        elif flags["italic"]:
            text = f"*{text}*"
        if flags["strike"]:
            text = f"~~{text}~~"
        if flags["underline"]:
            text = f"<u>{text}</u>"
        return text

    def _render_inline(self, node, style_map):
        if node.__class__.__name__ == "Text":
            return self._escape_markdown_text(node.data or "")

        local = self._node_local_name(node)
        if not local:
            return ""

        if local == "span":
            raw = "".join(self._render_inline(child, style_map) for child in node.childNodes)
            flags = self._inline_style_flags(node.getAttribute("stylename"), style_map)
            return self._apply_inline_styles(raw, flags)
        if local == "a":
            label = "".join(self._render_inline(child, style_map) for child in node.childNodes).strip()
            href = node.getAttribute("href")
            return f"[{label}]({href})" if href else label
        if local == "line-break":
            return "  \n"
        if local == "tab":
            return "    "
        if local == "s":
            count = node.getAttribute("c")
            try:
                n = int(count) if count else 1
            except Exception:
                n = 1
            return " " * n

        return "".join(self._render_inline(child, style_map) for child in getattr(node, "childNodes", []))

    def _render_paragraph(self, node, style_map):
        local = self._node_local_name(node)
        text = "".join(self._render_inline(child, style_map) for child in node.childNodes).strip()
        if not text:
            return ""

        if local == "h":
            try:
                level = int(node.getAttribute("outlinelevel") or 1)
            except Exception:
                level = 1
            level = max(1, min(6, level))
            return f"{'#' * level} {text}"

        style_name = node.getAttribute("stylename")
        heading_level = self._heading_level_from_style(style_name, style_map)
        if heading_level:
            return f"{'#' * heading_level} {text}"

        has_span_child = any(self._node_local_name(child) == "span" for child in node.childNodes)
        if not has_span_child:
            paragraph_flags = self._inline_style_flags(style_name, style_map)
            text = self._apply_inline_styles(text, paragraph_flags)

        # 合并相邻的粗体片段，避免出现 **A****B** 这种噪声
        while True:
            merged = re.sub(r"\*\*([^*\n]+)\*\*\s*\*\*([^*\n]+)\*\*", r"**\1\2**", text)
            if merged == text:
                break
            text = merged
        return text

    def _render_list(self, list_node, style_map, level=0):
        lines = []
        indent = "  " * level

        for child in list_node.childNodes:
            if self._node_local_name(child) != "list-item":
                continue

            first_line_written = False
            for item_child in child.childNodes:
                local = self._node_local_name(item_child)
                if local in ("p", "h"):
                    content = self._render_paragraph(item_child, style_map)
                    if not content:
                        continue
                    if not first_line_written:
                        lines.append(f"{indent}- {content}")
                        first_line_written = True
                    else:
                        lines.append(f"{indent}  {content}")
                elif local == "list":
                    lines.extend(self._render_list(item_child, style_map, level + 1))

            if not first_line_written:
                lines.append(f"{indent}-")
        return lines

    def _table_rows(self, table_node):
        rows = []
        for child in table_node.childNodes:
            local = self._node_local_name(child)
            if local == "table-row":
                rows.append(child)
            elif local in ("table-header-rows", "table-rows"):
                for row in child.childNodes:
                    if self._node_local_name(row) == "table-row":
                        rows.append(row)
        return rows

    def _render_table(self, table_node, style_map):
        rows = []
        for row in self._table_rows(table_node):
            cells = []
            for cell in row.childNodes:
                local = self._node_local_name(cell)
                if local not in ("table-cell", "covered-table-cell"):
                    continue
                parts = []
                for block in cell.childNodes:
                    block_local = self._node_local_name(block)
                    if block_local in ("p", "h"):
                        content = self._render_paragraph(block, style_map)
                        if content:
                            parts.append(content)
                cell_text = "<br>".join(parts).strip()
                cells.append(cell_text)
            if any(cell.strip() for cell in cells):
                rows.append(cells)

        if not rows:
            return []

        col_count = max(len(r) for r in rows)
        normalized_rows = [r + [""] * (col_count - len(r)) for r in rows]
        header = normalized_rows[0]

        lines = [
            "| " + " | ".join(header) + " |",
            "| " + " | ".join(["---"] * col_count) + " |",
        ]
        for r in normalized_rows[1:]:
            lines.append("| " + " | ".join(r) + " |")
        return lines

    def _render_blocks(self, nodes, style_map):
        lines = []
        for node in nodes:
            local = self._node_local_name(node)
            if local in ("p", "h"):
                content = self._render_paragraph(node, style_map)
                if content:
                    lines.append(content)
                    lines.append("")
            elif local == "list":
                list_lines = self._render_list(node, style_map)
                if list_lines:
                    lines.extend(list_lines)
                    lines.append("")
            elif local == "table":
                table_lines = self._render_table(node, style_map)
                if table_lines:
                    lines.extend(table_lines)
                    lines.append("")
            elif local in ("section",):
                child_lines = self._render_blocks(node.childNodes, style_map)
                if child_lines:
                    lines.extend(child_lines)
        return lines

    def export_markdown(self):
        title = os.path.splitext(os.path.basename(self.input_file))[0]
        try:
            doc = load(self.input_file)
            style_map = self._build_style_map(doc)
            lines = [f"# {self._escape_markdown_text(title)}", ""]
            lines.extend(self._render_blocks(doc.text.childNodes, style_map))
            content = "\n".join(lines).strip() + "\n"
            with open(self.output_md, "w", encoding="utf-8") as f:
                f.write(content)
            print(f"[√] MD 生成成功: {self.output_md}")
        except Exception as e:
            print("[!] MD 生成失败:", e)

    def cleanup_docx(self):
        if os.path.exists(self.output_docx):
            try:
                os.remove(self.output_docx)
                print(f"[√] 已删除中间 DOCX: {self.output_docx}")
            except Exception as e:
                print("[!] 删除 DOCX 失败:", e)

    def run(self, export_pdf=True, export_markdown=True):
        if export_markdown:
            self.export_markdown()
        if not export_pdf:
            return
        raw_paragraphs, paragraphs = self.read_odt()
        removed_count = len(raw_paragraphs) - len(paragraphs)
        print(f"读取到 {len(raw_paragraphs)} 个段落，去重后 {len(paragraphs)} 个（删除重复 {removed_count} 个）")
        print(f"[√] 已更新源 ODT: {self.input_file}")
        if export_pdf:
            self.add_content(paragraphs)
            self.doc.save(self.output_docx)
            print(f"[√] DOCX 生成成功: {self.output_docx}")
            self.export_pdf()
            self.cleanup_docx()


def main():
    parser = argparse.ArgumentParser(description="ODT 文档排版并按需导出 PDF 或 Markdown")
    parser.add_argument("file", help="输入 .odt 文件")
    parser.add_argument(
        "--format",
        choices=["pdf", "markdown", "both"],
        default="both",
        help="选择输出格式：pdf / markdown / both（默认）"
    )
    args = parser.parse_args()
    if not os.path.exists(args.file):
        print("文件不存在")
        return
    app = PaperFormatter(args.file)
    export_pdf = args.format in ("pdf", "both")
    export_markdown = args.format in ("markdown", "both")
    app.run(export_pdf=export_pdf, export_markdown=export_markdown)


if __name__ == "__main__":
    main()
