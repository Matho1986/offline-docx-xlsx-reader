package com.example.offlinedocxxlsxreader

import org.apache.poi.xwpf.extractor.XWPFWordExtractor
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFRun
import java.io.InputStream

class DocxReader {
    data class DocxContent(
        val html: String,
        val plainText: String
    )

    fun parse(inputStream: InputStream): DocxContent {
        inputStream.use { stream ->
            val document = XWPFDocument(stream)
            val htmlBody = buildHtmlBody(document.paragraphs)
            val plainText = XWPFWordExtractor(document).text
            return DocxContent(
                html = HtmlTemplates.wrap(htmlBody),
                plainText = plainText.trim()
            )
        }
    }

    private fun buildHtmlBody(paragraphs: List<XWPFParagraph>): String {
        val builder = StringBuilder()
        var listOpen = false

        for (paragraph in paragraphs) {
            val text = buildParagraphRuns(paragraph)
            if (text.isBlank()) {
                continue
            }
            val isList = paragraph.numID != null
            if (isList && !listOpen) {
                builder.append("<ul>")
                listOpen = true
            } else if (!isList && listOpen) {
                builder.append("</ul>")
                listOpen = false
            }

            if (isList) {
                builder.append("<li>").append(text).append("</li>")
            } else {
                val headingTag = headingTag(paragraph)
                if (headingTag != null) {
                    builder.append("<").append(headingTag).append(">")
                        .append(text)
                        .append("</").append(headingTag).append(">")
                } else {
                    builder.append("<p>").append(text).append("</p>")
                }
            }
        }

        if (listOpen) {
            builder.append("</ul>")
        }

        return builder.toString()
    }

    private fun headingTag(paragraph: XWPFParagraph): String? {
        val style = paragraph.style ?: return null
        val normalized = style.lowercase()
        if (!normalized.contains("heading")) return null
        val level = normalized.filter { it.isDigit() }.toIntOrNull() ?: return "h2"
        val clamped = level.coerceIn(1, 6)
        return "h$clamped"
    }

    private fun buildParagraphRuns(paragraph: XWPFParagraph): String {
        val builder = StringBuilder()
        for (run in paragraph.runs) {
            val text = run.textValue() ?: continue
            if (text.isBlank()) continue
            var segment = escapeHtml(text)
            if (run.isBold) segment = "<strong>$segment</strong>"
            if (run.isItalic) segment = "<em>$segment</em>"
            if (run.underline != null && run.underline.value != null) {
                segment = "<u>$segment</u>"
            }
            builder.append(segment)
        }
        return builder.toString()
    }

    private fun XWPFRun.textValue(): String? {
        return text() ?: getText(0)
    }

    private fun escapeHtml(text: String): String {
        return text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("\"", "&quot;")
            .replace("'", "&#39;")
    }
}
