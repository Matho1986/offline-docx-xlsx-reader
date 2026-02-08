package com.example.offlinedocxxlsxreader

import java.io.InputStream

class DocxReader {
    data class DocxContent(
        val html: String,
        val plainText: String
    )

    fun parse(inputStream: InputStream): DocxContent {
        val plainText = DocxTextExtractor.extract(inputStream)
        val htmlBody = buildPlainTextHtml(plainText)
        return DocxContent(
            html = HtmlTemplates.wrap(htmlBody),
            plainText = plainText.trim()
        )
    }

    private fun buildPlainTextHtml(text: String): String {
        if (text.isBlank()) {
            return "<p></p>"
        }
        val escaped = escapeHtml(text)
        return "<pre>$escaped</pre>"
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
