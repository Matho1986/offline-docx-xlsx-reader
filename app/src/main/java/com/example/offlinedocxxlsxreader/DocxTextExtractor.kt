package com.example.offlinedocxxlsxreader

import org.apache.poi.xwpf.usermodel.IBodyElement
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import org.apache.poi.xwpf.usermodel.XWPFTable
import java.io.InputStream

object DocxTextExtractor {

    fun extract(inputStream: InputStream): String {
        inputStream.use { stream ->
            val doc = XWPFDocument(stream)
            val sb = StringBuilder()

            // Wichtig: bodyElements hält die Reihenfolge aus dem Dokument
            for (el: IBodyElement in doc.bodyElements) {
                when (el.elementType) {
                    IBodyElement.ElementType.PARAGRAPH -> {
                        val p = el as XWPFParagraph
                        val text = p.text?.trimEnd().orEmpty()
                        if (text.isNotBlank()) {
                            sb.append(text).append("\n")
                        } else {
                            // Leerzeilen beibehalten (optional)
                            // sb.append("\n")
                        }
                    }

                    IBodyElement.ElementType.TABLE -> {
                        val table = el as XWPFTable
                        appendTable(table, sb)
                        sb.append("\n") // Abstand nach Tabellenblock
                    }

                    else -> {
                        // Andere Elemente (z.B. Content Controls) erstmal ignorieren
                    }
                }
            }

            // Aufräumen: nicht endlos Leerzeilen am Ende
            return sb.toString().trimEnd()
        }
    }

    private fun appendTable(table: XWPFTable, sb: StringBuilder) {
        for (row in table.rows) {
            val cellTexts = row.tableCells.map { cell ->
                // Eine Zelle kann mehrere Paragraphen haben
                val cellText = cell.paragraphs
                    .map { it.text.orEmpty().trim() }
                    .filter { it.isNotEmpty() }
                    .joinToString(separator = " ")

                cellText
            }

            // Ganze Tabellenzeile als tab-getrennte Zellen ausgeben
            sb.append(cellTexts.joinToString(separator = "\t"))
            sb.append("\n")
        }
    }
}
