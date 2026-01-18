package com.example.offlinedocxxlsxreader

import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.FormulaEvaluator
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.InputStream

class XlsxReader {
    data class XlsxSheet(
        val name: String,
        val html: String,
        val tsv: String
    )

    data class XlsxContent(
        val sheets: List<XlsxSheet>
    )

    fun parse(inputStream: InputStream): XlsxContent {
        inputStream.use { stream ->
            val workbook = XSSFWorkbook(stream)
            val formatter = DataFormatter()
            val evaluator = workbook.creationHelper.createFormulaEvaluator()
            val sheets = mutableListOf<XlsxSheet>()

            for (i in 0 until workbook.numberOfSheets) {
                val sheet = workbook.getSheetAt(i)
                val rows = mutableListOf<List<String>>()
                var maxColumns = 0
                if (sheet.lastRowNum >= sheet.firstRowNum) {
                    for (rowIndex in sheet.firstRowNum..sheet.lastRowNum) {
                        val row = sheet.getRow(rowIndex)
                        val values = if (row != null) {
                            readRow(row, formatter, evaluator)
                        } else {
                            emptyList()
                        }
                        maxColumns = maxOf(maxColumns, values.size)
                        rows.add(values)
                    }
                }

                val normalizedRows = rows.map { row ->
                    if (row.size < maxColumns) {
                        row + List(maxColumns - row.size) { "" }
                    } else {
                        row
                    }
                }

                val htmlBody = buildHtmlTable(normalizedRows)
                val tsv = buildTsv(normalizedRows)

                sheets.add(
                    XlsxSheet(
                        name = sheet.sheetName,
                        html = HtmlTemplates.wrap(htmlBody),
                        tsv = tsv.trim()
                    )
                )
            }

            return XlsxContent(sheets)
        }
    }

    private fun readRow(
        row: Row,
        formatter: DataFormatter,
        evaluator: FormulaEvaluator
    ): List<String> {
        val values = mutableListOf<String>()
        val lastCell = row.lastCellNum.toInt().coerceAtLeast(0)
        for (cellIndex in 0 until lastCell) {
            val cell = row.getCell(cellIndex)
            val cellValue = if (cell == null) {
                ""
            } else {
                formatter.formatCellValue(cell, evaluator)
            }
            values.add(cellValue)
        }
        return values
    }

    private fun buildHtmlTable(rows: List<List<String>>): String {
        val builder = StringBuilder()
        builder.append("<div class=\"table-container\"><table><tbody>")
        for (row in rows) {
            builder.append("<tr>")
            for (cell in row) {
                builder.append("<td>").append(escapeHtml(cell)).append("</td>")
            }
            builder.append("</tr>")
        }
        builder.append("</tbody></table></div>")
        return builder.toString()
    }

    private fun buildTsv(rows: List<List<String>>): String {
        return rows.joinToString("\n") { row ->
            row.joinToString("\t") { cell -> cell.replace("\t", " ") }
        }
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
