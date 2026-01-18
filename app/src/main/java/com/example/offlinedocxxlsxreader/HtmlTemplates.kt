package com.example.offlinedocxxlsxreader

object HtmlTemplates {
    private const val BASE_CSS = """
        body {
            font-family: sans-serif;
            line-height: 1.6;
            color: #111827;
            margin: 16px;
            max-width: 960px;
        }
        h1, h2, h3, h4, h5, h6 {
            color: #1e3a8a;
            margin-top: 24px;
        }
        ul {
            padding-left: 20px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 12px;
            font-size: 14px;
        }
        th, td {
            border: 1px solid #e5e7eb;
            padding: 6px 8px;
            vertical-align: top;
            text-align: left;
        }
        .table-container {
            overflow-x: auto;
        }
    """

    fun wrap(bodyContent: String): String {
        return """
            <!doctype html>
            <html lang="de">
                <head>
                    <meta charset="utf-8" />
                    <meta name="viewport" content="width=device-width, initial-scale=1" />
                    <style>
                        $BASE_CSS
                    </style>
                </head>
                <body>
                    $bodyContent
                </body>
            </html>
        """.trimIndent()
    }
}
