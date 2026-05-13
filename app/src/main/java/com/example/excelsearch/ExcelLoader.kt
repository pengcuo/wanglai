package com.example.excelsearch

import android.util.Xml
import org.xmlpull.v1.XmlPullParser
import java.io.ByteArrayInputStream
import java.io.InputStream
import java.util.zip.ZipInputStream

/**
 * Minimal .xlsx reader. An xlsx file is a ZIP that contains:
 *   - xl/sharedStrings.xml  (deduplicated string table, optional)
 *   - xl/worksheets/sheetN.xml  (actual cells; values reference the string table)
 * We only need the first sheet and the first four columns.
 */
object ExcelLoader {

    private val HEADER_FIRST_CELLS = setOf("名字", "姓名", "name", "Name", "NAME")

    fun load(input: InputStream): List<ExcelRow> {
        val files = HashMap<String, ByteArray>()
        ZipInputStream(input).use { zis ->
            var entry = zis.nextEntry
            while (entry != null) {
                if (!entry.isDirectory) {
                    files[entry.name] = zis.readBytes()
                }
                entry = zis.nextEntry
            }
        }

        val shared = files["xl/sharedStrings.xml"]?.let(::parseSharedStrings) ?: emptyList()

        val sheetBytes = files["xl/worksheets/sheet1.xml"]
            ?: files.entries.firstOrNull {
                it.key.startsWith("xl/worksheets/sheet") && it.key.endsWith(".xml")
            }?.value
            ?: return emptyList()

        val rawRows = parseSheet(sheetBytes, shared)

        return rawRows.mapIndexedNotNull { idx, cells ->
            val name = cells.getOrNull(0)?.trim().orEmpty()
            if (name.isEmpty()) return@mapIndexedNotNull null
            if (idx == 0 && name in HEADER_FIRST_CELLS) return@mapIndexedNotNull null
            ExcelRow(
                name = name,
                location = cells.getOrNull(1)?.trim().orEmpty(),
                amount = cells.getOrNull(2)?.trim().orEmpty(),
                page = extractPageNumber(cells.getOrNull(3)?.trim().orEmpty()),
            )
        }
    }

    private fun parseSharedStrings(bytes: ByteArray): List<String> {
        val result = ArrayList<String>()
        val parser = Xml.newPullParser()
        parser.setInput(ByteArrayInputStream(bytes), "UTF-8")
        val buf = StringBuilder()
        var inSi = false
        var event = parser.eventType
        while (event != XmlPullParser.END_DOCUMENT) {
            when (event) {
                XmlPullParser.START_TAG -> if (parser.name == "si") {
                    inSi = true
                    buf.setLength(0)
                }
                XmlPullParser.TEXT -> if (inSi) buf.append(parser.text)
                XmlPullParser.END_TAG -> if (parser.name == "si") {
                    result.add(buf.toString())
                    inSi = false
                }
            }
            event = parser.next()
        }
        return result
    }

    private fun parseSheet(bytes: ByteArray, shared: List<String>): List<List<String>> {
        val rows = ArrayList<List<String>>()
        val parser = Xml.newPullParser()
        parser.setInput(ByteArrayInputStream(bytes), "UTF-8")

        var currentRow: HashMap<Int, String>? = null
        var cellRef: String? = null
        var cellType: String? = null
        val text = StringBuilder()
        var inValue = false
        var inInlineText = false

        var event = parser.eventType
        while (event != XmlPullParser.END_DOCUMENT) {
            when (event) {
                XmlPullParser.START_TAG -> when (parser.name) {
                    "row" -> currentRow = HashMap()
                    "c" -> {
                        cellRef = parser.getAttributeValue(null, "r")
                        cellType = parser.getAttributeValue(null, "t")
                        text.setLength(0)
                    }
                    "v" -> {
                        inValue = true
                        text.setLength(0)
                    }
                    "t" -> if (cellType == "inlineStr" || cellType == "str") {
                        inInlineText = true
                        text.setLength(0)
                    }
                }
                XmlPullParser.TEXT -> if (inValue || inInlineText) text.append(parser.text)
                XmlPullParser.END_TAG -> when (parser.name) {
                    "v" -> {
                        inValue = false
                        val raw = text.toString()
                        val resolved = when (cellType) {
                            "s" -> raw.toIntOrNull()?.let { shared.getOrNull(it) }.orEmpty()
                            "b" -> if (raw == "1") "TRUE" else "FALSE"
                            null, "n" -> cleanNumber(raw)
                            else -> raw
                        }
                        val col = colIndex(cellRef)
                        if (col >= 0) currentRow?.put(col, resolved)
                    }
                    "t" -> if (inInlineText) {
                        val col = colIndex(cellRef)
                        if (col >= 0) currentRow?.put(col, text.toString())
                        inInlineText = false
                    }
                    "row" -> {
                        val r = currentRow
                        if (r != null && r.isNotEmpty()) {
                            val maxCol = r.keys.max()
                            rows.add((0..maxCol).map { r[it].orEmpty() })
                        }
                        currentRow = null
                    }
                }
            }
            event = parser.next()
        }
        return rows
    }

    /** "A1" -> 0, "B3" -> 1, "AA2" -> 26 ... */
    private fun colIndex(ref: String?): Int {
        if (ref.isNullOrEmpty()) return -1
        var col = 0
        for (c in ref) {
            if (c in 'A'..'Z') col = col * 26 + (c - 'A' + 1)
            else if (c in 'a'..'z') col = col * 26 + (c - 'a' + 1)
            else break
        }
        return col - 1
    }

    private val DIGITS_REGEX = Regex("\\d+")

    private fun extractPageNumber(s: String): String =
        DIGITS_REGEX.find(s)?.value.orEmpty()

    /** Strip the trailing ".0" that Excel emits for integer-valued numbers. */
    private fun cleanNumber(s: String): String {
        if (s.isEmpty()) return s
        val d = s.toDoubleOrNull() ?: return s
        return if (d == d.toLong().toDouble()) d.toLong().toString() else s
    }
}
