package com.example.excelsearch

import net.sourceforge.pinyin4j.PinyinHelper
import net.sourceforge.pinyin4j.format.HanyuPinyinCaseType
import net.sourceforge.pinyin4j.format.HanyuPinyinOutputFormat
import net.sourceforge.pinyin4j.format.HanyuPinyinToneType

/**
 * Pinyin-aware match key for a name.
 *
 * `full` is the concatenated lowercase pinyin (e.g. "李大" -> "lida"),
 * `initials` is the leading letters only (e.g. "李大" -> "ld").
 * Non-Chinese letters are kept as-is so mixed input still works.
 *
 * Both forms are stored already-normalized: ng -> n and l -> n,
 * so front/back nasal and l/n confusions match by substring directly.
 */
class PinyinKey(val full: String, val initials: String) {

    companion object {
        private val format = HanyuPinyinOutputFormat().apply {
            caseType = HanyuPinyinCaseType.LOWERCASE
            toneType = HanyuPinyinToneType.WITHOUT_TONE
        }

        fun build(name: String): PinyinKey {
            val fullBuf = StringBuilder()
            val initBuf = StringBuilder()
            for (c in name) {
                val py = pinyinOf(c)
                if (py != null) {
                    fullBuf.append(py)
                    initBuf.append(py[0])
                } else if (c.isLetterOrDigit()) {
                    fullBuf.append(c.lowercaseChar())
                    initBuf.append(c.lowercaseChar())
                }
            }
            return PinyinKey(
                full = normalize(fullBuf.toString()),
                initials = normalize(initBuf.toString()),
            )
        }

        fun normalize(input: String): String {
            if (input.isEmpty()) return input
            val lower = input.lowercase()
            val sb = StringBuilder(lower.length)
            var i = 0
            while (i < lower.length) {
                val c = lower[i]
                if (!c.isLetter()) { i++; continue }
                if (c == 'n' && i + 1 < lower.length && lower[i + 1] == 'g') {
                    sb.append('n')
                    i += 2
                    continue
                }
                sb.append(if (c == 'l') 'n' else c)
                i++
            }
            return sb.toString()
        }

        private fun pinyinOf(c: Char): String? {
            if (c.code < 0x4E00 || c.code > 0x9FFF) return null
            val arr = runCatching { PinyinHelper.toHanyuPinyinStringArray(c, format) }
                .getOrNull()
            return arr?.firstOrNull()?.takeIf { it.isNotEmpty() }
        }
    }
}
