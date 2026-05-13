package com.example.excelsearch

import net.sourceforge.pinyin4j.PinyinHelper
import net.sourceforge.pinyin4j.format.HanyuPinyinCaseType
import net.sourceforge.pinyin4j.format.HanyuPinyinOutputFormat
import net.sourceforge.pinyin4j.format.HanyuPinyinToneType

/**
 * Pinyin-aware match key supporting polyphone characters.
 *
 * - `full`: normalized concatenation of each char's first reading. Used for sort
 *   and as a tie-breaker; not relied on for matching when polyphones exist.
 * - `initialSlots[i]`: the set of possible normalized initial letters at char
 *   position `i`. Multi-reading chars contribute multiple options.
 * - `fullVariants`: every normalized full-pinyin concatenation that the readings
 *   admit, capped so pathological names don't explode.
 *
 * Both representations are pre-normalized: `ng -> n` and `l -> n`, so front/back
 * nasal and l/n confusions still match by substring.
 */
class PinyinKey(
    val full: String,
    private val initialSlots: List<Set<Char>>,
    private val fullVariants: List<String>,
) {

    fun matchesInitials(normalizedQuery: String): Boolean {
        if (normalizedQuery.isEmpty()) return false
        if (initialSlots.size < normalizedQuery.length) return false
        val last = initialSlots.size - normalizedQuery.length
        for (start in 0..last) {
            var ok = true
            for (j in normalizedQuery.indices) {
                if (normalizedQuery[j] !in initialSlots[start + j]) {
                    ok = false
                    break
                }
            }
            if (ok) return true
        }
        return false
    }

    fun matchesFull(normalizedQuery: String): Boolean {
        if (normalizedQuery.isEmpty()) return false
        return fullVariants.any { it.contains(normalizedQuery) }
    }

    companion object {
        private const val VARIANTS_CAP = 16

        private val format = HanyuPinyinOutputFormat().apply {
            caseType = HanyuPinyinCaseType.LOWERCASE
            toneType = HanyuPinyinToneType.WITHOUT_TONE
        }

        fun build(name: String): PinyinKey {
            val firstFull = StringBuilder()
            val initialSlots = mutableListOf<Set<Char>>()
            val syllableSlots = mutableListOf<List<String>>()

            for (c in name) {
                val readings = readingsOf(c)
                if (readings != null) {
                    val normalizedSyllables = readings.map { normalize(it) }
                        .filter { it.isNotEmpty() }
                        .distinct()
                    if (normalizedSyllables.isEmpty()) continue
                    firstFull.append(normalizedSyllables[0])
                    initialSlots.add(normalizedSyllables.map { it[0] }.toSet())
                    syllableSlots.add(normalizedSyllables)
                } else if (c.isLetterOrDigit()) {
                    val lower = c.lowercaseChar()
                    val n = normalize(lower.toString())
                    if (n.isEmpty()) continue
                    firstFull.append(n)
                    initialSlots.add(setOf(n[0]))
                    syllableSlots.add(listOf(n))
                }
            }

            return PinyinKey(
                full = firstFull.toString(),
                initialSlots = initialSlots,
                fullVariants = enumerateVariants(syllableSlots),
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

        private fun readingsOf(c: Char): List<String>? {
            if (c.code < 0x4E00 || c.code > 0x9FFF) return null
            val arr = runCatching { PinyinHelper.toHanyuPinyinStringArray(c, format) }
                .getOrNull() ?: return null
            val list = arr.filter { it.isNotEmpty() }.distinct()
            return list.ifEmpty { null }
        }

        private fun enumerateVariants(slots: List<List<String>>): List<String> {
            if (slots.isEmpty()) return listOf("")
            var results: List<String> = listOf("")
            for (slot in slots) {
                val expanded = ArrayList<String>(results.size * slot.size)
                outer@ for (prefix in results) {
                    for (syllable in slot) {
                        expanded.add(prefix + syllable)
                        if (expanded.size >= VARIANTS_CAP) break@outer
                    }
                }
                results = expanded
                if (results.size >= VARIANTS_CAP) break
            }
            return results
        }
    }
}
