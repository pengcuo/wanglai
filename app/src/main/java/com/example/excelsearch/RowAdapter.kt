package com.example.excelsearch

import android.text.Spannable
import android.text.SpannableString
import android.text.style.BackgroundColorSpan
import android.util.TypedValue
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.CheckBox
import android.widget.TextView
import androidx.core.content.ContextCompat
import androidx.recyclerview.widget.RecyclerView

class RowAdapter(
    private val onFinishedChanged: (rowId: Int, finished: Boolean) -> Unit,
) : RecyclerView.Adapter<RowAdapter.VH>() {

    data class Item(val rowId: Int, val row: ExcelRow)

    private var items: List<Item> = emptyList()
    private var finishedIds: Set<Int> = emptySet()
    private var query: String = ""
    private var textScale: Float = 1f

    fun submit(newItems: List<Item>, q: String, finished: Set<Int>, scale: Float) {
        items = newItems
        query = q
        finishedIds = finished
        textScale = scale
        notifyDataSetChanged()
    }

    fun updateFinished(finished: Set<Int>) {
        finishedIds = finished
        notifyDataSetChanged()
    }

    fun updateScale(scale: Float) {
        textScale = scale
        notifyDataSetChanged()
    }

    class VH(view: View) : RecyclerView.ViewHolder(view) {
        val name: TextView = view.findViewById(R.id.name)
        val amount: TextView = view.findViewById(R.id.amount)
        val location: TextView = view.findViewById(R.id.location)
        val page: TextView = view.findViewById(R.id.page)
        val check: CheckBox = view.findViewById(R.id.finishedCheck)
    }

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): VH {
        val v = LayoutInflater.from(parent.context).inflate(R.layout.item_row, parent, false)
        return VH(v)
    }

    override fun onBindViewHolder(holder: VH, position: Int) {
        val ctx = holder.itemView.context
        val item = items[position]
        val row = item.row
        val highlightBg = ContextCompat.getColor(ctx, R.color.highlight_match)
        holder.name.text = highlight(row.name, query, highlightBg)
        holder.amount.text = ctx.getString(R.string.amount_format, row.amount.ifEmpty { "—" })
        holder.location.text = highlightInTemplate(
            ctx.getString(R.string.location_format, row.location.ifEmpty { "—" }),
            row.location,
            query,
            highlightBg,
        )
        holder.page.text = ctx.getString(R.string.page_format, row.page.ifEmpty { "—" })

        applyScale(holder.name, 17f)
        applyScale(holder.amount, 13f)
        applyScale(holder.location, 13f)
        applyScale(holder.page, 13f)
        applyScale(holder.check, 12f)

        holder.check.setOnCheckedChangeListener(null)
        holder.check.isChecked = item.rowId in finishedIds
        holder.check.setOnCheckedChangeListener { _, isChecked ->
            onFinishedChanged(item.rowId, isChecked)
        }
    }

    override fun getItemCount(): Int = items.size

    private fun applyScale(view: TextView, baseSp: Float) {
        view.setTextSize(TypedValue.COMPLEX_UNIT_SP, baseSp * textScale)
    }

    private fun highlight(text: String, q: String, bgColor: Int): CharSequence {
        if (q.isEmpty()) return text
        val idx = text.indexOf(q, ignoreCase = true)
        if (idx < 0) return text
        val span = SpannableString(text)
        span.setSpan(
            BackgroundColorSpan(bgColor),
            idx, idx + q.length,
            Spannable.SPAN_EXCLUSIVE_EXCLUSIVE,
        )
        return span
    }

    private fun highlightInTemplate(
        full: String,
        substring: String,
        q: String,
        bgColor: Int,
    ): CharSequence {
        if (q.isEmpty() || substring.isEmpty()) return full
        val matchIdx = substring.indexOf(q, ignoreCase = true)
        if (matchIdx < 0) return full
        val anchor = full.indexOf(substring)
        if (anchor < 0) return full
        val start = anchor + matchIdx
        val end = start + q.length
        val span = SpannableString(full)
        span.setSpan(
            BackgroundColorSpan(bgColor),
            start, end,
            Spannable.SPAN_EXCLUSIVE_EXCLUSIVE,
        )
        return span
    }
}
