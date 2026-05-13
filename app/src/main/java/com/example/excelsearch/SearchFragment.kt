package com.example.excelsearch

import android.content.Context
import android.content.Intent
import android.content.SharedPreferences
import android.net.Uri
import android.os.Bundle
import android.text.Editable
import android.text.TextWatcher
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.AdapterView
import android.widget.ArrayAdapter
import android.widget.Button
import android.widget.EditText
import android.widget.ProgressBar
import android.widget.Spinner
import android.widget.TextView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.fragment.app.Fragment
import androidx.lifecycle.lifecycleScope
import androidx.recyclerview.widget.DividerItemDecoration
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext

class SearchFragment : Fragment() {

    private lateinit var prefs: SharedPreferences
    private lateinit var statusText: TextView
    private lateinit var searchInput: EditText
    private lateinit var resultList: RecyclerView
    private lateinit var progressBar: ProgressBar
    private lateinit var progressText: TextView
    private lateinit var sortSpinner: Spinner

    private val adapter = RowAdapter(::onFinishedChanged)

    @Volatile private var entries: List<RowEntry> = emptyList()
    private var currentUri: Uri? = null
    private var finishedIds: MutableSet<Int> = mutableSetOf()
    private var textScale: Float = 1f
    private var sortMode: Int = SORT_DEFAULT

    private data class RowEntry(
        val id: Int,
        val row: ExcelRow,
        val key: PinyinKey,
    )

    private val pickFile = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        uri ?: return@registerForActivityResult
        runCatching {
            requireContext().contentResolver.takePersistableUriPermission(
                uri, Intent.FLAG_GRANT_READ_URI_PERMISSION
            )
        }
        prefs.edit().putString(KEY_LAST_URI, uri.toString()).apply()
        loadFile(uri)
    }

    override fun onCreateView(
        inflater: LayoutInflater,
        container: ViewGroup?,
        savedInstanceState: Bundle?,
    ): View = inflater.inflate(R.layout.fragment_search, container, false)

    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
        super.onViewCreated(view, savedInstanceState)

        prefs = requireContext()
            .getSharedPreferences("excel_search", Context.MODE_PRIVATE)
        textScale = prefs.getFloat(KEY_TEXT_SCALE, 1f).coerceIn(MIN_SCALE, MAX_SCALE)
        sortMode = prefs.getInt(KEY_SORT_MODE, SORT_DEFAULT)
            .coerceIn(SORT_DEFAULT, SORT_BY_LOCATION)

        statusText = view.findViewById(R.id.statusText)
        searchInput = view.findViewById(R.id.searchInput)
        resultList = view.findViewById(R.id.resultList)
        progressBar = view.findViewById(R.id.finishedProgress)
        progressText = view.findViewById(R.id.progressText)
        sortSpinner = view.findViewById(R.id.sortSpinner)

        resultList.layoutManager = LinearLayoutManager(requireContext())
        resultList.adapter = adapter
        resultList.addItemDecoration(
            DividerItemDecoration(requireContext(), DividerItemDecoration.VERTICAL)
        )

        view.findViewById<Button>(R.id.pickFileButton).setOnClickListener {
            pickFile.launch(arrayOf(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel",
                "application/octet-stream",
                "*/*",
            ))
        }
        view.findViewById<Button>(R.id.fontDecreaseButton)
            .setOnClickListener { adjustScale(-SCALE_STEP) }
        view.findViewById<Button>(R.id.fontIncreaseButton)
            .setOnClickListener { adjustScale(+SCALE_STEP) }

        setupSortSpinner()

        searchInput.addTextChangedListener(object : TextWatcher {
            override fun beforeTextChanged(s: CharSequence?, start: Int, count: Int, after: Int) {}
            override fun onTextChanged(s: CharSequence?, start: Int, before: Int, count: Int) {}
            override fun afterTextChanged(s: Editable?) {
                applyFilter(s?.toString().orEmpty())
            }
        })

        updateProgress()
        restoreLastFile()
    }

    private fun setupSortSpinner() {
        val spinnerAdapter = ArrayAdapter.createFromResource(
            requireContext(),
            R.array.sort_options,
            android.R.layout.simple_spinner_item,
        ).also {
            it.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item)
        }
        sortSpinner.adapter = spinnerAdapter
        sortSpinner.setSelection(sortMode, false)
        sortSpinner.onItemSelectedListener = object : AdapterView.OnItemSelectedListener {
            override fun onItemSelected(
                parent: AdapterView<*>?,
                view: View?,
                position: Int,
                id: Long,
            ) {
                if (position == sortMode) return
                sortMode = position
                prefs.edit().putInt(KEY_SORT_MODE, position).apply()
                applyFilter(searchInput.text.toString())
            }
            override fun onNothingSelected(parent: AdapterView<*>?) {}
        }
    }

    private fun restoreLastFile() {
        val saved = prefs.getString(KEY_LAST_URI, null) ?: return
        val uri = runCatching { Uri.parse(saved) }.getOrNull() ?: return
        val hasPermission = requireContext().contentResolver
            .persistedUriPermissions.any {
                it.uri == uri && it.isReadPermission
            }
        if (!hasPermission) {
            prefs.edit().remove(KEY_LAST_URI).apply()
            statusText.text = getString(R.string.reload_failed)
            return
        }
        loadFile(uri)
    }

    private fun loadFile(uri: Uri) {
        statusText.text = getString(R.string.loading)
        viewLifecycleOwner.lifecycleScope.launch {
            val result = withContext(Dispatchers.IO) {
                runCatching {
                    requireContext().contentResolver.openInputStream(uri)?.use { input ->
                        ExcelLoader.load(input)
                    } ?: emptyList()
                }
            }
            result.fold(
                onSuccess = { rows ->
                    currentUri = uri
                    entries = rows.mapIndexed { idx, row ->
                        RowEntry(idx, row, PinyinKey.build(row.name))
                    }
                    finishedIds = loadFinishedIds(uri)
                    statusText.text = getString(R.string.loaded_count, rows.size)
                    applyFilter(searchInput.text.toString())
                    updateProgress()
                },
                onFailure = { err ->
                    err.printStackTrace()
                    statusText.text = getString(R.string.load_failed)
                    Toast.makeText(
                        requireContext(),
                        getString(R.string.load_failed_detail, err.message.orEmpty()),
                        Toast.LENGTH_LONG,
                    ).show()
                },
            )
        }
    }

    private fun applyFilter(rawQuery: String) {
        val q = rawQuery.trim()
        val filtered: List<RowEntry> = if (q.isEmpty()) {
            entries
        } else {
            val normalizedQ = PinyinKey.normalize(q)
            entries.filter { entry ->
                entry.row.name.contains(q, ignoreCase = true) ||
                    (normalizedQ.isNotEmpty() && (
                        entry.key.initials.contains(normalizedQ) ||
                            entry.key.full.contains(normalizedQ)
                        ))
            }
        }
        val sorted = sortEntries(filtered)
        adapter.submit(
            sorted.map { RowAdapter.Item(it.id, it.row) },
            q,
            finishedIds,
            textScale,
        )
    }

    private fun sortEntries(list: List<RowEntry>): List<RowEntry> = when (sortMode) {
        SORT_BY_NAME -> list.sortedWith(
            compareBy(NAME_COMPARATOR) { it.key.full.ifEmpty { it.row.name } }
        )
        SORT_BY_LOCATION -> list.sortedWith(
            compareBy(NAME_COMPARATOR) { it.row.location }
        )
        else -> list
    }

    private fun onFinishedChanged(rowId: Int, finished: Boolean) {
        val changed = if (finished) finishedIds.add(rowId) else finishedIds.remove(rowId)
        if (!changed) return
        persistFinishedIds()
        updateProgress()
    }

    private fun adjustScale(delta: Float) {
        val newScale = (textScale + delta).coerceIn(MIN_SCALE, MAX_SCALE)
        if (newScale == textScale) return
        textScale = newScale
        prefs.edit().putFloat(KEY_TEXT_SCALE, newScale).apply()
        adapter.updateScale(newScale)
    }

    private fun updateProgress() {
        val total = entries.size
        val done = finishedIds.size
        progressBar.max = total.coerceAtLeast(1)
        progressBar.progress = done
        progressText.text = getString(R.string.progress_format, done, total)
    }

    private fun loadFinishedIds(uri: Uri): MutableSet<Int> {
        val raw = prefs.getStringSet(finishedKey(uri), emptySet()) ?: emptySet()
        return raw.mapNotNullTo(mutableSetOf()) { it.toIntOrNull() }
    }

    private fun persistFinishedIds() {
        val uri = currentUri ?: return
        prefs.edit()
            .putStringSet(finishedKey(uri), finishedIds.map(Int::toString).toSet())
            .apply()
    }

    private fun finishedKey(uri: Uri): String =
        "finished_" + Integer.toHexString(uri.toString().hashCode())

    companion object {
        private const val KEY_LAST_URI = "last_uri"
        private const val KEY_TEXT_SCALE = "text_scale"
        private const val KEY_SORT_MODE = "sort_mode"
        private const val MIN_SCALE = 0.7f
        private const val MAX_SCALE = 2.0f
        private const val SCALE_STEP = 0.1f

        private const val SORT_DEFAULT = 0
        private const val SORT_BY_NAME = 1
        private const val SORT_BY_LOCATION = 2

        private val NAME_COMPARATOR =
            java.text.Collator.getInstance(java.util.Locale.CHINA).apply {
                strength = java.text.Collator.PRIMARY
            }
    }
}
