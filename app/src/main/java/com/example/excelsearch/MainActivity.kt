package com.example.excelsearch

import android.app.Activity
import android.content.Context
import android.content.Intent
import android.content.SharedPreferences
import android.net.Uri
import android.os.Bundle
import android.speech.RecognizerIntent
import android.text.Editable
import android.text.TextWatcher
import android.widget.Button
import android.widget.EditText
import android.widget.ImageButton
import android.widget.ProgressBar
import android.widget.TextView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import androidx.recyclerview.widget.DividerItemDecoration
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.util.Locale

class MainActivity : AppCompatActivity() {

    private lateinit var prefs: SharedPreferences
    private lateinit var statusText: TextView
    private lateinit var searchInput: EditText
    private lateinit var resultList: RecyclerView
    private lateinit var progressBar: ProgressBar
    private lateinit var progressText: TextView

    private val adapter = RowAdapter(::onFinishedChanged)

    @Volatile private var entries: List<RowEntry> = emptyList()
    private var currentUri: Uri? = null
    private var finishedIds: MutableSet<Int> = mutableSetOf()
    private var textScale: Float = 1f

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
            contentResolver.takePersistableUriPermission(
                uri, Intent.FLAG_GRANT_READ_URI_PERMISSION
            )
        }
        prefs.edit().putString(KEY_LAST_URI, uri.toString()).apply()
        loadFile(uri)
    }

    private val voiceSearch = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        if (result.resultCode == Activity.RESULT_OK) {
            val matches = result.data
                ?.getStringArrayListExtra(RecognizerIntent.EXTRA_RESULTS)
            matches?.firstOrNull()?.let { spoken ->
                searchInput.setText(spoken)
                searchInput.setSelection(spoken.length)
            }
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        prefs = getSharedPreferences("excel_search", Context.MODE_PRIVATE)
        textScale = prefs.getFloat(KEY_TEXT_SCALE, 1f).coerceIn(MIN_SCALE, MAX_SCALE)

        statusText = findViewById(R.id.statusText)
        searchInput = findViewById(R.id.searchInput)
        resultList = findViewById(R.id.resultList)
        progressBar = findViewById(R.id.finishedProgress)
        progressText = findViewById(R.id.progressText)

        resultList.layoutManager = LinearLayoutManager(this)
        resultList.adapter = adapter
        resultList.addItemDecoration(
            DividerItemDecoration(this, DividerItemDecoration.VERTICAL)
        )

        findViewById<Button>(R.id.pickFileButton).setOnClickListener {
            pickFile.launch(arrayOf(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "application/vnd.ms-excel",
                "application/octet-stream",
                "*/*",
            ))
        }
        findViewById<Button>(R.id.fontDecreaseButton).setOnClickListener { adjustScale(-SCALE_STEP) }
        findViewById<Button>(R.id.fontIncreaseButton).setOnClickListener { adjustScale(+SCALE_STEP) }

        findViewById<ImageButton>(R.id.voiceButton).setOnClickListener {
            startVoiceRecognition()
        }

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

    private fun restoreLastFile() {
        val saved = prefs.getString(KEY_LAST_URI, null) ?: return
        val uri = runCatching { Uri.parse(saved) }.getOrNull() ?: return
        val hasPermission = contentResolver.persistedUriPermissions.any {
            it.uri == uri && it.isReadPermission
        }
        if (!hasPermission) {
            prefs.edit().remove(KEY_LAST_URI).apply()
            statusText.text = getString(R.string.reload_failed)
            return
        }
        loadFile(uri)
    }

    private fun startVoiceRecognition() {
        try {
            val intent = Intent(RecognizerIntent.ACTION_RECOGNIZE_SPEECH).apply {
                putExtra(
                    RecognizerIntent.EXTRA_LANGUAGE_MODEL,
                    RecognizerIntent.LANGUAGE_MODEL_FREE_FORM,
                )
                putExtra(
                    RecognizerIntent.EXTRA_LANGUAGE,
                    Locale.SIMPLIFIED_CHINESE.toLanguageTag(),
                )
                putExtra(
                    RecognizerIntent.EXTRA_PROMPT,
                    getString(R.string.voice_prompt),
                )
            }
            voiceSearch.launch(intent)
        } catch (_: Exception) {
            Toast.makeText(this, R.string.voice_not_supported, Toast.LENGTH_SHORT).show()
        }
    }

    private fun loadFile(uri: Uri) {
        statusText.text = getString(R.string.loading)
        lifecycleScope.launch {
            val result = withContext(Dispatchers.IO) {
                runCatching {
                    contentResolver.openInputStream(uri)?.use { input ->
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
                        this@MainActivity,
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
        adapter.submit(
            filtered.map { RowAdapter.Item(it.id, it.row) },
            q,
            finishedIds,
            textScale,
        )
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
        private const val MIN_SCALE = 0.7f
        private const val MAX_SCALE = 2.0f
        private const val SCALE_STEP = 0.1f
    }
}
