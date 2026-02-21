package com.example.offlinedocxxlsxreader

import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.print.PrintAttributes
import android.print.PrintManager
import android.text.InputType
import android.util.Log
import android.widget.ArrayAdapter
import android.widget.EditText
import android.widget.TextView
import android.widget.Toast
import android.webkit.WebView
import android.webkit.WebViewClient
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AlertDialog
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.isVisible
import androidx.lifecycle.lifecycleScope
import com.example.offlinedocxxlsxreader.databinding.ActivityMainBinding
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.util.Locale

class MainActivity : AppCompatActivity() {
    private lateinit var binding: ActivityMainBinding
    private val docxReader = DocxReader()
    private val xlsxReader = XlsxReader()

    private var currentUri: Uri? = null
    private var currentFileType: String? = null
    private var currentSheetName: String? = null
    private var currentShareText: String? = null
    private var currentXlsxContent: XlsxReader.XlsxContent? = null
    private var currentSearchQuery: String? = null
    private var currentTextExtension: String? = null
    private var pendingCreateMimeType: String = MIME_TEXT_PLAIN
    private var pendingCreateExtension: String = "txt"
    private var pendingCreateDisplayName: String = DEFAULT_NEW_FILE_NAME
    private var searchDialog: AlertDialog? = null
    private var searchInputView: EditText? = null
    private var searchCountView: TextView? = null
    private var textSearchMatches: List<IntRange> = emptyList()
    private var textSearchIndex = -1
    private var printWebView: WebView? = null
    private var printDone = false

    private val viewerPreferences by lazy {
        getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
    }

    private val openDocumentLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri ->
        uri?.let { handleUri(it, persistable = true) }
    }

    private val createDocumentLauncher = registerForActivityResult(
        ActivityResultContracts.CreateDocument("*/*")
    ) { uri ->
        uri?.let { saveTextToUri(it) }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        setSupportActionBar(binding.toolbar)
        setupWebView()
        setupUi()

        val restoredUri = savedInstanceState?.getParcelable<Uri>(STATE_URI)
        val restoredType = savedInstanceState?.getString(STATE_FILE_TYPE)
        currentSheetName = savedInstanceState?.getString(STATE_SHEET_NAME)
        currentShareText = savedInstanceState?.getString(STATE_SHARE_TEXT)
        currentTextExtension = savedInstanceState?.getString(STATE_TEXT_EXTENSION)
        currentUri = restoredUri
        currentFileType = restoredType

        if (restoredUri != null && restoredType != null) {
            when (restoredType) {
                FILE_TYPE_DOCX -> loadDocx(restoredUri)
                FILE_TYPE_XLSX -> loadXlsx(restoredUri)
                FILE_TYPE_TEXT -> loadText(restoredUri)
            }
        } else {
            handleIncomingIntent(intent)
        }
    }

    override fun applyOverrideConfiguration(overrideConfiguration: android.content.res.Configuration?) {
        val prefs = getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        val language = prefs.getString(PREF_LANGUAGE, LANGUAGE_DEFAULT) ?: LANGUAGE_DEFAULT
        val locale = Locale(language)
        Locale.setDefault(locale)
        val config = android.content.res.Configuration(overrideConfiguration ?: resources.configuration)
        config.setLocale(locale)
        super.applyOverrideConfiguration(config)
    }

    override fun onNewIntent(intent: Intent) {
        super.onNewIntent(intent)
        handleIncomingIntent(intent)
    }

    override fun onSaveInstanceState(outState: Bundle) {
        super.onSaveInstanceState(outState)
        outState.putParcelable(STATE_URI, currentUri)
        outState.putString(STATE_FILE_TYPE, currentFileType)
        outState.putString(STATE_SHEET_NAME, currentSheetName)
        outState.putString(STATE_SHARE_TEXT, currentShareText)
        outState.putString(STATE_TEXT_EXTENSION, currentTextExtension)
    }

    override fun onCreateOptionsMenu(menu: android.view.Menu): Boolean {
        menuInflater.inflate(R.menu.menu_main, menu)
        return true
    }

    override fun onOptionsItemSelected(item: android.view.MenuItem): Boolean {
        return when (item.itemId) {
            R.id.action_search -> {
                showSearchDialog()
                true
            }

            R.id.action_new -> {
                showNewDocumentDialog()
                true
            }

            R.id.action_save_as -> {
                saveAs()
                true
            }

            R.id.action_share -> {
                shareCurrentContent()
                true
            }

            R.id.action_print -> {
                printCurrentContent()
                true
            }

            R.id.action_font_size -> {
                showFontSizeDialog()
                true
            }

            R.id.action_language -> {
                showLanguageDialog()
                true
            }

            else -> super.onOptionsItemSelected(item)
        }
    }

    private fun setupUi() {
        binding.buttonOpen.setOnClickListener {
            openDocumentLauncher.launch(
                arrayOf(
                    MIME_DOCX,
                    MIME_XLSX,
                    MIME_TEXT_PLAIN,
                    MIME_JSON,
                    MIME_CSV,
                    MIME_MARKDOWN
                )
            )
        }

        binding.buttonCopy.setOnClickListener {
            val content = getCurrentContentForExport()
            if (content.isNullOrBlank()) return@setOnClickListener
            val clipboard = getSystemService(Context.CLIPBOARD_SERVICE) as ClipboardManager
            clipboard.setPrimaryClip(ClipData.newPlainText(getString(R.string.app_name), content))
            Toast.makeText(this, getString(R.string.toast_copy_done), Toast.LENGTH_SHORT).show()
        }

        binding.sheetSelector.setOnItemClickListener { _, _, position, _ ->
            currentXlsxContent?.sheets?.getOrNull(position)?.let { sheet ->
                showHtml(sheet.html)
                currentShareText = sheet.tsv
                currentSheetName = sheet.name
            }
        }
    }

    private fun setupWebView() {
        binding.webView.settings.apply {
            javaScriptEnabled = false
            allowFileAccess = false
            allowContentAccess = false
            domStorageEnabled = false
            setSupportZoom(true)
            builtInZoomControls = true
            displayZoomControls = false
            useWideViewPort = true
            loadWithOverviewMode = true
            setAllowFileAccessFromFileURLs(false)
            setAllowUniversalAccessFromFileURLs(false)
        }
        binding.webView.setFindListener { _, numberOfMatches, isDoneCounting ->
            if (!isDoneCounting) return@setFindListener
            updateSearchCount(numberOfMatches)
        }
        applyTextZoom(getSavedTextZoom())
    }

    private fun handleIncomingIntent(intent: Intent?) {
        if (intent == null) return
        when (intent.action) {
            Intent.ACTION_VIEW -> {
                intent.data?.let { handleUri(it, persistable = false) }
            }

            Intent.ACTION_SEND -> {
                val uri = intent.getParcelableExtra<Uri>(Intent.EXTRA_STREAM)
                uri?.let { handleUri(it, persistable = false) }
            }
        }
    }

    private fun handleUri(uri: Uri, persistable: Boolean) {
        if (persistable) {
            val flags = Intent.FLAG_GRANT_READ_URI_PERMISSION
            try {
                contentResolver.takePersistableUriPermission(uri, flags)
            } catch (_: SecurityException) {
                // ignore
            }
        }

        val detectedMimeType = contentResolver.getType(uri)
        val mimeType = when {
            detectedMimeType == MIME_DOCX || detectedMimeType == MIME_XLSX || isSupportedTextMime(detectedMimeType) -> detectedMimeType
            else -> guessMimeType(uri)
        }
        when {
            mimeType == MIME_DOCX -> {
                currentUri = uri
                currentFileType = FILE_TYPE_DOCX
                loadDocx(uri)
            }

            mimeType == MIME_XLSX -> {
                currentUri = uri
                currentFileType = FILE_TYPE_XLSX
                loadXlsx(uri)
            }

            isSupportedTextMime(mimeType) -> {
                currentUri = uri
                currentFileType = FILE_TYPE_TEXT
                currentTextExtension = extensionFromUri(uri)
                loadText(uri)
            }

            else -> Toast.makeText(this, getString(R.string.error_wrong_format), Toast.LENGTH_SHORT).show()
        }
    }

    private fun guessMimeType(uri: Uri): String? {
        return when (extensionFromUri(uri)) {
            "docx" -> MIME_DOCX
            "xlsx" -> MIME_XLSX
            "txt", "log" -> MIME_TEXT_PLAIN
            "json" -> MIME_JSON
            "csv" -> MIME_CSV
            "md" -> MIME_MARKDOWN
            else -> null
        }
    }

    private fun extensionFromUri(uri: Uri): String? {
        val name = uri.lastPathSegment ?: return null
        return name.substringAfterLast('.', "").lowercase(Locale.ROOT).ifBlank { null }
    }

    private fun isSupportedTextMime(mimeType: String?): Boolean {
        return mimeType in SUPPORTED_TEXT_MIMES
    }

    private fun loadDocx(uri: Uri) {
        binding.sheetSelectorContainer.isVisible = false
        binding.editorContainer.isVisible = false
        currentXlsxContent = null
        currentSearchQuery = null
        binding.webView.clearMatches()
        binding.webView.isVisible = true
        binding.toolbar.subtitle = getString(R.string.button_open)

        lifecycleScope.launch {
            val result = withContext(Dispatchers.IO) {
                runCatching {
                    contentResolver.openInputStream(uri)?.use { docxReader.parse(it) }
                }
            }
            result.onSuccess { content ->
                if (content == null) {
                    showError()
                    return@onSuccess
                }
                currentShareText = content.plainText
                showHtml(content.html)
            }.onFailure {
                showError()
            }
        }
    }

    private fun loadXlsx(uri: Uri) {
        binding.sheetSelectorContainer.isVisible = true
        binding.editorContainer.isVisible = false
        currentSearchQuery = null
        binding.webView.clearMatches()
        binding.webView.isVisible = true
        binding.toolbar.subtitle = getString(R.string.button_open)

        lifecycleScope.launch {
            val result = withContext(Dispatchers.IO) {
                runCatching {
                    contentResolver.openInputStream(uri)?.use { xlsxReader.parse(it) }
                }
            }
            result.onSuccess { content ->
                if (content == null || content.sheets.isEmpty()) {
                    showError()
                    return@onSuccess
                }
                currentXlsxContent = content
                val adapter = ArrayAdapter(
                    this@MainActivity,
                    android.R.layout.simple_list_item_1,
                    content.sheets.map { it.name }
                )
                binding.sheetSelector.setAdapter(adapter)
                val selectedSheet = currentSheetName?.let { name ->
                    content.sheets.firstOrNull { it.name == name }
                } ?: content.sheets.first()
                currentSheetName = selectedSheet.name
                binding.sheetSelector.setText(selectedSheet.name, false)
                currentShareText = selectedSheet.tsv
                showHtml(selectedSheet.html)
            }.onFailure {
                showError()
            }
        }
    }

    private fun loadText(uri: Uri) {
        binding.sheetSelectorContainer.isVisible = false
        binding.webView.isVisible = false
        binding.editorContainer.isVisible = true
        currentXlsxContent = null
        currentSearchQuery = null
        textSearchMatches = emptyList()
        textSearchIndex = -1
        binding.toolbar.subtitle = getString(R.string.button_open)

        lifecycleScope.launch {
            val result = withContext(Dispatchers.IO) {
                runCatching {
                    contentResolver.openInputStream(uri)?.bufferedReader()?.use { it.readText() }
                }
            }
            result.onSuccess { content ->
                if (content == null) {
                    showError()
                    return@onSuccess
                }
                bindTextContent(content)
            }.onFailure {
                showError()
            }
        }
    }

    private fun bindTextContent(content: String) {
        binding.editor.setText(content)
        currentShareText = content
        val extension = currentTextExtension
        binding.editor.inputType = InputType.TYPE_CLASS_TEXT or InputType.TYPE_TEXT_FLAG_MULTI_LINE
        if (extension == "json" || extension == "log") {
            binding.editor.typeface = android.graphics.Typeface.MONOSPACE
        } else {
            binding.editor.typeface = android.graphics.Typeface.DEFAULT
        }
    }

    private fun showHtml(html: String) {
        binding.editorContainer.isVisible = false
        binding.webView.isVisible = true
        applyTextZoom(getSavedTextZoom())
        binding.webView.loadDataWithBaseURL(
            "about:blank",
            html,
            "text/html",
            "utf-8",
            null
        )
    }

    private fun showSearchDialog() {
        if (currentFileType != FILE_TYPE_DOCX && currentFileType != FILE_TYPE_XLSX && currentFileType != FILE_TYPE_TEXT) {
            Toast.makeText(this, getString(R.string.toast_open_file_first), Toast.LENGTH_SHORT).show()
            return
        }
        if (searchDialog?.isShowing == true) {
            searchInputView?.requestFocus()
            return
        }
        val dialogView = layoutInflater.inflate(R.layout.dialog_search, null)
        val input = dialogView.findViewById<EditText>(R.id.searchInput)
        val countView = dialogView.findViewById<TextView>(R.id.searchCount)
        searchInputView = input
        searchCountView = countView
        val existingQuery = currentSearchQuery.orEmpty()
        input.setText(existingQuery)
        if (existingQuery.isBlank()) {
            updateSearchCount(0)
        } else {
            refreshSearchMatches(existingQuery)
        }

        val dialog = AlertDialog.Builder(this)
            .setTitle(getString(R.string.search_dialog_title))
            .setView(dialogView)
            .setNegativeButton(getString(R.string.search_prev), null)
            .setPositiveButton(getString(R.string.search_next), null)
            .setNeutralButton(getString(R.string.search_close), null)
            .create()

        dialog.setOnShowListener {
            dialog.getButton(AlertDialog.BUTTON_NEGATIVE).setOnClickListener {
                performSearch(forward = false)
            }
            dialog.getButton(AlertDialog.BUTTON_POSITIVE).setOnClickListener {
                performSearch(forward = true)
            }
            dialog.getButton(AlertDialog.BUTTON_NEUTRAL).setOnClickListener {
                clearSearchState()
                dialog.dismiss()
            }
        }
        dialog.setOnDismissListener {
            searchDialog = null
            searchInputView = null
            searchCountView = null
        }
        searchDialog = dialog
        dialog.show()
    }

    private fun performSearch(forward: Boolean) {
        val query = searchInputView?.text?.toString()?.trim().orEmpty()
        if (query.isBlank()) {
            clearSearchState()
            return
        }

        if (currentSearchQuery != query) {
            currentSearchQuery = query
            refreshSearchMatches(query)
        }

        if (currentFileType == FILE_TYPE_TEXT) {
            if (textSearchMatches.isEmpty()) {
                updateSearchCount(0)
                return
            }
            textSearchIndex = if (forward) {
                (textSearchIndex + 1).mod(textSearchMatches.size)
            } else {
                if (textSearchIndex <= 0) textSearchMatches.lastIndex else textSearchIndex - 1
            }
            val range = textSearchMatches[textSearchIndex]
            binding.editor.requestFocus()
            binding.editor.setSelection(range.first, range.last + 1)
            updateSearchCount(textSearchMatches.size)
            return
        }

        binding.webView.findNext(forward)
    }

    private fun refreshSearchMatches(query: String) {
        if (currentFileType == FILE_TYPE_TEXT) {
            val fullText = binding.editor.text?.toString().orEmpty()
            val loweredText = fullText.lowercase(Locale.ROOT)
            val loweredQuery = query.lowercase(Locale.ROOT)
            if (loweredQuery.isBlank()) {
                textSearchMatches = emptyList()
            } else {
                val matches = mutableListOf<IntRange>()
                var index = loweredText.indexOf(loweredQuery)
                while (index >= 0) {
                    matches += index until (index + loweredQuery.length)
                    index = loweredText.indexOf(loweredQuery, index + loweredQuery.length)
                }
                textSearchMatches = matches
            }
            textSearchIndex = -1
            updateSearchCount(textSearchMatches.size)
        } else {
            binding.webView.findAllAsync(query)
        }
    }

    private fun clearSearchState() {
        currentSearchQuery = null
        if (currentFileType == FILE_TYPE_TEXT) {
            textSearchMatches = emptyList()
            textSearchIndex = -1
            val cursor = binding.editor.selectionStart.coerceAtLeast(0)
            binding.editor.setSelection(cursor)
            updateSearchCount(0)
        } else {
            binding.webView.clearMatches()
            updateSearchCount(0)
        }
    }

    private fun updateSearchCount(count: Int) {
        searchCountView?.text = getString(R.string.search_count, count)
    }

    private fun showError() {
        Toast.makeText(this, getString(R.string.error_open_failed), Toast.LENGTH_SHORT).show()
    }

    private fun getCurrentContentForExport(): String? {
        return when (currentFileType) {
            FILE_TYPE_TEXT -> binding.editor.text?.toString()
            else -> currentShareText
        }
    }

    private fun shareCurrentContent() {
        val content = getCurrentContentForExport()
        if (content.isNullOrBlank()) {
            Toast.makeText(this, getString(R.string.toast_share_failed), Toast.LENGTH_SHORT).show()
            return
        }
        val shareIntent = Intent(Intent.ACTION_SEND).apply {
            type = "text/plain"
            putExtra(Intent.EXTRA_TEXT, content)
        }
        try {
            startActivity(Intent.createChooser(shareIntent, getString(R.string.button_share)))
        } catch (_: Exception) {
            Toast.makeText(this, getString(R.string.toast_share_failed), Toast.LENGTH_SHORT).show()
        }
    }

    private fun printCurrentContent() {
        if (isFinishing || isDestroyed) {
            Log.w(TAG, "printCurrentContent skipped: activity finishing=$isFinishing destroyed=$isDestroyed")
            return
        }

        when (currentFileType) {
            FILE_TYPE_DOCX, FILE_TYPE_XLSX -> {
                val jobName = "Offline Reader - Druck"
                val printManager = getSystemService(Context.PRINT_SERVICE) as PrintManager
                val printAdapter = binding.webView.createPrintDocumentAdapter(jobName)
                printManager.print(jobName, printAdapter, PrintAttributes.Builder().build())
            }

            FILE_TYPE_TEXT -> {
                val extension = currentTextExtension?.lowercase(Locale.ROOT)
                val content = when (extension) {
                    "txt", "json" -> readTextFromUriUtf8(currentUri) ?: binding.editor.text?.toString().orEmpty()
                    "log", "csv", "md", null -> binding.editor.text?.toString().orEmpty()
                    else -> {
                        Log.w(TAG, "Print unsupported text extension: $extension")
                        Toast.makeText(this, getString(R.string.toast_print_unavailable), Toast.LENGTH_SHORT).show()
                        return
                    }
                }
                if (content.isBlank()) {
                    Log.w(TAG, "Print unavailable: empty text content for extension=$extension")
                    Toast.makeText(this, getString(R.string.toast_print_unavailable), Toast.LENGTH_SHORT).show()
                    return
                }

                val webView = WebView(this).apply {
                    settings.javaScriptEnabled = false
                }
                printWebView = webView
                printDone = false
                webView.webViewClient = object : WebViewClient() {
                    override fun onPageFinished(view: WebView, url: String?) {
                        if (printDone) return
                        printDone = true
                        val jobName = "Offline Reader - Text"
                        val printManager = getSystemService(Context.PRINT_SERVICE) as PrintManager
                        val printAdapter = view.createPrintDocumentAdapter(jobName)
                        printManager.print(jobName, printAdapter, PrintAttributes.Builder().build())
                        view.webViewClient = WebViewClient()
                        printWebView = null
                    }
                }
                webView.loadDataWithBaseURL(
                    "about:blank",
                    HtmlTemplates.wrap("<pre>${escapeHtml(content)}</pre>"),
                    "text/html",
                    "utf-8",
                    null
                )
            }

            else -> {
                Log.w(TAG, "Print unsupported file type: $currentFileType")
                Toast.makeText(this, getString(R.string.toast_open_file_first), Toast.LENGTH_SHORT).show()
            }
        }
    }

    private fun readTextFromUriUtf8(uri: Uri?): String? {
        if (uri == null) return null
        return runCatching {
            contentResolver.openInputStream(uri)?.bufferedReader(Charsets.UTF_8)?.use { it.readText() }
        }.onFailure {
            Log.w(TAG, "Failed reading text for print from URI=$uri", it)
        }.getOrNull()
    }

    private fun showFontSizeDialog() {
        if (currentFileType != FILE_TYPE_DOCX && currentFileType != FILE_TYPE_XLSX && currentFileType != FILE_TYPE_TEXT) {
            Toast.makeText(
                this,
                getString(R.string.toast_font_size_unavailable),
                Toast.LENGTH_SHORT
            ).show()
            return
        }
        val zoomValues = intArrayOf(80, 100, 120, 140)
        val zoomLabels = arrayOf(
            getString(R.string.font_size_small),
            getString(R.string.font_size_normal),
            getString(R.string.font_size_large),
            getString(R.string.font_size_xlarge)
        )
        val currentZoom = getSavedTextZoom()
        val checkedItem = zoomValues.indexOf(currentZoom).takeIf { it >= 0 } ?: DEFAULT_FONT_INDEX
        AlertDialog.Builder(this)
            .setTitle(getString(R.string.font_size_title))
            .setSingleChoiceItems(zoomLabels, checkedItem) { dialog, which ->
                val selectedZoom = zoomValues[which]
                saveTextZoom(selectedZoom)
                applyTextZoom(selectedZoom)
                dialog.dismiss()
            }
            .setNegativeButton(android.R.string.cancel, null)
            .show()
    }

    private fun showLanguageDialog() {
        val languageOptions = arrayOf(
            getString(R.string.language_option_de),
            getString(R.string.language_option_en),
            getString(R.string.language_option_nl)
        )
        val languageCodes = arrayOf(LANGUAGE_DE, LANGUAGE_EN, LANGUAGE_NL)
        val currentLanguage = viewerPreferences.getString(PREF_LANGUAGE, LANGUAGE_DEFAULT) ?: LANGUAGE_DEFAULT
        val checkedItem = languageCodes.indexOf(currentLanguage).takeIf { it >= 0 } ?: 0
        AlertDialog.Builder(this)
            .setTitle(getString(R.string.language_dialog_title))
            .setSingleChoiceItems(languageOptions, checkedItem) { dialog, which ->
                val selectedLanguage = languageCodes[which]
                if (selectedLanguage != currentLanguage) {
                    viewerPreferences.edit().putString(PREF_LANGUAGE, selectedLanguage).apply()
                    recreate()
                } else {
                    dialog.dismiss()
                }
            }
            .setNegativeButton(android.R.string.cancel, null)
            .show()
    }

    private fun showNewDocumentDialog() {
        val options = listOf(
            NewFileOption(getString(R.string.new_file_txt), "txt", MIME_TEXT_PLAIN),
            NewFileOption(getString(R.string.new_file_json), "json", MIME_JSON),
            NewFileOption(getString(R.string.new_file_log), "log", MIME_TEXT_PLAIN),
            NewFileOption(getString(R.string.new_file_csv), "csv", MIME_CSV),
            NewFileOption(getString(R.string.new_file_md), "md", MIME_MARKDOWN)
        )

        AlertDialog.Builder(this)
            .setTitle(getString(R.string.action_new))
            .setItems(options.map { it.label }.toTypedArray()) { _, which ->
                val selected = options[which]
                currentFileType = FILE_TYPE_TEXT
                currentUri = null
                currentTextExtension = selected.extension
                currentXlsxContent = null
                currentSheetName = null
                currentSearchQuery = null
                currentShareText = ""
                binding.sheetSelectorContainer.isVisible = false
                binding.webView.isVisible = false
                binding.editorContainer.isVisible = true
                bindTextContent("")
                pendingCreateMimeType = selected.mimeType
                pendingCreateExtension = selected.extension
                pendingCreateDisplayName = "$DEFAULT_NEW_FILE_NAME.${selected.extension}"
                Toast.makeText(this, getString(R.string.toast_new_file_created), Toast.LENGTH_SHORT).show()
            }
            .show()
    }

    private fun saveAs() {
        if (currentFileType != FILE_TYPE_TEXT) {
            Toast.makeText(this, getString(R.string.toast_save_only_text), Toast.LENGTH_SHORT).show()
            return
        }

        val extension = currentTextExtension ?: pendingCreateExtension
        val mimeType = when (extension) {
            "json" -> MIME_JSON
            "csv" -> MIME_CSV
            "md" -> MIME_MARKDOWN
            else -> MIME_TEXT_PLAIN
        }

        pendingCreateMimeType = mimeType
        pendingCreateExtension = extension
        pendingCreateDisplayName = "$DEFAULT_NEW_FILE_NAME.$extension"
        createDocumentLauncher.launch(pendingCreateDisplayName)
    }

    private fun saveTextToUri(uri: Uri) {
        val textContent = binding.editor.text?.toString().orEmpty()

        lifecycleScope.launch {
            val result = withContext(Dispatchers.IO) {
                runCatching {
                    contentResolver.openOutputStream(uri, "wt")?.bufferedWriter()?.use { writer ->
                        writer.write(textContent)
                    }
                }
            }

            result.onSuccess {
                handleUri(uri, persistable = false)
                Toast.makeText(this@MainActivity, getString(R.string.toast_save_success), Toast.LENGTH_SHORT).show()
            }.onFailure {
                Toast.makeText(this@MainActivity, getString(R.string.toast_save_failed), Toast.LENGTH_SHORT).show()
            }
        }
    }

    private fun getSavedTextZoom(): Int {
        return viewerPreferences.getInt(PREF_TEXT_ZOOM, DEFAULT_TEXT_ZOOM)
    }

    private fun saveTextZoom(value: Int) {
        viewerPreferences.edit().putInt(PREF_TEXT_ZOOM, value).apply()
    }

    private fun applyTextZoom(value: Int) {
        binding.webView.settings.textZoom = value
        binding.editor.textSize = BASE_EDITOR_TEXT_SIZE_SP * (value / 100f)
    }

    private fun escapeHtml(text: String): String {
        return text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
    }

    companion object {
        private const val TAG = "MainActivity"
        private const val MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        private const val MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        private const val MIME_TEXT_PLAIN = "text/plain"
        private const val MIME_JSON = "application/json"
        private const val MIME_CSV = "text/csv"
        private const val MIME_MARKDOWN = "text/markdown"

        private val SUPPORTED_TEXT_MIMES = setOf(MIME_TEXT_PLAIN, MIME_JSON, MIME_CSV, MIME_MARKDOWN)

        private const val FILE_TYPE_DOCX = "docx"
        private const val FILE_TYPE_XLSX = "xlsx"
        private const val FILE_TYPE_TEXT = "text"

        private const val STATE_URI = "state_uri"
        private const val STATE_FILE_TYPE = "state_file_type"
        private const val STATE_SHEET_NAME = "state_sheet_name"
        private const val STATE_SHARE_TEXT = "state_share_text"
        private const val STATE_TEXT_EXTENSION = "state_text_extension"

        private const val PREFS_NAME = "viewer_preferences"
        private const val PREF_TEXT_ZOOM = "pref_text_zoom"
        private const val PREF_LANGUAGE = "pref_language"
        private const val DEFAULT_TEXT_ZOOM = 100
        private const val DEFAULT_FONT_INDEX = 1
        private const val BASE_EDITOR_TEXT_SIZE_SP = 16f
        private const val DEFAULT_NEW_FILE_NAME = "new_document"

        private const val LANGUAGE_DE = "de"
        private const val LANGUAGE_EN = "en"
        private const val LANGUAGE_NL = "nl"
        private const val LANGUAGE_DEFAULT = LANGUAGE_DE
    }

    private data class NewFileOption(
        val label: String,
        val extension: String,
        val mimeType: String
    )
}
