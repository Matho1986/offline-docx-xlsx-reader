package com.example.offlinedocxxlsxreader

import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.webkit.WebResourceRequest
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.ArrayAdapter
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.isVisible
import androidx.lifecycle.lifecycleScope
import com.example.offlinedocxxlsxreader.databinding.ActivityMainBinding
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext

class MainActivity : AppCompatActivity() {
    private lateinit var binding: ActivityMainBinding
    private val docxReader = DocxReader()
    private val xlsxReader = XlsxReader()

    private var currentUri: Uri? = null
    private var currentFileType: String? = null
    private var currentSheetName: String? = null
    private var currentShareText: String? = null
    private var currentXlsxContent: XlsxReader.XlsxContent? = null

    private val openDocumentLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri ->
        uri?.let { handleUri(it, persistable = true) }
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
        currentUri = restoredUri
        currentFileType = restoredType

        if (restoredUri != null && restoredType != null) {
            when (restoredType) {
                FILE_TYPE_DOCX -> loadDocx(restoredUri)
                FILE_TYPE_XLSX -> loadXlsx(restoredUri)
            }
        } else {
            handleIncomingIntent(intent)
        }
    }

    override fun onNewIntent(intent: Intent?) {
        super.onNewIntent(intent)
        handleIncomingIntent(intent)
    }

    override fun onSaveInstanceState(outState: Bundle) {
        super.onSaveInstanceState(outState)
        outState.putParcelable(STATE_URI, currentUri)
        outState.putString(STATE_FILE_TYPE, currentFileType)
        outState.putString(STATE_SHEET_NAME, currentSheetName)
        outState.putString(STATE_SHARE_TEXT, currentShareText)
    }

    private fun setupUi() {
        binding.buttonOpen.setOnClickListener {
            openDocumentLauncher.launch(arrayOf(MIME_DOCX, MIME_XLSX))
        }

        binding.buttonCopy.setOnClickListener {
            val content = currentShareText
            if (content.isNullOrBlank()) return@setOnClickListener
            val clipboard = getSystemService(Context.CLIPBOARD_SERVICE) as ClipboardManager
            clipboard.setPrimaryClip(ClipData.newPlainText(getString(R.string.app_name), content))
            Toast.makeText(this, getString(R.string.toast_copy_done), Toast.LENGTH_SHORT).show()
        }

        binding.buttonShare.setOnClickListener {
            val content = currentShareText
            if (content.isNullOrBlank()) return@setOnClickListener
            val shareIntent = Intent(Intent.ACTION_SEND).apply {
                type = "text/plain"
                putExtra(Intent.EXTRA_TEXT, content)
            }
            try {
                startActivity(Intent.createChooser(shareIntent, getString(R.string.button_share)))
            } catch (ex: Exception) {
                Toast.makeText(this, getString(R.string.toast_share_failed), Toast.LENGTH_SHORT).show()
            }
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
            setAllowFileAccessFromFileURLs(false)
            setAllowUniversalAccessFromFileURLs(false)
        }
        binding.webView.webViewClient = object : WebViewClient() {
            override fun shouldOverrideUrlLoading(view: WebView?, request: WebResourceRequest?): Boolean {
                return true
            }
        }
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
                // Ignore if permission cannot be persisted.
            }
        }

        val mimeType = contentResolver.getType(uri) ?: guessMimeType(uri)
        when (mimeType) {
            MIME_DOCX -> {
                currentUri = uri
                currentFileType = FILE_TYPE_DOCX
                loadDocx(uri)
            }
            MIME_XLSX -> {
                currentUri = uri
                currentFileType = FILE_TYPE_XLSX
                loadXlsx(uri)
            }
            else -> Toast.makeText(this, getString(R.string.error_wrong_format), Toast.LENGTH_SHORT).show()
        }
    }

    private fun guessMimeType(uri: Uri): String? {
        val name = uri.lastPathSegment ?: return null
        return when {
            name.endsWith(".docx", ignoreCase = true) -> MIME_DOCX
            name.endsWith(".xlsx", ignoreCase = true) -> MIME_XLSX
            else -> null
        }
    }

    private fun loadDocx(uri: Uri) {
        binding.sheetSelectorContainer.isVisible = false
        currentXlsxContent = null
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

    private fun showHtml(html: String) {
        binding.webView.loadDataWithBaseURL(
            "about:blank",
            html,
            "text/html",
            "utf-8",
            null
        )
    }

    private fun showError() {
        Toast.makeText(this, getString(R.string.error_open_failed), Toast.LENGTH_SHORT).show()
    }

    companion object {
        private const val MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        private const val MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        private const val FILE_TYPE_DOCX = "docx"
        private const val FILE_TYPE_XLSX = "xlsx"
        private const val STATE_URI = "state_uri"
        private const val STATE_FILE_TYPE = "state_file_type"
        private const val STATE_SHEET_NAME = "state_sheet_name"
        private const val STATE_SHARE_TEXT = "state_share_text"
    }
}
