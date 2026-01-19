package com.example.offlinedocxxlsxreader

import android.content.ClipData
import android.content.ClipboardManager
import android.content.Context
import android.content.Intent
import android.graphics.Bitmap
import android.graphics.Color
import android.graphics.pdf.PdfRenderer
import android.net.Uri
import android.os.Bundle
import android.os.ParcelFileDescriptor
import android.view.LayoutInflater
import android.view.Menu
import android.view.MenuItem
import android.view.ViewGroup
import android.webkit.WebResourceRequest
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.ArrayAdapter
import android.widget.EditText
import android.widget.ImageView
import android.widget.TextView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AlertDialog
import androidx.appcompat.app.AppCompatActivity
import androidx.core.view.isVisible
import androidx.lifecycle.lifecycleScope
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
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
    private var pdfRenderer: PdfRenderer? = null
    private var pdfFileDescriptor: ParcelFileDescriptor? = null
    private var searchDialog: AlertDialog? = null
    private var searchInputView: EditText? = null
    private var searchCountView: TextView? = null
    private var currentSearchQuery: String? = null

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
                FILE_TYPE_PDF -> loadPdf(restoredUri)
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

    override fun onCreateOptionsMenu(menu: Menu): Boolean {
        menuInflater.inflate(R.menu.menu_main, menu)
        return true
    }

    override fun onOptionsItemSelected(item: MenuItem): Boolean {
        return when (item.itemId) {
            R.id.action_search -> {
                showSearchDialog()
                true
            }
            R.id.action_share -> {
                shareCurrentContent()
                true
            }
            else -> super.onOptionsItemSelected(item)
        }
    }

    override fun onDestroy() {
        super.onDestroy()
        closePdfRenderer()
    }

    private fun setupUi() {
        binding.buttonOpen.setOnClickListener {
            openDocumentLauncher.launch(arrayOf(MIME_DOCX, MIME_XLSX, MIME_PDF))
        }

        binding.buttonCopy.setOnClickListener {
            if (currentFileType == FILE_TYPE_PDF) {
                Toast.makeText(
                    this,
                    getString(R.string.toast_pdf_text_unavailable),
                    Toast.LENGTH_SHORT
                ).show()
                return@setOnClickListener
            }
            val content = currentShareText
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

        binding.pdfRecyclerView.layoutManager = LinearLayoutManager(this)
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
        binding.webView.setFindListener { _, numberOfMatches, isDoneCounting ->
            if (!isDoneCounting) return@setFindListener
            updateSearchCount(numberOfMatches)
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
            MIME_PDF -> {
                currentUri = uri
                currentFileType = FILE_TYPE_PDF
                loadPdf(uri)
            }
            else -> Toast.makeText(this, getString(R.string.error_wrong_format), Toast.LENGTH_SHORT).show()
        }
    }

    private fun guessMimeType(uri: Uri): String? {
        val name = uri.lastPathSegment ?: return null
        return when {
            name.endsWith(".docx", ignoreCase = true) -> MIME_DOCX
            name.endsWith(".xlsx", ignoreCase = true) -> MIME_XLSX
            name.endsWith(".pdf", ignoreCase = true) -> MIME_PDF
            else -> null
        }
    }

    private fun loadDocx(uri: Uri) {
        binding.sheetSelectorContainer.isVisible = false
        currentXlsxContent = null
        currentSearchQuery = null
        binding.webView.clearMatches()
        closePdfRenderer()
        binding.pdfRecyclerView.isVisible = false
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
        closePdfRenderer()
        currentSearchQuery = null
        binding.webView.clearMatches()
        binding.pdfRecyclerView.isVisible = false
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

    private fun loadPdf(uri: Uri) {
        currentUri = uri
        currentFileType = FILE_TYPE_PDF
        binding.sheetSelectorContainer.isVisible = false
        currentXlsxContent = null
        currentShareText = null
        currentSearchQuery = null
        binding.webView.clearMatches()
        binding.toolbar.subtitle = getString(R.string.button_open)
        binding.webView.isVisible = false
        binding.pdfRecyclerView.isVisible = true
        closePdfRenderer()

        lifecycleScope.launch {
            val result = withContext(Dispatchers.IO) {
                runCatching {
                    val descriptor = contentResolver.openFileDescriptor(uri, "r") ?: return@runCatching null
                    PdfBundle(descriptor, PdfRenderer(descriptor))
                }
            }
            result.onSuccess { bundle ->
                if (bundle == null) {
                    showError()
                    return@onSuccess
                }
                pdfFileDescriptor = bundle.descriptor
                pdfRenderer = bundle.renderer
                val width = resources.displayMetrics.widthPixels
                binding.pdfRecyclerView.adapter = PdfPageAdapter(bundle.renderer, width)
            }.onFailure {
                showError()
            }
        }
    }

    private fun showHtml(html: String) {
        binding.pdfRecyclerView.isVisible = false
        binding.webView.isVisible = true
        binding.webView.loadDataWithBaseURL(
            "about:blank",
            html,
            "text/html",
            "utf-8",
            null
        )
    }

    private fun showSearchDialog() {
        if (currentFileType == FILE_TYPE_PDF) {
            Toast.makeText(this, getString(R.string.toast_pdf_search_unavailable), Toast.LENGTH_SHORT).show()
            return
        }
        if (currentFileType != FILE_TYPE_DOCX && currentFileType != FILE_TYPE_XLSX) {
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
            binding.webView.findAllAsync(existingQuery)
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
                binding.webView.clearMatches()
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
            currentSearchQuery = null
            binding.webView.clearMatches()
            updateSearchCount(0)
            return
        }
        if (currentSearchQuery != query) {
            currentSearchQuery = query
            binding.webView.findAllAsync(query)
        }
        binding.webView.findNext(forward)
    }

    private fun updateSearchCount(count: Int) {
        searchCountView?.text = getString(R.string.search_count, count)
    }

    private fun showError() {
        Toast.makeText(this, getString(R.string.error_open_failed), Toast.LENGTH_SHORT).show()
    }

    private fun shareCurrentContent() {
        val content = when (currentFileType) {
            FILE_TYPE_PDF -> getString(R.string.pdf_share_hint)
            else -> currentShareText
        }
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
        } catch (ex: Exception) {
            Toast.makeText(this, getString(R.string.toast_share_failed), Toast.LENGTH_SHORT).show()
        }
    }

    private fun closePdfRenderer() {
        binding.pdfRecyclerView.adapter = null
        pdfRenderer?.close()
        pdfRenderer = null
        pdfFileDescriptor?.close()
        pdfFileDescriptor = null
    }

    companion object {
        private const val MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        private const val MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        private const val MIME_PDF = "application/pdf"
        private const val FILE_TYPE_DOCX = "docx"
        private const val FILE_TYPE_XLSX = "xlsx"
        private const val FILE_TYPE_PDF = "pdf"
        private const val STATE_URI = "state_uri"
        private const val STATE_FILE_TYPE = "state_file_type"
        private const val STATE_SHEET_NAME = "state_sheet_name"
        private const val STATE_SHARE_TEXT = "state_share_text"
    }

    private data class PdfBundle(
        val descriptor: ParcelFileDescriptor,
        val renderer: PdfRenderer
    )

    private class PdfPageAdapter(
        private val renderer: PdfRenderer,
        private val targetWidth: Int
    ) : RecyclerView.Adapter<PdfPageAdapter.PageViewHolder>() {

        override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): PageViewHolder {
            val view = LayoutInflater.from(parent.context)
                .inflate(R.layout.item_pdf_page, parent, false)
            return PageViewHolder(view)
        }

        override fun onBindViewHolder(holder: PageViewHolder, position: Int) {
            val page = renderer.openPage(position)
            val width = if (targetWidth > 0) targetWidth else page.width
            val scale = width.toFloat() / page.width.toFloat()
            val height = (page.height * scale).toInt().coerceAtLeast(1)
            val bitmap = Bitmap.createBitmap(width, height, Bitmap.Config.ARGB_8888)
            bitmap.eraseColor(Color.WHITE)
            page.render(bitmap, null, null, PdfRenderer.Page.RENDER_MODE_FOR_DISPLAY)
            page.close()
            holder.imageView.setImageBitmap(bitmap)
        }

        override fun onViewRecycled(holder: PageViewHolder) {
            holder.imageView.setImageDrawable(null)
        }

        override fun getItemCount(): Int = renderer.pageCount

        class PageViewHolder(itemView: android.view.View) : RecyclerView.ViewHolder(itemView) {
            val imageView: ImageView = itemView.findViewById(R.id.pdfPageImage)
        }
    }
}
