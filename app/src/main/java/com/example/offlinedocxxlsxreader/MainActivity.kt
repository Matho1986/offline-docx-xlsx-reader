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
import android.print.PrintAttributes
import android.print.PrintManager
import android.view.GestureDetector
import android.view.LayoutInflater
import android.view.Menu
import android.view.MenuItem
import android.view.MotionEvent
import android.view.ScaleGestureDetector
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
    private var pdfRenderer: PdfRenderer? = null
    private var pdfFileDescriptor: ParcelFileDescriptor? = null
    private var searchDialog: AlertDialog? = null
    private var searchInputView: EditText? = null
    private var searchCountView: TextView? = null
    private var currentSearchQuery: String? = null
    private var pdfScaleFactor = PDF_SCALE_DEFAULT
    private var pdfTranslationX = 0f
    private var pdfTranslationY = 0f
    private var pdfLastTouchX = 0f
    private var pdfLastTouchY = 0f
    private var isPdfDragging = false
    private lateinit var pdfScaleDetector: ScaleGestureDetector
    private lateinit var pdfGestureDetector: GestureDetector
    private val viewerPreferences by lazy {
        getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
    }

    private val openDocumentLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri ->
        uri?.let { handleUri(it, persistable = true) }
    }

    override fun attachBaseContext(newBase: Context) {
        val prefs = newBase.getSharedPreferences(PREFS_NAME, Context.MODE_PRIVATE)
        val language = prefs.getString(PREF_LANGUAGE, LANGUAGE_DEFAULT) ?: LANGUAGE_DEFAULT
        super.attachBaseContext(updateLocale(newBase, language))
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        setSupportActionBar(binding.toolbar)
        setupWebView()
        setupUi()
        setupPdfGestures()

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
        binding.pdfRecyclerView.setOnTouchListener { _, event ->
            if (currentFileType != FILE_TYPE_PDF) return@setOnTouchListener false
            pdfScaleDetector.onTouchEvent(event)
            val gestureHandled = pdfGestureDetector.onTouchEvent(event)

            if (pdfScaleDetector.isInProgress) {
                isPdfDragging = false
                return@setOnTouchListener true
            }

            when (event.actionMasked) {
                MotionEvent.ACTION_DOWN -> {
                    pdfLastTouchX = event.x
                    pdfLastTouchY = event.y
                    isPdfDragging = false
                }
                MotionEvent.ACTION_POINTER_DOWN -> {
                    isPdfDragging = false
                }
                MotionEvent.ACTION_POINTER_UP -> {
                    val remainingIndex = if (event.actionIndex == 0) 1 else 0
                    if (remainingIndex < event.pointerCount) {
                        pdfLastTouchX = event.getX(remainingIndex)
                        pdfLastTouchY = event.getY(remainingIndex)
                    }
                    isPdfDragging = false
                }
                MotionEvent.ACTION_MOVE -> {
                    if (event.pointerCount == 1 && pdfScaleFactor > PDF_SCALE_DEFAULT) {
                        val deltaX = event.x - pdfLastTouchX
                        val deltaY = event.y - pdfLastTouchY
                        pdfTranslationX += deltaX
                        pdfTranslationY += deltaY
                        clampPdfTranslation()
                        applyPdfScale()
                        pdfLastTouchX = event.x
                        pdfLastTouchY = event.y
                        isPdfDragging = true
                        return@setOnTouchListener true
                    }
                }
                MotionEvent.ACTION_UP,
                MotionEvent.ACTION_CANCEL -> {
                    isPdfDragging = false
                }
            }

            pdfScaleDetector.isInProgress || gestureHandled || isPdfDragging
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
        binding.webView.webViewClient = object : WebViewClient() {
            override fun shouldOverrideUrlLoading(view: WebView?, request: WebResourceRequest?): Boolean {
                return true
            }
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
        resetPdfScale()
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
                applyPdfScale()
            }.onFailure {
                showError()
            }
        }
    }

    private fun setupPdfGestures() {
        pdfScaleDetector = ScaleGestureDetector(this, object : ScaleGestureDetector.SimpleOnScaleGestureListener() {
            override fun onScale(detector: ScaleGestureDetector): Boolean {
                pdfScaleFactor = (pdfScaleFactor * detector.scaleFactor)
                    .coerceIn(PDF_SCALE_MIN, PDF_SCALE_MAX)
                applyPdfScale()
                return true
            }
        })
        pdfGestureDetector = GestureDetector(this, object : GestureDetector.SimpleOnGestureListener() {
            override fun onDoubleTap(e: MotionEvent): Boolean {
                if (currentFileType != FILE_TYPE_PDF) return false
                resetPdfScale()
                applyPdfScale()
                return true
            }
        })
    }

    private fun applyPdfScale() {
        binding.pdfRecyclerView.pivotX = binding.pdfRecyclerView.width / 2f
        binding.pdfRecyclerView.pivotY = 0f
        if (pdfScaleFactor <= PDF_SCALE_DEFAULT) {
            pdfTranslationX = 0f
            pdfTranslationY = 0f
        } else {
            clampPdfTranslation()
        }
        binding.pdfRecyclerView.scaleX = pdfScaleFactor
        binding.pdfRecyclerView.scaleY = pdfScaleFactor
        binding.pdfRecyclerView.translationX = pdfTranslationX
        binding.pdfRecyclerView.translationY = pdfTranslationY
    }

    private fun resetPdfScale() {
        pdfScaleFactor = PDF_SCALE_DEFAULT
        pdfTranslationX = 0f
        pdfTranslationY = 0f
    }

    private fun clampPdfTranslation() {
        val width = binding.pdfRecyclerView.width.toFloat()
        val height = binding.pdfRecyclerView.height.toFloat()
        if (width <= 0f || height <= 0f) {
            pdfTranslationX = 0f
            pdfTranslationY = 0f
            return
        }
        val maxTranslationX = ((width * pdfScaleFactor) - width) / 2f
        val maxTranslationY = (height * pdfScaleFactor) - height
        pdfTranslationX = pdfTranslationX.coerceIn(-maxTranslationX, maxTranslationX)
        pdfTranslationY = pdfTranslationY.coerceIn(-maxTranslationY, 0f)
    }

    private fun showHtml(html: String) {
        binding.pdfRecyclerView.isVisible = false
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

    private fun printCurrentContent() {
        when (currentFileType) {
            FILE_TYPE_DOCX, FILE_TYPE_XLSX -> {
                val jobName = "Offline Reader - Druck"
                val printManager = getSystemService(Context.PRINT_SERVICE) as PrintManager
                val printAdapter = binding.webView.createPrintDocumentAdapter(jobName)
                printManager.print(jobName, printAdapter, PrintAttributes.Builder().build())
            }
            FILE_TYPE_PDF -> {
                Toast.makeText(this, getString(R.string.toast_pdf_print_unavailable), Toast.LENGTH_SHORT).show()
            }
            else -> {
                Toast.makeText(this, getString(R.string.toast_open_file_first), Toast.LENGTH_SHORT).show()
            }
        }
    }

    private fun showFontSizeDialog() {
        if (currentFileType != FILE_TYPE_DOCX && currentFileType != FILE_TYPE_XLSX) {
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

    private fun updateLocale(context: Context, language: String): Context {
        val locale = Locale(language)
        Locale.setDefault(locale)
        val config = android.content.res.Configuration(context.resources.configuration)
        config.setLocale(locale)
        return context.createConfigurationContext(config)
    }

    private fun getSavedTextZoom(): Int {
        return viewerPreferences.getInt(PREF_TEXT_ZOOM, DEFAULT_TEXT_ZOOM)
    }

    private fun saveTextZoom(value: Int) {
        viewerPreferences.edit().putInt(PREF_TEXT_ZOOM, value).apply()
    }

    private fun applyTextZoom(value: Int) {
        binding.webView.settings.textZoom = value
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
        private const val PREFS_NAME = "viewer_preferences"
        private const val PREF_TEXT_ZOOM = "pref_text_zoom"
        private const val PREF_LANGUAGE = "pref_language"
        private const val DEFAULT_TEXT_ZOOM = 100
        private const val DEFAULT_FONT_INDEX = 1
        private const val PDF_SCALE_DEFAULT = 1.0f
        private const val PDF_SCALE_MIN = 1.0f
        private const val PDF_SCALE_MAX = 3.0f
        private const val LANGUAGE_DE = "de"
        private const val LANGUAGE_EN = "en"
        private const val LANGUAGE_NL = "nl"
        private const val LANGUAGE_DEFAULT = LANGUAGE_DE
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
