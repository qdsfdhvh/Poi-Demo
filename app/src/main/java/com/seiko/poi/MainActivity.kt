package com.seiko.poi

import android.net.Uri
import android.os.Bundle
import android.util.Log
import android.view.View
import android.widget.Button
import android.widget.TextView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.lifecycle.lifecycleScope
import io.noties.markwon.Markwon
import io.noties.markwon.ext.tables.TablePlugin
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.InputStream

class MainActivity : AppCompatActivity(R.layout.activity_main), View.OnClickListener {

    companion object {
        private const val TAG = "MainActivity"
    }

    private val markwon by lazy(LazyThreadSafetyMode.NONE) {
        Markwon.builder(this)
            .usePlugin(TablePlugin.create(this))
            .build()
    }

    private val filePicker =
        registerForActivityResult(ActivityResultContracts.OpenDocument()) { uri ->
            if (uri == null) {
                return@registerForActivityResult
            }
            readExcelWithUri(uri)
        }

    private lateinit var markdown: TextView

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        markdown = findViewById(R.id.markdown)
        findViewById<Button>(R.id.button1).setOnClickListener(this)
        findViewById<Button>(R.id.button2).setOnClickListener(this)
    }

    override fun onClick(v: View) {
        when (v.id) {
            R.id.button1 -> readExcelWithAssets()
            R.id.button2 -> filePicker.launch(EXCEL_MIME_TYPES)
        }
    }

    private fun readExcelWithAssets() {
        val inputStream = assets.open("Qianji_MuBan_V2.xls")
        readExcel(inputStream)
    }

    private fun readExcelWithUri(uri: Uri) {
        val inputStream = Utils.getUriInputStream(this, uri) ?: return
        readExcel(inputStream)
    }

    private fun readExcel(inputStream: InputStream) {
        lifecycleScope.launchWhenResumed {
            runCatching {
                withContext(Dispatchers.IO) {
                    Utils.readExcelToMarkDown(inputStream)
                }
            }.onFailure {
                showToast(it.message)
                Log.e(TAG, "Read excel error", it)
            }.onSuccess {
                Log.d(TAG, "MD: $it")
                showTable(it)
            }
        }
    }

    private fun showTable(text: String) {
        markwon.setMarkdown(markdown, text)
    }

    private fun showToast(msg: String?) {
        Toast.makeText(this, msg, Toast.LENGTH_SHORT).show()
    }
}