package com.seiko.poi

import android.content.Context
import android.net.Uri
import android.os.Build
import org.apache.poi.hssf.usermodel.HSSFDateUtil
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.formula.eval.ErrorEval
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.InputStream
import java.text.DateFormat
import java.text.SimpleDateFormat
import java.util.*

object Utils {

    @Throws(RuntimeException::class)
    fun getUriInputStream(context: Context, uri: Uri): InputStream? {
        return context.contentResolver.openAssetFileDescriptor(uri, "r")?.createInputStream()
    }

    /**
     * 解析excel表格转成markdown
     * PS: 根据情况调整修改
     */
    @Throws(RuntimeException::class)
    fun readExcelToMarkDown(inputStream: InputStream): String {
        try {
            val workbook = if (isUnsupportedDevice) {
                HSSFWorkbook(inputStream)
            } else {
                XSSFWorkbook(inputStream)
            }
            val sheet = workbook.getSheetAt(0)
            val formulaEvaluator = workbook.creationHelper.createFormulaEvaluator()

            val sb = StringBuilder()

            var row: Row
            var cell: Cell?
            val rowsCount = sheet.physicalNumberOfRows
            var cellsCount = 0

            for (i in 0 until rowsCount) {
                row = sheet.getRow(i)

                // 只取第一行的列数(标题)
                if (i == 0) {
                    cellsCount = row.physicalNumberOfCells
                }

                sb.append("|")
                for (j in 0 until cellsCount) {
                    cell = row.getCell(j)
                    if (cell == null) {
                        sb.append("|")
                        continue
                    }
                    val value = getCellAsString(cell, formulaEvaluator)
                    sb.append(value).append("|")
                }
                sb.append("\n")

                // |:----:|:----:|
                if (i == 0) {
                    sb.append("|")
                    for (i2 in 0 until cellsCount) {
                        sb.append(":----:|")
                    }
                    sb.append("\n")
                }
            }

            return sb.toString()
        } catch (e: Exception) {
            throw RuntimeException(e)
        }
    }

    /**
     * 读取excel中每一格的内容
     */
    private fun getCellAsString(cell: Cell, formulaEvaluator: FormulaEvaluator): String {
        val cellValue = formulaEvaluator.evaluate(cell) ?: return ""
        return when (cell.cellTypeEnum) {
            CellType.STRING -> cellValue.stringValue
            CellType.BOOLEAN -> cellValue.booleanValue.toString()
            CellType.NUMERIC -> {
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    val date = DateUtil.getJavaCalendar(cellValue.numberValue)
                    SimpleDateFormat.getDateTimeInstance(
                        DateFormat.DEFAULT,
                        DateFormat.DEFAULT,
                        Locale.CHINESE
                    ).format(date.time)
                } else {
                    cellValue.numberValue.toString()
                }
            }
            CellType.ERROR -> ErrorEval.getText(cellValue.errorValue.toInt())
            CellType._NONE -> "<error unexpected cell type>"
            CellType.BLANK -> ""
            CellType.FORMULA -> ""
            else -> "null"
        }
    }

    private val isUnsupportedDevice by lazy {
        Build.VERSION.SDK_INT < Build.VERSION_CODES.LOLLIPOP
    }

    init {
        System.setProperty(
            "org.apache.poi.javax.xml.stream.XMLInputFactory",
            "com.fasterxml.aalto.stax.InputFactoryImpl"
        )
        System.setProperty(
            "org.apache.poi.javax.xml.stream.XMLOutputFactory",
            "com.fasterxml.aalto.stax.OutputFactoryImpl"
        )
        System.setProperty(
            "org.apache.poi.javax.xml.stream.XMLEventFactory",
            "com.fasterxml.aalto.stax.EventFactoryImpl"
        )
    }
}