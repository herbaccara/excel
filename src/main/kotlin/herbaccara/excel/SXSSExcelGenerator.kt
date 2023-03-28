package herbaccara.excel

import herbaccara.excel.annotation.ExcelColumn
import herbaccara.excel.annotation.ExcelSheet
import herbaccara.excel.annotation.ExcelStyle
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import java.io.OutputStream
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.*

class SXSSExcelGenerator<T>(
    clazz: Class<T>
) : ExcelGenerator<T> {

    private val workbook: SXSSFWorkbook
    private val sheet: SXSSFSheet
    private var currentRowIndex: Int = 0
    private val cellInfos: List<CellInfo>

    init {
        val excelSheet = clazz.getAnnotation(ExcelSheet::class.java) ?: throw IllegalArgumentException("")

        cellInfos = clazz.declaredFields.mapNotNull {
            val excelColumn = it.getAnnotation(ExcelColumn::class.java)
            if (excelColumn != null) {
                val excelStyle = it.getAnnotation(ExcelStyle::class.java)
                CellInfo(it.apply { isAccessible = true }, excelColumn, excelStyle)
            } else {
                null
            }
        }.sortedBy { it.excelColumn.order }

        workbook = SXSSFWorkbook()
        sheet = workbook.createSheet(excelSheet.value)

        val headerRow = sheet.createRow(currentRowIndex++)
        cellInfos.forEachIndexed { index, cellInfo ->
            val cell = headerRow.createCell(index)
            cell.setCellValue(cellInfo.field.name)
        }
    }

    protected fun setCellValue(cell: Cell, value: Any?) {
        when (value) {
            is Number -> cell.setCellValue(value.toDouble())
            is Boolean -> cell.setCellValue(value)
            is LocalDateTime -> cell.setCellValue(value)
            is LocalDate -> cell.setCellValue(value)
            is Date -> cell.setCellValue(value)
            is RichTextString -> cell.setCellValue(value)
            else -> cell.setCellValue(value?.toString()?.ifBlank { "" } ?: "")
        }
    }

    override fun addRows(items: List<T>) {
        items.forEach { item ->
            val row = sheet.createRow(currentRowIndex++)
            cellInfos.forEachIndexed { columnIndex, cellInfo ->
                val cell = row.createCell(columnIndex)
                val value = cellInfo.field.get(item)
                setCellValue(cell, value)
            }
        }
    }

    override fun write(os: OutputStream) {
        os.use {
            workbook.run {
                write(it)
                close()
                dispose()
            }
        }
    }
}
