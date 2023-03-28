package herbaccara.excel

import herbaccara.excel.annotation.ExcelColumn
import herbaccara.excel.annotation.ExcelSheet
import herbaccara.excel.annotation.ExcelStyle
import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.OutputStream
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.*

class SingleSheetExcelGenerator<T> @JvmOverloads constructor(
    clazz: Class<T>,
    excelType: ExcelType = ExcelType.SXSSF
) : ExcelGenerator<T> {

    companion object {
        private val DEFAULT_HEADER_STYLE = "${this::class.java.name}.DEFAULT_HEADER_STYLE"
        private val DEFAULT_BODY_STYLE = "${this::class.java.name}.DEFAULT_BODY_STYLE"
    }

    private val workbook: Workbook = when (excelType) {
        ExcelType.HSSF -> HSSFWorkbook()
        ExcelType.XSSF -> XSSFWorkbook()
        ExcelType.SXSSF -> SXSSFWorkbook()
    }

    private val sheet: Sheet
    private var currentRowIndex: Int = 0
    private val cellInfos: List<CellInfo>
    private val styles: MutableMap<String, CellStyle> = mutableMapOf()

    init {
        val excelSheet = clazz.getAnnotation(ExcelSheet::class.java) ?: throw IllegalArgumentException("")

        sheet = workbook.createSheet(excelSheet.value).apply {
            defaultColumnWidth = excelSheet.columnWidth
            defaultRowHeight = excelSheet.rowHeight
        }
        createStyle(DEFAULT_HEADER_STYLE, excelSheet.headerStyle)
        createStyle(DEFAULT_BODY_STYLE, excelSheet.bodyStyle)

        cellInfos = clazz.declaredFields
            .mapNotNull { field ->
                val excelColumn = field.getAnnotation(ExcelColumn::class.java)
                if (excelColumn != null) {
                    CellInfo(field.apply { isAccessible = true }, excelColumn).also { cellInfo ->
                        val excelStyle = field.getAnnotation(ExcelStyle::class.java)
                        if (excelStyle != null) {
                            createStyle(cellInfo.styleName(), excelStyle)
                        }
                    }
                } else {
                    null
                }
            }
            .sortedBy { it.excelColumn.order }

        val headerRow = sheet.createRow(currentRowIndex++)
        cellInfos.forEachIndexed { index, cellInfo ->
            val cell = headerRow.createCell(index)
            cell.setCellValue(cellInfo.field.name)
            cell.cellStyle = styles[DEFAULT_HEADER_STYLE]
        }
    }

    private fun createStyle(styleName: String, excelStyle: ExcelStyle) {
        if (ExcelStyle.isDefault(excelStyle)) return

        val font = workbook.createFont().apply {
            if (excelStyle.fontName.isNotBlank()) {
                fontName = excelStyle.fontName
            }
            fontHeight = excelStyle.fontHeight
            color = excelStyle.fontColor.index
            bold = excelStyle.fontBold
            italic = excelStyle.fontItalic
            strikeout = excelStyle.fontStrikeout
            underline = excelStyle.fontUnderline.toByte()
        }

        val cellStyle = workbook.createCellStyle().apply {
            setFont(font)
            shrinkToFit = excelStyle.shrinkToFit
            fillPattern = excelStyle.fillPattern
            fillForegroundColor = excelStyle.fillForegroundColor.index
            fillBackgroundColor = excelStyle.fillBackgroundColor.index
            alignment = excelStyle.alignment
            verticalAlignment = excelStyle.verticalAlignment

            borderTop = excelStyle.borderStyle
            borderLeft = excelStyle.borderStyle
            borderRight = excelStyle.borderStyle
            borderBottom = excelStyle.borderStyle

            topBorderColor = excelStyle.borderColor.index
            leftBorderColor = excelStyle.borderColor.index
            rightBorderColor = excelStyle.borderColor.index
            bottomBorderColor = excelStyle.borderColor.index
        }

        styles[styleName] = cellStyle
    }

    protected fun setCellValue(cell: Cell, value: Any?) {
        when (value) {
            is Number -> cell.setCellValue(value.toDouble())
            is Boolean -> cell.setCellValue(value)
            is LocalDateTime -> cell.setCellValue(value)
            is LocalDate -> cell.setCellValue(value)
            is Date -> cell.setCellValue(value)
            is RichTextString -> cell.setCellValue(value)
            else -> {
                val str = value?.toString()?.ifBlank { "" } ?: ""
                cell.setCellValue(str)
                @Suppress("HttpUrlsUsage")
                if (str.startsWith("http://") || str.startsWith("https://")) {
                    val hyperlink = workbook.creationHelper.createHyperlink(HyperlinkType.URL).apply {
                        address = str
                    }
                    cell.hyperlink = hyperlink
                }
            }
        }
    }

    override fun addRows(items: List<T>) {
        items.forEach { item ->
            val row = sheet.createRow(currentRowIndex++)
            cellInfos.forEachIndexed { columnIndex, cellInfo ->
                val cell = row.createCell(columnIndex).apply {
                    cellStyle = styles[cellInfo.styleName()] ?: styles[DEFAULT_BODY_STYLE]
                }
                val value = cellInfo.field.get(item)
                setCellValue(cell, value)
            }
        }
    }

    override fun write(os: OutputStream) {
        os.use {
            workbook.write(it)
            workbook.close()
            if (workbook is SXSSFWorkbook) {
                workbook.dispose()
            }
        }
    }
}
