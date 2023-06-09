package herbaccara.excel

import herbaccara.excel.annotation.*
import herbaccara.excel.dataformat.DataFormatStrategy
import herbaccara.excel.extension.setBorder
import herbaccara.excel.extension.setBorderColor
import herbaccara.excel.style.DefaultExcelCellStyle
import herbaccara.excel.style.ExcelCellStyle
import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.streaming.SXSSFSheet
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.OutputStream
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.*
import kotlin.reflect.KClass
import kotlin.reflect.full.createInstance

abstract class AbstractExcelGenerator<T>(
    clazz: Class<T>,
    private val excelType: ExcelType,
    dataFormatStrategy: DataFormatStrategy
) : ExcelGenerator<T> {

    companion object {
        private val DEFAULT_HEADER_STYLE = "${this::class.java.name}.DEFAULT_HEADER_STYLE"
        private val DEFAULT_BODY_STYLE = "${this::class.java.name}.DEFAULT_BODY_STYLE"
        private val excelCellStyleMap: MutableMap<KClass<out ExcelCellStyle>, ExcelCellStyle> = mutableMapOf()
    }

    protected val workbook: Workbook = when (excelType) {
        ExcelType.HSSF -> HSSFWorkbook()
        ExcelType.XSSF -> XSSFWorkbook()
        ExcelType.SXSSF -> SXSSFWorkbook()
    }

    protected val maxRows = workbook.spreadsheetVersion.maxRows
    protected val defaultSheetName: String
    protected val defaultColumnWidth: Int
    protected val defaultRowHeight: Short
    protected val freezePane: Boolean

    protected val styles: MutableMap<String, CellStyle> = mutableMapOf()

    protected val cellInfos: List<CellInfo>

    override fun excelType(): ExcelType = excelType

    override fun fileExtension(): String = when (excelType) {
        ExcelType.HSSF -> "xls"
        ExcelType.XSSF, ExcelType.SXSSF -> "xlsx"
    }

    override fun mediaType(): String = when (excelType) {
        ExcelType.HSSF -> "application/vnd.ms-excel"
        ExcelType.XSSF, ExcelType.SXSSF -> "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }

    init {
        val excelSheet = requireNotNull(clazz.getAnnotation(ExcelSheet::class.java))
        defaultSheetName = excelSheet.value.ifBlank { "Sheet" }
        defaultColumnWidth = excelSheet.columnWidth
        defaultRowHeight = excelSheet.rowHeight
        freezePane = excelSheet.freezeHeaderPane

        styles[DEFAULT_HEADER_STYLE] = if (excelSheet.headerStyleClass == DefaultExcelCellStyle::class) {
            createCellStyle(excelSheet.headerStyle)
        } else {
            createCellStyle(excelSheet.headerStyleClass)
        }

        styles[DEFAULT_BODY_STYLE] = if (excelSheet.bodyStyleClass == DefaultExcelCellStyle::class) {
            createCellStyle(excelSheet.bodyStyle)
        } else {
            createCellStyle(excelSheet.bodyStyleClass)
        }

        cellInfos = clazz.declaredFields
            .mapNotNull { field ->
                val excelColumn = field.getAnnotation(ExcelColumn::class.java)
                if (excelColumn != null) {
                    CellInfo(
                        field.apply { isAccessible = true },
                        excelColumn.value,
                        excelColumn.order,
                        excelColumn.width,
                        excelColumn.autoSize
                    ).also { cellInfo ->
                        val excelStyleClass = field.getAnnotation(ExcelStyleClass::class.java)
                        val excelStyle = field.getAnnotation(ExcelStyle::class.java)

                        val style = if (excelStyleClass != null) {
                            createCellStyle(excelStyleClass.value)
                        } else if (excelStyle != null) {
                            createCellStyle(excelStyle)
                        } else {
                            createCellStyle(bodyCellStyle())
                        }

                        // 0 이면 한번도 설정을 안한 상태
                        if (style.dataFormat == 0.toShort()) {
                            val type = cellInfo.field.type
                            val dataFormat = dataFormatStrategy.apply(createDataFormat(), type)
                            style.dataFormat = dataFormat
                        }
                        styles[cellInfo.styleName()] = style
                    }
                } else {
                    null
                }
            }
            .let { items ->
                when (excelSheet.fieldSort) {
                    Sort.NONE -> items
                    Sort.NAME -> items.sortedBy { it.name() }
                    Sort.ORDER -> items.sortedBy { it.order }
                }
            }
    }

    protected fun createSheet(suffix: String = ""): Sheet {
        return workbook.createSheet("$defaultSheetName$suffix").also {
            it.defaultColumnWidth = defaultColumnWidth
            it.defaultRowHeight = defaultRowHeight
            if (freezePane) {
                it.createFreezePane(0, 1)
            }
            for ((index, cellInfo) in cellInfos.withIndex()) {
                it.setColumnWidth(index, cellInfo.width)
            }
        }
    }

    // data format

    protected fun createDataFormat(): DataFormat = workbook.createDataFormat()

    protected fun createDataFormat(format: String): Short = workbook.createDataFormat().getFormat(format)

    // hyper link

    private val hyperlinkPatterns = listOf("http://", "https://", "mailto:")

    protected fun createHyperlink(type: HyperlinkType): Hyperlink = workbook.creationHelper.createHyperlink(type)

    protected fun createHyperlink(value: String): Hyperlink {
        @Suppress("HttpUrlsUsage")
        val type = when {
            value.startsWith("http://") || value.startsWith("https://") -> HyperlinkType.URL
            value.startsWith("mailto:") -> HyperlinkType.EMAIL
            else -> HyperlinkType.NONE
        }

        return createHyperlink(type).apply {
            address = value
        }
    }

    // font

    protected fun createFont(): Font = workbook.createFont()

    // cell style

    protected fun headerCellStyle(): CellStyle = styles[DEFAULT_HEADER_STYLE]!!

    protected fun bodyCellStyle(): CellStyle = styles[DEFAULT_BODY_STYLE]!!

    protected fun createCellStyle(): CellStyle = workbook.createCellStyle()

    protected fun createCellStyle(source: CellStyle): CellStyle =
        workbook.createCellStyle().apply { cloneStyleFrom(source) }

    protected fun createCellStyle(excelStyleClass: KClass<out ExcelCellStyle>): CellStyle {
        val excelCellStyle = excelCellStyleMap.getOrPut(excelStyleClass) {
            excelStyleClass.createInstance()
        }
        return excelCellStyle.apply(workbook)
    }

    protected fun createCellStyle(excelStyle: ExcelStyle): CellStyle {
        val font = createFont().apply {
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

        val cellStyle = createCellStyle().apply {
            setFont(font)

            fillPattern = excelStyle.fillPattern
            fillForegroundColor = excelStyle.fillForegroundColor.index
            fillBackgroundColor = excelStyle.fillBackgroundColor.index

            alignment = excelStyle.alignment
            verticalAlignment = excelStyle.verticalAlignment

            setBorder(excelStyle.borderStyle)

            setBorderColor(excelStyle.borderColor.index)

            shrinkToFit = excelStyle.shrinkToFit
            wrapText = excelStyle.wrapText
            quotePrefixed = excelStyle.quotePrefixed
            rotation = excelStyle.rotation

            if (excelStyle.dataFormat.isNotBlank()) {
                dataFormat = createDataFormat(excelStyle.dataFormat)
            }
        }

        return cellStyle
    }

    protected fun setCellValue(cell: Cell, value: Any?) {
        when (value) {
            is Number -> cell.setCellValue(value.toDouble())
            is Boolean -> cell.setCellValue(value)
            is LocalDateTime -> cell.setCellValue(value)
            is LocalDate -> cell.setCellValue(value)
            is Date -> cell.setCellValue(value)
            is Calendar -> cell.setCellValue(value)
            is RichTextString -> cell.setCellValue(value)
            is Enum<*> -> cell.setCellValue(value.name)
            else -> {
                val str = value?.toString()?.ifBlank { "" } ?: ""
                cell.setCellValue(str)
                if (hyperlinkPatterns.any { str.startsWith(it) }) {
                    cell.hyperlink = createHyperlink(str)
                }
            }
        }
    }

    protected fun renderHeader(sheet: Sheet, rownum: Int) {
        val row = sheet.createRow(rownum)
        cellInfos.forEachIndexed { index, cellInfo ->
            val cell = row.createCell(index)

            cell.cellStyle = headerCellStyle()
            cell.setCellValue(cellInfo.name())
        }
    }

    protected fun renderBody(item: T, sheet: Sheet, rownum: Int) {
        val row = sheet.createRow(rownum)
        cellInfos.forEachIndexed { index, cellInfo ->
            val cell = row.createCell(index)

            cell.cellStyle = styles[cellInfo.styleName()]
            setCellValue(cell, cellInfo.field.get(item))
        }
    }

    protected fun autoSizeColumns(sheet: Sheet) {
        for ((index, cellInfo) in cellInfos.withIndex()) {
            if (cellInfo.autoSize) {
                if (sheet is SXSSFSheet) {
                    sheet.trackColumnForAutoSizing(index)
                }
                sheet.autoSizeColumn(index)
            }
        }
    }

    override fun write(os: OutputStream) {
        os.use {
            workbook.write(it)
            workbook.close()
            (workbook as? SXSSFWorkbook)?.dispose()
        }
    }
}
