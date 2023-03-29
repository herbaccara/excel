package herbaccara.excel

import herbaccara.excel.annotation.*
import herbaccara.excel.dataformat.DataFormatStrategy
import herbaccara.excel.style.DefaultExcelCellStyle
import herbaccara.excel.style.ExcelCellStyle
import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.*
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
    excelType: ExcelType,
    dataFormatStrategy: DataFormatStrategy
) : ExcelGenerator<T> {

    companion object {
        internal val DEFAULT_HEADER_STYLE = "${this::class.java.name}.DEFAULT_HEADER_STYLE"
        internal val DEFAULT_BODY_STYLE = "${this::class.java.name}.DEFAULT_BODY_STYLE"
    }

    protected val workbook: Workbook = when (excelType) {
        ExcelType.HSSF -> HSSFWorkbook()
        ExcelType.XSSF -> XSSFWorkbook()
        ExcelType.SXSSF -> SXSSFWorkbook()
    }

    protected val maxRows = workbook.spreadsheetVersion.maxRows
    protected val columnWidth: Int
    protected val rowHeight: Short

    protected val styles: MutableMap<String, CellStyle> = mutableMapOf()

    protected val cellInfos: List<CellInfo>

    init {
        val excelSheet = requireNotNull(clazz.getAnnotation(ExcelSheet::class.java))
        columnWidth = excelSheet.columnWidth
        rowHeight = excelSheet.rowHeight

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
                        ExcelColumn(excelColumn.value.ifBlank { field.name }, excelColumn.order)
                    ).also { cellInfo ->
                        val excelStyleClass = field.getAnnotation(ExcelStyleClass::class.java)
                        val excelStyle = field.getAnnotation(ExcelStyle::class.java)

                        val style = if (excelStyleClass != null) {
                            createCellStyle(excelStyleClass.value)
                        } else if (excelStyle != null) {
                            createCellStyle(excelStyle)
                        } else {
                            workbook.createCellStyle().apply {
                                cloneStyleFrom(styles[DEFAULT_BODY_STYLE])
                            }
                        }

                        // 0 이면 한번도 설정을 안한 상태
                        if (style.dataFormat == 0.toShort()) {
                            val type = cellInfo.field.type
                            val dataFormat = dataFormatStrategy.apply(workbook.createDataFormat(), type)
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
                    Sort.NAME -> items.sortedBy { it.excelColumn.value }
                    Sort.ORDER -> items.sortedBy { it.excelColumn.order }
                }
            }
    }

    protected fun createSheet(name: String): Sheet {
        return workbook.createSheet(name).apply {
            defaultColumnWidth = columnWidth
            defaultRowHeight = rowHeight
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

    protected fun createCellStyle(): CellStyle = workbook.createCellStyle()

    protected fun createCellStyle(excelStyleClass: KClass<out ExcelCellStyle>): CellStyle {
        val createInstance = excelStyleClass.createInstance()
        return createInstance.apply(workbook)
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

    override fun write(os: OutputStream) {
        os.use {
            workbook.write(it)
            workbook.close()
            (workbook as? SXSSFWorkbook)?.dispose()
        }
    }
}
