package herbaccara.excel

import herbaccara.excel.annotation.ExcelStyle
import herbaccara.excel.style.ExcelCellStyle
import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.streaming.SXSSFWorkbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.OutputStream
import java.time.LocalDate
import java.time.LocalDateTime
import java.util.*
import kotlin.reflect.KClass
import kotlin.reflect.full.createInstance

abstract class AbstractExcelGenerator<T>(excelType: ExcelType) : ExcelGenerator<T> {

    companion object {
        internal val DEFAULT_HEADER_STYLE = "${this::class.java.name}.DEFAULT_HEADER_STYLE"
        internal val DEFAULT_BODY_STYLE = "${this::class.java.name}.DEFAULT_BODY_STYLE"
    }

    protected val workbook: Workbook = when (excelType) {
        ExcelType.HSSF -> HSSFWorkbook()
        ExcelType.XSSF -> XSSFWorkbook()
        ExcelType.SXSSF -> SXSSFWorkbook()
    }

    protected abstract val styles: MutableMap<String, CellStyle>

    protected fun createStyle(styleName: String, excelStyleClass: KClass<out ExcelCellStyle>) {
        val createInstance = excelStyleClass.createInstance()
        val cellStyle = createInstance.apply(workbook)
        styles[styleName] = cellStyle
    }

    protected fun createStyle(styleName: String, excelStyle: ExcelStyle) {
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

            if (excelStyle.dataFormat.isNotBlank()) {
                val dataFormat = workbook.createDataFormat()
                this.dataFormat = dataFormat.getFormat(excelStyle.dataFormat)
            }
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
            is Calendar -> cell.setCellValue(value)
            is RichTextString -> cell.setCellValue(value)
            is Enum<*> -> cell.setCellValue(value.name)
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

    override fun write(os: OutputStream) {
        os.use {
            workbook.write(it)
            workbook.close()
            (workbook as? SXSSFWorkbook)?.dispose()
        }
    }
}
