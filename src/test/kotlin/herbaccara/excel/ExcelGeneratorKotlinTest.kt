package herbaccara.excel

import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.jupiter.api.Test

class ExcelGeneratorKotlinTest {

    @Test
    fun test2() {
        val sxssExcelGenerator = SXSSExcelGenerator(Pojo::class.java)
        println()
    }

    @Test
    fun test() {
        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet()
        val row = sheet.createRow(0)
        val cell = row.createCell(0)
        cell.cellType = CellType.STRING

        val createCellStyle = workbook.createCellStyle()
        createCellStyle.borderTop = BorderStyle.NONE
        createCellStyle.alignment = HorizontalAlignment.CENTER
        createCellStyle.verticalAlignment = VerticalAlignment.CENTER
        createCellStyle.fillBackgroundColor = IndexedColors.BLACK.index
        createCellStyle.fillPattern = FillPatternType.NO_FILL
        val red: Byte = 255.toByte()
        println()
        val xssfColor = XSSFColor(byteArrayOf(red, red, red), DefaultIndexedColorMap())

        val createFont = workbook.createFont()
        createFont.bold = true
        createFont.italic = true
        createFont.underline = HSSFFont.U_NONE
        createFont.fontHeightInPoints

        println("test")
    }
}
