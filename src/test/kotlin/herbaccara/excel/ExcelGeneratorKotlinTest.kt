package herbaccara.excel

import org.apache.poi.hssf.usermodel.HSSFFont
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap
import org.apache.poi.xssf.usermodel.XSSFColor
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.jupiter.api.Test
import java.io.File
import java.nio.file.Files

class ExcelGeneratorKotlinTest {

    @Test
    fun test2() {
        val excelGenerator: ExcelGenerator<Pojo2> = SingleSheetExcelGenerator(Pojo2::class.java)
        excelGenerator.addRows(
            listOf(
//                Pojo2("가", "블라블라\b블~라~블~라~"),
//                Pojo2("나", ""),
//                Pojo2("다", "https://www.google.com"),
//                Pojo2("라", "https://www.naver.com", 100000),
                Pojo2("마", "https://www.naver.com", 100000).apply {
                    this.costasd = 10
                }
            )
        )
        excelGenerator.write(Files.newOutputStream(File("src/test/resources/test.xlsx").toPath()))
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
