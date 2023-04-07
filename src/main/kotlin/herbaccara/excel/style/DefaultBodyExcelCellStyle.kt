package herbaccara.excel.style

import herbaccara.excel.extension.setBorder
import org.apache.poi.ss.usermodel.*

class DefaultBodyExcelCellStyle : ExcelCellStyle {

    override fun apply(workbook: Workbook): CellStyle {
        return workbook.createCellStyle().apply {
            alignment = HorizontalAlignment.RIGHT
            verticalAlignment = VerticalAlignment.CENTER
            setBorder(BorderStyle.THIN)
        }
    }
}
