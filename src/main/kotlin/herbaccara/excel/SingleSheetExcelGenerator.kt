package herbaccara.excel

import herbaccara.excel.dataformat.DataFormatStrategy
import herbaccara.excel.dataformat.DefaultDataFormatStrategy
import org.apache.poi.ss.usermodel.Sheet

class SingleSheetExcelGenerator<T> @JvmOverloads constructor(
    clazz: Class<T>,
    excelType: ExcelType = ExcelType.SXSSF,
    dataFormatStrategy: DataFormatStrategy = DefaultDataFormatStrategy()
) : AbstractExcelGenerator<T>(clazz, excelType, dataFormatStrategy) {

    protected val sheet: Sheet
    protected var currentRowIndex: Int = 0

    init {
        sheet = createSheet()
        renderHeader()
    }

    protected fun renderHeader() {
        val row = sheet.createRow(currentRowIndex++)
        cellInfos.forEachIndexed { index, cellInfo ->
            val cell = row.createCell(index)

            cell.cellStyle = headerCellStyle()
            cell.setCellValue(cellInfo.excelColumn.value)
        }
    }

    protected fun renderBody(item: T) {
        val row = sheet.createRow(currentRowIndex++)
        cellInfos.forEachIndexed { index, cellInfo ->
            val cell = row.createCell(index)

            cell.cellStyle = styles[cellInfo.styleName()]
            setCellValue(cell, cellInfo.field.get(item))
        }
    }

    override fun addRows(items: List<T>) {
        items.forEach(::renderBody)
    }
}
