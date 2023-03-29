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
        renderHeader(sheet, currentRowIndex++)
    }

    protected fun renderBody(item: T) {
        renderBody(item, sheet, currentRowIndex++)
    }

    override fun addRows(items: List<T>) {
        items.forEach(::renderBody)
    }
}
