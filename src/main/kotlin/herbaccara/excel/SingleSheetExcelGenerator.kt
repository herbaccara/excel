package herbaccara.excel

import herbaccara.excel.dataformat.DataFormatStrategy
import herbaccara.excel.dataformat.DefaultDataFormatStrategy
import org.apache.poi.ss.usermodel.Sheet
import java.io.OutputStream

class SingleSheetExcelGenerator<T> @JvmOverloads constructor(
    clazz: Class<T>,
    dataFormatStrategy: DataFormatStrategy = DefaultDataFormatStrategy(),
    excelType: ExcelType = ExcelType.SXSSF
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

    override fun write(os: OutputStream) {
        autoSizeColumns(sheet)
        super.write(os)
    }
}
