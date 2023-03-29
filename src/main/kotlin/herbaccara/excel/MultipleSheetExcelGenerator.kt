package herbaccara.excel

import herbaccara.excel.dataformat.DataFormatStrategy
import herbaccara.excel.dataformat.DefaultDataFormatStrategy
import org.apache.poi.ss.usermodel.Sheet

class MultipleSheetExcelGenerator<T> @JvmOverloads constructor(
    clazz: Class<T>,
    chunk: Int? = null,
    excelType: ExcelType = ExcelType.SXSSF,
    dataFormatStrategy: DataFormatStrategy = DefaultDataFormatStrategy()
) : AbstractExcelGenerator<T>(clazz, excelType, dataFormatStrategy) {

    protected val sheets: MutableList<Sheet> = mutableListOf()
    protected lateinit var currentSheet: Sheet
    protected var currentRowIndex: Int = 0
    protected var rownum = 0
    protected val chunk: Int = if (chunk == null || chunk > maxRows) maxRows else chunk

    init {
        createSheetWithHeader()
    }

    protected fun createSheetWithHeader() {
        currentSheet = createSheet("_${(sheets.size + 1)}").also {
            sheets.add(it)
        }
        currentRowIndex = 0
        renderHeader()
    }

    protected fun renderHeader() {
        renderHeader(currentSheet, currentRowIndex++)
    }

    protected fun renderBody(item: T) {
        renderBody(item, currentSheet, currentRowIndex++)
    }

    override fun addRows(items: List<T>) {
        items.forEach { item ->
            if (rownum > 0 && rownum % chunk == 0) {
                createSheetWithHeader()
            }
            renderBody(item)
            rownum++
        }
    }
}
