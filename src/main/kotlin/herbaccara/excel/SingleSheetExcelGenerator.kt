package herbaccara.excel

import herbaccara.excel.annotation.*
import herbaccara.excel.style.DefaultExcelCellStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Sheet

class SingleSheetExcelGenerator<T> @JvmOverloads constructor(
    clazz: Class<T>,
    excelType: ExcelType = ExcelType.SXSSF
) : AbstractExcelGenerator<T>(excelType) {

    protected val sheet: Sheet
    protected var currentRowIndex: Int = 0
    protected val cellInfos: List<CellInfo>
    override val styles: MutableMap<String, CellStyle> = mutableMapOf()

    init {
        val excelSheet = clazz.getAnnotation(ExcelSheet::class.java) ?: throw IllegalArgumentException("")

        sheet = workbook.createSheet(excelSheet.value).apply {
            defaultColumnWidth = excelSheet.columnWidth
            defaultRowHeight = excelSheet.rowHeight
        }

        // header style
        if (excelSheet.headerStyleClass == DefaultExcelCellStyle::class) {
            createStyle(DEFAULT_HEADER_STYLE, excelSheet.headerStyle)
        } else {
            createStyle(DEFAULT_HEADER_STYLE, excelSheet.headerStyleClass)
        }

        // body style
        if (excelSheet.bodyStyleClass == DefaultExcelCellStyle::class) {
            createStyle(DEFAULT_BODY_STYLE, excelSheet.bodyStyle)
        } else {
            createStyle(DEFAULT_BODY_STYLE, excelSheet.bodyStyleClass)
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

                        if (excelStyleClass != null) {
                            createStyle(cellInfo.styleName(), excelStyleClass.value)
                        } else if (excelStyle != null) {
                            createStyle(cellInfo.styleName(), excelStyle)
                        }
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

        renderHeader()
    }

    protected fun renderHeader() {
        val row = sheet.createRow(currentRowIndex++)
        cellInfos.forEachIndexed { index, cellInfo ->
            val cell = row.createCell(index)

            cell.cellStyle = styles[DEFAULT_HEADER_STYLE]
            cell.setCellValue(cellInfo.excelColumn.value)
        }
    }

    protected fun renderBody(item: T) {
        val row = sheet.createRow(currentRowIndex++)
        cellInfos.forEachIndexed { index, cellInfo ->
            val cell = row.createCell(index)

            cell.cellStyle = styles[cellInfo.styleName()] ?: styles[DEFAULT_BODY_STYLE]
            setCellValue(cell, cellInfo.field.get(item))
        }
    }

    override fun addRows(items: List<T>) {
        items.forEach(::renderBody)
    }
}
