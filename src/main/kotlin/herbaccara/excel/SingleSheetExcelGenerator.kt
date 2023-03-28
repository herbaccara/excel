package herbaccara.excel

import herbaccara.excel.annotation.ExcelColumn
import herbaccara.excel.annotation.ExcelSheet
import herbaccara.excel.annotation.ExcelStyle
import herbaccara.excel.annotation.Sort
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Sheet

class SingleSheetExcelGenerator<T> @JvmOverloads constructor(
    clazz: Class<T>,
    excelType: ExcelType = ExcelType.SXSSF
) : AbstractExcelGenerator<T>(excelType) {

    private val sheet: Sheet
    private var currentRowIndex: Int = 0
    private val cellInfos: List<CellInfo>
    override val styles: MutableMap<String, CellStyle> = mutableMapOf()

    init {
        val excelSheet = clazz.getAnnotation(ExcelSheet::class.java) ?: throw IllegalArgumentException("")

        sheet = workbook.createSheet(excelSheet.value).apply {
            defaultColumnWidth = excelSheet.columnWidth
            defaultRowHeight = excelSheet.rowHeight
        }
        createStyle(DEFAULT_HEADER_STYLE, excelSheet.headerStyle)
        createStyle(DEFAULT_BODY_STYLE, excelSheet.bodyStyle)

        cellInfos = clazz.declaredFields
            .mapNotNull { field ->
                val excelColumn = field.getAnnotation(ExcelColumn::class.java)
                if (excelColumn != null) {
                    CellInfo(
                        field.apply { isAccessible = true },
                        ExcelColumn(excelColumn.value.ifBlank { field.name }, excelColumn.order)
                    ).also { cellInfo ->
                        val excelStyle = field.getAnnotation(ExcelStyle::class.java)
                        if (excelStyle != null) {
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

        val headerRow = sheet.createRow(currentRowIndex++)
        cellInfos.forEachIndexed { index, cellInfo ->
            val cell = headerRow.createCell(index)
            cell.setCellValue(cellInfo.excelColumn.value)
            cell.cellStyle = styles[DEFAULT_HEADER_STYLE]
        }
    }

    override fun addRows(items: List<T>) {
        items.forEach { item ->
            val row = sheet.createRow(currentRowIndex++)
            cellInfos.forEachIndexed { columnIndex, cellInfo ->
                val cell = row.createCell(columnIndex).apply {
                    cellStyle = styles[cellInfo.styleName()] ?: styles[DEFAULT_BODY_STYLE]
                }
                val value = cellInfo.field.get(item)
                setCellValue(cell, value)
            }
        }
    }
}
