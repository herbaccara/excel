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
        styles[DEFAULT_HEADER_STYLE] = if (excelSheet.headerStyleClass == DefaultExcelCellStyle::class) {
            createStyle(excelSheet.headerStyle)
        } else {
            createStyle(excelSheet.headerStyleClass)
        }

        // body style
        val bodyStyle = if (excelSheet.bodyStyleClass == DefaultExcelCellStyle::class) {
            createStyle(excelSheet.bodyStyle)
        } else {
            createStyle(excelSheet.bodyStyleClass)
        }
        styles[DEFAULT_BODY_STYLE] = bodyStyle

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

                        val style = if (excelStyleClass != null) {
                            createStyle(excelStyleClass.value)
                        } else if (excelStyle != null) {
                            createStyle(excelStyle)
                        } else {
                            workbook.createCellStyle().apply {
                                cloneStyleFrom(bodyStyle)
                            }
                        }

                        // 0 이면 한번도 설정을 안한 상태
                        if (style.dataFormat == 0.toShort()) {
                            val type = cellInfo.field.type.kotlin

                            val integerTypes = listOf(Byte::class, Short::class, Int::class, Long::class)
                            val realTypes = listOf(Float::class, Double::class)

                            val isIntegerType = integerTypes.contains(type) // #,##0
                            val isRealType = realTypes.contains(type) // #,##0.00

                            // TODO : type 에 따른 dataFormat 처리
                        }
                        styles[cellInfo.styleName()] = style
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

            cell.cellStyle = styles[cellInfo.styleName()]
            setCellValue(cell, cellInfo.field.get(item))
        }
    }

    override fun addRows(items: List<T>) {
        items.forEach(::renderBody)
    }
}
