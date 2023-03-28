package herbaccara.excel

import herbaccara.excel.annotation.ExcelColumn
import herbaccara.excel.annotation.ExcelStyle
import java.lang.reflect.Field

data class CellInfo(
    val field: Field,
    val excelColumn: ExcelColumn,
    val excelStyle: ExcelStyle? = null
)
