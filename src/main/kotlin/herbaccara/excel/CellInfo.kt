package herbaccara.excel

import herbaccara.excel.annotation.ExcelColumn
import java.lang.reflect.Field

data class CellInfo(
    val field: Field,
    val excelColumn: ExcelColumn
) {
    fun styleName(): String {
        return "${field.name}Style"
    }
}
