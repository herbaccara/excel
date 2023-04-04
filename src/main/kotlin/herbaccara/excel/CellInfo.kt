package herbaccara.excel

import java.lang.reflect.Field

data class CellInfo(
    val field: Field,
    private val name: String,
    val order: Int,
    val width: Int
) {
    fun name(): String {
        return name.ifBlank { field.name }
    }

    fun styleName(): String {
        return "${field.name}Style"
    }
}
