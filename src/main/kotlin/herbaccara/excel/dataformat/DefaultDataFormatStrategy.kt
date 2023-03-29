package herbaccara.excel.dataformat

import org.apache.poi.ss.usermodel.DataFormat

class DefaultDataFormatStrategy : DataFormatStrategy {

    override fun apply(dataFormat: DataFormat, type: Class<*>): Short {
        val kClass = type.kotlin

        val integerTypes = listOf(Byte::class, Short::class, Int::class, Long::class)
        val realTypes = listOf(Float::class, Double::class)

        val format = when {
            integerTypes.contains(kClass) -> "#,##0"
            realTypes.contains(kClass) -> "#,##0.00"
            else -> ""
        }

        return dataFormat.getFormat(format)
    }
}
