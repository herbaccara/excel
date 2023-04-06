package herbaccara.excel

import java.io.OutputStream

interface ExcelGenerator<T> {

    fun excelType(): ExcelType

    fun fileExtension(): String

    fun mediaType(): String

    fun addRows(items: List<T>)

    fun write(os: OutputStream)
}
