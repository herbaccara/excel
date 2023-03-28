package herbaccara.excel

import java.io.OutputStream

interface ExcelGenerator<T> {

    fun addRows(items: List<T>)

    fun write(os: OutputStream)
}
