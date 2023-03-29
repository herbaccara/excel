package herbaccara.excel.dataformat

import org.apache.poi.ss.usermodel.DataFormat

interface DataFormatStrategy {

    fun apply(dataFormat: DataFormat, type: Class<*>): Short
}
