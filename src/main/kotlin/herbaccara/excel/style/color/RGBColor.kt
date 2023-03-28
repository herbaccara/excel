package herbaccara.excel.style.color

import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap
import org.apache.poi.xssf.usermodel.XSSFColor

data class RGBColor(val red: Byte, val green: Byte, val blue: Byte) : Color {
    constructor(red: Int, green: Int, blue: Int) : this(red.toByte(), green.toByte(), blue.toByte())

    override fun toXSSFColor(): XSSFColor {
        return XSSFColor(byteArrayOf(red, green, blue), DefaultIndexedColorMap())
    }
}
