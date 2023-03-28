package herbaccara.excel.style.color

import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap
import org.apache.poi.xssf.usermodel.XSSFColor

data class IndexedColor(val color: IndexedColors) : Color {
    override fun toXSSFColor(): XSSFColor {
        return XSSFColor(color, DefaultIndexedColorMap())
    }
}
