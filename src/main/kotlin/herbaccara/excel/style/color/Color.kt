package herbaccara.excel.style.color

import org.apache.poi.xssf.usermodel.XSSFColor

interface Color {
    fun toXSSFColor(): XSSFColor
}
