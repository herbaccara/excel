package herbaccara.excel.extension

import herbaccara.excel.style.color.RGBColor
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFColor

fun CellStyle.setBorder(borderStyle: BorderStyle) {
    borderTop = borderStyle
    borderLeft = borderStyle
    borderRight = borderStyle
    borderBottom = borderStyle
}

fun CellStyle.setBorderColor(color: Short) {
    topBorderColor = color
    leftBorderColor = color
    rightBorderColor = color
    bottomBorderColor = color
}

fun XSSFCellStyle.setFillForegroundColor(color: RGBColor) {
    setFillBackgroundColor(XSSFColor(color.byteArray(), DefaultIndexedColorMap()))
}
