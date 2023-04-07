package herbaccara.excel.extension

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle

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
