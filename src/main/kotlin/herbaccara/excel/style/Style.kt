package herbaccara.excel.style

import herbaccara.excel.style.color.Color
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment

data class Style @JvmOverloads constructor(
    val font: Font? = null,
    val shrinkToFit: Boolean = false,
    val pattern: FillPatternType = FillPatternType.NO_FILL,
    val foregroundColor: Color? = null,
    val backgroundColor: Color? = null,
    val borderStyle: BorderStyle = BorderStyle.NONE,
    val borderColor: Color? = null,
    val alignment: HorizontalAlignment? = null,
    val verticalAlignment: VerticalAlignment? = null
)
