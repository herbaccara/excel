package herbaccara.excel.style

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.VerticalAlignment

data class Style @JvmOverloads constructor(
    val font: Font? = null,

    val fontName: String,
    val fontHeightInPoints: Short = 220,
    val fontBold: Boolean = false,
    val fontItalic: Boolean = false,
    val fontStrikeout: Boolean = false,
    val fontUnderline: Underline = Underline.NONE,
    val shrinkToFit: Boolean = false,
    val pattern: FillPatternType = FillPatternType.NO_FILL,
    val foregroundColor: Short = 64,
    val backgroundColor: Short = 64,
    val borderStyle: BorderStyle = BorderStyle.NONE,
    val borderColor: Short = 8,
    val alignment: HorizontalAlignment = HorizontalAlignment.GENERAL,
    val verticalAlignment: VerticalAlignment = VerticalAlignment.BOTTOM
)
