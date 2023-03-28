package herbaccara.excel.annotation

import org.apache.poi.ss.usermodel.*

@Target(AnnotationTarget.CLASS, AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
@MustBeDocumented
annotation class ExcelStyle(
    val fontName: String = "",
    val fontColor: IndexedColors = IndexedColors.BLACK, // 8
    val fontHeight: Short = 220,
    val fontBold: Boolean = false,
    val fontItalic: Boolean = false,
    val fontStrikeout: Boolean = false,
    val fontUnderline: Underline = Underline.NONE,
    val shrinkToFit: Boolean = false,
    val fillPattern: FillPatternType = FillPatternType.NO_FILL,
    val fillForegroundColor: IndexedColors = IndexedColors.AUTOMATIC, // 64
    val fillBackgroundColor: IndexedColors = IndexedColors.AUTOMATIC, // 64
    val borderStyle: BorderStyle = BorderStyle.NONE,
    val borderColor: IndexedColors = IndexedColors.BLACK, // 8
    val alignment: HorizontalAlignment = HorizontalAlignment.GENERAL,
    val verticalAlignment: VerticalAlignment = VerticalAlignment.BOTTOM,
    val dataFormat: String = ""
) {
    companion object {
        private val DefaultExcelStyle = ExcelStyle()

        fun isDefault(style: ExcelStyle): Boolean {
            return DefaultExcelStyle == style
        }
    }
}
