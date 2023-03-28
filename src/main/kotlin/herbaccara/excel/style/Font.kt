package herbaccara.excel.style

data class Font @JvmOverloads constructor(
    val name: String,
    val heightInPoints: Short? = null,
    val bold: Boolean = false,
    val italic: Boolean = false,
    val strikeout: Boolean = false,
    val underline: Underline = Underline.NONE
)
