package herbaccara.excel.style

import org.apache.poi.ss.usermodel.Font

enum class Underline {
    NONE, SINGLE, DOUBLE, SINGLE_ACCOUNTING, DOUBLE_ACCOUNTING;

    fun toByte(): Byte {
        return when (this) {
            Underline.NONE -> Font.U_NONE
            Underline.SINGLE -> Font.U_SINGLE
            Underline.DOUBLE -> Font.U_DOUBLE
            Underline.SINGLE_ACCOUNTING -> Font.U_SINGLE_ACCOUNTING
            Underline.DOUBLE_ACCOUNTING -> Font.U_DOUBLE_ACCOUNTING
        }
    }
}
