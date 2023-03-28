package herbaccara.excel

import herbaccara.excel.annotation.ExcelColumn
import herbaccara.excel.annotation.ExcelSheet
import herbaccara.excel.annotation.ExcelStyle
import herbaccara.excel.style.BodyCellStyle
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors

@ExcelSheet(
    value = "pojos",
    headerStyle = ExcelStyle(
        fillPattern = FillPatternType.SOLID_FOREGROUND,
        fillForegroundColor = IndexedColors.GREEN,
        borderStyle = BorderStyle.THIN,
        fontBold = true
    ),
    bodyStyleClass = BodyCellStyle::class
)
class Pojo2 {
    constructor(foo: String, bar: String) {
        this.foo = foo
        this.bar = bar
    }

    constructor(foo: String, bar: String, cost: Int?) {
        this.foo = foo
        this.bar = bar
        this.cost = cost
    }

    @ExcelColumn(value = "a푸")
    @ExcelStyle(fontBold = true)
    var foo: String

    @ExcelColumn(value = "b파")
    var bar: String

    @ExcelColumn("비용")
    var cost: Int? = null

    @ExcelColumn("비용222")
    var costasd: Long? = null
}
