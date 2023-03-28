package herbaccara.excel.annotation

import herbaccara.excel.style.DefaultExcelCellStyle
import herbaccara.excel.style.ExcelCellStyle
import kotlin.reflect.KClass

@Target(AnnotationTarget.CLASS)
@Retention(AnnotationRetention.RUNTIME)
@MustBeDocumented
annotation class ExcelSheet(
    val value: String = "sheet",
    val columnWidth: Int = 8,
    val rowHeight: Short = 300,
    val headerStyle: ExcelStyle = ExcelStyle(),
    val headerStyleClass: KClass<out ExcelCellStyle> = DefaultExcelCellStyle::class,
    val bodyStyle: ExcelStyle = ExcelStyle(),
    val bodyStyleClass: KClass<out ExcelCellStyle> = DefaultExcelCellStyle::class,
    val fieldSort: Sort = Sort.NONE
)
