package herbaccara.excel.annotation

@Target(AnnotationTarget.CLASS)
@Retention(AnnotationRetention.RUNTIME)
@MustBeDocumented
annotation class ExcelSheet(
    val value: String = "sheet",
    val columnWidth: Int = 8,
    val rowHeight: Short = 300,
    val headerStyle: ExcelStyle = ExcelStyle(),
    val bodyStyle: ExcelStyle = ExcelStyle(),
    val fieldSort: Sort = Sort.NONE
)
