package herbaccara.excel.annotation

@Target(AnnotationTarget.CLASS)
@Retention(AnnotationRetention.RUNTIME)
@MustBeDocumented
annotation class ExcelSheet(
    val value: String = "sheet",
    val headerStyles: Array<ExcelStyle> = [],
    val bodyStyles: Array<ExcelStyle> = []
)
