package herbaccara.excel.annotation

@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
@MustBeDocumented
annotation class ExcelColumn(
    val value: String = "",
    val order: Int = 0,
    val width: Int = 2048,
    val autoSize: Boolean = false
)
