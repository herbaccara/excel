package herbaccara.excel.annotation

import herbaccara.excel.style.ExcelCellStyle
import org.apache.poi.ss.usermodel.*
import kotlin.reflect.KClass

@Target(AnnotationTarget.FIELD)
@Retention(AnnotationRetention.RUNTIME)
@MustBeDocumented
annotation class ExcelStyleClass(val value: KClass<out ExcelCellStyle>)
