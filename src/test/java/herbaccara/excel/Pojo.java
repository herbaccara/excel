package herbaccara.excel;

import herbaccara.excel.annotation.ExcelColumn;
import herbaccara.excel.annotation.ExcelSheet;
import herbaccara.excel.annotation.ExcelStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

@ExcelSheet(
        value = "pojos",
        headerStyle = @ExcelStyle(
                fillPattern =  FillPatternType.SOLID_FOREGROUND,
                fillForegroundColor = IndexedColors.GREEN,
                borderStyle = BorderStyle.THIN,
                fontBold = true
        ),
        bodyStyle = @ExcelStyle(
                fontItalic = true
        )
)
public class Pojo {

    public Pojo(final String foo, final String bar) {
        this.foo = foo;
        this.bar = bar;
    }

    @ExcelColumn(value = "푸", order = 1)
    @ExcelStyle(fontBold = true)
    private String foo;

    @ExcelColumn(value = "바", order = 2)
    private String bar;

    public String getFoo() {
        return foo;
    }

    public void setFoo(final String foo) {
        this.foo = foo;
    }

    public String getBar() {
        return bar;
    }

    public void setBar(final String bar) {
        this.bar = bar;
    }
}
