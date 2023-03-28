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

    public Pojo(final String foo, final String bar, final Integer cost) {
        this.foo = foo;
        this.bar = bar;
        this.cost = cost;
    }

    @ExcelColumn(value = "a푸")
    @ExcelStyle(fontBold = true)
    private String foo;

    @ExcelColumn(value = "b파")
    private String bar;

    @ExcelColumn
    @ExcelStyle(dataFormat = "#,##0")
    private Integer cost;

    public Integer getCost() {
        return cost;
    }

    public void setCost(final Integer cost) {
        this.cost = cost;
    }

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
