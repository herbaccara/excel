package herbaccara.excel;

import herbaccara.excel.annotation.ExcelColumn;
import herbaccara.excel.annotation.ExcelSheet;
import herbaccara.excel.annotation.ExcelStyle;
import herbaccara.excel.style.BodyCellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Random;

@ExcelSheet(
        value = "pojos",
        headerStyle = @ExcelStyle(
                fillPattern = FillPatternType.SOLID_FOREGROUND,
                fillForegroundColor = IndexedColors.GREEN,
                borderStyle = BorderStyle.THIN,
                fontBold = true
        ),
        bodyStyleClass = BodyCellStyle.class,
        freezeHeaderPane = true
)
public class Pojo {

    public Pojo(final String foo, final String bar) {
        this.foo = foo;
        this.bar = bar;
    }

    @ExcelColumn(value = "a푸")
    @ExcelStyle(fontBold = true)
    private String foo;

    @ExcelColumn(value = "b파")
    private String bar;

    @ExcelColumn("비용")
//    @ExcelStyleClass(CostCellStyle.class)
    private Integer cost = new Random().nextInt(1_000_000);

    @ExcelColumn("비용xxxxxxx")
    private float asdasd = new Random().nextFloat(1_000_000);

    @ExcelColumn("long~~~~")
    private Long longggg = new Random().nextLong(1_000_000);

    public Long getLongggg() {
        return longggg;
    }

    public void setLongggg(final Long longggg) {
        this.longggg = longggg;
    }

    public float getAsdasd() {
        return asdasd;
    }

    public void setAsdasd(final float asdasd) {
        this.asdasd = asdasd;
    }

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
