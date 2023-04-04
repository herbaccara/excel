package herbaccara.excel;

import herbaccara.excel.annotation.ExcelColumn;
import herbaccara.excel.annotation.ExcelSheet;
import herbaccara.excel.annotation.ExcelStyle;
import herbaccara.excel.style.BodyCellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
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

    @ExcelColumn("date")
    private LocalDate date = LocalDate.now();

    @ExcelColumn(value = "date2", width = 4096)
    private Date date2 = new Date();

    public Date getDate2() {
        return date2;
    }

    public void setDate2(final Date date2) {
        this.date2 = date2;
    }

    public LocalDate getDate() {
        return date;
    }

    public void setDate(final LocalDate date) {
        this.date = date;
    }

    @ExcelColumn("dateTime")
    private LocalDateTime dateTime = LocalDateTime.now();

    public LocalDateTime getDateTime() {
        return dateTime;
    }

    public void setDateTime(final LocalDateTime dateTime) {
        this.dateTime = dateTime;
    }

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
