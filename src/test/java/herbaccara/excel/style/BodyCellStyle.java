package herbaccara.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.jetbrains.annotations.NotNull;

public class BodyCellStyle implements ExcelCellStyle {

    @NotNull
    @Override
    public CellStyle apply(@NotNull final Workbook workbook) {
        final CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
        return cellStyle;
    }
}
