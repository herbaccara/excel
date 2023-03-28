package herbaccara.excel.style;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;
import org.jetbrains.annotations.NotNull;

public class CostCellStyle extends BodyCellStyle {

    @NotNull
    @Override
    public CellStyle apply(@NotNull final Workbook workbook) {
        final CellStyle cellStyle = super.apply(workbook);

        final DataFormat dataFormat = workbook.createDataFormat();
        final short format = dataFormat.getFormat("#,##0");

        cellStyle.setDataFormat(format);
        return cellStyle;
    }
}
