package herbaccara.excel;

import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.util.Arrays;

public class ExcelGeneratorJavaTest {

    @Test
    public void test() throws IOException {
        final ExcelGenerator<Pojo> excelGenerator = new SingleSheetExcelGenerator<>(Pojo.class);
        excelGenerator.addRows(Arrays.asList(
                new Pojo("가", "블라블라\b블~라~블~라~"),
                new Pojo("다", null),
                new Pojo("바", "https://www.google.com")
        ));
        excelGenerator.write(Files.newOutputStream(new File("src/test/resources/test.xlsx").toPath()));
    }
}
