import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import static org.hamcrest.CoreMatchers.is;
import static org.hamcrest.Matchers.closeTo;
import static org.junit.Assert.assertThat;


/**
 * @author hf_cherish
 * @date 2018/7/25
 */
public class XLSXTest {
    @Test
    void should_support_basic_sheet_operation() throws IOException, InvalidFormatException {
//        load excel
        XSSFWorkbook workbook = new XSSFWorkbook(getFile("test.xlsx"));

//        get sheet by name
        String sheetName = "Product Mix";
        XSSFSheet sheet = workbook.getSheet(sheetName);
        assertThat(sheet.getSheetName(), is(sheetName));

//        get cell by coordinate
        XSSFCell tvsetNumber = getCell(sheet, "D4");
        assertThat(tvsetNumber.getNumericCellValue(), closeTo(100, 0.1));

//        get formula cell by coordinate
        XSSFCell total = getCell(sheet, "D13");
        assertThat(total.getNumericCellValue(), closeTo(16000, 0.1));

//        change cell value
        tvsetNumber.setCellValue(200);
        assertThat(tvsetNumber.getNumericCellValue(), closeTo(200, 0.1));

//        refresh formulas
        XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
        assertThat(total.getNumericCellValue(), closeTo(23500, 0.1));

//        write back to file
        workbook.write(new FileOutputStream(getFile("update.xlsx")));

//        check data updated
        XSSFWorkbook updateWorkBook = new XSSFWorkbook(getFile("update.xlsx"));
        assertThat(getCell(updateWorkBook.getSheet(sheetName), "D4").getNumericCellValue(), closeTo(200, 0.1));
        assertThat(getCell(updateWorkBook.getSheet(sheetName), "D13").getNumericCellValue(), closeTo(23500, 0.1));
    }

    @Test
    void should_able_to_get_cell_by_name() throws IOException, InvalidFormatException {
        XSSFWorkbook workbook = new XSSFWorkbook(getFile("test.xlsx"));

        XSSFCell cell = getCellByName(workbook, "test_name");

        assertThat(cell.getStringCellValue(), is("TV set"));


    }

    private XSSFCell getCellByName(XSSFWorkbook workbook, String cellName) {
        XSSFName test_name = workbook.getName(cellName);
        AreaReference areaReference = new AreaReference(test_name.getRefersToFormula(), SpreadsheetVersion.EXCEL2007);

        CellReference firstCell = areaReference.getFirstCell();
        XSSFSheet sheet = workbook.getSheet(firstCell.getSheetName());
        XSSFRow row = sheet.getRow(firstCell.getRow());
        return row.getCell(firstCell.getCol());
    }

    private void copyFileUsingStream(File source, File dest) throws IOException {
        FileInputStream inputStream = null;
        FileOutputStream outputStream = null;
        try {
            inputStream = new FileInputStream(source);
            outputStream = new FileOutputStream(dest);

            byte[] buffer = new byte[1024];

            int length;

            while ((length = inputStream.read(buffer)) > 0) {
                outputStream.write(buffer, 0, length);
            }
        } finally {
            inputStream.close();
            outputStream.close();
        }
    }

    private File getFile(String name) {
        String file = getClass().getResource(name).getFile();
        return new File(file);
    }

    private XSSFCell getCell(XSSFSheet sheet, String cellCoordinate) {
        CellReference d4 = new CellReference(cellCoordinate);
        XSSFRow row = sheet.getRow(d4.getRow());
        return row.getCell(d4.getCol());
    }
}
