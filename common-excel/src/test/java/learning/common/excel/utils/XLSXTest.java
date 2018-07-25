package learning.common.excel.utils;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;

import static learning.common.excel.utils.CellUtils.getCell;
import static learning.common.excel.utils.CellUtils.getCellByName;
import static learning.common.excel.utils.WorkbookUtils.writeExcel;
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
        XSSFWorkbook workbook = new XSSFWorkbook(getFile("/test.xlsx"));

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
        writeExcel(workbook, getFile("/update.xlsx"));

//        check data updated
        XSSFWorkbook updateWorkBook = new XSSFWorkbook(getFile("/update.xlsx"));
        assertThat(getCell(updateWorkBook.getSheet(sheetName), "D4").getNumericCellValue(), closeTo(200, 0.1));
        assertThat(getCell(updateWorkBook.getSheet(sheetName), "D13").getNumericCellValue(), closeTo(23500, 0.1));
    }

    @Test
    void should_able_to_get_cell_by_name() throws IOException, InvalidFormatException {
        XSSFWorkbook workbook = new XSSFWorkbook(getFile("/test.xlsx"));

        XSSFCell cell = getCellByName(workbook, "test_name");

        assertThat(cell.getStringCellValue(), is("TV set"));
    }

    private File getFile(String name) {
        String file = getClass().getResource(name).getFile();
        return new File(file);
    }

}
