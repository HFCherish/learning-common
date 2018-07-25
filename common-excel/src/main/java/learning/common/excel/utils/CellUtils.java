package learning.common.excel.utils;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;

/**
 * @author hf_cherish
 * @date 2018/7/25
 */
public class CellUtils {
    public static XSSFCell getCellByName(XSSFWorkbook workbook, String cellName) {
        XSSFName cellNameObj = workbook.getName(cellName);
        AreaReference areaReference = new AreaReference(cellNameObj.getRefersToFormula(), SpreadsheetVersion.EXCEL2007);

        CellReference firstCell = areaReference.getFirstCell();
        XSSFSheet sheet = workbook.getSheet(firstCell.getSheetName());
        return getCell(sheet, firstCell);
    }

    public static XSSFCell getCell(XSSFSheet sheet, String cellCoordinate) {
        CellReference cellReference = new CellReference(cellCoordinate);
        return getCell(sheet, cellReference);
    }

    public static XSSFCell getCell(XSSFSheet sheet, CellReference cellReference) {
        XSSFRow row = sheet.getRow(cellReference.getRow());
        return row.getCell(cellReference.getCol());
    }
}
