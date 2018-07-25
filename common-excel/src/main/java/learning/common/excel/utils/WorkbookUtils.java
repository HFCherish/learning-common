package learning.common.excel.utils;

import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @author hf_cherish
 * @date 2018/7/25
 */
public class WorkbookUtils {
    public static void writeExcel(XSSFWorkbook workbook, File target) throws IOException {
        writeExcel(workbook, target, true);
    }

    public static void writeExcel(XSSFWorkbook workbook, File target, boolean refreshForFormulaCells) throws IOException {
        if (refreshForFormulaCells) {
            XSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
        }
        workbook.write(new FileOutputStream(target));
    }
}
