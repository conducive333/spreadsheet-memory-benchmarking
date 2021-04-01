package testutils;

import static org.junit.Assert.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.function.BiFunction;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Random;
import java.io.File;

import creator.Creatable;
import creator.Creator;

public class ExcelTestingUtils extends TestingUtils {
    
    /**
     * Creates a formula-value spreadsheet and its corresponding value-only spreadsheet
     * using `creatable`. The spreadsheet is saved as a .xlsx file.
     * 
     * @param creatable
     * @param rows
     * @param cols
     * @param rand
     * @return An array, ARR, of two files. ARR[0] is the value-only file and ARR[1] is 
     * the formula-value file.
     */
    public static File[] createExcelFiles (Creatable creatable, int rows, int cols, Random rand) {
        Creator.createExcelSheet(creatable, TestingUtils.F_FOLDER.toString(), TestingUtils.V_FOLDER.toString(), rows, cols, rand);
        return new File[] {
            Path.of(TestingUtils.V_FOLDER.toString(), "vo-" + rows + ".xlsx").toFile(),
            Path.of(TestingUtils.F_FOLDER.toString(), "fv-" + rows + ".xlsx").toFile()
        };
    }

    /**
     * Checks the number of sheets, number of rows, and number of cells in each row.
     * Verifies that the first `cols` columns contain values, (but does not necessarily
     * check the values themselves), and verifies that the next `cols` columns have the 
     * expected formula structure.
     * 
     * @param formulaValueFile
     * @param rows
     * @param cols
     * @param rand
     * @param getExpectedFormula
     */
    public static void integrationTest (Creatable creatable, int rows, int cols, Random rand, BiFunction<Integer, Integer, String> getExpectedFormula) {
        File[] files = createExcelFiles(creatable, rows, cols, rand);
        assertTrue(allFilesExist(files));
        checkExcelVOWorkbook(files[0], rows, cols);
        checkExcelFVWorkbook(files[1], rows, cols, getExpectedFormula);
    }

    private static void checkExcelVOWorkbook (File valueOnlyFile, int rows, int cols) {
        try (XSSFWorkbook vWorkbook = new XSSFWorkbook(valueOnlyFile)) {
            assertEquals(1, vWorkbook.getNumberOfSheets());
            int actualRowCount = 0;
            for (Row row : vWorkbook.getSheetAt(0)) {
                assertEquals(cols * 2, row.getPhysicalNumberOfCells());
                for (Cell cell : row) { checkExcelNumericCell(cell); }
                actualRowCount++;
            }
            assertEquals(rows, actualRowCount);
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
            fail("Could not open file.");
        }
    }

    private static void checkExcelFVWorkbook (File formulaValueFile, int rows, int cols, BiFunction<Integer, Integer, String> getExpectedFormula) {
        try (XSSFWorkbook fWorkbook = new XSSFWorkbook(formulaValueFile)) {
            assertEquals(1, fWorkbook.getNumberOfSheets());
            int actualRowCount = 0;
            for (Row row : fWorkbook.getSheetAt(0)) {
                assertEquals(cols * 2, row.getPhysicalNumberOfCells());
                for (int c = 0; c < cols; c++) {
                    checkExcelNumericCell(row.getCell(c));
                    checkExcelFormulaCell(row.getCell(c + cols), getExpectedFormula.apply(row.getRowNum(), c));
                }
                actualRowCount++;
            }
            assertEquals(rows, actualRowCount);
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
            fail("Could not open file.");
        }
    }

    private static void checkExcelNumericCell (Cell cell) {
        try { cell.getNumericCellValue(); }
        catch (IllegalStateException e) { fail("Cell is a string not a double!"); }
        catch (NumberFormatException e) { fail("Cell did not contain a parsable double!"); }
    }

    private static void checkExcelFormulaCell (Cell cell, String expectedFormula) {
        try { assertEquals(expectedFormula, cell.getCellFormula()); } 
        catch (IllegalStateException e) { fail(String.format("Formula not found at %s", cell.getAddress().toString())); }
    }

}
