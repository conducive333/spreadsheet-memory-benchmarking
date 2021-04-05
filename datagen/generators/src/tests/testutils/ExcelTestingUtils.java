package testutils;

import static org.junit.Assert.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.function.BiFunction;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Iterator;
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
     * For the formula-value workbook, this test makes sure that the number of sheets equals 1, 
     * the number of rows and columns match the input rows and columns, the first `cols` columns 
     * contain values that are between [0, rows * cols), and the next `cols` columns have the 
     * expected formula structure. For the value-only workbook, this test checks the number of 
     * sheets, rows, and columns just as before. It also verifies that each cell has a value that
     * matches its corresponding evaluated result in the formula-value workbook.
     * 
     * @param creatable
     * @param rows
     * @param cols
     * @param rand
     * @param getExpectedFormula
     */
    public static void integrationTest (Creatable creatable, int rows, int cols, Random rand, BiFunction<Integer, Integer, String> getExpectedFormula) {
        File[] files = createExcelFiles(creatable, rows, cols, rand);
        assertTrue(allFilesExist(files));
        checkExcelFVWorkbook(files[1], rows, cols, rand, getExpectedFormula);
        checkExcelVOWorkbook(files[0], files[1], rows, cols, rand);
        TestingUtils.deleteFiles();
    }

    private static void checkExcelFVWorkbook (File formulaValueFile, int rows, int cols, Random rand, BiFunction<Integer, Integer, String> getExpectedFormula) {
        try (XSSFWorkbook fWorkbook = new XSSFWorkbook(formulaValueFile)) {
            assertEquals(1, fWorkbook.getNumberOfSheets());
            int actualRowCount = 0;
            for (Row row : fWorkbook.getSheetAt(0)) {
                assertEquals(cols * 2, row.getPhysicalNumberOfCells());
                for (int c = 0; c < cols; c++) {
                    checkExcelNumericCell(row.getCell(c), rand, 0, rows * cols);
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

    private static void checkExcelVOWorkbook (File valueOnlyFile, File formulaValueFile, int rows, int cols, Random rand) {
        try (XSSFWorkbook vWorkbook = new XSSFWorkbook(valueOnlyFile); XSSFWorkbook fWorkbook = new XSSFWorkbook(formulaValueFile)) {
            assertEquals(1, vWorkbook.getNumberOfSheets());
            Iterator<Row> formulaRows = fWorkbook.getSheetAt(0).iterator();
            FormulaEvaluator evaluator = fWorkbook.getCreationHelper().createFormulaEvaluator();
            for (Row vRow : vWorkbook.getSheetAt(0)) {
                if (formulaRows.hasNext()) {
                    Row fRow = formulaRows.next();
                    assertEquals(cols * 2, vRow.getPhysicalNumberOfCells());
                    for (int c = 0; c < cols; c++) {
                        checkIfCellsAreEqual(evaluator, vRow.getCell(c)        , fRow.getCell(c));
                        checkIfCellsAreEqual(evaluator, vRow.getCell(c + cols) , fRow.getCell(c + cols));
                    }
                } else {
                    fail("Value-only sheet has more rows than formula-value sheet.");
                }
            }
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
            fail("Could not open file.");
        }
    }

    private static double tryToGetNumericValue (Cell cell) {
        assertEquals("Numeric value not found at " + cell.getAddress(), cell.getCellType(), CellType.NUMERIC);
        try { return cell.getNumericCellValue(); }
        catch (NumberFormatException e) { fail("Cell " + cell.getAddress() + " does not contain a parsable double: " + cell.getStringCellValue()); }
        return Double.NaN;
    }

    private static void checkIfCellsAreEqual (FormulaEvaluator evaluator, Cell vCell, Cell fCell) {
        double vVal = tryToGetNumericValue(vCell);
        double fVal = evaluator.evaluate(fCell).getNumberValue();
        assertTrue(String.format("Value-only cell (%f) does not match formula-value cell (%f)",  vVal, fVal), vVal == fVal);
    }

    private static void checkExcelNumericCell (Cell cell, Random rand, double inclusiveLowerBound, double exclusiveUpperBound) {
        double v = tryToGetNumericValue(cell);
        if (rand != null) {
            assertTrue(v + " is not in [" + inclusiveLowerBound + ", " + exclusiveUpperBound + ")", v >= inclusiveLowerBound && v < exclusiveUpperBound);
        }
    }

    private static void checkExcelFormulaCell (Cell cell, String expectedFormula) {
        assertEquals("Formula not found at " + cell.getAddress(), cell.getCellType(), CellType.FORMULA);
        assertEquals(expectedFormula, cell.getCellFormula());
    }

}
