package testutils;

import static org.junit.Assert.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.function.BiFunction;
import java.util.OptionalLong;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Iterator;
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
     * @param seed
     * @return An array, ARR, of two files. ARR[0] is the value-only file and ARR[1] is 
     * the formula-value file.
     */
    public static File[] createExcelFiles (Creatable creatable, int rows, int cols, OptionalLong seed) {
        Creator.createExcelSheet(creatable, TestingUtils.F_FOLDER.toString(), TestingUtils.V_FOLDER.toString(), rows, cols, seed);
        return new File[] {
            Path.of(TestingUtils.V_FOLDER.toString(), "vo-" + rows + ".xlsx").toFile(),
            Path.of(TestingUtils.F_FOLDER.toString(), "fv-" + rows + ".xlsx").toFile()
        };
    }

    /**
     * For the formula-value workbook, this test makes sure that the number of sheets equals
     * 1, the number of rows and columns match the expected rows and columns, the first `cols` 
     * columns contain values that are between [0, rows * cols), and the next `cols` columns 
     * have the expected formula structure. For the value-only workbook, this test checks the
     * number of sheets, rows, and columns just as before. It also verifies that each cell has
     * a value that matches its corresponding evaluated result in the formula-value workbook.
     * 
     * @param creatable
     * @param rows
     * @param expectedRows
     * @param cols
     * @param expectedCols
     * @param seed
     * @param getExpectedFormula
     */
    public static void integrationTest (Creatable creatable, int rows, int expectedRows, int cols, int expectedCols, OptionalLong seed, BiFunction<Integer, Integer, String> getExpectedFormula) {
        File[] files = createExcelFiles(creatable, rows, cols, seed);
        assertTrue(allFilesExist(files));
        checkExcelFVWorkbook(files[1], rows, expectedRows, cols, expectedCols, seed, getExpectedFormula);
        checkExcelVOWorkbook(files[0], files[1], cols);
        TestingUtils.deleteFiles();
    }

    private static void checkExcelFVWorkbook (File formulaValueFile, int rows, int expectedRows, int cols, int expectedCols, OptionalLong seed, BiFunction<Integer, Integer, String> getExpectedFormula) {
        try (XSSFWorkbook fWorkbook = new XSSFWorkbook(formulaValueFile)) {
            assertEquals(1, fWorkbook.getNumberOfSheets());
            int actualRowCount = 0;
            for (Row row : fWorkbook.getSheetAt(0)) {
                assertEquals(expectedCols, row.getPhysicalNumberOfCells());
                for (int c = 0; c < cols; c++) {
                    assertCellIsNotNull(row, c);
                    assertCellIsNotNull(row, c + cols);
                    checkExcelNumericCell(row.getCell(c), seed, 0, rows * cols);
                    checkExcelFormulaCell(row.getCell(c + cols), getExpectedFormula.apply(row.getRowNum(), c));
                }
                actualRowCount++;
            }
            assertEquals(expectedRows, actualRowCount);
        } catch (InvalidFormatException | IOException e) {
            e.printStackTrace();
            fail("Could not open file.");
        }
    }

    private static void checkExcelVOWorkbook (File valueOnlyFile, File formulaValueFile, int cols) {
        try (XSSFWorkbook vWorkbook = new XSSFWorkbook(valueOnlyFile); XSSFWorkbook fWorkbook = new XSSFWorkbook(formulaValueFile)) {
            assertEquals(1, vWorkbook.getNumberOfSheets());
            Iterator<Row> formulaRows = fWorkbook.getSheetAt(0).iterator();
            FormulaEvaluator evaluator = fWorkbook.getCreationHelper().createFormulaEvaluator();
            for (Row vRow : vWorkbook.getSheetAt(0)) {
                if (formulaRows.hasNext()) {
                    Row fRow = formulaRows.next();
                    assertEquals(fRow.getPhysicalNumberOfCells(), vRow.getPhysicalNumberOfCells());
                    for (int c = 0; c < cols; c++) {
                        assertCellIsNotNull(vRow, c);
                        assertCellIsNotNull(vRow, c + cols);
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

    private static void assertCellIsNotNull (Row row, int col) {
        assertNotNull(
            String.format("Found null cell at row index %d and column index %d"
                , row.getRowNum()
                , col
            )
            , row.getCell(col)
        );
    }

    private static double tryToGetNumericValue (Cell cell) {
        assertEquals(String.format("Numeric value not found at %s", cell.getAddress()), cell.getCellType(), CellType.NUMERIC);
        try { return cell.getNumericCellValue(); }
        catch (NumberFormatException e) { fail(String.format("Cell %s does not contain a parsable double: %s",  cell.getAddress(), cell.getStringCellValue())); }
        return Double.NaN;
    }

    private static void checkIfCellsAreEqual (FormulaEvaluator evaluator, Cell vCell, Cell fCell) {
        CellType ct = evaluator.evaluateFormulaCell(fCell);
        if (ct != CellType.ERROR) {
            double vVal = tryToGetNumericValue(vCell);
            double fVal = evaluator.evaluate(fCell).getNumberValue();
            assertTrue(String.format("Value-only cell (%s=%f) does not match formula-value cell (%s=%f)",  vCell.getAddress(), vVal, fCell.getAddress(), fVal), vVal == fVal);
        } else {
            if (fCell.getCellType() == CellType.FORMULA) {
                System.out.println("WARNING: POI could not evaluate " + fCell.getCellFormula() + ". It is recommended to check this test case's sheet creator manually.");
            } else {
                fail(String.format("Found unexpected item at %s: %s", fCell.getAddress(), evaluator.evaluate(fCell).toString()));
            }
        }
    }

    private static void checkExcelNumericCell (Cell cell, OptionalLong seed, double inclusiveLowerBound, double exclusiveUpperBound) {
        double v = tryToGetNumericValue(cell);
        if (seed.isPresent()) {
            assertTrue(String.format("%f is not in [%f, %f)", v, inclusiveLowerBound, exclusiveUpperBound), v >= inclusiveLowerBound && v < exclusiveUpperBound);
        }
    }

    private static void checkExcelFormulaCell (Cell cell, String expectedFormula) {
        assertEquals("Formula not found at " + cell.getAddress(), cell.getCellType(), CellType.FORMULA);
        assertEquals(expectedFormula, cell.getCellFormula());
    }

}
