import org.junit.*;

import static org.junit.Assert.*;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import org.jopendocument.dom.spreadsheet.SpreadSheet;
import org.jopendocument.dom.spreadsheet.Sheet;

import java.util.function.BiFunction;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Random;
import java.io.File;

import creator.*;
import sums.*;

public class SumsTests {

    private static final Random RANDOM = new Random(42L);
    private static final File TEMP_DIR = Path.of("src", "tests", "temp").toFile();
    private static final File V_FOLDER = Path.of(TEMP_DIR.toString(), "value-only"   ).toFile();
    private static final File F_FOLDER = Path.of(TEMP_DIR.toString(), "formula-value").toFile();

    @BeforeClass
    public static void createDirectories () {
        if (!V_FOLDER.exists()) {
            V_FOLDER.mkdirs();
        } else { deleteAllFilesInDirectory(V_FOLDER); }
        if (!F_FOLDER.exists()) {
            F_FOLDER.mkdirs();
        } else { deleteAllFilesInDirectory(F_FOLDER); }
    }

    @AfterClass
    public static void deleteDirectories () {
        V_FOLDER.delete();
        F_FOLDER.delete();
        TEMP_DIR.delete();
    }

    @After
    public void deleteFiles () {
        deleteAllFilesInDirectory(V_FOLDER);
        deleteAllFilesInDirectory(F_FOLDER);
    }

    @Test
    public void testCalcFilesCreated () {
        int rows = 1 + RANDOM.nextInt(10), cols = 1 + RANDOM.nextInt(10);
        File[] files;
        files = createCalcFiles(new CompleteBipartiteSum(), rows, cols);
        assertTrue(allFilesExist(files));
        files = createCalcFiles(new SingleCellSum(), rows, cols);
        assertTrue(allFilesExist(files));
        files = createCalcFiles(new RunningSum(), rows, cols);
        assertTrue(allFilesExist(files));
    }

    @Test
    public void testExcelCompleteBipartiteSum () {
        int rows = 1 + RANDOM.nextInt(10), cols = 1 + RANDOM.nextInt(10);
        File[] files = createExcelFiles(new CompleteBipartiteSum(), rows, cols);
        assertTrue(allFilesExist(files));
        checkExcelVOWorkbook(files[0], rows, cols);
        checkExcelFVWorkbook(files[1], rows, cols, (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), rows)
        );
    }

    // @Test
    // public void testCalcCompleteBipartiteSum () {
    //     int rows = 10, cols = 1;
    //     File[] files = createExcelFiles(new CompleteBipartiteSum(), rows, cols);
    //     assertTrue(allFilesExist(files));
    //     checkCalcVOWorkbook(files[0], rows, cols);
    //     checkCalcFVWorkbook(files[0], rows, cols, (currRowIdx, currColIdx) 
    //         -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), rows)
    //     );
    //     deleteAllFiles(files);
    // }

    @Test
    public void testExcelSingleCellSumStructure () {
        int rows = 1 + RANDOM.nextInt(10), cols = 1 + RANDOM.nextInt(10);
        File[] files = createExcelFiles(new SingleCellSum(), rows, cols);
        assertTrue(allFilesExist(files));
        checkExcelVOWorkbook(files[0], rows, cols);
        checkExcelFVWorkbook(files[1], rows, cols, (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            return String.format("SUM(%s%d:%s%d)", col, currRowIdx + 1, col, currRowIdx + 1);
        });
    }

    @Test
    public void testExcelRunningSumStructure () {
        int rows = 1 + RANDOM.nextInt(10), cols = 1 + RANDOM.nextInt(10);
        File[] files = createExcelFiles(new RunningSum(), rows, cols);
        assertTrue(allFilesExist(files));
        checkExcelVOWorkbook(files[0], rows, cols);
        checkExcelFVWorkbook(files[1], rows, cols, (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), currRowIdx + 1)
        );
    }

    /** Miscellaneous Helpers */

    private static boolean allFilesExist (File[] files) {
        for (File f : files) {
            if (!f.exists()) { return false; }
        }
        return true;
    }

    private static void deleteAllFilesInDirectory (File dir) {
        for (File f: dir.listFiles()) {
            f.delete();
        }
    }

    /** EXCEL Helper Methods */

    private File[] createExcelFiles (Creatable c, int rows, int cols) {
        Creator.createExcelSheet(c, F_FOLDER.toString(), V_FOLDER.toString(), rows, cols, null);
        return new File[] {
            Path.of(V_FOLDER.toString(), "vo-" + rows + ".xlsx").toFile(),
            Path.of(F_FOLDER.toString(), "fv-" + rows + ".xlsx").toFile()
        };
    }

    private void checkExcelNumericCell (Cell cell) {
        try { cell.getNumericCellValue(); }
        catch (IllegalStateException e) { fail("Cell is a string not a double!"); }
        catch (NumberFormatException e) { fail("Cell did not contain a parsable double!"); }
    }

    private void checkExcelFormulaCell (Cell cell, String expectedFormula) {
        try { assertEquals(expectedFormula, cell.getCellFormula()); } 
        catch (IllegalStateException e) { fail(String.format("Formula not found at %s", cell.getAddress().toString())); }
    }

    private void checkExcelVOWorkbook (File valueOnlyFile, int rows, int cols) {
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
        }
    }

    private void checkExcelFVWorkbook (File formulaValueFile, int rows, int cols, BiFunction<Integer, Integer, String> getExpectedFormula) {
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
        }
    }

    /** CALC Helper Methods */

    // TODO: Finish testing calc

    private File[] createCalcFiles (Creatable c, int rows, int cols) {
        Creator.createCalcSheet(c, F_FOLDER.toString(), V_FOLDER.toString(), rows, cols, RANDOM);
        return new File[] {
            Path.of(V_FOLDER.toString(), "vo-" + rows + ".ods").toFile(),
            Path.of(F_FOLDER.toString(), "fv-" + rows + ".ods").toFile()
        };
    }

    private void checkCalcValueCell (Object cellValue) {
        assertTrue(cellValue instanceof Double);
    }

    private void checkCalcFormulaCell (String cellFormula, String expectedFormula) {
        assertEquals(cellFormula, cellFormula);
    }

    private void checkCalcVOWorkbook (File formulaValueFile, int rows, int cols) {
        // try {
        //     Sheet sheet = SpreadSheet.createFromFile(formulaValueFile).getSheet(0);
        //     assertEquals(cols * 2, sheet.getColumnCount());
        //     assertEquals(rows, sheet.getRowCount());
        //     for () {
        //         assertEquals(cols * 2, row.getPhysicalNumberOfCells());
        //         for (int c = 0; c < cols; c++) {
        //             checkCalcNumericCell(row.getCell(c));
        //             checkCalcFormulaCell(row.getCell(c + cols), getExpectedFormula.apply(row.getRowNum(), c + cols));
        //         }
        //     }            
        // } catch (IOException e) {
        //     e.printStackTrace();
        // }
    }

    private void checkCalcFVWorkbook (File formulaValueFile, int rows, int cols, BiFunction<Integer, Integer, String> getExpectedFormula) {
        // try (XSSFWorkbook fWorkbook = new XSSFWorkbook(formulaValueFile)) {
        //     assertEquals(1, fWorkbook.getNumberOfSheets());
        //     int actualRowCount = 0;
        //     for (Row row : fWorkbook.getSheetAt(0)) {
        //         assertEquals(cols * 2, row.getPhysicalNumberOfCells());
        //         for (int c = 0; c < cols; c++) {
        //             checkExcelNumericCell(row.getCell(c));
        //             checkExcelFormulaCell(row.getCell(c + cols), getExpectedFormula.apply(row.getRowNum(), c + cols));
        //         }
        //         actualRowCount++;
        //     }
        //     assertEquals(rows, actualRowCount);
        // } catch (InvalidFormatException | IOException e) {
        //     e.printStackTrace();
        // }
    }

}
