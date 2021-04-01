import org.junit.*;
import static org.junit.Assert.*;

import testutils.ExcelTestingUtils;
import testutils.CalcTestingUtils;
import testutils.TestingUtils;

import sums.CompleteBipartiteSum;
import sums.SingleCellSum;
import sums.RunningSum;

import org.apache.poi.ss.util.CellReference;

import java.util.Random;

public class SumsTests {

    /**
     * This is for generating random spreadsheet sizes.
     */
    private static final Random RANDOM = new Random(42L);

    /** 
     * The number of rows and columns to generate for a 
     * test spreadsheet is chosen between 1 (inclusive)
     * and the number below (exclusive).
     */
    private static final int EXCLUSIVE_UPPER_BOUND = 5;

    /**
     * The number of rows and columns to use. These are 
     * randomly assigned before each test case.
     */
    private int rows, cols;

    @BeforeClass
    public static void createDirectories () {
        if (!TestingUtils.V_FOLDER.exists()) {
            TestingUtils.V_FOLDER.mkdirs();
        } else { 
            TestingUtils.deleteAllFilesInDirectory(TestingUtils.V_FOLDER); 
        }
        if (!TestingUtils.F_FOLDER.exists()) {
            TestingUtils.F_FOLDER.mkdirs();
        } else { 
            TestingUtils.deleteAllFilesInDirectory(TestingUtils.F_FOLDER); 
        }
    }

    @Before
    public void setInputSize () {
        this.rows = 1 + RANDOM.nextInt(EXCLUSIVE_UPPER_BOUND);
        this.cols = 1 + RANDOM.nextInt(EXCLUSIVE_UPPER_BOUND);
    }

    @After
    public void deleteFiles () {
        TestingUtils.deleteAllFilesInDirectory(TestingUtils.V_FOLDER);
        TestingUtils.deleteAllFilesInDirectory(TestingUtils.F_FOLDER);
    }

    @AfterClass
    public static void deleteDirectories () {
        TestingUtils.V_FOLDER.delete();
        TestingUtils.F_FOLDER.delete();
        TestingUtils.TEMP_DIR.delete();
    }

    /** Complete Bipartite Sum Tests */

    @Test
    public void testExcelCompleteBipartiteSumRand () {
        ExcelTestingUtils.integrationTest(new CompleteBipartiteSum(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(this.cols - 1), this.rows)
        );
    }

    @Test
    public void testExcelCompleteBipartiteSumNoRand () {
        ExcelTestingUtils.integrationTest(new CompleteBipartiteSum(), this.rows, this.cols, null, (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(this.cols - 1), this.rows)
        );
    }

    @Test
    public void testCalcCompleteBipartiteSumRand () {
        CalcTestingUtils.integrationTest(new CompleteBipartiteSum(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(this.cols - 1), this.rows)
        );
    }

    @Test
    public void testCalcCompleteBipartiteSumNoRand () {
        CalcTestingUtils.integrationTest(new CompleteBipartiteSum(), this.rows, this.cols, null, (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(this.cols - 1), this.rows)
        );
    }

    /** Single Cell Sum Tests */

    @Test
    public void testExcelSingleCellSumRand () {
        ExcelTestingUtils.integrationTest(new SingleCellSum(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            return String.format("SUM(%s%d:%s%d)", col, currRowIdx + 1, col, currRowIdx + 1);
        });
    }

    @Test
    public void testExcelSingleCellSumNoRand () {
        ExcelTestingUtils.integrationTest(new SingleCellSum(), this.rows, this.cols, null, (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            return String.format("SUM(%s%d:%s%d)", col, currRowIdx + 1, col, currRowIdx + 1);
        });
    }

    @Test
    public void testCalcSingleCellSumRand () {
        CalcTestingUtils.integrationTest(new SingleCellSum(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            return String.format("SUM(%s%d:%s%d)", col, currRowIdx + 1, col, currRowIdx + 1);
        });
    }

    @Test
    public void testCalcSingleCellSumNoRand () {
        CalcTestingUtils.integrationTest(new SingleCellSum(), this.rows, this.cols, null, (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            return String.format("SUM(%s%d:%s%d)", col, currRowIdx + 1, col, currRowIdx + 1);
        });
    }

    /** Running Sum Tests */

    @Test
    public void testExcelRunningSumRand () {
        ExcelTestingUtils.integrationTest(new RunningSum(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), currRowIdx + 1)
        );
    }

    @Test
    public void testExcelRunningSumNoRand () {
        ExcelTestingUtils.integrationTest(new RunningSum(), this.rows, this.cols, null, (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), currRowIdx + 1)
        );
    }

    @Test
    public void testCalcRunningSumRand () {
        CalcTestingUtils.integrationTest(new RunningSum(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), currRowIdx + 1)
        );
    }

    @Test
    public void testCalcRunningSumNoRand () {
        CalcTestingUtils.integrationTest(new RunningSum(), this.rows, this.cols, null, (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), currRowIdx + 1)
        );
    }

    /** Miscellaneous Calc Tests */

    @Test
    public void testCreateCalcFilesRand () {
        long seed = 42L;
        assertTrue(TestingUtils.allFilesExist(CalcTestingUtils.createCalcFiles(new CompleteBipartiteSum(), rows, cols, new Random(seed))));
        assertTrue(TestingUtils.allFilesExist(CalcTestingUtils.createCalcFiles(new SingleCellSum(), rows, cols, new Random(seed))));
        assertTrue(TestingUtils.allFilesExist(CalcTestingUtils.createCalcFiles(new RunningSum(), rows, cols, new Random(seed))));
    }

    @Test
    public void testCreateCalcFilesNoRand () {
        assertTrue(TestingUtils.allFilesExist(CalcTestingUtils.createCalcFiles(new CompleteBipartiteSum(), rows, cols, null)));
        assertTrue(TestingUtils.allFilesExist(CalcTestingUtils.createCalcFiles(new SingleCellSum(), rows, cols, null)));
        assertTrue(TestingUtils.allFilesExist(CalcTestingUtils.createCalcFiles(new RunningSum(), rows, cols, null)));
    }
}
