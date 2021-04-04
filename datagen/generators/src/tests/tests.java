import org.junit.*;

import static org.junit.Assert.*;

import org.apache.poi.ss.util.CellReference;

import testutils.ExcelTestingUtils;
import testutils.CalcTestingUtils;
import testutils.TestingUtils;

import java.util.function.BiFunction;
import java.util.Random;
import java.io.File;

import creator.Creatable;
import vlookups.*;
import sums.*;

public class tests {

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

    @Before
    public void setInputSize () {
        this.rows = 1 + RANDOM.nextInt(EXCLUSIVE_UPPER_BOUND);
        this.cols = 1 + RANDOM.nextInt(EXCLUSIVE_UPPER_BOUND);
    }

    @BeforeClass
    public static void createDirectories () {
        TestingUtils.createDirectories();
    }

    @AfterClass
    public static void deleteDirectories () {
        TestingUtils.deleteDirectories();
    }

    private void testEmpty (Creatable creatable, Random rand, BiFunction<Integer, Integer, String> getExpectedFormula) {
        ExcelTestingUtils.integrationTest(creatable, 0, 0, null, getExpectedFormula);
        ExcelTestingUtils.integrationTest(creatable, 0, 0, rand, getExpectedFormula);
        CalcTestingUtils.integrationTest(creatable, 0, 0, null, getExpectedFormula);
        CalcTestingUtils.integrationTest(creatable, 0, 0, rand, getExpectedFormula);
    }

    private void runAllIntegrationTests (Creatable creatable, int rows, int cols, Random rand, BiFunction<Integer, Integer, String> getExpectedFormula) {
        this.testEmpty(creatable, rand, getExpectedFormula);
        ExcelTestingUtils.integrationTest(creatable, rows, cols, null, getExpectedFormula);
        ExcelTestingUtils.integrationTest(creatable, rows, cols, rand, getExpectedFormula);
        CalcTestingUtils.integrationTest(creatable, rows, cols, null, getExpectedFormula);
        CalcTestingUtils.integrationTest(creatable, rows, cols, rand, getExpectedFormula);
    }

    @Test
    public void testCompleteBipartiteSum () {
        this.runAllIntegrationTests(new CompleteBipartiteSum(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(this.cols - 1), this.rows)
        );
    }

    @Test
    public void testRunningSum () {
        this.runAllIntegrationTests(new RunningSum(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(this.cols - 1), currRowIdx + 1)
        );
    }

    @Test
    public void testRunningSumAvg () {
        this.runAllIntegrationTests(new RunningSumAvg(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) -> 
            String.format("%s(A1:%s%d)"
                , currRowIdx % 2 == 0 ? "SUM" : "AVERAGE"
                , CellReference.convertNumToColString(this.cols - 1)
                , currRowIdx + 1
            )
        );
    }

    @Test
    public void testSingleCellSum () {
        this.runAllIntegrationTests(new SingleCellSum(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            return String.format("SUM(%s%d:%s%d)", col, currRowIdx + 1, col, currRowIdx + 1);
        });
    }

    @Test
    public void testSingleCellSumAvg () {
        this.runAllIntegrationTests(new SingleCellSumAvg(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            String frm = currRowIdx % 2 == 0 ? "SUM" : "AVERAGE";
            return String.format("%s(%s%d:%s%d)", frm, col, currRowIdx + 1, col, currRowIdx + 1);
        });
    }

    @Test
    public void testRunningVlookup () {
        this.runAllIntegrationTests(new RunningVlookup(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) 
            -> String.format("VLOOKUP(A%d, A1:%s%d, %d, FALSE)"
                , currRowIdx + 1
                , CellReference.convertNumToColString(this.cols - 1)
                , currRowIdx + 1
                , this.cols
            )
        );
    }

    @Test
    public void testSingleCellVlookup () {
        this.runAllIntegrationTests(new SingleCellVlookup(), this.rows, this.cols, new Random(42L), (currRowIdx, currColIdx) -> { 
            String col = CellReference.convertNumToColString(currColIdx);
            int row = currRowIdx + 1;
            return String.format("VLOOKUP(%s%d, %s%d:%s%d, 1, FALSE)", col, row, col, row, col, row);
        });
    }

    @Test
    public void testCreateCalcFiles () {
        Creatable[] creatables = { 
            new CompleteBipartiteSum(),
            new RunningSum(),
            new RunningSumAvg(),
            new SingleCellSum(),
            new SingleCellSumAvg(),
            new RunningVlookup(),
            new SingleCellVlookup() 
        };

        /** Make sure RNG path works */
        for (Creatable c : creatables) {
            File[] files = CalcTestingUtils.createCalcFiles(c, rows, cols, new Random(42L));            
            assertTrue(TestingUtils.allFilesExist(files));
            TestingUtils.deleteFiles();
        }

        /** Make sure non-RNG path works */
        for (Creatable c : creatables) {
            File[] files = CalcTestingUtils.createCalcFiles(c, rows, cols, null);            
            assertTrue(TestingUtils.allFilesExist(files));
            TestingUtils.deleteFiles();
        }
    }
}
