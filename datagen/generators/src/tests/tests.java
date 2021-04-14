import org.junit.*;

import static org.junit.Assert.*;

import org.apache.poi.ss.util.CellReference;

import testutils.ExcelTestingUtils;
import testutils.CalcTestingUtils;
import testutils.TestingUtils;

import java.util.function.BiFunction;
import java.util.OptionalLong;
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
    private static final int EXCLUSIVE_UPPER_BOUND = 10;

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

    private void testEmpty (Creatable creatable, long seed, BiFunction<Integer, Integer, String> getExpectedExcelFormula, BiFunction<Integer, Integer, String> getExpectedCalcFormula) {
        ExcelTestingUtils.integrationTest(creatable, 0, 0, 0, 0, OptionalLong.empty(), getExpectedExcelFormula);
        ExcelTestingUtils.integrationTest(creatable, 0, 0, 0, 0, OptionalLong.of(seed), getExpectedExcelFormula);
        CalcTestingUtils.integrationTest(creatable, 0, 0, 0, 0, OptionalLong.empty(), getExpectedCalcFormula);
        CalcTestingUtils.integrationTest(creatable, 0, 0, 0, 0, OptionalLong.of(seed), getExpectedCalcFormula);
    }

    private void runAllIntegrationTests (Creatable creatable, int rows, int expectedRows, int cols, int expectedCols, long seed, BiFunction<Integer, Integer, String> getExpectedExcelFormula, BiFunction<Integer, Integer, String> getExpectedCalcFormula) {
        this.testEmpty(creatable, seed, getExpectedExcelFormula, getExpectedCalcFormula);
        ExcelTestingUtils.integrationTest(creatable, rows, expectedRows, cols, expectedCols, OptionalLong.empty(), getExpectedExcelFormula);
        ExcelTestingUtils.integrationTest(creatable, rows, expectedRows, cols, expectedCols, OptionalLong.of(seed), getExpectedExcelFormula);
        CalcTestingUtils.integrationTest(creatable, rows, expectedRows, cols, expectedCols, OptionalLong.empty(), getExpectedCalcFormula);
        CalcTestingUtils.integrationTest(creatable, rows, expectedRows, cols, expectedCols, OptionalLong.of(seed), getExpectedCalcFormula);
    }

    private void runAllIntegrationTests (Creatable creatable, int rows, int expectedRows, int cols, int expectedCols, long seed, BiFunction<Integer, Integer, String> getExpectedFormula) {
        this.runAllIntegrationTests(creatable, rows, expectedRows, cols, expectedCols, seed, getExpectedFormula, getExpectedFormula);
    }

    @Test
    public void testCreateCalcFiles () {
        Creatable[] creatables = { 
            new CompleteBipartiteSum(),
            new MixedRangeSum(),
            new NoEdgeSum(),
            new OverlappingSum(),
            new RunningSum(),
            new CompleteBipartiteSumWithConstant(),
            new SingleCellSum(),
            new CompleteBipartiteVlookup(),
            new SingleCellVlookup() 
        };

        /** Check if RNG path works */
        for (Creatable c : creatables) {
            File[] files = CalcTestingUtils.createCalcFiles(c, rows, cols, OptionalLong.of(42L));            
            assertTrue(TestingUtils.allFilesExist(files));
            TestingUtils.deleteFiles();
        }

        /** Check if non-RNG path works */
        for (Creatable c : creatables) {
            File[] files = CalcTestingUtils.createCalcFiles(c, rows, cols, OptionalLong.empty());            
            assertTrue(TestingUtils.allFilesExist(files));
            TestingUtils.deleteFiles();
        }
    }

    @Test
    public void testCompleteBipartiteSum () {
        this.runAllIntegrationTests(new CompleteBipartiteSum(), this.rows, this.rows, this.cols, this.cols * 2, 42L, (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(this.cols - 1), this.rows)
        );
    }

    @Test
    public void testCompleteBipartiteSumWithConstant() {
        this.runAllIntegrationTests(new CompleteBipartiteSumWithConstant(), this.rows, this.rows, this.cols, this.cols * 2, 42L, (currRowIdx, currColIdx)
            -> String.format("SUM(A1:%s%d) + %d", CellReference.convertNumToColString(this.cols - 1), this.rows, currRowIdx + 1)
        );
    }

    @Test
    public void testMixedRangeSum () {
        this.runAllIntegrationTests(new MixedRangeSum(), this.rows, this.rows, this.cols, this.cols * 2, 42L, (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            int row = currRowIdx + 1;
            return String.format("SUM(%s%d:%s%d) + SUM(A1:%s%d)", col, row, col, row
                , CellReference.convertNumToColString(this.cols - 1)
                , this.rows
            );
        });
    }

    @Test
    public void testNoEdgeSum () {
        
        final long seed = 42;
        
        // Same as FILL_VALUE in BaseSum
        final double fillValue = 1.0;
        
        // These will be refreshed after each integration test
        Random[]    rand = { null };
        int[]       rows = { 0 };

        this.runAllIntegrationTests(new NoEdgeSum(), this.rows, this.rows, this.cols, this.cols * 2, seed, (currRowIdx, currColIdx) -> {
            
            // This will be true once we perform one integration test.
            // Integration tests alternate between non-random and random
            // sheets, so we need to reset these values after each test.
            if (rows[0] == this.rows * this.cols) {
                rand[0] = (rand[0] == null ? new Random(seed) : null);
                rows[0] = 0;
            }

            // Return the correct formula for the corresponding integration
            // test.
            rows[0]++;
            if (rand[0] != null) {
                return String.format("SUM(%f)", (double) rand[0].nextInt(this.rows * this.cols));
            } else {
                return String.format("SUM(%f)", fillValue);
            }
            
        });
    }

    @Test
    public void testOverlappingSum () {
        final int windowSize = 2;
        this.runAllIntegrationTests(new OverlappingSum(), this.rows, this.rows, this.cols, this.cols * 2, 42L, (currRowIdx, currColIdx) 
            -> String.format("SUM(%s%d:%s%d)"
                , CellReference.convertNumToColString(currColIdx)
                , currRowIdx + 1
                , CellReference.convertNumToColString(currColIdx)
                , currRowIdx + windowSize
            )
        );
    }

    @Test
    public void testRunningSum () {
        this.runAllIntegrationTests(new RunningSum(), this.rows, this.rows, this.cols, this.cols * 2, 42L, (currRowIdx, currColIdx) 
            -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(this.cols - 1), currRowIdx + 1)
        );
    }

    @Test
    public void testSingleCellSum () {
        this.runAllIntegrationTests(new SingleCellSum(), this.rows, this.rows, this.cols, this.cols * 2, 42L, (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            return String.format("SUM(%s%d:%s%d)", col, currRowIdx + 1, col, currRowIdx + 1);
        });
    }

    @Test
    public void testCompleteBipartiteVlookup () {
        this.runAllIntegrationTests(new CompleteBipartiteVlookup(), this.rows, this.rows, 1, 3, 42L
            , (currRowIdx, currColIdx) -> String.format("VLOOKUP(C%d, A1:A%d, 1, FALSE)", currRowIdx + 1, this.rows)
            , (currRowIdx, currColIdx) -> String.format("VLOOKUP(C%d; A1:A%d; 1; 0)"    , currRowIdx + 1, this.rows)
        );
    }

    @Test
    public void testSameCellVlookup () {
        // Sample testing code:
        this.runAllIntegrationTests(new SameCellVlookup(), this.rows, this.rows, 1, 3, 42L
            , (currRowIdx, currColIdx) -> String.format("VLOOKUP(C%d, A1:A1, 1, FALSE)", currRowIdx + 1)
            , (currRowIdx, currColIdx) -> String.format("VLOOKUP(C%d; A1:A1; 1; 0)"    , currRowIdx + 1)
        );
    }

    @Test
    public void testSingleCellVlookup () {
        this.runAllIntegrationTests(new SingleCellVlookup(), this.rows, this.rows, 1, 3, 42L
            , (currRowIdx, currColIdx) -> String.format("VLOOKUP(C%1$d, A%1$d:A%1$d, 1, FALSE)" , currRowIdx + 1)
            , (currRowIdx, currColIdx) -> String.format("VLOOKUP(C%1$d; A%1$d:A%1$d; 1; 0)"     , currRowIdx + 1)
        );
    }

}
