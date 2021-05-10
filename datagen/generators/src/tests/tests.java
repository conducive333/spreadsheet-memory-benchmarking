import org.junit.*;

import static org.junit.Assert.*;

import testutils.ExcelTestingUtils;
import testutils.CalcTestingUtils;
import testutils.TestingUtils;

import java.util.function.BiFunction;
import java.util.OptionalLong;
import java.util.AbstractMap;
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
     * Random values for each test sheet will be chosen 
     * from the interval [0, UPPR).
     */
    private static final int UPPR = 10;

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

    private void runEmptyIntegrationTests (Creatable creatable, long seed, BiFunction<Integer, Integer, String> getExpectedExcelFormula, BiFunction<Integer, Integer, String> getExpectedCalcFormula) {
        ExcelTestingUtils.integrationTest(creatable, 0, 0, 0, 0, OptionalLong.empty(), UPPR, getExpectedExcelFormula);
        ExcelTestingUtils.integrationTest(creatable, 0, 0, 0, 0, OptionalLong.of(seed), UPPR, getExpectedExcelFormula);
        CalcTestingUtils.integrationTest(creatable, 0, 0, 0, 0, OptionalLong.empty(), UPPR, getExpectedCalcFormula);
        CalcTestingUtils.integrationTest(creatable, 0, 0, 0, 0, OptionalLong.of(seed), UPPR, getExpectedCalcFormula);
    }

    private void runEmptyIntegrationTests (Creatable creatable, long seed, BiFunction<Integer, Integer, String> getExpectedFormula) {
        this.runEmptyIntegrationTests(creatable, seed, getExpectedFormula, getExpectedFormula);
    }

    private void runAllIntegrationTests (Creatable creatable, int rows, int expectedRows, int cols, int expectedCols, long seed, BiFunction<Integer, Integer, String> getExpectedExcelFormula, BiFunction<Integer, Integer, String> getExpectedCalcFormula) {
        ExcelTestingUtils.integrationTest(creatable, rows, expectedRows, cols, expectedCols, OptionalLong.empty(), UPPR, getExpectedExcelFormula);
        ExcelTestingUtils.integrationTest(creatable, rows, expectedRows, cols, expectedCols, OptionalLong.of(seed), UPPR, getExpectedExcelFormula);
        CalcTestingUtils.integrationTest(creatable, rows, expectedRows, cols, expectedCols, OptionalLong.empty(), UPPR, getExpectedCalcFormula);
        CalcTestingUtils.integrationTest(creatable, rows, expectedRows, cols, expectedCols, OptionalLong.of(seed), UPPR, getExpectedCalcFormula);
    }

    private void runAllIntegrationTests (Creatable creatable, int rows, int expectedRows, int cols, int expectedCols, long seed, BiFunction<Integer, Integer, String> getExpectedFormula) {
        this.runAllIntegrationTests(creatable, rows, expectedRows, cols, expectedCols, seed, getExpectedFormula, getExpectedFormula);
    }

    @Test
    public void testCreateCalcFiles () {

        sums.specialsums.BaseSpecialSum.setMaxRows(EXCLUSIVE_UPPER_BOUND * 2);

        Creatable[] creatables = new Creatable[] {
            new CompleteBipartiteSum                    (UPPR),
            new CompleteBipartiteSumWithConstant        (UPPR),
            new MixedRangeSum                           (UPPR),
            new NoEdgeSum                               (UPPR),
            new OverlappingSum                          (UPPR),
            new RunningSum                              (UPPR),
            new SingleCellSum                           (UPPR),
            new CompleteBipartiteVlookup                (UPPR),
            new SameCellVlookup                         (UPPR),
            new SingleCellVlookup                       (UPPR),
            new sums.specialsums.CompleteBipartiteSum   (UPPR),
            new sums.specialsums.MixedRangeSum          (UPPR),
            new sums.specialsums.NoEdgeSum              (UPPR),
            new sums.specialsums.OverlappingSum         (UPPR)
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
    public void testCompleteBipartiteSum() {
        final long seed = 42;
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getCompleteBipartiteSum(0, 0, UPPR);
        this.runEmptyIntegrationTests(pair.getKey(), seed, pair.getValue());
        pair = TestingUtils.getCompleteBipartiteSum(rows, cols, UPPR);
        this.runAllIntegrationTests(pair.getKey(), rows, rows, cols, cols * 2, seed, pair.getValue());
    }

    @Test
    public void testCompleteBipartiteSumWithConstant() {
        final long seed = 42;
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getCompleteBipartiteSumWithConstant(0, 0, UPPR);
        this.runEmptyIntegrationTests(pair.getKey(), seed, pair.getValue());
        pair = TestingUtils.getCompleteBipartiteSumWithConstant(rows, cols, UPPR);
        this.runAllIntegrationTests(pair.getKey(), rows, rows, cols, cols * 2, seed, pair.getValue());
    }

    @Test
    public void testMixedRangeSum () {
        final long seed = 42;
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getMixedRangeSum(0, 0, UPPR);
        this.runEmptyIntegrationTests(pair.getKey(), seed, pair.getValue());
        pair = TestingUtils.getMixedRangeSum(rows, cols, UPPR);
        this.runAllIntegrationTests(pair.getKey(), rows, rows, cols, cols * 2, seed, pair.getValue());
    }

    @Test
    public void testNoEdgeSum () {
        final long seed = 42;
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getNoEdgeSum(0, 0, UPPR, seed);
        this.runEmptyIntegrationTests(pair.getKey(), seed, pair.getValue());
        pair = TestingUtils.getNoEdgeSum(rows, cols, UPPR, seed);
        this.runAllIntegrationTests(pair.getKey(), rows, rows, cols, cols * 2, seed, pair.getValue());
    }

    @Test
    public void testOverlappingSum () {
        final long seed = 42;
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getOverlappingSum(0, 0, UPPR, OverlappingSum.WINDOW_SZE);
        this.runEmptyIntegrationTests(pair.getKey(), seed, pair.getValue());
        pair = TestingUtils.getOverlappingSum(rows, cols, UPPR, OverlappingSum.WINDOW_SZE);
        this.runAllIntegrationTests(pair.getKey(), rows, rows, cols, cols * 2, seed, pair.getValue());
    }

    @Test
    public void testRunningSum () {
        final long seed = 42;
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getRunningSum(0, 0, UPPR);
        this.runEmptyIntegrationTests(pair.getKey(), seed, pair.getValue());
        pair = TestingUtils.getRunningSum(rows, cols, UPPR);
        this.runAllIntegrationTests(pair.getKey(), rows, rows, cols, cols * 2, seed, pair.getValue());
    }

    @Test
    public void testSingleCellSum () {
        final long seed = 42;
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getSingleCellSum(0, 0, UPPR);
        this.runEmptyIntegrationTests(pair.getKey(), seed, pair.getValue());
        pair = TestingUtils.getSingleCellSum(rows, cols, UPPR);
        this.runAllIntegrationTests(pair.getKey(), rows, rows, cols, cols * 2, seed, pair.getValue());
    }

    @Test
    public void testCompleteBipartiteVlookup () {
        final long seed = 42;
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        BiFunction<Integer, Integer, String> getExpectedExcelFormula;
        BiFunction<Integer, Integer, String> getExpectedLibreFormula;

        pair = TestingUtils.getCompleteBipartiteVlookup(0, 0, UPPR, false);
        getExpectedExcelFormula = pair.getValue();
        pair = TestingUtils.getCompleteBipartiteVlookup(0, 0, UPPR, true);
        getExpectedLibreFormula = pair.getValue();
        this.runEmptyIntegrationTests(pair.getKey(), seed, getExpectedExcelFormula, getExpectedLibreFormula);

        pair = TestingUtils.getCompleteBipartiteVlookup(rows, cols, UPPR, false);
        getExpectedExcelFormula = pair.getValue();
        pair = TestingUtils.getCompleteBipartiteVlookup(rows, cols, UPPR, true);
        getExpectedLibreFormula = pair.getValue();
        this.runAllIntegrationTests(pair.getKey(), rows, rows, 1, 3, seed, getExpectedExcelFormula, getExpectedLibreFormula);
    }

    @Test
    public void testSameCellVlookup () {
        // NOTE: This sheet creator will always create sheets with #N/A values, which makes it 
        // difficult to test. With that in mind, these types of sheets should be inspected manually.
    }

    @Test
    public void testSingleCellVlookup () {
        final long seed = 42;
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        BiFunction<Integer, Integer, String> getExpectedExcelFormula;
        BiFunction<Integer, Integer, String> getExpectedLibreFormula;

        pair = TestingUtils.getSingleCellVlookup(0, 0, UPPR, false);
        getExpectedExcelFormula = pair.getValue();
        pair = TestingUtils.getSingleCellVlookup(0, 0, UPPR, true);
        getExpectedLibreFormula = pair.getValue();
        this.runEmptyIntegrationTests(pair.getKey(), seed, getExpectedExcelFormula, getExpectedLibreFormula);

        pair = TestingUtils.getSingleCellVlookup(rows, cols, UPPR, false);
        getExpectedExcelFormula = pair.getValue();
        pair = TestingUtils.getSingleCellVlookup(rows, cols, UPPR, true);
        getExpectedLibreFormula = pair.getValue();
        this.runAllIntegrationTests(pair.getKey(), rows, rows, 1, 3, seed, getExpectedExcelFormula, getExpectedLibreFormula);
    }

    @Test
    public void testSpecialOverlappingSum () {
        final long seed = 42;
        sums.specialsums.OverlappingSum.setMaxRows(EXCLUSIVE_UPPER_BOUND * 2);
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getSpecialOverlappingSum(rows, cols, UPPR, sums.specialsums.OverlappingSum.WINDOW_SZE, seed);
        this.runAllIntegrationTests(pair.getKey(), rows, sums.specialsums.OverlappingSum.MAX_V_ROWS, cols, cols * 2, seed, pair.getValue());
    }

    @Test
    public void testSpecialNoEdgeSum () {
        final long seed = 42;
        sums.specialsums.NoEdgeSum.setMaxRows(EXCLUSIVE_UPPER_BOUND * 2);
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getSpecialNoEdgeSum(rows, cols, UPPR, seed);
        this.runAllIntegrationTests(pair.getKey(), rows, sums.specialsums.NoEdgeSum.MAX_V_ROWS, cols, cols * 2, seed, pair.getValue());
    }

    @Test
    public void testSpecialMixedRangeSum () {
        final long seed = 42;
        sums.specialsums.MixedRangeSum.setMaxRows(EXCLUSIVE_UPPER_BOUND * 2);
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getSpecialMixedRangeSum(rows, cols, UPPR, seed);
        this.runAllIntegrationTests(pair.getKey(), rows, sums.specialsums.MixedRangeSum.MAX_V_ROWS, cols, cols * 2, seed, pair.getValue());
    }

    @Test
    public void testSpecialCompleteBipartiteSum () {
        final long seed = 42;
        sums.specialsums.CompleteBipartiteSum.setMaxRows(EXCLUSIVE_UPPER_BOUND * 2);
        AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> pair;
        pair = TestingUtils.getSpecialCompleteBipartiteSum(rows, cols, UPPR, seed);
        this.runAllIntegrationTests(pair.getKey(), rows, sums.specialsums.CompleteBipartiteSum.MAX_V_ROWS, cols, cols * 2, seed, pair.getValue());
    }

}
