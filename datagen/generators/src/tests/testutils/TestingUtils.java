package testutils;

import org.apache.poi.ss.util.CellReference;

import java.util.function.BiFunction;
import java.util.AbstractMap;
import java.util.ArrayDeque;
import java.util.ArrayList;
import java.util.Deque;
import java.util.Random;
import java.util.List;

import java.nio.file.Path;
import java.io.File;

import creator.Creatable;
import vlookups.*;
import sums.*;

public class TestingUtils {

    /**
     * These are for testing file creation.
     */
    public static final File TEMP_DIR = Path.of("src", "tests", "temp").toFile();
    public static final File V_FOLDER = Path.of(TEMP_DIR.toString(), "value-only"   ).toFile();
    public static final File F_FOLDER = Path.of(TEMP_DIR.toString(), "formula-value").toFile();

    /**
     * Creates two folders, `V_FOLDER` and `F_FOLDER`, in `TEMP_DIR`. 
     * If these folders exist, all files in them are deleted.
     */
    public static void createDirectories () {
        TestingUtils.createDirectory(TestingUtils.V_FOLDER);
        TestingUtils.createDirectory(TestingUtils.F_FOLDER);
    }

    /**
     * Deletes all folders created by `createDirectories()`.
     */
    public static void deleteDirectories () {
        TestingUtils.V_FOLDER.delete();
        TestingUtils.F_FOLDER.delete();
        TestingUtils.TEMP_DIR.delete();
    }

    /**
     * Deletes all non-directory files in `V_FOLDER` and `F_FOLDER`.
     */
    public static void deleteFiles () {
        TestingUtils.deleteAllFilesInDirectory(TestingUtils.V_FOLDER);
        TestingUtils.deleteAllFilesInDirectory(TestingUtils.F_FOLDER);
    }

    /**
     * Tests if all files in `files` exist.
     * 
     * @param files
     * @return True if all files in `files` exist.
     */
    public static boolean allFilesExist (File[] files) {
        for (File f : files) {
            if (!f.exists()) { return false; }
        }
        return true;
    }

    /**
     * Creates `dir` and all necessary parent folders.
     * If the folder exists all files in it are deleted.
     * 
     * @param dir
     */
    private static void createDirectory (File dir) {
        if (!dir.exists()) {
            dir.mkdirs();
        } else {
            TestingUtils.deleteAllFilesInDirectory(dir); 
        }
    }

    /**
     * Deletes all non-directory files in `dir`.
     * 
     * @param dir
     */
    private static void deleteAllFilesInDirectory (File dir) {
        for (File f : dir.listFiles()) {
            if (!f.isDirectory()) {
                f.delete();
            }
        }
    }

    /** 
     * These methods will return a pair where the key is the spreasheet
     * creator and the value is a function that may be used to get the 
     * expected formula of the creator. 
     */

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getCompleteBipartiteSum(int rows, int cols, int uppr) {
        return new AbstractMap.SimpleEntry<>(new CompleteBipartiteSum(uppr), (currRowIdx, currColIdx) -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), rows));
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getCompleteBipartiteSumWithConstant(int rows, int cols, int uppr) {
        return new AbstractMap.SimpleEntry<>(new CompleteBipartiteSumWithConstant(uppr), (currRowIdx, currColIdx) -> String.format("SUM(A1:%s%d) + %d", CellReference.convertNumToColString(cols - 1), rows, currRowIdx + 1));
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getOverlappingSum(int rows, int cols, int uppr, int windowSize) {
        return new AbstractMap.SimpleEntry<>(new OverlappingSum(uppr), (currRowIdx, currColIdx) -> String.format("SUM(%s%d:%s%d)", CellReference.convertNumToColString(currColIdx), currRowIdx + 1, CellReference.convertNumToColString(currColIdx), currRowIdx + windowSize));
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getRunningSum(int rows, int cols, int uppr) {
        return new AbstractMap.SimpleEntry<>(new RunningSum(uppr), (currRowIdx, currColIdx) -> String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), currRowIdx + 1));
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getCompleteBipartiteVlookup(int rows, int cols, int uppr, boolean calc) {
        if (calc) {
            return new AbstractMap.SimpleEntry<>(new CompleteBipartiteVlookup(uppr), (currRowIdx, currColIdx) -> String.format("VLOOKUP(C%d; A1:A%d; 1; 0)", currRowIdx + 1, rows));
        } else {
            return new AbstractMap.SimpleEntry<>(new CompleteBipartiteVlookup(uppr), (currRowIdx, currColIdx) -> String.format("VLOOKUP(C%d, A1:A%d, 1, FALSE)", currRowIdx + 1, rows));
        }
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getSingleCellVlookup(int rows, int cols, int uppr, boolean calc) {
        if (calc) {
            return new AbstractMap.SimpleEntry<>(new SingleCellVlookup(uppr), (currRowIdx, currColIdx) -> String.format("VLOOKUP(C%1$d; A%1$d:A%1$d; 1; 0)", currRowIdx + 1));
        } else {
            return new AbstractMap.SimpleEntry<>(new SingleCellVlookup(uppr), (currRowIdx, currColIdx) -> String.format("VLOOKUP(C%1$d, A%1$d:A%1$d, 1, FALSE)", currRowIdx + 1));
        }
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getSingleCellSum(int rows, int cols, int uppr) {
        return new AbstractMap.SimpleEntry<>(new SingleCellSum(uppr), (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            return String.format("SUM(%s%d:%s%d)", col, currRowIdx + 1, col, currRowIdx + 1);
        });
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getMixedRangeSum(int rows, int cols, int uppr) {
        return new AbstractMap.SimpleEntry<>(new MixedRangeSum(uppr), (currRowIdx, currColIdx) -> {
            String col = CellReference.convertNumToColString(currColIdx);
            int row = currRowIdx + 1;
            return String.format("SUM(%s%d:%s%d) + SUM(A1:%s%d)", col, row, col, row
                , CellReference.convertNumToColString(cols - 1)
                , rows
            );
        });
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getNoEdgeSum(int rows, int cols, int uppr, long seed) {
        Random[] randRef = { null };
        int[] numRows = { 0 };
        return new AbstractMap.SimpleEntry<>(new NoEdgeSum(uppr), (currRowIdx, currColIdx) -> {
            
            // This will be true once we perform one integration test.
            // Integration tests alternate between non-random and random
            // sheets, so we need to reset these values after each test.
            if (numRows[0] == rows * cols) {
                randRef[0] = (randRef[0] == null ? new Random(seed) : null);
                numRows[0] = 0;
            }

            // Return the correct formula for the corresponding integration
            // test.
            numRows[0]++;
            if (randRef[0] != null) {
                return String.format("SUM(%f)", (double) randRef[0].nextInt(uppr));
            } else {
                return String.format("SUM(%f)", BaseSum.FILL_VALUE);
            }
            
        });
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getSpecialOverlappingSum(int rows, int cols, int uppr, int windowSize, long seed) {
        Random[] randRef = { null };
        int[] numRows = { 0 };
        List<Double> values = new ArrayList<>();
        return new AbstractMap.SimpleEntry<>(new sums.specialsums.OverlappingSum(uppr), (currRowIdx, currColIdx) -> {
        
            // This will be true once we perform one integration test.
            // Integration tests alternate between non-random and random
            // sheets, so we need to reset these values after each test.
            if (numRows[0] == sums.specialsums.OverlappingSum.MAX_V_ROWS * cols) {
                randRef[0] = (randRef[0] == null ? new Random(seed) : null);
                values.clear();
                if (randRef[0] != null) {
                    for (int i = 0; i < sums.specialsums.OverlappingSum.MAX_V_ROWS * cols; i++) {
                        values.add((double) randRef[0].nextInt(uppr));
                    }
                }
                numRows[0] = 0;
            }

            // Return the correct formula for the corresponding integration
            // test.
            numRows[0]++;
            if (currRowIdx < rows) {
                return String.format("SUM(%s%d:%s%d)"
                    , CellReference.convertNumToColString(currColIdx)
                    , currRowIdx + 1
                    , CellReference.convertNumToColString(currColIdx)
                    , currRowIdx + windowSize
                );
            } else {
                if (randRef[0] != null) {
                    double total = 0.0;
                    for (int i = 0; i < windowSize; i++) {
                        if (cols * (currRowIdx + i) + currColIdx < values.size()) {
                            total += values.get(cols * (currRowIdx + i) + currColIdx);
                        } else break;
                    }
                    return total + "";
                } else {
                    return (Math.min(windowSize, sums.specialsums.OverlappingSum.MAX_V_ROWS - currRowIdx) * BaseSum.FILL_VALUE) + "";
                }
            }

        });
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getSpecialNoEdgeSum(int rows, int cols, int uppr, long seed) {
        Random[] randRef = { null };
        int[] numRows = { 0 };
        return new AbstractMap.SimpleEntry<>(new sums.specialsums.NoEdgeSum(uppr), (currRowIdx, currColIdx) -> {
            
            // This will be true once we perform one integration test.
            // Integration tests alternate between non-random and random
            // sheets, so we need to reset these values after each test.
            if (numRows[0] == sums.specialsums.NoEdgeSum.MAX_V_ROWS * cols) {
                randRef[0] = (randRef[0] == null ? new Random(seed) : null);
                numRows[0] = 0;
            }

            // Return the correct formula for the corresponding integration
            // test.
            numRows[0]++;
            if (currRowIdx < rows) {
                if (randRef[0] != null) {
                    return String.format("SUM(%f)", (double) randRef[0].nextInt(uppr));
                } else {
                    return String.format("SUM(%f)", BaseSum.FILL_VALUE);
                }
            } else {
                if (randRef[0] != null) {
                    return ((double) randRef[0].nextInt(uppr)) + "";
                } else {
                    return BaseSum.FILL_VALUE + "";
                }
            }
            
        });
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getSpecialMixedRangeSum(int rows, int cols, int uppr, long seed) {
        Random[] randRef = { null };
        int[] numRows = { 0 };
        double[] total = { 0 };
        Deque<Double> values = new ArrayDeque<>();
        return new AbstractMap.SimpleEntry<>(new sums.specialsums.MixedRangeSum(uppr), (currRowIdx, currColIdx) -> {
        
            // This will be true once we perform one integration test.
            // Integration tests alternate between non-random and random
            // sheets, so we need to reset these values after each test.
            if (numRows[0] == sums.specialsums.MixedRangeSum.MAX_V_ROWS * cols) {
                randRef[0] = (randRef[0] == null ? new Random(seed) : null);
                if (randRef[0] != null) {
                    values.clear();
                    total[0] = 0;
                    for (int i = 0; i < sums.specialsums.MixedRangeSum.MAX_V_ROWS * cols; i++) {
                        double val = (double) randRef[0].nextInt(uppr);
                        values.add(val);
                        total[0] += val;
                    }
                } else {
                    total[0] = BaseSum.FILL_VALUE * sums.specialsums.MixedRangeSum.MAX_V_ROWS * cols;
                }
                numRows[0] = 0;
            }

            // Return the correct formula for the corresponding integration
            // test.
            numRows[0]++;
            if (currRowIdx < rows) {

                // We pop a value here so that the random numbers generated 
                // by the spreadsheet creator align with the ones we generate
                // in this function.
                if (randRef[0] != null) values.pop();
                
                return String.format("SUM(%s%d:%s%d) + SUM(A1:%s%d)"
                    , CellReference.convertNumToColString(currColIdx)
                    , currRowIdx + 1
                    , CellReference.convertNumToColString(currColIdx)
                    , currRowIdx + 1
                    , CellReference.convertNumToColString(cols - 1)
                    , sums.specialsums.MixedRangeSum.MAX_V_ROWS
                );
            } else {
                if (randRef[0] != null) {
                    return (total[0] + values.pop()) + "";
                } else {
                    return (total[0] + BaseSum.FILL_VALUE) + "";
                }
            }

        });
    }

    public static AbstractMap.SimpleEntry<Creatable, BiFunction<Integer, Integer, String>> getSpecialCompleteBipartiteSum(int rows, int cols, int uppr, long seed) {
        Random[] randRef = { null };
        int[] numRows = { 0 };
        double[] total = { 0 };
        return new AbstractMap.SimpleEntry<>(new sums.specialsums.CompleteBipartiteSum(uppr), (currRowIdx, currColIdx) -> {
        
            // This will be true once we perform one integration test.
            // Integration tests alternate between non-random and random
            // sheets, so we need to reset these values after each test.
            if (numRows[0] == sums.specialsums.CompleteBipartiteSum.MAX_V_ROWS * cols) {
                randRef[0] = (randRef[0] == null ? new Random(seed) : null);
                if (randRef[0] != null) {
                    total[0] = 0;
                    for (int i = 0; i < sums.specialsums.CompleteBipartiteSum.MAX_V_ROWS * cols; i++) {
                        total[0] += (double) randRef[0].nextInt(uppr);
                    }
                } else {
                    total[0] = BaseSum.FILL_VALUE * sums.specialsums.CompleteBipartiteSum.MAX_V_ROWS * cols;
                }
                numRows[0] = 0;
            }

            // Return the correct formula for the corresponding integration
            // test.
            numRows[0]++;
            if (currRowIdx < rows) {                
                return String.format("SUM(A1:%s%d)"
                    , CellReference.convertNumToColString(cols - 1)
                    , sums.specialsums.CompleteBipartiteSum.MAX_V_ROWS
                );
            } else {
                return total[0] + "";
            }

        });
    }
}
