import java.util.concurrent.ExecutorService;

import util.SimpleThreadPoolExecutor;
import util.Stopwatch;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Random;
import java.io.File;

import creator.Creatable;
import creator.Creator;
import sums.*;

public class Main {

    /** Edit zone: */
    private static final Creatable  INST = new CompleteBipartiteSum();
    private static final Path       PATH = Path.of("..", "exp");    
    private static final String     FLDR = "RCBS-1col-xlsx";
    private static final Random     RAND = new Random(42L);
    private static final boolean    XLSX = true;
    private static final int        STEP = 10000;
    private static final int        ROWS = 0;
    private static final int        COLS = 1;
    private static final int        ITER = 10;
    private static       int        POOL = 5;
    /***************/

    private static String[] createDirectories () {
        File vDir = Path.of(Main.PATH.toString(), Main.FLDR, "value-only"   ).toFile();
        File fDir = Path.of(Main.PATH.toString(), Main.FLDR, "formula-value").toFile();
        if (!vDir.exists()) { vDir.mkdirs(); }
        if (!fDir.exists()) { fDir.mkdirs(); }
        return new String[] { fDir.toString(), vDir.toString() };
    }

    private static void createSpreadsheet (String fPath, String vPath, int rows) {
        System.out.println("Creating a sheet with " + rows + " row(s)");
        if (Main.XLSX) {
            Creator.createExcelSheet(Main.INST, fPath, vPath, rows, Main.COLS, Main.RAND);
        } else {
            Creator.createCalcSheet(Main.INST, fPath, vPath, rows, Main.COLS, Main.RAND);
        }
    }

    public static void main (String[] args) throws IOException {

        /** Setup */
        String[]    paths = Main.createDirectories();
        Stopwatch   stopw = new Stopwatch();

        /** Create datasets */
        stopw.start();
        if (Main.POOL == 1) {
            for (int i = 0, r = Main.ROWS; i < Main.ITER; i++, r += Main.STEP) {
                Main.createSpreadsheet(paths[0], paths[1], r);
            }
        } else {
            ExecutorService exc = SimpleThreadPoolExecutor.getNewExecutor(Main.POOL);
            for (int i = 0, r = Main.ROWS; i < Main.ITER; i++, r += Main.STEP) {
                int[] rSize = {r};
                exc.submit(() -> { Main.createSpreadsheet(paths[0], paths[1], rSize[0]); });
            }
            SimpleThreadPoolExecutor.wait(exc);
        }
        stopw.printDuration();

    }
}
