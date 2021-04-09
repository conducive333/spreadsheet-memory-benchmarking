import java.util.concurrent.ExecutorService;
import java.io.FileInputStream;
import java.util.OptionalLong;
import java.util.Properties;
import java.io.IOException;
import java.nio.file.Path;
import java.io.File;

import vlookups.*;
import creator.*;
import utils.*;
import sums.*;

public class Main {

    private static final Creatable      INST;
    private static final Path           PATH;
    private static final OptionalLong   SEED;
    private static final boolean        XLSX;
    private static final int            STEP;
    private static final int            ROWS;
    private static final int            COLS;
    private static final int            ITRS;
    private static final int            POOL;

    static {
        Properties pr = new Properties();
        try(FileInputStream in = new FileInputStream("config")) { pr.load(in); }
        catch (IOException e) { e.printStackTrace(); System.exit(1); }
        PATH = Path.of(pr.getProperty("PATH"));
        INST = Main.resolveName(pr.getProperty("INST"));
        SEED = Main.resolveSeed(pr.getProperty("SEED"));
        STEP = Integer.parseInt(pr.getProperty("STEP"));
        ROWS = Integer.parseInt(pr.getProperty("ROWS"));
        COLS = Integer.parseInt(pr.getProperty("COLS"));
        ITRS = Integer.parseInt(pr.getProperty("ITRS"));
        POOL = Integer.parseInt(pr.getProperty("POOL"));
        XLSX = Boolean.parseBoolean(pr.getProperty("XLSX"));
    }

    /**
     * @param s
     * @return The spreadsheet creator corresponding to `s`.
     */
    private static Creatable resolveName (String s) {
        if (s.equals("CompleteBipartiteSum"))               return new CompleteBipartiteSum();
        if (s.equals("CompleteBipartiteSumWithConstant"))   return new CompleteBipartiteSumWithConstant();
        if (s.equals("MixedRangeSum"))                      return new MixedRangeSum();
        if (s.equals("RunningSum"))                         return new RunningSum();
        if (s.equals("SingleCellSum"))                      return new SingleCellSum();
        if (s.equals("CompleteBipartiteVlookup"))           return new CompleteBipartiteVlookup();
        if (s.equals("SingleCellVlookup"))                  return new SingleCellVlookup();
        return null;
    }

    /**
     * @param s
     * @return An OptionalLong containing `s` interpreted as a 
     * long. If `s` is empty, returns an empty OptionalLong.
     */
    private static OptionalLong resolveSeed (String s) {
        if (s.length() == 0) {
            return OptionalLong.empty();
        }
        return OptionalLong.of(Long.parseLong(s));
    }

    /**
     * @return An array of strings, ARR, where ARR[0] is the path
     * to the formula-value directory and ARR[1] is the path to 
     * the value-only directory.
     */
    private static String[] createDirectories () {
        File vDir = Path.of(Main.PATH.toString(), "value-only"   ).toFile();
        File fDir = Path.of(Main.PATH.toString(), "formula-value").toFile();
        if (!vDir.exists()) { vDir.mkdirs(); }
        if (!fDir.exists()) { fDir.mkdirs(); }
        return new String[] { fDir.toString(), vDir.toString() };
    }

    /**
     * A wrapper method for creating spreadsheets.
     * 
     * @param fPath
     * @param vPath
     * @param rows
     */
    private static void createSpreadsheet (String fPath, String vPath, int rows) {
        System.out.println("Creating a sheet with " + rows + " row(s)");
        if (Main.XLSX) {
            Creator.createExcelSheet(Main.INST, fPath, vPath, rows, Main.COLS, Main.SEED);
        } else {
            Creator.createCalcSheet(Main.INST, fPath, vPath, rows, Main.COLS, Main.SEED);
        }
    }

    public static void main (String[] args) throws IOException {

        /** Setup */
        String[]    paths = Main.createDirectories();
        Stopwatch   stopw = new Stopwatch();

        /** Create datasets */
        stopw.start();
        if (Main.POOL == 1) {
            for (int i = 0, r = Main.ROWS; i < Main.ITRS; i++, r += Main.STEP) {
                Main.createSpreadsheet(paths[0], paths[1], r);
            }
        } else {
            ExecutorService exc = SimpleThreadPoolExecutor.getNewExecutor(Main.POOL);
            for (int i = 0, r = Main.ROWS; i < Main.ITRS; i++, r += Main.STEP) {
                int[] rSize = {r};
                exc.submit(() -> { Main.createSpreadsheet(paths[0], paths[1], rSize[0]); });
            }
            SimpleThreadPoolExecutor.wait(exc);
        }
        stopw.printDuration();

    }
}
