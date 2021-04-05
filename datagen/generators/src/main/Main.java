import java.util.concurrent.ExecutorService;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.io.File;

import java.util.Properties;
import java.util.Random;

import vlookups.*;
import creator.*;
import utils.*;
import sums.*;

public class Main {

    private static Creatable  INST;
    private static Path       PATH;
    private static Random     RAND;
    private static boolean    XLSX;
    private static int        STEP;
    private static int        ROWS;
    private static int        COLS;
    private static int        ITRS;
    private static int        POOL;

    static {
        try {
            FileInputStream in  = new FileInputStream("config");
            Properties      pr  = new Properties();
            pr.load(in);
            INST = Main.resolveName(pr.getProperty("INST"));
            PATH = Path.of(pr.getProperty("PATH"));
            RAND = pr.getProperty("RAND").length() == 0 ? null : new Random(Long.parseLong(pr.getProperty("RAND")));
            XLSX = Boolean.parseBoolean(pr.getProperty("XLSX"));
            STEP = Integer.parseInt(pr.getProperty("STEP"));
            ROWS = Integer.parseInt(pr.getProperty("ROWS"));
            COLS = Integer.parseInt(pr.getProperty("COLS"));
            ITRS = Integer.parseInt(pr.getProperty("ITRS"));
            POOL = Integer.parseInt(pr.getProperty("POOL"));
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Creatable resolveName (String s) {
        if (s.equals("CompleteBipartiteSum"))   return new CompleteBipartiteSum();
        if (s.equals("RunningSum"))             return new RunningSum();
        if (s.equals("RunningSumAvg"))          return new RunningSumAvg();
        if (s.equals("SingleCellSum"))          return new SingleCellSum();
        if (s.equals("SingleCellSumAvg"))       return new SingleCellSumAvg();
        if (s.equals("RunningVlookup"))         return new RunningVlookup();
        if (s.equals("SingleCellVlookup"))      return new SingleCellVlookup();
        return null;
    }

    private static String[] createDirectories () {
        File vDir = Path.of(Main.PATH.toString(), "value-only"   ).toFile();
        File fDir = Path.of(Main.PATH.toString(), "formula-value").toFile();
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
