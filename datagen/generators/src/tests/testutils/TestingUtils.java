package testutils;

import java.nio.file.Path;
import java.io.File;

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

}
