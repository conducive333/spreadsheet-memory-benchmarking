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
     * Deletes all non-directory files in `dir`.
     * 
     * @param dir
     */
    public static void deleteAllFilesInDirectory (File dir) {
        for (File f : dir.listFiles()) {
            if (!f.isDirectory()) {
                f.delete();
            }
        }
    }

}
