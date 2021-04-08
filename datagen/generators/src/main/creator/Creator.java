package creator;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.github.jferard.fastods.AnonymousOdsFileWriter;
import com.github.jferard.fastods.OdsFactory;
import com.github.jferard.fastods.Table;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.io.File;

import java.util.logging.Logger;
import java.util.OptionalLong;
import java.util.Locale;

public abstract class Creator {

    private static final OdsFactory odsFactory = OdsFactory.create(Logger.getLogger("logger"), Locale.US);

    public static void createExcelSheet (Creatable createable, String fPath, String vPath, int rows, int cols, OptionalLong seed) {
        try (SXSSFWorkbook fWorkbook = new SXSSFWorkbook(1); SXSSFWorkbook vWorkbook = new SXSSFWorkbook(1)) {
            String fName = Path.of(fPath, "fv-" + rows + ".xlsx").toString();
            String vName = Path.of(vPath, "vo-" + rows + ".xlsx").toString();
            if (!(new File(fName)).exists() || !(new File(vName)).exists()) {
                fWorkbook.setCompressTempFiles(true);
                vWorkbook.setCompressTempFiles(true);
                SXSSFSheet fSheet = fWorkbook.createSheet("Sheet1");
                SXSSFSheet vSheet = vWorkbook.createSheet("Sheet1");
                if (seed.isPresent()) {
                    createable.createRandomExcelSheet(fSheet, vSheet, rows, cols, seed.getAsLong());
                } else {
                    createable.createExcelSheet(fSheet, vSheet, rows, cols);
                }
                Creator.saveWorkbook(fWorkbook, fName);
                Creator.saveWorkbook(vWorkbook, vName);
                fWorkbook.dispose();
                vWorkbook.dispose();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void createCalcSheet (Creatable createable, String fPath, String vPath, int rows, int cols, OptionalLong seed) {
        File fName = Path.of(fPath, "fv-" + rows + ".ods").toFile();
        File vName = Path.of(vPath, "vo-" + rows + ".ods").toFile();
        if (!fName.exists() || !vName.exists()) {
            try {
                final AnonymousOdsFileWriter fWriter = Creator.odsFactory.createWriter();
                final AnonymousOdsFileWriter vWriter = Creator.odsFactory.createWriter();
                Table fSheet = fWriter.document().addTable("Sheet1");
                Table vSheet = vWriter.document().addTable("Sheet1");
                if (seed.isPresent()) {
                    createable.createRandomCalcSheet(fSheet, vSheet, rows, cols, seed.getAsLong());
                } else {
                    createable.createCalcSheet(fSheet, vSheet, rows, cols);
                }
                fWriter.saveAs(fName);
                vWriter.saveAs(vName);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static void saveWorkbook (Workbook wb, String name) {
        try (FileOutputStream fileOut = new FileOutputStream(name)) {
            wb.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
}
