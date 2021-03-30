package creator;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import com.github.jferard.fastods.Table;

import java.io.IOException;
import java.util.Random;

public interface Creatable {

    public void createExcelSheet        (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols);
    public void createRandomExcelSheet  (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols, Random rand);
    public void createCalcSheet         (Table fSheet, Table vSheet, int rows, int cols) throws IOException;
    public void createRandomCalcSheet   (Table fSheet, Table vSheet, int rows, int cols, Random rand) throws IOException;

}
