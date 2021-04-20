package creator;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import com.github.jferard.fastods.Table;
import java.io.IOException;

public interface Creatable {

    public void createExcelSheet        (SXSSFSheet fSheet, SXSSFSheet  vSheet              );
    public void createRandomExcelSheet  (SXSSFSheet fSheet, SXSSFSheet  vSheet, long seed   );
    public void createCalcSheet         (Table      fSheet, Table       vSheet              )   throws IOException;
    public void createRandomCalcSheet   (Table      fSheet, Table       vSheet, long seed   )   throws IOException;

}
