package vlookups;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFRow;

import com.github.jferard.fastods.TableRowImpl;
import com.github.jferard.fastods.Table;

import java.io.IOException;
import java.util.Random;
import java.util.List;

import creator.Creatable;

public class SingleCellVlookup extends BaseVlookup implements Creatable {
    /**
     * Creates a spreadsheet with the following structure:
     * 
     * 0    |   A   |               B                   |   C   |
     * ----------------------------------------------------------
     * 1    |   ?   |   =VLOOKUP(C1, A1:A1, 1, FALSE)   |   ?   |
     * 2    |   ?   |   =VLOOKUP(C2, A2:A2, 1, FALSE)   |   ?   |
     * ...  |   ... |   ...                             |   ?   |
     * N    |   ?   |   =VLOOKUP(CN, AN:AN, 1, FALSE)   |   ?   |
     * 
     * ? = a UNIQUE random value (or a placeholder value if no random 
     * seed is specified). Unlike the classes in the sums package, the
     * COLS parameter is ignored for this class. The values in column 
     * A will be in the same order as the values in column C.
     */

    private static final double FILL_VALUE = 1.0;
    private static final String EXCEL_FSTR = "VLOOKUP(C%1$d, A%1$d:A%1$d, 1, FALSE)";
    private static final String LIBRE_FSTR = "VLOOKUP(C%1$d; A%1$d:A%1$d; 1; 0)";

    @Override
    public void createExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols) {
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            fRow.createCell(0).setCellValue(FILL_VALUE);
            vRow.createCell(0).setCellValue(FILL_VALUE);
            fRow.createCell(1).setCellFormula(String.format(EXCEL_FSTR, r + 1));
            vRow.createCell(1).setCellValue(FILL_VALUE);
            fRow.createCell(2).setCellValue(FILL_VALUE);
            vRow.createCell(2).setCellValue(FILL_VALUE);
        }
    }

    @Override
    public void createRandomExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols, long seed) {
        List<Double> vals = super.getShuffledConsecutiveNumbers(new Random(seed), rows);
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            fRow.createCell(0).setCellValue(vals.get(r));
            vRow.createCell(0).setCellValue(vals.get(r));
            fRow.createCell(1).setCellFormula(String.format(EXCEL_FSTR, r + 1));
            vRow.createCell(1).setCellValue(vals.get(r));
            fRow.createCell(2).setCellValue(vals.get(r));
            vRow.createCell(2).setCellValue(vals.get(r));
        }
    }

    @Override
    public void createCalcSheet(Table fSheet, Table vSheet, int rows, int cols) throws IOException {
        for (int r = 0; r < rows; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            fRow.getOrCreateCell(0).setFloatValue(FILL_VALUE);
            vRow.getOrCreateCell(0).setFloatValue(FILL_VALUE);
            fRow.getOrCreateCell(1).setFormula(String.format(LIBRE_FSTR, r + 1));
            vRow.getOrCreateCell(1).setFloatValue(FILL_VALUE);
            fRow.getOrCreateCell(2).setFloatValue(FILL_VALUE);
            vRow.getOrCreateCell(2).setFloatValue(FILL_VALUE);
        }
    }

    @Override
    public void createRandomCalcSheet(Table fSheet, Table vSheet, int rows, int cols, long seed) throws IOException {
        List<Double> vals = super.getShuffledConsecutiveNumbers(new Random(seed), rows);
        for (int r = 0; r < rows; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            fRow.getOrCreateCell(0).setFloatValue(vals.get(r));
            vRow.getOrCreateCell(0).setFloatValue(vals.get(r));
            fRow.getOrCreateCell(1).setFormula(String.format(LIBRE_FSTR, r + 1));
            vRow.getOrCreateCell(1).setFloatValue(vals.get(r));
            fRow.getOrCreateCell(2).setFloatValue(vals.get(r));
            vRow.getOrCreateCell(2).setFloatValue(vals.get(r));
        }
    }
    
}
