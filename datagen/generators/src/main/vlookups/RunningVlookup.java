package vlookups;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.ss.util.CellReference;

import com.github.jferard.fastods.TableRowImpl;
import com.github.jferard.fastods.Table;

import java.io.IOException;
import java.util.Collections;
import java.util.ArrayList;
import java.util.Random;
import java.util.List;

import creator.Creatable;

public class RunningVlookup implements Creatable {
    /**
     * Creates a spreadsheet with the following structure:
     * 
     * 0    |   A   |               B                   |
     * --------------------------------------------------
     * 1    |   ?   |   =VLOOKUP(A1, A1:A1, 1, FALSE)   |
     * 2    |   ?   |   =VLOOKUP(A2, A1:A2, 1, FALSE)   |
     * ...  |   ... |   ...                             |
     * N    |   ?   |   =VLOOKUP(AN, A1:AN, 1, FALSE)   |
     * 
     * ? = a UNIQUE random value (or a placeholder value 
     * if no random seed is specified). The COLS parameter 
     * controls the number of columns with values AND the
     * number of columns with formulae. For example, if 
     * COLS = 2, then there will be 2 columns of values and 
     * 2 columns of formulae.
     */

    private static final double FILL_VALUE = 1.0;
    private static final String CREATE_STR = "VLOOKUP(A%d, A1:%s%d, %d, FALSE)";

    private List<Double> getShuffledConsecutiveNumbers (Random rand, int inclusiveLowerBound, int exclusiveUpperBound) {
        List<Double> vals = new ArrayList<>(exclusiveUpperBound - inclusiveLowerBound);
        for (int i = inclusiveLowerBound; i < exclusiveUpperBound; i++) { vals.add((double) i); }
        Collections.shuffle(vals, rand);
        return vals;
    }

    @Override
    public void createExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols) {
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.createCell(c).setCellValue(FILL_VALUE);
                vRow.createCell(c).setCellValue(FILL_VALUE);
                fRow.createCell(c + cols).setCellFormula(
                    String.format(CREATE_STR
                        , r + 1
                        , CellReference.convertNumToColString(cols - 1)
                        , r + 1
                        , cols
                    )
                );
                vRow.createCell(c + cols).setCellValue(FILL_VALUE);
            }
        }
    }

    @Override
    public void createRandomExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols, Random rand) {
        int lookupIndex = cols - 1;
        List<Double> vals = this.getShuffledConsecutiveNumbers(rand, 0, rows * cols);
        for (int r = 0; r < rows; r++) {
            double lookup = vals.get(cols * r + lookupIndex);
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                double val = vals.get(cols * r + c);
                fRow.createCell(c).setCellValue(val);
                vRow.createCell(c).setCellValue(val);
                fRow.createCell(c + cols).setCellFormula(
                    String.format(CREATE_STR
                        , r + 1
                        , CellReference.convertNumToColString(cols - 1)
                        , r + 1
                        , lookupIndex + 1
                    )
                );
                vRow.createCell(c + cols).setCellValue(lookup);
            }
        }
    }

    @Override
    public void createCalcSheet(Table fSheet, Table vSheet, int rows, int cols) throws IOException {
        for (int r = 0; r < rows; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
                vRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
                fRow.getOrCreateCell(c + cols).setFormula(
                    String.format(CREATE_STR
                        , r + 1
                        , CellReference.convertNumToColString(cols - 1)
                        , r + 1
                        , cols
                    )
                );
                vRow.getOrCreateCell(c + cols).setFloatValue(FILL_VALUE);
            }
        }
    }

    @Override
    public void createRandomCalcSheet(Table fSheet, Table vSheet, int rows, int cols, Random rand) throws IOException {
        int lookupIndex = cols - 1;
        List<Double> vals = this.getShuffledConsecutiveNumbers(rand, 0, rows * cols);
        for (int r = 0; r < rows; r++) {
            double lookup = vals.get(cols * r + lookupIndex);
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            for (int c = 0; c < cols; c++) {
                double val = vals.get(cols * r + c);
                fRow.getOrCreateCell(c).setFloatValue(val);
                vRow.getOrCreateCell(c).setFloatValue(val);
                fRow.getOrCreateCell(c + cols).setFormula(
                    String.format(CREATE_STR
                        , r + 1
                        , CellReference.convertNumToColString(cols - 1)
                        , r + 1
                        , lookupIndex + 1
                    )
                );
                vRow.getOrCreateCell(c + cols).setFloatValue(lookup);
            }
        }
    }
    
}
