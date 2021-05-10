package sums.specialsums;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.ss.util.CellReference;

import com.github.jferard.fastods.TableRowImpl;
import com.github.jferard.fastods.Table;

import java.util.ArrayList;
import java.io.IOException;
import java.util.Random;
import java.util.List;

import creator.Creatable;

public class OverlappingSum extends BaseSpecialSum implements Creatable {
    /**
     * Creates a spreadsheet with the following structure:
     * 
     *      0       |   A   |                       B                       |
     * ----------------------------------------------------------------------
     *      1       |   ?   |   =SUM(A[1]:A[2])                             |
     *      2       |   ?   |   =SUM(A[2]:A[3])                             |
     *     ...      |   ... |   ...                                         |
     *      N       |   ?   |   =SUM(A[N]:A[N+1])                           |
     *    N + 1     |   ?   |   =EVAL(SUM(A[N+1]:A[N+2]))                   |
     *     ...      |   ?   |   ...                                         |
     *  MAX_V_ROWS  |   ?   |   =EVAL(SUM(A[MAX_V_ROWS]:A[MAX_V_ROWS+1]))   |
     * 
     * ? = a random value (or a placeholder value if no random seed is specified). 
     * The COLS parameter controls the number of columns with values AND the number 
     * of columns with formulae. For example, if COLS = 2, then there will be 2 
     * columns of values and 2 columns of formulae. The window size adjusts the
     * number of cells in the range. In the example above, WINDOW_SZE = 2. EVAL(...)
     * means the evaluated result of the input formula and N is the number of rows.
     */

    private static final String CREATE_STR = "SUM(%s%d:%s%d)";
    public  static final int    WINDOW_SZE = 500;

    public OverlappingSum (int uppr) {
        super(uppr);
    }

    @Override
    public void createExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols) {
        for (int r = 0; r < MAX_V_ROWS; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.createCell(c).setCellValue(FILL_VALUE);
                vRow.createCell(c).setCellValue(FILL_VALUE);
                if (r < rows) {
                    fRow.createCell(c + cols).setCellFormula(
                        String.format(CREATE_STR
                            , CellReference.convertNumToColString(c)
                            , r + 1
                            , CellReference.convertNumToColString(c)
                            , r + WINDOW_SZE
                        )
                    );
                } else {
                    fRow.createCell(c + cols).setCellValue(Math.min(WINDOW_SZE, MAX_V_ROWS - r) * FILL_VALUE);                    
                }
                vRow.createCell(c + cols).setCellValue(Math.min(WINDOW_SZE, MAX_V_ROWS - r) * FILL_VALUE);
            }
        }
    }

    @Override
    public void createRandomExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols, long seed) {
        List<Double> values = new ArrayList<>();
        super.randomlyFillList(values, MAX_V_ROWS * cols, new Random(seed));
        for (int r = 0; r < MAX_V_ROWS; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.createCell(c).setCellValue(values.get(cols * r + c));
                vRow.createCell(c).setCellValue(values.get(cols * r + c));

                // This ensures that we don't run into an out of bounds error
                double total = 0.0;
                for (int i = 0; i < WINDOW_SZE; i++) {
                    if (cols * (r + i) + c < values.size()) 
                        total += values.get(cols * (r + i) + c);
                    else break;
                }
                
                if (r < rows) {
                    fRow.createCell(c + cols).setCellFormula(
                        String.format(CREATE_STR
                            , CellReference.convertNumToColString(c)
                            , r + 1
                            , CellReference.convertNumToColString(c)
                            , r + WINDOW_SZE
                        )
                    );
                } else {
                    fRow.createCell(c + cols).setCellValue(total);
                }
                vRow.createCell(c + cols).setCellValue(total);
            }
        }
    }

    @Override
    public void createCalcSheet(Table fSheet, Table vSheet, int rows, int cols) throws IOException {
        for (int r = 0; r < MAX_V_ROWS; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
                vRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
                if (r < rows) {
                    fRow.getOrCreateCell(c + cols).setFormula(
                        String.format(CREATE_STR
                            , CellReference.convertNumToColString(c)
                            , r + 1
                            , CellReference.convertNumToColString(c)
                            , r + WINDOW_SZE
                        )
                    );
                } else {
                    fRow.getOrCreateCell(c + cols).setFormula((Math.min(WINDOW_SZE, MAX_V_ROWS - r) * FILL_VALUE) + "");
                }
                vRow.getOrCreateCell(c + cols).setFloatValue(Math.min(WINDOW_SZE, MAX_V_ROWS - r) * FILL_VALUE);
            }
        }
    }

    @Override
    public void createRandomCalcSheet(Table fSheet, Table vSheet, int rows, int cols, long seed) throws IOException {
        List<Double> values = new ArrayList<>();
        super.randomlyFillList(values, MAX_V_ROWS * cols, new Random(seed));
        for (int r = 0; r < MAX_V_ROWS; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.getOrCreateCell(c).setFloatValue(values.get(cols * r + c));
                vRow.getOrCreateCell(c).setFloatValue(values.get(cols * r + c));

                // This ensures that we don't run into an out of bounds error
                double total = 0.0;
                for (int i = 0; i < WINDOW_SZE; i++) {
                    if (cols * (r + i) + c < values.size()) 
                        total += values.get(cols * (r + i) + c);
                    else break;
                }

                if (r < rows) {
                    fRow.getOrCreateCell(c + cols).setFormula(
                        String.format(CREATE_STR
                            , CellReference.convertNumToColString(c)
                            , r + 1
                            , CellReference.convertNumToColString(c)
                            , r + WINDOW_SZE
                        )
                    );
                } else {
                    fRow.getOrCreateCell(c + cols).setFormula(total + "");
                }
                vRow.getOrCreateCell(c + cols).setFloatValue(total);
            }
        }
    }

}
