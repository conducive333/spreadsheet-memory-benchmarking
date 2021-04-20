package sums;

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

public class OverlappingSum extends BaseSum implements Creatable {
    /**
     * Creates a spreadsheet with the following structure:
     * 
     * 0    |   A   |           B           |
     * --------------------------------------
     * 1    |   ?   |   =SUM(A1:A2)         |
     * 2    |   ?   |   =SUM(A2:A3)         |
     * ...  |   ... |   ...                 |
     * N    |   ?   |   =SUM(A[N]:A[N+1])   |
     * 
     * ? = a random value (or a placeholder value if no
     * random seed is specified). The COLS parameter 
     * controls the number of columns with values AND
     * the number of columns with formulae. For example,
     * if COLS = 2, then there will be 2 columns of values
     * and 2 columns of formulae. The window size adjusts
     * the number of cells in the range. In the example 
     * above, WINDOW_SZE = 2.
     */

    private static final String CREATE_STR = "SUM(%s%d:%s%d)";
    public  static final int    WINDOW_SZE = 500;

    public OverlappingSum (int uppr) {
        super(uppr);
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
                        , CellReference.convertNumToColString(c)
                        , r + 1
                        , CellReference.convertNumToColString(c)
                        , r + WINDOW_SZE
                    )
                );
                vRow.createCell(c + cols).setCellValue(Math.min(WINDOW_SZE, rows - r) * FILL_VALUE);
            }
        }
    }

    @Override
    public void createRandomExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols, long seed) {
        List<Double> values = new ArrayList<>();
        super.randomlyFillList(values, rows * cols, new Random(seed));
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.createCell(c).setCellValue(values.get(cols * r + c));
                vRow.createCell(c).setCellValue(values.get(cols * r + c));
                fRow.createCell(c + cols).setCellFormula(
                    String.format(CREATE_STR
                        , CellReference.convertNumToColString(c)
                        , r + 1
                        , CellReference.convertNumToColString(c)
                        , r + WINDOW_SZE
                    )
                );
                double total = 0.0;
                for (int i = 0; i < WINDOW_SZE; i++) {
                    if (cols * (r + i) + c < values.size()) 
                        total += values.get(cols * (r + i) + c);
                    else break;
                }
                vRow.createCell(c + cols).setCellValue(total);
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
                        , CellReference.convertNumToColString(c)
                        , r + 1
                        , CellReference.convertNumToColString(c)
                        , r + WINDOW_SZE
                    )
                );
                vRow.getOrCreateCell(c + cols).setFloatValue(Math.min(WINDOW_SZE, rows - r) * FILL_VALUE);
            }
        }
    }

    @Override
    public void createRandomCalcSheet(Table fSheet, Table vSheet, int rows, int cols, long seed) throws IOException {
        List<Double> values = new ArrayList<>();
        super.randomlyFillList(values, rows * cols, new Random(seed));
        for (int r = 0; r < rows; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.getOrCreateCell(c).setFloatValue(values.get(cols * r + c));
                vRow.getOrCreateCell(c).setFloatValue(values.get(cols * r + c));
                fRow.getOrCreateCell(c + cols).setFormula(
                    String.format(CREATE_STR
                        , CellReference.convertNumToColString(c)
                        , r + 1
                        , CellReference.convertNumToColString(c)
                        , r + WINDOW_SZE
                    )
                );
                double total = 0.0;
                for (int i = 0; i < WINDOW_SZE; i++) {
                    if (cols * (r + i) + c < values.size()) 
                        total += values.get(cols * (r + i) + c);
                    else break;
                }
                vRow.getOrCreateCell(c + cols).setFloatValue(total);
            }
        }
    }

}
