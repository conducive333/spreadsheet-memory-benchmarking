package sums.specialsums;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.ss.util.CellReference;

import com.github.jferard.fastods.TableRowImpl;
import com.github.jferard.fastods.Table;

import java.io.IOException;
import java.util.ArrayDeque;
import java.util.Random;
import java.util.Deque;

import creator.Creatable;

public class MixedRangeSum extends BaseSpecialSum implements Creatable {
  /**
   * Creates a spreadsheet with the following structure:
   * 
   *    0       |   A   |                                   B                                 |
   *  -----------------------------------------------------------------------------------------
   *     1      |   ?   |   =SUM(A[1]:A[1]) + SUM(A[1]:A[MAX_V_ROWS])                         |
   *     2      |   ?   |   =SUM(A[2]:A[2]) + SUM(A[1]:A[MAX_V_ROWS])                         |
   *    ...     |  ...  |   ...                                                               |
   *     N      |   ?   |   =SUM(A[N]:A[N]) + SUM(A[1]:A[MAX_V_ROWS])                         |
   *   N + 1    |   ?   |   =EVAL(SUM(A[N+1]:A[N+1]) + SUM(A[1]:A[MAX_V_ROWS]))               |
   *    ...     |   ?   |    ...                                                              |
   * MAX_V_ROWS |   ?   |   =EVAL(SUM(A[MAX_V_ROWS]:A[MAX_V_ROWS]) + SUM(A[1]:A[MAX_V_ROWS])) |
   * 
   * ? = a random value (or a placeholder value if no random seed is specified). The COLS parameter 
   * controls the number of columns with values AND the number of columns with formulae. For example,
   * if COLS = 2, then there will be 2 columns of values and 2 columns of formulae. The window size 
   * adjusts the number of cells in the range. In the example above, WINDOW_SZE = 2. EVAL(...) means
   * the evaluated result of the input formula and N is the number of rows.
   */

  private static final String CREATE_STR = "SUM(%s%d:%s%d) + SUM(A1:%s%d)";

  public MixedRangeSum (int uppr) {
    super(uppr);
  }

  @Override
  public void createExcelSheet(SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols) {
    for (int r = 0; r < MAX_V_ROWS; r++) {
      SXSSFRow fRow = fSheet.createRow(r);
      SXSSFRow vRow = vSheet.createRow(r);
      for (int c = 0; c < cols; c++) {
        String col = CellReference.convertNumToColString(c);
        int row = r + 1;
        fRow.createCell(c).setCellValue(FILL_VALUE);
        vRow.createCell(c).setCellValue(FILL_VALUE);
        if (r < rows) {
          fRow.createCell(c + cols).setCellFormula(
            String.format(CREATE_STR
              , col
              , row
              , col
              , row
              , CellReference.convertNumToColString(cols - 1)
              , MAX_V_ROWS
            )
          );
        } else {
          fRow.createCell(c + cols).setCellValue(FILL_VALUE + (FILL_VALUE * MAX_V_ROWS * cols));
        }
        vRow.createCell(c + cols).setCellValue(FILL_VALUE + (FILL_VALUE * MAX_V_ROWS * cols));
      }
    }
  }

  @Override
  public void createRandomExcelSheet(SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols, long seed) {
    Deque<Double> values = new ArrayDeque<>();
    double total = super.randomlyFillDeque(values, MAX_V_ROWS * cols, new Random(seed));
    for (int r = 0; r < MAX_V_ROWS; r++) {
      SXSSFRow fRow = fSheet.createRow(r);
      SXSSFRow vRow = vSheet.createRow(r);
      for (int c = 0; c < cols; c++) {
        double num = values.pop();
        String col = CellReference.convertNumToColString(c);
        int row = r + 1;
        fRow.createCell(c).setCellValue(num);
        vRow.createCell(c).setCellValue(num);
        if (r < rows) {
          fRow.createCell(c + cols).setCellFormula(
            String.format(CREATE_STR
              , col
              , row
              , col
              , row
              , CellReference.convertNumToColString(cols - 1)
              , MAX_V_ROWS
            )
          );
        } else {
          fRow.createCell(c + cols).setCellValue(num + total);
        }
        vRow.createCell(c + cols).setCellValue(num + total);
      }
    }
  }

  @Override
  public void createCalcSheet(Table fSheet, Table vSheet, int rows, int cols) throws IOException {
    for (int r = 0; r < MAX_V_ROWS; r++) {
      TableRowImpl fRow = fSheet.getRow(r);
      TableRowImpl vRow = vSheet.getRow(r);
      for (int c = 0; c < cols; c++) {
        String col = CellReference.convertNumToColString(c);
        int row = r + 1;
        fRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
        vRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
        if (r < rows) {
          fRow.getOrCreateCell(c + cols).setFormula(
            String.format(CREATE_STR
              , col
              , row
              , col
              , row
              , CellReference.convertNumToColString(cols - 1)
              , MAX_V_ROWS
            )
          );
        } else {
          fRow.getOrCreateCell(c + cols).setFormula((FILL_VALUE + (FILL_VALUE * MAX_V_ROWS * cols)) + "");
        }
        vRow.getOrCreateCell(c + cols).setFloatValue(FILL_VALUE + (FILL_VALUE * MAX_V_ROWS * cols));
      }
    }
  }

  @Override
  public void createRandomCalcSheet(Table fSheet, Table vSheet, int rows, int cols, long seed) throws IOException {
    Deque<Double> values = new ArrayDeque<>();
    double total = super.randomlyFillDeque(values, MAX_V_ROWS * cols, new Random(seed));
    for (int r = 0; r < MAX_V_ROWS; r++) {
      TableRowImpl fRow = fSheet.getRow(r);
      TableRowImpl vRow = vSheet.getRow(r);
      for (int c = 0; c < cols; c++) {
        double num = values.pop();
        String col = CellReference.convertNumToColString(c);
        int row = r + 1;
        fRow.getOrCreateCell(c).setFloatValue(num);
        vRow.getOrCreateCell(c).setFloatValue(num);
        if (r < rows) {
          fRow.getOrCreateCell(c + cols).setFormula(
            String.format(CREATE_STR
              , col
              , row
              , col
              , row
              , CellReference.convertNumToColString(cols - 1)
              , MAX_V_ROWS
            )
          );
        } else {
          fRow.getOrCreateCell(c + cols).setFormula((num + total) + "");          
        }
        vRow.getOrCreateCell(c + cols).setFloatValue(num + total);
      }
    }
  }

}
