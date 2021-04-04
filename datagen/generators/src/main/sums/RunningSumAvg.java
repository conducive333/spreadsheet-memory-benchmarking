package sums;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.ss.util.CellReference;

import com.github.jferard.fastods.TableRowImpl;
import com.github.jferard.fastods.Table;

import java.io.IOException;
import java.util.Random;

import creator.Creatable;

public class RunningSumAvg implements Creatable {
    /**
     * Creates a spreadsheet with the following structure:
     * 
     * 0    |   A   |       B           |
     * ----------------------------------
     * 1    |   ?   |   =SUM(A1:A1)     |
     * 2    |   ?   |   =AVERAGE(A1:A2) |
     * 3    |   ?   |   =SUM(A1:A3)     |
     * 4    |   ?   |   =AVERAGE(A1:A4) |
     * ...  |   ... |   ...             |
     * 
     * ? = a random value (or a placeholder value if no
     * random seed is specified). The COLS parameter 
     * controls the number of columns with values AND
     * the number of columns with formulae. For example,
     * if COLS = 2, then there will be 2 columns of values
     * and 2 columns of formulae.
     */

    private static final double FILL_VALUE = 1.0;
    private static final String CREATE_STR = "%s(A1:%s%d)";

    @Override
    public void createExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols) {
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            String frmula = r % 2 == 0 ? "SUM" : "AVERAGE";
            double divide = r % 2 == 0 ? 1 :  ((r + 1) * cols);
            for (int c = 0; c < cols; c++) {
                fRow.createCell(c).setCellValue(FILL_VALUE);
                vRow.createCell(c).setCellValue(FILL_VALUE);
                fRow.createCell(c + cols).setCellFormula(
                    String.format(CREATE_STR
                        , frmula
                        , CellReference.convertNumToColString(cols - 1)
                        , r + 1
                    )
                );
                vRow.createCell(c + cols).setCellValue((FILL_VALUE * (r + 1) * cols) / divide);
            }
        }   
    }

    @Override
    public void createRandomExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols, Random rand) {
        double total = 0.0;
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            double[] vals = new double[cols];
            String frmula = r % 2 == 0 ? "SUM" : "AVERAGE";
            double divide = r % 2 == 0 ? 1 : cols * r + cols;
            for (int c = 0; c < cols; c++) { 
                vals[c] = (double) rand.nextInt(rows * cols);
                total += vals[c];
            }
            for (int c = 0; c < cols; c++) {
                fRow.createCell(c).setCellValue(vals[c]);
                vRow.createCell(c).setCellValue(vals[c]);
                fRow.createCell(c + cols).setCellFormula(
                    String.format(CREATE_STR
                        , frmula
                        , CellReference.convertNumToColString(cols - 1)
                        , r + 1
                    )
                );
                vRow.createCell(c + cols).setCellValue(total / divide);
            }
        }
    }

    @Override
    public void createCalcSheet(Table fSheet, Table vSheet, int rows, int cols) throws IOException {
        for (int r = 0; r < rows; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            String frmula = r % 2 == 0 ? "SUM" : "AVERAGE";
            double divide = r % 2 == 0 ? 1 :  ((r + 1) * cols);
            for (int c = 0; c < cols; c++) {
                fRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
                vRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
                fRow.getOrCreateCell(c + cols).setFormula(String.format(CREATE_STR, frmula, CellReference.convertNumToColString(cols - 1), r + 1));
                vRow.getOrCreateCell(c + cols).setFloatValue((FILL_VALUE * (r + 1) * cols) / divide);
            }
        }
    }

    @Override
    public void createRandomCalcSheet(Table fSheet, Table vSheet, int rows, int cols, Random rand) throws IOException {
        double total = 0.0;
        for (int r = 0; r < rows; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            double[] vals = new double[cols];
            String frmula = r % 2 == 0 ? "SUM" : "AVERAGE";
            double divide = r % 2 == 0 ? 1 :  ((r + 1) * cols);
            for (int c = 0; c < cols; c++) { 
                vals[c] = (double) rand.nextInt(rows * cols);
                total += vals[c];
            }
            for (int c = 0; c < cols; c++) {
                fRow.getOrCreateCell(c).setFloatValue(vals[c]);
                vRow.getOrCreateCell(c).setFloatValue(vals[c]);
                fRow.getOrCreateCell(c + cols).setFormula(String.format(CREATE_STR, frmula, CellReference.convertNumToColString(cols - 1), r + 1));
                vRow.getOrCreateCell(c + cols).setFloatValue(total / divide);
            }
        }
    }

}
