package sums;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.ss.util.CellReference;

import com.github.jferard.fastods.TableRowImpl;
import com.github.jferard.fastods.Table;

import java.io.IOException;
import java.util.Random;

import creator.Creatable;

public class RunningSum implements Creatable {
    /**
     * Creates a spreadsheet with the following structure:
     * 
     * 0    |   A   |       B       |
     * ------------------------------
     * 1    |   ?   |   =SUM(A1:A1) |
     * 2    |   ?   |   =SUM(A1:A2) |
     * ...  |   ... |   ...         |
     * N    |   ?   |   =SUM(A1:AN) |
     * 
     * ? = a random value (or a placeholder value if no
     * random seed is specified). The COLS parameter 
     * controls the number of columns with values AND
     * the number of columns with formulae. For example,
     * if COLS = 2, then there will be 2 columns of values
     * and 2 columns of formulae.
     * 
     */

    private static final double FILL_VALUE = 1.0;

    @Override
    public void createExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols) {
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.createCell(c).setCellValue(FILL_VALUE);
                vRow.createCell(c).setCellValue(FILL_VALUE);
                fRow.createCell(c + cols).setCellFormula(String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), r + 1));
                vRow.createCell(c + cols).setCellValue(FILL_VALUE * (r + 1) * cols);
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
            for (int c = 0; c < cols; c++) { 
                vals[c] = (double) rand.nextInt(rows * cols);
                total += vals[c];
            }
            for (int c = 0; c < cols; c++) {
                fRow.createCell(c).setCellValue(vals[c]);
                vRow.createCell(c).setCellValue(vals[c]);
                fRow.createCell(c + cols).setCellFormula(String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), r + 1));
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
                fRow.getOrCreateCell(c + cols).setFormula(String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), r + 1));
                vRow.getOrCreateCell(c + cols).setFloatValue((double) (FILL_VALUE * (r + 1) * cols));
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
            for (int c = 0; c < cols; c++) { 
                vals[c] = (double) rand.nextInt(rows * cols);
                total += vals[c];
            }
            for (int c = 0; c < cols; c++) {
                fRow.getOrCreateCell(c).setFloatValue(vals[c]);
                vRow.getOrCreateCell(c).setFloatValue(vals[c]);
                fRow.getOrCreateCell(c + cols).setFormula(String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), r + 1));
                vRow.getOrCreateCell(c + cols).setFloatValue(total);
            }
        }
    }

}
