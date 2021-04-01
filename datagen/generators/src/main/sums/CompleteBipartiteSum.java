package sums;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.ss.util.CellReference;

import com.github.jferard.fastods.TableRowImpl;
import com.github.jferard.fastods.Table;

import java.util.ArrayDeque;
import java.io.IOException;
import java.util.Random;
import java.util.Deque;

import creator.Creatable;

public class CompleteBipartiteSum implements Creatable {
    /**
     * Creates a spreadsheet with the following structure:
     * 
     * 0    |   A   |       B       |
     * ------------------------------
     * 1    |   ?   |   =SUM(A1:AN) |
     * 2    |   ?   |   =SUM(A1:AN) |
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
                fRow.createCell(c + cols).setCellFormula(String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), rows));
                vRow.createCell(c + cols).setCellValue(FILL_VALUE * rows * cols);
            }
        }   
    }

    @Override
    public void createRandomExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols, Random rand) {
        double total = 0.0;
        Deque<Double> values = new ArrayDeque<>();
        for (int i = 0; i < rows * cols; i++) {
            int num = rand.nextInt(rows * cols);
            values.add((double) num);
            total += num;
        }
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                double num = values.pop();
                fRow.createCell(c).setCellValue(num);
                vRow.createCell(c).setCellValue(num);
                fRow.createCell(c + cols).setCellFormula(String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), rows));
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
                fRow.getOrCreateCell(c + cols).setFormula(String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), rows));
                vRow.getOrCreateCell(c + cols).setFloatValue((double) (FILL_VALUE * rows * cols));
            }
        }
    }

    @Override
    public void createRandomCalcSheet(Table fSheet, Table vSheet, int rows, int cols, Random rand) throws IOException {
        double total = 0.0;
        Deque<Integer> values = new ArrayDeque<>();
        for (int i = 0; i < rows * cols; i++) {
            int num = rand.nextInt(rows * cols);
            values.add(num);
            total += num;
        }
        for (int r = 0; r < rows; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            for (int c = 0; c < cols; c++) {
                double num = values.pop();
                fRow.getOrCreateCell(c).setFloatValue(num);
                vRow.getOrCreateCell(c).setFloatValue(num);
                fRow.getOrCreateCell(c + cols).setFormula(String.format("SUM(A1:%s%d)", CellReference.convertNumToColString(cols - 1), rows));
                vRow.getOrCreateCell(c + cols).setFloatValue(total);
            }
        }
    }

}
