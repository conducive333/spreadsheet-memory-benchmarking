package sums;

import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFRow;

import com.github.jferard.fastods.TableRowImpl;
import com.github.jferard.fastods.Table;

import java.io.IOException;
import java.util.Random;

import creator.Creatable;

public class NoEdgeSum extends BaseSum implements Creatable {
    /**
     * Creates a spreadsheet with the following structure:
     * 
     * 0    |   A   |       B       |
     * ------------------------------
     * 1    |   ?   |   =SUM(?)     |
     * 2    |   ?   |   =SUM(?)     |
     * ...  |   ... |   ...         |
     * N    |   ?   |   =SUM(?)     |
     * 
     * ? = a random value (or a placeholder value if no
     * random seed is specified). The COLS parameter 
     * controls the number of columns with values AND
     * the number of columns with formulae. For example,
     * if COLS = 2, then there will be 2 columns of values
     * and 2 columns of formulae.
     */

    private static final String CREATE_STR = "SUM(%f)";

    public NoEdgeSum (int rows, int cols, int uppr) {
        super(rows, cols, uppr);
    }

    @Override
    public void createExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet) {
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.createCell(c).setCellValue(FILL_VALUE);
                vRow.createCell(c).setCellValue(FILL_VALUE);
                fRow.createCell(c + cols).setCellFormula(String.format(CREATE_STR, FILL_VALUE));
                vRow.createCell(c + cols).setCellValue(FILL_VALUE);
            }
        }
    }

    @Override
    public void createRandomExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, long seed) {
        Random rand = new Random(seed);
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                double num = (double) rand.nextInt(uppr);
                fRow.createCell(c).setCellValue(num);
                vRow.createCell(c).setCellValue(num);
                fRow.createCell(c + cols).setCellFormula(String.format(CREATE_STR, num));
                vRow.createCell(c + cols).setCellValue(num);
            }
        }
    }

    @Override
    public void createCalcSheet(Table fSheet, Table vSheet) throws IOException {
        for (int r = 0; r < rows; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            for (int c = 0; c < cols; c++) {
                fRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
                vRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
                fRow.getOrCreateCell(c + cols).setFormula(String.format(CREATE_STR, FILL_VALUE));
                vRow.getOrCreateCell(c + cols).setFloatValue(FILL_VALUE);
            }
        }
    }

    @Override
    public void createRandomCalcSheet(Table fSheet, Table vSheet, long seed) throws IOException {
        Random rand = new Random(seed);
        for (int r = 0; r < rows; r++) {
            TableRowImpl fRow = fSheet.getRow(r);
            TableRowImpl vRow = vSheet.getRow(r);
            for (int c = 0; c < cols; c++) {
                double num = (double) rand.nextInt(uppr);
                fRow.getOrCreateCell(c).setFloatValue(num);
                vRow.getOrCreateCell(c).setFloatValue(num);
                fRow.getOrCreateCell(c + cols).setFormula(String.format(CREATE_STR, num));
                vRow.getOrCreateCell(c + cols).setFloatValue(num);
            }
        }
    }

}