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

public class SingleCellSum implements Creatable {
    
    private static final double FILL_VALUE = 1.0;

    @Override
    public void createExcelSheet (SXSSFSheet fSheet, SXSSFSheet vSheet, int rows, int cols) {
        for (int r = 0; r < rows; r++) {
            SXSSFRow fRow = fSheet.createRow(r);
            SXSSFRow vRow = vSheet.createRow(r);
            for (int c = 0; c < cols; c++) {
                String col = CellReference.convertNumToColString(c);
                fRow.createCell(c).setCellValue(FILL_VALUE);
                vRow.createCell(c).setCellValue(FILL_VALUE);
                fRow.createCell(c + cols).setCellFormula(String.format("SUM(%s%d:%s%d)", col, r + 1, col, r + 1));
                vRow.createCell(c + cols).setCellValue(1);
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
                String col = CellReference.convertNumToColString(c);
                fRow.createCell(c).setCellValue(num);
                vRow.createCell(c).setCellValue(num);
                fRow.createCell(c + cols).setCellFormula(String.format("SUM(%s%d:%s%d)", col, r + 1, col, r + 1));
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
                String col = CellReference.convertNumToColString(c);
                fRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
                vRow.getOrCreateCell(c).setFloatValue(FILL_VALUE);
                fRow.getOrCreateCell(c + cols).setFormula(String.format("SUM(%s%d:%s%d)", col, r + 1, col, r + 1));
                vRow.getOrCreateCell(c + cols).setFloatValue(rows * cols);
            }
        }
    }

    @Override
    public void createRandomCalcSheet(Table fSheet, Table vSheet, int rows, int cols, Random rand) throws IOException {
        int total = 0;
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
                int num = values.pop();
                String col = CellReference.convertNumToColString(c);
                fRow.getOrCreateCell(c).setFloatValue(num);
                vRow.getOrCreateCell(c).setFloatValue(num);
                fRow.getOrCreateCell(c + cols).setFormula(String.format("SUM(%s%d:%s%d)", col, r + 1, col, r + 1));
                vRow.getOrCreateCell(c + cols).setFloatValue(total);
            }
        }
    }

}
