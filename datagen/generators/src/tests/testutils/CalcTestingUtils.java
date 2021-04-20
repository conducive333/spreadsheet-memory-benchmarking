package testutils;

import org.mockito.Mockito;

import static org.mockito.ArgumentMatchers.anyDouble;
import static org.mockito.Mockito.when;

import com.github.jferard.fastods.TableRowImpl;
import com.github.jferard.fastods.TableCell;
import com.github.jferard.fastods.Table;

import java.util.function.BiFunction;
import java.util.OptionalLong;
import java.io.IOException;
import java.nio.file.Path;
import java.io.File;

import creator.Creatable;
import creator.Creator;

/**
 * 
 * 3/31/21 NOTE: jopendocument and SimpleODS have been 
 * very inconsistent when reading .ods files and don't
 * seem suitable for testing (at least at the time of 
 * this writing). The current workaround uses mocks to
 * check if functions are being called with the correct
 * parameters.
 * 
 * */
public class CalcTestingUtils extends TestingUtils {

    /**
     * Creates a formula-value spreadsheet and its corresponding value-only spreadsheet
     * using `creatable`. The spreadsheet is saved as a .ods file.
     * 
     * @param creatable
     * @param rows
     * @param cols
     * @param seed
     * @return An array, ARR, of two files. ARR[0] is the value-only file and ARR[1] is
     * the formula-value file.
     */
    public static File[] createCalcFiles (Creatable creatable, int rows, int cols, OptionalLong seed) {
        Creator.createCalcSheet(creatable, TestingUtils.F_FOLDER.toString(), TestingUtils.V_FOLDER.toString(), rows, cols, seed);
        return new File[] {
            Path.of(TestingUtils.V_FOLDER.toString(), "vo-" + rows + ".ods").toFile(),
            Path.of(TestingUtils.F_FOLDER.toString(), "fv-" + rows + ".ods").toFile()
        };
    }

    /**
     * Checks the number of rows and number of cells in each row. Verifies that the 
     * first `cols` columns contain values, (but does not necessarily check the values
     * themselves), and verifies that the next `cols` columns have the correct formula
     * structure.
     * 
     * @param creatable
     * @param rows
     * @param expectedRows
     * @param cols
     * @param expectedCols
     * @param seed
     * @param uppr
     * @param getExpectedFormula
     */
    public static void integrationTest (Creatable creatable, int rows, int expectedRows, int cols, int expectedCols, OptionalLong seed, int uppr, BiFunction<Integer, Integer, String> getExpectedFormula) {
        try {

            // Set up testing mocks
            TableCell[][]   fCellMocks  = new TableCell[expectedRows][expectedCols];
            TableCell[][]   vCellMocks  = new TableCell[expectedRows][expectedCols];
            Creatable       createMock  = Mockito.spy(creatable);
            Table           fSheetMock  = Mockito.mock(Table.class);
            Table           vSheetMock  = Mockito.mock(Table.class);
            for (int r = 0; r < expectedRows; r++) {
                TableRowImpl fRowMock = Mockito.mock(TableRowImpl.class);
                TableRowImpl vRowMock = Mockito.mock(TableRowImpl.class);
                when(fSheetMock.getRow(r)).thenReturn(fRowMock);
                when(vSheetMock.getRow(r)).thenReturn(vRowMock);
                for (int c = 0; c < expectedCols; c++) {
                    fCellMocks[r][c] = Mockito.mock(TableCell.class);
                    vCellMocks[r][c] = Mockito.mock(TableCell.class);
                    when(fRowMock.getOrCreateCell(c)).thenReturn(fCellMocks[r][c]);
                    when(vRowMock.getOrCreateCell(c)).thenReturn(vCellMocks[r][c]);
                }
            }
            
            // Call the real method
            if (seed.isPresent()) {
                createMock.createRandomCalcSheet(fSheetMock, vSheetMock, seed.getAsLong());
            } else {
                createMock.createCalcSheet(fSheetMock, vSheetMock);
            }

            // Check function calls and parameters
            for (int r = 0; r < rows; r++) {
                
                // Each row should have only been accessed one time
                Mockito.verify(fSheetMock, Mockito.times(1)).getRow(r);
                Mockito.verify(vSheetMock, Mockito.times(1)).getRow(r);


                for (int c = 0; c < cols; c++) {

                    // Each cell should have only been set one time
                    Mockito.verify(fCellMocks[r][c], Mockito.times(1)).setFloatValue(anyDouble());
                    Mockito.verify(vCellMocks[r][c], Mockito.times(1)).setFloatValue(anyDouble());
                    Mockito.verify(fCellMocks[r][c + cols], Mockito.times(1)).setFormula(getExpectedFormula.apply(r, c));
                    Mockito.verify(vCellMocks[r][c + cols], Mockito.times(1)).setFloatValue(anyDouble());

                }
            }

        } catch (IOException e) { e.printStackTrace(); }
    }
    
}
