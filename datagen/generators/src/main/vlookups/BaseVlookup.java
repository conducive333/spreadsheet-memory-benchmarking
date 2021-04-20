package vlookups;

import java.util.Collections;
import java.util.ArrayList;
import java.util.Random;
import java.util.List;

public abstract class BaseVlookup {

    public static final double FILL_VALUE = 1.0;
    
    public final int rows;
    public final int cols;
    public final int uppr;

    /**
     * 
     * @param rows
     * @param cols
     * @param uppr
     */
    public BaseVlookup (int rows, int cols, int uppr) {
        this.rows = rows;
        this.cols = cols;
        this.uppr = uppr;
    }

    /**
     * Creates a list from [0, `size`) and shuffles its ordering.
     * 
     * @param rand
     * @param size
     * @return The list as described above.
     */
    protected List<Double> getShuffledConsecutiveNumbers (Random rand, int size) {
        List<Double> nums = new ArrayList<>(size);
        for (int r = 0; r < size; r++) { nums.add((double) r); }
        Collections.shuffle(nums, rand);
        return nums;
    }

}
