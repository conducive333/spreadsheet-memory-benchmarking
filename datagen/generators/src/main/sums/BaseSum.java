package sums;

import java.util.Random;
import java.util.Deque;
import java.util.List;

public abstract class BaseSum {
    
    public static final double FILL_VALUE = 1.0;

    public final int uppr;

    /**
     * 
     * @param uppr
     */
    public BaseSum (int uppr) {
        this.uppr = uppr;
    }

    /**
     * Fills the input deque with `size` random
     * values from the interval [0, bound).
     * 
     * @param dequeToFill
     * @param size
     * @param rand
     * @return The sum of `dequeToFill`.
     */
    protected double randomlyFillDeque (Deque<Double> dequeToFill, int size, Random rand) {
        double total = 0.0;
        for (int i = 0; i < size; i++) {
            int num = rand.nextInt(this.uppr);
            dequeToFill.add((double) num);
            total += num;
        }
        return total;
    }

    /**
     * Fills the input list with `size` random
     * values from the interval [0, bound).
     * 
     * @param listToFill
     * @param size
     * @param rand
     * @return The sum of `dequeToFill`.
     */
    protected double randomlyFillList (List<Double> listToFill, int size, Random rand) {
        double total = 0.0;
        for (int i = 0; i < size; i++) {
            int num = rand.nextInt(this.uppr);
            listToFill.add((double) num);
            total += num;
        }
        return total;
    }

}
