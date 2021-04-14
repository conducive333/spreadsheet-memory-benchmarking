package sums;

import java.util.Random;
import java.util.Deque;
import java.util.List;

public abstract class BaseSum {
    
    protected static final double FILL_VALUE = 1.0;

    /**
     * Fills the input deque with `size` random
     * values from the interval [0, bound).
     * 
     * @param dequeToFill
     * @param size
     * @param rand
     * @param bound
     * @return The sum of `dequeToFill`.
     */
    protected double randomlyFillDeque (Deque<Double> dequeToFill, int size, Random rand, int bound) {
        double total = 0.0;
        for (int i = 0; i < size; i++) {
            int num = rand.nextInt(bound);
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
     * @param bound
     * @return The sum of `dequeToFill`.
     */
    protected double randomlyFillList (List<Double> listToFill, int size, Random rand, int bound) {
        double total = 0.0;
        for (int i = 0; i < size; i++) {
            int num = rand.nextInt(bound);
            listToFill.add((double) num);
            total += num;
        }
        return total;
    }

}
