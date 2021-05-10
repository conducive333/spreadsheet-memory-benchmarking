package sums.specialsums;

import sums.BaseSum;

public abstract class BaseSpecialSum extends BaseSum {
    
  public static int MAX_V_ROWS = 1000000;

  public BaseSpecialSum(int uppr) {
    super(uppr);
  }
  
  public static void setMaxRows (int val) {
    MAX_V_ROWS = val;
  }

}
