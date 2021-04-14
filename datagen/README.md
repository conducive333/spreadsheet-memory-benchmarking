# Spreadsheet Dataset Generator

## Overview
- This script will generate two types of spreadsheet datasets:
    1. Formula-value: A spreadsheet with both values and formulae. The layout of the spreadsheet is determined by the `INST` parmeter. Currently, the possible values are:
        - "CompleteBipartiteSum"
        - "CompleteBipartiteSumWithConstant"
        - "MixedRangeSum"
        - "NoEdgeSum"
        - "OverlappingSum"
        - "RunningSum"
        - "SingleCellSum"
        - "CompleteBipartiteVlookup"    (only 1 column supported)
        - "SameCellVlookup"             (only 1 column supported, double check structure before using)
        - "SingleCellVlookup"           (only 1 column supported)
    
        **NOTE:** The values above are the names of classes in the source code. The comments of each class will show you the respective spreadsheet it will create.

    2. Value-only: Same as the formula-value, but all formulae are replaced by their evaluated result.

## How to use this script:

### Method 1 (Preferred):
1. Download the `.jar` file associated with this project (should be on the "Releases" page).
2. Fill in the desired script arguments in `createdata.py`.
3. Run the script!

### Method 2:
1. Create a file named `config` (no extension) in the root directory of the entire project (i.e. the directory where requirements.txt is). The config file should have the following fields:
    - `INST`  : The name of the class to use.
    - `PATH`  : Specifies where to create the datasets.
    - `SEED`  : The seed to use when generating random values. If left empty, then sheets will be filled with the same placeholder value (currently set to 1).
    - `XLSX`  : If true, creates `.xlsx` spreadsheets. Otherwise, creates `.ods` spreadsheets.
    - `STEP`  : The number of rows to increment by on the next iteration.
    - `ROWS`  : The starting number of rows.
    - `COLS`  : The number of columns to create.
    - `ITRS`  : The number of iterations to perform.
    - `POOL`  : The number of threads to use. If set to 1, then the main thread will be used (i.e. no multithreading).

2. Run the script from `Main.java`.

## Sample `config` file:
```
INST=CompleteBipartiteSum
PATH=datasets
SEED=42
XLSX=true
STEP=1
ROWS=100
COLS=1
ITER=1
POOL=1
```

## Notes
- **WARNING:** There is currently no mechanism that detects if any of the config fields are invalid.
