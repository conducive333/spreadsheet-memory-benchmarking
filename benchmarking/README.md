# Memory Benchmarking Code

## Table of Contents
- Directories:
    - `datagen` : Contains the code for generating experimental datasets. It is capable of generating `.ods` and `.xlsx` files. 
    - `excelmem` : Contains the code for benchmarking Excel's memory consumption.
    - `libremem` : Contains the code for benchmarking LibreOffice's memory consumption.

- Files:
    - `memscript.py`        : The script for running memory benchmark experiments.
    - `visualizer.ipynb`    : Same as `memscript.py`, except that experimental results may be viewed as soon as they become available .
    - `util.py`             : Contains helper code for running experiments.

## Explanation of Script Arguments
- `CONFIG_ARGS` : controls the types of spreadsheets to create. Each field is described in detail in the `datagen` project's README.
- `OUTPUT_PATH` : this specifies where the output of the benchmarking run should be placed. Any folders in the path that don't exist will be created automatically. This path will be appended directly after the path to the project's root directory.
- `INTEGER_ARG` : If, in your `CONFIG_ARGS`, you specify `xlsx=True`, then this is the number of trials to perform for each Excel spreadsheet. Otherwise this is the number of poll seconds to use for each LibreOffice Calc spreadsheet. The number of trials and the number of poll seconds are described in the README's of excelmem and libremem respectively.
- `SOFFICE_DIR` : The absolute path to soffice.

## How to Run Each Script
- `memscript.py`
    1. If you cloned this repository, it should have come with a file, `main.jar`, in the root directory. If it does not exist, export `datagen` as a jar before continuing.
    2. Once you have the jar file, fill in each script argument.
    3. Run the script.

- `visualizer.ipynb`
    1. If you cloned this repository, it should have come with a file, `main.jar`, in the root directory. If it does not exist, export `datagen` as a jar before continuing.
    2. Once you have the jar file, find the cell which sets the script's arguments and fill them in.
    3. Run all cells in order.

## Notes
- This code is intended to be run on a Windows machine. Errors may appear on other systems.