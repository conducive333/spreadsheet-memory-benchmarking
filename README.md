# Memory Benchmarking Code

## Table of Contents
- Directories:
    - `datagen`
        - Contains the code for generating experimental datasets. It is capable of generating `.ods` and `.xlsx` files.
    - `pymem`
        - `excelmem`    : Contains the code for benchmarking Excel's memory consumption.
        - `libremem`    : Contains the code for benchmarking LibreOffice's memory consumption.
        - `utils`       : Contains various Python helper code.

- Files:
    - `memscript.py`        : The script for running memory benchmark experiments.
    - `visualizer.ipynb`    : Same as `memscript.py`, except that experimental results are plotted as soon as they become available.

## Explanation of Script Arguments
- `CONFIG_ARGS` : controls the types of spreadsheets to create. Each field is described in detail in the `datagen` project's README. If you'd like to reuse a dataset that already exists (and prevent the script from re-generating a potentially large dataset), then pass its path to `PATH` and set the `XLSX` parameter appropriately. In the case that `PATH` already points to a pre-existing directory of benchmarking files, the other config fields will be ignored (with the exception of `XLSX`). The `PATH` variable will be appended after the path to the project's root directory.
- `OUTPUT_PATH` : this specifies where the output of the benchmarking run should be placed. Any folders in the path that don't exist will be created automatically. This path will be appended directly after the path to the project's root directory.
- `INTEGER_ARG` : If, in your `CONFIG_ARGS`, you specify `xlsx=True`, then this is the number of trials to perform for each Excel spreadsheet. Otherwise this is the number of poll seconds to use for each LibreOffice Calc spreadsheet. The number of trials and the number of poll seconds are described in the README's of excelmem and libremem respectively.
- `SOFFICE_DIR` : The absolute path to soffice.

## How to Run Each Script
- `memscript.py`
    1. Install all python dependencies specified in `requirements.txt`
    2. Download `main.jar` from this project's 'Releases' page on Github and place it in the root directory of this project.
    3. Once you have the jar file, fill in each script argument.
    4. Run the script.

- `visualizer.ipynb`
    1. Install all python dependencies specified in `requirements.txt`
    2. Download `main.jar` from this project's 'Releases' page on Github and place it in the root directory of this project.
    3. Once you have the jar file, find the cell which sets the script's arguments and fill them in.
    4. Run all cells in order (i.e. from top to bottom).

## Notes
- This code is intended to be run on a Windows machine. Errors may appear on other systems.