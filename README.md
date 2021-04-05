# Memory Benchmarking Code

## Table of Contents
- Files:
    - `createdata.py`       : A script that just generates spreadsheets.
    - `memscript.py`        : A script for running memory benchmarking experiments.
    - `visualizer.ipynb`    : Same as `memscript.py`, except that experimental results are plotted as soon as they become available.

- Directories:
    - `datagen`
        - Contains the Java code for generating experimental datasets. It supports `.ods` and `.xlsx` formats.
    - `pymem`
        - `excelmem`    : Contains the code for benchmarking Excel's memory consumption.
        - `libremem`    : Contains the code for benchmarking LibreOffice's memory consumption.
        - `utils`       : Contains various Python helper code.

## Explanation of Script Arguments
- `CONFIG_ARGS` : controls the types of spreadsheets to create. This parameter consists of several sub-fields all of which are detailed in the `datagen` project's README.
    - **IMPORTANT:** If `PATH` does not point to an existing directory, then the script will use all the config fields to generate the specified dataset from scratch. However, if `PATH` points to a dataset that was created from a previous run or by `createdata.py`, then the script will use it for benchmarking and NOT re-create it. In this case, all other config fields except `XLSX` will be ignored. This can be useful if you want to re-use a dataset and avoid the trouble of re-generating a potentially large series of spreadsheet files. 
    - **NOTE:** The `PATH` variable will be appended after the absolute path to the project's root directory.
- `OUTPUT_PATH` : this specifies where the benchmarking results should be placed. Any folders in the path that don't exist will be created automatically. This path will be appended directly after the path to the project's root directory.
- `INTEGER_ARG` : If, in your `CONFIG_ARGS`, you specify `xlsx=True`, then this is the number of trials to perform for each Excel spreadsheet. Otherwise this is the number of poll seconds to use for each LibreOffice Calc spreadsheet. The number of trials and the number of poll seconds are described in the README's of `pymem/excelmem` and `pymem/libremem` respectively.
- `SOFFICE_DIR` : The absolute path to soffice.

## How to Run Each Script
- `createdata.py`
    1. Install all python dependencies specified in `requirements.txt`.
    2. Download `datagen.jar` from this project's "Releases" page and place it in the root directory of this project.
    3. Fill in each script argument with the desired parameters.
    4. Run the script.

- `memscript.py`
    1. Install all python dependencies specified in `requirements.txt`.
    2. If you'd like to create your own datasets, download `datagen.jar` from this project's "Releases" page and place it in the root directory of this project. Otherwise, feel free to skip this step.
    3. Fill in each script argument with the desired parameters.
    4. Run the script.

- `visualizer.ipynb`
    1. Install all python dependencies specified in `requirements.txt`.
    2. If you'd like to create your own datasets, download `datagen.jar` from this project's "Releases" page and place it in the root directory of this project. Otherwise, feel free to skip this step.
    3. Find the cell which sets the script's arguments and fill them in with the desired parameters.
    4. Run all cells in order (i.e. from top to bottom).

## Notes
- This code is intended to be run on a Windows machine. Errors may appear on other systems.