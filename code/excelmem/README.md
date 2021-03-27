# Excel Memory Benchmarking

## Table of Contents:
- `excelmem.py`

    - Takes in the following parameters as input:

        - `INPUTS_PATH` : A directory with at least two folders of experimental Excel files. All input files are assumed to have a name that matches:
            ```
            <prefix>-<size>.xlsx
            ```
        - `FV_INPUTDIR` : The name of the folder with formula-value datasets (assumed to be inside `INPUTS_PATH`)
        - `VO_INPUTDIR` : The name of the folder with formula-value datasets (assumed to be inside `INPUTS_PATH`)
        - `OUTPUT_PATH` : The path to a directory where all experimental folders will be created
        - `OUTDIR_NAME` : The name of the directory to write results to (overwites any existing file(s) with the same name without warning)
        - `TRIALS`      : Number of trials to perform for each run
    
    - Outputs an excel file called `memory.xlsx` with the following schema:

        - `Rows` : The number of rows in the file
        - `Value Peak WSS (MB)` : The peak working set size in megabytes for value-only spreadsheets
        - `Value WSS (MB)`  : The working set size in megabytes for value-only spreadsheets
        - `Value USS (MB)`  : The unique set size in megabytes for value-only spreadsheets
        - `Formula Peak WSS (MB)` : The peak working set size in megabytes for formula-value spreadsheets
        - `Formula WSS (MB)` : The working set size in megabytes for formula-value spreadsheets
        - `Formula USS (MB)` : The unique set size in megabytes for formula-value spreadsheets

- `excelmemvis.py`     : Same as `excelmem.py`, but modified so that it may be called from a jupyter notebook.

- `memdata.py`         : A helper class for collecting memory measurements.