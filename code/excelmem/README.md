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
    
    - Outputs an excel file called `memory.xlsx` with the following memory measurements (in megabytes) for both formula-value and value-only spreadsheets:

        - `peak_nonpaged_pool`  : The peak nonpaged pool usage
        - `peak_paged_pool`     : The peak paged pool usage
        - `peak_pagefile`       : The peak value in bytes of the Commit Charge during the lifetime of this process
        - `nonpaged_pool`       : The nonpaged pool usage
        - `paged_pool`          : The paged pool usage
        - `peak_wset`           : The peak working set size
        - `pagefile`            : The Commit Charge value
        - `private`             : Same as `pagefile`
        - `wset`                : The working set size
        - `rss`                 : The resident set size
        - `uss`                 : The unique set size
        - `vms`                 : The virtual memory size

        See the [PROCESS_MEMORY_COUNTERS_EX](https://docs.microsoft.com/en-us/windows/win32/api/psapi/ns-psapi-process_memory_counters_ex) structure doc for more detailed explanations.

- `excelmemvis.py`     : Same as `excelmem.py`, but modified so that it may be called from a jupyter notebook.

- `memdata.py`         : A helper class for collecting memory measurements.