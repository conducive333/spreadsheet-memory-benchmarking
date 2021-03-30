# Excel Memory Benchmarking

## Table of Contents:
- `mem.py`

    - Takes in the following parameters as input:

        - `INPUTS_PATH` : A directory with at least two folders of experimental Excel files. All input files are assumed to have a name that matches:
            ```
            <prefix>-<size>.xlsx
            ```
        - `OUTPUT_PATH` : Specifies where to dump the output of the script.
        - `FV_INPUTDIR` : The name of the folder with formula-value datasets (assumed to be inside `INPUTS_PATH`).
        - `VO_INPUTDIR` : The name of the folder with value-only datasets (assumed to be inside `INPUTS_PATH`).
        - `TOTL_TRIALS` : The number of trials to perform for each run. A trial is defined as opening a workbook, measuring its memory consumption, and closing the workbook. If the number of trials is greater than two, then the max and min are removed and the average measurement is taken over the resulting data points. Otherwise, the average measurement is taken without removing the max and min.
    
    - Outputs an excel file called `memory.xlsx` with the following memory measurements (in megabytes) for both formula-value and value-only spreadsheets:

        - `peak_nonpaged_pool`  : The peak nonpaged pool usage.
        - `peak_paged_pool`     : The peak paged pool usage.
        - `peak_pagefile`       : The peak value in bytes of the Commit Charge during the lifetime of this process.
        - `nonpaged_pool`       : The nonpaged pool usage.
        - `paged_pool`          : The paged pool usage.
        - `peak_wset`           : The peak working set size.
        - `pagefile`            : The Commit Charge value.
        - `private`             : Same as `pagefile`.
        - `wset`                : The working set size.
        - `rss`                 : The resident set size.
        - `uss`                 : The unique set size.
        - `vms`                 : The virtual memory size.

        See the [PROCESS_MEMORY_COUNTERS_EX](https://docs.microsoft.com/en-us/windows/win32/api/psapi/ns-psapi-process_memory_counters_ex) structure doc for more detailed explanations.

- `memvis.py`
    - Same as `mem.py`, but modified so that it may communicate intermediate results to a jupyter notebook.
