# LibreOffice Memory Benchmarking

## Table of Contents:
- `mem.py`

    - Takes in the following parameters as input:

        - `INPUTS_PATH` : A directory with at least two folders of experimental Calc files. All input files are assumed to have a name that matches:
            ```
            <prefix>-<size>.ods
            ```
        - `OUTPUT_PATH` : Specifies where to dump the output of the script.
        - `FV_INPUTDIR` : The name of the folder with formula-value datasets (assumed to be inside `INPUTS_PATH`).
        - `VO_INPUTDIR` : The name of the folder with value-only datasets (assumed to be inside `INPUTS_PATH`).
        - `POLLSECONDS` : The minimum number of seconds to wait before collecting another memory measurement.
        - `SOFFICEPATH` : The absolute path to soffice.
    
    - Outputs:
    
        - An excel file called `memory.xlsx` with the following memory measurements (in megabytes) for both formula-value and value-only spreadsheets:
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
        
        - Two directories, `fv-mem-curve` and `vo-mem-curve`, that contain JSON files. Each JSON file has the following format:
            ```
            { <timestamp> : 
              {
                peak_nonpaged_pool  : <value in bytes>
                peak_paged_pool     : <value in bytes>
                peak_pagefile       : <value in bytes>
                nonpaged_pool       : <value in bytes>
                paged_pool          : <value in bytes>
                peak_wset           : <value in bytes>
                pagefile            : <value in bytes>
                private             : <value in bytes>
                wset                : <value in bytes>
                rss                 : <value in bytes>
                uss                 : <value in bytes>
                vms                 : <value in bytes>
              }
            } 
            ```

- `memvis.py`
  - Same as `mem.py`, but modified so that it may communicate intermediate results to a jupyter notebook.
