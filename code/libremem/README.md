# LibreOffice Memory Benchmarking

## Table of Contents:
- `libremem.py`

    - Takes in the following parameters as input:

        - `INPUTS_PATH` : A directory with at least two folders of experimental Calc files. All input files are assumed to have a name that matches:
            ```
            <prefix>-<size>.ods
            ```
        - `FV_INPUTDIR` : The name of the folder with formula-value datasets (assumed to be inside `INPUTS_PATH`)
        - `VO_INPUTDIR` : The name of the folder with formula-value datasets (assumed to be inside `INPUTS_PATH`)
        - `OUTPUT_PATH` : The path to a directory where all experimental folders will be created
        - `OUTDIR_NAME` : The name of the directory to write results to (overwites any existing file(s) with the same name without warning)
        - `POLLSECONDS` : The minimum number of seconds to wait before collecting another memory measurement
        - `SOFFICE_PATH`: The absolute path to soffice
    
    - Outputs:
    
        - An excel file called `memory.xlsx` with the following schema:
            - `Rows` : The number of rows in the file
            - `Value Peak WSS (MB)` : The peak working set size in megabytes for value-only spreadsheets
            - `Value WSS (MB)`  : The working set size in megabytes for value-only spreadsheets
            - `Value USS (MB)`  : The unique set size in megabytes for value-only spreadsheets
            - `Formula Peak WSS (MB)` : The peak working set size in megabytes for formula-value spreadsheets
            - `Formula WSS (MB)` : The working set size in megabytes for formula-value spreadsheets
            - `Formula USS (MB)` : The unique set size in megabytes for formula-value spreadsheets
        
        - Two directories, `fv-mem-curve` and `vo-mem-curve`, that contain JSON files. Each JSON file has the following format:
            ```
            { <timestamp> : 
              {
                'Peak WSS' : Peak working set size in bytes,
                'WSS' : Working set size in bytes,
                'USS' : Unique set size in bytes,
              }
            } 
            ```

- `librememvis.py`     : Same as `libremem.py`, but modified so that it may be called from a jupyter notebook.
