import pymem.utils.memdata as memdata
import pymem.utils.excel as excel
import subprocess
import traceback
import datetime
import pathlib
import random
import pandas
import time
import sys
import os

# NOTE: Starting a fresh Excel process before the trials loop is much 
# faster than spawning one at the beginning of each iteration of the 
# loop. The former method also seems to give the same results as the 
# latter method.

def run(path, trials, prefix, results):
    """
    Measures the memory consumption of all .xlsx files specified by `path`.

    Parameters:
    -----------        
        path : path-like
            The path to a directory of benchmarking files.
        
        trials : int
            The number of trials to perform.

        prefix : str
            This will be prepended to the keys of the dictionary 
            returned by `psutil.memory_full_info()`.

        results : dict
            A dictionary to store the final memory measurements in.
    """

    # Collect and shuffle input files
    pairs = [(f, int(f[f.index('-')+1:f.index('.')])) for f in os.listdir(path) if f.endswith(".xlsx")]
    random.shuffle(pairs)

    # Iterate over input files
    for fname, rows in pairs:

        # Start a fresh Excel process
        xl = excel.Excel(start=True)

        try:
            
            # Get a new memory data collector
            collector = memdata.MemDataCollector()

            # Open, measure, close, repeat
            for t in range(trials):
                print(f"Opening {fname} (trial {t + 1})")
                xl.open_wb(os.path.join(pathlib.Path.cwd() / path, fname))
                collector.measure(xl.pid)
                xl.close_wb()
            
            # Update results
            if rows not in results: results[rows] = {}
            results[rows].update(collector.report(smooth=True, prefix=prefix, suffix=" (MB)", normalizer=1e6))

            # Show results
            print(results[rows])

        except Exception as e:
            
            # Completely end the program if an exception is raised
            traceback.print_exc()
            sys.exit()
        
        finally:

            # Close Excel
            xl.end()

def main(inputs_path
    , output_path
    , fv_inputdir="formula-value"
    , vo_inputdir="value-only"
    , totl_trials=1):

    # Ensures Excel is fully terminated before starting
    print("Closing all running instances of EXCEL.EXE (if any)")
    subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"], stderr=subprocess.DEVNULL)

    # Create directories
    if not os.path.exists(output_path): os.makedirs(output_path)

    # Run Experiments
    results = dict()                                                                # Container for experimental results        
    exptime = datetime.datetime.now()                                               # Start time
    run(os.path.join(inputs_path, vo_inputdir), totl_trials, "Value "  , results)   # Run value-only experiments
    run(os.path.join(inputs_path, fv_inputdir), totl_trials, "Formula ", results)   # Run formula-value experiments

    # Report timing stats
    exptime = (datetime.datetime.now() - exptime).total_seconds()
    print("\nTotal time (HH:MM:SS): {:02}:{:02}:{:02}".format(
        int(exptime // 3600), 
        int(exptime % 3600 // 60), 
        int(exptime % 60)
    ))

    # Write results to an Excel file
    results = pandas.DataFrame.from_dict(results, orient="index")
    results.index.rename("Rows", inplace=True)
    results.sort_index(inplace=True)
    results.to_excel(os.path.join(output_path, "memory.xlsx"))
    