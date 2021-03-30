import win32com.client
import subprocess
import traceback
import datetime
import memdata
import pathlib
import random
import pandas
import time
import sys
import os

# Benchmark all .xlsx files in path
def run(path, prefix, trials, results):
    
    # Collect and shuffle input files
    pairs = [(f, int(f[f.index('-')+1:f.index('.')])) for f in os.listdir(path) if f.endswith(".xlsx")]
    random.shuffle(pairs)

    # Iterate over input files
    for fname, rows in pairs:

        # Start a fresh Excel process
        excel = win32com.client.DispatchEx('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        try:
            
            # Get a new memory data collector
            collector = memdata.ExcelMemDataCollector(excel)

            # Open, measure, close, repeat
            for t in range(trials):
                print(f"Opening {fname} (trial {t + 1})")
                wb = excel.Workbooks.Open(os.path.join(pathlib.Path.cwd() / path, fname))
                collector.measure()
                wb.Close()

            # Free resources
            collector.close_handle()
            
            # Update results
            if rows not in results: results[rows] = {}
            results[rows].update(collector.report(smooth=True, prefix=prefix, suffix=" (MB)", normalizer=1e6))

            # Show results and clean up
            print(results[rows])

        except Exception as e:
            
            # Completely end the program if an exception is raised
            traceback.print_exc()
            sys.exit()

        finally:

            # Close Excel
            excel.Application.Quit()

if __name__ == "__main__":

    # Stores the absolute path to the project's root directory
    ROOT_DIR = pathlib.Path(os.path.dirname(os.path.realpath(__file__))).parent.absolute()

    # A directory with at least two folders of experimental Excel files
    INPUTS_PATH = os.path.join(ROOT_DIR, "input-data", "rscs-test")

    # The path to the output directory
    OUTPUT_PATH = os.path.join(ROOT_DIR, "results")

    # The name of the folder with formula-value datasets (assumed to be inside INPUTS_PATH)
    FV_INPUTDIR = "formula-value"

    # The name of the folder with formula-value datasets (assumed to be inside INPUTS_PATH)
    VO_INPUTDIR = "value-only"
    
    # The number of trials to perform for each run
    TOTL_TRIALS = 1

    # Ensures Excel is fully terminated before starting
    print("Closing all running instances of EXCEL.EXE (if any)")
    subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"], stderr=subprocess.DEVNULL)

    # Create directories
    if not os.path.exists(OUTPUT_PATH): os.makedirs(OUTPUT_PATH)

    # Run Experiments
    results = dict()                                                                # Container for experimental results        
    exptime = datetime.datetime.now()                                               # Start time
    run(os.path.join(INPUTS_PATH, VO_INPUTDIR), "Value "  , TOTL_TRIALS, results)   # Run value-only experiments
    run(os.path.join(INPUTS_PATH, FV_INPUTDIR), "Formula ", TOTL_TRIALS, results)   # Run formula-value experiments

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
    results.to_excel(os.path.join(OUTPUT_PATH, "memory.xlsx"))
    