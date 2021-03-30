import win32com.client
import subprocess
import traceback
import datetime
import pathlib
import random
import pandas
import time
import sys
import os

from . import memdata

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
    run(os.path.join(inputs_path, vo_inputdir), "Value "  , totl_trials, results)   # Run value-only experiments
    run(os.path.join(inputs_path, fv_inputdir), "Formula ", totl_trials, results)   # Run formula-value experiments

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
    