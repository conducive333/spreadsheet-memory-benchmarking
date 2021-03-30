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
def run(child_conn, path, trials, prefix, results):

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
                child_conn.send(f"Opening {fname} (trial {t + 1})")
                wb = excel.Workbooks.Open(os.path.join(pathlib.Path.cwd() / path, fname))
                collector.measure()
                wb.Close()

            # Free resources
            collector.close_handle()
            
            # Update results
            if rows not in results: results[rows] = {}
            results[rows].update(collector.report(smooth=True, prefix=prefix, suffix=" (MB)", normalizer=1e6))

            # Show results
            child_conn.send(results)
            child_conn.send(str(results[rows]))
        
        except Exception as e:
            
            # Exit the function if an exception is raised
            child_conn.send(traceback.format_exc())
            return

        finally:

            # Close Excel
            excel.Application.Quit()

def main(child_conn
    , inputs_path
    , output_path
    , fv_inputdir="formula-value"
    , vo_inputdir="value-only"
    , totl_trials=5):

    if child_conn is not None:
        
        try:

            # Ensures Excel is fully terminated before starting
            child_conn.send("Closing all running instances of EXCEL.EXE (if any)")
            subprocess.call(["taskkill", "/f", "/im", "EXCEL.EXE"], stderr=subprocess.DEVNULL)

            # Create directories
            if not os.path.exists(output_path): os.makedirs(output_path)

            # Run experiments
            results = dict()                                                                            # Container for experimental results        
            exptime = datetime.datetime.now()                                                           # Start time
            run(child_conn, os.path.join(inputs_path, vo_inputdir), totl_trials, "Value "  , results)   # Run value-only experiments
            run(child_conn, os.path.join(inputs_path, fv_inputdir), totl_trials, "Formula ", results)   # Run formula-value experiments
            
            # Report timing stats
            exptime = (datetime.datetime.now() - exptime).total_seconds()
            child_conn.send("\nTotal time (HH:MM:SS): {:02}:{:02}:{:02}".format(
                int(exptime // 3600), 
                int(exptime % 3600 // 60), 
                int(exptime % 60)
            ))

            # Push results to Excel
            results = pandas.DataFrame.from_dict(results, orient="index")
            results.index.rename("Rows", inplace=True)
            results.sort_index(inplace=True)
            results.to_excel(os.path.join(output_path, "memory.xlsx"))

        except Exception as e:

            # Send any errors as strings to the parent process
            child_conn.send(str(e))

        finally:
        
            # Pipe clean up
            child_conn.send(None)
            child_conn.close()
