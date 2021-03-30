import collections
import subprocess
import datetime
import pathlib
import random
import pandas
import psutil
import time
import json
import os

def get_soffice_bin_pid():
    """
    Finds the process ID of soffice.bin.

    Returns:
    --------
        The process ID of soffice.bin or None if it couldn't be found.
    """
    bin_pid = None
    for p in psutil.process_iter(['pid', 'name']):
        if 'soffice.bin' in p.info['name']:
            return p.info['pid']
    return bin_pid

def get_soffice_bin_mem():
    """
    Gets the memory usage of soffice.bin.

    Returns:
    --------
        A dictionary of memory data.
    """
    # Sometimes the PID of soffice.bin will be reassigned in the middle of data
    # collection and cause an exception. The while loop below should fix this.
    pid = [get_soffice_bin_pid()]
    while len(pid) > 0:
        try:
            memdata = psutil.Process(pid.pop()).memory_full_info()
        except Exception as e:
            pid.append(get_soffice_bin_pid())
        else:
            return {
                'peak_nonpaged_pool'    : memdata.peak_nonpaged_pool,
                'peak_paged_pool'       : memdata.peak_paged_pool,
                'peak_pagefile'         : memdata.peak_pagefile,
                'nonpaged_pool'         : memdata.nonpaged_pool,
                'paged_pool'            : memdata.paged_pool,
                'peak_wset'             : memdata.peak_wset,
                'pagefile'              : memdata.pagefile,
                'private'               : memdata.private,
                'wset'                  : memdata.wset,
                'rss'                   : memdata.rss,
                'uss'                   : memdata.uss,
                'vms'                   : memdata.vms
            }

def measure_calc_mem(soffice_path, in_path, out_path, poll_seconds, prefix="", suffix=" (Bytes)", normalizer=1):
    """
    Measures the memory consumption of the Calc file specified by `in_path`.
    All measurements are in bytes.

    Parameters:
    -----------
        soffice_path : str
            The absolute path to soffice.

        in_path : path-like
            A path to a directory of Calc files to open.

        out_path : path-like
            A path to an empty directory to place memory vs. time
            graph data.

        poll_seconds : int
            The minimum number of seconds to wait before polling
            the task manager for memory metrics.

        prefix : str
            This will be prepended to the keys of the dictionary 
            returned by `get_soffice_bin_mem()`.

        suffix : str
            This will be appended to the keys of the dictionary 
            returned by `get_soffice_bin_mem()`.
            
        normalizer : int
            The last memory measurements will be divided by this
            value.
    
    Returns:
    --------
        A dictionary containing the final memory measurements.
    """

    # Open the file in Calc without a GUI
    # See: https://help.libreoffice.org/3.3/Common/Starting_the_Software_With_Parameters
    proc = subprocess.Popen([soffice_path, '--headless', '-o', in_path]
        , stderr=subprocess.DEVNULL
        , stdin=subprocess.DEVNULL
    )

    # Get file name for printing purposes
    file_name = str(os.path.basename(os.path.normpath(in_path)))

    # Repeatedly poll the task manager for memory info
    # Stop polling once USS memory stops changing
    iterate = 5
    memdict = collections.OrderedDict()
    while True:
        prevmem = get_soffice_bin_mem()
        memdict[str(datetime.datetime.now().replace(microsecond=0))] = prevmem.copy()
        print(file_name + ": " + str(prevmem))
        time.sleep(poll_seconds)
        currmem = get_soffice_bin_mem()
        memdict[str(datetime.datetime.now().replace(microsecond=0))] = currmem.copy()
        print(file_name + ": " + str(currmem))
        time.sleep(poll_seconds)
        if currmem["uss"] - prevmem["uss"] == 0:
            for _ in range(iterate):
                currmem = get_soffice_bin_mem()
                memdict[str(datetime.datetime.now().replace(microsecond=0))] = currmem
                print(file_name + ": " + str(currmem))
                time.sleep(poll_seconds)
            break

    # Clean up
    proc.terminate()
    proc.wait()

    # Save memory vs. time data as a JSON
    with open(out_path, "w") as f: json.dump(memdict, f)

    # Return the last memory measurement
    return { prefix + k + suffix : v / normalizer for k, v in next(reversed(memdict.values())).items() }

def run(soffice_path, in_path, out_path, poll_seconds, prefix, results):
    """
    Measures the memory consumption of all Calc files in `path`.
    All measurements are in bytes.

    Parameters:
    -----------
        soffice_path : str
            The absolute path to soffice.

        in_path : path-like
            A path to a directory of Calc files to open.

        out_path : path-like
            A path to an empty directory to place memory vs. time
            graph data.
        
        poll_seconds : int
            The minimum number of seconds to wait before polling
            the task manager for memory metrics.

        prefix : str
            This will be prepended to the keys of the dictionary 
            returned by `get_soffice_bin_mem()`.
        
        results : dict
            A dictionary to store the final memory measurements 
            in.
    """
    pairs = [(f, int(f[f.index('-')+1:f.index('.')])) for f in os.listdir(in_path) if f.endswith(".ods")]
    random.shuffle(pairs)
    for fname, rows in pairs:
        time.sleep(0.5)
        if rows not in results: results[rows] = {}
        results[rows].update(
            measure_calc_mem(soffice_path
                , os.path.join(in_path, fname)
                , os.path.join(out_path, fname.replace('.ods', '.json'))
                , poll_seconds
                , prefix=prefix
                , suffix=" (MB)"
                , normalizer=1e6
            )
        )

if __name__ == "__main__":

    # The absolute path to soffice
    SOFFICE_PATH = "C:/Program Files/LibreOffice/program/soffice"

    # A directory with at least two folders of experimental Calc files
    INPUTS_PATH = os.path.join("..", "..", "exp")

    # The name of the folder with formula-value datasets (assumed to be inside INPUTS_PATH)
    FV_INPUTDIR = "formula-value"

    # The name of the folder with formula-value datasets (assumed to be inside INPUTS_PATH)
    VO_INPUTDIR = "value-only"

    # The path to a directory where all experimental folders will be created
    OUTPUT_PATH = os.path.join("..", "results")

    # The name of the directory to write results to (overwites any existing file(s) with the same name without warning)
    OUTDIR_NAME = "test"

    # The minimum number of seconds to wait before polling the task manager for memory metrics
    POLLSECONDS = 1

    # Ensures Calc is fully terminated before starting
    subprocess.call(["taskkill", "/f", "/im", "soffice.exe"], stderr=subprocess.DEVNULL)

    # Create fancy directory structure
    output_fldr = os.path.join(OUTPUT_PATH, OUTDIR_NAME)
    vo_memcurve = os.path.join(output_fldr, "vo-mem-curve")
    fv_memcurve = os.path.join(output_fldr, "fv-mem-curve")
    if not os.path.exists(output_fldr): os.makedirs(output_fldr)
    if not os.path.exists(vo_memcurve): os.makedirs(vo_memcurve)
    if not os.path.exists(fv_memcurve): os.makedirs(fv_memcurve)

    # Run experiments
    results = {}
    exptime = datetime.datetime.now()
    run(SOFFICE_PATH, os.path.join(INPUTS_PATH, VO_INPUTDIR), vo_memcurve, POLLSECONDS, "Value "  , results)
    run(SOFFICE_PATH, os.path.join(INPUTS_PATH, FV_INPUTDIR), fv_memcurve, POLLSECONDS, "Formula ", results)

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
    results.to_excel(os.path.join(output_fldr, "memory.xlsx"))
