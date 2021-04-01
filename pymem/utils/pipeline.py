import pymem.excelmem.memvis as excel_memvis
import pymem.libremem.memvis as libre_memvis
import pymem.excelmem.mem as excel_mem
import pymem.libremem.mem as libre_mem
import multiprocessing
import subprocess
import pathlib
import os

from . import definitions

class ConfigArgs:
    def __init__(self, path, inst, rand, xlsx, step, rows, cols, itrs, pool):
        self.path = path
        self.inst = inst
        self.rand = rand
        self.xlsx = xlsx
        self.step = step
        self.rows = rows
        self.cols = cols
        self.itrs = itrs
        self.pool = pool

def __create_process(child_conn, xlsx, inputs_path, output_path, experiment_arg, sofficepath):
    """
    Performs each phase of the benchmarking pipeline:
        1. Generates datasets (if necessary)
        2. Runs the appropriate benchmarking code
        3. Saves results to the specified output path

    Parameter(s):
    -------------
        child_conn : multiprocessing.Connection
            A multiprocessing.Connection object for communicating
            with the parent process.

        xlsx : bool
            See the XLSX parameter in datagen's README.

        inputs_path : path-like
            See PATH parameter in datagen's README.

        output_path : path-like
            See README.

        experiment_arg : int
            See README.

        sofficepath : str or path-like
            See README.
    """
    if xlsx:
        return multiprocessing.Process(target=excel_memvis.main
            , args=(child_conn, inputs_path, output_path, )
            , kwargs={
                "fv_inputdir"   : "formula-value"
                , "vo_inputdir" : "value-only"
                , "totl_trials" : experiment_arg
            }
        )
    else:
        return multiprocessing.Process(target=libre_memvis.main
            , args=(child_conn, inputs_path, output_path, )
            , kwargs={
                "fv_inputdir"   : "formula-value"
                , "vo_inputdir" : "value-only"
                , "pollseconds" : experiment_arg
                , "sofficepath" : sofficepath
            }
        )

def create_datasets(config_args):
    """
    Creates a file named 'config' in this project's root directory
    using the members of `config_args`.

    Parameter(s):
    -------------
        config_args : ConfigArgs
            Specifies how to create the benchmarking dataset. See the
            datagen project's README for an explanation of each field.
    """
    cwd = pathlib.Path().cwd()
    os.chdir(definitions.ROOT_DIR)
    with open(os.path.join(definitions.ROOT_DIR, 'config'), 'w') as f:
        config_args_path = pathlib.Path(config_args.path).as_posix()
        f.write(f"INST={config_args.inst}\n")
        f.write(f"PATH={config_args_path}\n")
        f.write(f"RAND={config_args.rand}\n")
        f.write(f"XLSX={config_args.xlsx}\n")
        f.write(f"STEP={config_args.step}\n")
        f.write(f"ROWS={config_args.rows}\n")
        f.write(f"COLS={config_args.cols}\n")
        f.write(f"ITER={config_args.itrs}\n")
        f.write(f"POOL={config_args.pool}\n")
    subprocess.check_call(["java", "-jar", "main.jar"])
    os.remove(os.path.join(definitions.ROOT_DIR, 'config'))
    os.chdir(cwd)

def run(config_args, output_path, experiment_arg, sofficepath):
    """
    Performs each phase of the benchmarking pipeline:
        1. Generates datasets (if necessary)
        2. Runs the appropriate benchmarking code
        3. Saves results to the specified output path

    Parameter(s):
    -------------
        config_args : ConfigArgs
            See README.

        output_path : path-like
            See README.

        experiment_arg : int
            See README.

        sofficepath : str or path-like
            See README.
    """

    # Create relative paths
    relative_inputs_path = os.path.join(definitions.ROOT_DIR, config_args.path)
    relative_output_path = os.path.join(definitions.ROOT_DIR, output_path)

    # Create datasets (if necessary)
    if not os.path.exists(relative_inputs_path): create_datasets(config_args)

    # Run experiments
    if config_args.xlsx:
        excel_mem.main(relative_inputs_path, relative_output_path, totl_trials=experiment_arg)
    else:
        libre_mem.main(relative_inputs_path, relative_output_path, pollseconds=experiment_arg, sofficepath=sofficepath)

def run_vis(config_args, output_path, experiment_arg, sofficepath, fv_fig, vo_fig, barplt, column):
    """
    Updates figures with intermediate results and performs each phase 
    of the benchmarking pipeline:
        1. Generates datasets (if necessary)
        2. Runs the appropriate benchmarking code
        3. Saves results to the specified output path

    Parameter(s):
    -------------
        config_args : ConfigArgs
            See README.

        output_path : path-like
            See README.

        experiment_arg : int
            See README.

        sofficepath : str or path-like
            See README.

        fv_fig : plotly.graphobjects.FigureWidget
            A scatter plot of memory vs. rows. This will be updated with 
            formula-value benchmarking results as they become available. 
            You can choose which variable to plot on the y-axis using the 
            `column` parameter.

        vo_fig : plotly.graphobjects.FigureWidget
            A scatter plot of memory vs. rows. This will be updated with 
            value-only benchmarking results as they become available. You
            can choose which variable to plot on the y-axis using the 
            `column` parameter.

        barplt : plotly.graphobjects.FigureWidget
            A side-by-side bar plot of memory vs. rows. The bar plot at index
            0 is assumed to be for value-only experiments. The bar plot at index
            1 is assumed to be for formula-value experiments.

        column : str
            The memory variable to plot. Must be one of: 'peak_nonpaged_pool', 
            'peak_paged_pool', 'peak_pagefile', 'nonpaged_pool', 'paged_pool', 
            'peak_wset', 'pagefile', 'private', 'wset', 'rss', 'uss', 'vms'.
    """

    # Create relative paths
    relative_inputs_path = os.path.join(definitions.ROOT_DIR, config_args.path)
    relative_output_path = os.path.join(definitions.ROOT_DIR, output_path)

    # Create a process with experiment arguments
    parent_conn, child_conn = multiprocessing.Pipe()
    process = __create_process(child_conn
        , config_args.xlsx
        , relative_inputs_path
        , relative_output_path
        , experiment_arg
        , sofficepath
    )

    # Create datasets (if necessary)
    if not os.path.exists(relative_inputs_path): create_datasets(config_args)

    # Run experiments and update plots in real-time
    if process is not None:
        process.start()
        while True:
            item = parent_conn.recv()
            if type(item) == dict:
                vo_rowsizes = []; vo_memsizes = []; fv_rowsizes = []; fv_memsizes = []
                for r, d in sorted(item.items(), key=lambda pair: pair[0]):
                    if f"Value {column} (MB)"   in d: vo_rowsizes.append(str(r)); vo_memsizes.append(d[f"Value {column} (MB)"  ])
                    if f"Formula {column} (MB)" in d: fv_rowsizes.append(str(r)); fv_memsizes.append(d[f"Formula {column} (MB)"])
                barplt.data[0].x = vo_fig.data[0].x = vo_rowsizes
                barplt.data[1].x = fv_fig.data[0].x = fv_rowsizes
                barplt.data[0].y = vo_fig.data[0].y = vo_memsizes
                barplt.data[1].y = fv_fig.data[0].y = fv_memsizes
            if type(item) == str:   print(item)
            if item is None:        break
        process.join()
        process.close()
