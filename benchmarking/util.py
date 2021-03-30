from excelmem.memvis import main as excel_main_vis
from libremem.memvis import main as libre_main_vis
from excelmem.mem import main as excel_main
from libremem.mem import main as libre_main

import multiprocessing as mp
import subprocess
import pathlib
import shutil
import os

ROOT_DIR = pathlib.Path(os.path.dirname(os.path.realpath(__file__))).parent.absolute()

# TODO: Finish comments

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

def __create_datasets(config_args):
    """
    Creates a file named 'config' in this project's root directory.
    using the member variables of util.ConfigArgs.

    Parameter(s):
    -------------
        config_args : util.ConfigArgs
            A 
    """
    cwd = pathlib.Path().cwd()
    os.chdir(ROOT_DIR)
    with open(os.path.join(ROOT_DIR, 'config'), 'w') as f:
        f.write(f"INST={config_args.inst}\n")
        f.write(f"PATH={config_args.path}\n")
        f.write(f"RAND={config_args.rand}\n")
        f.write(f"XLSX={config_args.xlsx}\n")
        f.write(f"STEP={config_args.step}\n")
        f.write(f"ROWS={config_args.rows}\n")
        f.write(f"COLS={config_args.cols}\n")
        f.write(f"ITER={config_args.itrs}\n")
        f.write(f"POOL={config_args.pool}\n")
    subprocess.check_call(["java", "-jar", "main.jar"])
    os.chdir(cwd)

def __create_process(child_conn, xlsx, inputs_path, output_path, experiment_arg, sofficepath):
    if xlsx:
        return mp.Process(target=excel_main_vis
            , args=(child_conn, inputs_path, output_path, )
            , kwargs={
                "fv_inputdir"   : "formula-value"
                , "vo_inputdir" : "value-only"
                , "totl_trials" : experiment_arg
            }
        )
    else:
        return mp.Process(target=libre_main_vis
            , args=(child_conn, inputs_path, output_path, )
            , kwargs={
                "fv_inputdir"   : "formula-value"
                , "vo_inputdir" : "value-only"
                , "pollseconds" : experiment_arg
                , "sofficepath" : sofficepath
            }
        )

def run_vis(config_args, output_path, experiment_arg, sofficepath, fv_fig, vo_fig, barplt, column):
    """
    Parameters
    ----------
        column a string 
    """

    # Create relative paths
    relative_inputs_path = os.path.join(ROOT_DIR, config_args.path)
    relative_output_path = os.path.join(ROOT_DIR, output_path)

    # Create a process with experiment arguments
    parent_conn, child_conn = mp.Pipe()
    process = __create_process(child_conn
        , config_args.xlsx
        , relative_inputs_path
        , relative_output_path
        , experiment_arg
        , sofficepath
    )

    # Create datasets
    already_exists = os.path.exists(relative_inputs_path)
    if not already_exists: 
        __create_datasets(config_args)

    # Run experiments
    if process is not None:

        # Update plots in real-time
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

        # Move or copy dataset to output folder
        if already_exists:
            shutil.copytree(relative_inputs_path, os.path.join(relative_output_path, "dataset"))    
        else:
            shutil.move(relative_inputs_path, relative_output_path)

def run(config_args, output_path, experiment_arg, sofficepath):
    """
    """

    # Create relative paths
    relative_inputs_path = os.path.join(ROOT_DIR, config_args.path)
    relative_output_path = os.path.join(ROOT_DIR, output_path)

    # Create datasets
    already_exists = os.path.exists(relative_inputs_path)
    if not already_exists: 
        __create_datasets(config_args)

    # Run experiments
    if config_args.xlsx:
        excel_main(relative_inputs_path, relative_output_path, totl_trials=experiment_arg)
    else:
        libre_main(relative_inputs_path, relative_output_path, pollseconds=experiment_arg, sofficepath=sofficepath)

    # Move or copy dataset to output folder
    if already_exists:
        shutil.copytree(relative_inputs_path, os.path.join(relative_output_path, "dataset"))
    else:
        shutil.move(relative_inputs_path, relative_output_path)