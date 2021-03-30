from excelmem.excelmemvis import main as excel_main
from libremem.librememvis import main as libre_main

import plotly.graph_objects as go
import multiprocessing as mp
import subprocess
import os

def __create_datasets(inst, path, fldr, rand, xlsx, step, rows, cols, itrs, pool):
    with open('config', 'w') as f:
        f.write(
            f"""
            INST={inst}
            PATH={path}
            FLDR={fldr}
            RAND={rand}
            XLSX={xlsx}
            STEP={step}
            ROWS={rows}
            COLS={cols}
            ITER={itrs}
            POOL={pool}
            """
        )
    subprocess.check_call(["java", "-jar", "main.jar"])

def run(application, output_home_dir, experiment_name, experiment_arg, inst, path, fldr, rand, xlsx, step, rows, cols, itrs, pool, soffice_path):

    __create_datasets(inst, path, fldr, rand, xlsx, step, rows, cols, itrs, pool)

    parent_conn, child_conn = mp.Pipe()

    if application.lower() == "excel":
        process = mp.Process(target=excel_main
            , args=(child_conn, path,)
            , kwargs={
                "fv_inputdir"   : "formula-value"
                , "vo_inputdir" : "value-only"
                , "output_path" : os.path.join(output_home_dir, "excel")
                , "outdir_name" : outdir_name
                , "trials"      : experiment_arg
            }
        )

    elif application.lower() == "libre":
        process = mp.Process(target=libre_main
            , args=(child_conn, path,)
            , kwargs={
                "soffice_path"  : soffice_path
                , "fv_inputdir" : "formula-value"
                , "vo_inputdir" : "value-only"
                , "output_path" : os.path.join(output_name, "libre")
                , "outdir_name" : outdir_name
                , "pollseconds" : experiment_arg
            }
        )

    else:
        process = None
        print(f"{application} is an invalid option.")

