from code.excelmem.excelmemvis import main as excel_main
from code.libremem.librememvis import main as libre_main

import plotly.graph_objects as go
import multiprocessing as mp
import subprocess
import os

def __compile_java():
    pass

def __execute_java():
    pass

if __name__ == "__main__":
    su

# def __create_config(inst, path, fldr, rand, xlsx, step, rows, cols, itrs, pool):
#     with open('config', 'w') as f:
#         f.write(
#             f"""
#             INST={inst}
#             PATH={path}
#             FLDR={fldr}
#             RAND={rand}
#             XLSX={xlsx}
#             STEP={step}
#             ROWS={rows}
#             COLS={cols}
#             ITER={itrs}
#             POOL={pool}
#             """
#         )

# def run(application, output_name, outdir_name, script_arg, inst, path, fldr, rand, xlsx, step, rows, cols, itrs, pool, soffice_path):

#     # Create datasets
#     __create_config(inst, path, fldr, rand, xlsx, step, rows, cols, itrs, pool)
#     __compile_java()
#     __execute_java()

#     # Run experiments
#     parent_conn, child_conn = mp.Pipe()

#     if application.lower() == "excel":
#         process = mp.Process(target=excel_main
#             , args=(child_conn, path,)
#             , kwargs={
#                 "fv_inputdir"   : "formula-value"
#                 , "vo_inputdir" : "value-only"
#                 , "output_path" : os.path.join(output_name, "excel")
#                 , "outdir_name" : outdir_name
#                 , "trials"      : script_arg
#             }
#         )

#     elif application.lower() == "libre":
#         process = mp.Process(target=libre_main
#             , args=(child_conn, path,)
#             , kwargs={
#                 "soffice_path"  : soffice_path
#                 , "fv_inputdir" : "formula-value"
#                 , "vo_inputdir" : "value-only"
#                 , "output_path" : os.path.join(output_name, "libre")
#                 , "outdir_name" : outdir_name
#                 , "pollseconds" : script_arg
#             }
#         )

#     else:
#         process = None
#         print(f"{application} is an invalid option.")

