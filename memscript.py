import pymem.utils.pipeline as pipeline
import pathlib
import os

# A script for running memory benchmarking experiments
if __name__ == "__main__":

    # See README for explanations of each field
    INTEGER_ARG = 5
    OUTPUT_PATH = os.path.join("experiments", "results", "excel", "rcbs-5trials-rand-1col", "run2")
    SOFFICE_DIR = "C:/Program Files/LibreOffice/program/soffice"
    CONFIG_ARGS = pipeline.ConfigArgs(
        path=os.path.join("experiments", "results", "excel", "rcbs-5trials-rand-1col", "dataset")
        , inst="CompleteBipartiteSum"
        , rand=42
        , xlsx=True
        , step=10000
        , rows=0
        , cols=1
        , itrs=11
        , pool=5
    )

    # Create datasets, run experiments, save results
    pipeline.run(CONFIG_ARGS, OUTPUT_PATH, INTEGER_ARG, SOFFICE_DIR)
