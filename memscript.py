import pymem.utils.pipeline
import os

# A script for running memory benchmarking experiments
if __name__ == "__main__":

    # See README for explanations of each field
    INTEGER_ARG = 1
    OUTPUT_PATH = os.path.join("TEST")
    SOFFICE_DIR = "C:/Program Files/LibreOffice/program/soffice"
    CONFIG_ARGS = pymem.utils.pipeline.ConfigArgs(
        path=os.path.join("TEST", "dataset")
        , inst="CompleteBipartiteSum"
        , seed=42
        , xlsx=True
        , step=0
        , rows=100
        , cols=1
        , itrs=1
        , pool=1
        , uppr=10
    )

    # Create datasets, run experiments, save results
    pymem.utils.pipeline.run(CONFIG_ARGS, OUTPUT_PATH, INTEGER_ARG, SOFFICE_DIR)
