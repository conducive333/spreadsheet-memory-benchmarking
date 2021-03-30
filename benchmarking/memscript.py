import pathlib
import util
import os

# A script for running memory benchmarking experiments
if __name__ == "__main__":

    # See README for an explanation of each field
    INTEGER_ARG = 1
    OUTPUT_PATH = os.path.join("experiments", "results", "excel", "TEST")
    SOFFICE_DIR = "C:/Program Files/LibreOffice/program/soffice"
    CONFIG_ARGS = util.ConfigArgs(inst="CompleteBipartiteSum"
        , path=os.path.join("experiments", "input-data", "rscs-test")
        , rand=42
        , xlsx=False
        , step=10000
        , rows=0
        , cols=1
        , itrs=11
        , pool=5
    )

    # Create datasets, run experiments, save results
    util.run(CONFIG_ARGS, OUTPUT_PATH, INTEGER_ARG, SOFFICE_DIR)
