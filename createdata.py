import pymem.utils.pipeline
import os

if __name__ == "__main__":

    # CBS, MRS, OVS (window size set to 500), CONST_SUM

    # See datagen's README for explanations of each field
    CONFIG_ARGS = pymem.utils.pipeline.ConfigArgs(
        path=os.path.join("..", "experiments", "results", "experiment-09", "64bit-excel", "rcbs-5trials-rand-1col-main", "dataset")
        , inst="OverlappingSum"
        , seed=42
        , xlsx=True
        , step=10000
        , rows=0
        , cols=1
        , itrs=10
        , pool=1
        , uppr=100
    )

    pymem.utils.pipeline.create_datasets(CONFIG_ARGS)
