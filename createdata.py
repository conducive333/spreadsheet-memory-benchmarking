import pymem.utils.pipeline
import os

if __name__ == "__main__":

    # See datagen's README for explanations of each field
    CONFIG_ARGS = pymem.utils.pipeline.ConfigArgs(
        path=os.path.join("..", "experiments", "results", "libre", "experiment6", "r(same)cv-1sec-1col-5iter-main", "dataset")
        , inst="SameCellVlookup"
        , seed=42
        , xlsx=False
        , step=10000
        , rows=0
        , cols=1
        , itrs=10
        , pool=5
    )

    pymem.utils.pipeline.create_datasets(CONFIG_ARGS)
