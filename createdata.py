import pymem.utils.pipeline
import os

if __name__ == "__main__":

    # See datagen's README for explanations of each field
    CONFIG_ARGS = pymem.utils.pipeline.ConfigArgs(
        path=os.path.join("TEST")
        , inst="SingleCellSum"
        , seed=42
        , xlsx=False
        , step=100000
        , rows=100000
        , cols=1
        , itrs=10
        , pool=3
    )

    pymem.utils.pipeline.create_datasets(CONFIG_ARGS)
