import pymem.utils.pipeline
import os

if __name__ == "__main__":

    # See datagen's README for explanations of each field
    CONFIG_ARGS = pymem.utils.pipeline.ConfigArgs(
        path=os.path.join("TEST")
        , inst="SingleCellSum"
        , rand=42
        , xlsx=False
        , step=1
        , rows=0
        , cols=1
        , itrs=1
        , pool=1
    )

    pymem.utils.pipeline.create_datasets(CONFIG_ARGS)
