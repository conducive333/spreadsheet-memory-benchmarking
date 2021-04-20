import pymem.utils.pipeline
import os

if __name__ == "__main__":

    # See datagen's README for explanations of each field
    CONFIG_ARGS = pymem.utils.pipeline.ConfigArgs(
        path=os.path.join("TEST")
        , inst="OverlappingSum"
        , seed=42
        , xlsx=True
        , step=10000
        , rows=1000
        , cols=1
        , itrs=1
        , pool=1
        , uppr=""
    )

    pymem.utils.pipeline.create_datasets(CONFIG_ARGS)
