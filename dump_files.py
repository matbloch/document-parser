import argparse
import os

from layout import GNPLayout

# ================================= #
#              Main

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('direc', help='an integer for the accumulator')

    # parse arguments
    args = parser.parse_args()

    for subdir, dirs, files in os.walk(rootdir):
        for file in files:
            pass