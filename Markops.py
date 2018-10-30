#!/usr/local/bin/python3.6

import re, json, collections, argparse, matplotlib
import pandas as pd
import numpy as np

parser = argparse.ArgumentParser(description='Process for Markops')
parser.add_argument('config',nargs='?', help='Input config file', default="naming")
args = parser.parse_args()
print("-config", args.config)

test = collections.defaultdict(lambda: collections.defaultdict(dict))

with open("data", "rt") as f:
    config = json.load(f)
f.close()
print(config)




