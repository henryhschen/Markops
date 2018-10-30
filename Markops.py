#!/usr/local/bin/python3.6

import re, json, collections, argparse
import pandas as pd
import numpy as np
import matplotlib


parser = argparse.ArgumentParser(description='Process for Markops')
parser.add_argument('config',nargs='?', help='Input config file', default="naming")
args = parser.parse_args()
print("-config", args.config)

test = collections.defaultdict(lambda: collections.defaultdict(dict))

test["PV"]["name"] = "Production level"
test["PV"]["unit"] = "KU"
test["PC"]["name"] = "Production capacity"
test["PC"]["unit"] = "KU"

#with open("data", "w") as f:
#    json.dump(test, f, indent=4)

with open("data", "rt") as f:
    config = json.load(f)
f.close()
print(config)




