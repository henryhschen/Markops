#!/usr/local/bin/python3.6

import re, json, collections, argparse, matplotlib
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

parser = argparse.ArgumentParser(description='Process for Markops')
parser.add_argument('-excel',nargs='+', help='Input excel files', required=True)
#parser.add_argument('config',nargs='?', help='Input config file', default="naming")
#parser.add_argument('config',nargs='?', help='Input config file', default="naming")

try:
    args = parser.parse_args()
    print("-excel", args.excel)
except:
    parser.print_help()
    exit()

data = collections.defaultdict(lambda: collections.defaultdict(dict))
for f in sorted(args.excel):
    xls = pd.ExcelFile(f)
    f_name = re.search(r"^Period(\d)\.", f, re.IGNORECASE).groups()[0]
    for sheet in sorted(xls.sheet_names):
        parse = xls.parse(sheet)
        sheet_name = re.search(r"^(\w)_", sheet, re.IGNORECASE).groups()[0]
        data[f_name][sheet_name.upper()]= parse

#for f, v1 in data.items():
#    for s, v2 in v1.items():
#        print(f, s)

for f, v1 in data.items():
    v1["G"]["Type"] = v1["G"]["Product"]+v1["G"]["Company"]+"_"+v1["G"]["SM"]
    plt.plot(v1["G"]["Type"], v1["G"]["USWOT"])
#with open("data", "rt") as f:
#    config = json.load(f)
#f.close()
#print(config)




