#!/usr/local/bin/python3.6

import math, re, json, collections, argparse, matplotlib
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.text import TextPath

x_lower = 1
x_upper = 5

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

def iden_label (num):
    label = 0
    if(num < 30):
        label = "D"
    elif(num >= 30 and num < 50):
        label = "C"
    elif(num >= 50 and num < 70):
        label = "B"
    else:
        label = "A"
    return label

def iden_company(char):
    if(re.search(r"^A$", char, re.IGNORECASE)):
        linestyle = ":"
    elif(re.search(r"^K$", char, re.IGNORECASE)):
        linestyle = "-."
    elif(re.search(r"^P$", char, re.IGNORECASE)):
        linestyle = "--"
    else: 
        linestyle = "-"
    return linestyle

data = collections.defaultdict(lambda: collections.defaultdict(dict))
for f in sorted(args.excel):
    xls = pd.ExcelFile(f)
    f_name = re.search(r"^Period(\d)\.", f, re.IGNORECASE).groups()[0]
    for sheet in sorted(xls.sheet_names):
        parse = xls.parse(sheet)
        sheet_name = re.search(r"^(\w)_", sheet, re.IGNORECASE).groups()[0]
        data[f_name][sheet_name.upper()]= parse

# Deal with E for competitive
E_data = collections.defaultdict(lambda: collections.defaultdict(dict))
for f, v1 in sorted(data.items()):
    for r_i, row in v1["E"].iterrows():
        if(row["UDS"] != 0):
            umds = np.single(row["MDS"]/row["UDS"])
        else:
            umds = 0
        if(row["UIS"] != 0):
            umis = np.single(row["MIS"]/row["UIS"])
        else:
            umis = 0
        if(row["UTS"] != 0):
            umts = np.single(row["MTS"]/row["UTS"])
        else:
            umts = 0
        if("x" in E_data["Direct"][row[0]+row[1]]):
            E_data["Direct"][row[0]+row[1]]["x"].append(f)
            E_data["Direct"][row[0]+row[1]]["y"].append(umds)
            E_data["Indirect"][row[0]+row[1]]["x"].append(f)
            E_data["Indirect"][row[0]+row[1]]["y"].append(umis)
            E_data["Total"][row[0]+row[1]]["x"].append(f)
            E_data["Total"][row[0]+row[1]]["y"].append(umts)
            E_data["Surplus"][row[0]+row[1]]["x"].append(f)
            E_data["Surplus"][row[0]+row[1]]["y"].append(row["PV"]-row["UTS"])
            E_data["PC"][row[0]+row[1]]["x"].append(f)
            E_data["PC"][row[0]+row[1]]["y"].append(row["PC"])
            E_data["PV"][row[0]+row[1]]["x"].append(f)
            E_data["PV"][row[0]+row[1]]["y"].append(row["PV"])
        else:
            E_data["Direct"][row[0]+row[1]]["unit"] = "Volume"
            E_data["Direct"][row[0]+row[1]]["x"] = [f]
            E_data["Direct"][row[0]+row[1]]["y"] = [umds]
            E_data["Indirect"][row[0]+row[1]]["unit"] = "Volume"
            E_data["Indirect"][row[0]+row[1]]["x"] = [f]
            E_data["Indirect"][row[0]+row[1]]["y"] = [umis]
            E_data["Total"][row[0]+row[1]]["unit"] = "Volume"
            E_data["Total"][row[0]+row[1]]["x"] = [f]
            E_data["Total"][row[0]+row[1]]["y"] = [umts]
            E_data["Surplus"][row[0]+row[1]]["unit"] = "Volume"
            E_data["Surplus"][row[0]+row[1]]["x"] = [f]
            E_data["Surplus"][row[0]+row[1]]["y"] = [row["PC"]-row["UTS"]]
            E_data["PC"][row[0]+row[1]]["unit"] = "Volume"
            E_data["PC"][row[0]+row[1]]["x"] = [f]
            E_data["PC"][row[0]+row[1]]["y"] = [row["PC"]]
            E_data["PV"][row[0]+row[1]]["unit"] = "Volume"
            E_data["PV"][row[0]+row[1]]["x"] = [f]
            E_data["PV"][row[0]+row[1]]["y"] = [row["PV"]]

fig = plt.figure(num=None, figsize=(8, 6), facecolor='w', edgecolor='k')
fig.subplots_adjust(hspace=0.3, wspace=0.3)
i=0
for ttype, v1 in E_data.items():
    i += 1
    ax = fig.add_subplot(2, math.ceil(len(E_data)/2), i)
    for product, v2 in v1.items():
        search = re.search(r"^(\w)Z(\d+)$", product, re.IGNORECASE).groups()
        char = search[0]
        label= iden_label(int(search[1]))
        label = TextPath((0,0), str(char+label), linewidth=3)
        linestyle = iden_company(char)
        ax.plot(v2["x"], v2["y"], label=product, linestyle=linestyle, marker=label, markersize=20)
        ax.set_title(ttype)
    plt.xlim([x_lower,x_upper])
    ax.set_xlabel('Period')
    ax.set_ylabel(v2["unit"])
handles, labels = ax.get_legend_handles_labels()
labels, handles = zip(*sorted(zip(labels, handles), key=lambda t: t[0]))
fig.legend(handles, labels, loc='lower right')
fig.suptitle('E. Product Sales', fontsize=20)
fig.savefig('E.png')


# Deal with H for direct competitive
H_data = collections.defaultdict(lambda: collections.defaultdict(dict))
H_I_unit_data = collections.defaultdict(lambda: collections.defaultdict(dict))
for f, v1 in sorted(data.items()):
    #v1["H"]["Type"] = v1["H"]["Company"]+v1["H"]["Product"]
    for r_i, row in v1["H"].iterrows():
        for col in v1["H"].columns[2:]:
            #print(col, row[col], row[0]+row[1])
            if("x" in H_data[col][row[0]+row[1]]):
                H_data[col][row[0]+row[1]]["x"].append(f)
                H_data[col][row[0]+row[1]]["y"].append(row[col])
            else:
                if(re.search(r"^CT$", col, re.IGNORECASE)):
                    H_data[col][row[0]+row[1]]["unit"] = "Days"
                else:
                    H_data[col][row[0]+row[1]]["unit"] = "Sales"
                H_data[col][row[0]+row[1]]["x"] = [f]
                H_data[col][row[0]+row[1]]["y"] = [row[col]]
  
        
for f, v1 in sorted(data.items()):
    for r_i, row in v1["H"].iterrows():
        for col in v1["H"].columns[2:3]:
            #print(col, row[col], row[0]+row[1])
            swot = int(row[1].replace("Z", ""))
            unit_price =np.single(row[col]/swot)
            label= iden_label(int(search[1]))
            H_I_unit_data[f][row[0]+row[1]+label+"_Direct"]["UnitPrice"] = unit_price
            H_I_unit_data[f][row[0]+row[1]+label+"_Direct"]["swot"] = row[col]
            H_I_unit_data[f][row[0]+row[1]+label+"_Direct"]["seg"] = label

fig = plt.figure(num=None, figsize=(8, 6), facecolor='w', edgecolor='k')
fig.subplots_adjust(hspace=0.3, wspace=0.3)
i=0
for ttype, v1 in H_data.items():
    i += 1
    ax = fig.add_subplot(2, math.ceil(len(H_data)/2), i)
    for product, v2 in v1.items():
        search = re.search(r"^(\w)Z(\d+)$", product, re.IGNORECASE).groups()
        char = search[0]
        label= iden_label(int(search[1]))
        label = TextPath((0,0), str(char+label), linewidth=3)
        linestyle = iden_company(char)
        ax.plot(v2["x"], v2["y"], label=product, linestyle=linestyle, marker=label, markersize=20)
        ax.set_title(ttype)
    ax.set_xlabel('Period')
    ax.set_ylabel(v2["unit"])
    plt.xlim([x_lower,x_upper])
handles, labels = ax.get_legend_handles_labels()
labels, handles = zip(*sorted(zip(labels, handles), key=lambda t: t[0]))
fig.legend(handles, labels, loc='lower right')
fig.suptitle('H. Direct with Competitive', fontsize=20)
fig.savefig('H.png')


# Deal with I for indirect competitive
I_data = collections.defaultdict(lambda: collections.defaultdict(dict))
for f, v1 in sorted(data.items()):
    for r_i, row in v1["I"].iterrows():
        for col in v1["I"].columns[2:]:
            #print(col, row[col], row[0]+row[1])
            if("x" in I_data[col][row[0]+row[1]]):
                I_data[col][row[0]+row[1]]["x"].append(f)
                I_data[col][row[0]+row[1]]["y"].append(row[col])
            else:
                if(re.search(r"^CT$", col, re.IGNORECASE)):
                    I_data[col][row[0]+row[1]]["unit"] = "Days"
                else:
                    I_data[col][row[0]+row[1]]["unit"] = "Sales"
                I_data[col][row[0]+row[1]]["x"] = [f]
                I_data[col][row[0]+row[1]]["y"] = [row[col]]

for f, v1 in sorted(data.items()):
    for r_i, row in v1["I"].iterrows():
        for col in v1["I"].columns[2:3]:
            #print(col, row[col], row[0]+row[1])
            swot = int(row[1].replace("Z", ""))
            unit_price =np.single(row[col]/swot)
            label= iden_label(int(search[1]))
            H_I_unit_data[f][row[0]+row[1]+label+"_Indirect"]["UnitPrice"] = unit_price
            H_I_unit_data[f][row[0]+row[1]+label+"_Indirect"]["swot"] = row[col]
            H_I_unit_data[f][row[0]+row[1]+label+"_Indirect"]["seg"] = label



fig = plt.figure(num=None, figsize=(8, 6), facecolor='w', edgecolor='k')
fig.subplots_adjust(hspace=0.3, wspace=0.3)
i=0
for ttype, v1 in I_data.items():
    i += 1
    ax = fig.add_subplot(2, math.ceil(len(I_data)/2), i)
    for product, v2 in v1.items():
        search = re.search(r"^(\w)Z(\d+)$", product, re.IGNORECASE).groups()
        char = search[0]
        label= iden_label(int(search[1]))
        label = TextPath((0,0), str(char+label), linewidth=3)
        linestyle = iden_company(char)
        ax.plot(v2["x"], v2["y"], label=product, linestyle=linestyle, marker=label, markersize=20)
        ax.set_title(ttype)
    ax.set_xlabel('Period')
    ax.set_ylabel(v2["unit"])
    plt.xlim([x_lower,x_upper])
handles, labels = ax.get_legend_handles_labels()
labels, handles = zip(*sorted(zip(labels, handles), key=lambda t: t[0]))
fig.legend(handles, labels, loc='lower right')
fig.suptitle('I. Inirect with Competitive', fontsize=20)
fig.savefig('I.png')

# Deal with H and I
fig = plt.figure(num=None, figsize=(8, 6), facecolor='w', edgecolor='k')
fig.subplots_adjust(hspace=0.3, wspace=0.3)
i=0
for ttype, v1 in I_data.items():
    i += 1
    ax = fig.add_subplot(2, math.ceil(len(I_data)/2), i)
    H_I_data = collections.defaultdict(lambda: collections.defaultdict(dict))
    for product, v2 in v1.items():
        if("unit" in H_data[ttype][product]):
            H_I_data[product]["Direct"]["y"] = H_data[ttype][product]["y"]
            H_I_data[product]["Direct"]["x"] = H_data[ttype][product]["x"]
            H_I_data[product]["Direct"]["unit"] = H_data[ttype][product]["unit"]
            H_I_data[product]["Indirect"]["y"] = I_data[ttype][product]["y"]
            H_I_data[product]["Indirect"]["x"] = I_data[ttype][product]["x"]

    fig = plt.figure(num=None, figsize=(8, 6), facecolor='w', edgecolor='k')
    fig.subplots_adjust(hspace=0.3, wspace=0.3)
    i=0
    for product, z1 in sorted(H_I_data.items()):
        search = re.search(r"^(\w)Z(\d+)$", product, re.IGNORECASE).groups()
        char = search[0]
        label= iden_label(int(search[1]))        
        linestyle = iden_company(char)
        i += 1
        ax = fig.add_subplot(2, math.ceil(len(H_I_data)/2), i)
        for dir_indir, z2 in z1.items():
            #print(ttype, product, dir_indir, dir_indir[0:1], z2["x"], z2["y"])
            label = TextPath((0,0), str(dir_indir[0:1]), linewidth=3)
            ax.plot(z2["x"], z2["y"], label=product+"_"+dir_indir, linestyle=linestyle, marker=label, markersize=20)
            ax.set_title(product)
        ax.set_xlabel('Period')
        ax.set_ylabel(z1["Direct"]["unit"])

    labels, handles = zip(*sorted(zip(labels, handles), key=lambda t: t[0]))
    fig.legend(handles, labels, loc='lower right')
    fig.suptitle(ttype, fontsize=20)
    fig.savefig('H_I_'+ttype+'.png')

# Draw H and I unit price vs. swot
fig = plt.figure(num=None, figsize=(8, 6), facecolor='w', edgecolor='k')
fig.subplots_adjust(hspace=0.3, wspace=0.3)
#i=0
for period, v1 in sorted(H_I_unit_data.items()):
#    i+=1
#    ax = fig.add_subplot(2, math.ceil(len(I_data)/2), i)
    for product, v2 in sorted(v1.items()):
#        print(period, method, v2["y"])
#        for num in range(len(v2["product"])):
        print(period, product, v2)
#            search = re.search(r"^(\w)", method, re.IGNORECASE).groups()
#            label = TextPath((0,0), str(char+label+search[0]), linewidth=3)
#            linestyle = iden_company(char)
#            ax.plot(v2["x"], v2["y"], label=v2["product"][num], linestyle="", marker=label, markersize=20)
#    ax.set_title("Period "+period)
#    ax.set_xlabel('Period')
#    ax.set_ylabel('Unit Price')
#labels, handles = zip(*sorted(zip(labels, handles), key=lambda t: t[0]))
#fig.legend(handles, labels, loc='lower right')
#fig.suptitle("Price/Swot", fontsize=20)
#fig.savefig('H_I_Swot_Period.png')
         


#with open("data", "rt") as f:
#    config = json.load(f)
#f.close()

#with open('data.json', 'w') as outfile:
#    json.dump(E_data, outfile)


