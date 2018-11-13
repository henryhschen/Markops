#!/usr/local/bin/python3.6

import math, re, json, collections, argparse, matplotlib, os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.text import TextPath

x_lower = 0
x_upper = 2
rpt_out = "output"

parser = argparse.ArgumentParser(description='Process for Markops')
parser.add_argument('-excel',nargs='+', help='Input excel files', required=True)

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

def iden_range(label):
    if(re.search(r"D", label, re.IGNORECASE)):
        x_min, x_max = 0, 30
    elif(re.search(r"C", label, re.IGNORECASE)):
        x_min, x_max = 30, 50
    elif(re.search(r"B", label, re.IGNORECASE)):
        x_min, x_max = 50, 70
    else:
        x_min, x_max = 70, 100
    return x_min, x_max

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

def iden_company_color(char):
    if(re.search(r"^A$", char, re.IGNORECASE)):
        color = "r"
    elif(re.search(r"^K$", char, re.IGNORECASE)):
        color = "g"
    elif(re.search(r"^P$", char, re.IGNORECASE)):
        color = "y"
    else: 
        color = "b"
    return color

# Main
data = collections.defaultdict(lambda: collections.defaultdict(dict))
for f in sorted(args.excel):
    xls = pd.ExcelFile(f)
    f_name = re.search(r"Period(\d)\.", f, re.IGNORECASE).groups()[0]
    for sheet in sorted(xls.sheet_names):
        parse = xls.parse(sheet)
        sheet_name = re.search(r"^(\w)_", sheet, re.IGNORECASE).groups()[0]
        if(sheet_name == "Z"):
            parse = parse.fillna(0)
        data[f_name][sheet_name.upper()]= parse

# Create output directory
rpt_out += "_"+f_name
if(os.path.isdir(rpt_out)):
    os.system("rm -rf "+rpt_out)
os.system("mkdir "+rpt_out)
writer = pd.ExcelWriter(rpt_out+"/summary.xlsx")

# Deal with B for competitive
B_data_extra = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(dict)))
for f, v1 in sorted(data.items()):
    for c_i, value in v1["B"].iteritems():
        B_data_extra[f][c_i]["UVMC"] = value["UVMC"]
        B_data_extra[f][c_i]["IMSP"] = value["IMSP"]

# Deal with C for competitive
C_data_extra = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(dict)))
for f, v1 in sorted(data.items()):
    for c_i, value in v1["C"].iteritems():
        C_data_extra[f][c_i]["S"] = value["S"]
        C_data_extra[f][c_i]["CGS"] = value["CGS"]
        C_data_extra[f][c_i]["FMC"] = value["FMC"]

new_B_C = pd.DataFrame()
new_C = pd.DataFrame()
for per, v1 in sorted(C_data_extra.items()):
    for pro, v2 in sorted(v1.items()):

        # For B_C
        if(re.search(r"^Z", pro, re.IGNORECASE)):
            new_B_C.loc[per+"_"+pro, "UVMC"] = B_data_extra[per][pro]["UVMC"]
            new_B_C.loc[per+"_"+pro, "FMC"] = C_data_extra[per][pro]["FMC"]
            new_B_C.loc[per+"_"+pro, "UVMC/FMC"] = np.single(B_data_extra[per][pro]["UVMC"]/C_data_extra[per][pro]["FMC"])
            new_B_C.loc[per+"_"+pro, "UVMC/FMC+UVMC"] = np.single(B_data_extra[per][pro]["UVMC"]/C_data_extra[per][pro]["FMC"]+B_data_extra[per][pro]["UVMC"])


        #print(per, pro, v2)
        new_C.loc[per+"_"+pro, "Product"] = pro
        new_C.loc[per+"_"+pro, "S"] = v2["S"]
        if(re.search(r"^Z", pro, re.IGNORECASE)):
            new_C.loc[per+"_"+pro, "S/S_total"] = np.single(v2["S"]/C_data_extra[per]["Total"]["S"]*100)
        else:
            new_C.loc[per+"_"+pro, "S/S_total"] = "Nan"
        new_C.loc[per+"_"+pro, "CGS"] = v2["CGS"]
        if(re.search(r"^Z", pro, re.IGNORECASE)):
            new_C.loc[per+"_"+pro, "CGS/S_total"] = np.single(v2["CGS"]/C_data_extra[per]["Total"]["S"]*100)
        else:
            new_C.loc[per+"_"+pro, "CGS/S_total"] = "Nan"
        new_C.loc[per+"_"+pro, "FMC"] = v2["FMC"]
        if(re.search(r"^Z", pro, re.IGNORECASE)):
            new_C.loc[per+"_"+pro, "FMC/S_total"] = np.single(v2["CGS"]/C_data_extra[per]["Total"]["S"]*100)
        else:
            new_C.loc[per+"_"+pro, "FMC/S_total"] = "Nan"
        new_C.loc[per+"_"+pro, "period"] = per
        new_C.loc[per+"_"+pro, "product"] = pro

new_B_C.to_excel(writer, "B_UVMC_FMC")
new_C.to_excel(writer, "C_CGS_FMC_S")

# Deal with E for competitive
E_data = collections.defaultdict(lambda: collections.defaultdict(dict))
E_data_extra = pd.DataFrame()
E_data_extra2 = pd.DataFrame()
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
            E_data["Surplus"][row[0]+row[1]]["y"] = [row["PV"]-row["UTS"]]
            E_data["PC"][row[0]+row[1]]["unit"] = "Volume"
            E_data["PC"][row[0]+row[1]]["x"] = [f]
            E_data["PC"][row[0]+row[1]]["y"] = [row["PC"]]
            E_data["PV"][row[0]+row[1]]["unit"] = "Volume"
            E_data["PV"][row[0]+row[1]]["x"] = [f]
            E_data["PV"][row[0]+row[1]]["y"] = [row["PV"]]

        #For new data
        swot = int(row[1].replace("Z", ""))
        label= iden_label(swot)
        for col in v1["E"].columns[3:]:
            E_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", col] = row[col]
        E_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", "period"] = f
        E_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", "sale_method"] = "Direct"
        E_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", "company"] = row[0]
        E_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", "seg"] = label
        E_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", "product"] = row[1]

        # Deal with PC & PV
        E_data_extra2.loc[f+label+"_"+row[0]+row[1], "PV"] = row["PV"]
        E_data_extra2.loc[f+label+"_"+row[0]+row[1], "PC"] = row["PC"]
        E_data_extra2.loc[f+label+"_"+row[0]+row[1], "Utilizaton"] = np.single(row["PV"]/row["PC"]*100)
        E_data_extra2.loc[f+label+"_"+row[0]+row[1], "period"] = f
        E_data_extra2.loc[f+label+"_"+row[0]+row[1], "seg"] = label
        E_data_extra2.loc[f+label+"_"+row[0]+row[1], "company"] = row[0]
        E_data_extra2.loc[f+label+"_"+row[0]+row[1], "product"] = row[1]


E_data_extra = E_data_extra.sort_index()
E_data_extra2 = E_data_extra2.sort_index()
E_data_extra2.to_excel(writer, "E_PV_PC")


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
fig.savefig(rpt_out+'/E.png')

# Deal with F for competitive
F_data = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(dict)))
F_data2 = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(dict)))
F_data_extra = pd.DataFrame()
for f, v1 in sorted(data.items()):
    for r_i, row in v1["F"].iterrows():
        swot = int(row[1].replace("Z", ""))
        label= iden_label(swot)

        if(row["MDS_P"] != 0):
            if("x" in F_data["Direct"][label][row[0]+row[1]]):
                F_data["Direct"][label][row[0]+row[1]]["x"].append(float(f))
                F_data["Direct"][label][row[0]+row[1]]["ym"].append(row["MDS_P"])
                F_data["Direct"][label][row[0]+row[1]]["yu"].append(row["UDS_P"])
            else: 
                F_data["Direct"][label][row[0]+row[1]]["x"] = [float(f)]
                F_data["Direct"][label][row[0]+row[1]]["ym"] = [row["MDS_P"]]
                F_data["Direct"][label][row[0]+row[1]]["unitm"] = "MDS_P"
                F_data["Direct"][label][row[0]+row[1]]["yu"] = [row["UDS_P"]]
                F_data["Direct"][label][row[0]+row[1]]["unitu"] = "UDS_P"

        if(row["MIS_P"] != 0):
            if("x" in F_data["Indirect"][label][row[0]+row[1]]):
                F_data["Indirect"][label][row[0]+row[1]]["x"].append(float(f))
                F_data["Indirect"][label][row[0]+row[1]]["ym"].append(row["MIS_P"])
                F_data["Indirect"][label][row[0]+row[1]]["yu"].append(row["UIS_P"])
            else: 
                F_data["Indirect"][label][row[0]+row[1]]["x"] = [float(f)]
                F_data["Indirect"][label][row[0]+row[1]]["ym"] = [row["MIS_P"]]
                F_data["Indirect"][label][row[0]+row[1]]["unitm"] = "MIS_P"
                F_data["Indirect"][label][row[0]+row[1]]["yu"] = [row["UIS_P"]]
                F_data["Indirect"][label][row[0]+row[1]]["unitu"] = "UIS_P"
            
        if(row["MS_P"] != 0):
            if("x" in F_data["MarketShare"][label][row[0]+row[1]]):
                F_data["MarketShare"][label][row[0]+row[1]]["x"].append(float(f))
                F_data["MarketShare"][label][row[0]+row[1]]["ym"].append(row["MS_P"])
                F_data["MarketShare"][label][row[0]+row[1]]["yu"].append(row["US_P"])
            else: 
                F_data["MarketShare"][label][row[0]+row[1]]["x"] = [float(f)]
                F_data["MarketShare"][label][row[0]+row[1]]["ym"] = [row["MS_P"]]
                F_data["MarketShare"][label][row[0]+row[1]]["unitm"] = "MS_P"
                F_data["MarketShare"][label][row[0]+row[1]]["yu"] = [row["US_P"]]
                F_data["MarketShare"][label][row[0]+row[1]]["unitu"] = "US_P"
            
        if(row["MTS_P"] != 0):
            if("x" in F_data["Total"][row[0]+row[1]]):
                F_data["Total"][row[0]+row[1]]["x"].append(float(f))
                F_data["Total"][row[0]+row[1]]["ym"].append(row["MTS_P"])
                F_data["Total"][row[0]+row[1]]["yu"].append(row["UTS_P"])
            else: 
                F_data["Total"][row[0]+row[1]]["x"] = [float(f)]
                F_data["Total"][row[0]+row[1]]["ym"] = [row["MTS_P"]]
                F_data["Total"][row[0]+row[1]]["unitm"] = "MTS_P"
                F_data["Total"][row[0]+row[1]]["yu"] = [row["UTS_P"]]
                F_data["Total"][row[0]+row[1]]["unitu"] = "UTS_P"

        #For new data
        swot = int(row[1].replace("Z", ""))
        label= iden_label(swot)
        for col in v1["F"].columns[3:]:
            F_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", col] = row[col]
        F_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", "period"] = f
        F_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", "sale_method"] = "Direct"
        F_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", "company"] = row[0]
        F_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", "seg"] = label
        F_data_extra.loc[f+label+"_"+row[0]+row[1]+"_Direct", "product"] = row[1]

        #For F new data2
        if(label in F_data2[f][row[0]]):
            F_data2[f][row[0]][label] += row["MTS_P"]
        else:
            F_data2[f][row[0]][label] = row["MTS_P"]

F_data_extra = F_data_extra.sort_index()

fig = plt.figure(num=None, figsize=(8, 6), facecolor='w', edgecolor='k')
fig.subplots_adjust(hspace=0.3, wspace=0.3)
ax = fig.add_subplot(111)

# For new data2
new_F = pd.DataFrame()
for period, v1 in sorted(F_data2.items()):
    for company, v2 in sorted(v1.items()):
        com_sum = 0
        for label, v3 in sorted(v2.items()):
            com_sum += v3
            new_F.loc[period+label+"_"+company, "period"] = period
            new_F.loc[period+label+"_"+company, "company"] = company
            new_F.loc[period+label+"_"+company, "label"] = label
            new_F.loc[period+label+"_"+company, "Total_MS"] = v3
        new_F.loc[period+"_"+company, "period"] = period
        new_F.loc[period+"_"+company, "company"] = company
        new_F.loc[period+"_"+company, "label"] = "ALL"
        new_F.loc[period+"_"+company, "Total_MS"] = com_sum
        #print(period, company, label, v3)
new_F.to_excel(writer, "F_TotalMS")

def autolabel_value(rects):
    for rect in rects:
        h = rect.get_height()
        ax.text(rect.get_x()+rect.get_width()/2., rect.get_y()+rect.get_height()+0.01, '%d'%int(h),
                ha='center', va='bottom', color='b')
def autolabel_company(rects, label):
    for rect in rects:
        h = rect.get_height()
        ax.text(rect.get_x()+rect.get_width()/2., rect.get_y()+rect.get_height()+5, label, ha='center', va='bottom', color='b')

bar_width = 0.2
bar_width_t = 0.05
for ttype, v1 in F_data.items():
    if(re.search(r"Total", ttype)):
        fig = plt.figure(num=None, figsize=(8, 6), facecolor='w', edgecolor='k')
        fig.subplots_adjust(hspace=0.3, wspace=0.3)
        i = 0
        ax = fig.add_subplot(111)
        save_fig = []
        save_labels = []
        save_company = []
        for product, v2 in sorted(v1.items()):
            i += 1
            search = re.search(r"^(\w)Z(\d+)$", product, re.IGNORECASE).groups()
            char = search[0]
            label= char+iden_label(int(search[1]))
            save_company.append(char+search[1])
            color = iden_company_color(search[0])
            save_labels.append(char)
            save_fig.append(ax.bar([x+bar_width_t*i for x in v2["x"]], v2["ym"], width=bar_width_t, label=label, align='center', color=color, edgecolor="black"))
        
        ax.legend( save_fig, save_company, loc='upper left', ncol=4, fontsize="x-small") 
        for fig_i in range(len(save_fig)):
            autolabel_value(save_fig[fig_i])
            autolabel_company(save_fig[fig_i], save_labels[fig_i])
        ax.set_title("Integration")
        ax.set_xlim([x_lower,x_upper])
        #ax.set_ylim([0,100])
        ax.set_xlabel('Period')
        ax.set_ylabel(v2["unitm"])
        fig.suptitle('F. Money_MS_'+ttype, fontsize=20)
        fig.savefig(rpt_out+'/F_Money_MS_'+ttype+'.png')

    else:
        fig = plt.figure(num=None, figsize=(8, 6), facecolor='w', edgecolor='k')
        fig.subplots_adjust(hspace=0.3, wspace=0.3)
        j = 0 
        for quality, v2 in sorted(v1.items()):
            i = 0
            j += 1
            ax = fig.add_subplot(2, math.ceil(len(v1.keys())/2), j)
            save_fig = []
            save_labels = []
            save_company = []
            for product, v3 in sorted(v2.items()):
                i += 1
                search = re.search(r"^(\w)Z(\d+)$", product, re.IGNORECASE).groups()
                char = search[0]
                label= char+iden_label(int(search[1]))
                save_company.append(char+search[1])
                color = iden_company_color(search[0])
                save_labels.append(char)
                #print(search[0], color)
                save_fig.append(ax.bar([x+bar_width*i for x in v3["x"]], v3["ym"], width=bar_width, label=label, align='center', color=color, edgecolor="black"))
                #print(ttype, quality, product, v3["x"], v3["ym"])


            ax.legend( save_fig, save_company, loc='upper left', ncol=4, fontsize="x-small") 
            for fig_i in range(len(save_fig)):
                autolabel_value(save_fig[fig_i])
                autolabel_company(save_fig[fig_i], save_labels[fig_i])
            ax.set_title("Quality: "+quality)
            ax.set_xlim([x_lower,x_upper])
            ax.set_ylim([0,100])
            #ax.autoscale(tight=True)
            ax.set_xlabel('Period')
            ax.set_ylabel(v3["unitm"])
        if(re.search(r"MarketShare", ttype, re.IGNORECASE)):
            real_type = "Direct+Indirect"
        else:
            real_type = ttype
        fig.suptitle('F. Money_MS_'+real_type, fontsize=20)
        fig.savefig(rpt_out+'/F_Money_MS_'+real_type+'.png')


# Deal with H for direct competitive
H_data = collections.defaultdict(lambda: collections.defaultdict(dict))
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
  
H_I_up_sort_period = pd.DataFrame()
for f, v1 in sorted(data.items()):
    for r_i, row in v1["H"].iterrows():
        for col in v1["H"].columns[2:3]:
            #print(col, row[col], row[0]+row[1])
            swot = int(row[1].replace("Z", ""))
            unit_price =np.single(row[col]/swot)
            label= iden_label(swot)
            H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "unitprice"] = unit_price
            H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "swot"] = swot
            H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "UP"] = row[col]
            H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "seg"] = label
        
        SF = row.loc["SF"]
        SS = row.loc["SS"]
        TS = row.loc["TS"]
        total = SF + SS + TS
        SF_per = np.single(SF/total*100)
        SS_per = np.single(SS/total*100)
        TS_per = np.single(TS/total*100)
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "SF_per"] = SF_per
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "SS_per"] = SS_per
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "TS_per"] = TS_per
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "Total_Market_per"] = SF_per + SS_per + TS_per
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "SF"] = SF
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "SS"] = SS
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "TS"] = TS
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "Total_Market"] = total

        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "period"] = f
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "sale_method"] = "Direct"
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "company"] = row[0]
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "seg"] = label
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Direct", "product"] = row[1]


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
fig.savefig(rpt_out+'/H.png')


# Deal with I for indirect competitive
I_data = collections.defaultdict(lambda: collections.defaultdict(dict))
I_data_extra = collections.defaultdict(lambda: collections.defaultdict(dict))
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
        swot = int(row[1].replace("Z", ""))
        if(re.search(r"^W", row["Company"], re.IGNORECASE)):
            #print(r_i, row)
            I_data_extra[f][row["Product"]]["UP"] = row["UP"]
            I_data_extra[f][row["Product"]]["number"] = swot

        #C_data_extra[f][c_i]["S"] = value["S"]

        for col in v1["I"].columns[2:3]:
            #print(col, row[col], row[0]+row[1])
            swot = int(row[1].replace("Z", ""))
            unit_price =np.single(row[col]/swot)
            label= iden_label(swot)
            H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "unitprice"] = unit_price
            H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "swot"] = swot
            H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "UP"] = row[col]
            H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "seg"] = label

        SF = row.loc["SF"]
        SS = row.loc["SS"]
        TS = row.loc["TS"]
        total = SF + SS + TS
        SF_per = np.single(SF/total*100)
        SS_per = np.single(SS/total*100)
        TS_per = np.single(TS/total*100)
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "SF_per"] = SF_per
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "SS_per"] = SS_per
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "TS_per"] = TS_per
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "Total_Market_per"] = SF_per + SS_per + TS_per
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "SF"] = SF
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "SS"] = SS
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "TS"] = TS
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "Total_Market"] = total

        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "period"] = f
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "sale_method"] = "Indirect"
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "company"] = row[0]
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "seg"] = label
        H_I_up_sort_period.loc[f+label+"_"+row[0]+row[1]+"_Indirect", "product"] = row[1]



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
fig.savefig(rpt_out+'/I.png')

new_B_I = pd.DataFrame()
for per, v1 in sorted(I_data_extra.items()):
    for pro, v2 in sorted(v1.items()):
        #print(f, pro, B_data_extra[per][pro]["IMSP"], I_data_extra[per][pro]["UP"])
        new_B_I.loc[per+"_"+pro+"_Indirect", "IMSP"] = B_data_extra[per][pro]["IMSP"]
        new_B_I.loc[per+"_"+pro+"_Indirect", "UP"] = I_data_extra[per][pro]["UP"]
        new_B_I.loc[per+"_"+pro+"_Indirect", "UP/Swot"] = np.single(I_data_extra[per][pro]["UP"]/I_data_extra[per][pro]["number"])
        new_B_I.loc[per+"_"+pro+"_Indirect", "UP-IMSP+UP*1.25"] = I_data_extra[per][pro]["UP"] - B_data_extra[per][pro]["IMSP"] + B_data_extra[per][pro]["IMSP"]*1.25
        new_B_I.loc[per+"_"+pro+"_Indirect", "(UP-IMSP+UP*1.25)/UP"] = np.single((I_data_extra[per][pro]["UP"] - B_data_extra[per][pro]["IMSP"] + B_data_extra[per][pro]["IMSP"]*1.25)/I_data_extra[per][pro]["number"])
        new_B_I.loc[per+"_"+pro+"_Indirect", "100*(UP-IMSP)/UP"] = np.single((I_data_extra[per][pro]["UP"] - B_data_extra[per][pro]["IMSP"]) / I_data_extra[per][pro]["UP"] *100)

new_B_I.to_excel(writer, "B_I_indirect")

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
    fig.savefig(rpt_out+'/H_I_'+ttype+'.png')

# Draw H and I unit price vs. swot
H_I_up_sort_period = H_I_up_sort_period.sort_index()
#H_I_up_sort_period.to_csv("H_I_unitprice_by_period.csv")

# Deal with J for indirect competitive
J_data_extra = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(dict)))
for f, v1 in sorted(data.items()):
    J_ori = v1["J"]
    #print(f, v1["J"].keys())
    for c_i, value in v1["J"].iteritems():
        sear = re.search(r"^(MV)_(\d+)$", c_i, re.IGNORECASE)
        if(sear):
            sear = sear.groups()
            MV = sear[0]
            per = sear[1]
            for quality in ['A', 'B', 'C', 'D']:                
                #print(f, quality, MV, per, c_i)
                if(per == f):
                    #print("Current", quality)
                    J_data_extra[f][quality]["current"] = J_ori[ J_ori["Seg"]== quality ][c_i].values[0]
                else:
                    #print("Next")
                    J_data_extra[f][quality]["next"] = J_ori[ J_ori["Seg"]== quality ][c_i].values[0]
                #print(J_ori[ J_ori["Seg"]== quality ][c_i].values[0])

new_J = pd.DataFrame()
for per, v1 in sorted(J_data_extra.items()):
    for quality, v2 in sorted(v1.items()):
        #print(per, quality, v2["current"], v2["next"])
        new_J.loc[per+"_"+quality, "MV"] = v2["current"]
        new_J.loc[per+"_"+quality, "MV_Forecast"] = v2["next"]

        new_J.loc[per+"_"+quality, "MV_Forecast_div_current_per"] = v2["next"]/v2["current"]
        per_next = str(int(per) + 1)
        if("current" in J_data_extra[per_next][quality]):
            new_J.loc[per+"_"+quality, "MV_next_div_current_per"] = J_data_extra[per_next][quality]["current"]/v2["current"]
new_J.to_excel(writer, "J_MV")
        
# Real E_H_I
def color_larger_1000_red(val):
    color = 'red' if val > 1000 else 'black'
    return 'color: %s' % color

period_uniq = H_I_up_sort_period["period"].unique()
company_uniq = H_I_up_sort_period["company"].unique()
product_uniq = H_I_up_sort_period["product"].unique()
sale_method_uniq = H_I_up_sort_period["sale_method"].unique()
seg_uniq = H_I_up_sort_period["seg"].unique()

E_F_H_I_market_sort_comany =pd.DataFrame()
od = H_I_up_sort_period
od2 = E_data_extra
od3 = F_data_extra
for per in period_uniq:
    for c in company_uniq:
        tmp_pd = pd.DataFrame()
        tmp_pd2 = pd.DataFrame()
        tmp_pd3 = pd.DataFrame()
        for pro in product_uniq:
            for s in sale_method_uniq:
                data1 = od.loc[(od["period"] == per) & (od["company"] == c) & (od["product"] == pro) & (od["sale_method"] == s)]
                data2 = od2.loc[(od2["period"] == per) & (od2["company"] == c) & (od2["product"] == pro) & (od2["sale_method"] == s)]
                data3 = od3.loc[(od3["period"] == per) & (od3["company"] == c) & (od3["product"] == pro) & (od3["sale_method"] == s)]
                if(not data1.empty):
                    tmp_pd = pd.concat([tmp_pd, data1])
                if(not data2.empty):
                    tmp_pd2 = pd.concat([tmp_pd2, data2])
                if(not data3.empty):
                    tmp_pd3 = pd.concat([tmp_pd3, data3])
    
        tmp_pd = tmp_pd.sort_index()
        new_pd = pd.DataFrame()
        total_mt = 0
        for index, row in tmp_pd.iterrows():
            total_mt += row["SF"]
            total_mt += row["SS"]
            total_mt += row["TS"]
            new_pd.loc[index, "unitprice"] = row["unitprice"]
            new_pd.loc[index, "swot"] = row["swot"]
            new_pd.loc[index, "UP"] = row["UP"]
        for index, row in tmp_pd.iterrows():
            SF_per = np.single(row["SF"]/total_mt*100)
            SS_per = np.single(row["SS"]/total_mt*100)
            TS_per = np.single(row["TS"]/total_mt*100)
            new_pd.loc[index, "SF_per"] = SF_per
            new_pd.loc[index, "SS_per"] = SS_per
            new_pd.loc[index, "TS_per"] = TS_per
            new_pd.loc[index, "Total_Market_per"] = SF_per+SS_per+TS_per
            new_pd.loc[index, "SF"] = row["SF"]
            new_pd.loc[index, "SS"] = row["SS"]
            new_pd.loc[index, "TS"] = row["TS"]
            new_pd.loc[index, "Total_Market"] = total_mt
            
        # Add F sheet
        for index3, row3 in tmp_pd3.iterrows():
            for col3 in tmp_pd3.columns[:]:
                new_pd.loc[index3, col3] = row3[col3]
        
        # Add E sheet
        for index, row in tmp_pd2.iterrows():
            for col in tmp_pd2.columns[:]:
                new_pd.loc[index, col] = row[col]
       
        # Name
        for index, row in tmp_pd.iterrows():
            new_pd.loc[index, "period"] = row["period"]
            new_pd.loc[index, "sale_method"] = row["sale_method"]
            new_pd.loc[index, "company"] = row["company"]
            new_pd.loc[index, "seg"] = row["seg"]
            new_pd.loc[index, "product"] = row["product"]
        E_F_H_I_market_sort_comany = pd.concat([E_F_H_I_market_sort_comany, new_pd])
#E_F_H_I_market_sort_comany.to_csv("E_F_H_I_market_sort_comany.csv")
style = E_F_H_I_market_sort_comany.style.applymap(color_larger_1000_red, subset = ["SF", "SS", "TS"])
E_F_H_I_market_sort_comany.to_excel(writer, "E_F_H_I_market_sort_comany")
style.to_excel(writer, "E_F_H_I_market_sort_comany")

#E_F_H_I_market_sort_seg.to_csv("E_F_H_I_market_sort_seg.csv")
E_F_H_I_market_sort_seg = E_F_H_I_market_sort_comany.sort_values(['period' , 'seg', 'sale_method'])
style = E_F_H_I_market_sort_seg.style.applymap(color_larger_1000_red, subset = ["SF", "SS", "TS"])
E_F_H_I_market_sort_seg.to_excel(writer, "E_F_H_I_market_sort_seg")
style.to_excel(writer, "E_F_H_I_market_sort_seg")

# Draw swot vs. price
draw_data = collections.defaultdict(lambda: collections.defaultdict(dict))
draw_w_data = collections.defaultdict(lambda: collections.defaultdict(dict))
draw = H_I_up_sort_period
for per in period_uniq:
    for seg in seg_uniq:
        for s in sale_method_uniq:
            #print(per, seg, s)
            method = re.search(r"^(\w)", s, re.IGNORECASE).groups()[0]
            others = draw.loc[(draw["period"] == per) & (draw["seg"] == seg) & (draw["company"] != "W") & (draw["sale_method"] == s)]
            ones = draw.loc[(draw["period"] == per) & (draw["seg"] == seg) & (draw["company"] == "W") & (draw["sale_method"] == s)]
            
            # others
            if(not others.empty):
                for other, value in others.iterrows():
                    if("x" in draw_data[per][seg]):
                        draw_data[per][seg]["x"].append(value["swot"])
                        draw_data[per][seg]["y"].append(value["UP"])
                        draw_data[per][seg]["product"].append(value["company"]+str(int(value["swot"]))+method)
                    else:
                        #print(value["swot"], value["UP"], value["company"], value["product"], method)
                        draw_data[per][seg]["x"] = [value["swot"]]
                        draw_data[per][seg]["y"] = [value["UP"]]
                        draw_data[per][seg]["product"] = [value["company"]+str(int(value["swot"]))+method]
            #ones
            if(not ones.empty):
                for one, value in ones.iterrows():
                    if("x" in draw_w_data[per][seg]):
                        draw_w_data[per][seg]["x"].append(value["swot"])
                        draw_w_data[per][seg]["y"].append(value["UP"])
                        draw_w_data[per][seg]["product"].append(value["company"]+str(int(value["swot"]))+method)
                    else:
                        #print(value["swot"], value["UP"], value["company"], value["product"], method)
                        draw_w_data[per][seg]["x"] = [value["swot"]]
                        draw_w_data[per][seg]["y"] = [value["UP"]]
                        draw_w_data[per][seg]["product"] = [value["company"]+str(int(value["swot"]))+method]
        
fig = plt.figure(num=None, figsize=(8, 6), facecolor='w', edgecolor='k')
fig.subplots_adjust(hspace=0.3, wspace=0.3)
for per, v1 in sorted(draw_data.items()):
    for seg, v2 in sorted(v1.items()):
        ax = fig.add_subplot(111)
        # For other companies
        if(len(v2["x"]) > 1):
            fit = np.polyfit(v2["x"], v2["y"], 1)
            fit_fn = np.poly1d(fit)
            x_min, x_max= iden_range(seg)
            rrange = [i for i in range(x_min, x_max, 1)]
            #label = []
            #for p in v2["product"]:
                #label.append(TextPath(0,0), str(p))
            #    print(TextPath(0,0))
            #print(label)
            ax.plot(v2["x"], v2["y"], 'ko', rrange, fit_fn(rrange), '-k')

        else:
            ax.plot(v2["x"], v2["y"], 'ko')
        for i, label in enumerate(v2["product"]):
            #print("Others", i, label, v2["x"][i], v2["y"][i])
            ax.annotate(label, xy=(v2["x"][i], v2["y"][i]), xytext=(v2["x"][i], v2["y"][i]-10), color='black')

        # For W company
        if("x" in draw_w_data[per][seg]):
            ax.plot(draw_w_data[per][seg]["x"], draw_w_data[per][seg]["y"], 'go')
            for i, label in enumerate(draw_w_data[per][seg]["product"]):
                #print("W", i, label, v2["x"][i], v2["y"][i])
                ax.annotate(label, xy=(draw_w_data[per][seg]["x"][i], draw_w_data[per][seg]["y"][i]), xytext=(draw_w_data[per][seg]["x"][i], draw_w_data[per][seg]["y"][i]+10), color='green')
    ax.set_xlabel('Swot')
    ax.set_ylabel('Price')
    plt.xlim([0,100])
    fig.suptitle("Period"+per, fontsize=20)
    fig.savefig(rpt_out+'/F_price_period'+per+'.png')


# Z(Projdction) vs C(real) for product contribution
Z_C_data = collections.defaultdict(lambda: collections.defaultdict(lambda: collections.defaultdict(dict)))
for f, v1 in sorted(data.items()):
    for row in v1["Z"]:
        # Projection
        for ttype, value in v1["Z"][row].iteritems():
            #print(f, row, ttype, value)
            f_next = str(int(f)+1)
            Z_C_data[f_next][row][ttype]["project"] = value
        # Real
        for ttype, value in v1["C"][row].iteritems():
            Z_C_data[f][row][ttype]["real"] = value

new_Z_C = pd.DataFrame()

for period, v1 in sorted(Z_C_data.items()):
    for product, v2 in sorted(v1.items()):
        for ttype, v3 in v2.items():
            if("project" in v3 and "real" in v3):
                #print(period, product, ttype, v3)
                new_Z_C.loc[period+"_"+product+"_project", ttype] = v3["project"]
                new_Z_C.loc[period+"_"+product+"_real", ttype] = v3["real"]
                if(v3["real"] != 0 and v3["project"]!=0):
                    new_Z_C.loc[period+"_"+product+"_real", ttype+"_compare"] = np.single(v3["real"]/v3["project"]*100)
new_Z_C.to_excel(writer, "Z_C_project_vs_real")



#with open("data", "rt") as f:
#    config = json.load(f)
#f.close()

#with open('data.json', 'w') as outfile:
#    json.dump(E_data, outfile)

writer.save()
