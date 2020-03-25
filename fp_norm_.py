#!/usr/bin/python3

"""

"""
import sys
import subprocess
import pkg_resources

required = {'openpyxl', 'numpy', 'pandas', 're', 'statistics', 'math', 'mathplotlib.pyplot'}
installed = {pkg.key for pkg in pkg_resources.working_set}
missing = required - installed

if missing:
    python = sys.executable
    subprocess.check_call([python, '-m', 'pip', 'install', *missing], stdout=subprocess.DEVNULL)

# import sys
# import os
from openpyxl import load_workbook
import numpy as np
import pandas as pd
import re
from statistics import mean
import math
import matplotlib.pyplot as plt

# ################### 
# ## INPUTS

## File path to the xlsx file
file = sys.argv[1]
plotName = sys.argv[2]
resultsName = sys.argv[3]
sampleNames = sys.argv[4]




def analyzeFP(file, plotName, resultsName, sampleNames=None):

    ## Load excel workbook 
    wb = load_workbook(file)  
    
    ## Specify the sheet with raw values
    data = wb["All Cycles"]
    
    ## Pull values from workbook sheet into a pd df
    data = pd.DataFrame(data.values)
    
    ## Grab complete data (drop rows if there are more than 3 NaNs -- just in case)
    data = data.dropna(axis=0, thresh=3).reset_index(drop=True)
    
    ## Change header names to the first row
    data.columns = data.iloc[0]
    
    ## Grab the first time point
    timePoint1 = data.iloc[1,3]
    
    ## Only iter for the time (in seconds)
    timeIter = int(re.split(r"\s", timePoint1)[2])
    
    ## Drop all non-data (samples etc)
    data = data.iloc[2:, 1:]
    
    ## Grab all the sample names 
    dataNames = data.iloc[:,0]
    
    ## Regular expression to find the 'Polarized based on blank corrected' value
    data = data.filter(regex='Polarization based on Blank corrected')
    
    ## Combine sample names with the raw data
    data = pd.concat([dataNames, data], axis=1)
    
    
    ## Iterate to generate all the different time points in the experiment, 
    ## based on the first time point and the number of cycles present in raw data
    ## Also convert time to minutes format
    timePointAll = pd.DataFrame([x*(timeIter/60) for x in range(0,data.shape[1]-1)], 
                                columns = ['time'])
    
    
    ## Dictionary for keeping track of all reps (based on same sample ID)
    allReps = {}
    
    
    ## Loop through all data 
    for ii in range(len(data)):
        ## Grab sample name
        tempSample = data.iloc[ii,0]
        ## Grab sample data
        tempData = data.iloc[ii,1:].tolist()
        
        ## If this is the first time sample is seen
        ## Add new key/value pair [value = list of list(data)]
        if tempSample not in allReps.keys():
            allReps[tempSample] = [tempData]
        ## Else append to key with same name
        else:
            allReps[tempSample].append(tempData)
    
    
    
    
    ## taking the average if there are multiple technical controls
    if len(allReps[dataNames.iloc[0]]) == 1:
        avgReps = allReps
        
    else:    
        avgReps = {}
        
        for ii in allReps.keys():
            # print(ii)
            avgReps[ii] = [mean(k) for k in zip(allReps[ii][0], 
                                                allReps[ii][1], allReps[ii][2])]        
                
    pos = np.array(sum([value for key, value in avgReps.items() if 'positive' in key.lower()],[]))
    neg = np.array(sum([value for key, value in avgReps.items() if 'negative' in key.lower()],[]))
    
    
    
    
    normRep = {}
    
    for ii in avgReps.keys():
        if "Positive" in ii:
            pass
        elif "Negative" in ii:
            pass
        else:
            val = np.array(avgReps[ii])
            normRep[ii] = (val - pos) / (neg - pos) * 100
    
    dataNorm = pd.concat([timePointAll, pd.DataFrame.from_dict(normRep)], 
                                 axis=1)
    
   
    if sampleNames != None:
        newSampleNames = ['time']
        sampleNames = map(str, sampleNames)
        newSampleNames.extend(sampleNames)
        dataNorm.columns = newSampleNames
    
    
    dataNorm.to_csv(resultsName, index=False)
    
    dataNormMelt = pd.melt(dataNorm,id_vars=['time'],
                           var_name='samples', value_name='values')
    
    
    ## Plotting the normalized data
    
    fig, ax = plt.subplots()
    colors = plt.cm.Spectral(np.linspace(0,1,dataNormMelt['samples'].nunique()))
    ax.set_prop_cycle('color', colors)
    
    for key, grp in dataNormMelt.groupby(['samples']):
        ax = grp.plot(ax=ax, kind='line', 
                      x='time', y='values', label=key, 
                      figsize=(16,12), 
                      yticks=range(0, int(math.ceil(max(dataNormMelt['values'])/10.0)*10), 10), 
                      xticks=range(0, int(math.ceil(max(dataNormMelt['time']) /10.0)*10), 10),
                      sort_columns = False,
                      fontsize = 16)
    ax.set(xlabel = 'Time (min)')
    ax.set(ylabel = '% Substrate Remaining')
    plt.rc('axes', labelsize=16)
    plt.rc('legend', fontsize=12) 
    plt.legend(loc='best')
    fig.savefig(plotName, bbox_inches='tight', transparent=True)
    plt.show()



if __name__ == '__main__':
    if len(sys.argv) == 4:
        analyzeFP(file, plotName, resultsName, sampleNames=sampleNames)
    else:
        analyzeFP(file, plotName, resultsName, sampleNames=None)




   