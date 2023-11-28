import os
import argparse
import pathlib
import glob
import pandas as pd
import numpy as np
from countrylist import euls,euplusls,noneuls,allcountryls

land_transition_matrix_sheet = 'Table4.1'
#Rows for FROM Land Use Class in Table4.1
from_ls = ["Forest land \(managed\)","Forest land \(unmanaged\)","Cropland","Grassland \(managed\)","Grassland \(unmanaged\)",
           "Wetlands \(managed","Wetlands \(unmanaged\)","Settlements","Other land"]
#Result sheet names for Land use change classes
sheet_name_dict = {0:['FL(manag.)->FL(manag.)','FL(manag.)->FL(unmanag.)','FL(manag.)->CL',
                      'FL(manag.)->GL(manag.)','FL(manag.)->GL(unmanag.)','FL(manag.)->WL(manag.)','FL(manag.)->WL(unmanag.)',
                      'FL(manag.)->SL','FL(manag.)->OL'],
                   1:['FL(unmanag.)->FL(manag.)','FL(unmanag.)->FL(unmanag.)','FL(unmanag.)->CL',
                      'FL(unmanag.)->GL(manag.)','FL(unmanag.)->GL(unmanag.)','FL(unmanag.)->WL(manag.)','FL(unmanag.)->WL(unmanag.)',
                      'FL(unmanag.)->SL','FL(unmanag.)->OL'],
                   2:['CL->FL(manag.)','CL->FL(unmanag.)','CL->CL',
                      'CL->GL(manag.)','CL->GL(unmanag.)','CL->WL(manag.)','CL->WL(unmanag.)',
                      'CL->SL','CL->OL'],
                   3:['GL(manag.)->FL(manag.)','GL(manag.)->FL(unmanag.)','GL(manag.)->CL',
                      'GL(manag.)->GL(manag.)','GL(manag.)->GL(unmanag.)','GL(manag.)->WL(manag.)','GL(manag.)->WL(unmanag.)',
                      'GL(manag.)->SL','GL(manag.)->OL'],
                   4:['GL(unmanag.)->FL(manag.)','GL(unmanag.)->FL(unmanag.)','GL(unmanag.)->CL',
                      'GL(unmanag.)->GL(manag.)','GL(unmanag.)->GL(unmanag.)','GL(unmanag.)->WL(manag.)','GL(unmanag.)->WL(unmanag.)',
                      'GL(unmanag.)->SL','GL(unmanag.)->OL'],
                   5:['WL(manag.)->FL(manag.)','WL(manag.)->FL(unmanag.)','WL(manag.)->CL',
                      'WL(manag.)->GL(manag.)','WL(manag.)->GL(unmanag.)','WL(manag.)->WL(manag.)','WL(manag.)->WL(unmanag.)',
                      'WL(manag.)->SL','WL(manag.)->OL'],
                   6:['WL(unmanag.)->FL(manag.)','WL(unmanag.)->FL(unmanag.)','WL(unmanag.)->CL',
                      'WL(unmanag.)->GL(manag.)','WL(unmanag.)->GL(unmanag.)','WL(unmanag.)->WL(manag.)','WL(unmanag.)->WL(unmanag.)',
                      'WL(unmanag.)->SL','WL(unmanag.)->OL'],
                   7:['SL->FL(manag.)','SL->FL(unmanag.)','SL->CL',
                      'SL->GL(manag.)','SL->GL(unmanag.)','SL->WL(manag.)','SL->WL(unmanag.)',
                      'SL->SL','SL->OL'],
                   8:['OL->FL(manag.)','OL->FL(unmanag.)','OL->CL',
                      'OL->GL(manag.)','OL->GL(unmanag.)','OL->WL(manag.)','OL->WL(unmanag.)',
                      'OL->SL','OL->OL']
                   }

def CreateLandTransitionMatrix(writer,directory,countryls,sheet:str,sheet_name:str,from_row:str,to_col:int,start:int,end:int):
    """
    Read CRFReporter Reporting tables and create Land transition sheets for each country and year.
    Use Excel sheet Table 4.1 Land Transition Matrix
    \param writer Excel writer
    \param directory The direactory where the Reporting tables are located
    \param countryls List of countries to be used
    \param sheet Land Transition Matrix sheet name
    \param sheet_name Excel sheet for a single Land Transition results
    \param from_row Name of the row name in *from_ls* (*FROM* Land Use classes Table4.1)
    \param to_col Column number for *TO*  Land Use Class
    \param start Inventory start year (1990)
    \param end Inventory end year
    """
    
    datarowlss=[]
    for country in countryls:
        rowls=[]
        #List all excel files and sort the files in ascending order (1990,1991,...,2021)
        #Exclude years in 1980's that some countries report
        excelfilels=list(set(glob.glob(directory+'/'+country+'/*.xlsx'))-set(glob.glob(directory+'/'+country+'/*_198??*.xlsx')))
        excelfilels=sorted(excelfilels)
        print(country,sheet_name)
        for file in excelfilels:
            xlsx = pd.ExcelFile(file)
            df = pd.read_excel(xlsx,sheet,keep_default_na=False,na_values=[''])
            row = df[df[df.columns[0]].str.contains(from_row)==True]
            rowls.append(row.iloc[0,to_col])
        datarowlss.append(rowls)
    dftotal = pd.DataFrame(datarowlss)
    dftotal.index = countryls
    dftotal.columns =  list(range(start,end+1))
    dftotal.to_excel(writer,sheet_name,na_rep='NaN')
    
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-d","--directory",dest="f1",required=True,help="Inventory Parties Directory")
    parser.add_argument("-s","--start",dest="f2",required=True,help="Inventory start year (usually 1990)")
    parser.add_argument("-e","--end",dest="f3",required=True,help="Inventory end year")
    group=parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--eu",action="store_true",dest="eu",default=False,help="EU countries")
    group.add_argument("--euplus",action="store_true",dest="euplus",default=False,help="EU countries plus GBR, ISL and NOR")
    group.add_argument("-a","--all",action="store_true",dest="all",default=False,help="All countries (EU+others")
    group.add_argument("-c","--countries",dest="country",type=str,nargs='+',help="List of countries")
    group.add_argument("-l","--list",action="store_true",dest="countryls",default=False,
                       help="List files in Inventory Parties Directory")
              
    args = parser.parse_args()
    directory=args.f1
    print("Inventory Parties directory",directory)
    inventory_start=int(args.f2)
    print("Inventory start",inventory_start)
    inventory_end=int(args.f3)
    print("Inventory end",inventory_end)
    file_prefix = 'EU'
    if args.eu:
        print("Using EU  countries")
        countryls=euls
    elif args.euplus:
        print("Using EU  countries plus GBR, ISL, NOR")
        countryls=euplusls
        file_prefix = 'EU_GBR_ISL_NOR'
    elif args.all:
        print("Using all countries")
        countryls = euls+noneuls
        file_prefix='EU_and_Others'
    elif args.countryls:
        print("Listing countries in",args.f1)
        ls = glob.glob(args.f1+'/???')
        countryls = [pathlib.Path(x).name for x in ls]
        countryls.sort()
        file_prefix = pathlib.Path(args.f1).name
    else:
        print("Using countries", args.country) 
        countryls=args.country
        file_prefix=countryls[0]
        for country in countryls[1:]:
            file_prefix = file_prefix+"_"+country

    writer = pd.ExcelWriter(file_prefix+'_Table4.1_Land_Transition_Matrix_'+str(inventory_start)+'_'+str(inventory_end)+'.xlsx',
                            engine='xlsxwriter')
    #1. Table4.1 Land transition matrix
    index = 0
    for land_use_class in from_ls:
        col=1
        sheet_name_ls = sheet_name_dict[index]
        index=index+1
        for sheet_name in sheet_name_ls:
            CreateLandTransitionMatrix(writer,directory,countryls,land_transition_matrix_sheet,sheet_name,land_use_class,col,
                                       inventory_start,inventory_end)
            col=col+1
    writer.close()
