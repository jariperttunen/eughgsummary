import os
import argparse
import pathlib
import glob
import pandas as pd
import numpy as np
from countrylist import euls,euplusls,noneuls,allcountryls

land_use_change_table_ls = ['Table4.A','Table4.B','Table4.C','Table4.D','Table4.E','Table4.F']
land_use_change_rows_ls = ['2.1','2.2','2.3','2.4','2.5']
#Table4.D Wetlands is a special case, but these name should find the right rows
land_use_change_wl_rows_ls = ['2.1. Land','2.2 Land','2.3 Land']
land_use_change_sheet_name_ls = ['CL-FL','GL-FL','WL-FL','SL-FL','OL-FL',
                                 'FL-CL','GL-CL','WL-CL','SL-CL','OL-CL',
                                 'FL-GL','CL-GL','WL-GL','SL-GL','OL-GL',
                                 '2.1Land-WLpeat_extraction','2.2Land-WLflooded','2.3Land-WLother',
                                 'FL-SL','CL-SL','GL-SL','WL-SL','OL-SL',
                                 'FL-OL','CL-OL','GL-OL','WL-OL','SL-OL'
                                 ]

def EULandUseChange(excel_writer,directory,countryls,sheet:str,sheet_name:str,from_row:str,from_col:int,start:int,end:int):
    """
    Read CRFReporter Reporting tables Table4.A-Table4-F and create Land Use Change Total Area sheets for each country and year.
    \note The algorithm is identical to the one in *eulandtransitionmatrix.py*
    \param writer Excel writer
    \param directory The direactory where the Reporting tables are located
    \param countryls List of countries to be used
    \param sheet Land Transition Matrix sheet name
    \param sheet_name Excel sheet for a single Land Transition results
    \param from_row Name of the row name in *from_ls* (*FROM* Land Use classes Table4.1)
    \param from_col Column number for (total) area Land Use Class
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
        print(country,sheet,sheet_name)
        for excel_file in excelfilels:
            xlsx = pd.ExcelFile(excel_file)
            df = pd.read_excel(xlsx,sheet,keep_default_na=False,na_values=[''])
            row = df[df[df.columns[0]].str.contains(from_row)==True]
            rowls.append(row.iloc[0,from_col])
        datarowlss.append(rowls)
    dftotal = pd.DataFrame(datarowlss)
    dftotal.index = countryls
    dftotal.columns =  list(range(start,end+1))
    dftotal.to_excel(excel_writer,sheet_name,na_rep='NaN')

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
        #The ??? mean three letter acronyms matched 
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

    excel_writer = pd.ExcelWriter(file_prefix+'_Table4A-Table4F_land_use_change'+str(inventory_start)+'_'+str(inventory_end)+'.xlsx',engine='xlsxwriter')
    sheet_name_index=0
    for land_use_sheet in land_use_change_table_ls:
        land_use_row_ls = land_use_change_rows_ls
        #Special case with Wetlands, only three rows
        if land_use_sheet == 'Table4.D':
            land_use_row_ls = land_use_change_wl_rows_ls 
        for land_use_row in land_use_row_ls:
            EULandUseChange(excel_writer,directory,countryls,land_use_sheet,land_use_change_sheet_name_ls[sheet_name_index],land_use_row,2,inventory_start,inventory_end)
            sheet_name_index=sheet_name_index+1
    excel_writer.save()
