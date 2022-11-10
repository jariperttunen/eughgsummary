import os
import argparse
import pathlib
import glob
import pandas as pd
import numpy as np
from countrylist import euls,euplusls,noneuls,allcountryls
import EUutility

#Excel sheets needed
sheet_ls = ['Table4.A','Table4.B','Table4.C','Table4.D']
#Search dataframe row with these names
substr_ls = ['A. Total forest land','B. Total Cropland','C. Total grassland','D. Total wetlands']
#Column names in the resulting excel. Note a bit different sheet structure.
Table4A_columns_ls = ['TotalArea(kha)','MineralSoil(kha)','OrganicSoil(kha)','LivingBMGains(tC/ha)','LivinBMLosses(tC/ha)','LivingBMNetChange(tC/ha)',
                      'DeadWoodNetChange(tC/ha)','LitterNetChange(tC/ha)','MineralSoilsNetChange(tC/ha)','OrganicSoilsNetChange(tC/ha)',
                      'LivingBMGains(ktC)','LivingBMLosses(ktC)','LivingBMNetChange(ktC)','DeadwoodNetChange(ktC)','LitterNetChange(ktC)',
                      'MineralSoilsNetChange(ktC)','OrganicSoilsNetChange(ktC)','NetCO2(ktCO2)']
Table4B_columns_ls = ['TotalArea(kha)','MineralSoil(kha)','OrganicSoil(kha)','LivingBMGains(tC/ha)','LivinBMLosses(tC/ha)','LivingBMNetChange(tC/ha)',
                      'DeadOrganicNetChange(tC/ha)','MineralSoilsNetChange(tC/ha)','OrganicSoilsNetChange(tC/ha)',
                      'LivingBMGains(ktC)','LivingBMLosses(ktC)','LivingBMNetChange(ktC)','DeadOrganicNetChange(ktC)',
                      'MineralSoilsNetChange(ktC)','OrganicSoilsNetChange(ktC)','NetCO2(ktCO2)']
Table4C_columns_ls = Table4B_columns_ls
Table4D_columns_ls = Table4B_columns_ls
columns_lss = [Table4A_columns_ls,Table4B_columns_ls,Table4C_columns_ls,Table4D_columns_ls]


def CreateEUTable4Total(writer,data_dir,countryls:list,inv_start:int,inv_end:int):
    """
    Collect row 10 from CRFReporter Excel files Table4 A,B,C and D 
    \pre It is assumed that immediate subdirectory of data_dir contains country directories denoted by three letter acronym.
    For each country
       For each Table4.[A,B,C,D]
          For each each year
            Collect row 10 (Total) to dataframe from CRFReporter Excel files
          Create excel sheet
    Save Excel file
    \param writer  Excel writer
    \param data_dir Data location
    \param countryls List of countries
    \param inv_start Inventroy start year, 1990
    \param inv_end Inventory end year
    \return the Excle writer with data
    \post Units are as in Excel files (no conversion to CO2)
    """

    excelfilels=sorted(glob.glob(data_dir+'/'+country+'/'+country+'*.xlsx'))
    for (sheet,row_name,columns_ls) in zip(sheet_ls,substr_ls,columns_lss):
        print(country,sheet)
        data_row_lss=[]
        for excel_file in excelfilels:
            xlsx = pd.ExcelFile(excel_file)
            df = pd.read_excel(xlsx,sheet,keep_default_na=False,na_values=['MISSING_VALUE'])
            #Find row by its name as Dataframe
            row_df = df[df[df.columns[0]].str.contains(substr_ls[substr_ls.index(row_name)])==True]
            #Row as Series
            row_s = row_df.iloc[0,:]
            #Row as list
            row_ls = list(row_s)
            #Delete two first element: Title and Subdivision in CRFReporter excel
            del row_ls[0:2]
            #Append one year data
            data_row_lss.append(row_ls)
            #Data collected for one sheet, create data frame
        df = pd.DataFrame(data_row_lss)
        df.columns=columns_ls
        df.index=list(range(inv_start,(inv_end+1)))
        df.to_excel(writer,sheet_name=country+sheet,na_rep='NaN')
    return writer

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-d","--directory",dest="f1",required=True,help="Inventory Parties Directory")
    parser.add_argument("-s","--start",type=int,dest="f2",required=True,help="Inventory start year (1990)")
    parser.add_argument("-e","--end",type=int,dest="f3",required=True,help="Inventory end year")
    group=parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--eu",action="store_true",dest="eu",default=False,help="EU countries")
    group.add_argument("--euplus",action="store_true",dest="euplus",default=False,help="EU countries plus GBR, ISL, NOR")
    group.add_argument("-a","--all",action="store_true",dest="all",default=False,help="All countries (EU+others")
    group.add_argument("-c","--countries",dest="country",type=str,nargs='+',help="List of countries from the official acronyms separated by spaces")
    args = parser.parse_args()
    directory=args.f1
    print("Inventory Parties data directory:",directory)
    inventory_start=int(args.f2)
    print("Inventory start:",inventory_start)
    inventory_end=int(args.f3)
    print("Inventory end:",inventory_end)
    file_prefix = 'EU'
    countryls=[]
    if args.eu:
        print("Using EU  countries")
        countryls=euls
    elif args.euplus:
        print("Using EU  countries plus GBR, ISL and NOR")
        countryls=euplusls
        file_prefix = 'EU_GBR_ISL_NOR'
    elif args.all:
        print("Using all countries")
        countryls = allcountryls
        file_prefix='EU_and_Others'
    else:
        print("Using countries:", args.country) 
        countryls=args.country
        file_prefix=args.country[0]
        for country in args.country[1:]:
            file_prefix = file_prefix+"_"+country
    file_name = file_prefix+'_Restoration_'+str(inventory_start)+'_'+str(inventory_end)+'.xlsx'
    countryls = sorted(countryls)
    writer = pd.ExcelWriter(file_name,engine='xlsxwriter')
    writer =  CreateEUTable4Total(writer,directory,countryls,inventory_start,inventory_end)
    print("Writing results to:",file_name)
    writer.save()
    print("Done")
