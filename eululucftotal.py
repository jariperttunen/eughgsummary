import os
import argparse
import pathlib
import glob
import pandas as pd
import numpy as np
from countrylist import euls,euplusls,noneuls,allcountryls

#LULUCF Total table name
sheetls = ['Table4']
#Sheet names in resulting excel file
table4_sheet_name_ls=['4.Total LULUCF','A.Forest land','B.Cropland','C.Grassland',
                      'D.Wetlands','E.Settlements','F.Other land','G.HWP','H.Other']
#Search dataframe row with these names
table4_row_substr_ls=['4. Total','A. Forest','B. Cropland','C. Grassland','D. Wetlands',
                   'E. Settlements','F. Other land','G. Harvested', 'H. Other \(please specify\)']
#Dataframe columns for emissions
netco2_col=1
ch4_col=2
n2o_col=3

def sum_emissions(e1,e2):
    """Utility function to sum emissions. Emission can be a number or  anotation key"""
    try:#Both are numbers
        return float(e1)+float(e2)
    except:#The other one is not 
        try:#Check first e1
            return float(e1)
        except:
            try:#Check e2
                return float(e2)
            except:#Both are notation keys
                return e1+','+e2

def convertco2eq(e,gwp):
    """Convert emission 'e' to CO2eq, 'e' can be notation key"""
    try:
        return float(gwp*e)
    except:
        return e
            
def CreateLULUCFTotalSheet(writer,directory,countryls,sheet,row_name,result_sheet_name,start,end,ch4gwp,n2ogwp):
    """
    Read CRFReporter Reporting table files (excel) for given EU countries
    for each inventory year. Find the given sheet and the given row (inventory item)
    and create a data frame row for each country for the CO2 net emission for each inventory year
    (last cell in the given row). This way one excel sheet is created including all EU countries.
    writer: excel writer that collects all Reporting tables into one excel file
    directory: directory for the countries (each country is a directory containing excel files) 
    countryls: list of (EU) countries
    sheet: the name of the excel sheet to be read
    row_name: the row names to pick up in the sheet
    col: column index to data
    result_sheet_name: sheet name in the output excel file
    start: inventory start year
    end: inventory end year
    ch4gwp: CH4 global warming potential
    n2ogwp: N2O global warming potential
    """
    data_row_ls =[]
    #For each country
    for country in countryls:
        excelfilels=glob.glob(directory+'/'+country+'*/*.xlsx')
        #File name allows sorting by year 
        excelfilels=sorted(excelfilels)
        print(country.upper(),row_name)
        if excelfilels ==[]:
            print("Missing country", country)
        else:
            country_row_ls=[]
            i = start
            #For each country read GHG inventory files for each year and collect data
            for excel_file in excelfilels:
                xlsx = pd.ExcelFile(excel_file)
                #NA is numpy missing value, but a notation key in GHG, empty string is a true missing value 
                df = pd.read_excel(xlsx,sheet,keep_default_na=False,na_values=['MISSING_VALUE'])
                #Find the values with the right row and columns
                row=df[df[df.columns[0]].str.contains(row_name)==True]
                #co2
                co2_val = row.iloc[0,netco2_col]
                #ch4
                ch4_val = row.iloc[0,ch4_col]
                ch4_co2eq_val = convertco2eq(ch4_val,ch4gwp)
                #n2o
                n2o_val = row.iloc[0,n2o_col]
                n2o_co2eq_val = convertco2eq(n2o_val,n2ogwp)
                #sum co2+ch4_co2eq+n2o_co2eq
                co2eq_tot = sum_emissions(co2_val,sum_emissions(ch4_co2eq_val,n2o_co2eq_val))
                #add_result to row
                country_row_ls.append(co2eq_tot)
                print(i)
                i=i+1
            #add country row to data row list
            data_row_ls.append(country_row_ls)
    #Create dataframe from data row list
    df = pd.DataFrame(data_row_ls)
    df.index=countryls
    df.columns=list(range(start,(end+1)))
    #Add dataframe to excel. Again:make NaN as missing value, not NA
    df.to_excel(writer,sheet_name=result_sheet_name,na_rep='NaN')
     

 
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-d","--directory",dest="f1",required=True,help="Inventory Parties Directory")
    parser.add_argument("-s","--start",type=int,dest="f2",required=True,help="Inventory start year (usually 1990)")
    parser.add_argument("-e","--end",type=int,dest="f3",required=True,help="Inventory end year")
    group=parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--eu",action="store_true",dest="eu",default=False,help="EU countries")
    group.add_argument("--euplus",action="store_true",dest="euplus",default=False,help="EU countries plus GBR, ISL, NOR")
    group.add_argument("-a","--all",action="store_true",dest="all",default=False,help="All countries (EU+others")
    group.add_argument("-c","--countries",dest="country",type=str,nargs='+',help="List of countries from the official acronyms separated by spaces")
    parser.add_argument('--GWP',type=str,dest='gwp',default="AR4",help="Global warming potential, AR4 (GHG inventory, default) or AR5 (e.g. scenarios)")
    args = parser.parse_args()
    #AR4, GHG default
    ch4co2eq = 25 
    n2oco2eq = 298 
    if args.gwp == 'AR5':
        ch4co2eq = 28
        n2oco2eq = 265
    print("Using GWP:", ch4co2eq,n2oco2eq)
    directory=args.f1
    print("Inventory Parties directory",directory)
    inventory_start=int(args.f2)
    print("Inventory start",inventory_start)
    inventory_end=int(args.f3)
    print("Inventory end",inventory_end)
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
        print("Using countries", args.country) 
        countryls=args.country
        file_prefix=args.country[0]
        for country in args.country[1:]:
            file_prefix = file_prefix+"_"+country

    writer = pd.ExcelWriter(file_prefix+'_Table4TotalLULUCF_CO2eq_'+str(inventory_start)+'_'+str(inventory_end)+'.xlsx',
                            engine='xlsxwriter')
    for (row_name,sheet_name) in zip(table4_row_substr_ls,table4_sheet_name_ls):
        CreateLULUCFTotalSheet(writer,args.f1,countryls,sheetls[0],row_name,sheet_name,inventory_start,inventory_end,ch4co2eq,n2oco2eq)
    print("Writing file",file_prefix+'_Table4TotalLULUCF_CO2eq.xlsx')
    writer.save()
