import pandas as pd
import numpy as np
import math
import numbers
import glob
import xlsxwriter
import argparse

directory='EU-MS/2017'
inventory_start=1990
inventory_end=2015
#List of excel sheets needed
sheetls = ['Table4.A','Table4.B','Table4.C',
           'Table4.D','Table4.E','Table4.F']
euls=['FIN','AUT','BEL','BGR','CYP','CZE','DEU',
      'DNK','ESP','EST','FRA','GBR','GRC','HRV',
      'HUN','IRL','ITA','LTU','LUX','LVA',
      'MLT','NLD','POL','PRT','ROU','SVK',
      'SVN','SWE']
noneuls=['CAN','CHE','EUA','ISL','JPN','LIE','MCO','NOR','RUS','TUR','USA']
countryls=euls+noneuls

table4A_total_FL_ls = ['A. Total forest land']
table4A_total_FL_sheet_name_ls = ['4A Total FL']

table4A_row_ls = ['1. Forest','2.1 Cropland','2.2 Grassland','2.3 Wetlands',
                  '2.4 Settlements','2.5 Other']
table4A_sheet_name_ls = ['4A1 FL remaining FL','4A2.1 CL to FL','4A2.2 GL to FL',
                         '4A2.3 WL to FL','4A2.4 SL to FL','4A2.5 OL to FL']

table4B_row_ls = ['1. Cropland','2.1 Forest','2.2 Grassland','2.3 Wetlands',
                  '2.4 Settlements','2.5 Other']
table4B_sheet_name_ls = ['4B1 CL remaining CL','4B2.1 FL to CL','4B2.2 GL to CL',
                         '4B2.3 WL to CL','4B2.4 SL to CL','4B2.5 OL to CL']
    
table4C_row_ls = ['1. Grassland','2.1 Forest','2.2 Cropland','2.3 Wetlands',
                  '2.4 Settlements','2.5 Other']
table4C_sheet_name_ls = ['4C1 GL remaining GL','4C2.1 FL to GL','4C2.2 CL to GL',
                         '4C2.3 WL to GL','4C2.4 SL to GL','4C2.5 OL to GL']
    
table4D_row_ls = ['1. Wetlands','2.1 Land converted to peat',
                  '2.2 Land converted to flooded','2.3 Land converted to other']
table4D_sheet_name_ls = ['4D1 WL remaining WL','4D2.1 Land to Peat extraction',
                         '4D2.2 Land to flooded land','4D2.3 Land to other WL']
    
table4E_row_ls = ['1. Settlements','2.1 Forest','2.2 Cropland','2.3 Grassland',
                  '2.4 Wetlands','2.5 Other']
table4E_sheet_name_ls = ['4E1 SL remaining SL','4E2.1 FL to SL','4E2.2 CL to SL',
                         '4E2.3 GL to SL','4E2.4 WL to SL','4E2.5 OL to SL']

table4F_row_ls = ['1. Other','2.1 Forest','2.2 Cropland','2.3 Grassland',
                  '2.4 Wetlands','2.5 Settlements']
table4F_sheet_name_ls = ['4F1 OL remaining OL','4F2.1 FL to OL','4F2.2 CL to OL',
                         '4F2.3 GL to OL','4F2.4 WL to OL','4F2.5 SL to OL']


def CreateExcelSheet(writer,writer2,directory,countryls,sheet,row_name,col,sheet_name,start,end,conv=1.0):
    """Read CRFReporter Reporting table files (excel) for given EU countries
       for each inventory year. Find the given sheet and the given row (inventory item)
       and create a data frame row for each country for the CO2 net emission for each inventory year
       (last cell in the given row). This way one excel sheet is created including all EU countries.
       writer: excel writer that collects all Reporting tables into one excel file
       writer2: excel writer that writes one Reporting table
       countryls: list of EU countries
       sheet: the name of the excel sheet to be read
       row_name: the name of the row in the sheet
       col: column index to data
       sheet_name: sheet name in the output excel file
    """ 
    data_row_ls=[]
    for country in countryls:
        print(country,row_name,sheet_name)
        country=country.lower()
        #Move dirfilels as parameter or implement more generic way to retrieve
        #excel files for one country.
        dirfilels=glob.glob(directory+'/'+country+'*/*.xlsx')
        dirfilels=sorted(dirfilels)
        row_ls=[]
        i=inventory_start
        if len(dirfilels)==0:
            row_ls=['?']*len(list(range(start,(end+1))))
            print('Missing country',country)
        for file in dirfilels:
            xlsx = pd.ExcelFile(file)
            #The default set if missing values include NA,override default values and set
            #empty string ('') as the missing value
            df = pd.read_excel(xlsx,sheet,keep_default_na=False,na_values=[''])
            index = list(df.columns)[0]
            row=df[df[index].str.contains(row_name)==True]
            number=0
            try:
                number = float(conv*row.iloc[0,col])
            except:
                number=row.iloc[0,col]
            row_ls.append(number)
            print(i,number)
            i=i+1
        if len(row_ls)!=len(list(range(start,(end+1)))):
            print(country,len(row_ls))
        data_row_ls.append(row_ls)
    df1 = pd.DataFrame(data_row_ls)
    df1.index=countryls
    df1.columns=list(range(start,(end+1)))
    df1.to_excel(writer,sheet_name=sheet_name,na_rep='NaN')
    df1.to_excel(writer2,sheet_name=sheet_name,na_rep='NaN')

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-d","--directory",dest="f1",help="Inventory Parties Directory (default EU-MH/2017)")
    parser.add_argument("-s","--start",dest="f2",help="Inventory start year (usually 1990)")
    parser.add_argument("-e","--end",dest="f3",help="Inventory end year (default 2015")
    parser.add_argument("-a","--all",action="store_true",dest="f4",default=False,help="All countries (EU+others")
    
    args = parser.parse_args()
    if args.f1 != None:
        directory=args.f1
    print("Inventory Parties directory",directory)
    if args.f2 != None:
        inventory_start=int(args.f2)
    print("Inventory start",inventory_start)
    if args.f3 != None:
        inventory_end=int(args.f3)
    print("Inventory end",inventory_end)
    if args.f4 == True:
        print("Using all countries")
        countryls=euls+noneuls
    else:
        print("Using EU countries")
        countryls=euls
        
    writer = pd.ExcelWriter('EU_NetCO2_emissions_removals_1990_2015_submitted.xlsx',
                            engine='xlsxwriter')
    #Table 4 A Total forest land
    writer2 = pd.ExcelWriter('EU_NetCO2_'+sheetls[0]+'Total_FL_1990_2015.xlsx',
                            engine='xlsxwriter')
    for (row_name,sheet_name) in zip(table4A_total_FL_ls,table4A_total_FL_sheet_name_ls):
        CreateExcelSheet(writer,writer2,directory,countryls,sheetls[0],row_name,19,sheet_name,inventory_start,inventory_end)
    writer2.save()
    #1. Table 4.A
    writer2 = pd.ExcelWriter('EU_NetCO2_'+sheetls[0]+'1990_2015.xlsx',
                            engine='xlsxwriter')
    for (row_name,sheet_name) in zip(table4A_row_ls,table4A_sheet_name_ls):
        CreateExcelSheet(writer,writer2,directory,countryls,sheetls[0],row_name,19,sheet_name,inventory_start,inventory_end)
    writer2.save()
    #2. Table 4.B
    writer2 = pd.ExcelWriter('EU_NetCO2_'+sheetls[1]+'.xlsx',
                            engine='xlsxwriter')
    for (row_name,sheet_name) in zip(table4B_row_ls,table4B_sheet_name_ls):
        CreateExcelSheet(writer,writer2,directory,countryls,sheetls[1],row_name,17,sheet_name,inventory_start,inventory_end)
    writer2.save()
    #3. Table 4.C
    writer2 = pd.ExcelWriter('EU_NetCO2_'+sheetls[2]+'.xlsx',
                            engine='xlsxwriter')
    for (row_name,sheet_name) in zip(table4C_row_ls,table4C_sheet_name_ls):
        CreateExcelSheet(writer,writer2,directory,countryls,sheetls[2],row_name,17,sheet_name,inventory_start,inventory_end)
    writer2.save()

    #4. Table 4.D
    writer2 = pd.ExcelWriter('EU_NetCO2_'+sheetls[3]+'.xlsx',
                            engine='xlsxwriter')
    for (row_name,sheet_name) in zip(table4D_row_ls,table4D_sheet_name_ls):
        CreateExcelSheet(writer,writer2,directory,countryls,sheetls[3],row_name,17,sheet_name,inventory_start,inventory_end)
    writer2.save()
    #5. Table 4.E
    writer2 = pd.ExcelWriter('EU_NetCO2_'+sheetls[4]+'.xlsx',
                            engine='xlsxwriter')
    for (row_name,sheet_name) in zip(table4E_row_ls,table4E_sheet_name_ls):
        CreateExcelSheet(writer,writer2,directory,countryls,sheetls[4],row_name,17,sheet_name,inventory_start,inventory_end)
    writer2.save()
    #6. Table 4.F
    writer2 = pd.ExcelWriter('EU_NetCO2_'+sheetls[5]+'.xlsx',
                            engine='xlsxwriter')
    for (row_name,sheet_name) in zip(table4F_row_ls,table4F_sheet_name_ls):
        CreateExcelSheet(writer,writer2,directory,countryls,sheetls[5],row_name,17,sheet_name,inventory_start,inventory_end)
    writer2.save()
    writer.save()


