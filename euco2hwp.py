import pandas as pd
import numpy as np
import math
import numbers
import glob
from optparse import OptionParser as OP

directory='EU-MS/2017'
inventory_start=1990
inventory_end=2015

#List of excel sheets needed
euls=['FIN','AUT','BEL','BGR','CYP','CZE','DEU',
      'DNK','ESP','EST','FRA','GBR','GRC','HRV',
      'HUN','IRL','ITA','LTU','LUX','LVA',
      'MLT','NLD','POL','PRT','ROU','SVK',
      'SVN','SWE']
noneuls=['CAN','CHE','EUA','ISL','JPN','LIE','MCO','NOR','RUS','TUR','USA']
countryls=euls+noneuls
sheetls = ['Table4.Gs1']
table4Gs1_sheet_name_ls=['Table4.Gs1 Total HWP','Table4.Gs1 Total HWP Domestic','Table4.Gs1 Total HWP Exported']
table4Gs1_row_ls=['TOTAL HWP','Total']

def CreateHWPExcelSheet(writer,directory,countryls,sheet,row_name_ls,col,sheet_name_ls,start,end):
    """Read CRFReporter Reporting table files (excel) for given EU countries
       for each inventory year. Find the given sheet and the given row (inventory item)
       and create a data frame row for each country for the CO2 net emission for each inventory year
       (last cell in the given row). This way one excel sheet is created including all EU countries.
       writer: excel writer that collects all Reporting tables into one excel file
       directory: directory for the countries (each country is a directory containing excel files) 
       countryls: list of (EU) countries
       sheet: the name of the excel sheet to be read
       row_name_ls: the row names to pick up in the sheet
       col: column index to data
       sheet_name_ls: sheet names (1.HWP Total, 2.HWP Domestic, 3.HW Exported) in the output excel file
       start: inventory start year
       end: inventory end year
    """
    data_row_ls0=[]
    data_row_ls1=[]
    data_row_ls2=[]
    approachA_set=set()
    for country in countryls:
        country=country.lower()
        #List all excel files and sort the files in ascending order (1990,1991,...,2015)
        excelfilels=glob.glob(directory+'/'+country+'*/*.xlsx')
        excelfilels=sorted(excelfilels)
        print(country.upper(),sheet_name_ls[0],sheet_name_ls[1],sheet_name_ls[2])
        row_ls0=[]
        row_ls1=[]
        row_ls2=[]
        i=start
        if excelfilels==[]:
            print("Missing country",country)
            row_ls0=['?']*len(list(range(start,(end+1))))
            row_ls1=['?']*len(list(range(start,(end+1))))
            row_ls2=['?']*len(list(range(start,(end+1))))
        for file in excelfilels:
            print(i)
            i=i+1
            xlsx = pd.ExcelFile(file)
            #The default set if missing values include NA,override default values and set
            #empty string ('') as the missing value
            df1 = pd.read_excel(xlsx,sheet,keep_default_na=False,na_values=[''])
            index = list(df1.columns)[0]
            #'TOTAL HWP' give two rows
            row0=df1[df1[index].str.contains(row_name_ls[0])==True]
            #'Total' also gives two rows, choose both
            row=df1[df1[index].str.contains(row_name_ls[1])==True]
            row_ls1.append(row.iloc[0,col])
            row_ls2.append(row.iloc[1,col])
            #To ease data processing in excel add up Domestic HWP and Exported HWP Totals
            #and insert the sum to row_ls0 instead of NaN from TOTAL HWP
            #The options are 1) notation key, 2) number or 3) missing value
            #Check first if both are string
            if isinstance(row.iloc[0,col],str) and isinstance(row.iloc[1,col],str):
                row_ls0.append(row.iloc[0,col]+','+row.iloc[1,col])
            #If either is a string the return the other one. It must be a number
            elif not isinstance(row.iloc[0,col],str) and  isinstance(row.iloc[1,col],str):
                row_ls0.append(row.iloc[0,col])
            elif isinstance(row.iloc[0,col],str) and not isinstance(row.iloc[1,col],str):
                row_ls0.append(row.iloc[1,col])
            #Now we can check against NaN (missing value), both must be numbers
            elif not np.isnan(row.iloc[0,col]) and not np.isnan(row.iloc[1,col]):
                row_ls0.append(row.iloc[0,col]+row.iloc[1,col])
            else:
                #Party (country) has reported total only
                #'TOTAL HWP' give two rows.
                #Test if the second one has NaN or not
                #If notnull is true, i.e. data exists, then the country has used Approach B
                if pd.notnull(row0.iloc[1,col]):
                    row_ls0.append(row0.iloc[1,col])
                #Otherwise this is Approach A
                else:
                    row_ls0.append(row0.iloc[0,col])
                    approachA_set.add(country.upper())
        data_row_ls0.append(row_ls0)
        data_row_ls1.append(row_ls1)
        data_row_ls2.append(row_ls2)
    data_row_ls0.append(sorted(list(approachA_set))+['']*(len(list(range(start,end+1)))-len(list(approachA_set))))
    df_total = pd.DataFrame(data_row_ls0)
    df_domestic = pd.DataFrame(data_row_ls1)
    df_export = pd.DataFrame(data_row_ls2)
    print(len(data_row_ls0),len(countryls))
    df_total.index=countryls+['Approach A']
    df_total.columns=list(range(start,end+1))
    df_total.to_excel(writer,sheet_name=sheet_name_ls[0],na_rep='NaN')
    df_domestic.index=countryls
    df_domestic.columns=list(range(start,end+1))
    df_domestic.to_excel(writer,sheet_name=sheet_name_ls[1],na_rep='NaN')
    df_export.index=countryls
    df_export.columns=list(range(start,end+1))
    df_export.to_excel(writer,sheet_name=sheet_name_ls[2],na_rep='NaN')

#writer = pd.ExcelWriter('Table4.Gs1.xlsx',
#                        engine='xlsxwriter')
#1. Table4G.s1
#CreateHWPExcelSheet(writer,euls,sheetls[0],table4Gs1_row_ls[0],5,table4Gs1_sheet_name_ls)
#writer.save()


if __name__ == "__main__":
    parser = OP()
    parser.add_option("-d","--directory",dest="f1",help="Inventory Parties Directory")
    parser.add_option("-s","--start",dest="f2",help="Inventory start year (1990)")
    parser.add_option("-e","--end",dest="f3",help="Inventory end year")
    parser.add_option("-a","--all",action="store_true",dest="f4",default=False,help="All countries")
    
    (options,args) = parser.parse_args()
    if options.f1 != None:
        directory=options.f1
    print("Inventory Parties directory",directory)
    if options.f2 != None:
        inventory_start=int(options.f2)
    print("Inventory start",inventory_start)
    if options.f3 != None:
        inventory_end=int(options.f3)
    print("Inventory end",inventory_end)
    if options.f4 == True:
        print("Using all countries")
        countryls=euls+noneuls
    else:
        print("Using EU countries")
        countryls=euls
        
    writer = pd.ExcelWriter('Table4.Gs1_with_ApproachA.xlsx',
                        engine='xlsxwriter')
    #1. Table4G.s1
    CreateHWPExcelSheet(writer,'EU-MS/2017/',countryls,sheetls[0],table4Gs1_row_ls,5,table4Gs1_sheet_name_ls,inventory_start,inventory_end)
    writer.save()
