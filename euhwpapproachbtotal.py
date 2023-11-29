import os
import argparse
import pathlib
import glob
import pandas as pd
import numpy as np
from countrylist import euls,euplusls,noneuls,allcountryls

hwp_table = 'Table4.Gs1'
hwp_row = 'Total'
hwp_row_total = 'TOTAL HWP'
hwp_sheet_name_ls = ['Table4.Gs1_HWP_domestic_gains','Table4.Gs1_HWP_domestic_losses',
                     'Table4.Gs1_HWP_exported_gains','Table4.Gs1_HWP_exported_losses']

def color_orange(val):
    """
    Highlight in orange, to be used with row selection of countries
    with HWP Total only
    """
    return  'background-color: orange'

def EUHwpApproachBTotal(excel_writer,directory:str,countryls:list,sheet:str,
                        from_row:str,from_row_total:str,col_gains:int,col_losses:int,start:int,end:int):
    """
    Read CRFReporter Reporting table Table4.Gs1 excel sheets for Total HWP (tC) domestic and exported, gains and losses.
    \note The algorithm is similar to the one in *eulandtransitionmatrix.py*
    \param excel_writer Excel writer
    \param directory The direactory where the Reporting tables are located
    \param countryls List of countries to be used
    \param sheet Land Transition Matrix sheet name
    \param from_row Name of the row ('Total')
    \param col_gains Column number for HWP domestic gains
    \param col_losses Column number for HWP domestic losses
    \param start Inventory start year (1990)
    \param end Inventory end year
    """
    #Keep track of the countries that have HWP Total only
    hwp_total_only_set=set()
    domestic_gains_lss=[]
    domestic_losses_lss=[]
    exported_gains_lss=[]
    exported_losses_lss=[]
    for country in countryls:
        rowls1=[]
        rowls2=[]
        rowls3=[]
        rowls4=[]
        #List all excel files and sort the files in ascending order (1990,1991,...,2021)
        #Exclude years in 1980's that some countries report
        excelfilels=list(set(glob.glob(directory+'/'+country+'/*.xlsx'))-set(glob.glob(directory+'/'+country+'/*_198??*.xlsx')))
        excelfilels=sorted(excelfilels)
        print(country,sheet)
        for excel_file in excelfilels:
            xlsx = pd.ExcelFile(excel_file)
            df = pd.read_excel(xlsx,sheet,keep_default_na=False,na_values=[''])
            #from_row  'Total' should give two rows exactly
            rows2 = df[df[df.columns[0]].str.fullmatch(from_row)==True]
            #Domestic gains
            rowls1.append(rows2.iloc[0,col_gains])
            #Domestic losses
            rowls2.append(rows2.iloc[0,col_losses])
            #Exported gains
            rowls3.append(rows2.iloc[1,col_gains])
            #Exported losses
            rowls4.append(rows2.iloc[1,col_losses])
            if np.isnan(rows2.iloc[0,col_gains]):
                #Add country to total only 
                hwp_total_only_set.add(country)
                rows_tmp = df[df[df.columns[0]].str.contains(from_row_total)==True]
                rowls1.pop()
                rowls1.append(rows_tmp.iloc[1,col_gains])
                rowls2.pop()
                rowls2.append(rows_tmp.iloc[1,col_losses])
        domestic_gains_lss.append(rowls1)
        domestic_losses_lss.append(rowls2)
        exported_gains_lss.append(rowls3)
        exported_losses_lss.append(rowls4)
    df_domestic_gains = pd.DataFrame(domestic_gains_lss)
    df_domestic_losses = pd.DataFrame(domestic_losses_lss)
    df_exported_gains = pd.DataFrame(exported_gains_lss)
    df_exported_losses = pd.DataFrame(exported_losses_lss)
    df_domestic_gains.index = countryls
    df_domestic_gains.columns = list(range(start,end+1))
    df_domestic_losses.index = countryls
    df_domestic_losses.columns = list(range(start,end+1))
    df_exported_gains.index = countryls
    df_exported_gains.columns = list(range(start,end+1))
    df_exported_losses.index = countryls
    df_exported_losses.columns = list(range(start,end+1))
    #Coloring HWP Total only country rows
    countrlyls = sorted(list(hwp_total_only_set))
    print("COUNTRIES",countrlyls)
    idx = pd.IndexSlice
    print(idx)
    idx_slice = idx['AUT',:]
    df_domestic_gains = df_domestic_gains.style.apply(color_orange,subset=idx_slice,axis=1)
    df_domestic_gains.to_excel(excel_writer,hwp_sheet_name_ls[0],na_rep='NaN')
    df_domestic_losses.to_excel(excel_writer,hwp_sheet_name_ls[1],na_rep='NaN')
    df_exported_gains.to_excel(excel_writer,hwp_sheet_name_ls[2],na_rep='NaN')
    df_exported_losses.to_excel(excel_writer,hwp_sheet_name_ls[3],na_rep='NaN')

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

    excel_writer = pd.ExcelWriter(file_prefix+'_Table4.Gs1_HWPApproachBTotal_'+str(inventory_start)+'_'+str(inventory_end)+'.xlsx',engine='xlsxwriter')
    EUHwpApproachBTotal(excel_writer,directory,countryls,hwp_table,hwp_row,hwp_row_total,1,2,inventory_start,inventory_end)
    excel_writer.close()
