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
hwp_sheet_name = 'HWP Table4.Gs1'
hwp_columns_ls = ["Country","Year","Gains(tC)","Gains(CO2)","TG/DG/EG","Losses(tC)","Losses(CO2)","TL/DL/EL"]
def color_orange(val):
    """
    Highlight in orange, to be used with row selection of countries
    with HWP Total only
    """
    return  'background-color: orange'

def convert_co2(x):
    try:
        x = 44.0/12.0*x
        return x
    except:
        return x
    
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
    #Resulting data frame where the results are in columnwise order
    #in one excel sheet.
    df_result_tC = pd.DataFrame()
    for country in countryls:
        row_country_ls=[]
        row_year_ls=[]
        #Yhe lists for types of gains and losses 
        row_tot_gains_type_ls=[]
        row_tot_losses_type_ls=[]
        row_domestic_gains_type_ls=[]
        row_domestic_losses_type_ls=[]
        row_exported_gains_type_ls=[]
        row_exported_losses_type_ls=[]
        #The lists for gains and losses
        row_tot_gains_tC_ls=[]
        row_tot_losses_tC_ls=[]
        row_domestic_gains_tC_ls=[]
        row_domestic_losses_tC_ls=[]
        row_exported_gains_tC_ls=[]
        row_exported_losses_tC_ls=[]
        #List all excel files and sort the files in ascending order (1990,1991,...,2021)
        #Exclude years in 1980's that some countries report
        excelfilels=list(set(glob.glob(directory+'/'+country+'/*.xlsx'))-set(glob.glob(directory+'/'+country+'/*_198??*.xlsx')))
        excelfilels=sorted(excelfilels)
        print(country,sheet)
        year = start
        #Some countries report HWP tot only
        is_hwptot_only = False
        for excel_file in excelfilels:
            row_country_ls.append(country)
            row_year_ls.append(year)
            year = year+1
            #xlsx_file = pd.ExcelFile(excel_file)
            df = pd.read_excel(excel_file,sheet,keep_default_na=False,na_values=[''])
            #from_row  'Total' should give two rows exactly
            rows1 = df[df[df.columns[0]].str.fullmatch(from_row)==True]
            if np.isnan(rows1.iloc[0,col_gains]):
                #Country has HWP total only
                rows1 = df[df[df.columns[0]].str.contains(from_row_total)==True]
                row_tot_gains_type_ls.append('TG')
                row_tot_losses_type_ls.append('TL')
                row_tot_gains_tC_ls.append(rows1.iloc[1,col_gains])
                row_tot_losses_tC_ls.append(rows1.iloc[1,col_losses])
                is_hwptot_only = True
            else:
                row_domestic_gains_type_ls.append('DG')
                row_domestic_losses_type_ls.append('DL')
                row_exported_gains_type_ls.append('EG')
                row_exported_losses_type_ls.append('EL')
                #Domestic gains
                row_domestic_gains_tC_ls.append(rows1.iloc[0,col_gains])
                #Domestic losses
                row_domestic_losses_tC_ls.append(rows1.iloc[0,col_losses])
                #Exported gains
                row_exported_gains_tC_ls.append(rows1.iloc[1,col_gains])
                #Exported losses
                row_exported_losses_tC_ls.append(rows1.iloc[1,col_losses])
        #One country is done. Collect the results in a list (list of lists), make a data frame *and transpose it*.
        #Append to the final data frame. The results are now in columnwise order
        country_lss1=[]
        country_lss2=[]
        if is_hwptot_only:
            row_tot_gains_CO2_ls = [convert_co2(x) for x in row_tot_gains_tC_ls]
            row_tot_losses_CO2_ls = [convert_co2(x) for x in row_tot_losses_tC_ls]
            country_lss1 = [row_country_ls,row_year_ls,row_tot_gains_tC_ls,row_tot_gains_CO2_ls ,row_tot_gains_type_ls,
                            row_tot_losses_tC_ls,row_tot_losses_CO2_ls,row_tot_losses_type_ls]
        else:
            row_domestic_gains_CO2_ls = [convert_co2(x) for x in row_domestic_gains_tC_ls]
            row_domestic_losses_CO2_ls = [convert_co2(x) for x in row_domestic_losses_tC_ls]
            row_exported_gains_CO2_ls = [convert_co2(x) for x in row_exported_gains_tC_ls]
            row_exported_losses_CO2_ls = [convert_co2(x) for x in row_exported_losses_tC_ls]
            country_lss1 = [row_country_ls,row_year_ls,
                            row_domestic_gains_tC_ls,row_domestic_gains_CO2_ls,row_domestic_gains_type_ls,
                            row_domestic_losses_tC_ls,row_domestic_losses_CO2_ls,row_domestic_losses_type_ls]
            country_lss2 = [row_country_ls,row_year_ls,
                            row_exported_gains_tC_ls,row_exported_gains_CO2_ls,row_exported_gains_type_ls,
                            row_exported_losses_tC_ls,row_exported_losses_CO2_ls,row_exported_losses_type_ls]
        
        #Create dataframe for one country
        df1_country = pd.DataFrame(country_lss1)
        df2_country = pd.DataFrame(country_lss2)
        #Transpose to get columnwise results
        df1T_country = df1_country.transpose()
        df2T_country = df2_country.transpose()
        #Append to final dataframe
        df_result_tC = pd.concat([df_result_tC,df1T_country,df2T_country])
    #All countries collected, insert dataframe to excel
    df_result_tC.columns = hwp_columns_ls 
    df_result_tC.to_excel(excel_writer,hwp_sheet_name,na_rep='NaN',engine='xlsxwriter')

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
