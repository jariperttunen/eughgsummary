import euco2
import pandas as pd
import numpy as np
import math
import numbers
import glob
import xlsxwriter
import argparse

sheet_ls=['Summary2','Table4.B','Table4.C','Table3.D']
tableSummary2_row_ls=['3.  Agriculture','4. Land use, land-use change and forestry','Total CO2 equivalent emissions, including indirect CO2,  without land use, land-use change and forestry']							
tableSummary2_sheet_ls=['3.Agriculture Total ktCO2','4.LULUCF Total ktCO2','Total ktCO2 (Stat Fi L68)']
tableSummary2_row_ls2=['Total CO2 equivalent emissions without land use, land-use change and forestry','Total CO2 equivalent emissions with land use, land-use change and forestry','Total CO2 equivalent emissions, including indirect CO2,  with land use, land-use change and forestry']
tableSummary2_sheet_ls2=['Summary2 Total ktC L66','Summary2 Total ktC L67','Summary2 Total ktC L69']
table4B_row_ls=['B. Total Cropland']
table4B_sheet_ls=['4B CL Total Org ktCO2 (3.666)']
table4B_row_ls2=['B. Total Cropland']
table4B_sheet_ls2=['4B CL Total ktCO2']
table4C_row_ls=['C. Total grassland']
table4C_sheet_ls=['4C GL Total Org ktCO2 (3.666)']
table4C_row_ls2=['C. Total grassland']
table4C_sheet_ls2=['4C GL Total ktCO2']
table3D_row_ls=['6.   Cultivation of organic soils']
table3D_sheet_ls=['3D Histosols ktCO2 (298)']
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("-d","--directory",type=str,dest="f1",help="Inventory Parties Directory (e.g. EU-MS/2017)")
    parser.add_argument("-s","--start",type=int,dest="f2",help="Inventory start year (usually 1990)")
    parser.add_argument("-e","--end",type=int,dest="f3",help="Inventory end year (e.g. 2015)")
    parser.add_argument("-a","--all",action="store_true",dest="f4",default=False,help="All countries")
    args = parser.parse_args()
    start_year = args.f2
    end_year = args.f3
    writer = pd.ExcelWriter('EU_CO2_LULUCF_AGRICULTURE_TOTAL_'+str(start_year)+'_'+str(end_year)+'.xlsx',
                            engine='xlsxwriter')
    #Table Summary2
    writer2 = pd.ExcelWriter('EU_CO2_'+'LULUCF'+'_'+str(start_year)+'_'+str(end_year)+'.xlsx',
                            engine='xlsxwriter')
    #countryls=euco2.euls
    countryls=['FIN','AUT']
    directory=args.f1
    for (row_name,sheet_name) in zip(tableSummary2_row_ls,tableSummary2_sheet_ls):
        euco2.CreateExcelSheet(writer,writer2,directory,countryls,sheet_ls[0],row_name,9,sheet_name,
                               start_year,end_year,conv=1.0)
    #44/12
    for (row_name,sheet_name) in zip(table4B_row_ls,table4B_sheet_ls):
        euco2.CreateExcelSheet(writer,writer2,directory,countryls,sheet_ls[1],row_name,16,sheet_name,
                               start_year,end_year,conv=-44.0/12.0)
    for (row_name,sheet_name) in zip(table4B_row_ls2,table4B_sheet_ls2):
        euco2.CreateExcelSheet(writer,writer2,directory,countryls,sheet_ls[1],row_name,17,sheet_name,
                               start_year,end_year,conv=1.0)
    #44/12
    for (row_name,sheet_name) in zip(table4C_row_ls,table4C_sheet_ls):
        euco2.CreateExcelSheet(writer,writer2,directory,countryls,sheet_ls[2],row_name,16,sheet_name,
                               start_year,end_year,conv=-44.0/12.0)
    for (row_name,sheet_name) in zip(table4C_row_ls2,table4C_sheet_ls2):
        euco2.CreateExcelSheet(writer,writer2,directory,countryls,sheet_ls[2],row_name,17,sheet_name,
                               start_year,end_year,conv=1.0)
    #298
    for (row_name,sheet_name) in zip(table3D_row_ls,table3D_sheet_ls):
        euco2.CreateExcelSheet(writer,writer2,directory,countryls,sheet_ls[3],row_name,4,sheet_name,
                               start_year,end_year,conv=298.0)
    #L66, L67, L69
    for (row_name,sheet_name) in zip(tableSummary2_row_ls2,tableSummary2_sheet_ls2):
        euco2.CreateExcelSheet(writer,writer2,directory,countryls,sheet_ls[0],row_name,9,sheet_name,
                               start_year,end_year,conv=1.0)
    writer2.close()
    writer.close()
