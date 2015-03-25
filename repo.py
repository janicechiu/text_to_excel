# 1 & 2: Top 1000 TI Parts Searched/Clicked - FC/OM: come from topPart_Texas_FindChips/OEMsTrade_Click/Search
# 3: Top 50 Competitor Part - FC/OM: come from FindChips_competitor_report & OEMsTrade_competitor_report

from pandas import DataFrame, read_csv, read_excel, Series
import pandas as pd
import csv
from xlsxwriter.workbook import Workbook

month ='2015-02'
with open("topPart_Texas_FindChips_Search_"+month+".tsv",'r') as tsvin :
    with open("1 - Top 1000 TI Parts Searched - FC "+month+".csv", 'w') as csvout:
        tsvin = csv.reader(tsvin, dialect =csv.excel_tab)
        csvout = csv.writer(csvout, delimiter =',')
        for tmp in tsvin:
            csvout.writerows([tmp])

############# for report 1 #####################

x = pd.read_csv("topPart_Texas_FindChips_Search_"+month+".tsv",sep='\t')
writer = pd.ExcelWriter("1 - Top 1000 TI Parts Searched - FC "+month+".xlsx", engine ='xlsxwriter')
x.to_excel(writer, sheet_name ='Sheet1', startrow = 0, startcol = 0, index =False)


x = pd.read_csv("topPart_Texas_OEMsTrade_Click_"+month+".tsv",sep='\t')
writer = pd.ExcelWriter("2 - Top 1000 TI Parts Clicked - OM "+month+".xlsx", engine ='xlsxwriter')
x.to_excel(writer, sheet_name ='Sheet1', startrow = 0, startcol = 0, index =False)
    
x = pd.read_csv("topPart_Texas_OEMsTrade_Search_"+month+".tsv",sep='\t')
writer = pd.ExcelWriter("2 - Top 1000 TI Parts Searched - OM "+month+".xlsx", engine ='xlsxwriter')
x.to_excel(writer, sheet_name ='Sheet1', startrow = 0, startcol = 0, index =False)


################ for report 3 ##################################################
x = pd.read_csv("FindChips_competitor_report_"+month+".tsv",sep='\t')
writer = pd.ExcelWriter("3 - Top 50 Competitor Parts - FC "+month+".xlsx", engine ='xlsxwriter')
x.to_excel(writer, sheet_name ='Sheet1', startrow = 0, startcol = 0, index =False)

x = pd.read_csv("OEMsTrade_competitor_report_"+month+".tsv",sep='\t')
writer = pd.ExcelWriter("3 - Top 50 Competitor Parts - OM "+month+".xlsx", engine ='xlsxwriter')
x.to_excel(writer, sheet_name ='Sheet1', startrow = 0, startcol = 0, index =False)

################# for report 4 ##################################################

x = pd.read_csv("FindChips_mfrclicksbycategory_"+month+".tsv",sep='\t')
writer = pd.ExcelWriter("4 - Manufacturer Clicks by Category - FC "+month+".xlsx", engine ='xlsxwriter')
x.to_excel(writer, sheet_name ='Sheet1', startrow = 0, startcol = 0, index =False)

x = pd.read_csv("OEMsTrade_mfrclicksbycategory_"+month+".tsv",sep='\t')
writer = pd.ExcelWriter("4 - Manufacturer Clicks by Category - OM "+month+".xlsx", engine ='xlsxwriter')
x.to_excel(writer, sheet_name ='Sheet1', startrow = 0, startcol = 0, index =False)

##################for report 2 & 5 FC ####################################################

dist =['analog','atmel','freescale','integrated','linear','maxim','nxp','on','rohm','silicon','stmicroelectronics']
writer = pd.ExcelWriter("5 - Top 1000 OPN - FC "+month+".xlsx", engine ='xlsxwriter')
for i in dist:
    x = pd.read_csv("topPart_"+i+"_FindChips_"+month+".tsv",sep='\t')
    x.to_excel(writer, sheet_name = i, startrow = 0, startcol = 0, index =False)
    worksheet = writer.sheets[i]
    worksheet.set_column('A:E',20)
    

x = pd.read_csv("topPart_Texas_FindChips_Click_"+month+".tsv",sep='\t')
x.iloc[:,[0,1,2,3,5]].to_excel(writer, sheet_name = 'TI', startrow = 0, startcol = 0, index =False)

writer = pd.ExcelWriter("1 - Top 1000 TI Parts Clicked - FC "+month+".xlsx", engine ='xlsxwriter')
x.to_excel(writer, sheet_name ='Sheet1', startrow = 0, startcol = 0, index =False)

##################for report 2 & 5 OM ####################################################

writer = pd.ExcelWriter("5 - Top 1000 OPN - OM "+month+".xlsx", engine ='xlsxwriter')
for i in dist:
    x = pd.read_csv("topPart_"+i+"_OEMsTrade_"+month+".tsv",sep='\t')
    x.to_excel(writer, sheet_name = i, startrow = 0, startcol = 0, index =False)
    worksheet = writer.sheets[i]
    worksheet.set_column('A:E',20)


x = pd.read_csv("topPart_Texas_OEMsTrade_Click_"+month+".tsv",sep='\t')
x.iloc[:,[0,1,2,3,5]].to_excel(writer, sheet_name = 'TI', startrow = 0, startcol = 0, index =False)

writer = pd.ExcelWriter("1 - Top 1000 TI Parts Clicked - OM "+month+".xlsx", engine ='xlsxwriter')
x.to_excel(writer, sheet_name ='Sheet1', startrow = 0, startcol = 0, index =False)
















