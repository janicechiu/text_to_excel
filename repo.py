# top 1000 TI Parts Searched/Clicked - FC/OM: come from topPart_Texas_FindChips/OEMsTrade_Click/Search
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

x = DataFrame.from_csv("topPart_Texas_FindChips_Click_"+month+".tsv",sep='\t')

print (x)


writer = pd.ExcelWriter("1 - Top 1000 TI Parts Clicked - FC "+month+".xlsx", engine ='xlsxwriter')


x.to_excel(writer, sheet_name ='Sheet1', startrow = 0, startcol = 0, index =False)

    

