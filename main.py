from autocorrect import Speller
import pandas as pd
from pandas import *

#Takes a list and autocorrects strings in list 
def list_auto(list):
    for x in range(len(list)):
        if type(list[x]) is str:
            word = str(spell(list[x]))
            if list[x] != word:
                list[x] = str(word)

#Higlights cell in excel sheet if string is misspelled 
def misspell_cell(val):
    if type(val) is str:
        color = 'yellow' if val != str(spell(val)) else 'white'
        return "background-color: %s" % color

if __name__ == "__main__":
    spell = Speller(lang='en') 
    writer = pd.ExcelWriter('testing_corrected.xlsx', engine='xlsxwriter')
    
    df = pd.read_excel('testing.xlsx')
    
    col_names = df.columns.tolist()         #get column names to list and autocorrect strings 
    list_auto(col_names)

    arr = df.to_numpy()

    new_df = pd.DataFrame(arr, columns = col_names)

    df_styled = new_df.style\
        .applymap(misspell_cell) 
    
    df_styled.to_excel(writer, sheet_name="Sheet1", startrow=1, index=False, header=False)    #write to excel
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    header_format = workbook.add_format()            #format for column 
    header_format.set_font_color('#000000')
    header_format.set_font_size(12)
    
    for col_num, value in enumerate(new_df.columns.values):       #write to excel 
        worksheet.write(0, col_num, value, header_format)
    
    writer.close()

