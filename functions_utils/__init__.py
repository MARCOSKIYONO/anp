# Needs Libraries
import os, sys
import win32com.client
import pandas as pd
from datetime import date, datetime
import string
import urllib.request


def download_file(download_url, filename):
    try:
        response = urllib.request.urlopen(download_url)    
        file = open(filename , 'wb')
        file.write(response.read())
        file.close()
    except Exception as e:
        print("Program download_file error:{}".format(e))        


def get_subject_positions(df, pivot_name):
    try:
        if pivot_name:
            df_new= df[ (df['PT Name']==pivot_name) & (df['Source Name'] == 'Jan') ]
            first_space, first_letter, inicial_number = df_new.iloc[0]['Heading'].strip().split('$')

            first_space , first_letter, first_number, last_letter, last_number = df_new.iloc[0]['PT Address'].strip().split('$')        
            
            return inicial_number, first_letter, last_number, last_letter
    except Exception as e:
        print("Program get_subject_positions error:{}".format(e))    


def load_list(inicial_number, first_letter, last_number, last_letter,excel, sheet, prefix):
    try:
        import string
        abc=string.ascii_uppercase

        lista =[]

        year_number = int(inicial_number) - 1
        for letter in abc:
            if letter > first_letter and letter <= last_letter:
                cell_year = letter+ str(year_number)
                item = {"sheet": prefix + str(int(excel.Worksheets(sheet).Range(cell_year).Value)), "range":letter+ str(inicial_number) }
                lista.append(item)
        return lista
    except Exception as e:
        print("Program load_list error:{}".format(e))    


def tranform_df(df, id_vars, drop_columns):
    try:
        df = df.drop_duplicates()
        
        df_temp= df.melt(id_vars=id_vars, var_name= "month", value_name="volume")
        
        df_temp = df_temp[df_temp['month'] != 'TOTAL']
        
        df_temp["volume"] = df_temp["volume"].fillna(0)
        
        df_temp["uf"] =df_temp["ESTADO"].map(Ufs)
        
        df_temp["year_month"] = df_temp["ANO"].astype(str) + df_temp["month"].map(Months)
        
        df_temp["created_at"] = now
        
        df_temp = df_temp.drop(drop_columns, axis=1)
        
        df_temp = df_temp[['year_month','uf','COMBUSTÍVEL','UNIDADE','volume']]
        
        df_temp.rename(columns={'COMBUSTÍVEL': 'product', 'UNIDADE': 'unit'}, inplace=True)
        
        return df_temp
        
    except Exception as e:
        print("Program tranform_df error:{}".format(e))      


def load_total(inicial_number, first_letter, last_number, last_letter,excel, sheet, prefix):
    try:
        import string
        abc=string.ascii_uppercase

        lista =[]

        year_number = int(inicial_number) - 1
        for letter in abc:
            if letter > first_letter and letter <= last_letter:
                cell_year = letter+ str(year_number)
                cell_value= letter+ str(last_number)
                item = {"year": int(excel.Worksheets(sheet).Range(cell_year).Value), "volume": float(excel.Worksheets(sheet).Range(cell_value).Value)}
                lista.append(item)
        return lista
    except Exception as e:
        print("Program load_Total error:{}".format(e))            
        

def create_html(html_file_name, df_sales_diesel, df_sales_general):
    try:
        # Generate HTML result
        with open(html_file_name, 'w') as _file:
            _file.write('<center>' 
                        +'<h1> ANP - Summary Report </h1><br><hr>'
                        +'<h2> Sales of diesel by UF and type </h2>' + df_sales_diesel.to_html(index=False,border=1,justify="center") + '<br><hr>'
                        +'<h2> Sales of oil derivative fuels by UF and product </h2>' + df_sales_general.to_html(index=False,border=1,justify="center") + '<be><hr>'
                        +'<h1> The biggest difference founded in the volume of <font color="red">405.399 </font> in the report "Sales of oil derivative fuels by UF and product" in 2020 is due to the difference in Excel itself, between the source data and the pivot table, the data calculated in the program keeps the same difference, the others volumes the difference is in the 7th decimal place </h1><br><hr>'
                        +'</center>')    
    except Exception as e:
        print("Program create_html error:{}".format(e)) 

