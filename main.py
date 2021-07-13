# LIBRARIES
import os, sys
import win32com.client
import pandas as pd
from datetime import date, datetime
import string
import urllib.request
import webbrowser

from functions_utils import load_Total, tranform_df, load_list, get_subject_positions, download_file
from utils import Months, Ufs, Sub_Detail_Pivot_tables, sub_expand_sales

# CONFIGURATIONS
#pd.set_option('precision', 9) # Trying to find out a way equalize the resuls when the difference it´s in the above the 7# decimal precision, but didn´t work
pd.options.display.float_format = '{: .12f}'.format
pd.set_option('mode.chained_assignment', None)

Always_Download = True # If need always download file before execute set True else False

Debug_Visible_Excel = True # If need see excel file while the program executes set True, else set False

# LOCAL VARIABLES
donwload_url = "http://www.anp.gov.br/arquivos/dados-estatisticos/vendas-combustiveis/vendas-combustiveis-m3.xls"

path = os.path.abspath(os.getcwd())

filename= "vendas-combustiveis-m3.xls"

html_file_name = "anp_sales.html"

Complete_File_Name = "{}\{}".format(path,filename)

Module_General_Name = 'Modulo_General'
Main_Sheet_Name = 'Plan1'
Index_Sheet_name = 'Index'

Prefix_Sheet_Sales_General_Name = 'Sales_General_'
Prefix_Sheet_Sales_Diesel_Name = 'Sales_Diesel_'

Index_Sub_Name = 'Get_Pivot_Details'
Sales_General_Sub_Name = 'Get_Sales_General'
Sales_Diesel_Sub_Name = 'Get_Sales_Diesel'

Pivot_table_Sales_General_Name = 'Tabela dinâmica1'
Pivot_table_Sales_Diesel_Name = 'Tabela dinâmica9'

now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

Sheets_List = [Index_Sheet_name,Prefix_Sheet_Sales_General_Name,Prefix_Sheet_Sales_Diesel_Name]

def increase_excel_macros():
    try:    
        # open up an instance of Excel with the win32com driver
        excel = win32com.client.Dispatch("Excel.Application")

        # do the operation in background without actually opening Excel
        excel.Visible = Debug_Visible_Excel

        # open the excel workbook from the specified file
        workbook = excel.Workbooks.Open(Filename=Complete_File_Name)

        excel.DisplayAlerts = False
        workbook.DoNotPromptForConvert = True
        workbook.CheckCompatibility = False

        # Remove Modules
        if Always_Download == False:
            for i in workbook.VBProject.VBComponents:        
                xlmodule = workbook.VBProject.VBComponents(i.Name)
                if xlmodule.Type in [1, 2, 3] and Module_General_Name in i.Name:       
                    print("Removing module {} ...".format(i.Name))
                    workbook.VBProject.VBComponents.Remove(xlmodule)

        # Remove Sheets
        if Always_Download == False:
            for sh in workbook.Sheets:
                Find=False
                for string in Sheets_List:
                    if string in sh.Name :
                        Find = True
                        continue
                if Find == True:
                    print("Removing sheet {} ...".format(sh.Name))
                    excel.Worksheets(sh.Name).Delete() 

        # Add General Module 
        excelModule = workbook.VBProject.VBComponents.Add(1)
        excelModule.name = Module_General_Name

        # Add Macro Get Details Pivot Table
        excelModule.CodeModule.AddFromString(Sub_Detail_Pivot_tables.format(Index_Sub_Name,Index_Sheet_name) )    

        # Add Macro Expand General Sales and Diesel Sales
        excelModule.CodeModule.AddFromString(sub_expand_sales.format(subname=Sales_General_Sub_Name,  mainsheet=Main_Sheet_Name) )

        # run the macro
        excel.Application.Run(Index_Sub_Name)

        # save the workbook and close
        excel.Workbooks(1).Close(SaveChanges=1)
        excel.Application.Quit()

        # garbage collection
        del excel

    except Exception as e:
        print(e)    
        

def get_anp_file():
    try:   
        # get directory where the script is located and Check File Exists
        Current_Path = os.path.abspath(os.getcwd())
        print(Current_Path)
        filepath= "{}\{}".format(Current_Path,filename)

        if os.path.exists(Complete_File_Name):
            print("File Exists !")
            if Always_Download:
                download_file(donwload_url, Complete_File_Name)
        else:
            print("File Do Not Exists !")
            download_file(donwload_url, Complete_File_Name)
    except Exception as e:
        print(e)     


def main():
    try:   
        # get ANP File
        get_anp_file()
    
        # Increase Module and Macros Excel
        increase_excel_macros()
        
        # Loading Pivot Table Info
        df = pd.read_excel(Complete_File_Name, sheet_name=Index_Sheet_name)
        
        # Expand Source Data from Pivot Tables
        try:
            # Open Excel
            # open up an instance of Excel with the win32com driver
            excel = win32com.client.Dispatch("Excel.Application")

            # do the operation in background without actually opening Excel
            excel.Visible = Debug_Visible_Excel

            # open the excel workbook from the specified file
            workbook = excel.Workbooks.Open(Filename=Complete_File_Name)
                    
            # Get Sales General sheets positions
            inicial_number, first_letter, last_number, last_letter = get_subject_positions(df,Pivot_table_Sales_General_Name)
            List_Sales_General = load_list(inicial_number, first_letter, last_number, last_letter , excel,Main_Sheet_Name, Prefix_Sheet_Sales_General_Name)    
            df_total_sales_general = pd.DataFrame(load_Total(inicial_number, first_letter, last_number, last_letter , excel,Main_Sheet_Name, Prefix_Sheet_Sales_General_Name))

            # Get Sales Diesel sheets positions
            inicial_number, first_letter, last_number, last_letter = get_subject_positions(df,Pivot_table_Sales_Diesel_Name)
            List_Sales_Diesel = load_list(inicial_number, first_letter, last_number, last_letter , excel,Main_Sheet_Name, Prefix_Sheet_Sales_Diesel_Name)
            df_total_sales_diesel = pd.DataFrame(load_Total(inicial_number, first_letter, last_number, last_letter , excel,Main_Sheet_Name, Prefix_Sheet_Sales_General_Name))

            # Generate details Sales General infos
            for item in List_Sales_General:
                print(item['sheet'], item['range'])
                excel.Application.Run(Sales_General_Sub_Name,item['sheet'] , item['range'])

            # Generate details Sales Diesel infos
            for item in List_Sales_Diesel:
                print(item['sheet'], item['range'])
                excel.Application.Run(Sales_General_Sub_Name,item['sheet'] , item['range'])
                
            # save the workbook and close
            excel.Workbooks(1).Close(SaveChanges=1)
            excel.Application.Quit()

        except Exception as e:
            excel.Workbooks(1).Close(SaveChanges=0)
            excel.Application.Quit()        
            print(e)       
        
        # Load Data frame Sales General infos
        df_Sales_General = pd.DataFrame()
        for item in List_Sales_General:
            print(Complete_File_Name, item['sheet'], item['range'])
            df = pd.read_excel(Complete_File_Name,
                           sheet_name = item['sheet'])
            df_Sales_General = pd.concat([df_Sales_General,df])

        # Load Data frame Sales Diesel infos
        df_Sales_Diesel = pd.DataFrame()
        for item in List_Sales_Diesel:
            print(Complete_File_Name, item['sheet'], item['range'])
            df = pd.read_excel(Complete_File_Name,
                           sheet_name = item['sheet'])
            df_Sales_Diesel = pd.concat([df_Sales_Diesel,df])
            
        # Treatments Sales General data
        df_Sales_General = tranform_df(df_Sales_General, ["COMBUSTÍVEL","ANO","REGIÃO","ESTADO","UNIDADE"], ['ANO', 'REGIÃO','ESTADO','month'])

        # Treatments Sales Diesel data
        df_Sales_Diesel = tranform_df(df_Sales_Diesel, ["COMBUSTÍVEL","ANO","SEGMENTO","ESTADO","UNIDADE"],['ANO', 'SEGMENTO','ESTADO','month']) 

        # Summarizing Sales General by Year
        df_sum_sales_general = df_Sales_General['volume'].groupby(df_Sales_General['year_month'].str.slice(0,4)).sum().reset_index()
        df_sum_sales_general['year_month'] = df_sum_sales_general['year_month'].astype(int)
        
        # Summarizing Sales Diesel by Year
        df_sum_sales_diesel = df_Sales_Diesel['volume'].groupby(df_Sales_Diesel['year_month'].str.slice(0,4)).sum().reset_index()
        df_sum_sales_diesel['year_month'] = df_sum_sales_diesel['year_month'].astype(int)        

        # Merge Sales General Source Data and Pivot Table and Calculate difference between source data and pivot table               
        merged_sales_general = pd.merge(left=df_total_sales_general, right=df_sum_sales_general, left_on='year', right_on='year_month', suffixes=('_Pivot', '_Source'))
        merged_sales_general['volume_Pivot_Minus_volume_Source'] = merged_sales_general['volume_Source'] - merged_sales_general['volume_Pivot'] 
        merged_sales_general.drop('year_month', axis=1, inplace=True)
        merged_sales_general
        
        # Merge Sales Diesel Source Data and Pivot Table  and Calculate difference between source data and pivot table
        merged_sales_diesel = pd.merge(left=df_total_sales_diesel, right=df_sum_sales_diesel, left_on='year', right_on='year_month', suffixes=('_Pivot', '_Source'))
        merged_sales_diesel['volume_Pivot_Minus_volume_Source'] = merged_sales_diesel['volume_Source'] - merged_sales_diesel['volume_Pivot'] 
        merged_sales_diesel.drop('year_month', axis=1, inplace=True)
        merged_sales_diesel        

        # Generate HTML result
        with open(html_file_name, 'w') as _file:
            _file.write('<center>' 
                        +'<h1> ANP - Summary Report </h1><br><hr>'
                        +'<h2> Sales of diesel by UF and type </h2>' + merged_sales_diesel.to_html(index=False,border=1,justify="center") + '<br><hr>'
                        +'<h2> Sales of oil derivative fuels by UF and product </h2>' + merged_sales_general.to_html(index=False,border=1,justify="center") + '<be><hr>'
                        +'<h1> The biggest difference founded in the volume of <font color="red">405.399 </font> in the report "Sales of oil derivative fuels by UF and product" in 2020 is due to the difference in Excel itself, between the source data and the pivot table, the data calculated in the program keeps the same difference, the others volumes the difference is in the 7th decimal place </h1><br><hr>'
                        +'</center>')
        
        # Showing Result HTML
        webbrowser.open(html_file_name)
        
    except Exception as e:
        print(e)       
    
if __name__ == ‘__main__’:
    main()
