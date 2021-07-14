# LIBRARIES
import os, sys
import win32com.client
import pandas as pd
from datetime import date, datetime
import string
import urllib.request
import webbrowser

from functions_utils import load_total, tranform_df, load_list, get_subject_positions, download_file, create_html
from utils import Months, Ufs, Sub_Detail_Pivot_tables, sub_expand_sales

# CONFIGURATIONS
#pd.set_option('precision', 9) # Trying to find out a way equalize the resuls when the difference it´s in the above the 7# decimal precision, but didn´t work
pd.options.display.float_format = '{: .12f}'.format
pd.set_option('mode.chained_assignment', None)

Always_Download = True # If need always download file before execute set True else False

Debug_Visible_Excel = True # If need see excel file while the program executes set True, else set False

# LOCAL VARIABLES
Donwload_URL = "http://www.anp.gov.br/arquivos/dados-estatisticos/vendas-combustiveis/vendas-combustiveis-m3.xls"

path = os.path.abspath(os.getcwd())

filename= "vendas-combustiveis-m3.xls"

html_file_name = "anp_sales.html"

Complete_File_Name = "{}\{}".format(path,filename)

Module_General_Name = 'Modulo_General'
Sheet_Name_Main = 'Plan1'
Sheet_Name_Index = 'Index'

Sheet_Prefix_Name_Sales_General = 'Sales_General_'
Sheet_Prefix_Name_Sales_Diesel = 'Sales_Diesel_'

Sub_Name_Index_ = 'Get_Pivot_Details'
Sub_Name_Sales = 'Get_Sales_General'

Pivot_Table_Name_Sales_General = 'Tabela dinâmica1'
Pivot_Table_Name_Sales_Diesel = 'Tabela dinâmica9'

now = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")

Sheets_List = [Sheet_Name_Index,Sheet_Prefix_Name_Sales_General,Sheet_Prefix_Name_Sales_Diesel]

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
        excelModule.CodeModule.AddFromString(Sub_Detail_Pivot_tables.format(Sub_Name_Index_,Sheet_Name_Index) )    

        # Add Macro Expand General Sales and Diesel Sales
        excelModule.CodeModule.AddFromString(sub_expand_sales.format(subname=Sub_Name_Sales,  mainsheet=Sheet_Name_Main) )

        # run the macro
        excel.Application.Run(Sub_Name_Index_)

        # save the workbook and close
        excel.Workbooks(1).Close(SaveChanges=1)
        excel.Application.Quit()

        # garbage collection
        del excel

    except Exception as e:
        # save the workbook and close
        excel.Workbooks(1).Close(SaveChanges=0)
        excel.Application.Quit()    
        print("Program increase_excel_macros error:{}".format(e))  
        

def get_anp_file():
    try:   
        # get directory where the script is located and Check File Exists
        Current_Path = os.path.abspath(os.getcwd())
        print(Current_Path)
        filepath= "{}\{}".format(Current_Path,filename)

        if os.path.exists(Complete_File_Name):
            print("File Exists !")
            if Always_Download:
                download_file(Donwload_URL, Complete_File_Name)
        else:
            print("File Do Not Exists !")
            download_file(Donwload_URL, Complete_File_Name)
    except Exception as e:
        print("Program get_anp_file error:{}".format(e)) 

def sum_df(df):
    df = df['volume'].groupby(df['year_month'].str.slice(0,4)).sum().reset_index()
    df['year_month'] = df['year_month'].astype(int)
    return df

def merge_dfs(df_rigth,df_left):
        df_ = pd.merge(left=df_left, right=df_rigth, left_on='year', right_on='year_month', suffixes=('_Pivot', '_Source'))
        df_['volume_Pivot_Minus_volume_Source'] = df_['volume_Source'] - df_['volume_Pivot'] 
        df_.drop('year_month', axis=1, inplace=True)
        return df_    
        
def main():
    try:   
        # get ANP File
        get_anp_file()
    
        # Increase Module and Macros Excel
        increase_excel_macros()
        
        # Loading Pivot Table Info
        df = pd.read_excel(Complete_File_Name, sheet_name=Sheet_Name_Index)
        
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
            inicial_number, first_letter, last_number, last_letter = get_subject_positions(df,Pivot_Table_Name_Sales_General)
            List_Sales_General = load_list(inicial_number, first_letter, last_number, last_letter , excel,Sheet_Name_Main, Sheet_Prefix_Name_Sales_General)    
            df_total_sales_general = pd.DataFrame(load_total(inicial_number, first_letter, last_number, last_letter , excel,Sheet_Name_Main, Sheet_Prefix_Name_Sales_General))

            # Get Sales Diesel sheets positions
            inicial_number, first_letter, last_number, last_letter = get_subject_positions(df,Pivot_Table_Name_Sales_Diesel)
            List_Sales_Diesel = load_list(inicial_number, first_letter, last_number, last_letter , excel,Sheet_Name_Main, Sheet_Prefix_Name_Sales_Diesel)
            df_total_sales_diesel = pd.DataFrame(load_total(inicial_number, first_letter, last_number, last_letter , excel,Sheet_Name_Main, Sheet_Prefix_Name_Sales_General))

            # Generate details Sales General infos
            for item in List_Sales_General:
                print(item['sheet'], item['range'])
                excel.Application.Run(Sub_Name_Sales,item['sheet'] , item['range'])

            # Generate details Sales Diesel infos
            for item in List_Sales_Diesel:
                print(item['sheet'], item['range'])
                excel.Application.Run(Sub_Name_Sales,item['sheet'] , item['range'])
                
            # save the workbook and close
            excel.Workbooks(1).Close(SaveChanges=1)
            excel.Application.Quit()

        except Exception as e:
            excel.Workbooks(1).Close(SaveChanges=0)
            excel.Application.Quit()   
            print("Program main Exand Excel error:{}".format(e))
            return False
        
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

        print("Summarizing Sales General info ...")
        # Summarizing Sales General by Year
        df_sum_sales_general = sum_df(df_Sales_General)
        
        print("Summarizing Sales Diesel info ...")
        # Summarizing Sales Diesel by Year
        df_sum_sales_diesel = sum_df(df_Sales_Diesel)      

        print("Merging Sales General info ...")    
        # Merge Sales General Source Data and Pivot Table and Calculate difference between source data and pivot table 
        merged_sales_general = merge_dfs(df_sum_sales_general,df_total_sales_general)
        
        print("Merging Sales Diesel info ...")    
        # Merge Sales Diesel Source Data and Pivot Table  and Calculate difference between source data and pivot table
        merged_sales_diesel = merge_dfs(df_sum_sales_diesel,df_total_sales_diesel)      

        # Generate HTML result
        create_html(html_file_name, merged_sales_diesel, merged_sales_general)
        
        # Showing Result HTML
        webbrowser.open(html_file_name)
        
    except Exception as e:
        print("Program main error:{}".format(e))  
    
if __name__ == ‘__main__’:
    main()
