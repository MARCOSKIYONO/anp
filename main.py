# LIBRARIES
import os, sys
import win32com.client
import pandas as pd
from datetime import date, datetime
import string
import urllib.request
import webbrowser

from functions_utils import load_total, transform_df, load_list, get_subject_positions, download_file, create_html
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
        excelModule.CodeModule.AddFromString(Sub_Detail_Pivot_tables.format(Sub_Name_Index,Sheet_Name_Index) )    

        # Add Macro Expand General Sales and Diesel Sales
        excelModule.CodeModule.AddFromString(sub_expand_sales.format(subname=Sub_Name_Sales,  mainsheet=Sheet_Name_Main) )

        # run the macro
        excel.Application.Run(Sub_Name_Index)

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
    try:
        df = df['volume'].groupby(df['year_month'].str.slice(0,4)).sum().reset_index()
        df['year_month'] = df['year_month'].astype(int)
        return df
    except Exception as e:
        print("Program sum_df error:{}".format(e))         


def merge_dfs(df_rigth,df_left):
    try:
        df_ = pd.merge(left=df_left, right=df_rigth, left_on='year', right_on='year_month', suffixes=('_Pivot', '_Source'))
        df_['volume_Pivot_Minus_volume_Source'] = df_['volume_Source'] - df_['volume_Pivot'] 
        df_.drop('year_month', axis=1, inplace=True)
        return df_    
    except Exception as e:
        print("Program sum_df error:{}".format(e))        


def load_df(lista_sheets):
    try:
        if lista_sheets:
            df_ = pd.DataFrame()
            for item in lista_sheets:
                df = pd.read_excel(Complete_File_Name,
                               sheet_name = item['sheet'])
                df_ = pd.concat([df_,df])
            return df_
    except Exception as e:
        print("Program load_df error:{}".format(e))         


def expand_excel_extract_data_sources(Complete_File_Name,sheet_name):
    try:
        
        # open up an instance of Excel with the win32com driver
        excel = win32com.client.Dispatch("Excel.Application")

        # do the operation in background without actually opening Excel
        excel.Visible = Debug_Visible_Excel

        # open the excel workbook from the specified file
        workbook = excel.Workbooks.Open(Filename=Complete_File_Name)
        
        # loading Pivot Table Info
        df = pd.read_excel(Complete_File_Name, sheet_name)
        
        # extracting Total Sales General       
        inicial_number, first_letter, last_number, last_letter = get_subject_positions(df,Pivot_Table_Name_Sales_General)
        List_Sales_General = load_list(inicial_number, first_letter, last_number, last_letter , excel,Sheet_Name_Main, Sheet_Prefix_Name_Sales_General)    
        df_total_sales_general = pd.DataFrame(load_total(inicial_number, first_letter, last_number, last_letter , excel,Sheet_Name_Main, Sheet_Prefix_Name_Sales_General))

        # extracting Total Sales Diesel   
        inicial_number, first_letter, last_number, last_letter = get_subject_positions(df,Pivot_Table_Name_Sales_Diesel)
        List_Sales_Diesel = load_list(inicial_number, first_letter, last_number, last_letter , excel,Sheet_Name_Main, Sheet_Prefix_Name_Sales_Diesel)
        df_total_sales_diesel = pd.DataFrame(load_total(inicial_number, first_letter, last_number, last_letter , excel,Sheet_Name_Main, Sheet_Prefix_Name_Sales_General))
            
        # expanding excel Sales General     
        for item in List_Sales_General:
            print("Expanding ", item['sheet'])
            excel.Application.Run(Sub_Name_Sales,item['sheet'] , item['range'])

        # expanding excel Sales Diesel              
        for item in List_Sales_Diesel:
            print("Expanding ", item['sheet'])
            excel.Application.Run(Sub_Name_Sales,item['sheet'] , item['range'])

        # save the workbook and close
        excel.Workbooks(1).Close(SaveChanges=1)
        excel.Application.Quit()
        
        # extracting Sales General info
        df_Sales_General = load_df(List_Sales_General)

        # extracting Sales Diesel info 
        df_Sales_Diesel = load_df(List_Sales_Diesel)

        return df_total_sales_general, df_Sales_General, df_total_sales_diesel, df_Sales_Diesel
    
    except Exception as e:
        excel.Workbooks(1).Close(SaveChanges=0)
        excel.Application.Quit()   
        df_total_sales_general = None 
        List_Sales_General = None 
        df_total_sales_diesel = None 
        df_Sales_Diesel = None 
        print("Program main Exand Excel error:{}".format(e))        


def main():
    try:   
        
        # downloading ANS File
        get_anp_file()
    
        # adding Module and Macros Excel
        increase_excel_macros()

        # expand Source Data from Pivot Tables and extract data from sheets
        df_total_sales_general, df_Sales_General, df_total_sales_diesel, df_Sales_Diesel = expand_excel_extract_data_sources(Complete_File_Name, Sheet_Name_Index)
        
        if df_total_sales_general.empty or df_Sales_General.empty or df_total_sales_diesel.empty or df_Sales_Diesel.empty:        
            raise ValueError('Program main error: Extracting')
            return False

        # transformation Sales General info 
        df_Sales_General = transform_df(df_Sales_General, ["COMBUSTÍVEL","ANO","REGIÃO","ESTADO","UNIDADE"], ['ANO', 'REGIÃO','ESTADO','month'], now )

        # transformation Sales Diesel info  
        df_Sales_Diesel = transform_df(df_Sales_Diesel, ["COMBUSTÍVEL","ANO","SEGMENTO","ESTADO","UNIDADE"],['ANO', 'SEGMENTO','ESTADO','month'], now ) 

        # summarizing Sales General info 
        df_sum_sales_general = sum_df(df_Sales_General)
        
        # summarizing Sales Diesel info
        df_sum_sales_diesel = sum_df(df_Sales_Diesel)      

        # merging Sales General info Calculate difference between source data and pivot table   
        merged_sales_general = merge_dfs(df_sum_sales_general,df_total_sales_general)
        
        # merging Sales Diesel info Calculate difference between source data and pivot table   
        merged_sales_diesel = merge_dfs(df_sum_sales_diesel,df_total_sales_diesel)
  
        # generating HTML info 
        create_html(html_file_name, merged_sales_diesel, merged_sales_general)
        
        # browsing Compare result  
        webbrowser.open(html_file_name)

    except Exception as e:
        print("Program main error:{}".format(e))  
        
        
if __name__ == ‘__main__’:
    main()
