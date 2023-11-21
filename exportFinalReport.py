import csv
import openpyxl
import pandas as pd
import pdb
from win32com.client import Dispatch
import os
from datetime import date
from pathlib import Path
from openpyxl import Workbook, load_workbook
from copy import copy
import xlwings as xw
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl import Workbook
import pythoncom
import datetime
import re






# create equip 
def get_equipments_list(equip_table):
    
# 1: create euiq_dict {
#       "block": "BuildingA',
#       "floor":"Basement 1",
#       "area":"Floor Area",
#       "type":"ASEF",
#       "title":"ASEF-01-01 to ASEF-01-07",
#       "quantity":2,
#       "no_of_points": 5,
#       "smart_points":2,
#       "demarcation": "Base build",
#       "plusSpares":8,
#       "MCC": "MCC-B-CPEF",
#       "ElectricalLoad": FALSE/TRUE
#              }

#### 2: create equi_qty per mcc location {'MMC-A' : {"fcu" : 6, "ahu":9}}
    with open(equip_table,"r") as equipTable:      
        csvreader = csv.reader(equipTable)
        header = next(csvreader)
        equipments_list = [] # type: list 
        equip_qty_per_MCC_dict = {}
        block_set = set()
        floor_num_dict = {}
        floor_set = set()
        for row in csvreader:
            if row[3].split("_")[0] == "Others":
           
                equip_dict = {
                    "block": row[0],
                    "floor": row[1],
                    "area":  row[2],
                    "type":f'Others({row[3].split("_",1)[1].replace('_', ' ')})',
                    "title":row[4],
                    "quantity":int(row[6]),
                    "no_of_points":int(row[7]),
                    "smart_points":int(row[8]),
                    "demarcation": row[5],
                    "plusSpares":int(row[9]),
                    "MCC":row[10],
                    "Electrical_load":row[11]
                }
            else:
                equip_dict = {
                "block": row[0],
                "floor": row[1],
                "area":  row[2],
                "type":row[3],
                "title":row[4],
                "quantity":int(row[6]),
                "no_of_points":int(row[7]),
                "smart_points":int(row[8]),
                "demarcation": row[5],
                "plusSpares":int(row[9]),
                "MCC":row[10],
                "Electrical_load":row[11]
            }
            equipments_list.append(
                equip_dict
            )
            block_set.add(row[0])
           
        ### only electruical load == true is needed to be filled in electric load table
            if row[11] == "TRUE":
                
                if row[10] in equip_qty_per_MCC_dict:
                    equip_qty_per_MCC_dict[row[10]][row[4]] = int(row[6])
                else:
                    equip_qty_per_MCC_dict[row[10]]= {row[4]:int(row[6])} 

            if 'floor' in row[1].lower() or 'level' in row[1].lower():
                floor_set.add(row[1].lower().strip().replace(" ", ""))
                # update when iterate equip
                floor_num_dict[row[0]]  = len(floor_set)
    block_list = list(sorted(block_set))
    #floor_num_dict : the number of the floor in one block {"Block A": 10, "Block B": 11}         [for point summay sheet]
    # block_list: the list of all the blocks in one project ["Block A", "Block B"]                [for iterate in summary sheet]
    # equip_qty_per_MCC_dict: the total quantity of equipment in one MCC location {'MCCZ-1': 1}   [for electrical load]
   
    return equipments_list,equip_qty_per_MCC_dict,block_list,floor_num_dict

def get_points_list(point_table):  
# point_list
# create euiq_dict {
#       "type":"ASEF",
#       "project_id":12345,
#       "title":"supply temperature",
#       "category": "Standard",
#       "point_type":"Monitor",
#       "point_selected": "Y",
#       "signal_type":"Digital OUT",
#       "cable_type":"BACnet IP",
#       "functional_description":"this is description",
#       
#              }
    point_equip_type_set = set()
    with open(point_table,"r") as pointTable:      
        csvreader = csv.reader(pointTable)
        header = next(csvreader)
        point_list = [] # type: list 
        for row in csvreader:
            
            point_dict = {
                "type":row[0],
                "description":row[1],
                "category": row[3],
                "point_type":row[2],
                "signal_type":row[4],
                "cable_type":row[5],
                "functional_description":row[6],
                "comments":row[7]
            }
            point_list.append(
                point_dict
            )
            point_equip_type_set.add(row[0])
    return point_list, point_equip_type_set


def createPointSummaryDF(equip_table):
###### point summary
    table_row_num =[]
    point_sum_row_num =[]   ### the list of point sum table
    area_count_per_floor_list = {}    # the dict of area count for each block {'Block A': Series(2,3,4,4)}
    tot_point_per_block = {} # the total number of points per block {"Block A": 89} sum of plus spares
    tot_point_per_floor_point = {} # the total number of points per floor {"Block A":series ( 4, 34) } sum of plus spares
    point_total = pd.DataFrame()
    equipments_list,_, block_list, floor_num_dict = get_equipments_list(equip_table)
    for i, block in enumerate(block_list):
        equipments_list_per_block= []
        
        for equip in equipments_list:  
            if equip['block'] == block:
                equipments_list_per_block.append(equip)

        equip_df = pd.DataFrame(equipments_list_per_block)
        equip_df['project_points'] = equip_df['quantity'] * equip_df['no_of_points'] 
        equip_df['totals'] = equip_df['plusSpares'] 
        point_sum =  equip_df[["block","area","demarcation","project_points","smart_points","plusSpares","totals"]]
        point_sum = point_sum.sort_values(["block","area"])
        point_sum = point_sum.groupby(["block","area","demarcation"]).agg({"project_points":'sum',"smart_points":"sum","plusSpares":"sum","totals":"sum"}, inplace = True)
        tot_point_per_block[block] = point_sum.groupby('block').sum()['plusSpares']
        

        point_sum = point_sum.reset_index( drop=False)
        # Rename the columns of each DataFrame to concate
        point_sum_renamed = point_sum.rename(columns={point_sum.columns[i]: 'common_name' for i in range(len(point_sum.columns))})
        # row number of point summary for each block (df) + 1(table1 header) + 1 (table1 header)
        point_sum_row_num.append(point_sum.shape[0] + 2)

###### 1 empty line
 # Create a series to be added between dataframe
        empty_line = pd.Series([" "] * 7)
        title = pd.Series(["Floor Summary","Area","Demarcation", 'Project Points','Smart Points','Plus Spare', 'Totals'])
        #### need to let placeholder keep different otherwise the cell will be merged and later format cannot write in this merged cell
        header_line = pd.Series(["8"] * 7)
        title_placeholder = pd.Series(["*"] * 7)
        

    ###### point breakdown
        equip_df1 = pd.DataFrame(equipments_list)
        equip_df1['project_points'] = equip_df1['quantity'] * equip_df1['no_of_points'] 
        equip_df1['totals'] = equip_df1['plusSpares'] 
        point_bd =  equip_df1[["floor","area","demarcation","project_points","smart_points","plusSpares","totals"]]
        point_bd = point_bd.sort_values(["floor","area","demarcation"])
        point_bd = point_bd.groupby(["floor","area","demarcation"]).agg({"project_points":"sum","smart_points":"sum","plusSpares":"sum","totals":"sum"}, inplace = True)
        point_bd =point_bd.reset_index( drop=False)
        tot_point_per_floor_point[block] = point_bd.groupby(['floor']).sum()['plusSpares']
        point_bd_renamed = point_bd.rename(columns={point_bd.columns[i]: 'common_name' for i in range(len(point_bd.columns))})
        area_count = point_bd.groupby('floor').count()['totals']
        area_count_per_floor_list[block] = area_count
        

  
        if i == len(block_list) - 1:
            # Concatenate the DataFrames
            df_concat = pd.concat([point_sum_renamed,
                               empty_line.set_axis(point_sum_renamed.columns).to_frame().T,  
                               title.set_axis(point_sum_renamed.columns).to_frame().T,
                               point_bd_renamed])
           
        else:
        # Concatenate the DataFrames
            df_concat = pd.concat([point_sum_renamed,
                                empty_line.set_axis(point_sum_renamed.columns).to_frame().T,  
                                title.set_axis(point_sum_renamed.columns).to_frame().T,
                                point_bd_renamed,
                                empty_line.set_axis(point_sum_renamed.columns).to_frame().T,
                                header_line.set_axis(point_sum_renamed.columns).to_frame().T,
                                title_placeholder.set_axis(point_sum_renamed.columns).to_frame().T])
            
        df_concat.columns = ["Summary","Area","Demarcation",'Project Points','Smart Points','Plus Spare', 'Totals']
        df_concat.set_index(['Summary',"Area","Demarcation",'Project Points'], inplace= True)
        # block breakdown header(1) + point summary title(1) + point summary(only data no title) + emptyline(1) + title(1) + point breakdown row (only df no title)
        
        table_row_num.append(point_sum_renamed.shape[0] + point_bd_renamed.shape[0] + 4)
        point_total =  pd.concat([point_total,df_concat])


    #point_sum_row_num
    return point_total, point_sum_row_num, area_count_per_floor_list,block_list, table_row_num,tot_point_per_block,tot_point_per_floor_point

def createEquipSummaryDF(equip_table):
    #record the total row num of each table(block)
    table_row_num =[]
    equip_sum_row_num =[]
    equip_total = pd.DataFrame()
    equip_count_per_floor_list = {}
    tot_point_per_block = {}
    tot_point_per_floor_equip = {}
######  equip smmmary  
    equipments_list, _ , block_list, _= get_equipments_list(equip_table)
    
######FILTER ACCORDING TO BLOCK
    for i, block in enumerate(block_list):
        equipments_list_per_block= []
        
        for equip in equipments_list:  
            if equip['block'] == block:
                equipments_list_per_block.append(equip)
         

        equip_df = pd.DataFrame(equipments_list_per_block)
        equip_df['totals'] = equip_df['plusSpares'] 
        equip_df['sub_totals'] = equip_df['quantity'] * equip_df['plusSpares'] 
        equip_sum =  equip_df[["block","type","demarcation","quantity","no_of_points","smart_points","plusSpares","sub_totals","totals"]]
        equip_sum = equip_sum.sort_values(["block","type","demarcation"])  
        equip_sum = equip_sum.groupby(["block","type","demarcation"]).agg({"quantity":'sum',"no_of_points":"mean", "smart_points":"sum","plusSpares":"sum","sub_totals":"sum", "totals":"sum"}, inplace = True)
        tot_point_per_block[block] = equip_sum.groupby('block').sum()['sub_totals']
        equip_sum = equip_sum.reset_index(drop=False)
        # Rename the columns of each DataFrame to concate
        equip_sum_renamed = equip_sum.rename(columns={equip_sum.columns[i]: 'common_name' for i in range(len(equip_sum.columns))})

        equip_sum_row_num.append(equip_sum.shape[0] + 2)
        
    ###### 1 empty line
    # Create a series to be added between dataframe
        empty_line = pd.Series([" "] * 9)
        title = pd.Series(["Floor Summary","Equipment Type","Demarcation", 'Total Quantities','Unit Points','Smart Points','Plus Spares', 'Sub-Totals','Totals'])
        #### otherwise the cell will be merged and later format cannot write in this merged cell
        header_line = pd.Series(["8"] * 9)
        title_placeholder = pd.Series(["*"] * 9)

    # equip breakdown 
        equip_df1 = pd.DataFrame(equipments_list_per_block)
        equip_df1['totals'] = equip_df1['plusSpares'] 
        equip_df1['sub_totals'] = equip_df1['quantity'] * equip_df1['plusSpares'] 
        equip_bd =  equip_df1[["floor","type","demarcation","quantity","no_of_points","smart_points","plusSpares","sub_totals","totals"]]
        equip_bd = equip_bd.sort_values(["floor","type","demarcation"])
        equip_bd = equip_bd.groupby(["floor","type","demarcation"]).agg({"quantity":'sum',"no_of_points":"mean", "smart_points":"sum","plusSpares":"sum","sub_totals":"sum", "totals":"sum"}, inplace = True)
        equip_bd = equip_bd.reset_index( drop=False)
        equip_count = equip_bd.groupby('floor').count()['totals']
        equip_count_per_floor_list[block] = equip_count
        
        tot_point_per_floor_equip[block] = equip_bd.groupby('floor').sum()['sub_totals']
        equip_bd_renamed = equip_bd.rename(columns={equip_bd.columns[i]: 'common_name' for i in range(len(equip_bd.columns))})

        if i == len(block_list) - 1:
            # last table no need to add placeholder
          
            equip_concat = pd.concat([equip_sum_renamed,
                                      empty_line.set_axis(equip_sum_renamed.columns).to_frame().T,  
                                      title.set_axis(equip_sum_renamed.columns).to_frame().T,
                                      equip_bd_renamed]) 
        # Concatenate the DataFrames
        else:
            equip_concat = pd.concat([equip_sum_renamed,
                                      empty_line.set_axis(equip_sum_renamed.columns).to_frame().T, 
                                      title.set_axis(equip_sum_renamed.columns).to_frame().T,
                                      equip_bd_renamed,
                                      empty_line.set_axis(equip_sum_renamed.columns).to_frame().T,
                                      header_line.set_axis(equip_sum_renamed.columns).to_frame().T,
                                      title_placeholder.set_axis(equip_sum_renamed.columns).to_frame().T])
            
        # header (breakdown)(1) +  equip summary title(1) + point summary(only data no title) + emptyline(1) + title(1) + point breakdown row (only df no title)
        table_row_num.append(equip_sum_renamed.shape[0] + equip_bd_renamed.shape[0] + 4)
        equip_total =  pd.concat([equip_total,equip_concat])


    equip_total.columns = ["Floor Summary","Equipment Type","Demarcation",'Total Quantities','Unit Points','Smart Points', 'Plus Spares', 'Sub-Totals' ,'Totals']
    equip_total.set_index(['Floor Summary',"Equipment Type","Demarcation",'Total Quantities'], inplace= True)
    # total row num  of equip sum(table 1) 1 header + 1 title + df
    
    return equip_total,equip_sum_row_num,equip_count_per_floor_list,block_list, table_row_num,tot_point_per_block, tot_point_per_floor_equip

def create_point_sum_format(writer,sheet_name,point_sum_row_num,point_sum_df,area_count_per_floor_list,block_list,table_row_num,tot_point_per_block,tot_point_per_floor_point):
    
    # point_sum_iloc (x, 3)
    # point_sum_df.index.values[0][3] 
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]
 
     #set whole sheet format
    workbook.formats[0].set_font_name('Cambria')
    workbook.formats[0].set_font_size(8)
    workbook.formats[0].set_align('center')
    workbook.formats[0].set_align('vcenter')

     ## set width
    worksheet.set_column('B:B', 29)
    worksheet.set_column('C:C', 20)
    worksheet.set_column('D:D', 20) 
    worksheet.set_column('E:E', 15)
    worksheet.set_column('F:F', 15)
    worksheet.set_column('G:G', 15)
    worksheet.set_column('H:H', 10)

    offset = 0
    
    # set cannot access by index, so convert it into list to get index
    
    for i, block in enumerate(block_list):
        ######### HEADER STYLE
        # Add Header format
        format_header = workbook.add_format({'bg_color': '#094121','bold': True,'font_color': 'white','font_size':16,'align': 'center','font_name':'Cambria','valign': 'vcenter','bottom':5,'top':5, 'left':5, 'right':5})
    
        # set Header place
        title =  f"{block} - Floor Breakdown"
        worksheet.merge_range(f'B{2+offset}:H{2+offset}',title,format_header)
        worksheet.set_row(1 + offset, 45)

        ##### set 1st table title
        
        border_format = workbook.add_format({'bold': True,'bottom':5,'top':5, 'left':5, 'right':5,'font_name':'Cambria','font_size':10,'valign': 'vcenter','bold': True,'align': 'center'})
        
        worksheet.set_row(2 + offset, 33)
        worksheet.write(f'B{3+offset}',"Building Summary",border_format)
        worksheet.write(f'C{3+offset}',"Locations",border_format)
        worksheet.write(f'D{3+offset}',"Demarcation",border_format)
        worksheet.write(f'E{3+offset}',"Project points",border_format)
        worksheet.write(f'F{3+offset}',"Smart Points",border_format)
        worksheet.write(f'G{3+offset}',"Plus Spares",border_format)
        worksheet.write(f'H{3+offset}',"Totals",border_format)


        ##### summary total merge 
        ## 3 is the 1 header 2 title 1 jumpback to last row
        
        worksheet.merge_range(f'H{4 + offset}:H{4 + offset + point_sum_row_num[i] - 3}', tot_point_per_block[block])



        ######## 2nd TITLE STYLE
        # title border formate
        # same num in excel
      
        border_format = workbook.add_format({'bg_color': '#094121','bold': True,'font_color': 'white','bottom': 5,'top':5, 'left':5, 'right':5,'font_name':'Cambria','font_size':10,'valign': 'vcenter','bold': True,'align': 'center'})
        # set_row start from 0---> row num - 1
        # empty 1(begin of table) + empty line (end of table) + 1 new line 
        point_bd_row = 3 + offset + point_sum_row_num[i]
        worksheet.set_row(point_bd_row -1, 33)
        worksheet.write(f'B{point_bd_row}',"Floor Summary",border_format)
        worksheet.write(f'C{point_bd_row}',"Area",border_format)
        worksheet.write(f'D{point_bd_row}',"Demarcation",border_format)
        worksheet.write(f'E{point_bd_row}',"Project Points",border_format)
        worksheet.write(f'F{point_bd_row}',"Smart Points",border_format)
        worksheet.write(f'G{point_bd_row}',"Plus Spare",border_format)
        worksheet.write(f'H{point_bd_row}',"Totals",border_format)
        worksheet.insert_image("B2", "logo.png",{"x_offset": 30, "y_offset": 10})

    #####  MERGE total
        #1 empty at beginig of table A , 1 empty row at end of table A + 1 2nd table title + 1 start from next one
        
        start = offset + point_sum_row_num[i] + 4
        for num , count in enumerate(area_count_per_floor_list[block]):
            
            worksheet.merge_range(f'H{start}:H{start + count - 1}', tot_point_per_floor_point[block].iloc[num])
            start += count 

        offset += table_row_num[i]  + 1
       
        
        
 
def create_equip_sum_format(writer,sheet_name,equip_sum_row_num,equip_sum_df,equip_count_per_floor_list,block_list,table_row_num,tot_point_per_block_equip, tot_point_per_floor_equip):
    ##equip_sum_row_num: [75,5]
    ##equip_sum_df.index(75, 4)
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]
     #set whole sheet format
    workbook.formats[0].set_font_name('Cambria')
    workbook.formats[0].set_font_size(8)
    workbook.formats[0].set_align('center')
    workbook.formats[0].set_align('vcenter')

     ## set width
    worksheet.set_column('B:B', 29)
    worksheet.set_column('C:C', 29)
    worksheet.set_column('D:D', 20) 
    worksheet.set_column('E:E', 18)
    worksheet.set_column('F:F', 15)
    worksheet.set_column('G:G', 15)
    worksheet.set_column('H:H', 15)
    worksheet.set_column('I:I', 15)
    worksheet.set_column('J:J', 10)

    offset = 0
    
    # set cannot access by index, so convert it into list to get index
    

    for i, block in enumerate(block_list):

        ######### HEADER STYLE
        # Add Header format
        format_header = workbook.add_format({'bg_color': '#094121','bold': True,'font_color': 'white','font_size':16,'align': 'center','font_name':'Cambria','valign': 'vcenter','bottom': 5,'top':5, 'left':5, 'right':5})
    
        # set Header place
        title =  f"{block} - Equipment Breakdown"
    
        worksheet.merge_range(f'B{2 + offset}:J{2 + offset}',title,format_header)
        
        worksheet.set_row(1 + offset, 45) 

        ####### 1st table title style
        border_format = workbook.add_format({'bold': True,'bottom': 5,'top':5, 'left':5, 'right':5,'font_name':'Cambria','font_size':10,'valign': 'vcenter','bold': True,'align': 'center'})
        ###setrow start from 0
        worksheet.set_row(2 + offset, 33)
        worksheet.write(f'B{3 + offset}',"Building Summary",border_format)
        worksheet.write(f'C{3 + offset}',"Equipment Type",border_format)
        worksheet.write(f'D{3 + offset}',"Demarcation",border_format)
        worksheet.write(f'E{3 + offset}',"Total Quantities",border_format)
        worksheet.write(f'F{3 + offset}',"Unit Points",border_format)
        worksheet.write(f'G{3 + offset}',"Smart Points",border_format)
        worksheet.write(f'H{3 + offset}',"Plus Spares",border_format)
        worksheet.write(f'I{3 + offset}',"Sub-Totals",border_format)
        worksheet.write(f'J{3 + offset}',"Totals",border_format)

        
        ### summary total merge 
        # 3 is the 1 header 2 title 1 jumpback to last row
        worksheet.merge_range(f'J{4 + offset}:J{4 + offset + equip_sum_row_num[i] - 3 }', tot_point_per_block_equip[block])


        ####### 2nd TITLE STYLE
        #title border formate
        border_format = workbook.add_format({'bg_color': '#094121','bold': True,'font_color': 'white','bottom': 5,'top':5, 'left':5, 'right':5,'font_name':'Cambria','font_size':10,'valign': 'vcenter','bold': True,'align': 'center'})
        # empty 1(begin of table) + empty line (end of table) + 1 new line 
        equip_bd_row = 3 + offset + equip_sum_row_num[i]
        worksheet.set_row(equip_bd_row - 1, 33)
        worksheet.write(f'B{equip_bd_row}',"Floor Summary",border_format)
        worksheet.write(f'C{equip_bd_row}',"Equipment Type",border_format)
        worksheet.write(f'D{equip_bd_row}',"Demarcation",border_format)
        worksheet.write(f'E{equip_bd_row}',"Total Quantities",border_format)
        worksheet.write(f'F{equip_bd_row}',"Unit Points",border_format)
        worksheet.write(f'G{equip_bd_row}',"Smart Points",border_format)
        worksheet.write(f'H{equip_bd_row}',"Plus Spares",border_format)
        worksheet.write(f'I{equip_bd_row}',"Sub-Totals",border_format)
        worksheet.write(f'J{equip_bd_row}',"Totals",border_format)
        worksheet.insert_image("B2", "logo.png",{"x_offset": 30, "y_offset": 10})

  
        #######  MERGE total
        # 1 empty at beginig of table A , 1 empty row at end of table A + 1 2nd table title + start from next one
       
        start = offset + equip_sum_row_num[i] + 4
        for num , count in enumerate( equip_count_per_floor_list[block]):
            #worksheet.merge_range(f'J{start}:J{ start + count - 1}', tot_point_per_floor_equip[block].iloc[c])
            
            worksheet.merge_range(f'J{start}:J{ start + count - 1}',tot_point_per_floor_equip[block].iloc[num])
            start += count 
           




        offset += table_row_num[i] + 1
        
    
        
def createPointsDF(group,point_equip_type_set, point_list ):       
    filter_equip = group[["Equipment","Title","Floor","Demarcation","Quantity"]]
    filter_equip.sort_values(["Equipment","Title","Floor"])
    # get qty info based on groupby character
    qty_info = filter_equip.groupby(["Equipment","Title","Floor","Demarcation"]).sum()
    point_records_list=[]
    for pinfo , pty in qty_info.iterrows():
        #if equip info can br fount in point list (try catch?)
        if pinfo[0] in point_equip_type_set:
            #find pointdescroption
            
            for p in  point_list:
                # find point list according to the key
                if p['type'] == pinfo[0] :
                    if pinfo[3] == "Landlord":
                        # Other_core -> Other(Core)
                        if pinfo[0].split("_")[0] == "Others":
                            point_record = {
                                    "Equipment": f"Others({pinfo[0].split("_")[1].replace('_', ' ')})",
                                    "Equipment Tag": f"Others({pinfo[1].split("_")[1].replace('_', ' ')})",
                                    "LL-QTY": pty.Quantity,
                                    "FO-QTY":0,
                                    "Floor": pinfo[2],
                                    "Point Description": p["description"],
                                    "Points Type":p["point_type"],
                                    "Signal Type":p["signal_type"],
                                    "Cable Type": p["cable_type"],
                                    "Functional Description":p["functional_description"],
                                    "Comments":""
                                    
                                }
                        else:
                            point_record = {
                                    "Equipment": pinfo[0],
                                    "Equipment Tag": pinfo[1],
                                    "LL-QTY": pty.Quantity,
                                    "FO-QTY":0,
                                    "Floor": pinfo[2],
                                    "Point Description": p["description"],
                                    "Points Type":p["point_type"],
                                    "Signal Type":p["signal_type"],
                                    "Cable Type": p["cable_type"],
                                    "Functional Description":p["functional_description"],
                                    "Comments":""
                                    
                                }
                
                    else:
                        if pinfo[0].split("_")[0] == "Others":
                            point_record = {
                                        "Equipment":  f"Others({pinfo[0].split("_")[1]})",
                                        "Equipment Tag": f"Others({pinfo[1].split("_")[1]})",
                                        "LL-QTY": 0,
                                        "FO-QTY":pty.Quantity,
                                        "Floor": pinfo[2],
                                        "Point Description": p["description"],
                                        "Points Type":p["point_type"],
                                        "Signal Type":p["signal_type"],
                                        "Cable Type": p["cable_type"],
                                        "Functional Description":p["functional_description"],
                                        "Comments":""
                                    }   
                        else:
                            point_record = {
                                    "Equipment": pinfo[0],
                                    "Equipment Tag": pinfo[1],
                                    "LL-QTY": 0,
                                    "FO-QTY":pty.Quantity,
                                    "Floor": pinfo[2],
                                    "Point Description": p["description"],
                                    "Points Type":p["point_type"],
                                    "Signal Type":p["signal_type"],
                                    "Cable Type": p["cable_type"],
                                    "Functional Description":p["functional_description"],
                                    "Comments":""
                                }           
                    point_records_list.append(point_record)
    
    point_record_df = pd.DataFrame( point_records_list)
    point_record_df.set_index(["Equipment","Equipment Tag","LL-QTY","FO-QTY","Floor","Point Description"], inplace = True)
    return point_record_df
        #create_report_per_MCC( point_record_df,MCCname[0],writer)


def createPointsDF2(group,point_equip_type_set, point_list ):  
    group.sort_values(["Equipment","Title","Floor"])
    point_total = pd.DataFrame()
    row_number_list = []
    
    floor_name_list = []
    
    total_BMS_point_list_per_floor=[]
    for floor_name , group_per_floor in group.groupby(['Floor']):
        floor_name_list.append(floor_name[0])
        point_record_per_floor = []
        
        
        filter_equip = group_per_floor[["Equipment","Title","Floor","Demarcation","Quantity"]]
        filter_equip.sort_values(["Equipment","Title","Floor"])
        # get qty info based on groupby character
        qty_info = filter_equip.groupby(["Equipment","Title","Floor","Demarcation"]).sum()
        
        total_point_per_floor = 0
        empty_line = pd.Series([" "] * 11)
        total_placeholder = pd.Series(["x"] * 11)
        header_placeholder = pd.Series(["h"] * 11)
        title_placeholder = pd.Series(["t"] * 11)
        for pinfo , pty in qty_info.iterrows():
            
            #qty.info: pinfo[0]: FCU   pty 65
            point_records_list_per_equip=[]
            
            #if equip info can be found in point list (DB)
            if pinfo[0] in point_equip_type_set:
                
                #find pointdescroption     
                for p in  point_list:
                    # find point list according to the key p
                    if p['type'] == pinfo[0] :
                        # count += 1
                        
                         #to fill the qty into LL or FO
                        if pinfo[3] == "Landlord":
                            # Other_core -> Other(Core)
                            if pinfo[0].split("_")[0] == "Others":
                                point_record = {
                                        "Equipment": f"Others({pinfo[0].split("_")[1].replace('_', ' ')})",
                                        "Equipment Tag": f"Others({pinfo[1].split("_")[1].replace('_', ' ')})",
                                        "LL-QTY": pty.Quantity,
                                        "FO-QTY":0,
                                        "Floor": pinfo[2],
                                        "Point Description": p["description"],
                                        "Points Type":p["point_type"],
                                        "Signal Type":p["signal_type"],
                                        "Cable Type": p["cable_type"],
                                        "Functional Description":p["functional_description"],
                                        "Comments":""
                                        
                                    }
                            else:
                                point_record = {
                                        "Equipment": pinfo[0],
                                        "Equipment Tag": pinfo[1],
                                        "LL-QTY": pty.Quantity,
                                        "FO-QTY":0,
                                        "Floor": pinfo[2],
                                        "Point Description": p["description"],
                                        "Points Type":p["point_type"],
                                        "Signal Type":p["signal_type"],
                                        "Cable Type": p["cable_type"],
                                        "Functional Description":p["functional_description"],
                                        "Comments":""
                                        
                                    }
                
                        else:
                            if pinfo[0].split("_")[0] == "Others":
                                point_record = {
                                            "Equipment":  f"Others({pinfo[0].split("_")[1]})",
                                            "Equipment Tag": f"Others({pinfo[1].split("_")[1]})",
                                            "LL-QTY": 0,
                                            "FO-QTY":pty.Quantity,
                                            "Floor": pinfo[2],
                                            "Point Description": p["description"],
                                            "Points Type":p["point_type"],
                                            "Signal Type":p["signal_type"],
                                            "Cable Type": p["cable_type"],
                                            "Functional Description":p["functional_description"],
                                            "Comments":""
                                        }   
                            else:
                                
                                point_record = {
                                        "Equipment": pinfo[0],
                                        "Equipment Tag": pinfo[1],
                                        "LL-QTY": 0,
                                        "FO-QTY":pty.Quantity,
                                        "Floor": pinfo[2],
                                        "Point Description": p["description"],
                                        "Points Type":p["point_type"],
                                        "Signal Type":p["signal_type"],
                                        "Cable Type": p["cable_type"],
                                        "Functional Description":p["functional_description"],
                                        "Comments":""
                                    }
                        point_records_list_per_equip.append(point_record)
                        point_record_per_floor.append(point_record)
            
                
            else:
                print("This equipment cannot be found in point list table")
            
            total_point_per_equip = pty.Quantity * len(point_records_list_per_equip)
            
            total_point_per_floor += total_point_per_equip

        
        point_record_per_floor = pd.DataFrame( point_record_per_floor)  
        row_number_list.append(point_record_per_floor.shape[0] + 2)
        total_BMS_point_list_per_floor.append(total_point_per_floor)
           
            

        if point_total.empty:  
            
            point_total = pd.concat([point_total, 
                                    point_record_per_floor])
        else:
           
            point_total = pd.concat([point_total,
                                     empty_line.set_axis(point_record_per_floor.columns).to_frame().T,
                                    total_placeholder.set_axis(point_record_per_floor.columns).to_frame().T,
                                    empty_line.set_axis(point_record_per_floor.columns).to_frame().T,
                                    header_placeholder.set_axis(point_record_per_floor.columns).to_frame().T,
                                    title_placeholder.set_axis(point_record_per_floor.columns).to_frame().T,
                                    point_record_per_floor])
    
    
        
    point_total.set_index(["Equipment","Equipment Tag","LL-QTY","FO-QTY","Floor","Point Description"], inplace = True)
    # point_total: point detail for all floor (total dataframe)
    # floor_name list: the list of floor name for header
    # row_number_list: total point (row num) per floor
    # total row number per floor(Header + title + content) NO EMPTY ROW
    # total_BMS_point_list_per_floor : total BMS points per floor(for total value)
    return point_total, row_number_list, floor_name_list,total_BMS_point_list_per_floor



def create_point_list_sheet_format(tot_df,MCCname,row_number_list, floor_name_list,total_BMS_point_list_per_floor,spare_points,writer):
    # Convert the merged dataframe to an XlsxWriter Excel object.
   
    sheet_name = f"{MCCname} Point List"
    header_start_row = 1
    tot_df.to_excel(writer, sheet_name= sheet_name,startrow = 2,startcol = 1)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets[sheet_name]

    #set whole sheet format
    workbook.formats[0].set_font_name('Cambria')
    workbook.formats[0].set_font_size(8)


    border_format = workbook.add_format({'bottom': 2,'top':4, 'left':4, 'right':4,'font_name':'Cambria','font_size':8})


    ## set width
    worksheet.set_column('B:B', 30)
    worksheet.set_column('C:C', 30)
    worksheet.set_column('D:D', 10) 
    worksheet.set_column('E:E', 10)
    worksheet.set_column('F:F', 20)
    worksheet.set_column('G:G', 35)
    worksheet.set_column('H:H', 30)
    worksheet.set_column('I:I', 30)
    worksheet.set_column('J:J', 30)
    worksheet.set_column('K:K', 120)
    worksheet.set_column('L:L', 30)

    offset = 0
    
    for count, floor_name in enumerate(floor_name_list):

        ######### HEADER STYLE
        # Add Header format
        format_header = workbook.add_format({'bg_color': '#094121','bold': True,'font_color': 'white','font_size':16,'align': 'center','font_name':'Cambria','valign': 'vcenter','bottom': 2,'top':2, 'left':2, 'right':2})
    
        # set Header place
        title =  f"{MCCname} - {floor_name}"
        
        worksheet.merge_range(f'B{2 + offset} :L{2 + offset}',title,format_header)
        worksheet.set_row(header_start_row + offset, 45)
        
            ######## TITLE STYLE
        # title border formate
        border_format = workbook.add_format({'bottom': 2,'top':2, 'left':2, 'right':2,'font_name':'Cambria','font_size':11,'valign': 'vcenter','bold': True,'align': 'center'})
        worksheet.set_row(2 + offset, 33)
        worksheet.write(f'B{3 + offset}',"Equipment",border_format)
        worksheet.write(f'C{3 + offset}',"Equipment Tag",border_format)
        worksheet.write(f'D{3 + offset}',"LL-QTY",border_format)
        worksheet.write(f'E{3 + offset}',"FO-QTY",border_format)
        worksheet.write(f'F{3 + offset}',"Floor",border_format)
        worksheet.write(f'G{3 + offset}',"Point Description",border_format)
        worksheet.write(f'H{3 + offset}',"Points Type",border_format)
        worksheet.write(f'I{3 + offset}',"Signal Type",border_format)
        worksheet.write(f'J{3 + offset}',"Cable Type",border_format)
        worksheet.write(f'K{3 + offset}',"Functional Description",border_format)
        worksheet.write(f'L{3 + offset}',"Comments",border_format)



    ######## TABLE BORDER STYLE
    
    
        ######## CELL STYLE dashed and center (equip col to  floor col)
        cell1_format = workbook.add_format({'bottom': 4,'top':4, 'left':4, 'right':4,'font_name':'Cambria','font_size':8,'align': 'center','valign': 'vcenter'})
        left_border_format = workbook.add_format({'font_name':'Cambria','valign': 'vcenter','align': 'center','font_size':8,'right':4,'left':2,'bottom': 4,'top':4})
        right_border_format = workbook.add_format({'font_name':'Cambria','valign': 'vcenter','align': 'center','font_size':8,'right':2,'left':4,'bottom': 4,'top':4})
        bottom_border_format = workbook.add_format({'font_name':'Cambria','valign': 'vcenter','align': 'center','font_size':28,'right':4,'left':4,'bottom': 2,'top':4})
        description_cell__format = workbook.add_format({'font_name':'Cambria','align': 'left','font_size':8,'right':4,'left':2,'bottom': 4,'top':4})
    
        df_tot_row = row_number_list[count] - 2
        for row in range(df_tot_row):
         
            # row_num_list contain title and header, so - 2
            for col in range(5):
               
                if col == 0:            
                    # left border
                    worksheet.write(row + 3 + offset, col + 1, tot_df.index.values[row][col], left_border_format)
                elif col == 4:
                    # right border
                    worksheet.write(row + 3 + offset, col + 1, tot_df.index.values[row][col], right_border_format)          
                else:
                    worksheet.write(row + 3 + offset, col + 1, tot_df.index.values[row][col], cell1_format)

                # last row
                if row == df_tot_row - 1 :
                    # bottom border
                
                    worksheet.write(row + 3 + offset, col + 1,tot_df.index.values[row][col], bottom_border_format)

            #description have diff format
            worksheet.write(row + 3 + offset, 6, tot_df.index.values[row][5], description_cell__format)


        description_cell__button_format = workbook.add_format({'font_name':'Cambria','align': 'left','font_size':8,'right':2,'left':2,'bottom': 2,'top':4})
        worksheet.write(df_tot_row + 2 + offset, 6, tot_df.index.values[df_tot_row - 1][5], description_cell__button_format)
        # left bottom corner
        left_corner_format = workbook.add_format({'font_name':'Cambria','align': 'left','font_size':8,'right':4,'left':2,'bottom': 2,'top':4})
        worksheet.write(df_tot_row + 2 + offset, 1, tot_df.index.values[df_tot_row - 1][1], left_corner_format)
            




            ####### CELL STYLE dashed and left align (pd col to  comment col)
        cell2_format = workbook.add_format({'bottom': 4,'top':4, 'left':4, 'right':4,'font_name':'Cambria','font_size':8,'align': 'left'})
        left_border_alignleft_format = workbook.add_format({'font_name':'Cambria','valign': 'vcenter','align': 'left','font_size':8,'right':4,'left':2,'bottom': 4,'top':4})
        right_border__aligenleft_format = workbook.add_format({'font_name':'Cambria','valign': 'vcenter','align': 'left','font_size':8,'right':2,'left':4,'bottom': 4,'top':4})
        for row in range(df_tot_row):
            for col in range(tot_df.shape[1]):
                if col == tot_df.shape[1] - 1:
                    # right border
                    worksheet.write(row + 3 + offset, col + 7, tot_df.iloc[row,col], right_border__aligenleft_format)         
                else:
                    worksheet.write(row + 3 + offset, col + 7, tot_df.iloc[row,col], cell2_format)
                if row == df_tot_row - 1:
                    # bottom border
                    bottom_border_left_format = workbook.add_format({'font_name':'Cambria','valign': 'vcenter','align': 'left','font_size':8,'right':4,'left':4,'bottom': 2,'top':4})
                    worksheet.write(row + 3 + offset, col + 7, tot_df.iloc[row,col], bottom_border_left_format)
        # right corner
        right_corner_format = workbook.add_format({'font_name':'Cambria','align': 'left','font_size':8,'right':2,'left':4,'bottom': 2,'top':4})
        worksheet.write(df_tot_row + 2 + offset, 11, tot_df.iloc[df_tot_row - 1, 4], right_corner_format)



        # write the total BMS point

        # all bold stype
        all_bold_format = workbook.add_format({'font_name':'Cambria','align': 'center','valign': 'vcenter','font_size':10,'right':2,'left':2,'bottom': 2,'top':2,'bold': True,})
        worksheet.write(df_tot_row + 4 + offset, 2 , total_BMS_point_list_per_floor[count], all_bold_format)
        worksheet.write(df_tot_row + 4 + offset, 7 , total_BMS_point_list_per_floor[count] * (1 + spare_points), all_bold_format)

        worksheet.write(df_tot_row + 4 + offset, 1 , "Total BMS Points", all_bold_format)
        # merge range start from 0. so need to add 4 + 1
        worksheet.merge_range(f'E{df_tot_row + 5 + offset} :G{df_tot_row + 5 + offset}',f"Total BMS Points (+ {format(spare_points, '.2%')} Spare Capacity)",all_bold_format)


        # + 1 empty row + 1 total row + 1 emoty row
        offset += row_number_list[count] + 3

    worksheet.insert_image("B2", "logo.png",{"x_offset": 30, "y_offset": 10})

def createReportExcel(point_file, equip_file, project_file):
    project_name = point_file.split("-")[1]
    report_date = point_file.split("-")[2] + "-" + point_file.split("-")[3] +"-" + point_file.split("-")[4].split(".")[0]
    report_name = f'G:\\Ethos Digital\\BMS Points Generator Reports\\{project_name} - Points Schedule - {report_date}.xlsx'
    file_path = os.path.abspath(report_name)
   
    equip_sum,equip_sum_row_num, equip_sum_count_list,block_list,equip_table_row_num, tot_point_per_block_equip, tot_point_per_floor_equip = createEquipSummaryDF(equip_file)
    point_sum,point_sum_row_num ,point_sum_count_list,block_list,point_table_row_num, tot_point_per_block, tot_point_per_floor_point= createPointSummaryDF(equip_file)

    spare_point = pd.read_csv(project_file)
    spare_points = spare_point['SpareBMSPoints'][0]
    with pd.ExcelWriter(file_path, engine = 'xlsxwriter') as excel_writer:
       # project_sum.to_excel(excel_writer, sheet_name='Points Summary', startrow= 2, startcol=1)
        point_sum.to_excel(excel_writer, sheet_name='Points Summary', startrow= 2, startcol=1)
        create_point_sum_format(excel_writer,'Points Summary',point_sum_row_num, point_sum,point_sum_count_list,block_list,point_table_row_num,tot_point_per_block,tot_point_per_floor_point)
        equip_sum.to_excel(excel_writer, sheet_name='Equipment Summary', startrow= 2, startcol=1)
        create_equip_sum_format(excel_writer, 'Equipment Summary',equip_sum_row_num, equip_sum, equip_sum_count_list,block_list,equip_table_row_num,tot_point_per_block_equip, tot_point_per_floor_equip)
        # iterate each MCC create excel

        point_list,point_equip_type_set = get_points_list(point_file)
        df_equip = pd.read_csv(equip_file)
        grouped =  df_equip.groupby(['MCCLocation'])

        for MCCname, group in grouped:
            point_record_df,row_number_list, floor_name_list,total_BMS_point_list_per_floor = createPointsDF2(group, point_equip_type_set, point_list)
            create_point_list_sheet_format(point_record_df,MCCname[0],row_number_list, floor_name_list,total_BMS_point_list_per_floor,spare_points, excel_writer)
    return equip_sum, equip_sum_row_num, equip_sum_count_list, equip_table_row_num, block_list, point_sum,point_sum_row_num,point_sum_count_list,point_table_row_num, report_name

def copy_project_sum(report_name):
    file_path = os.path.abspath('G:\\Ethos Digital\\BMS Points Generator Reports\\Template\\BMS Export Template.xlsx')
 
    path1 = file_path
    xl=Dispatch("Excel.Application",pythoncom.CoInitialize())
    path2 = report_name
   


    wb_res = xl.Workbooks.Open(Filename=path1)
    wb_des = xl.Workbooks.Open(Filename=path2)

    ws1 = wb_res.Worksheets(1)
    ws1.Copy(Before=wb_des.Worksheets(1))
    

    wb_res.Close(SaveChanges=False)
    wb_des.Close(SaveChanges=True)
    xl.Quit()
   

def edit_project_summary(projectName, current_date,report_name):
   
    wb = load_workbook(report_name)
    ws = wb['Project Summary']
    projectFile = Path(f'G:\\Ethos Digital\\BMS Points Generator Reports\\Points Schedule - {projectName} - {current_date.day:02d}-{current_date.month:02d}-{current_date.year}.csv\\Projects.csv')
    dfProject = pd.read_csv(projectFile)
    ws['C2'] = dfProject.loc[0]['JobName']
    ws['C4'] = dfProject.loc[0]['JobName']
    ws['C6'] = dfProject.loc[0]['Client']
    ws['C8'] = dfProject.loc[0]['Designer']
    ws['F6'] = dfProject.loc[0]['Verifier']
    ws['F8'] = dfProject.loc[0]['BuildingType']
    
    formatted_date = current_date.strftime('%d/%m/%Y')
    ws['F2'] = formatted_date
    wb.save(report_name)

def edit_electrical_loads(qty_per_MCC,report_name):
    wb = load_workbook(report_name)


    ### add bottom border
    bottom_border = Border(left=Side(style='dashed'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))
    
    cell_border = Border(left=Side(style='dashed'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='dashed'))
    left_cell_border = Border(left=Side(style='thick'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='dashed'))

    right_cell_border = Border(left=Side(style='dashed'), 
                     right=Side(style='thick'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='dashed'))
    left_corner_cell_border= Border(left=Side(style='thick'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))
    right_corner_cell_border= Border(left=Side(style='dashed'), 
                     right=Side(style='thick'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))
    
    for key, innerdict in qty_per_MCC.items():
       
    ## change the table TITLE to MCC name
        wb.active = wb[f"{key} - Electrical Loads"]
        ws = wb.active
        ws['B2'].value = key
        count = 11
        # copy equipment name into excel and set border style 
        pattern = re.compile(r'\w+ to \w+')
        
        for equip, qty in innerdict.items():
            # initialize flag
            pattern_found = False
            if pattern.search(equip):
                # found -> "FCU-01-01 to FCU-01-50"
                # flag means value assigned different when later in the loop
                pattern_found = True
                start_val = int(equip.split(" ")[0].split("-")[2])
            for i in range(qty):
                # use flag to decide which assignment 
                if pattern_found:
                    ws[f'B{count + i}'].value = equip.split("-")[0] + "-" + equip.split("-")[1] + "-" + str(format(start_val + i,'02'))
                    
                else: 
                    ws[f'B{count + i}'].value = equip
                # column style design after each row filled
                for col in range (2,28):
                    if col == 2 :
                        ws.cell(row = count + i, column = col ).border = left_cell_border
                    elif col == 27:
                        ws.cell(row = count + i, column = col ).border = right_cell_border
                    else:
                        ws.cell(row = count + i, column = col ).border = cell_border
            ws.column_dimensions['B'].width = 25
            count += i + 1

        #### bottom border style
        for col in range(2,28):
            if col == 2 :
                        ws.cell(row = count, column = col ).border = left_corner_cell_border
            elif col == 27:
                        ws.cell(row = count, column = col ).border = right_corner_cell_border
            else:
                ws.cell(row = count, column = col).border = bottom_border

        ##### add formula of bottom graph 
        for col in range(12, 16):
            col_letter = openpyxl.utils.get_column_letter(col)
            ## park load formula
            ws[f'{col_letter}{count + 2}'].value = f'=SUM({col_letter}{11}:{col_letter}{count})' 
            # peak load formula
            ws[f'{col_letter}{count + 3}'] = f'=IF($P$14="3Ph",(({col_letter}{count + 2}*1000)/((SQRT(3))*400)),(({col_letter}{count + 2}*1000)/230))' 
            # diversified peal load amps
            ws[f'{col_letter}{count + 5}'] = f'=SUM({col_letter}{count + 2}:{col_letter}{count + 3})'     

    wb.save(report_name)

def copy_electrical_loads_title(qty_per_MCC,report_name):

    file_path = os.path.abspath('G:\\Ethos Digital\\BMS Points Generator Reports\\Template\\BMS Export Template.xlsx')
 
    path1 = file_path
 
    path2 = report_name
    xl = Dispatch("Excel.Application")


    wb_res = xl.Workbooks.Open(Filename=path1)
    wb_des = xl.Workbooks.Open(Filename=path2)

    ws_res = wb_res.Worksheets(2)
  

    for key, innerdict in qty_per_MCC.items():
        
            
    
        ws2 = wb_des.Worksheets.Add(After = wb_des.Worksheets['Project Summary'])
        ws2.Name = f"{key} - Electrical Loads"
        
        # Copy a range from the original worksheet
        # row of equip qty + offset
        start_row_sum = sum(innerdict.values()) + 12
        ws_res.Range("A1:AA10").Copy(ws2.Range("A1:AA10"))
        ws_res.Range("B43:AA49").Copy(ws2.Range(f"B{start_row_sum}:AA{start_row_sum + 6}"))


    wb_res.Close(SaveChanges=False)
    wb_des.Close(SaveChanges=True)
    xl.Quit()

def equip_summary_format(equip_sum,equip_sum_row_num,equip_sum_count_list,equip_table_row_num,block_list,report_name):
    path = report_name
    wb = load_workbook(path)
    ws = wb['Equipment Summary']
    bottom_border = Border(left=Side(style='dashed'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))
    
    cell_border = Border(left=Side(style='dashed'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='dashed'))
    left_cell_border = Border(left=Side(style='thick'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='dashed'))

    right_cell_border = Border(left=Side(style='dashed'), 
                     right=Side(style='thick'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='dashed'))
    left_corner_cell_border= Border(left=Side(style='thick'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))
    right_corner_cell_border= Border(left=Side(style='dashed'), 
                     right=Side(style='thick'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))
    
    font = Font(name='Cambria',size = 8)
    left_align = Alignment(horizontal='left', vertical='center')
    center_align = Alignment(horizontal='center', vertical='center')
    tot_col = equip_sum.shape[1] + len(equip_sum.index[0])

    offset = 0

    for i, block in enumerate(block_list):
        
    ####### 1: 1ST TABLE FORMAT 
        for row in range(equip_sum_row_num[i] -2):
            for col in range(1, tot_col  + 1):
                # start from 3rd row, 2nd col
               
                start_row = row + 4 + offset
                start_col = col + 1
                ws.cell(row = start_row, column = start_col ).font = font
                ws.cell(row = start_row, column = start_col ).alignment = center_align
                #### start from 1
                if col == 1 or col == 2:

                    ws.cell(row = start_row, column = start_col ).border = left_cell_border
                
                elif (col == tot_col) or (col == len(equip_sum.index[0])):
                    ws.cell(row = start_row, column = start_col ).border = right_cell_border
                elif (col == 2) or (col == 3):
                    ws.cell(row = start_row, column = start_col ).alignment = left_align
                else:
                    ws.cell(row = start_row, column = start_col ).border = cell_border
      
        


        ####### 2: 2ND TABLE FORMAT 
    
       
        # -2 : 1 empty row, 1 2nd title row
        bd_row_num = equip_table_row_num[i] - equip_sum_row_num[i] - 2
        
        # start from 1st table row + 6, 2nd col
      
        start_bd_row = equip_sum_row_num[i]  + 4 + offset

     
        for row2 in range(bd_row_num):
            for col2 in range(1, tot_col  + 1):    
                start_row2 = start_bd_row + row2 
                start_col2 =  col2 + 1  
                ws.cell(row = start_row2, column = start_col2 ).font = font
                ws.cell(row = start_row2, column = start_col2 ).alignment = center_align
                ####There are 2 logic need to check in one loop
                if (col2 == 2) or (col2 == 3):
                    ws.cell(row = start_row2, column = start_col2 ).alignment = left_align
                if col2 == 1 or col2 == 2:
                    ws.cell(row = start_row2, column = start_col2 ).border = left_cell_border
                elif (col2 == tot_col) or (col2 == len(equip_sum.index[0])):
                    ws.cell(row = start_row2, column = start_col2 ).border = right_cell_border 
                else:
                    ws.cell(row = start_row2, column = start_col2 ).border = cell_border
        
        
        # 3: SET BOTTOM BORDER FOR 2 TABLES
        # border under each equipment in breakdown
        start = 3 + equip_sum_row_num[i] + offset
    
        for row in equip_sum_count_list[block]:
            for col in range(1, 10):
                    start_col = col + 1
                
                    ws.cell(row = start + row, column = start_col).border = bottom_border
                    ws.cell(row = equip_sum_row_num[i] + 1 + offset , column = start_col).border = bottom_border
                    ws.cell(row = equip_sum.shape[0] + 3 , column = start_col).border = bottom_border

                    if col == 1 or col == 2:
                                # 2nd table equip bottom border
                                ws.cell(row = start + row, column = start_col ).border = left_corner_cell_border
                                ## 1st table bottom border
                                ws.cell(row = equip_sum_row_num[i] + 1 + offset , column = start_col).border = left_corner_cell_border
                                ## 2nd table
                                ws.cell(row =equip_sum.shape[0] + 3, column = start_col ).border = left_corner_cell_border
                    elif (col == 9) or (col == 4) :
                                # 1st table bottom border
                                ws.cell(row = equip_sum_row_num[i] + 1 + offset, column = start_col).border = right_corner_cell_border
                                ws.cell(row = start + row, column = start_col ).border = right_corner_cell_border
                                ## 2nd table
                                ws.cell(row = equip_sum.shape[0] + 3, column = start_col ).border = right_corner_cell_border
            start = start + row 
                
        offset += equip_table_row_num[i] + 1

    wb.save(path)
    wb.close()


def point_summary_format(point_sum,point_sum_row_num,point_sum_count_list,point_table_row_num, block_list,floor_num_dict,report_name):
    path = report_name
    wb = load_workbook(path)
    ws = wb['Points Summary']
    bottom_border = Border(left=Side(style='dashed'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))
    
    cell_border = Border(left=Side(style='dashed'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='dashed'))
    left_cell_border = Border(left=Side(style='thick'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='dashed'))

    right_cell_border = Border(left=Side(style='dashed'), 
                     right=Side(style='thick'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='dashed'))
    left_corner_cell_border= Border(left=Side(style='thick'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))
    right_corner_cell_border= Border(left=Side(style='dashed'), 
                     right=Side(style='thick'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))
    
    font = Font(name='Cambria',size = 8)
    left_align = Alignment(horizontal='left', vertical='center')
    center_align = Alignment(horizontal='center', vertical='center')
    tot_col = point_sum.shape[1] + len(point_sum.index[0])

    offset = 0
    for i, block in enumerate(block_list):

    
        ####### 1: 1ST TABLE FORMAT 
        for row in range(point_sum_row_num[i] - 2):
            for col in range(1, tot_col  + 1):
                # start from 3rd row, 2nd col
                start_row = row + 4 + offset
                start_col = col + 1
                ws.cell(row = start_row, column = start_col ).font = font
                ws.cell(row = start_row, column = start_col ).alignment = center_align
                #### start from 1
                if col == 1 or col == 2:

                    ws.cell(row = start_row, column = start_col ).border = left_cell_border
                
                elif (col == tot_col) or (col == len(point_sum.index[0])):
                    ws.cell(row = start_row, column = start_col ).border = right_cell_border
                elif (col == 2) or (col == 3):
                    ws.cell(row = start_row, column = start_col ).alignment = left_align
                else:
                    ws.cell(row = start_row, column = start_col ).border = cell_border

                
                if(ws.cell(row=start_row, column=3).value == "Floor Area"):
                 
                    ws[f'C{start_row}'] = f'Floor Area ({floor_num_dict[block]})'


        ####### 2: 2ND TABLE FORMAT 
        ### -2 because 1 empty row, 1 2nd table title
        bd_row_num = point_table_row_num[i] - point_sum_row_num[i] - 2
        
        # start from 1st table row + 1 emptyrow(begining of table) + 1 empty row(end of table) + 1: 2nd table title + 1(cell begining from 1 )
        start_bd_row =  point_sum_row_num[i] + 4 + offset
        # row start from 0. 
        for row in range(bd_row_num):
            for col in range(1, tot_col  + 1):    
                start_row = start_bd_row + row
                start_col =  col + 1  
                ws.cell(row = start_row, column = start_col ).font = font
                ws.cell(row = start_row, column = start_col ).alignment = center_align
                if (col == 2) or (col == 3):
                    ws.cell(row = start_row, column = start_col ).alignment = left_align 
                if col == 1 or col == 2:
                    ws.cell(row = start_row, column = start_col ).border = left_cell_border
                elif (col == tot_col) or (col == len(point_sum.index[0])):
                    ws.cell(row = start_row, column = start_col ).border = right_cell_border
                
                else:
                    ws.cell(row = start_row, column = start_col ).border = cell_border
        
        # 3: SET BOTTOM BORDER FOR 2 TABLES
        # border under each equipment in breakdown
        #4: 1 emptyrow(begining of table) + 1 empty row(end of table) + 1: 2nd table title 
        start = point_sum_row_num[i] + 3 + offset
        # row start from boundary
        for row in point_sum_count_list[block]:
            for col in range(2,9):
                   
                    ws.cell(row = start + row, column = col).border = bottom_border
                    ws.cell(row = point_sum_row_num[i] + 1 , column = col).border = bottom_border
                    ws.cell(row = point_sum.shape[0] + 3 , column = col).border = bottom_border

                    if col == 2 or col == 3:
                                # 2nd table equip bottom border
                                ws.cell(row = start + row, column = col ).border = left_corner_cell_border
                                ## 1st table bottom border
                                ws.cell(row = point_sum_row_num[i] + 1 + offset , column = col).border = left_corner_cell_border
                                ## 2nd table
                                ws.cell(row =point_sum.shape[0] + 3, column = col ).border = left_corner_cell_border
                    elif (col == 8) or (col == 5) :
                                ## 1st table bottom border
                                ws.cell(row = point_sum_row_num[i] + 1 + offset , column = col).border = right_corner_cell_border

                                ws.cell(row = start + row, column = col ).border = right_corner_cell_border
                                ## 2nd table
                                ws.cell(row = point_sum.shape[0] + 3, column = col ).border = right_corner_cell_border
            start = start + row 
        offset += point_table_row_num[i] + 1
        
            

    wb.save(path)
    wb.close()

def apply_border_again(block_list,point_sum_row_num,point_sum):
    path = 'G:\\Ethos Digital\\BMS Points Generator Reports\\BMS Final Export.xlsx'
    wb = load_workbook(path)
    ws = wb['Points Summary']
    bottom_border = Border(left=Side(style='dashed'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))

    left_corner_cell_border= Border(left=Side(style='thick'), 
                     right=Side(style='dashed'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))
    right_corner_cell_border= Border(left=Side(style='dashed'), 
                     right=Side(style='thick'), 
                     top=Side(style='dashed'), 
                     bottom=Side(style='thick'))

      
    wb.save(path)
    wb.close()


def get_final_report(projectName):
    
    current_date = date.today()
    #current_date = datetime.datetime(2023, 11, 15)
    equip_sum, equip_sum_row_num, equip_sum_count_list, equip_table_row_num, block_list,point_sum, point_sum_row_num,point_sum_count_list, point_table_row_num,report_name = createReportExcel(
                      f'G:\\Ethos Digital\\BMS Points Generator Reports\\Points Schedule - {projectName} - {current_date.day:02d}-{current_date.month:02d}-{current_date.year}.csv\\Points.csv',
                      f'G:\\Ethos Digital\\BMS Points Generator Reports\\Points Schedule - {projectName} - {current_date.day:02d}-{current_date.month:02d}-{current_date.year}.csv\\Equipment.csv',
                      f'G:\\Ethos Digital\\BMS Points Generator Reports\\Points Schedule - {projectName} - {current_date.day:02d}-{current_date.month:02d}-{current_date.year}.csv\\Projects.csv')
    
    #openpyxl to format the cell font
    ## need get equip per MCC {'MMC -A' : {"fcu" : 6, "ahu":9}}
    ### num of equipment = row of that equip
    _, qty_per_MCC ,_,floor_num_dict= get_equipments_list(f'G:\\Ethos Digital\\BMS Points Generator Reports\\Points Schedule - {projectName} - {current_date.day:02d}-{current_date.month:02d}-{current_date.year}.csv\\Equipment.csv') 
    point_summary_format(point_sum,point_sum_row_num,point_sum_count_list,point_table_row_num, block_list,floor_num_dict,report_name)
    equip_summary_format(equip_sum,equip_sum_row_num,equip_sum_count_list,equip_table_row_num, block_list,report_name)
   
    
    
    copy_project_sum(report_name)
    ##### EDIT copied sheet(Project summary and electrical loads)  
    edit_project_summary(projectName,current_date,report_name)


  
    copy_electrical_loads_title(qty_per_MCC,report_name)
    edit_electrical_loads(qty_per_MCC,report_name)
    
    #openpyxl didn't preserve all the original information when it loaded the workbook, some of the original formatting might be lost, so we need to reapply missing feature
    apply_border_again(block_list,point_sum_row_num,point_sum)


#if __name__=="__main__":
    #get_final_report('Camden Yard Script Test')
   # get_final_report('Linkedin 4 Wilton Park')