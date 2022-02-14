# -*- coding: utf-8 -*-
"""
Created on Sun Jul 18 08:19:22 2021

@author: Fu
"""


def remove_header(dataframe):
    dataframe=dataframe[dataframe['station'] != "station"]
    dataframe.reset_index(drop=True, inplace=True)
    return dataframe



def find_machine_short_name(machine_table,machine_long_name):
    return machine_table[machine_table['long_name'] == machine_long_name]['short_name'].values[0]



def filter_machine(dataframe,machine_name):
    return dataframe[dataframe['station'].str.contains(machine_name)]


import pandas as pd
import matplotlib.pyplot as plt
import os
import numpy as np
import xlsxwriter
from io import BytesIO
import sys



#data_type_mode = int(input('Data tail?\n1.CSV\n2.xlsx\n'))
data_type_mode = 1
print("Date type: CSV")
try:
    if data_type_mode == 1:
        def read_csv(filename):
            file_read = filename+'.CSV'
            dataframe = pd.read_csv(file_read,error_bad_lines = False)
            return remove_header(dataframe)
    if data_type_mode == 2:
        def read_xlsx(filename):
            file_read = filename+'.xlsx'
            dataframe = pd.read_excel(file_read)
            return remove_header(dataframe)
except:
    print('Check data type again!')

#=============================================================================================================================
#Intial setting




path = './data/'
master_path = './master_data/'
all_path = './all_data/'
tester_color = ['blue']
master_color = ['red']
box_histor_color = ['cyan','salmon','chartreuse','gray','tan','lightblue','coral','ivory']
line_histor_color = ['indigo','crimson','green','fuchsia','brown','tomato','azure','sienna']

machine_redundant = 3 #Số lượng ký tự thừa ở cuối tên máy. Vi du: VCMH8902T01 thì thừa ở đây là T01, hoặc T02,... vậy điền là 3

master_machine_redundant = 1 # Tương tự như trên nhưng cho master

#=================================================================================================================================

# # Read setting file


setting_df = pd.read_excel('Setting.xlsx',header= None)


machine_table = pd.read_excel('Machine_info.xlsx')

file_name = setting_df.iloc[0,1]

plotting_mode = int(input('Plotting mode?\n1.Machine-Machine\n2.Machine-Golden\n'))

if plotting_mode == 2:
    log_df = read_csv(path+file_name)
    master_df = read_csv(master_path+file_name)
    all_machine = log_df['station'].drop_duplicates().apply(lambda x: x[:-1*machine_redundant]).to_list()
    master_machine = master_df['station'].drop_duplicates().apply(lambda x: x[:-1*master_machine_redundant]).values[0]


    # # Data processing

    # ### Pick machine first

    master_machine_short_name = find_machine_short_name(machine_table,master_machine)

    for machine_count in range(len(all_machine)):
        machine_pick = all_machine[machine_count]
        
         
        tester_machine_short_name = find_machine_short_name(machine_table,machine_pick)
        print("Start plotting: {}".format(tester_machine_short_name))
        
        
        
        workbook = xlsxwriter.Workbook('./result/'+tester_machine_short_name+'.xlsx')
    
    
    # ### Pick log
    
    
    
        for i in range(0,len(setting_df),6):
    
        
        
            file_name = setting_df.iloc[i,1]
            print('Log: {}'.format(file_name))
            item = setting_df.iloc[i+1,1:].dropna().to_list()
            
            warning_spec = setting_df.iloc[i+5,1:].to_list()
    
            log_df = read_csv(path+file_name)
            
            log_df = filter_machine(log_df,machine_pick)
            
            master_log_df = read_csv(master_path+file_name)
            
            worksheet = workbook.add_worksheet(file_name)
            worksheet.set_landscape()
            worksheet.set_paper(9)
            worksheet.set_margins(left=0.25, right=0.25, top=0.75, bottom=0.75)
            
            cell_format1 = workbook.add_format()
            cell_format1.set_bg_color('lime')
            cell_format2=  workbook.add_format()
            cell_format2.set_bg_color('yellow')
            cell_format3 = workbook.add_format()
            cell_format3.set_bg_color('orange')
            cell_format4 = workbook.add_format()
            cell_format4.set_bg_color('red')
            
            row = 0
            col = 0
            col_data = 12
        
            for j in range(len(item)):
    
                
            
                print('Item: {}'.format(item[j]))
                
                
    
            
                log_df[item[j]] = pd.to_numeric(log_df[item[j]])
            
    
            
                master_log_df[item[j]] = pd.to_numeric(master_log_df[item[j]])
    
            
            
                sockets = log_df['site'].sort_values().drop_duplicates().to_list()
    
            
                box_values = []
                for socket in sockets:
                    temp = log_df[log_df['site'] == socket][item[j]].to_list()
                    box_values.append(temp)
    
                sockets_master = master_log_df['site'].sort_values().drop_duplicates().to_list()
        
                box_values_master = []
                for socket in sockets_master:
                    temp = master_log_df[master_log_df['site'] == socket][item[j]].to_list()
                    box_values_master.append(temp)
                    
                    
                # Tinh toan gia tri  delta va ghi vao file ------------------------------------------------------------------------------
                check_warning = 0
                
                if warning_spec[j] != 'na':
                    warning = warning_spec[j].split(',')
                    warning =list(map(float,warning))
                    check_warning = 1
                
                worksheet.write(row,col_data, 'Socket')
                worksheet.write(row+1,col_data, 'Median tester')
                worksheet.write(row+2,col_data, 'Median master')
                worksheet.write(row+3,col_data, 'Delta')
                worksheet.write(row+4,col_data, 'delta = master_value - tester_value')    

    
                col_temp = col_data+1
                dem=0
                # try:
                for  socket_value,socket_master_value in zip(box_values,box_values_master):
                    
                    worksheet.write(row,col_temp,str(sockets[dem]))
                    worksheet.write(row+1,col_temp,round(np.median(socket_value),4))
                    worksheet.write(row+2,col_temp,round(np.median(socket_master_value),4))
                    try:
                        temp_delta = np.median(socket_master_value)-np.median(socket_value)
                        if check_warning == 1:
                            if abs(temp_delta) < warning[0]:
                                worksheet.write(row+3,col_temp,round(temp_delta,4),cell_format1)
                            if warning[0] < abs(temp_delta) < warning[1]:
                                worksheet.write(row+3,col_temp,round(temp_delta,4),cell_format2)
                            if warning[1] < abs(temp_delta) < warning[2]:
                                worksheet.write(row+3,col_temp,round(temp_delta,4),cell_format3)
                            if abs(temp_delta) > warning[2]:
                                worksheet.write(row+3,col_temp,round(temp_delta,4),cell_format4)   
                        else:
                            worksheet.write(row+3,col_temp,round(temp_delta,4))
                        
                    except:
                        worksheet.write(row,col_temp, 'NA')
                    
                    col_temp += 1
                    dem += 1
                    
                # worksheet.write(row+3,col,round(np.median(socket_master_value)-np.median(socket_value),4))
                
                # except:
                #     print("Master hoac tester bi thieu du lieu socket")
                
                
    
                total_data = box_values + box_values_master
                total_socket_tester = []
                total_socket_master = []
                for socket in  sockets:
                    total_socket_tester.append('T'+str(socket))
                for socket in  sockets_master:
                    total_socket_master.append('M'+str(socket))
                total_socket = total_socket_tester + total_socket_master
            
            
            # # Vẽ đồ thị box-plot ----------------------------------------------------------------------------------------------------
    
            
            
                colors = tester_color*len(total_socket_tester)+master_color*len(total_socket_master)
            
    
            
            
                fig, ax = plt.subplots(nrows=1, ncols=1, figsize=(9, 4.5))
                
                bplot = ax.boxplot(total_data,
                                    vert=True,  # vertical box alignment
                                    patch_artist=True,  # fill with color
                                    labels=total_socket)  # will be used to label x-ticks
                #ax.set_title('Test item: +item[j]+'\nTester: '+tester_machine_short_name+)
                ax.set_title('Test item: {}\nTester: {}   Master: {} '.format(item[j],tester_machine_short_name,master_machine_short_name))
                fig.tight_layout()             
                        
                for patch, color in zip(bplot['boxes'], colors):
                    patch.set_facecolor(color)
                
                for i in range(len(total_socket)):
                    ax.text(i+1.35,np.median(total_data[i]),round(np.median(total_data[i]),3),fontsize=9,rotation = 'vertical')
            
            
            # # Vẽ đồ thị historgram
    
            
            
                # fig1, ax1 = plt.subplots(nrows=1, ncols=1, figsize=(5, 3.5))
                # fig2, ax2 = plt.subplots(nrows=1, ncols=1, figsize=(5, 3.5))
                # for k,socket in enumerate(sockets):
                # #for k in range(3):
                #     n, bins, patches = ax1.hist(box_values[k],bins = 30,label=str(socket), alpha=.6, edgecolor='black',color=box_histor_color[k])
                #     sigma = np.std(box_values[k])
                #     mu = np.average(box_values[k])
                #     if k%3 == 1:
                #         line_style = '--'
                #     elif k%3 == 2:
                #         line_style = '-'
                #     else:
                #         line_style = '-.'
                #     y = ((1 / (np.sqrt(2 * np.pi) * sigma)) *np.exp(-0.5 * (1 / sigma * (bins - mu))**2))
                #     #ax1.plot(bins, y, line_style,label =str(socket),lw=2.3)
                #     ax2.plot(bins, y, line_style,label =str(socket),lw=1.5)
                # ax2.set_title('Socket comparasion')
                # fig2.legend(bbox_to_anchor =(0.97,0.895),ncol = 2, fontsize = 'x-small')
                # fig2.tight_layout()
            
            
            # # Tạo báo cáo
    
            
            
                
                for m,plot in enumerate([fig]):
                    imgdata = BytesIO()
                    plot.savefig(imgdata, format="jpg",dpi = 150)
                    worksheet.insert_image(row,col,"",{'image_data': imgdata,'x_scale': 0.75, 'y_scale': 0.75})
                plt.close(fig)
                
                row += 20
        try:    
            workbook.close()
        except:
            if str(input('Dong file excel dang chay! va press OK: ')) == 'OK':
                workbook.close()
            else:
                print('Chua dong file ma da an OK hoac go sai chu OK roi. The thi chiu')
                         
        
elif plotting_mode == 1:
    log_df = read_csv(all_path+file_name)
    
    master_df = read_csv(master_path+file_name)
    
    all_machine = log_df['station'].drop_duplicates().apply(lambda x: x[:-1*machine_redundant]).to_list()
    
    master_machine = master_df['station'].drop_duplicates().apply(lambda x: x[:-1*master_machine_redundant]).values[0]
    
    master_machine_short_name = find_machine_short_name(machine_table,master_machine)
    
    workbook = xlsxwriter.Workbook('./result/Summary.xlsx',{'nan_inf_to_errors': True})
    
    for i in range(0,len(setting_df),5):
    
    
        file_name = setting_df.iloc[i,1]
        print('Log: {}'.format(file_name))
        item = setting_df.iloc[i+1,1:].dropna().to_list()

        
        log_df = read_csv(all_path+file_name)
        
        for j in range(len(item)):
    
            print('Item: {}'.format(item[j]))
            
            log_df[item[j]] = pd.to_numeric(log_df[item[j]])
            
            tester_value = []
            tester_machine =[]
            
            master_value = []
        
            for machine in all_machine:
                golden_flag = 0
                
                if machine == master_machine:
                    golden_flag = 1
                
                log_df_temp = filter_machine(log_df,machine)
                
                temp = log_df_temp[item[j]].to_list()
                
                tester_machine_short_name = find_machine_short_name(machine_table,machine)
                
                if golden_flag == 0:
                    tester_value.append(temp)
                    tester_machine.append(tester_machine_short_name)
                else:
                    master_value.append(temp)
            
            total_data = tester_value + master_value
            
            total_machine = tester_machine + [master_machine_short_name]
            
            
            # Ve box-plot:
            colors = tester_color*len(tester_machine)+master_color
            
            fig, ax = plt.subplots(nrows=1, ncols=1, figsize=(9, 4.5))
    
            bplot = ax.boxplot(total_data,
                                vert=True,  # vertical box alignment
                                patch_artist=True,  # fill with color
                                labels=total_machine)  # will be used to label x-ticks
            ax.set_title('Test item: {}\nTester: All   Master: {} '.format(item[j],master_machine_short_name))
            ax.tick_params(axis='x', labelrotation=90)
            fig.tight_layout()             
            
            for patch, color in zip(bplot['boxes'], colors):
                patch.set_facecolor(color)
              
            for k in range(len(total_machine)):
                ax.text(k+1.35,np.median(total_data[k]),round(np.median(total_data[k]),4),fontsize=9,rotation = 'vertical')
            
            #Tạo báo cáo :
            worksheet = workbook.add_worksheet(item[j])
            worksheet.set_landscape()
            worksheet.set_paper(9)
            worksheet.set_margins(left=0.25, right=0.25, top=0.75, bottom=0.75)
            
            imgdata = BytesIO()
            fig.savefig(imgdata, format="jpg",dpi = 150)
            
            row = 0  
            col = 0
            worksheet.insert_image(row,col,"",{'image_data': imgdata,'x_scale': 0.75, 'y_scale': 0.75})
            worksheet.write('A18', 'Machine descriptive statistics')
            worksheet.write('A19', 'Machine')
            worksheet.write('A20', 'Median')
            worksheet.write('A21', 'Mean')
            worksheet.write('A22', 'Max')
            worksheet.write('A23', 'Min')
            worksheet.write('A24', 'Delta')
            
            
            worksheet.write('A26', 'delta = master_value - tester_value')
            
            master_median = np.median(master_value)
            
            worksheet.write('B19', master_machine_short_name+' (Golden)')
            try:
                worksheet.write('B20', round(master_median,4))
            except:
                worksheet.write('B20', 'NA')
            try:
                worksheet.write('B21', round(np.mean(master_value),4))
            except:
                worksheet.write('B21', 'NA')
            try:
                worksheet.write('B22', round(np.max(master_value),4))
            except:
                worksheet.write('B22', 'NA')
            try:
                worksheet.write('B23', round(np.min(master_value),4))
            except:
                worksheet.write('B23', 'NA')
            
            worksheet.write('B24', 0)
            
            
            row = 18
            col = 2
            for machine_num in range(len(tester_machine)):
                worksheet.write(row,col,tester_machine[machine_num])
                tester_median = np.median(tester_value[machine_num])
                try:
                    worksheet.write(row+1,col,round(tester_median,4))
                except:
                    worksheet.write(row+1,col, 'NA')
                try:
                    worksheet.write(row+2,col,round(np.mean(tester_value[machine_num]),4))
                except:
                    worksheet.write(row+2,col, 'NA')
                try:
                    worksheet.write(row+3,col,round(np.max(tester_value[machine_num]),4))
                except:
                    worksheet.write(row+3,col, 'NA')
                try:
                    worksheet.write(row+4,col,round(np.min(tester_value[machine_num]),4))
                except:
                    worksheet.write(row+4,col, 'NA')
                try:
                    worksheet.write(row+5,col,round(master_median -tester_median ,4))
                except:
                    worksheet.write(row+5,col, 'NA')

                col = col + 1
                
                    
    workbook.close()         
    
    
    

else:
    print("WRONG INPUT")
    sys.exit(0)

print("Xong film!")

sys.exit(0)