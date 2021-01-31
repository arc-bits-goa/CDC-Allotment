import pandas as pd
import numpy as np
import xlsxwriter 

global max_seats_dict, registered_seats_dict

#Create dicts of option:capacity and option:registered (registered is 0 now)
max_seats_df = pd.read_excel('combination.xlsx', sheet_name='Seats')
max_seats_dict = {} #eg. 'CS4': 42
registered_seats_dict = {} #eg. 'CS4': 0
for i in range(len(max_seats_df['option'])): 
    max_seats_dict[max_seats_df['option'][i].strip()]=max_seats_df['seats'][i]
    registered_seats_dict[max_seats_df['option'][i].strip()] = 0


#Create DF of ID No | PR No. DF sorted on PR NO.
pr_df_2019 = pd.read_excel('GOA PR.xlsx', sheet_name='2019')
pr_df_2019 = pr_df_2019[['CAMPUS_ID', 'PR NO.']]
pr_df_2019 = pr_df_2019.sort_values(by=['PR NO.'])
pr_df_2019 = pr_df_2019.reset_index(drop = True)


#CS Responses:
global cs_lab_dict #id_no: preference list
#eg. 2019A7PS0036G:['CS2', 'CS3', 'CS5', 'CS4', 'CS1']
#create such list for each response
cs_lab_dict = {}
responses_df = pd.read_excel('./Responses/CS Form (Responses).xlsx')
no_of_options = 5
no_of_cols_to_skip = 3
for i in range(len(responses_df['ID Number'])):
    current_id = responses_df['ID Number'][i].strip().upper()
    current_pref_list = ['','','','','']
    current_row = responses_df.iloc[i]
    for i in range(no_of_options):
        pref_no = current_row[i+no_of_cols_to_skip][0]
        list_index = int(pref_no) - 1
        current_option = 'CS' + str(i+1)
        current_pref_list[list_index] = current_option
    cs_lab_dict[current_id] = current_pref_list
global cs_lab_allot
cs_lab_allot = {}
#eg. 2019A7PS0036G:'CS2'
def allotCS(current_id):
    global registered_seats_dict, max_seats_dict,cs_lab_allot,cs_lab_dict
    if current_id not in cs_lab_dict: #don't allot if ID not in reponses
        return
    current_pref_list = cs_lab_dict[current_id]
    for i in range(len(current_pref_list)):
        current_option = current_pref_list[i]
        #check for seats availability for current (i+1)th preference
        if registered_seats_dict[current_option] < max_seats_dict[current_option]:
            cs_lab_allot[current_id] = current_option
            registered_seats_dict[current_option] += 1
            return



#ECE Responses:
global ece_lab_dict #id_no: preference list
#eg. 2019AAPS0036G:['ECE2','ECE1','ECE3']
#create such list for each response
ece_lab_dict = {}
responses_df = pd.read_excel('./Responses/ECE Form (Responses).xlsx')
no_of_options = 3
no_of_cols_to_skip = 3
for i in range(len(responses_df['ID Number'])):
    current_id = responses_df['ID Number'][i].strip().upper()
    current_pref_list = ['','','']
    current_row = responses_df.iloc[i]
    for i in range(no_of_options):
        pref_no = current_row[i+no_of_cols_to_skip][0]
        list_index = int(pref_no) - 1
        current_option = 'ECE' + str(i+1)
        current_pref_list[list_index] = current_option
    ece_lab_dict[current_id] = current_pref_list
global ece_lab_allot
ece_lab_allot = {}
#eg. 2019AAPS0036G:'ECE2'
def allotECE_lab(current_id):
    global registered_seats_dict, max_seats_dict,ece_lab_allot,ece_lab_dict
    if current_id not in ece_lab_dict: #don't allot if ID not in reponses
        return
    current_pref_list = ece_lab_dict[current_id]
    for i in range(len(current_pref_list)):
        current_option = current_pref_list[i]
        #check for seats availability for current (i+1)th preference
        if registered_seats_dict[current_option] < max_seats_dict[current_option]:
            ece_lab_allot[current_id] = current_option
            registered_seats_dict[current_option] += 1
            return




#EEE Responses:
global eee_lab_dict #id_no: preference list
#eg. 2019AAPS0036G:['EEE2','ECE1',...]
#create such list for each response
eee_lab_dict = {}
responses_df = pd.read_excel('./Responses/EEE form (Responses).xlsx')
no_of_options = 6
no_of_cols_to_skip = 3
for i in range(len(responses_df['ID Number'])):
    current_id = responses_df['ID Number'][i].strip().upper()
    current_pref_list = ['','','','','','']
    current_row = responses_df.iloc[i]
    for i in range(no_of_options):
        pref_no = current_row[i+no_of_cols_to_skip][0]
        list_index = int(pref_no) - 1
        current_option = 'EEE' + str(i+1)
        current_pref_list[list_index] = current_option
    eee_lab_dict[current_id] = current_pref_list
global eee_lab_allot
eee_lab_allot = {}
#eg. 2019AAPS0036G:'EEE2'
def allotEEE_lab(current_id):
    global registered_seats_dict, max_seats_dict,eee_lab_allot,eee_lab_dict
    if current_id not in eee_lab_dict: #don't allot if ID not in reponses
        return
    current_pref_list = eee_lab_dict[current_id]
    for i in range(len(current_pref_list)):
        current_option = current_pref_list[i]
        #check for seats availability for current (i+1)th preference
        if registered_seats_dict[current_option] < max_seats_dict[current_option]:
            eee_lab_allot[current_id] = current_option
            registered_seats_dict[current_option] += 1
            return



#ENI Responses:
global eni_lab_dict #id_no: preference list
#eg. 2019AAPS0036G:['ENI2','ENI1',...]
#create such list for each response
eni_lab_dict = {}
responses_df = pd.read_excel('./Responses/ENI Form (Responses).xlsx')
no_of_options = 6
no_of_cols_to_skip = 3
for i in range(len(responses_df['ID Number'])):
    current_id = responses_df['ID Number'][i].strip().upper()
    current_pref_list = ['','','','','','']
    current_row = responses_df.iloc[i]
    for i in range(no_of_options):
        pref_no = current_row[i+no_of_cols_to_skip][0]
        list_index = int(pref_no) - 1
        current_option = 'ENI' + str(i+1)
        current_pref_list[list_index] = current_option
    eni_lab_dict[current_id] = current_pref_list
global eni_lab_allot
eni_lab_allot = {}
#eg. 2019AAPS0036G:'ENI2'
def allotENI_lab(current_id):
    global registered_seats_dict, max_seats_dict,eni_lab_allot,eni_lab_dict
    if current_id not in eni_lab_dict: #don't allot if ID not in reponses
        return
    current_pref_list = eni_lab_dict[current_id]
    for i in range(len(current_pref_list)):
        current_option = current_pref_list[i]
        #check for seats availability for current (i+1)th preference
        if registered_seats_dict[current_option] < max_seats_dict[current_option]:
            eni_lab_allot[current_id] = current_option
            registered_seats_dict[current_option] += 1
            return


#iterate through sorted PR NO. list and allot
for i in range(len(pr_df_2019['PR NO.'])):
    current_id = pr_df_2019['CAMPUS_ID'][i].strip().upper()
    current_branch = current_id[4:6]
    if current_branch=='A7':
        allotCS(current_id)
    elif current_branch=='AA':
        allotECE_lab(current_id)
    elif current_branch=='A8':
        allotENI_lab(current_id)
    elif current_branch=='A3':
        allotEEE_lab(current_id)

class_map = pd.read_excel('class_emted.xlsx', sheet_name='ClassNumb')
course_map = pd.read_excel('class_emted.xlsx', sheet_name='ClassCode')
time_map = pd.read_excel('class_emted.xlsx', sheet_name='ClassTime')

course_dict={}
class_dict={}
time_dict={}
for i in range(course_map.shape[0]):
    course_dict[course_map['Course'][i]]=[course_map['o1'][i],course_map['o2'][i],course_map['o3'][i],course_map['o4'][i]]
for i in range(class_map.shape[0]):
    class_dict[class_map['Course'][i]]=[class_map['o1'][i],class_map['o2'][i],class_map['o3'][i],class_map['o4'][i]]
for i in range(time_map.shape[0]):
    time_dict[time_map['Course'][i]]=[time_map['o1'][i],time_map['o2'][i],time_map['o3'][i],time_map['o4'][i]]
    
id_to_name={}
pr_df_2019 = pd.read_excel('GOA PR.xlsx', sheet_name='2019')
for i in range(pr_df_2019.shape[0]):
    current_id = pr_df_2019['CAMPUS_ID'][i].strip()
    current_name=pr_df_2019['NAME'][i].strip()
    id_to_name[current_id]=current_name

def solve_CS(curr_id ,labs, lab_allot):
    for i in range(3):
        try:
            p=course_dict[lab_allot[curr_id]]
            l=[curr_id,p[i],class_dict[lab_allot[curr_id]][i],time_dict[lab_allot[curr_id]][i]]
            labs.append(l)
            print(l)
        except:
            print(curr_id)
def solve_REST(curr_id ,labs, lab_allot):
    for i in range(4):
        try:
            p=course_dict[lab_allot[curr_id]]
            l=[curr_id,p[i],class_dict[lab_allot[curr_id]][i],time_dict[lab_allot[curr_id]][i]]
            labs.append(l)
            print(l)
        except:
            print(curr_id)
        #labs.append(l)
    
CS_labs=[]
ECE_labs=[]
EEE_labs=[]
ENI_labs=[]

for i in range(len(pr_df_2019['PR NO.'])):
    current_id = pr_df_2019['CAMPUS_ID'][i].strip()
    current_branch = current_id[4:6]
    if current_branch=='A3' or current_branch=='a3':
        solve_REST(current_id,EEE_labs,eee_lab_allot)
    elif current_branch=='A8' or current_branch=='a8':
        solve_REST(current_id,ENI_labs,eni_lab_allot)
    elif current_branch=='AA' or current_branch=='aa':
        solve_REST(current_id,ECE_labs,ece_lab_allot)
        
import xlsxwriter 
def make_file_cs(course,CS_labs):
    workbook = xlsxwriter.Workbook('output_cs.xlsx')
    workbook.add_worksheet(course)
    cell_format = workbook.add_format({'bold': True, 'align': 'center'})
    cell_format3 =workbook.add_format ({'align': 'center'})
    worksheet=workbook.get_worksheet_by_name(course)
    worksheet.set_column(0,200, 50)
    worksheet.write(0,0,'StudentID',cell_format)
    worksheet.write(0,1,'StudentName',cell_format)
    worksheet.write(0,2,'Section',cell_format)
    worksheet.write(0,3,'ClassNbr',cell_format)
    worksheet.write(0,4,'Time',cell_format)
    j=1
    for i in CS_labs:
            worksheet.write(j,0,i[0],cell_format3)
            worksheet.write(j,1,id_to_name[i[0]],cell_format3)
            worksheet.write(j,2,i[1],cell_format3)
            worksheet.write(j,3,i[2],cell_format3)
            worksheet.write(j,4,i[3],cell_format3)
            j=j+1
            #print(i)
    workbook.close()
    
def make_file_rest(course,rest_labs):
    workbook = xlsxwriter.Workbook('output_emted_'+course+'.xlsx')
    workbook.add_worksheet(course)
    cell_format = workbook.add_format({'bold': True, 'align': 'center'})
    cell_format3 =workbook.add_format ({'align': 'center'})
    worksheet=workbook.get_worksheet_by_name(course)
    worksheet.set_column(0,200, 50)
    worksheet.write(0,0,'StudentID',cell_format)
    worksheet.write(0,1,'StudentName',cell_format)
    worksheet.write(0,2,'Section',cell_format)
    worksheet.write(0,3,'ClassNbr',cell_format)
    worksheet.write(0,4,'Time',cell_format)
    j=1
    for i in rest_labs:
            worksheet.write(j,0,i[0],cell_format3)
            worksheet.write(j,1,id_to_name[i[0]],cell_format3)
            worksheet.write(j,2,i[1],cell_format3)
            worksheet.write(j,3,i[2],cell_format3)
            worksheet.write(j,4,i[3],cell_format3)
            j=j+1
            #print(i)
    workbook.close()

# make_file_cs('CS',CS_labs)
make_file_rest('EEE',EEE_labs)
make_file_rest('ECE',ECE_labs)
make_file_rest('ENI',ENI_labs)
