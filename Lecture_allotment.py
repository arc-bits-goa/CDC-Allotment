import pandas as pd
import numpy as np


global max_seats_dict, registered_seats_dict

#Create dicts of option:capacity and option:registered (registered is 0 now)
max_seats_df = pd.read_excel('combination.xlsx', sheet_name='seat_lec')
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


global em_pref_dict, dd_pref_list, em_allot_dict, dd_allot_dict
em_pref_dict = {} #2019AAPS0205G: ['EM2','EM1']
dd_pref_dict = {} #2019AAPS0205G: ['DD3','DD1',..]
em_allot_dict = {} #2019AAPS0205G: 'EM2_ECE'
dd_allot_dict = {} #2019AAPS0205G: 'DD3_ECE'

def fill_pref_dict(file_path, no_of_cols_to_skip_em, no_of_cols_to_skip_dd): 
    global em_pref_dict, dd_pref_list
    responses_df = pd.read_excel(file_path)
    no_of_options_em = 2
    no_of_options_dd = 4
    for i in range(len(responses_df['ID Number'])):
        current_id = responses_df['ID Number'][i].strip().upper()
        current_pref_list_em = ['','']
        current_pref_list_dd = ['','','','']
        current_row = responses_df.iloc[i]
        #em pref list:
        for i in range(no_of_options_em):
            pref_no = current_row[i+no_of_cols_to_skip_em][0]
            list_index = int(pref_no) - 1
            current_option = 'EM' + str(i+1)
            current_pref_list_em[list_index] = current_option
        em_pref_dict[current_id] = current_pref_list_em
        #dd pref list:
        for i in range(no_of_options_dd):
            pref_no = current_row[i+no_of_cols_to_skip_dd][0]
            list_index = int(pref_no) - 1
            current_option = 'DD' + str(i+1)
            current_pref_list_dd[list_index] = current_option
        dd_pref_dict[current_id] = current_pref_list_dd

fill_pref_dict('./Responses/ECE Form (Responses).xlsx', 11, 13)
fill_pref_dict('./Responses/EEE Form (Responses).xlsx', 14, 16)
fill_pref_dict('./Responses/ENI Form (Responses).xlsx', 14, 16)


def allot_lecture(current_id):
    global registered_seats_dict, max_seats_dict, em_allot_dict, em_pref_dict, dd_allot_dict, dd_pref_dict
    if current_id not in em_pref_dict: #don't allot if ID not in reponses
        return
    current_pref_list_em = em_pref_dict[current_id]
    current_branch = current_id[4:6]
    if current_branch == 'AA':
        current_branch = '_ECE'
    elif current_branch == 'A3':
        current_branch = '_EEE'
    elif current_branch == 'A8':
        current_branch = '_ENI'
    #allot em lecture
    for i in range(len(current_pref_list_em)):
        current_option = current_pref_list_em[i] + current_branch
        #check for seats availability for current (i+1)th preference
        if registered_seats_dict[current_option] < max_seats_dict[current_option]:
            em_allot_dict[current_id] = current_option
            registered_seats_dict[current_option] += 1
            break
    #allot dd lecture
    current_pref_list_dd = dd_pref_dict[current_id]
    for i in range(len(current_pref_list_dd)):
        current_option = current_pref_list_dd[i] + current_branch
        #check for seats availability for current (i+1)th preference
        if registered_seats_dict[current_option] < max_seats_dict[current_option]:
            dd_allot_dict[current_id] = current_option
            registered_seats_dict[current_option] += 1
            break


#iterate through sorted PR NO. list and allot lectures
for i in range(len(pr_df_2019['PR NO.'])):
    current_id = pr_df_2019['CAMPUS_ID'][i].strip().upper()
    current_branch = current_id[4:6]
    if current_branch=='AA' or current_branch=='A8' or current_branch=='A3':
        allot_lecture(current_id)


#generate temporary output excel files
combined_output_df = pd.DataFrame(columns=['ID Number', 'Allotted Option'])
em_output_df = pd.DataFrame(columns=['ID Number', 'Allotted Option'])
for id_no, allotted_option in em_allot_dict.items():
    combined_output_df = combined_output_df.append(pd.Series([id_no, allotted_option], index = combined_output_df.columns), ignore_index=True)
    em_output_df = em_output_df.append(pd.Series([id_no, allotted_option], index = em_output_df.columns), ignore_index=True)
em_output_df.to_excel('Temporary EM allotments.xlsx', index=False)

dd_output_df = pd.DataFrame(columns=['ID Number', 'Allotted Option'])
for id_no, allotted_option in dd_allot_dict.items():
    combined_output_df = combined_output_df.append(pd.Series([id_no, allotted_option], index = combined_output_df.columns), ignore_index=True)
    dd_output_df = dd_output_df.append(pd.Series([id_no, allotted_option], index = dd_output_df.columns), ignore_index=True)
dd_output_df.to_excel('Temporary DD allotments.xlsx', index=False)

combined_output_df.to_excel('Temporary Combined Lecture allotments.xlsx', index=False)




#generate final outputs
class_map = pd.read_excel('classs_lec.xlsx', sheet_name='ClassNumb')
course_map = pd.read_excel('classs_lec.xlsx', sheet_name='ClassCode')
time_map = pd.read_excel('classs_lec.xlsx', sheet_name='ClassTime')

course_dict={}
class_dict={}
time_dict={}
for i in range(course_map.shape[0]):
    course_dict[course_map['Course'][i]]=course_map['o1'][i]
for i in range(class_map.shape[0]):
    class_dict[class_map['Course'][i]]=class_map['o1'][i]
for i in range(time_map.shape[0]):
    time_dict[time_map['Course'][i]]=time_map['o1'][i]
    
id_to_name={}
pr_df_2019 = pd.read_excel('GOA PR.xlsx', sheet_name='2019')
for i in range(pr_df_2019.shape[0]):
    current_id = pr_df_2019['CAMPUS_ID'][i].strip().upper()
    current_name=pr_df_2019['NAME'][i].strip().upper()
    id_to_name[current_id]=current_name

def solve_CS(curr_id ,labs, lab_allot):
    for i in range(3):
        try:
            p=course_dict[lab_allot[curr_id]]
            l=[curr_id,p[i],class_dict[lab_allot[curr_id]][i],time_dict[lab_allot[curr_id]][i]]
            #labs.append(l)
           # print(l)
        except:
            print(curr_id)
def solve_REST(curr_id ,labs, dd_lab_allot,em_lab_allot):
    try:
        p=course_dict[dd_lab_allot[curr_id]]
        l=[curr_id,p,class_dict[dd_lab_allot[curr_id]],time_dict[dd_lab_allot[curr_id]]]
        labs.append(l)
        print(l)
        p=course_dict[em_lab_allot[curr_id]]
        l=[curr_id,p,class_dict[em_lab_allot[curr_id]],time_dict[em_lab_allot[curr_id]]]
        labs.append(l)
    except:
        print(curr_id)
            #labs.append(l)
    
CS_labs=[]
ECE_labs=[]
EEE_labs=[]
ENI_labs=[]

for i in range(len(pr_df_2019['PR NO.'])):
    current_id = pr_df_2019['CAMPUS_ID'][i].strip().upper()
    current_branch = current_id[4:6]
    if current_branch=='A3' or current_branch=='a3':
        solve_REST(current_id,EEE_labs,dd_allot_dict,em_allot_dict)
    elif current_branch=='A8' or current_branch=='a8':
        solve_REST(current_id,ENI_labs,dd_allot_dict,em_allot_dict)
    elif current_branch=='AA' or current_branch=='aa':
        solve_REST(current_id,ECE_labs,dd_allot_dict,em_allot_dict)
        
import xlsxwriter 
def make_file_cs(course,CS_labs):
    workbook = xlsxwriter.Workbook('output_lect_cs.xlsx')
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
    z=0
    workbook = xlsxwriter.Workbook('output_lec_'+course+'.xlsx')
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
            z+=1
            worksheet.write(j,0,i[0],cell_format3)
            worksheet.write(j,1,id_to_name[i[0]],cell_format3)
            worksheet.write(j,2,i[1],cell_format3)
            worksheet.write(j,3,i[2],cell_format3)
            worksheet.write(j,4,i[3],cell_format3)
            j=j+1
            #print(i)
    print(z)
    workbook.close()

#make_file_cs('CS',CS_labs)
make_file_rest('EEE',EEE_labs)
make_file_rest('ECE',ECE_labs)
make_file_rest('ENI',ENI_labs)
