import pandas as pd

idlist = []
#Create DF of ID No | PR No. DF sorted on PR NO.
pr_df_2019 = pd.read_excel('GOA PR.xlsx', sheet_name='2019')
pr_df_2019 = pr_df_2019[['CAMPUS_ID', 'PR NO.']]
pr_df_2019 = pr_df_2019.sort_values(by=['PR NO.'])
pr_df_2019 = pr_df_2019.reset_index(drop = True)

for i in range(len(pr_df_2019)):
    idlist.append(pr_df_2019['CAMPUS_ID'][i])

global wrong_id_entries
wrong_id_entries = []
def check_id_present(file_path):
    global wrong_id_entries
    responses_df = pd.read_excel(file_path)
    for i in range(len(responses_df['ID Number'])):
        current_id = responses_df['ID Number'][i].strip().upper()
        current_email = responses_df['Email Address'][i].strip()
        current_last4 = current_email[5:9]
        if current_last4 != current_id[8:12]:
            print("ID-emailID mismatch:" + current_id)
        if current_id not in idlist:
            wrong_id_entries.append([current_id, file_path])

check_id_present('./Responses/CS Form (Responses).xlsx')
check_id_present('./Responses/ECE Form (Responses).xlsx')
check_id_present('./Responses/EEE Form (Responses).xlsx')
check_id_present('./Responses/ENI Form (Responses).xlsx')

print(wrong_id_entries)