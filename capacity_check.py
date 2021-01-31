import pandas as pd

idlist = []
#Create DF of ID No | PR No. DF sorted on PR NO.
pr_df_2019 = pd.read_excel('GOA PR.xlsx', sheet_name='2019')
pr_df_2019 = pr_df_2019[['CAMPUS_ID', 'PR NO.']]
pr_df_2019 = pr_df_2019.sort_values(by=['PR NO.'])
pr_df_2019 = pr_df_2019.reset_index(drop = True)

cs_count = 0
ece_count = 0
eee_count = 0
eni_count = 0

for i in range(len(pr_df_2019)):
    current_id = pr_df_2019['CAMPUS_ID'][i].strip().upper()
    current_branch = current_id[4:6]
    if current_branch == 'AA':
        ece_count += 1
    elif current_branch == 'A3':
        eee_count += 1
    elif current_branch == 'A8':
        eni_count += 1
    elif current_branch == 'A7':
        cs_count += 1

print('student counts:')
print([cs_count, ece_count, eee_count, eni_count])
# cs   ece  eee eni
# [170, 90, 83, 84] student count
#



