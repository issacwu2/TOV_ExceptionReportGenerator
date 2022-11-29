import pandas as pd
import time
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import os
import shutil
import openpyxl
import numpy as np

def clean_case(case):
    case_fix = case.drop(columns=['exception type_x', 'level_x', 'startKm_x', 'endKm_x',
                                    'length_x', 'maxValue_x', 'maxLocation_x', 'track type_x',
                                    'key', 'startKm_shift_x', 'endKm_shift_x',
                                  'maxLocation_shift_x', 'startKm_shift_y', 'endKm_shift_y', 'maxLocation_shift_y', 'Overlap_x',
                                  'Tension Length_x', 'Landmark_x'])\
        .rename({'id_y': 'id', 'exception type_y': 'exception type', 'level_y': 'level', 'startKm_y': 'startKm',
                 'endKm_y': 'endKm', 'length_y': 'length', 'maxValue_y': 'maxValue', 'maxLocation_y': 'maxLocation',
                 'track type_y': 'track type', 'Overlap_y': 'Overlap', 'Tension Length_y': 'Tension Length',
                 'Landmark_y': 'Landmark'}, axis=1)
    if 'previous' in case_fix.columns:
        case_fix['previous'] = case_fix['previous'].astype(str) + ', ' + case_fix['id_x'].astype(str)
    else:
        case_fix['previous'] = case_fix['id_x'].astype(str)

    case_fix = case_fix.reset_index().drop(columns=['index', 'id_x'])
    return case_fix[['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue',
                     'maxLocation', 'track type', 'Overlap', 'Tension Length', 'Landmark', 'previous']]


def find_repeated(df1, df2, df1_shift, df2_shift):
    if df1.empty or df2.empty:
        return pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length',
                                     'maxValue', 'maxLocation', 'track type', 'Overlap', 'Tension Length', 'Landmark', 'previous'])
    else:
        for all in ['startKm', 'endKm', 'maxLocation']:

            df1[all + '_shift'] = df1[all] + df1_shift
            df2[all + '_shift'] = df2[all] + df2_shift
        merged = df1.assign(key=1).merge(df2.assign(key=1), on='key')

        # ---------- case1 = 1st exception behind, 2nd at front ----------
        case1 = merged.query('(`startKm_x`.between(`startKm_y`, `endKm_y`)) & '
                           '(`endKm_x` > `endKm_y`) &'
                           '(`maxLocation_x`.between(`startKm_x`, `endKm_y`)) &'
                           '(`maxLocation_y`.between(`startKm_x`, `endKm_y`))', engine='python')

        # ---------- case2 = 1st exception at front, 2nd behind ----------
        case2 = merged.query('(`startKm_y`.between(`startKm_x`, `endKm_x`)) & '
                           '(`endKm_x` < `endKm_y`) &'
                           '(`maxLocation_x`.between(`startKm_y`, `endKm_x`)) &'
                           '(`maxLocation_y`.between(`startKm_y`, `endKm_x`))', engine='python')

        # ---------- case3 = 2nd exception covering whole 1st  ----------
        case3 = merged.query('(`startKm_x` >= `startKm_y`) & '
                           '(`endKm_x` <= `endKm_y`) & '
                           '(`maxLocation_x`.between(`startKm_x`, `endKm_x`)) &'
                           '(`maxLocation_y`.between(`startKm_x`, `endKm_x`))', engine='python')

        # ---------- case4 = 1st exception covering whole 2nd  ----------
        case4 = merged.query('(`startKm_x` <= `startKm_y`) & '
                           '(`endKm_x` >= `endKm_y`) & '
                           '(`maxLocation_x`.between(`startKm_y`, `endKm_y`)) &'
                           '(`maxLocation_y`.between(`startKm_y`, `endKm_y`))', engine='python')

        # ======== Reserved part: repeat logic with chainage shift =========
        # ---------- case1 = 1st exception behind, 2nd at front ----------
        # case1 = merged.query('(`startKm_shift_x`.between(`startKm_shift_y`, `endKm_shift_y`)) & '
        #                    '(`endKm_shift_x` > `endKm_shift_y`) &'
        #                    '(`maxLocation_shift_x`.between(`startKm_shift_x`, `endKm_shift_y`)) &'
        #                    '(`maxLocation_shift_y`.between(`startKm_shift_x`, `endKm_shift_y`))', engine='python')
        #
        # # ---------- case2 = 1st exception at front, 2nd behind ----------
        # case2 = merged.query('(`startKm_shift_y`.between(`startKm_shift_x`, `endKm_shift_x`)) & '
        #                    '(`endKm_shift_x` < `endKm_shift_y`) &'
        #                    '(`maxLocation_shift_x`.between(`startKm_shift_y`, `endKm_shift_x`)) &'
        #                    '(`maxLocation_shift_y`.between(`startKm_shift_y`, `endKm_shift_x`))', engine='python')
        #
        # # ---------- case3 = 2nd exception covering whole 1st  ----------
        # case3 = merged.query('(`startKm_shift_x` >= `startKm_shift_y`) & '
        #                    '(`endKm_shift_x` <= `endKm_shift_y`) & '
        #                    '(`maxLocation_shift_x`.between(`startKm_shift_x`, `endKm_shift_x`)) &'
        #                    '(`maxLocation_shift_y`.between(`startKm_shift_x`, `endKm_shift_x`))', engine='python')
        #
        # # ---------- case4 = 1st exception covering whole 2nd  ----------
        # case4 = merged.query('(`startKm_shift_x` <= `startKm_shift_y`) & '
        #                    '(`endKm_shift_x` >= `endKm_shift_y`) & '
        #                    '(`maxLocation_shift_x`.between(`startKm_shift_y`, `endKm_shift_y`)) &'
        #                    '(`maxLocation_shift_y`.between(`startKm_shift_y`, `endKm_shift_y`))', engine='python')
        return pd.concat([clean_case(case1), clean_case(case2), clean_case(case3), clean_case(case4)])\
            .drop_duplicates()\
            .reset_index().drop('index', axis=1)


print('================ Finding Repeated Exception Tool ================')
print('This is a tool to find repeated exceptions of 2/3 consecutive TOV Exception Reports')
check = input('Compare 2 or 3 TOV exception reports? (input 2/3):')
while not check in ['2', '3']:
    print('Error! Make sure you only input 2 or 3')
    check = input('Compare 2 or 3 TOV exception reports? (input 2/3):')
print('===========================================================')
print('You are going to compare ' + check + ' TOV Exception Reports')

print('Select 1st Exception Report after 3 seconds...')
for i in range(3, 0, -1):
    print(f"{i}", end="\r", flush=True)
    time.sleep(1)

# ---------- allow user to select Exception Report files -------
root = tk.Tk()
root.withdraw()
df1_path = filedialog.askopenfilename()
print('Selected: ' + os.path.basename(df1_path))
df1_W = pd.read_excel(df1_path, sheet_name='wear exception')
df1_LH = pd.read_excel(df1_path, sheet_name='low height exception')
df1_HH = pd.read_excel(df1_path, sheet_name='high height exception')
df1_SL = pd.read_excel(df1_path, sheet_name='stagger left exception')
df1_SR = pd.read_excel(df1_path, sheet_name='stagger right exception')
df1_SL = df1_SL.loc[df1_SL['Overlap'] != 'Y']
df1_SR = df1_SR.loc[df1_SR['Overlap'] != 'Y']
# df1_SL = df1_SL.loc[df1_SL['level'] != 'L3']
# df1_SR = df1_SR.loc[df1_SR['level'] != 'L3']


print('Select 2nd Exception Report after 3 seconds...')
for i in range(3, 0, -1):
    print(f"{i}", end="\r", flush=True)
    time.sleep(1)
df2_path = filedialog.askopenfilename()
print('Selected: ' + os.path.basename(df2_path))
df2_W = pd.read_excel(df2_path, sheet_name='wear exception')
df2_LH = pd.read_excel(df2_path, sheet_name='low height exception')
df2_HH = pd.read_excel(df2_path, sheet_name='high height exception')
df2_SL = pd.read_excel(df2_path, sheet_name='stagger left exception')
df2_SR = pd.read_excel(df2_path, sheet_name='stagger right exception')
df2_SL = df2_SL.loc[df2_SL['Overlap'] != 'Y']
df2_SR = df2_SR.loc[df2_SR['Overlap'] != 'Y']
# df2_L3 = pd.concat([df2_SL.loc[df2_SL['level'] == 'L3'], df2_SR.loc[df2_SR['level'] == 'L3']])
# df2_SL = df2_SL.loc[df2_SL['level'] != 'L3']
# df2_SR = df2_SR.loc[df2_SR['level'] != 'L3']


# ---------- only for AEL and TCL -------
if check == '3':
    print('Select 3rd Exception Report after 3 seconds...')
    for i in range(3, 0, -1):
        print(f"{i}", end="\r", flush=True)
        time.sleep(1)
    df3_path = filedialog.askopenfilename()
    print('Selected: ' + os.path.basename(df3_path))
    df3_W = pd.read_excel(df3_path, sheet_name='wear exception')
    df3_LH = pd.read_excel(df3_path, sheet_name='low height exception')
    df3_HH = pd.read_excel(df3_path, sheet_name='high height exception')
    df3_SL = pd.read_excel(df3_path, sheet_name='stagger left exception')
    df3_SR = pd.read_excel(df3_path, sheet_name='stagger right exception')
    df3_SL = df3_SL.loc[df3_SL['Overlap'] != 'Y']
    df3_SR = df3_SR.loc[df3_SR['Overlap'] != 'Y']
    # df3_L3 = pd.concat([df3_SL.loc[df3_SL['level'] == 'L3'], df3_SR.loc[df3_SR['level'] == 'L3']])
    # df3_SL = df3_SL.loc[df3_SL['level'] != 'L3']
    # df3_SR = df3_SR.loc[df3_SR['level'] != 'L3']
# ---------- only for AEL and TCL -------

# ---------- allow user to select Exception Report files -------


# ---------- find repeated exception for each type ----------
repeated_W = find_repeated(df1_W, df2_W, 0, 0)
repeated_LH = find_repeated(df1_LH, df2_LH, 0, 0)
repeated_HH = find_repeated(df1_HH, df2_HH, 0, 0)
repeated_SL = find_repeated(df1_SL, df2_SL, 0, 0)
repeated_SR = find_repeated(df1_SR, df2_SR, 0, 0)

# ---------- only for AEL and TCL -------
if check == '3':
    triple_repeated_W = find_repeated(repeated_W, df3_W, 0, 0)
    triple_repeated_LH = find_repeated(repeated_LH, df3_LH, 0, 0)
    triple_repeated_HH = find_repeated(repeated_HH, df3_HH, 0, 0)
    triple_repeated_SL = find_repeated(repeated_SL, df3_SL, 0, 0)
    triple_repeated_SR = find_repeated(repeated_SR, df3_SR, 0, 0)
# ---------- only for AEL and TCL -------

# ---------- find repeated exception for each type ----------


# ---------- output results ----------
print('Saving as Excel at', datetime.now())
print('Done!')
input('Press Enter to select save location')
directory = filedialog.askdirectory()
if check == '2':
    final_path = directory + '/' + os.path.basename(os.path.splitext(df2_path)[0]) + '_repeated.xlsx'
elif check == '3':
    final_path = directory + '/' + os.path.basename(os.path.splitext(df3_path)[0]) + '_triple_repeated.xlsx'


# ----- save as -----
temp = openpyxl.load_workbook('EAL metadata.xlsx')
sheets = temp.sheetnames
for s in sheets:
    if s != 'template':
        del temp[s]
ss = temp['template']
ss.title = 'Summary'
temp.save(final_path)
if check == '3':
    update = pd.concat([triple_repeated_SL, triple_repeated_SR, triple_repeated_W, triple_repeated_HH, triple_repeated_LH], ignore_index=True, sort=False)
else:
    update = pd.concat([repeated_SL, repeated_SR, repeated_W, repeated_HH, repeated_LH], ignore_index=True, sort=False)

update['support'] = np.nan
update['empty1'] = np.nan
update['empty2'] = np.nan
update['empty3'] = np.nan
update['empty4'] = np.nan
update = update[['id', 'startKm', 'endKm', 'length', 'exception type', 'maxValue', 'maxLocation', 'support',
                 'Tension Length', 'track type', 'level', 'empty1', 'empty2', 'empty3', 'empty4', 'previous']]


with pd.ExcelWriter(final_path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
	update.to_excel(writer, sheet_name="Summary", startrow=2, header=None, index=False)
# ----- save as -----


# with pd.ExcelWriter(final_path) as writer:
#     if check == '3':
#         triple_repeated_W.sort_values('startKm').to_excel(writer, sheet_name='wear exception', index=False)
#         triple_repeated_LH.sort_values('startKm').to_excel(writer, sheet_name='low height exception', index=False)
#         triple_repeated_HH.sort_values('startKm').to_excel(writer, sheet_name='high height exception', index=False)
#         triple_repeated_SL.sort_values('startKm').to_excel(writer, sheet_name='stagger left exception', index=False)
#         triple_repeated_SR.sort_values('startKm').to_excel(writer, sheet_name='stagger right exception', index=False)
#         df3_L3.to_excel(writer, sheet_name='L3 alarms', index=False)
#
#     else:
#         repeated_W.sort_values('startKm').to_excel(writer, sheet_name='wear exception', index=False)
#         repeated_LH.sort_values('startKm').to_excel(writer, sheet_name='low height exception', index=False)
#         repeated_HH.sort_values('startKm').to_excel(writer, sheet_name='high height exception', index=False)
#         repeated_SL.sort_values('startKm').to_excel(writer, sheet_name='stagger left exception', index=False)
#         repeated_SR.sort_values('startKm').to_excel(writer, sheet_name='stagger right exception', index=False)
#         df2_L3.to_excel(writer, sheet_name='L3 alarms', index=False)
#
# os.startfile(final_path)