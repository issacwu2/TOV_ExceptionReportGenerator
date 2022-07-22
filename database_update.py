import pandas as pd
import openpyxl
import numpy as np
import datetime

# ----- function to find last row with non-empty value for appending new exception -----
def get_maximum_rows(*, sheet_object):
    rows = 0
    for max_row, row in enumerate(sheet_object, 1):
        if not all(col.value is None for col in row):
            rows += 1
    return rows + 1
# ----- function to find last row with non-empty value for appending new exception -----


def clean_case(case):
	return case


def find_repeated(old_db, new, new_shift):

	# for all in ['startKm', 'endKm', 'maxLocation']:
	# 	old_db[all + '_shift'] = old_db[all] + old_db['chainage shift']
	# 	new[all + '_shift'] = new[all] + new_shift

	merged = old_db.assign(key=1).merge(new.assign(key=1), on='key').query('(`Exception` == `exception type`) & `Line_x` == `Line_y`')

	# ---------- case1 = 1st exception behind, 2nd at front ----------
	case1 = merged.query('(`startKm_x`.between(`startKm_y`, `endKm_y`)) & '
					   '(`endKm_x` > `endKm_y`)', engine='python')

	# ---------- case2 = 1st exception at front, 2nd behind ----------
	case2 = merged.query('(`startKm_y`.between(`startKm_x`, `endKm_x`)) & '
					   '(`endKm_x` < `endKm_y`)', engine='python')

	# ---------- case3 = 2nd exception covering whole 1st  ----------
	case3 = merged.query('(`startKm_x` >= `startKm_y`) & '
					   '(`endKm_x` <= `endKm_y`)', engine='python')

	# ---------- case4 = 1st exception covering whole 2nd  ----------
	case4 = merged.query('(`startKm_x` <= `startKm_y`) & '
					   '(`endKm_x` >= `endKm_y`)', engine='python')
	return pd.concat([clean_case(case1), clean_case(case2), clean_case(case3), clean_case(case4)]) \
		.drop_duplicates() \
		.reset_index(drop=True)


new_exception_path = 'C:/Users/issac/Coding Projects/TOV_ExceptionReportGenerator/KTL data report/jun 2022/20220627_KTL_UT_Exception Report_repeated.xlsx'
old_db_path = 'C:/Users/issac/Coding Projects/TOV_ExceptionReportGenerator/KTL data report/Exception Follow-up Master List- KTL up to mar 2022-test.xlsx'

# ----- read data -----
workbook = openpyxl.load_workbook(old_db_path)
L3_sheet_max_row = get_maximum_rows(sheet_object=workbook['L3 alarms'])
summary_sheet_max_row = get_maximum_rows(sheet_object=workbook['Summary'])

old_db_complete = pd.read_excel(old_db_path, sheet_name='Summary', header=3)
old_db = pd.read_excel(old_db_path, sheet_name='Summary', usecols='A:P, T', header=3)
# ----- ignore FALSE alarm > 1 year -----
old_db = old_db.loc[~(old_db['False Alarm? (Y/N)'].eq('Yes') & old_db['Completion Date'].lt(np.datetime64('today')-365))]
# ----- ignore FALSE alarm > 1 year -----

# ----- new exception to be added to db -----
exception_df_list = [
	pd.read_excel(new_exception_path, sheet_name='wear exception'),
	pd.read_excel(new_exception_path, sheet_name='low height exception'),
	pd.read_excel(new_exception_path, sheet_name='high height exception'),
	pd.read_excel(new_exception_path, sheet_name='stagger left exception'),
	pd.read_excel(new_exception_path, sheet_name='stagger right exception')
]
shift = pd.read_excel(new_exception_path, sheet_name='chainage shift').values.flatten().tolist()[0]

new_exception_all = pd.concat(exception_df_list).reset_index(drop=True)
new_exception_all['exception type'] = new_exception_all['exception type'].str.replace('Low ', '').str.replace('High ', '')
# ----- new exception to be added to db -----

# ----- L3 -----
L3_new = pd.read_excel(new_exception_path, sheet_name='L3 alarms')
L3_old = pd.read_excel(old_db_path, sheet_name='L3 alarms', header=3)
# ----- L3 -----

# ----- read data -----

# ----- align column format -----
for all in [new_exception_all, L3_new]:
	all['Date'] = pd.to_datetime(all.id.str[0:8], format='%Y%m%d')
	all['Line'] = all.id.str[9:15].replace('_', ' ', regex=True).str.rstrip()
# ----- align column format -----

# ----- L3 alarm to be handled separately -----

# ----- handle L3 alarms with "Reoccurrence" not empty -----
L3_append_re = find_repeated(L3_old, L3_new, 0)[['ID', 'Reoccurrence', 'id']].groupby(['ID', 'Reoccurrence'])['id'].apply(list).reset_index(name='reoccur-list')
L3_append_re['Reoccurrence'] = L3_append_re['Reoccurrence'].apply(lambda x: f'{x:.0f}' if type(x) == float else x).astype(str)
L3_append_re['reoccur-list-2'] = L3_append_re.apply(lambda x: np.append(x['Reoccurrence'], x['reoccur-list']), axis=1).apply(', '.join)
# ----- handle L3 alarms with "Reoccurrence" not empty -----


# ----- handle L3 alarms with "Reoccurrence"  empty -----
L3_append_re2 = find_repeated(L3_old, L3_new, 0)[['ID', 'Reoccurrence', 'id']]
L3_append_re2 = L3_append_re2[L3_append_re2['Reoccurrence'].isna()].groupby('ID')['id'].apply(list).reset_index(name='reoccur-list')
L3_append_re2['Reoccurrence'] = L3_append_re2['reoccur-list'].apply(', '.join)
# ----- handle L3 alarms with "Reoccurrence"  empty -----

# ------ update L3 Reoccurrence -----
L3_update = L3_old.copy()
L3_update['Date'] = L3_update['Date'].dt.date
for row, index in L3_append_re.iterrows():
	L3_update.loc[(L3_update['ID'] == index['ID']), 'Reoccurrence'] = index['reoccur-list-2']

for row, index in L3_append_re2.iterrows():
	L3_update.loc[(L3_update['ID'] == index['ID']), 'Reoccurrence'] = index['Reoccurrence']
# ------ update L3 Reoccurrence -----

# ----- new L3 alarms -----
L3_new_selected = L3_new.copy()
for repeated in find_repeated(L3_old, L3_new, 0)['id'].drop_duplicates():
	L3_new_selected.drop(L3_new_selected[L3_new_selected.id == repeated].index, inplace=True)
L3_new_selected = L3_new_selected.rename({'id': 'ID', 'exception type': 'Exception',
									'level': 'Level',
									'length': 'Length',
									'track type': 'Track Type',
									'location type': 'Location Type'}, axis=1)
L3_new_selected['Reoccurrence'] = np.nan
L3_new_selected['Date'] = L3_new_selected['Date'].dt.date
L3_new_selected['Length'] = L3_new_selected['Length']*1000
L3_new_selected['chainage shift'] = shift
L3_new_selected = L3_new_selected[['ID', 'Date', 'Line', 'Exception', 'Level', 'startKm', 'endKm', 'Length', 'maxValue',
       'maxLocation', 'Track Type', 'Location Type', 'chainage shift', 'Reoccurrence']]
# ----- new L3 alarms -----

# ----- L3 alarm to be handled separately -----


# ----- Append exception ID in column 'reoccurrence' if old db already had a follow up action -----
append_re = find_repeated(old_db, new_exception_all, shift)[['ID', 'Reoccurrence', 'id']]
append_re_grp = append_re.groupby(['ID', 'Reoccurrence'])['id'].apply(list).reset_index(name='reoccur-list')
append_re_grp['Reoccurrence'] = append_re_grp['Reoccurrence'].apply(lambda x: f'{x:.0f}' if type(x) == float else x).astype(str)
append_re_grp['reoccur-list-2'] = append_re_grp.apply(lambda x: np.append(x['Reoccurrence'], x['reoccur-list']), axis=1).apply(', '.join)

# append_re = find_repeated(old_db, new_exception_all, shift)[['ID', 'Reoccurrence', 'id']]
# append_re['Reoccurrence'] = append_re['Reoccurrence'].astype(str) + ', ' + append_re['id'].astype(str)

old_db_update = old_db_complete.copy()
old_db_update['Date'] = old_db_update['Date'].dt.date
old_db_update['Target Date'] = old_db_update['Target Date'].dt.date
old_db_update['Completion Date'] = old_db_update['Completion Date'].dt.date
old_db_update['Completion Date.1'] = old_db_update['Completion Date.1'].apply(lambda x: x.date() if isinstance(x, datetime.datetime) else x)
old_db_update['Target Date.1'] = old_db_update['Target Date.1'].dt.date
for index, row in append_re_grp.iterrows():
	old_db_update.loc[(old_db_update['ID'] == row['ID']), 'Reoccurrence'] = row['reoccur-list-2']
# ----- Append exception ID in column 'reoccurrence' if old db already had a follow up action -----

# ----- Append new exceptions into old_db table section 0. -----
new_selected_l = append_re['id'].values.tolist()
new_selected = new_exception_all.query('~id.isin(@new_selected_l)')
new_selected = new_selected.rename({'id': 'ID', 'exception type': 'Exception',
									'level': 'Level',
									'length': 'Length',
									'track type': 'Track Type',
									'location type': 'Location Type',
									'previous': 'Reoccurence'}, axis=1)

new_selected['chainage shift'] = shift
new_selected['Length'] = new_selected['Length']*1000
new_selected.loc[(new_selected['Level'] == 'L1'), 'Target Date'] = new_selected['Date'] + datetime.timedelta(days=14)
new_selected.loc[(new_selected['Level'] == 'L2'), 'Target Date'] = new_selected['Date'] + datetime.timedelta(days=30)
new_selected = new_selected[['ID', 'Date', 'Line', 'Exception', 'Level', 'startKm', 'endKm', 'Length', 'maxValue',
       'maxLocation', 'chainage shift', 'Track Type', 'Location Type', 'Reoccurence', 'Target Date']]

new_selected['Date'] = new_selected['Date'].dt.date
new_selected['Target Date'] = new_selected['Target Date'].dt.date

# ----- Append new exceptions into old_db table section 0. -----
with pd.ExcelWriter(old_db_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
	old_db_update.to_excel(writer, sheet_name='Summary', startrow=4, header=None, index=False)
	new_selected.to_excel(writer, sheet_name='Summary', startrow=summary_sheet_max_row, header=None, index=False)
	L3_update.to_excel(writer, sheet_name='L3 alarms', startrow=4, header=None, index=False)
	L3_new_selected.to_excel(writer, sheet_name='L3 alarms', startrow=L3_sheet_max_row, header=None, index=False)