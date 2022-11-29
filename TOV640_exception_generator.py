import numpy as np
import pandas as pd
import os
import re
from datetime import datetime
import time
import tkinter as tk
from tkinter import filedialog
pd.options.mode.chained_assignment = None

print('================ TOV640 Exception Report Generator ================')
print('Before you start, please go through the following instructions or otherwise you might experience errors and the results could be inaccurate!')
input("Press Enter to continue...")
print('0. Make sure all metadata.xlsx of each line are in the same folder with this script.')
print('If not, please move them together and restart the program.')
input("Press Enter to continue...")
print('When the program is running, MAKE SURE YOU:')
print('1. Input the name of Line e.g. EAL/TML...')
print('2. Input the direction of track, i.e. UP/DN')
print('3. Select .DATAC file of TOV data to generate exception report')
print('5. Input the date of when that TOV data is obtained, in YYYY/MM/DD')
input("Press Enter to continue...")
print('=========================== IMPORTANT!!! ===========================')
input('Make sure you read through the above instructions carefully and press enter to start... ')

line = input('Line: ')
while not line in ['EAL', 'TML']:
    print('Please make sure you input correct line:')
    line = input('line:')

if line == 'TML':
    section = input('WRL/KSL/ETSE/MOL? :')
    while not section in ['WRL', 'KSL', 'ETSE', 'MOL']:
        print('Please make sure your input is either WRL/KSL/ETSE/MOL')
        section = input('WRL/KSL/ETSE/MOL? :')

if line == 'EAL':
    section = input('LMC? [y/n]: ')
    while not section in ['y', 'n']:
        section = input('LMC? [y/n]: ')
    if section == 'y':
        section = 'LMC'
    else:
        section = input('RAC? [y/n]: ')
        while not section in ['y', 'n']:
            section = input('RAC? [y/n]: ')
        if section == 'y':
            section = 'RAC'
        else:
            section = input('LOW S1? [y/n]: ')
            while not section in ['y', 'n']:
                section = input('LOW S1? [y/n]: ')
            if section == 'y':
                section = 'LOW'
                track = 'S1'
            else:
                section = input('Please input the section range, e.g. UNI-TAP:')
                while not re.match('^[A-Z]{3}-[A-Z]{3}', section):
                    print('Please make sure you input correct section range, e.g. UNI-TAP')
                    section = input('Please input the section range:')

if section != 'LOW':
    track = input('UP/DN? : ')
    while not track in ['DN', 'UP']:
        print('Please make sure you input either UP / DN')
        track = input('UP/DN? : ')
else:
    pass

d = input('Date (YYYY/MM/DD) : ')
while not re.match('^\d{4}\/\d{2}\/\d{2}$', d):
    print('Please make sure you in put the date in correct format [YYYY/MM/DD]')
    d = input('Date (YYYY/MM/DD) : ')

date = d.replace('/', '')

print('select the data report in .DATAC format after 1 second...')
for i in range(1,0,-1):
    print(f"{i}", end="\r", flush=True)
    time.sleep(1)
# ---------- allow user to select csv files -------
root = tk.Tk()
root.withdraw()
raw_data_path = filedialog.askopenfilename()
raw = pd.read_csv(raw_data_path, sep=';', engine='python').apply(lambda x: x.str.strip() if x.dtype == "object" else x)
print('Selected: ' + os.path.basename(raw_data_path))
# ---------- allow user to select csv files -------

# --------- Load metadata ----------
if section in ['LMC', 'RAC', 'LOW', 'WRL', 'KSL', 'ETSE', 'MOL']:
    track_type = pd.read_excel('./' + line + ' metadata.xlsx', sheet_name=section + ' ' + track + ' track type')
    overlap = pd.read_excel('./' + line + ' metadata.xlsx', sheet_name=section + ' ' + track + ' Tension Length')
    landmark = pd.read_excel('./' + line + ' metadata.xlsx', sheet_name=section + ' ' + track + ' Landmark')
else:
    track_type = pd.read_excel('./' + line + ' metadata.xlsx', sheet_name=track + ' track type')
    overlap = pd.read_excel('./' + line + ' metadata.xlsx', sheet_name=track + ' Tension Length')
    landmark = pd.read_excel('./' + line + ' metadata.xlsx', sheet_name=track + ' Landmark')
threshold = pd.read_excel('./' + line + ' metadata.xlsx', sheet_name='threshold')

overlap['Overlap'].fillna('N', inplace=True)
# --------- Load metadata --------
print('Loading...')

# ----------- for debugging ----------
# section = 'LMC'
# line = 'EAL'
# track = 'DN'
# date = '20220610'
# track_type = pd.read_excel('C:/Users/issac/Coding Projects/TOV_ExceptionReportGenerator/640/TOV RAW/EAL metadata.xlsx', sheet_name='DN track type')
# threshold = pd.read_excel('C:/Users/issac/Coding Projects/TOV_ExceptionReportGenerator/640/TOV RAW/EAL metadata.xlsx', sheet_name='threshold')
# overlap = pd.read_excel('C:/Users/issac/Coding Projects/TOV_ExceptionReportGenerator/640/TOV RAW/EAL metadata.xlsx', sheet_name='DN Tension Length')
# raw = pd.read_csv('C:/Users/issac/Coding Projects/TOV_ExceptionReportGenerator/640/TOV RAW/2206/EAL_D1_2206.datac', sep=';').apply(lambda x: x.str.strip() if x.dtype == "object" else x)
# ----------- for debugging ----------

raw.columns = raw.columns.str.replace(' ', '')
raw.drop(raw[raw.KM == 'KM'].index, inplace=True)

raw = raw.rename({'STG1c': 'stagger1',
                  'STG2c': 'stagger2',
                  'STG3c': 'stagger3',
                  'STG4c': 'stagger4',
                  'RWH1mm': 'wear1',
                  'RWH2mm': 'wear2',
                  'RWH3mm': 'wear3',
                  'RWH4mm': 'wear4',
                  'WHGT1c': 'height1',
                  'WHGT2c': 'height2',
                  'WHGT3c': 'height3',
                  'WHGT4c': 'height4',
                  'LINE': 'Line',
                  'TRACK': 'Track'}, axis=1)

raw[['KM', 'LOCATION', 'height1', 'height2', 'height3', 'height4', 'wear1', 'wear2', 'wear3', 'wear4', 'stagger1', 'stagger2', 'stagger3',
	 'stagger4']] = raw[['KM', 'LOCATION', 'height1', 'height2', 'height3', 'height4', 'wear1', 'wear2', 'wear3', 'wear4', 'stagger1', 'stagger2', 'stagger3',
	 'stagger4']].apply(pd.to_numeric)

raw['Km'] = raw['KM'] + raw['LOCATION']*0.001
raw['height1'] = raw['height1'] + 5300
raw['height2'] = raw['height2'] + 5300
raw['height3'] = raw['height3'] + 5300
raw['height4'] = raw['height4'] + 5300
raw = raw[['Line', 'Track', 'Km', 'height1', 'height2', 'height3', 'height4', 'wear1', 'wear2', 'wear3', 'wear4', 'stagger1', 'stagger2', 'stagger3', 'stagger4']]

raw['Km'] = raw['Km']\
    .round(decimals=3)
# set to chainage accuracy to 0.001km = 1m


# ---------- To remove unreasonable CW height data ----------
WH_min = raw.groupby('Km')[['height1', 'height2', 'height3', 'height4']].min().reset_index()
WH_min.loc[(WH_min['height1'] < 3500) | (WH_min['height2'] < 3500) | (WH_min['height3'] < 3500) | (WH_min['height4'] < 3500), 'error'] = WH_min['Km']
WH_error = WH_min['error'].dropna().to_list()
WH = raw[['Km', 'height1', 'height2', 'height3', 'height4']]
WH_cleaned = WH[~WH.Km.isin(WH_error)]
# ---------- To remove unreasonable CW height data ----------

# ---------- preprocess the data before generating exception ----------
stagger_left = raw.groupby('Km')[['stagger1', 'stagger2', 'stagger3', 'stagger4']]\
    .max()\
    .reset_index()
stagger_right = raw.groupby('Km')[['stagger1', 'stagger2', 'stagger3', 'stagger4']]\
    .min()\
    .reset_index()

wear_min = raw.groupby('Km')[['wear1', 'wear2', 'wear3', 'wear4']]\
    .min()\
    .reset_index()

WH_cleaned_max = WH_cleaned.groupby('Km')[['height1', 'height2', 'height3', 'height4']]\
    .max()\
    .reset_index()
WH_cleaned_min = WH_cleaned.groupby('Km')[['height1', 'height2', 'height3', 'height4']]\
    .min()\
    .reset_index()
# ---------- preprocess the data before generating exception ----------


# ---------- identify track type (tangent/curve) ----------
stagger_left = stagger_left \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

stagger_right = stagger_right \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

wear_min = wear_min \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

WH_cleaned_max = WH_cleaned_max \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

WH_cleaned_min = WH_cleaned_min \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)
# ---------- identify track type (tangent/curve) ----------


# ---------- find extreme value amongst 4 channels ----------
WH_cleaned_min['maxValue'] = WH_cleaned_min[['height1', 'height2', 'height3', 'height4']].min(axis=1)
WH_cleaned_max['maxValue'] = WH_cleaned_max[['height1', 'height2', 'height3', 'height4']].max(axis=1)

wear_min['maxValue'] = wear_min[['wear1', 'wear2', 'wear3', 'wear4']].min(axis=1)

stagger_left['maxValue'] = stagger_left[['stagger1', 'stagger2', 'stagger3', 'stagger4']].max(axis=1)
stagger_right['maxValue'] = stagger_right[['stagger1', 'stagger2', 'stagger3', 'stagger4']].min(axis=1)
# ---------- find extreme value amongst 4 channels ----------


# ---------- load line related threshold from metadata ----------
if section =='KSL':
    KSL_tangent_stagger_L1_min = threshold.loc[(threshold['Class'] == section) &
                                               (threshold['Track Type'] == 'Tangent') &
                                               (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
    KSL_tangent_stagger_L2_min = threshold.loc[(threshold['Class'] == section) &
                                               (threshold['Track Type'] == 'Tangent') &
                                               (threshold['Exc Type'] == 'Stagger L2')]['min'].values.item()
    KSL_tangent_stagger_L2_max = threshold.loc[(threshold['Class'] == section) &
                                               (threshold['Track Type'] == 'Tangent') &
                                               (threshold['Exc Type'] == 'Stagger L2')]['max'].values.item()
    KSL_tangent_stagger_L3_min = threshold.loc[(threshold['Class'] == section) &
                                               (threshold['Track Type'] == 'Tangent') &
                                               (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
    KSL_tangent_stagger_L3_max = threshold.loc[(threshold['Class'] == section) &
                                               (threshold['Track Type'] == 'Tangent') &
                                               (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()

    KSL_curve_stagger_L1_min = threshold.loc[(threshold['Class'] == section) &
                                               (threshold['Track Type'] == 'Curve') &
                                               (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
    KSL_curve_stagger_L3_min = threshold.loc[(threshold['Class'] == section) &
                                               (threshold['Track Type'] == 'Curve') &
                                               (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
    KSL_curve_stagger_L3_max = threshold.loc[(threshold['Class'] == section) &
                                               (threshold['Track Type'] == 'Curve') &
                                               (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()
else:
    tangent_stagger_L1_min = threshold.loc[(threshold['Class'] == 'both') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
    tangent_stagger_L2_min = threshold.loc[(threshold['Class'] == 'both') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L2')]['min'].values.item()
    tangent_stagger_L2_max = threshold.loc[(threshold['Class'] == 'both') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L2')]['max'].values.item()
    tangent_stagger_L3_min = threshold.loc[(threshold['Class'] == 'both') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
    tangent_stagger_L3_max = threshold.loc[(threshold['Class'] == 'both') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()

    curve_stagger_L1_min = threshold.loc[(threshold['Class'] == 'both') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
    curve_stagger_L2_min = threshold.loc[(threshold['Class'] == 'both') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L2')]['min'].values.item()
    curve_stagger_L2_max = threshold.loc[(threshold['Class'] == 'both') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L2')]['max'].values.item()
    curve_stagger_L3_min = threshold.loc[(threshold['Class'] == 'both') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
    curve_stagger_L3_max = threshold.loc[(threshold['Class'] == 'both') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()

wear_L1_max = threshold.loc[threshold['Exc Type'] == 'Wire Wear L1']['max'].values.item()
wear_L2_min = threshold.loc[threshold['Exc Type'] == 'Wire Wear L2']['min'].values.item()
wear_L2_max = threshold.loc[threshold['Exc Type'] == 'Wire Wear L2']['max'].values.item()

high_height_L1_min = threshold.loc[threshold['Exc Type'] == 'High Height L1']['min'].values.item()
high_height_L2_min = threshold.loc[threshold['Exc Type'] == 'High Height L2']['min'].values.item()
high_height_L2_max = threshold.loc[threshold['Exc Type'] == 'High Height L2']['max'].values.item()

if section in ['LMC', 'WRL', 'KSL', 'ETSE', 'MOL']:
    class_low_height_L1_max = threshold.loc[(threshold['Class'] == section) & (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
    class_low_height_L2_min = threshold.loc[(threshold['Class'] == section) & (threshold['Exc Type'] == 'Low Height L2')]['min'].values.item()
    class_low_height_L2_max = threshold.loc[(threshold['Class'] == section) & (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()
elif line == 'EAL' and section != 'LMC':
    class_low_height_L1_max = threshold.loc[(threshold['Class'] == 'both') & (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
    class_low_height_L2_min = threshold.loc[(threshold['Class'] == 'both') & (threshold['Exc Type'] == 'Low Height L2')]['min'].values.item()
    class_low_height_L2_max = threshold.loc[(threshold['Class'] == 'both') & (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()


# if section == 'LMC':
#     SHS_LMC_low_height_L1_max = threshold.loc[(threshold['Location Type'] == 'SHS-LMC') & (threshold['Exc Type'] == 'Low Height L1')]['max'].values.item()
#     SHS_LMC_low_height_L2_min = threshold.loc[(threshold['Location Type'] == 'SHS-LMC') & (threshold['Exc Type'] == 'Low Height L2')]['min'].values.item()
#     SHS_LMC_low_height_L2_max = threshold.loc[(threshold['Location Type'] == 'SHS-LMC') & (threshold['Exc Type'] == 'Low Height L2')]['max'].values.item()

# ---------- load line related threshold from metadata ----------


# ---------- low height exception ----------
# WH_cleaned_min['L2'] = (WH_cleaned_min.maxValue <= (SHS_LMC_low_height_L2_max if section == 'LMC' else HUH_LOW_low_height_L2_max))
WH_cleaned_min['L2'] = (WH_cleaned_min.maxValue <= class_low_height_L2_max)
WH_cleaned_min['L2_id'] = (WH_cleaned_min.L2 != WH_cleaned_min.L2.shift()).cumsum()
WH_cleaned_min['L2_count'] = WH_cleaned_min.groupby(['L2', 'L2_id']).cumcount(ascending=False) + 1
WH_cleaned_min.loc[~WH_cleaned_min['L2'], 'L2_count'] = 0

if WH_cleaned_min['L2'].any():
    low_height_exception_full = WH_cleaned_min[(WH_cleaned_min['L2_count'] != 0)]

    low_height_exception_full.loc[low_height_exception_full['L2'], 'exception type'] = 'Low Height'

    low_height_exception = low_height_exception_full\
        .groupby(['exception type', 'L2_id'])\
        .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max']})
    low_height_exception.columns = low_height_exception.columns.map('_'.join)
    low_height_exception = low_height_exception\
        .assign(key=1)\
        .merge(low_height_exception_full.assign(key=1), on='key')\
        .query('`maxValue_min` == `maxValue` & `Km`.between(`Km_min`, `Km_max`)', engine='python')\
        .drop(columns=['maxValue_max', 'key', 'maxValue', 'L2_id', 'L2_count', 'height1',
                       'height2', 'height3', 'height4', 'L2'])\
        .rename({'Km': 'maxLocation', 'Km_min': 'startKm', 'Km_max': 'endKm',
                 'maxValue_min': 'maxValue'}, axis=1)\
        .reset_index()\
        .drop('index', axis=1)

    low_height_exception['length'] = low_height_exception['endKm'] - low_height_exception['startKm']
    low_height_exception = low_height_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]\
        .sort_values('startKm')\
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

    # ------ defining alarm level depending on maxValue only ------
    # low_height_exception.loc[(low_height_exception['maxValue'] <= (SHS_LMC_low_height_L1_max if section == 'LMC' else HUH_LOW_low_height_L1_max)), 'level'] = 'L1'
    low_height_exception.loc[(low_height_exception['maxValue'] <= class_low_height_L1_max), 'level'] = 'L1'
    low_height_exception['level'] = low_height_exception['level'].fillna('L2')
    low_height_exception = low_height_exception[
        ['exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']].reset_index()
    low_height_exception['id'] = date + '_' + line + '_' + section + '_' + track + '_' + 'LH' + low_height_exception['index'].astype(str)
    low_height_exception = low_height_exception[
        ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]
    # ------ defining alarm level depending on maxValue only ------

else:
    low_height_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type'])
# ---------- low height exception ----------

# ---------- high height exception ----------
WH_cleaned_max['L2'] = (WH_cleaned_max.maxValue >= high_height_L2_min)
WH_cleaned_max['L2_id'] = (WH_cleaned_max.L2 != WH_cleaned_max.L2.shift()).cumsum()
WH_cleaned_max['L2_count'] = WH_cleaned_max.groupby(['L2', 'L2_id']).cumcount(ascending=False) + 1
WH_cleaned_max.loc[~WH_cleaned_max['L2'], 'L2_count'] = 0

if WH_cleaned_max['L2'].any():
    high_height_exception_full = WH_cleaned_max[(WH_cleaned_max['L2_count'] != 0)]
    high_height_exception_full.loc[high_height_exception_full['L2'], 'exception type'] = 'High Height'

    high_height_exception = high_height_exception_full\
        .groupby(['exception type', 'L2_id'])\
        .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max']})
    high_height_exception.columns = high_height_exception.columns.map('_'.join)
    high_height_exception = high_height_exception\
        .assign(key=1)\
        .merge(high_height_exception_full.assign(key=1), on='key')\
        .query('`maxValue_max` == `maxValue` & `Km`.between(`Km_min`, `Km_max`)', engine='python')\
        .drop(columns=['maxValue_min', 'key', 'maxValue', 'L2_id', 'L2_count', 'height1',
                       'height2', 'height3', 'height4', 'L2'])\
        .rename({'Km': 'maxLocation', 'Km_min': 'startKm', 'Km_max': 'endKm',
                 'maxValue_max': 'maxValue'}, axis=1)\
        .reset_index()\
        .drop('index', axis=1)

    high_height_exception['length'] = high_height_exception['endKm'] - high_height_exception['startKm']
    high_height_exception = high_height_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]\
        .sort_values('startKm') \
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

    # ------ defining alarm level depending on maxValue only ------
    high_height_exception.loc[(high_height_exception['maxValue'] >= high_height_L1_min), 'level'] = 'L1'
    high_height_exception['level'] = high_height_exception['level'].fillna('L2')
    high_height_exception = high_height_exception[['exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']].reset_index()
    high_height_exception['id'] = date + '_' + line + '_' + section + '_' + track + '_' + 'HH' + high_height_exception['index'].astype(str)
    high_height_exception = high_height_exception[
        ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]
    # ------ defining alarm level depending on maxValue only ------

else:
    high_height_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type'])
# ---------- high height exception ----------

# ---------- Wire Wear exception ----------
wear_min['L2'] = (wear_min.maxValue <= wear_L2_max)
wear_min['L2_id'] = (wear_min.L2 != wear_min.L2.shift()).cumsum()
wear_min['L2_count'] = wear_min.groupby(['L2', 'L2_id']).cumcount(ascending=False) + 1
wear_min.loc[~wear_min['L2'], 'L2_count'] = 0

if wear_min['L2'].any():
    wear_exception_full = wear_min[(wear_min['L2_count'] != 0)]
    wear_exception_full.loc[wear_exception_full['L2'], 'exception type'] = 'Wire Wear'

    wear_exception = wear_exception_full\
        .groupby(['exception type', 'L2_id'])\
        .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max']})
    wear_exception.columns = wear_exception.columns.map('_'.join)
    wear_exception = wear_exception\
        .assign(key=1)\
        .merge(wear_exception_full.assign(key=1), on='key')\
        .query('`maxValue_min` == `maxValue` & `Km`.between(`Km_min`, `Km_max`)', engine='python')\
        .drop(columns=['maxValue_max', 'key', 'maxValue', 'L2_id', 'L2_count',
                       'wear1', 'wear2', 'wear3', 'wear4', 'L2'])\
        .rename({'Km': 'maxLocation', 'Km_min': 'startKm', 'Km_max': 'endKm',
                 'maxValue_min': 'maxValue'}, axis=1)\
        .reset_index()\
        .drop('index', axis=1)

    wear_exception['length'] = wear_exception['endKm'] - wear_exception['startKm']
    wear_exception = wear_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]\
        .sort_values('startKm') \
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

    # ------ defining alarm level depending on maxValue only ------
    wear_exception.loc[(wear_exception['maxValue'] <= wear_L1_max), 'level'] = 'L1'
    wear_exception['level'] = wear_exception['level'].fillna('L2')
    wear_exception = wear_exception[['exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']].reset_index()
    wear_exception['id'] = date + '_' + line + '_' + section + '_' + track + '_' + 'W' + wear_exception['index'].astype(str)
    wear_exception = wear_exception[
        ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]
    # ------ defining alarm level depending on maxValue only ------

else:
    wear_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type'])
# ---------- Wire Wear exception ----------

# ---------- Left stagger exception ----------
stagger_left['curve_L3'] = ((stagger_left.maxValue >= (KSL_curve_stagger_L3_min if section == 'KSL' else curve_stagger_L3_min)) & (stagger_left['track type'] == 'Curve'))
stagger_left['tangent_L3'] = ((stagger_left.maxValue >= (KSL_tangent_stagger_L3_min if section == 'KSL' else tangent_stagger_L3_min)) & (stagger_left['track type'] == 'Tangent'))
stagger_left['curve_L3_id'] = (stagger_left.curve_L3 != stagger_left.curve_L3.shift()).cumsum()
stagger_left['tangent_L3_id'] = (stagger_left.tangent_L3 != stagger_left.tangent_L3.shift()).cumsum()
stagger_left['curve_L3_count'] = stagger_left.groupby(['curve_L3', 'curve_L3_id']).cumcount(ascending=False) + 1
stagger_left['tangent_L3_count'] = stagger_left.groupby(['tangent_L3', 'tangent_L3_id']).cumcount(ascending=False) + 1
stagger_left.loc[~stagger_left['curve_L3'], 'curve_L3_count'] = 0
stagger_left.loc[~stagger_left['tangent_L3'], 'tangent_L3_count'] = 0

if stagger_left['curve_L3'].any() or stagger_left['tangent_L3'].any():

    stagger_left_exception_full = stagger_left[(stagger_left['curve_L3_count'] != 0)
                                    | (stagger_left['tangent_L3_count'] != 0)]

    stagger_left_exception_full.loc[stagger_left_exception_full['curve_L3'], 'exception type'] = 'Stagger'
    stagger_left_exception_full.loc[stagger_left_exception_full['tangent_L3'], 'exception type'] = 'Stagger'

    stagger_left_exception = stagger_left_exception_full\
        .groupby(['exception type', 'curve_L3_id', 'tangent_L3_id'])\
        .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max']})
    stagger_left_exception.columns = stagger_left_exception.columns.map('_'.join)
    stagger_left_exception = stagger_left_exception\
        .assign(key=1)\
        .merge(stagger_left_exception_full.assign(key=1), on='key')\
        .query('`maxValue_max` == `maxValue` & `Km`.between(`Km_min`, `Km_max`)', engine='python')\
        .drop(columns=['maxValue_min', 'key', 'maxValue', 'stagger1', 'stagger2', 'stagger3', 'stagger4', 'curve_L3',
                       'tangent_L3', 'curve_L3_id', 'tangent_L3_id', 'curve_L3_count', 'tangent_L3_count'])\
        .rename({'Km': 'maxLocation', 'Km_min': 'startKm', 'Km_max': 'endKm', 'maxValue_max': 'maxValue'}, axis=1)\
        .reset_index()\
        .drop('index', axis=1)

    stagger_left_exception['length'] = stagger_left_exception['endKm'] - stagger_left_exception['startKm']
    stagger_left_exception = stagger_left_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]\
        .sort_values('startKm') \
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

    # ------ defining alarm level depending on maxValue only ------
    if section != 'KSL':
        stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= curve_stagger_L2_min) & (
                    stagger_left_exception['maxValue'] < curve_stagger_L2_max) & (
                                                stagger_left_exception['track type'] == 'Curve'), 'level'] = 'L2'
    else:
        pass
    stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= (KSL_curve_stagger_L1_min if section == 'KSL' else curve_stagger_L1_min)) & (
                stagger_left_exception['track type'] == 'Curve'), 'level'] = 'L1'
    stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= (KSL_tangent_stagger_L2_min if section == 'KSL' else tangent_stagger_L2_min)) & (
                stagger_left_exception['maxValue'] < (KSL_tangent_stagger_L2_max if section == 'KSL' else tangent_stagger_L2_max)) & (
                                            stagger_left_exception['track type'] == 'Tangent'), 'level'] = 'L2'
    stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= (KSL_tangent_stagger_L1_min if section == 'KSL' else tangent_stagger_L1_min)) & (
                stagger_left_exception['track type'] == 'Tangent'), 'level'] = 'L1'

    stagger_left_exception['level'] = stagger_left_exception['level'].fillna('L3')
    stagger_left_exception = stagger_left_exception[
        ['exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']].reset_index()
    stagger_left_exception['id'] = date + '_' + line + '_' + section + '_' + track + '_' + 'SL' + stagger_left_exception['index'].astype(str)
    stagger_left_exception = stagger_left_exception[
        ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]
    # ------ defining alarm level depending on maxValue only ------

else:
    stagger_left_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type'])
# ---------- Left stagger exception ----------

# ---------- Right stagger exception ----------
stagger_right['curve_L3'] = ((stagger_right.maxValue <= -(KSL_curve_stagger_L3_min if section == 'KSL' else curve_stagger_L3_min)) & (stagger_right['track type'] == 'Curve'))
stagger_right['tangent_L3'] = ((stagger_right.maxValue <= -(KSL_tangent_stagger_L3_min if section == 'KSL' else tangent_stagger_L3_min)) & (stagger_right['track type'] == 'Tangent'))
stagger_right['curve_L3_id'] = (stagger_right.curve_L3 != stagger_right.curve_L3.shift()).cumsum()
stagger_right['tangent_L3_id'] = (stagger_right.tangent_L3 != stagger_right.tangent_L3.shift()).cumsum()
stagger_right['curve_L3_count'] = stagger_right.groupby(['curve_L3', 'curve_L3_id']).cumcount(ascending=False) + 1
stagger_right['tangent_L3_count'] = stagger_right.groupby(['tangent_L3', 'tangent_L3_id']).cumcount(ascending=False) + 1
stagger_right.loc[~stagger_right['curve_L3'], 'curve_L3_count'] = 0
stagger_right.loc[~stagger_right['tangent_L3'], 'tangent_L3_count'] = 0

if stagger_right['curve_L3'].any() or stagger_right['tangent_L3'].any():

    stagger_right_exception_full = stagger_right[(stagger_right['curve_L3_count'] != 0)
                                    | (stagger_right['tangent_L3_count'] != 0)]

    stagger_right_exception_full.loc[stagger_right_exception_full['curve_L3'], 'exception type'] = 'Stagger'
    stagger_right_exception_full.loc[stagger_right_exception_full['tangent_L3'], 'exception type'] = 'Stagger'

    stagger_right_exception = stagger_right_exception_full\
        .groupby(['exception type', 'curve_L3_id', 'tangent_L3_id'])\
        .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max']})
    stagger_right_exception.columns = stagger_right_exception.columns.map('_'.join)
    stagger_right_exception = stagger_right_exception\
        .assign(key=1)\
        .merge(stagger_right_exception_full.assign(key=1), on='key')\
        .query('`maxValue_min` == `maxValue` & `Km`.between(`Km_min`, `Km_max`)', engine='python')\
        .drop(columns=['maxValue_max', 'key', 'maxValue', 'stagger1', 'stagger2', 'stagger3', 'stagger4',
                       'curve_L3', 'tangent_L3', 'curve_L3_id', 'tangent_L3_id', 'curve_L3_count', 'tangent_L3_count'])\
        .rename({'Km': 'maxLocation', 'Km_min': 'startKm', 'Km_max': 'endKm', 'maxValue_min': 'maxValue'}, axis=1)\
        .reset_index()\
        .drop('index', axis=1)

    stagger_right_exception['length'] = stagger_right_exception['endKm'] - stagger_right_exception['startKm']
    stagger_right_exception = stagger_right_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]\
        .sort_values('startKm') \
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

    # ------ defining alarm level depends on maxValue only ------
    if section != 'KSL':
        stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -curve_stagger_L2_min) & (
                    stagger_right_exception['maxValue'] > -curve_stagger_L2_max) & (
                                                stagger_right_exception['track type'] == 'Curve'), 'level'] = 'L2'
    else:
        pass
    stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -(KSL_curve_stagger_L1_min if section == 'KSL' else curve_stagger_L1_min)) & (
                stagger_right_exception['track type'] == 'Curve'), 'level'] = 'L1'
    stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -(KSL_tangent_stagger_L2_min if section == 'KSL' else tangent_stagger_L2_min)) & (
                stagger_right_exception['maxValue'] > -(KSL_tangent_stagger_L2_max if section == 'KSL' else tangent_stagger_L2_max)) & (
                                            stagger_right_exception['track type'] == 'Tangent'), 'level'] = 'L2'
    stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -(KSL_tangent_stagger_L1_min if section == 'KSL' else tangent_stagger_L1_min)) & (
                stagger_right_exception['track type'] == 'Tangent'), 'level'] = 'L1'
    stagger_right_exception['level'] = stagger_right_exception['level'].fillna('L3')
    stagger_right_exception = stagger_right_exception[
        ['exception type', 'level', 'startKm', 'endKm', 'length',
         'maxValue', 'maxLocation', 'track type']].reset_index()
    stagger_right_exception['id'] = date + '_' + line + '_' + section + '_' + track + '_' + 'SR' + stagger_right_exception['index'].astype(str)
    stagger_right_exception = stagger_right_exception[
        ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type']]
    # ------ defining alarm level depends on maxValue only ------

else:
    stagger_right_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type'])
# ---------- Right stagger exception ----------

# ---------- indicate exceptions within overlap section ----------
hh_overlap_id = high_height_exception\
    .assign(key=1).merge(overlap.assign(key=1), on='key')\
    .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
    .reset_index(drop=True)[['id', 'Overlap', 'Tension Length']]

high_height_exception = pd.merge(high_height_exception, hh_overlap_id, on='id', how='outer')

lh_overlap_id = low_height_exception\
    .assign(key=1).merge(overlap.assign(key=1), on='key')\
    .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
    .reset_index(drop=True)[['id', 'Overlap', 'Tension Length']]

low_height_exception = pd.merge(low_height_exception, lh_overlap_id, on='id', how='outer')

left_overlap_id = stagger_left_exception\
    .assign(key=1).merge(overlap.assign(key=1), on='key')\
    .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
    .reset_index(drop=True)[['id', 'Overlap', 'Tension Length']]

stagger_left_exception = pd.merge(stagger_left_exception, left_overlap_id, on='id', how='outer')

right_overlap_id = stagger_right_exception\
    .assign(key=1).merge(overlap.assign(key=1), on='key')\
    .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
    .reset_index(drop=True)[['id', 'Overlap', 'Tension Length']]

stagger_right_exception = pd.merge(stagger_right_exception, right_overlap_id, on='id', how='outer')

wear_overlap_id = wear_exception\
    .assign(key=1).merge(overlap.assign(key=1), on='key')\
    .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
    .reset_index(drop=True)[['id', 'Overlap', 'Tension Length']]

wear_exception = pd.merge(wear_exception, wear_overlap_id, on='id', how='outer')
# ---------- indicate exceptions within overlap section ----------

# ---------- indicate Landmark for wear + stagger exceptions ----------
hh_landmark_id = high_height_exception\
    .assign(key=1).merge(landmark.assign(key=1), on='key')\
    .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
    .reset_index(drop=True)[['id', 'Landmark']]

high_height_exception = pd.merge(high_height_exception, hh_landmark_id, on='id', how='outer')
high_height_exception['Landmark'].fillna(value='typical', inplace=True)
high_height_exception = high_height_exception[['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue',
                                               'maxLocation', 'track type', 'Overlap', 'Tension Length', 'Landmark']]

lh_landmark_id = low_height_exception\
    .assign(key=1).merge(landmark.assign(key=1), on='key')\
    .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
    .reset_index(drop=True)[['id', 'Landmark']]

low_height_exception = pd.merge(low_height_exception, lh_landmark_id, on='id', how='outer')
low_height_exception['Landmark'].fillna(value='typical', inplace=True)
low_height_exception = low_height_exception[['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue',
                                               'maxLocation', 'track type', 'Overlap', 'Tension Length', 'Landmark']]

left_landmark_id = stagger_left_exception\
    .assign(key=1).merge(landmark.assign(key=1), on='key')\
    .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
    .reset_index(drop=True)[['id', 'Landmark']]

stagger_left_exception = pd.merge(stagger_left_exception, left_landmark_id, on='id', how='outer')
stagger_left_exception['Landmark'].fillna(value='typical', inplace=True)
stagger_left_exception = stagger_left_exception[['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue',
                                               'maxLocation', 'track type', 'Overlap', 'Tension Length', 'Landmark']]

right_landmark_id = stagger_right_exception\
    .assign(key=1).merge(landmark.assign(key=1), on='key')\
    .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
    .reset_index(drop=True)[['id', 'Landmark']]

stagger_right_exception = pd.merge(stagger_right_exception, right_landmark_id, on='id', how='outer')
stagger_right_exception['Landmark'].fillna(value='typical', inplace=True)
stagger_right_exception = stagger_right_exception[['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue',
                                               'maxLocation', 'track type', 'Overlap', 'Tension Length', 'Landmark']]

wear_landmark_id = wear_exception\
    .assign(key=1).merge(landmark.assign(key=1), on='key')\
    .query('`maxLocation`.between(`FromKM`, `ToKM`)', engine='python')\
    .reset_index(drop=True)[['id', 'Landmark']]

wear_exception = pd.merge(wear_exception, wear_landmark_id, on='id', how='outer')
wear_exception['Landmark'].fillna(value='typical', inplace=True)
wear_exception = wear_exception[['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue',
                                               'maxLocation', 'track type', 'Overlap', 'Tension Length', 'Landmark']]
# ---------- indicate Landmark for wear + stagger exceptions ----------

# ---------- output results ----------
print('Saving as Excel at', datetime.now())
print('Done!')
input('Press Enter to select save location')
directory = filedialog.askdirectory()

with pd.ExcelWriter(directory + '/' + date + '_' + line + '_' + track + '_' + section.replace('-', '_') + '_' + 'Exception Report.xlsx') as writer:
    wear_exception.to_excel(writer, sheet_name='wear exception', index=False)
    low_height_exception.to_excel(writer, sheet_name='low height exception', index=False)
    high_height_exception.to_excel(writer, sheet_name='high height exception', index=False)
    stagger_left_exception.to_excel(writer, sheet_name='stagger left exception', index=False)
    stagger_right_exception.to_excel(writer, sheet_name='stagger right exception', index=False)
# ---------- output results ----------

# os.startfile(directory + '/' + date + '_' + line + '_' + section.replace('-', '_') + '_' + track + '_' + 'Exception Report.xlsx')