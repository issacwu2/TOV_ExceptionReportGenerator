import pandas as pd
import os
import re
from datetime import datetime
import time
import tkinter as tk
from tkinter import filedialog
pd.options.mode.chained_assignment = None

print('================ TOV1050 Exception Report Generator ================')
print('Before you start, please go through the following instructions or otherwise you might experience errors and the results could be inaccurate!')
input("Press Enter to continue...")
print('0. Make sure all metadata.xlsx of each line are in the same folder with this script.')
print('If not, please move them together and restart the program.')
input("Press Enter to continue...")
print('When the program is running, MAKE SURE YOU:')
print('1. Input the name of Line e.g. TWL/KTL/AEL/TCL/DRL/ISL...')
print('2. Input the direction of track, i.e. UT/DT')
print('3. Input the chainage shift in Km as you marked on the TOV checklist after the TOV run')
print('4. Select ONLY the csv version of the data report to generate exception report')
print('5. Input the date of when that TOV data is obtained, in YYYY/MM/DD')
input("Press Enter to continue...")
print('=========================== IMPORTANT!!! ===========================')
input('Make sure you read through the above instructions carefully and press enter to start... ')

line = input('Line: ')
while not line in ['KTL', 'TWL', 'ISL', 'TKL', 'AEL', 'TCL', 'DRL']:
    print('Please make sure you input correct line:')
    line = input('line:')

if line in ['AEL', 'TCL']:
    section = input('Please input the section range, e.g. TSY-HOK:')
    while not re.match('^[A-Z]{3}-[A-Z]{3}', section):
        print('Please make sure you input correct section range, e.g. TSY-HOK')
        section = input('Please input the section range:')

if line is 'DRL':
    track = input('UT/PL? : ')
    while not track in ['PL', 'UT']:
        print('Please make sure you input either UT / PL')
        track = input('UT/PL? : ')
else:
    track = input('UT/DT? : ')
    while not track in ['DT', 'UT']:
        print('Please make sure you input either UT / DT')
        track = input('UT/DT? : ')

chainage_shift = float(input('Chainage Shift? (in Km and with +/- sign) : '))

d = input('Date (YYYY/MM/DD) : ')
while not re.match('^\d{4}\/\d{2}\/\d{2}$', d):
    print('Please make sure you in put the date in correct format [YYYY/MM/DD]')
    d = input('Date (YYYY/MM/DD) : ')

date = d.replace('/', '')

print('select the data report in csv format after 3 seconds...')
for i in range(3,0,-1):
    print(f"{i}", end="\r", flush=True)
    time.sleep(1)
# ---------- allow user to select csv files -------
root = tk.Tk()
root.withdraw()
raw_data_path = filedialog.askopenfilename()
raw = pd.read_csv(raw_data_path, low_memory=False, encoding='cp1252')
print('Selected: ' + os.path.basename(raw_data_path))
# ---------- allow user to select csv files -------

# --------- Load metadata ----------
track_type = pd.read_excel(line + ' metadata.xlsx', sheet_name=track + ' track type')
location_type = pd.read_excel(line + ' metadata.xlsx', sheet_name='location type')
threshold = pd.read_excel(line + ' metadata.xlsx', sheet_name='threshold')
# --------- Load metadata ----------
print('Loading...')

raw['Km'] = raw['Km']\
    .round(decimals=3) + chainage_shift
# set to chainage accuracy to 0.001km = 1m
# shift chainage according to input

raw = raw[['Km', 'HeightWire1 [mm]', 'HeightWire2 [mm]', 'HeightWire3 [mm]', 'HeightWire4 [mm]',
           'StaggerWire1 [mm]', 'StaggerWire2 [mm]', 'StaggerWire3 [mm]', 'StaggerWire4 [mm]',
           'WearWire1 [mm]', 'WearWire2 [mm]', 'WearWire3 [mm]', 'WearWire4 [mm]']]

raw = raw.rename({'StaggerWire1 [mm]': 'stagger1',
                  'StaggerWire2 [mm]': 'stagger2',
                  'StaggerWire3 [mm]': 'stagger3',
                  'StaggerWire4 [mm]': 'stagger4',
                  'WearWire1 [mm]': 'wear1',
                  'WearWire2 [mm]': 'wear2',
                  'WearWire3 [mm]': 'wear3',
                  'WearWire4 [mm]': 'wear4',
                  'HeightWire1 [mm]': 'height1',
                  'HeightWire2 [mm]': 'height2',
                  'HeightWire3 [mm]': 'height3',
                  'HeightWire4 [mm]': 'height4'}, axis=1)


# ---------- To remove non-float data in float columns ----------
# for column in raw[
# 	['height1', 'height2', 'height3', 'height4', 'wear1', 'wear2', 'wear3', 'wear4', 'stagger1', 'stagger2', 'stagger3',
# 	 'stagger4']].select_dtypes(include=[object]).columns:
# 	raw = raw[~raw[column].str.contains('[!@#$%^&*()a-zA-Z!@#$%^&*()]+$', na=False)]

raw = raw.replace('1.#IO', '')

raw[['height1', 'height2', 'height3', 'height4', 'wear1', 'wear2', 'wear3', 'wear4', 'stagger1', 'stagger2', 'stagger3',
	 'stagger4']] = raw[['height1', 'height2', 'height3', 'height4', 'wear1', 'wear2', 'wear3', 'wear4', 'stagger1', 'stagger2', 'stagger3',
	 'stagger4']].apply(pd.to_numeric)
# ---------- To remove non-float data in float columns ----------


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


# ---------- identify track type (tangent/curve) and location type (open/tunnel) ----------
stagger_left = stagger_left \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True) \
    .assign(key=1) \
    .merge(location_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

stagger_right = stagger_right \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True) \
    .assign(key=1) \
    .merge(location_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

wear_min = wear_min \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True) \
    .assign(key=1) \
    .merge(location_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

WH_cleaned_max = WH_cleaned_max \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True) \
    .assign(key=1) \
    .merge(location_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

WH_cleaned_min = WH_cleaned_min \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True) \
    .assign(key=1) \
    .merge(location_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)', engine='python') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)
# ---------- identify track type (tangent/curve) and location type (open/tunnel) ----------


# ---------- find extreme value amongst 4 channels ----------
WH_cleaned_min['maxValue'] = WH_cleaned_min[['height1', 'height2', 'height3', 'height4']].min(axis=1)
WH_cleaned_max['maxValue'] = WH_cleaned_max[['height1', 'height2', 'height3', 'height4']].max(axis=1)

wear_min['maxValue'] = wear_min[['wear1', 'wear2', 'wear3', 'wear4']].min(axis=1)

stagger_left['maxValue'] = stagger_left[['stagger1', 'stagger2', 'stagger3', 'stagger4']].max(axis=1)
stagger_right['maxValue'] = stagger_right[['stagger1', 'stagger2', 'stagger3', 'stagger4']].min(axis=1)
# ---------- find extreme value amongst 4 channels ----------


# ---------- load line related threshold from metadata ----------
open_tangent_stagger_L1_min = threshold.loc[(threshold['Location Type'] == 'Open') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
open_tangent_stagger_L2_min = threshold.loc[(threshold['Location Type'] == 'Open') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L2')]['min'].values.item()
open_tangent_stagger_L2_max = threshold.loc[(threshold['Location Type'] == 'Open') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L2')]['max'].values.item()
open_tangent_stagger_L3_min = threshold.loc[(threshold['Location Type'] == 'Open') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
open_tangent_stagger_L3_max = threshold.loc[(threshold['Location Type'] == 'Open') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()

tunnel_tangent_stagger_L1_min = threshold.loc[(threshold['Location Type'] == 'Tunnel') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
tunnel_tangent_stagger_L2_min = threshold.loc[(threshold['Location Type'] == 'Tunnel') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L2')]['min'].values.item()
tunnel_tangent_stagger_L2_max = threshold.loc[(threshold['Location Type'] == 'Tunnel') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L2')]['max'].values.item()
tunnel_tangent_stagger_L3_min = threshold.loc[(threshold['Location Type'] == 'Tunnel') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
tunnel_tangent_stagger_L3_max = threshold.loc[(threshold['Location Type'] == 'Tunnel') & (threshold['Track Type'] == 'Tangent') & (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()

open_curve_stagger_L1_min = threshold.loc[(threshold['Location Type'] == 'Open') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
open_curve_stagger_L2_min = threshold.loc[(threshold['Location Type'] == 'Open') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L2')]['min'].values.item()
open_curve_stagger_L2_max = threshold.loc[(threshold['Location Type'] == 'Open') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L2')]['max'].values.item()
open_curve_stagger_L3_min = threshold.loc[(threshold['Location Type'] == 'Open') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
open_curve_stagger_L3_max = threshold.loc[(threshold['Location Type'] == 'Open') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()

tunnel_curve_stagger_L1_min = threshold.loc[(threshold['Location Type'] == 'Tunnel') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L1')]['min'].values.item()
tunnel_curve_stagger_L2_min = threshold.loc[(threshold['Location Type'] == 'Tunnel') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L2')]['min'].values.item()
tunnel_curve_stagger_L2_max = threshold.loc[(threshold['Location Type'] == 'Tunnel') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L2')]['max'].values.item()
tunnel_curve_stagger_L3_min = threshold.loc[(threshold['Location Type'] == 'Tunnel') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L3')]['min'].values.item()
tunnel_curve_stagger_L3_max = threshold.loc[(threshold['Location Type'] == 'Tunnel') & (threshold['Track Type'] == 'Curve') & (threshold['Exc Type'] == 'Stagger L3')]['max'].values.item()

wear_L1_max = threshold.loc[threshold['Exc Type'] == 'Wire Wear L1']['max'].values.item()
wear_L2_min = threshold.loc[threshold['Exc Type'] == 'Wire Wear L2']['min'].values.item()
wear_L2_max = threshold.loc[threshold['Exc Type'] == 'Wire Wear L2']['max'].values.item()

high_height_L1_min = threshold.loc[threshold['Exc Type'] == 'High Height L1']['min'].values.item()
high_height_L2_min = threshold.loc[threshold['Exc Type'] == 'High Height L2']['min'].values.item()
high_height_L2_max = threshold.loc[threshold['Exc Type'] == 'High Height L2']['max'].values.item()

low_height_L1_max = threshold.loc[threshold['Exc Type'] == 'Low Height L1']['max'].values.item()
low_height_L2_min = threshold.loc[threshold['Exc Type'] == 'Low Height L2']['min'].values.item()
low_height_L2_max = threshold.loc[threshold['Exc Type'] == 'Low Height L2']['max'].values.item()
# ---------- load line related threshold from metadata ----------


# ---------- low height exception ----------
WH_cleaned_min['L2'] = (WH_cleaned_min.maxValue <= low_height_L2_max)
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
    low_height_exception = low_height_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type']]\
        .sort_values('startKm')\
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

    # ------ defining alarm level depending on maxValue only ------
    low_height_exception.loc[(low_height_exception['maxValue'] <= low_height_L1_max), 'level'] = 'L1'
    low_height_exception['level'] = low_height_exception['level'].fillna('L2')
    low_height_exception = low_height_exception[
        ['exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type',
         'location type']].reset_index()
    low_height_exception['id'] = date + '_' + line + '_' + track + '_' + 'LH' + low_height_exception['index'].astype(str)
    low_height_exception = low_height_exception[
        ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type',
         'location type']]
    # ------ defining alarm level depending on maxValue only ------

else:
    low_height_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type'])
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
    high_height_exception = high_height_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type']]\
        .sort_values('startKm') \
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

    # ------ defining alarm level depending on maxValue only ------
    high_height_exception.loc[(high_height_exception['maxValue'] >= high_height_L1_min), 'level'] = 'L1'
    high_height_exception['level'] = high_height_exception['level'].fillna('L2')
    high_height_exception = high_height_exception[['exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type',
         'location type']].reset_index()
    high_height_exception['id'] = date + '_' + line + '_' + track + '_' + 'HH' + high_height_exception['index'].astype(str)
    high_height_exception = high_height_exception[
        ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type',
         'location type']]
    # ------ defining alarm level depending on maxValue only ------

else:
    high_height_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type'])
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
    wear_exception = wear_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type']]\
        .sort_values('startKm') \
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

    # ------ defining alarm level depending on maxValue only ------
    wear_exception.loc[(wear_exception['maxValue'] <= wear_L1_max), 'level'] = 'L1'
    wear_exception['level'] = wear_exception['level'].fillna('L2')
    wear_exception = wear_exception[['exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type',
         'location type']].reset_index()
    wear_exception['id'] = date + '_' + line + '_' + track + '_' + 'W' + wear_exception['index'].astype(str)
    wear_exception = wear_exception[
        ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type',
         'location type']]
    # ------ defining alarm level depending on maxValue only ------

else:
    wear_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type'])
# ---------- Wire Wear exception ----------

# ---------- Left stagger exception ----------
stagger_left['open_curve_L3'] = ((stagger_left.maxValue >= open_curve_stagger_L3_min) & (stagger_left['track type'] == 'Curve') & (stagger_left['location type'] != 'Tunnel'))
stagger_left['open_tangent_L3'] = ((stagger_left.maxValue >= open_tangent_stagger_L3_min) & (stagger_left['track type'] == 'Tangent') & (stagger_left['location type'] != 'Tunnel'))
stagger_left['tunnel_L3'] = ((stagger_left.maxValue >= tunnel_curve_stagger_L3_min) & (stagger_left['location type'] == 'Tunnel'))
stagger_left['open_curve_L3_id'] = (stagger_left.open_curve_L3 != stagger_left.open_curve_L3.shift()).cumsum()
stagger_left['open_tangent_L3_id'] = (stagger_left.open_tangent_L3 != stagger_left.open_tangent_L3.shift()).cumsum()
stagger_left['tunnel_L3_id'] = (stagger_left.tunnel_L3 != stagger_left.tunnel_L3.shift()).cumsum()
stagger_left['open_curve_L3_count'] = stagger_left.groupby(['open_curve_L3', 'open_curve_L3_id']).cumcount(ascending=False) + 1
stagger_left['open_tangent_L3_count'] = stagger_left.groupby(['open_tangent_L3', 'open_tangent_L3_id']).cumcount(ascending=False) + 1
stagger_left['tunnel_L3_count'] = stagger_left.groupby(['tunnel_L3', 'tunnel_L3_id']).cumcount(ascending=False) + 1
stagger_left.loc[~stagger_left['open_curve_L3'], 'open_curve_L3_count'] = 0
stagger_left.loc[~stagger_left['open_tangent_L3'], 'open_tangent_L3_count'] = 0
stagger_left.loc[~stagger_left['tunnel_L3'], 'tunnel_L3_count'] = 0

if stagger_left['open_curve_L3'].any() \
        or stagger_left['open_tangent_L3'].any() \
        or stagger_left['tunnel_L3'].any():

    stagger_left_exception_full = stagger_left[(stagger_left['open_curve_L3_count'] != 0)
                                    | (stagger_left['open_tangent_L3_count'] != 0)
                                    | (stagger_left['tunnel_L3_count'] != 0)]

    stagger_left_exception_full.loc[stagger_left_exception_full['open_curve_L3'], 'exception type'] = 'Stagger'
    stagger_left_exception_full.loc[stagger_left_exception_full['open_tangent_L3'], 'exception type'] = 'Stagger'
    stagger_left_exception_full.loc[stagger_left_exception_full['tunnel_L3'], 'exception type'] = 'Stagger'

    stagger_left_exception = stagger_left_exception_full\
        .groupby(['exception type', 'open_curve_L3_id', 'open_tangent_L3_id', 'tunnel_L3_id'])\
        .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max']})
    stagger_left_exception.columns = stagger_left_exception.columns.map('_'.join)
    stagger_left_exception = stagger_left_exception\
        .assign(key=1)\
        .merge(stagger_left_exception_full.assign(key=1), on='key')\
        .query('`maxValue_max` == `maxValue` & `Km`.between(`Km_min`, `Km_max`)', engine='python')\
        .drop(columns=['maxValue_min', 'key', 'maxValue', 'stagger1', 'stagger2', 'stagger3', 'stagger4', 'open_curve_L3',
                       'open_tangent_L3', 'tunnel_L3', 'open_curve_L3_id', 'open_tangent_L3_id', 'tunnel_L3_id',
                       'open_curve_L3_count', 'open_tangent_L3_count', 'tunnel_L3_count'
                       ])\
        .rename({'Km': 'maxLocation', 'Km_min': 'startKm', 'Km_max': 'endKm', 'maxValue_max': 'maxValue'}, axis=1)\
        .reset_index()\
        .drop('index', axis=1)

    stagger_left_exception['length'] = stagger_left_exception['endKm'] - stagger_left_exception['startKm']
    stagger_left_exception = stagger_left_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type']]\
        .sort_values('startKm') \
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

    # ------ defining alarm level depending on maxValue only ------
    stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= open_curve_stagger_L2_min) & (
                stagger_left_exception['maxValue'] <open_curve_stagger_L2_max) & (
                                            stagger_left_exception['track type'] == 'Curve') & (
                                            stagger_left_exception['location type'] != 'Tunnel'), 'level'] = 'L2'
    stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= open_curve_stagger_L1_min) & (
                stagger_left_exception['track type'] == 'Curve') & (
                                            stagger_left_exception['location type'] != 'Tunnel'), 'level'] = 'L1'
    stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= open_tangent_stagger_L2_min) & (
                stagger_left_exception['maxValue'] <open_tangent_stagger_L2_max) & (
                                            stagger_left_exception['track type'] == 'Tangent') & (
                                            stagger_left_exception['location type'] != 'Tunnel'), 'level'] = 'L2'
    stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= open_tangent_stagger_L1_min) & (
                stagger_left_exception['track type'] == 'Tangent') & (
                                            stagger_left_exception['location type'] != 'Tunnel'), 'level'] = 'L1'
    stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= tunnel_curve_stagger_L2_min) & (
                stagger_left_exception['maxValue'] <tunnel_curve_stagger_L2_max) & (
                                            stagger_left_exception['location type'] == 'Tunnel'), 'level'] = 'L2'
    stagger_left_exception.loc[(stagger_left_exception['maxValue'] >= tunnel_curve_stagger_L1_min) & (
                stagger_left_exception['location type'] == 'Tunnel'), 'level'] = 'L1'
    stagger_left_exception['level'] = stagger_left_exception['level'].fillna('L3')
    stagger_left_exception = stagger_left_exception[
        ['exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type',
         'location type']].reset_index()
    stagger_left_exception['id'] = date + '_' + line + '_' + track + '_' + 'SL' + stagger_left_exception['index'].astype(str)
    stagger_left_exception = stagger_left_exception[
        ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type',
         'location type']]
    # ------ defining alarm level depending on maxValue only ------

else:
    stagger_left_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type'])
# ---------- Left stagger exception ----------

# ---------- Right stagger exception ----------
stagger_right['open_curve_L3'] = ((stagger_right.maxValue <= -open_curve_stagger_L3_min) & (stagger_right['track type'] == 'Curve') & (stagger_right['location type'] != 'Tunnel'))
stagger_right['open_tangent_L3'] = ((stagger_right.maxValue <= -open_tangent_stagger_L3_min) & (stagger_right['track type'] == 'Tangent') & (stagger_right['location type'] != 'Tunnel'))
stagger_right['tunnel_L3'] = ((stagger_right.maxValue <= -tunnel_curve_stagger_L3_min) & (stagger_right['location type'] == 'Tunnel'))
stagger_right['open_curve_L3_id'] = (stagger_right.open_curve_L3 != stagger_right.open_curve_L3.shift()).cumsum()
stagger_right['open_tangent_L3_id'] = (stagger_right.open_tangent_L3 != stagger_right.open_tangent_L3.shift()).cumsum()
stagger_right['tunnel_L3_id'] = (stagger_right.tunnel_L3 != stagger_right.tunnel_L3.shift()).cumsum()
stagger_right['open_curve_L3_count'] = stagger_right.groupby(['open_curve_L3', 'open_curve_L3_id']).cumcount(ascending=False) + 1
stagger_right['open_tangent_L3_count'] = stagger_right.groupby(['open_tangent_L3', 'open_tangent_L3_id']).cumcount(ascending=False) + 1
stagger_right['tunnel_L3_count'] = stagger_right.groupby(['tunnel_L3', 'tunnel_L3_id']).cumcount(ascending=False) + 1
stagger_right.loc[~stagger_right['open_curve_L3'], 'open_curve_L3_count'] = 0
stagger_right.loc[~stagger_right['open_tangent_L3'], 'open_tangent_L3_count'] = 0
stagger_right.loc[~stagger_right['tunnel_L3'], 'tunnel_L3_count'] = 0

if stagger_right['open_curve_L3'].any() \
        or stagger_right['open_tangent_L3'].any() \
        or stagger_right['tunnel_L3'].any():

    stagger_right_exception_full = stagger_right[(stagger_right['open_curve_L3_count'] != 0)
                                    | (stagger_right['open_tangent_L3_count'] != 0)
                                    | (stagger_right['tunnel_L3_count'] != 0)]


    stagger_right_exception_full.loc[stagger_right_exception_full['open_curve_L3'], 'exception type'] = 'Stagger'
    stagger_right_exception_full.loc[stagger_right_exception_full['open_tangent_L3'], 'exception type'] = 'Stagger'
    stagger_right_exception_full.loc[stagger_right_exception_full['tunnel_L3'], 'exception type'] = 'Stagger'

    stagger_right_exception = stagger_right_exception_full\
        .groupby(['exception type', 'open_curve_L3_id', 'open_tangent_L3_id', 'tunnel_L3_id'])\
        .agg({'Km': ['min', 'max'], 'maxValue': ['min', 'max']})
    stagger_right_exception.columns = stagger_right_exception.columns.map('_'.join)
    stagger_right_exception = stagger_right_exception\
        .assign(key=1)\
        .merge(stagger_right_exception_full.assign(key=1), on='key')\
        .query('`maxValue_max` == `maxValue` & `Km`.between(`Km_min`, `Km_max`)', engine='python')\
        .drop(columns=['maxValue_min', 'key', 'maxValue', 'stagger1', 'stagger2', 'stagger3', 'stagger4',
                       'open_curve_L3', 'open_tangent_L3', 'tunnel_L3', 'open_curve_L3_id',
                       'open_tangent_L3_id', 'tunnel_L3_id', 'open_curve_L3_count',
                       'open_tangent_L3_count', 'tunnel_L3_count'
                       ])\
        .rename({'Km': 'maxLocation', 'Km_min': 'startKm', 'Km_max': 'endKm', 'maxValue_max': 'maxValue'}, axis=1)\
        .reset_index()\
        .drop('index', axis=1)

    stagger_right_exception['length'] = stagger_right_exception['endKm'] - stagger_right_exception['startKm']
    stagger_right_exception = stagger_right_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type']]\
        .sort_values('startKm') \
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

    # ------ defining alarm level depends on maxValue only ------
    stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -open_curve_stagger_L2_min) & (
                stagger_right_exception['maxValue'] > -open_curve_stagger_L2_max) & (
                                            stagger_right_exception['track type'] == 'Curve') & (
                                            stagger_right_exception['location type'] != 'Tunnel'), 'level'] = 'L2'
    stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -open_curve_stagger_L1_min) & (
                stagger_right_exception['track type'] == 'Curve') & (
                                            stagger_right_exception['location type'] != 'Tunnel'), 'level'] = 'L1'
    stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -open_tangent_stagger_L2_min) & (
                stagger_right_exception['maxValue'] > -open_tangent_stagger_L2_max) & (
                                            stagger_right_exception['track type'] == 'Tangent') & (
                                            stagger_right_exception['location type'] != 'Tunnel'), 'level'] = 'L2'
    stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -open_tangent_stagger_L1_min) & (
                stagger_right_exception['track type'] == 'Tangent') & (
                                            stagger_right_exception['location type'] != 'Tunnel'), 'level'] = 'L1'
    stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -tunnel_curve_stagger_L2_min) & (
                stagger_right_exception['maxValue'] > -tunnel_curve_stagger_L2_max) & (
                                            stagger_right_exception['location type'] == 'Tunnel'), 'level'] = 'L2'
    stagger_right_exception.loc[(stagger_right_exception['maxValue'] <= -tunnel_curve_stagger_L1_min) & (
                stagger_right_exception['location type'] == 'Tunnel'), 'level'] = 'L1'
    stagger_right_exception['level'] = stagger_right_exception['level'].fillna('L3')
    stagger_right_exception = stagger_right_exception[
        ['exception type', 'level', 'startKm', 'endKm', 'length',
         'maxValue', 'maxLocation', 'track type', 'location type']].reset_index()
    stagger_right_exception['id'] = date + '_' + line + '_' + track + '_' + 'SL' + stagger_right_exception['index'].astype(str)
    stagger_right_exception = stagger_right_exception[
        ['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type',
         'location type']]
    # ------ defining alarm level depends on maxValue only ------

else:
    stagger_right_exception = pd.DataFrame(columns=['id', 'exception type', 'level', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type'])
# ---------- Right stagger exception ----------

# ---------- output results ----------
print('Saving as Excel at', datetime.now())
print('Done!')
input('Press Enter to select save location')
directory = filedialog.askdirectory()
if line in ['AEL', 'TCL']:
    line = line + '_' + section.replace('-', '_')
with pd.ExcelWriter(directory + '/' + date + '_' + line + '_' + track + '_' + 'Exception Report.xlsx') as writer:
    wear_exception.to_excel(writer, sheet_name='wear exception', index=False)
    low_height_exception.to_excel(writer, sheet_name='low height exception', index=False)
    high_height_exception.to_excel(writer, sheet_name='high height exception', index=False)
    stagger_left_exception.to_excel(writer, sheet_name='stagger left exception', index=False)
    stagger_right_exception.to_excel(writer, sheet_name='stagger right exception', index=False)
# ---------- output results ----------

os.startfile(directory + '/' + date + '_' + line + '_' + track + '_' + 'Exception Report.xlsx')