import pandas as pd
from datetime import datetime

print('read raw data and metadata start:', datetime.now())

raw = pd.read_excel('TWL may-aug-nov data reports/20211124_TWL_DT.xlsx', sheet_name='Data Report 1')

# --------- Load metadata ----------
track_type = pd.read_excel('TWL metadata.xlsx', sheet_name='DT track type')
location_type = pd.read_excel('TWL metadata.xlsx', sheet_name='location type')
threshold = pd.read_excel('TWL metadata.xlsx', sheet_name='threshold')
# --------- Load metadata ----------
print('read raw data and metadata end:', datetime.now())
print('set raw data accuracy to 0.001km start:', datetime.now())

raw['Km'] = raw['Km']\
    .round(decimals=3)\

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
# set to chainage accuracy to 0.001km = 1m
print('set raw data accuracy to 0.001km end:', datetime.now())

print('preprocess raw data start:', datetime.now())
# ---------- To remove unreasonable CW height data ----------
WH_min = raw.groupby('Km')[['height1', 'height2', 'height3', 'height4']].min().reset_index()
WH_min.loc[(WH_min['height1'] < 2000) | (WH_min['height2'] < 2000) | (WH_min['height3'] < 2000) | (WH_min['height4'] < 2000), 'error'] = WH_min['Km']
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
print('preprocess raw data end:', datetime.now())

print('query track and location type start:', datetime.now())
# ---------- identify track type (tangent/curve) and location type (open/tunnel) ----------
stagger_left = stagger_left \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True) \
    .assign(key=1) \
    .merge(location_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

stagger_right = stagger_right \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True) \
    .assign(key=1) \
    .merge(location_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

wear_min = wear_min \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True) \
    .assign(key=1) \
    .merge(location_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

WH_cleaned_max = WH_cleaned_max \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True) \
    .assign(key=1) \
    .merge(location_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)

WH_cleaned_min = WH_cleaned_min \
    .assign(key=1) \
    .merge(track_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True) \
    .assign(key=1) \
    .merge(location_type.assign(key=1), on='key') \
    .query('`Km`.between(`startKM`, `endKM`)') \
    .drop(columns=['startKM', 'endKM', 'key']) \
    .reset_index(drop=True)
# ---------- identify track type (tangent/curve) and location type (open/tunnel) ----------
print('query track and location type end:', datetime.now())

print('final extreme value out of 4 channels start:', datetime.now())
# ---------- find extreme value amongst 4 channels ----------
WH_cleaned_min['maxValue'] = WH_cleaned_min[['height1', 'height2', 'height3', 'height4']].min(axis=1)
WH_cleaned_max['maxValue'] = WH_cleaned_max[['height1', 'height2', 'height3', 'height4']].max(axis=1)

wear_min['maxValue'] = wear_min[['wear1', 'wear2', 'wear3', 'wear4']].min(axis=1)

stagger_left['maxValue'] = stagger_left[['stagger1', 'stagger2', 'stagger3', 'stagger4']].max(axis=1)
stagger_right['maxValue'] = stagger_right[['stagger1', 'stagger2', 'stagger3', 'stagger4']].min(axis=1)
# ---------- find extreme value amongst 4 channels ----------
print('final extreme value out of 4 channels end:', datetime.now())

print('load threshold value start:', datetime.now())
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
print('load threshold value end:', datetime.now())


print('generating exception start:', datetime.now())
# ---------- low height exception ----------
WH_cleaned_min['L2'] = (WH_cleaned_min.maxValue <= low_height_L2_max)
WH_cleaned_min['L2_id'] = (WH_cleaned_min.L2 != WH_cleaned_min.L2.shift()).cumsum()
WH_cleaned_min['L2_count'] = WH_cleaned_min.groupby(['L2', 'L2_id']).cumcount(ascending=False) + 1
WH_cleaned_min.loc[~WH_cleaned_min['L2'], 'L2_count'] = 0
if WH_cleaned_min['L2'].any():
    low_height_exception = WH_cleaned_min[(WH_cleaned_min['L2_count'] != 0)]
    low_height_exception.loc[low_height_exception['L2'], 'exception type'] = 'Low Height'

    low_height_exception = low_height_exception\
        .groupby(['exception type', 'L2_id'])[['Km', 'maxValue']]\
        .min()\
        .rename({'Km': 'startKm'}, axis=1)\
        .join(low_height_exception\
            .groupby(['exception type', 'L2_id'])[['Km']]\
            .max()\
            .rename({'Km': 'endKm'}, axis=1))\
        .assign(key=1)\
        .merge(low_height_exception.assign(key=1), on='key')\
        .query('`maxValue_x` == `maxValue_y` & `Km`.between(`startKm`, `endKm`)')\
        .drop(columns=['key', 'maxValue_y', 'L2_id', 'L2_count', 'height1', 'height2', 'height3', 'height4', 'L2'])\
        .rename({'Km': 'maxLocation', 'maxValue_x': 'maxValue'}, axis=1)\
        .reset_index()\
        .drop('index', axis=1)

    low_height_exception['length'] = low_height_exception['endKm'] - low_height_exception['startKm']
    low_height_exception = low_height_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type']]\
        .sort_values('startKm')\
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

else:
    low_height_exception = pd.DataFrame(columns=['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type'])
# ---------- low height exception ----------

# ---------- high height exception ----------
WH_cleaned_max['L2'] = (WH_cleaned_max.maxValue >= high_height_L2_min)
WH_cleaned_max['L2_id'] = (WH_cleaned_max.L2 != WH_cleaned_max.L2.shift()).cumsum()
WH_cleaned_max['L2_count'] = WH_cleaned_max.groupby(['L2', 'L2_id']).cumcount(ascending=False) + 1
WH_cleaned_max.loc[~WH_cleaned_max['L2'], 'L2_count'] = 0
if WH_cleaned_max['L2'].any():
    high_height_exception = WH_cleaned_max[(WH_cleaned_max['L2_count'] != 0)]
    high_height_exception.loc[high_height_exception['L2'], 'exception type'] = 'High Height'

    high_height_exception = high_height_exception\
        .groupby(['exception type', 'L2_id'])[['Km', 'maxValue']]\
        .min()\
        .rename({'Km': 'startKm'}, axis=1)\
        .join(high_height_exception\
            .groupby(['exception type', 'L2_id'])[['Km']]\
            .max()\
            .rename({'Km': 'endKm'}, axis=1))\
        .assign(key=1)\
        .merge(high_height_exception.assign(key=1), on='key')\
        .query('`maxValue_x` == `maxValue_y` & `Km`.between(`startKm`, `endKm`)')\
        .drop(columns=['key', 'maxValue_y', 'L2_id', 'L2_count', 'height1', 'height2', 'height3', 'height4', 'L2'])\
        .rename({'Km': 'maxLocation', 'maxValue_x': 'maxValue'}, axis=1)\
        .reset_index()\
        .drop('index', axis=1)

    high_height_exception['length'] = high_height_exception['endKm'] - high_height_exception['startKm']
    high_height_exception = high_height_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type']]\
        .sort_values('startKm') \
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

else:
    high_height_exception = pd.DataFrame(columns=['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type'])
# ---------- high height exception ----------

# ---------- Wire Wear exception ----------
wear_min['L2'] = (wear_min.maxValue <= wear_L2_max)
wear_min['L2_id'] = (wear_min.L2 != wear_min.L2.shift()).cumsum()
wear_min['L2_count'] = wear_min.groupby(['L2', 'L2_id']).cumcount(ascending=False) + 1
wear_min.loc[~wear_min['L2'], 'L2_count'] = 0
if wear_min['L2'].any():
    wear_exception = wear_min[(wear_min['L2_count'] != 0)]
    wear_exception.loc[wear_exception['L2'], 'exception type'] = 'Wire Wear'

    wear_exception = wear_exception\
        .groupby(['exception type', 'L2_id'])[['Km', 'maxValue']]\
        .min()\
        .rename({'Km': 'startKm'}, axis=1)\
        .join(wear_exception\
            .groupby(['exception type', 'L2_id'])[['Km']]\
            .max()\
            .rename({'Km': 'endKm'}, axis=1))\
        .assign(key=1)\
        .merge(wear_exception.assign(key=1), on='key')\
        .query('`maxValue_x` == `maxValue_y` & `Km`.between(`startKm`, `endKm`)')\
        .drop(columns=['key', 'maxValue_y', 'L2_id', 'L2_count', 'wear1', 'wear2', 'wear3', 'wear4', 'L2'])\
        .rename({'Km': 'maxLocation', 'maxValue_x': 'maxValue'}, axis=1)\
        .reset_index()\
        .drop('index', axis=1)

    wear_exception['length'] = wear_exception['endKm'] - wear_exception['startKm']
    wear_exception = wear_exception[['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type']]\
        .sort_values('startKm') \
        .drop_duplicates(subset=['exception type', 'startKm', 'endKm'])\
        .reset_index()\
        .drop('index', axis=1)

else:
    wear_exception = pd.DataFrame(columns=['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type'])
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
        .query('`maxValue_max` == `maxValue` & `Km`.between(`Km_min`, `Km_max`)')\
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
    # ------ defining alarm level depending on maxValue only ------

else:
    stagger_left_exception = pd.DataFrame(columns=['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type'])
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
        .query('`maxValue_max` == `maxValue` & `Km`.between(`Km_min`, `Km_max`)')\
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
    # ------ defining alarm level depends on maxValue only ------

else:
    stagger_right_exception = pd.DataFrame(columns=['exception type', 'startKm', 'endKm', 'length', 'maxValue', 'maxLocation', 'track type', 'location type'])


# ---------- Right stagger exception ----------
print('generating exception end:', datetime.now())

# ---------- output results ----------
with pd.ExcelWriter('TWL_nov_new.xlsx') as writer:
    wear_exception.to_excel(writer, sheet_name='wear exception')
    low_height_exception.to_excel(writer, sheet_name='low height exception')
    high_height_exception.to_excel(writer, sheet_name='high height exception')
    stagger_left_exception.to_excel(writer, sheet_name='stagger left exception')
    stagger_right_exception.to_excel(writer, sheet_name='stagger right exception')
# ---------- output results ----------
