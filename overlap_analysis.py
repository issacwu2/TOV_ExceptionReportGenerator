import pandas as pd
import plotly.express as px
import plotly.subplots as sp
import plotly.graph_objects as go
import time
import tkinter as tk
from tkinter import filedialog
pd.options.mode.chained_assignment = None


print('================ Overlap Analysis Tool ================')
print('This is a tool to visualize the TOV raw data')
print('select the data report in csv format after 3 seconds...')
for i in range(3,0,-1):
    print(f"{i}", end="\r", flush=True)
    time.sleep(1)
# ---------- allow user to select csv files -------
root = tk.Tk()
root.withdraw()
raw_data_path = filedialog.askopenfilename()
raw = pd.read_csv(raw_data_path, low_memory=False, encoding='cp1252')
# ---------- allow user to select csv files -------

print('Loading...')

raw['Km'] = raw['Km'].round(decimals=3)

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

raw = raw.replace('1.#IO', '')

raw[['height1', 'height2', 'height3', 'height4', 'wear1', 'wear2', 'wear3', 'wear4', 'stagger1', 'stagger2', 'stagger3',
	 'stagger4']] = raw[['height1', 'height2', 'height3', 'height4', 'wear1', 'wear2', 'wear3', 'wear4', 'stagger1',
                         'stagger2', 'stagger3', 'stagger4']].apply(pd.to_numeric)
# ---------- To remove non-float data in float columns ----------

# ---------- getting extreme stagger value by absolute max (regardless of +/- sign) ----------
stagger_condensed = raw.groupby(raw.Km)[['stagger1', 'stagger2', 'stagger3', 'stagger4']]\
    .agg([lambda x: max(x, key=abs)])
stagger_condensed.columns = ['stagger1', 'stagger2', 'stagger3', 'stagger4']
# ---------- getting extreme stagger value by absolute max (regardless of +/- sign) ----------

# ---------- getting extreme wear value by min value ----------
wear_condensed = raw.groupby(raw.Km)[['wear1', 'wear2', 'wear3', 'wear4']].agg([lambda x: min(x, key=abs)])
wear_condensed.columns = ['wear1', 'wear2', 'wear3', 'wear4']
# ---------- getting extreme wear value by min value ----------

# ---------- getting extreme height value by min or max (conditional) ----------
height_condensed = raw.groupby(raw.Km).agg({'height1': ['min', 'max'], 'height2': ['min', 'max'], 'height3': ['min', 'max'], 'height4': ['min', 'max']})
height_condensed.columns = height_condensed.columns.map('_'.join)
height_condensed.loc[(height_condensed['height1_max'] > 5299), 'height1'] = height_condensed['height1_max']
height_condensed.loc[(height_condensed['height1_max'] <= 5299), 'height1'] = height_condensed['height1_min']
height_condensed.loc[(height_condensed['height1_max'] > 5299), 'height1'] = height_condensed['height1_max']
height_condensed.loc[(height_condensed['height1_max'] <= 5299), 'height1'] = height_condensed['height1_min']
height_condensed.loc[(height_condensed['height2_max'] > 5299), 'height2'] = height_condensed['height2_max']
height_condensed.loc[(height_condensed['height2_max'] <= 5299), 'height2'] = height_condensed['height2_min']
height_condensed.loc[(height_condensed['height3_max'] > 5299), 'height3'] = height_condensed['height3_max']
height_condensed.loc[(height_condensed['height3_max'] <= 5299), 'height3'] = height_condensed['height3_min']
height_condensed.loc[(height_condensed['height4_max'] > 5299), 'height4'] = height_condensed['height4_max']
height_condensed.loc[(height_condensed['height4_max'] <= 5299), 'height4'] = height_condensed['height4_min']
height_condensed = height_condensed[['height1', 'height2', 'height3', 'height4']]
# ---------- getting extreme height value by min or max (conditional) ----------
# explanation: over-height is rare case, only when height is reaching upper bound, we choose max value over min.
# otherwise, min value is preferred


# ---------- creating stacked chart ----------
fig_stagger = px.line(stagger_condensed)
fig_height = px.line(height_condensed)
fig_wear = px.line(wear_condensed)

fig = sp.make_subplots(rows=3, cols=1, vertical_spacing=0.05, shared_xaxes=True,
                       subplot_titles=("Stagger", "CW Height", "CW Wearing"))

for d in fig_stagger.data:
    fig.add_trace((go.Scattergl(x=d['x'], y=d['y'], name=d['name'])), row=1, col=1)

for d in fig_height.data:
    fig.add_trace((go.Scattergl(x=d['x'], y=d['y'], name=d['name'])), row=2, col=1)

for d in fig_wear.data:
    fig.add_trace((go.Scattergl(x=d['x'], y=d['y'], name=d['name'])), row=3, col=1)


title = raw_data_path.split('/')[-1].split('.')[0]

fig.update_layout(xaxis3_rangeslider_visible=True,
                  xaxis3_rangeslider_thickness=0.05,
                  hovermode="x unified",
                  title="TOV data Visualizer" + ' - ' + title)
# ---------- creating stacked chart ----------

fig.show()

