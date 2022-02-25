import pandas as pd
import plotly.express as px
import plotly.subplots as sp
import plotly.graph_objects as go

# -------------- Read raw data --------------

raw = pd.read_csv('20211124_TWL_DT.csv').rename({'Km': 'chainage',
                                                 'StaggerWire1 [mm]': 'stagger1',
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

raw['chainage'] = raw['chainage']\
    .round(decimals=3)

stagger_condensed = raw.groupby(raw.chainage)['stagger1', 'stagger2', 'stagger3', 'stagger4']\
    .agg([lambda x: max(x, key=abs)])
stagger_condensed.columns = ['stagger1', 'stagger2', 'stagger3', 'stagger4']

wear_condensed = raw.groupby(raw.chainage)['wear1', 'wear2', 'wear3', 'wear4'].agg([lambda x: min(x, key=abs)])
wear_condensed.columns = ['wear1', 'wear2', 'wear3', 'wear4']

height_condensed = raw.groupby(raw.chainage).agg({'height1': ['min', 'max'], 'height2': ['min', 'max'], 'height3': ['min', 'max'], 'height4': ['min', 'max']})
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

fig.update_layout(xaxis3_rangeslider_visible=True,
                  xaxis3_rangeslider_thickness=0.05,
                  hovermode="x unified",
                  title="TOV data Visualizer")
fig.show()
