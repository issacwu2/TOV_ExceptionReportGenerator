import pandas as pd
import plotly.express as px
import plotly.subplots as sp
import plotly.graph_objects as go

# -------------- Read raw data --------------

raw = pd.read_excel('20211124_TWL_DT.xlsx', sheet_name='Data Report 1')
raw = raw.rename({'Km': 'chainage',
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


plotDF = raw.set_index('chainage')

# -------------- Read raw data --------------

fig_stagger = px.line(plotDF[['stagger1', 'stagger2', 'stagger3', 'stagger4']])
fig_height = px.line(plotDF[['height1', 'height2', 'height3', 'height4']])
fig_wear = px.line(plotDF[['wear1', 'wear2', 'wear3', 'wear4']])

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
                  hovermode="x unified")

fig.show()
