import pandas as pd
import plotly.express as px
import plotly.subplots as sp
import plotly.graph_objects as go

# -------------- Read raw data --------------

plotDF = pd.read_excel('test_plot_data_short.xlsx', sheet_name='data')
plotDF = plotDF.set_index('chainage')

# -------------- Read raw data --------------

fig_stagger = px.line(plotDF[['stagger c1', 'stagger c2', 'stagger c3', 'stagger c4']])
fig_height = px.line(plotDF[['height c1', 'height c2', 'height c3', 'height c4']])
fig_wear = px.line(plotDF[['wear c1', 'wear c2', 'wear c3', 'wear c4']])

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
