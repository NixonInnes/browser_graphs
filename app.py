from xlwings import Workbook, Sheet, Range, Chart
import pandas as pd
from plotly.offline import plot
from plotly.graph_objs import Figure, Layout, Scatter

wb = Workbook('Phase 1.xlsx')
graph_title = Range('Dash', 'B2').value
sheetname = Range('Dash', 'B3').value
x_axis = Range('Dash', 'B4').value or None


def new_df(sheetname, startcell='A1'):
    data = Range(sheetname, startcell).table.value
    temp_df = pd.DataFrame(data[1:], columns=data[0])
    return temp_df

def draw_traces(df, x=None):
    traces = []
    if x is not None and x in df.columns:
        X = df[x]
    else:
        X = df.index
    for d in [d for d in df.columns if d != x]:
        traces.append(
            Scatter(
                x=X, 
                y=df[d],
                name=d
            )
        )
    return traces

df = new_df(sheetname)
data = draw_traces(df, x=x_axis)

layout = Layout(
    title = graph_title,
    showlegend = True,
    hovermode = 'compare'
)

figure = Figure(data=data, layout=layout)
plot(figure)
