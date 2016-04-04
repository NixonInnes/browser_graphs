from xlwings import Workbook, Sheet, Range, Chart
import pandas as pd
from plotly.offline import plot
from plotly.graph_objs import Figure, Layout, Scatter

wb = Workbook('Example_Workbook.xlsx')
graph_title = Range('Dash', 'B2').value
sheetname = Range('Dash', 'B3').value

def new_df(sheetname, startcell='A1'):
    data = Range(sheetname, startcell).table.value
    temp_df = pd.DataFrame(data[1:], columns=data[0])
    return temp_df

df = new_df(sheetname)
X = df.index

def draw_traces(df):
    plots = []
    for d in df.columns:
        plots.append(
            Scatter(
                x=X, 
                y=df[d],
                name=d
            )
        )
    return traces

data = draw_traces(df)
layout = Layout(
    title = graph_title,
    showlegend = True,
    hovermode = 'compare'
)

figure = Figure(data=data, layout=layout)

plot(figure)
