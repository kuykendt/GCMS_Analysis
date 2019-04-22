'''
Plotly & Dash Interactive Analysis Tool for processing data from the Thermo-Scientific
GC-MS in Johnson Hall at Oregon State University

Written by: Tyler Kuykendall
Date: April 19, 2019
'''

# Imports Packages to use
import string # Imports String library
import re # Imports Regular Expressions
import numpy as np # Imports Numpy
import pandas as pd # Imports Pandas ()
import seaborn as sns # Imports Seaborn (Graphing/Visualization tool)
import matplotlib.pyplot as plt # Needed to change plot elements
import plotly.offline as pyo # Imports a version of Plotly for offline use
import plotly.graph_objs as go # Imports Plotly graphing library
import dash # Imports Dash stuff
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output # Allows for inputs and outputs

def chemSplit(formula):
    '''
    Takes in the chemical formula as provided by the GC-MS data
    and outputs a list of the C and H numbers in that formula
    '''

    # Splits formula into elements and numbers
    formSplit = re.split('(\d+)', formula)

    # Adds a 1 to the end if there is only one of an element
    if formSplit[len(formSplit) - 1] not in string.digits:
        formSplit.append('1')

    # Separates the elments and quantities into respective lists.
    elem = []
    quant = []
    for i in formSplit:
        if i in string.ascii_letters:
            elem.append(i)
        else:
            quant.append(i)

    # Stores the number of each element present in the formula. If the element
    # is not present it sets the number to 0
    if 'C' in elem:
        cNum = quant[elem.index('C')]
    else:
        cNum = 0

    if 'H' in elem:
        hNum = quant[elem.index('H')]
    else:
        hNum = 0

    if 'O' in elem:
        oNum = quant[elem.index('O')]
    else:
        oNum = 0

    # Returns a list of numbers corresponding to each element
    return [cNum, hNum, oNum]

def processing(df):
    '''
    This function will take in the prepared DataFrame that was constructed
    from the GC data and process it so that we can analyze the data
    quantitatively.
    '''

    # Groups up the alkanes by carbon number
    # Casts the C, H, and O numbers to be integers
    df = df.astype({'C Number': int, 'H Number': int, 'O Number': int})

    # Creates a GroupBy object so that we can operate around carbon numbers
    cNumGroups = df.groupby('C Number')

    # Creates a list of C Numbers and the sum of their sample fractions
    cAreas = cNumGroups['Sample Fraction'].sum()

    return [cAreas.index, cAreas.values]

# Creates a Dash "app" object
app = dash.Dash()

# Creates the layout for the app
app.layout = html.Div([
            dcc.Graph(id='graph'),
            dcc.Dropdown(
                id='dataset',
                options=[
                    {'label':'First Wax', 'value':0},
                    {'label':'April 12 Initial', 'value':1},
                    {'label':'April 12 Final', 'value':2}
                ],
                value=0)
])

# Creates the callback so that the dashboard can be updated
@app.callback(Output('graph',component_property='figure'),
            [Input(component_id='dataset', component_property='value')])
def update_figure(sheet):

    # Reads in the data to a Pandas DataFrame
    gc = pd.read_excel("PlasticCrudeAnalysis.xlsx", sheet_name=sheet)

    # Drops the first two rows to get rid of units and "TIC" column
    units = gc.iloc[0]
    gc.drop(index=[0, 1], inplace=True)
    gc.set_index("Peak", inplace=True);

    # Creates a column of the mole fractions of the sample (aka "Sample Fraction")
    gc['Sample Fraction'] = gc['Area '] / gc['Area '].sum()

    # Loops through the chemSplit function to get the number of Carbons, Hydrogens,
    # and Oxygens from the chemical formula
    for peak in gc.index:
        quants = chemSplit(gc["Chemical Formula"][peak])
        gc.loc[(peak, 'C Number')] = int(quants[0])
        gc.loc[(peak, 'H Number')] = int(quants[1])
        gc.loc[(peak, 'O Number')] = int(quants[2])

    # Creates Plotly elements
    trace1 = go.Bar(x=processing(gc[gc['O Number']==0])[0],
            y=processing(gc[gc['O Number']==0])[1], name='Hydrocarbons', marker={'color':'orange'})
    trace2 = go.Bar(x=processing(gc[gc['O Number']!=0])[0],
            y=processing(gc[gc['O Number']!=0])[1], name='Non-Hydrocarbons', marker={'color':'black'})
    data = [trace1, trace2] # Creates dataset
    layout = go.Layout(title='Carbon Distribution', yaxis={'title':'Mole Fraction'},
            xaxis={'title':'Carbon Number'}, barmode='stack') # Creates layout object
    fig = go.Figure(data=data, layout=layout)
    return {'data':data, 'layout':layout}

# Runs the Dash server
if __name__ == '__main__':
    app.run_server()
