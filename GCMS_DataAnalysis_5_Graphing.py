'''
Analysis Tool for processing data from the Thermo-Scientific
GC-MS in Johnson Hall at Oregon State University

Written by: Tyler Kuykendall
Date: April 18, 2019
'''

# Imports Packages to use
import string # Imports String library
import re # Imports Regular Expressions
import numpy as np # Imports Numpy
import pandas as pd # Imports Pandas ()
import seaborn as sns # Imports Seaborn (Graphing/Visualization tool)
import matplotlib.pyplot as plt # Needed to change plot elements

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

def gcAnalysis(df):
    '''
    This function will take in the prepared DataFrame that was constructed
    from the GC data and turn it into the plots that we want.
    '''

    # Groups up the alkanes by carbon number
    # Casts the C, H, and O numbers to be integers
    df = df.astype({'C Number': int, 'H Number': int, 'O Number': int})

    # Creates a GroupBy object so that we can operate around carbon numbers
    cNumGroups = df.groupby('C Number')

    # Creates a list of C Numbers and the sum of their sample fractions
    cAreas = cNumGroups['Sample Fraction'].sum()

    # Creates a Seaborn barplot object with C Number on the x-axis and
    # mole fraction on the y-axis
    plt.figure(figsize=(12,8)) # Scales the figure to an appropriate size
    sns.set_context('notebook', font_scale=1.25) # Enhances readibility
    sns.set_style({"ytick.direction": 'in', "xtick.bottom": False}) # Moves ticks
    plt.xlabel('Carbon Number') # Labels x-axis
    plt.ylabel('Mole Fraction') # Labels y-axis
    plt.title(input('Title: '))
    cDist = sns.barplot(cAreas.index, cAreas.values, palette='viridis') # Creates plot

    figName = input('\nFile Name: ') + '.jpg'
    cDist.figure.savefig(figName)

    plt.show()

# Reads in the data to a Pandas DataFrame
sheet = int(input("\nSelect a sheet to analyze (0, 1, 2, ...): "))
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

# Menu loop for data processing
while True:

    # Asks the user what data they would like to process
    resp = input('''
0: Process as sampled
1: Process only hydrocarbons
2: Process oxygenated compounds only
3: Process all scenarios

Please select an option: ''')

    # Processes data based on selection
    if resp == '0':
        gcAnalysis(gc)
        break
    elif resp == '1':
        gcAnalysis(gc[gc['O Number'] == 0])
        break
    elif resp == '2':
        gcAnalysis(gc[gc['O Number'] != 0])
    elif resp == '3':
        gcAnalysis(gc)
        gcAnalysis(gc[gc['O Number'] == 0])
        gcAnalysis(gc[gc['O Number'] != 0])
        break
    else:
        print("Selection invalid")
