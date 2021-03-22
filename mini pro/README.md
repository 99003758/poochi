
import pandas as pd

from openpyxl import load_workbook

import sys
import matplotlib.pyplot as plt

# taken an data frame for accesing output
df2=pd.DataFrame()

for a in range(int(input(" Enter the no of inputs you required:"))):
# for n number

    first = int(input("Enter the ps.no:"))
# first input
    second = (input("Enter the name:"))
# second input
    third = input("Enter the email:")
# third input

    class1= pd.read_excel('final sheet211.xlsx',sheet_name=['Sheet1', 'Sheet2', 'Sheet3', 'Sheet4', 'Sheet5'])
# reading all sheets in excel and reading it to a variable

    g = class1['Sheet1']
# sheet selectionand giving related data frame
    g = g[g['p.s.no'] == first]
# ps no validationand giving related data frame
    g = g[g['NAME'] == second]
# name validation and giving related data frame
    g = g[g['email'] == third]
# mail validation and giving related data frame

    df = pd.DataFrame(g)
# taking an another data frame
    if df.empty:
        sys.exit("ENTERED DETAILS ARE INCORRECT")

# if input is invalid it will show as above


    for i in class1.keys():
# iteration of sheets from 2 to 5
        sheet = class1[i]
        g = sheet[sheet['p.s.no'] == first]
# checking all inputs
        g = g[g['NAME'] == second]
        g = g[g['email'] == third]
        col = sheet.columns
# getting all columns into an output master sheet
        for j in col:
            #iteration of data in columns and sending to data frame
            df[j] = g[j]
#  check the input data entries is 1 or more and append data in to data frame
            if a== 0:
                df2[j]=g[j]
# if data entries are more than 1 concatination of data frames in to an final master sheet
    if a!=0:
        df2 = pd.concat([df, df2])
        print(df2)


# iterates through all sheets

k = load_workbook('final sheet.xlsx')
df2.to_excel('final sheet.xlsx', sheet_name='Sheet6')
print(df2)

# selecting inputs for ploting of graphs for selected data
df2.plot.bar(x='NAME',y='laptops')

plt.show()
#it wil show the plot i
k.save('.xlsx')
k.close()






