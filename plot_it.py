import pandas as pd
import xlrd # to make pandas work
import matplotlib.pyplot as plt


import glob
files = glob.glob('./results/*.xlsx')

for file in files:
    xl = pd.ExcelFile(file)
    sheet_name = xl.sheet_names[0]
    sheet = xl.parse(sheet_name)
    plot2 = sheet['State'].value_counts().plot.bar()
    result = sheet['State'].value_counts()
    #did = plt.plot(result)
    plt.xticks(rotation='vertical')
    plt.xlabel('State')
    plt.ylabel('Count')
    plt.tight_layout()
    #result.to_excel('./results/sums.xlsx', sheet_name='ResultSums2')
    plt.show()

you = input('Press enter to quit')