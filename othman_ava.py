import pandas as pd
import matplotlib.pyplot as plt
path_fl = r'D:\2023\November\Calories.xlsx'
df = pd.read_excel(path_fl)
df.head(5)
df.plot(kind = 'line', x='Duration', y='Calories')
df.plot.bar(stacked= True)
df.plot.barh()
plt.show()
