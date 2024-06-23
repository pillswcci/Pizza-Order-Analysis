
# import housing data
import pandas as pd
home_data = pd.read_csv("housing.csv")

# print data to check
print(home_data.head())

# create scatterplot
import seaborn as sns
sns.scatterplot(data = home_data,
                x = 'longitude',
                y = 'latitude',
                hue = 'median_house_value')

# display plot
import matplotlib.pyplot as plt
plt.show() 
