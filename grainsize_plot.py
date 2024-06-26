import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from matplotlib.font_manager import FontProperties
import os

font_prop = FontProperties(family='Verdana', size=30)

plt.rcParams['hatch.linewidth'] = 4

# Read Excel file
file_path = "BSc Thesis All Measurements.xlsx"
grain = pd.read_excel(file_path, 'Grains')
subgrain = pd.read_excel(file_path, 'Subgrains')

# Set width of bar
barWidth = 0.15
fig = plt.subplots(figsize =(24, 16))
colors = ['orange', 'deeppink', 'teal', 'cyan']

# Set height of bar
grainsize_li = grain.iloc[0, 4:8]
subgrainsize1_li = subgrain.iloc[0, 4:8]
subgrainsize2_li = subgrain.iloc[4, 4:8]
subgrainsize1_man = subgrain.iloc[20, 4:8]
grainsize_auto = grain.iloc[13, 4:8]

# Set position of bar on Y axis
br1 = np.arange(len(grainsize_li))
br2 = [x + barWidth for x in br1]
br3 = [x + 0.05 + barWidth for x in br2]
br4 = [x + barWidth for x in br3]
br5 = [x + barWidth for x in br4]

br1 = br1[::-1]
br2 = br2[::-1]
br3 = br3[::-1]
br4 = br4[::-1]
br5 = br5[::-1]

# Make the plot
plt.barh(br5, grainsize_li, color ='#008080', height = barWidth, 
        edgecolor ='white', label ='Line Intercept Method')
plt.barh(br4, grainsize_auto, color ='#0AFFFF', height = barWidth, 
        edgecolor ='white', label ='GOS Method')
plt.barh(br3, subgrainsize1_li, color ='#05c491', height = barWidth, 
        edgecolor ='whitesmoke', label ='Line Intercept Method 1\u00b0', hatch='|')
plt.barh(br2, subgrainsize2_li, color ='#edee00', height = barWidth, 
        edgecolor ='whitesmoke', label ='Line Intercept Method 2\u00b0', hatch='|')
plt.barh(br1, subgrainsize1_man, color ='#dd447c', height = barWidth, 
        edgecolor ='whitesmoke', label ='Manual Subgrain Method', hatch='|')

plt.ylabel('', size=40, fontproperties=font_prop)
plt.xlabel('(Sub)grain size (\u03bcm)', size=40, fontproperties=font_prop)
plt.yticks([r + 0.15 + barWidth for r in range(len(grainsize_li))], 
        ['KMA22-63 \n (230 m)', 'KMA22-74 \n (200 m)', 'KMA22-77 \n (-120 m)', 'KMA22-02 \n (-15 m)'], size=35)
plt.xticks(size=35, c='dimgray')
plt.legend(fontsize=20, prop=font_prop, loc='upper right')
plt.grid(axis="x", zorder=-2)
plt.rc('axes', axisbelow=True)

directory = # Insert file directory
filename = 'barchart2.png'
full_path = os.path.join(directory, filename)

plt.savefig(full_path, bbox_inches='tight', dpi=300)

plt.show()
