import matplotlib.pyplot as plt 
import pandas as pd 
from matplotlib.markers import MarkerStyle
from matplotlib.font_manager import FontProperties
import numpy as np
import os

font_prop = FontProperties(family='Verdana', size=30)

# Read Excel file
file_path = "BSc Thesis All Measurements.xlsx"
file = pd.read_excel(file_path, 'Subgrains')

# Correct column names and indices
subgrainsize2 = file.iloc[4, 4:8]
goddardcor2 = file.iloc[14, 4:8]
goddarduncor2 = file.iloc[17, 4:8]
subgrainsize_manual = file.iloc[20, 4:8]
goddardcor_manual = file.iloc[24, 4:8]
goddarduncor_manual = file.iloc[27, 4:8]

# Define values for labeling
kma2202_grainsize = file.iloc[4,4]
kma2202_goddardcor = file.iloc[14,4]
kma2202_goddarduncor= file.iloc[17,4]

kma2277_grainsize = file.iloc[4,5]
kma2277_goddardcor = file.iloc[14,5]
kma2277_goddarduncor = file.iloc[17,5]

kma2274_grainsize = file.iloc[4,6]
kma2274_goddardcor = file.iloc[14,6]

kma2263_grainsize = file.iloc[4,7]
kma2263_goddardcor = file.iloc[14,7]

# Calculate error bars
subgrainsize2_min = file.iloc[5, 4:8]
subgrainsize2_max = file.iloc[5, 4:8]
subgrainsize2_errors = [subgrainsize2_min, subgrainsize2_max]

subgrainsize_man_min = file.iloc[21, 4:8]
subgrainsize_man_max = file.iloc[21, 4:8]
subgrainsize_man_errors = [subgrainsize_man_min, subgrainsize_man_max]

goddardcor2_max = file.iloc[15, 4:8]
goddardcor2_min = file.iloc[16, 4:8]
goddardcor2_errors_lower = goddardcor2 - goddardcor2_min
goddardcor2_errors_upper = goddardcor2_max - goddardcor2
goddardcor2_errors = [goddardcor2_errors_lower, goddardcor2_errors_upper]

goddarduncor2_max = file.iloc[18, 4:8]
goddarduncor2_min = file.iloc[19, 4:8]
goddarduncor2_errors_lower = goddarduncor2 - goddarduncor2_min
goddarduncor2_errors_upper = goddarduncor2_max - goddarduncor2
goddarduncor2_errors = [goddarduncor2_errors_lower, goddarduncor2_errors_upper]

goddardcor_manual_max = file.iloc[25, 4:8]
goddardcor_manual_min = file.iloc[26, 4:8]
goddardcor_manual_errors_lower = goddardcor_manual - goddardcor_manual_min
goddardcor_manual_errors_upper = goddardcor_manual_max - goddardcor_manual
goddardcor_manual_errors = [goddardcor_manual_errors_lower, goddardcor_manual_errors_upper]

goddarduncor_manual_max = file.iloc[28, 4:8]
goddarduncor_manual_min = file.iloc[29, 4:8]
goddarduncor_manual_errors_lower = goddarduncor_manual - goddarduncor_manual_min
goddarduncor_manual_errors_upper = goddarduncor_manual_max - goddarduncor_manual
goddarduncor_manual_errors = [goddarduncor_manual_errors_lower, goddarduncor_manual_errors_upper]

# Define colors for plotting
colors = ['#dd447c', '#008080', '#edee00',  '#70feff']

# Plot figure
plt.figure(figsize=(24, 16))

# Create correct labels
plt.scatter(kma2202_grainsize, kma2202_goddardcor, label='Line Intercept Method - 2\u00b0', s=400, marker=MarkerStyle('o', 'full'), c='grey')
plt.scatter(kma2202_grainsize, kma2202_goddardcor, label='KMA22-02', s=400, marker=MarkerStyle('o', 'full'), c='#dd447c')
plt.scatter(kma2277_grainsize, kma2277_goddardcor, label='KMA22-77', s=400, marker=MarkerStyle('o', 'full'), c='#edee00')
plt.scatter(kma2274_grainsize, kma2274_goddardcor, label='KMA22-74', s=400, marker=MarkerStyle('o', 'full'), c='#008080')
plt.scatter(kma2263_grainsize, kma2263_goddardcor, label='KMA22-63', s=400, marker=MarkerStyle('o', 'full'), c='#70feff')
plt.scatter(kma2277_grainsize, kma2277_goddardcor, label='Goddard - corrected', s=400, marker=MarkerStyle('o', 'full'), c='black')
plt.scatter(kma2202_grainsize, kma2202_goddarduncor, label='Goddard - uncorrected', s=400, marker=MarkerStyle('^', 'full'), c='black')

# Plot values
for i in range(len(subgrainsize2)):
    plt.scatter(subgrainsize2[i], goddardcor2[i], s=400, marker=MarkerStyle('o', 'full'), linewidth=1, color=colors[i])
    plt.scatter(subgrainsize2[i], goddarduncor2[i], s=400, marker=MarkerStyle('^', 'full'), linewidth=1, color=colors[i])
   
# Plot errobars
plt.errorbar(subgrainsize2, goddarduncor2, xerr=subgrainsize2_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(subgrainsize2, goddardcor2, yerr=goddardcor2_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(subgrainsize2, goddarduncor2, yerr=goddarduncor2_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)

plt.xlabel('Subgrain size (\u03bcm)', size=40, fontproperties=font_prop)
plt.ylabel('Stress (MPa)', size=40, fontproperties=font_prop)
plt.grid(color='gainsboro', which="both")

xticks_values = np.arange(0, 450, 50)  
yticks_values = np.arange(0, 150, 10)  

plt.xticks(size=35, fontproperties=font_prop, c='dimgray')
plt.yticks(size=35, fontproperties=font_prop, c='dimgray')
plt.xticks(xticks_values, size=35, fontproperties=font_prop, c='dimgray')
plt.yticks(size=35, fontproperties=font_prop, c='dimgray')
plt.legend(fontsize=20, prop=font_prop)
plt.yscale('log')

directory =  # insert file directory
filename = 'subgrainsize2vstress.png'
full_path = os.path.join(directory, filename)

plt.savefig(full_path, bbox_inches='tight', dpi=300)

plt.show()