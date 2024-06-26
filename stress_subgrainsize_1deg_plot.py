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
subgrainsize1 = file.iloc[0, 4:8]
goddardcor1 = file.iloc[8, 4:8]
goddarduncor1 = file.iloc[11, 4:8]
subgrainsize_manual = file.iloc[20, 4:8]
goddardcor_manual = file.iloc[24, 4:8]
goddarduncor_manual = file.iloc[27, 4:8]

# Define values for labeling
kma2202_grainsize = file.iloc[0, 4]
kma2202_goddardcor = file.iloc[8, 4]
kma2202_goddarduncor = file.iloc[11, 4]

kma2277_grainsize = file.iloc[0, 5]
kma2277_goddardcor  = file.iloc[9, 4]

kma2274_grainsize = file.iloc[0, 6]
kma2274_goddardcor = 3.16

kma2263_grainsize = file.iloc[0, 7]
kma2263_goddardcor = 5.71

man_kma2202_grainsize = file.iloc[20, 4]
man_kma2202_goddardcor = file.iloc[24, 4]

# Calculate error bars
subgrainsize1_min = file.iloc[1, 4:8]
subgrainsize1_max = file.iloc[1, 4:8]
subgrainsize1_errors = [subgrainsize1_min, subgrainsize1_max]

subgrainsize2_min = file.iloc[5, 4:8]
subgrainsize2_max = file.iloc[5, 4:8]
subgrainsize2_errors = [subgrainsize2_min, subgrainsize2_max]

subgrainsize_man_min = file.iloc[21, 4:8]
subgrainsize_man_max = file.iloc[21, 4:8]
subgrainsize_man_errors = [subgrainsize_man_min, subgrainsize_man_max]

goddardcor1_max = file.iloc[9, 4:8]
goddardcor1_min = file.iloc[10, 4:8]
goddardcor1_errors_lower = goddardcor1 - goddardcor1_min
goddardcor1_errors_upper = goddardcor1_max - goddardcor1
goddardcor1_errors = [goddardcor1_errors_lower, goddardcor1_errors_upper]

goddarduncor1_max = file.iloc[12, 4:8]
goddarduncor1_min = file.iloc[13, 4:8]
goddarduncor1_errors_lower = goddarduncor1 - goddarduncor1_min
goddarduncor1_errors_upper = goddarduncor1_max - goddarduncor1
goddarduncor1_errors = [goddarduncor1_errors_lower, goddarduncor1_errors_upper]

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
colors = ['#dd447c', '#edee00', '#0AFFFF', '#008080']

# Plot figure
plt.figure(figsize=(24, 16))

# Create correct labels
plt.scatter(kma2202_grainsize, kma2202_goddardcor, label='Line Intercept Method - 1\u00b0', s=400, marker=MarkerStyle('o', 'full'), c='grey')
plt.scatter(man_kma2202_grainsize, man_kma2202_goddardcor, label='Manual Subgrain Method - 1\u00b0', s=400, marker=MarkerStyle('o', 'none'), c='gray')
plt.scatter(kma2202_grainsize, kma2202_goddardcor, label='KMA22-02', s=400, marker=MarkerStyle('o', 'full'), c='#dd447c')
plt.scatter(kma2277_grainsize, kma2277_goddardcor, label='KMA22-77', s=400, marker=MarkerStyle('o', 'full'), c='#edee00')
plt.scatter(kma2274_grainsize, kma2274_goddardcor, label='KMA22-74', s=400, marker=MarkerStyle('o', 'full'), c='#0AFFFF')
plt.scatter(kma2263_grainsize, kma2263_goddardcor, label='KMA22-63', s=400, marker=MarkerStyle('o', 'full'), c='#008080')
plt.scatter(kma2202_grainsize, kma2202_goddardcor, label='Goddard - corrected', s=400, marker=MarkerStyle('o', 'full'), c='black')
plt.scatter(kma2202_grainsize, kma2202_goddarduncor, label='Goddard - uncorrected', s=400, marker=MarkerStyle('^', 'full'), c='black')

# Plot values
for i in range(len(subgrainsize1)):
    plt.scatter(subgrainsize1[i], goddardcor1[i], s=400, marker=MarkerStyle('o', 'full'), color=colors[i])
    plt.scatter(subgrainsize1[i], goddarduncor1[i], s=400, marker=MarkerStyle('^', 'full'), color=colors[i])
    plt.scatter(subgrainsize_manual[i], goddardcor_manual[i], s=400, marker=MarkerStyle('o', 'full'), facecolors='white', linewidth=3, color=colors[i])
    plt.scatter(subgrainsize_manual[i], goddarduncor_manual[i], s=400, marker=MarkerStyle('^', 'full'), facecolors='white', linewidth=3, color=colors[i])

# Plot errobars
plt.errorbar(subgrainsize1, goddarduncor1, xerr=subgrainsize1_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(subgrainsize_manual, goddarduncor_manual, xerr=subgrainsize_man_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(subgrainsize1, goddardcor1, yerr=goddardcor1_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(subgrainsize1, goddarduncor1, yerr=goddarduncor1_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(subgrainsize_manual, goddardcor_manual, yerr=goddardcor_manual_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(subgrainsize_manual, goddarduncor_manual, yerr=goddarduncor_manual_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)

plt.xlabel('Subgrain size (\u03bcm)', size=40, fontproperties=font_prop)
plt.ylabel('Stress (MPa)', size=40, fontproperties=font_prop)
plt.grid(color='gainsboro', which="both")

xticks_values = np.arange(0, 300, 50)  
yticks_values = np.arange(0, 150, 10)  

plt.xticks(size=35, fontproperties=font_prop)
plt.yticks(size=35, fontproperties=font_prop)
plt.xticks(xticks_values, size=35, fontproperties=font_prop, c='dimgray')
plt.yticks(yticks_values, size=35, fontproperties=font_prop,  c='dimgray')
plt.legend(fontsize=20, prop=font_prop)
plt.yscale('log')

directory = # insert file directory
filename = 'subgrainsize1vstress.png'
full_path = os.path.join(directory, filename)

plt.savefig(full_path, bbox_inches='tight', dpi=300)

plt.show()