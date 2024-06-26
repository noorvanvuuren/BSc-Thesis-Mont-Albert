import matplotlib.pyplot as plt 
import pandas as pd 
from matplotlib.markers import MarkerStyle
import numpy as np
from matplotlib.font_manager import FontProperties
import os

font_prop = FontProperties(family='Verdana', size=30)

# Read Excel file
file_path = r"C:\Users\noorv\OneDrive\Picture's\AW - jaar 3\BSc Thesis\Edited Excel files\BSc Thesis All Measurements.xlsx"
file = pd.read_excel(file_path, 'Grains')

# Correct column names and indices
grainsize = file.iloc[0, 4:8]
stipp = file.iloc[4, 4:8]
cross = file.iloc[10, 4:8]
crosssliding = file.iloc[7, 4:8]
grainsize_auto = file.iloc[13, 4:8]
stipp_auto = file.iloc[17, 4:8]
cross_auto = file.iloc[23, 4:8]
crosssliding_auto = file.iloc[20, 4:8]

# Define values for labeling
kma2202_grainsize = file.iloc[0, 4]
kma2202_stipp = file.iloc[4, 4]
kma2202_crosssliding = file.iloc[7, 4]
kma2202_crossmicro = file.iloc[10, 4]

kma2277_grainsize = file.iloc[0, 5]
kma2277_stipp = file.iloc[4, 5]
kma2274_grainsize = file.iloc[0, 6]
kma2274_stipp = file.iloc[4, 6]
kma2263_grainsize = file.iloc[0, 7]
kma2263_stipp = file.iloc[4, 7]

akma2202_grainsize = file.iloc[13, 4]
akma2202_stipp = file.iloc[17, 4]

# Calculate error bars
grainsize_min = file.iloc[1, 4:8]
grainsize_max = file.iloc[1, 4:8]
grainsize_errors = [grainsize_min, grainsize_max]

grainsize_min_auto = file.iloc[14, 4:8]
grainsize_max_auto = file.iloc[14, 4:8]
grainsize_errors_auto = [grainsize_min_auto, grainsize_max_auto]

stipp_max = file.iloc[5, 4:8]
stipp_min = file.iloc[6, 4:8]
stipp_errors_lower = stipp - stipp_min
stipp_errors_upper = stipp_max - stipp
stipp_errors = [stipp_errors_lower, stipp_errors_upper]

cross_max = file.iloc[11, 4:8]
cross_min = file.iloc[12, 4:8]
cross_errors_lower = cross - cross_min
cross_errors_upper = cross_max - cross
cross_errors = [cross_errors_lower, cross_errors_upper]

crosssliding_min = file.iloc[9, 4:8]
crosssliding_max = file.iloc[8, 4:8]
crosssliding_errors_lower = crosssliding - crosssliding_min
crosssliding_errors_upper = crosssliding_max - crosssliding
crosssliding_errors = [crosssliding_errors_lower, crosssliding_errors_upper]

stipp_auto_max = file.iloc[18, 4:8]
stipp_auto_min = file.iloc[19, 4:8]
stipp_auto_errors_lower = stipp_auto - stipp_auto_min
stipp_auto_errors_upper = stipp_auto_max - stipp_auto
stipp_auto_errors = [stipp_auto_errors_lower, stipp_auto_errors_upper]

cross_auto_max = file.iloc[24, 4:8]
cross_auto_min = file.iloc[25, 4:8]
cross_auto_errors_lower = cross_auto - cross_auto_min
cross_auto_errors_upper = cross_auto_max - cross_auto
cross_auto_errors = [cross_auto_errors_lower, cross_auto_errors_upper]

crosssliding_auto_min = file.iloc[22, 4:8]
crosssliding_auto_max = file.iloc[21, 4:8]
crosssliding_auto_errors_lower = crosssliding_auto - crosssliding_auto_min
crosssliding_auto_errors_upper = crosssliding_auto_max - crosssliding_auto
crosssliding_auto_errors = [crosssliding_auto_errors_lower, crosssliding_auto_errors_upper]

# Define colors for plotting 
colors = ['#dd447c', '#edee00', '#0AFFFF', '#008080']

# Plot figure
plt.figure(figsize=(24, 16))

# Create correct labels
plt.scatter(kma2202_grainsize, kma2202_stipp, label='Line Intercept Method', s=400, marker=MarkerStyle('o', 'full'), color='gray')
plt.scatter(akma2202_grainsize, akma2202_stipp, label='GOS Method', s=400, marker=MarkerStyle('o', 'none'), linewidth=3, color='gray')

plt.scatter(kma2202_grainsize, kma2202_stipp, label='KMA22-02', s=400, marker=MarkerStyle('o', 'full'), color='#dd447c')
plt.scatter(kma2277_grainsize, kma2277_stipp, label='KMA22-77', s=400, marker=MarkerStyle('o', 'full'), color='#edee00')
plt.scatter(kma2274_grainsize, kma2274_stipp, label='KMA22-74', s=400, marker=MarkerStyle('o', 'full'), color='#0AFFFF')
plt.scatter(kma2263_grainsize, kma2263_stipp, label='KMA22-63', s=400, marker=MarkerStyle('o', 'full'), color='#008080')

plt.scatter(kma2202_grainsize, kma2202_stipp, label='Stipp et al.', s=400, marker=MarkerStyle('o', 'full'), c='black')
plt.scatter(kma2202_grainsize, kma2202_crossmicro, label='Cross et al. - 1 \u03bcm resolution ', s=400, marker=MarkerStyle('^', 'full'), c='black')
plt.scatter(kma2202_grainsize, kma2202_crosssliding, label='Cross et al. - sliding resolution', s=400, marker=MarkerStyle('D', 'full'), c='black')

# Plot values
plt.scatter(grainsize, stipp, s=400, marker=MarkerStyle('o', 'full'), color=colors)
plt.scatter(grainsize, cross, s=400, marker=MarkerStyle('^', 'full'), color=colors)
plt.scatter(grainsize, crosssliding, s=400, marker=MarkerStyle('D', 'full'), color=colors)
plt.scatter(grainsize_auto, stipp_auto, s=400, marker=MarkerStyle('o', 'full'), facecolors='white', linewidth=3, color=colors)
plt.scatter(grainsize_auto, cross_auto, s=400, marker=MarkerStyle('^', 'full'), facecolors='white', linewidth=3, color=colors)
plt.scatter(grainsize_auto, crosssliding_auto, s=400, marker=MarkerStyle('D', 'full'), facecolors='white', linewidth=3, color=colors)

# Plot errobars
plt.errorbar(grainsize, stipp, xerr=grainsize_errors, fmt='none', ecolor ='silver', capsize=10, ls='none',linewidth=2, zorder=-1)
plt.errorbar(grainsize_auto, stipp_auto, xerr=grainsize_errors_auto, fmt='none', elinewidth=2, ecolor ='silver', capsize=10,zorder=-1)
plt.errorbar(grainsize, stipp, yerr=stipp_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(grainsize, cross, yerr=cross_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(grainsize, crosssliding, yerr=crosssliding_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(grainsize_auto, stipp_auto, yerr=stipp_auto_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(grainsize_auto, cross_auto, yerr=cross_auto_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=10, ls='none',zorder=-1)
plt.errorbar(grainsize_auto, crosssliding_auto, yerr=crosssliding_auto_errors, fmt='none', elinewidth=2, ecolor ='silver', capsize=5, ls='none',zorder=-1)

plt.xlabel('Grain size (\u03bcm)', size=40, fontproperties=font_prop)
plt.ylabel('Stress (MPa)', size=40, fontproperties=font_prop)
plt.grid(color='gainsboro', which="both")

xticks_values = np.arange(0, 700, 50)  
yticks_values = [0, 10, 100]

plt.xticks(xticks_values, size=35, fontproperties=font_prop, c='dimgray')
plt.yticks(yticks_values, size=35, fontproperties=font_prop, c='dimgray')
plt.yscale('log')
plt.legend(prop=font_prop, fontsize=20)

directory =   # insert file directory
filename = 'grainsizevstress.png'
full_path = os.path.join(directory, filename)

plt.savefig(full_path, bbox_inches='tight', dpi=300)

plt.show()





