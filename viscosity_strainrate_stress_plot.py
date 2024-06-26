from matplotlib import pyplot as plt
import matplotlib.patches as patches
import numpy as np
from matplotlib.font_manager import FontProperties
import os

# Properties for plot
font_prop = FontProperties(family='Verdana')
colors = ['#dd447c', '#edee00', '#0AFFFF', '#008080', '#05c491']

# Define stress and viscosity values to use
stress = np.linspace(0, 4.0e7, 500)  # Pa
viscosities = [1e15, 1e16, 1e17, 1e18, 1e19]  # Pa·s

# Plot figure
plt.figure(figsize=(24, 16))

# Make strain rate calculation for each value of viscosity
for viscosity, color in zip(viscosities, colors):
    strain_rate = stress / (2 * viscosity)
    plt.plot(stress, strain_rate, label=f'Viscosity = {viscosity:.0e} Pa·s', linewidth=3, color=color)

# Insert rectangles to show range if necessary
# rect1 = patches.Rectangle((1.1E7, 3.1E-11), 2.4E7, 1.8E-8, linewidth=6, edgecolor='dimgray', facecolor='none', zorder=5)
# plt.gca().add_patch(rect1)

# rect2 = patches.Rectangle((7.0E6, 6.6E-14), 6.8E6, 1.1E-11, linewidth=6, linestyle='--', edgecolor='dimgray', facecolor='none', zorder=5)
# plt.gca().add_patch(rect2)

plt.xlabel('Stress (MPa)', size=40, fontproperties=font_prop)
plt.ylabel('Strain Rate (s$^{-1}$)', size=40, fontproperties=font_prop)
# plt.legend(loc='lower right')
# plt.rc('legend', fontsize=40)
plt.grid(True, which='both', color='gainsboro')
plt.xticks(fontsize=35, c='dimgray')
plt.yticks(fontsize=35, c='dimgray')
plt.ylim(1e-15, 1e-7)
plt.xlim(0, 4.0e7)
plt.yscale('log')

directory = # insert file directory
filename = 'viscosity.png'
full_path = os.path.join(directory, filename)

plt.savefig(full_path, bbox_inches='tight', dpi=300)

plt.show()





