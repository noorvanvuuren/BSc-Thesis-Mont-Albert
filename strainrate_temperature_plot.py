from matplotlib import pyplot as plt
import matplotlib.patches as patches
import numpy as np
from scipy.interpolate import interp1d
from matplotlib.font_manager import FontProperties
import os
import pandas as pd

font_prop = FontProperties(family='Verdana')

# Flow laws for subducted materials
# Change WATER FUGACITY string for geothermal gradient of interest; strain rate according to tectonic setting and shear zone thickness

R = 8.3144
T = 273 + np.arange(300, 1001, 50)        # Temperature [K]
TC = T-273                 # Temperature [C]
geotherm = 25              # geothermal gradient [degrees C/km]
rho = 2700                 # average crustal density [kg/m^3]
h = TC/geotherm*1e3        # depth [m]
g = 9.8                    # acceleration of gravity [m^2/s]
P = rho*g*h/1e9            # calculated pressure at each T for geotherm above

# # using UMN water fugacity calculator: fH2O from 300-1000C at 50C intervals
# # for a geothermal gradient of....

# % FH2O = 1e3*[0.203 0.469 0.944 1.758 2.923 4.545 6.692 9.627...
#           %  13.045 17.134 22.326 27.902 34.204 41.233 49.743]; %10c/km

# %FH2O = 1e3*[0.071 0.156 0.291 0.510 0.801 1.219 1.718 2.383...
#  %           3.120 4.060 5.047 6.274 7.513 9.026 10.509]; %15c/km

# %FH2O = 1e3*[0.042 0.086 0.156 0.259 0.399 0.599 0.833 1.115...
# %1.445 1.824 2.252 2.728 3.317 3.893 4.514]; %20c/km

FH2O = 1e3 * np.array([0.029, 0.059, 0.107, 0.170, 0.260, 0.377, 0.522, 0.678,
                      0.877, 1.106, 1.331, 1.612, 1.920, 2.254, 2.562])  # 25c/km


# %FH2O = 1e3*[0.024 0.045 0.080 0.129 0.195 0.270 0.369 0.487...
# % 0.623 0.756 0.924 1.109 1.310 1.493 1.720]; %30c/km

# Water fugacity interpolation
v = FH2O
tt = round(len(FH2O) / len(T), 1)

# Generate xq
xq = np.arange(1, len(FH2O) + 1, tt)

# Interpolate fH2O
interp_func = interp1d(np.arange(1, len(v) + 1), v, kind='linear')
fH2O = interp_func(xq)

# E = 1e-10 #strain rate /s
n = 4         # stress exponent
A = 10**-11.2  # pre-exponential constant
Q = 1.35e5    # activation energy
d = 10  # in microns
m = 0
fn = 1  # so that fH2O for everything but quartz goes to 1

# Define stress for each strain rate
s_grains_1 = 35.09127635
s_grains_2 = 13.82835386
s_subgrains_1 = 10.80230023
s_subgrains_2 = 6.984765561

# Write formulas for strain rates using correct stresses
E1 = (s_grains_1**n/d**m) * (A * fH2O**fn * np.exp(-Q / (R * T)))
E2 = (s_grains_2**n/d**m) * (A * fH2O**fn * np.exp(-Q / (R * T)))
E3 = (s_subgrains_1**n/d**m) * (A * fH2O**fn * np.exp(-Q / (R * T)))
E4 = (s_subgrains_2**n/d**m) * (A * fH2O**fn * np.exp(-Q / (R * T)))

# Interpolate the funciton and calculate values at certain T
interp_func_E1 = interp1d(T-273, E1, kind='linear')
E1_at_850 = interp_func_E1(850)
E1_at_900 = interp_func_E1(900)

interp_func_E2 = interp1d(T-273, E2, kind='linear')
E2_at_650 = interp_func_E2(650)
E2_at_700 = interp_func_E2(700)

interp_func_E3 = interp1d(T-273, E3, kind='linear')
E3_at_850 = interp_func_E3(850)
E3_at_800 = interp_func_E3(800)

interp_func_E4 = interp1d(T-273, E4, kind='linear')
E4_at_650 = interp_func_E4(650)
E4_at_600 = interp_func_E4(600)

# Plot figure
plt.figure(figsize=(24, 16))

# Plot strain rates
plt.semilogy(T-273, E1, '#dd447c', linewidth=3, label='Grains [< 120 m]')
plt.semilogy(T-273, E2, '#edee00', linewidth=3, linestyle='--', dashes=(10, 2), label='Grains [> 200 m]')
plt.semilogy(T-273, E3, '#0AFFFF', linewidth=3, label='Subgrains [< 120 m]')
plt.semilogy(T-273, E4, '#008080', linewidth=3, linestyle='--', dashes=(10, 2), label='Subgrains [> 200 m]')

# Plot vertical lines 
plt.plot([650, 650], [10**-15, E2_at_650],
         color='dimgray', linewidth=4, linestyle=':')
plt.plot([850, 850], [10**-15, E1_at_850],
         color='dimgray', linewidth=4, linestyle=':')

# Plot specific points of interest
plt.scatter(850, 8.108060377230426e-09, s=1200, color='thistle',
            edgecolor='dimgray', linewidth=4, zorder=4, marker='*')
plt.scatter(650, 3.584680891828387e-12, s=300, color='dimgray', zorder=4)
plt.scatter(850, 7.280903734389602e-11, s=300, color='dimgray', zorder=4)
plt.scatter(650, 2.333328571129852e-13, s=1200, color='thistle',
            edgecolor='dimgray', linewidth=4, zorder=4, marker='*')

# Plot rectangles to show range if needed
# rect1 = patches.Rectangle((800, E3_at_800), 100, (E1_at_900 - E3_at_800), linewidth=6, edgecolor='dimgray', facecolor='none', zorder=5)
# plt.gca().add_patch(rect1)

# rect2 = patches.Rectangle((600, E4_at_600), 100, (E2_at_700 - E4_at_600), linewidth=6, edgecolor='dimgray', linestyle='--', facecolor='none', zorder=5)
# plt.gca().add_patch(rect2)

plt.ylim([10**-15, 10**-8])
plt.xlim([500, 900])

xticks_values = np.arange(500, 1100, 100)

plt.legend(loc='upper left')
plt.rc('legend', fontsize=40)
plt.grid(color='gainsboro', which="minor")
plt.yscale('log')
plt.xscale('log')
plt.ylabel('Strain rate (s$^{-1}$)', size=40, fontproperties=font_prop)
plt.xlabel('Temperature (°C)', size=40, fontproperties=font_prop)
plt.xticks(xticks_values, fontsize=35, c='dimgray')
plt.yticks(fontsize=35, c='dimgray')
plt.axis([5e2, 10e2, 10e-16, 10e-8])

directory =  # insert directory
filename = 'strainrate.png'
full_path = os.path.join(directory, filename)

plt.savefig(full_path, bbox_inches='tight', dpi=300)

# Print values
print("E1_at_850:", E1_at_850)
print("E1_at_900:", E1_at_900)
print("E2_at_650:", E2_at_650)
print("E2_at_700:", E2_at_700)
print("E3_at_800:", E3_at_800)
print("E3_at_850:", E3_at_850)
print("E4_at_600:", E4_at_600)
print("E4_at_650:", E4_at_650)

# Save data in Excel file
data = {
    'Measurement': ['E1_at_850', 'E1_at_900', 'E2_at_650', 'E2_at_700', 'E3_at_800', 'E3_at_850', 'E4_at_600', 'E4_at_650'],
    'Value': [E1_at_850, E1_at_900, E2_at_650, E2_at_700, E3_at_800, E3_at_850, E4_at_600, E4_at_650]
}

# Create a DataFrame
df = pd.DataFrame(data)

directory = r"C:\Users\noorv\OneDrive\Picture's\AW - jaar 3\BSc Thesis\Edited Excel files"
file_path = os.path.join(directory, 'Strain Rate Viscosity.xlsx')

# Write the DataFrame to an Excel file
df.to_excel(file_path, index=False)

print(f"Data has been written to '{file_path}'")

