import matplotlib.pyplot as plt 
import numpy as np 
import pandas as pd 
from matplotlib.font_manager import FontProperties

font_prop = FontProperties(family='Verdana', size=45)

# Read Excel file and load data
file_path = "KMA22-02.xlsx"
file = pd.read_excel(file_path, 'Mean Orientation Spread')

mos = file['Mean'][1:]
counts = file['Cumulative'][1:] 

start_point = [counts.iloc[0], mos.iloc[0]]
end_point = [counts.iloc[-1], mos.iloc[-1]]

sorted_indices = np.argsort(mos)

counts_sorted = counts.iloc[sorted_indices]
mos_sorted = mos.iloc[sorted_indices]

data_sorted = np.column_stack((counts_sorted, mos_sorted))

# Define elbow function
def find_elbow(data, theta):
    """
    Find the index of the elbow point in the given data by rotating the data 
    vector and identifying the point with the minimum y-value in the rotated data.
    
    Parameters:
    data: A 2D array where each row represents a data point in the form [x, y].
    theta: The angle in radians by which to rotate the data.
    
    Returns:
    int: The index of the elbow point in the original data.
    """
    co = np.cos(theta)
    si = np.sin(theta)
    rotation_matrix = np.array(((co, -si), (si, co)))

    rotated_vector = np.dot(data, rotation_matrix)

    return np.argmin(rotated_vector[:, 1])

# Define function to calculate radiant for data rotation
def get_data_radiant(data):
    """
    Calculate the angle in radians for data rotation based on the diagonal of the data's bounding box.
    
    Parameters:
    data: A 2D array where each row represents a data point in the form [x, y].
    
    Returns:
    float: The angle in radians for rotating the data, computed as the arctangent of 
           the vertical span divided by the horizontal span of the data's bounding box.
    """
    return np.arctan2(data[:, 1].max() - data[:, 1].min(), 
                      data[:, 0].max() - data[:, 0].min())

theta = get_data_radiant(data_sorted)

# Get the point of maximum curvature (elbow)
elbow_index = find_elbow(data_sorted, theta)

# Plot figure
plt.figure(figsize=(24,16))

plt.plot(counts, mos, marker='o', markersize=2, color='#008080', linewidth=5, label='Mean Orientation Spread')

plt.scatter(data_sorted[elbow_index, 0], data_sorted[elbow_index, 1], color='r', s=200, zorder=3)  # Mark the elbow point

gos_value = data_sorted[elbow_index, 1]
plt.annotate(f'{gos_value:.2f}°', (data_sorted[elbow_index, 0], gos_value),
             textcoords="offset points", xytext=(10,-10), ha='left', color='black', fontsize=45)

plt.plot([data_sorted[:, 0].min(), data_sorted[elbow_index, 0]], [gos_value, gos_value], 
         color='black', linestyle='--', linewidth=6, label='GOS Threshold')

plt.xlabel('Cumulative number of grains', size=50, fontproperties=font_prop) 
plt.ylabel('Grain orientation spread (°)', size=50, fontproperties=font_prop) 
plt.grid(color='gainsboro')
plt.xticks(size=40, c='dimgray')
plt.yticks(size=40, c='dimgray')
plt.legend(prop=font_prop, fontsize=50)
plt.show()

