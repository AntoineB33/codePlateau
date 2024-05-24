import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from sklearn.cluster import DBSCAN
from scipy.signal import find_peaks

# Assuming the data is available in a CSV file or DataFrame with 'x' and 'y' columns
# For example purposes, we generate some mock data

# Generate mock data (replace this with actual data loading)
np.random.seed(42)
x = np.linspace(100000, 500000, 1000)
y = 700 - (np.sin(x / 5000) * 20 + np.random.normal(0, 5, len(x)))

# Create a DataFrame
data = pd.DataFrame({'x': x, 'y': y})

# Plot the data to visualize
plt.plot(data['x'], data['y'])
plt.title("Original Data")
plt.xlabel("x")
plt.ylabel("y")
plt.show()

# Use DBSCAN to find flat regions
# We first compute the difference between successive y-values
dy = np.diff(data['y'])
dy = np.append(dy, dy[-1])  # Keep the length the same

# Perform DBSCAN clustering on the derivative
epsilon = 0.1  # This threshold can be tuned
db = DBSCAN(eps=epsilon, min_samples=10).fit(dy.reshape(-1, 1))

# Extract flat intervals
labels = db.labels_
flat_intervals = []
current_interval = []

for i, label in enumerate(labels):
    if label == 0:  # 0 indicates flat region
        current_interval.append(data['x'].iloc[i])
    else:
        if current_interval:
            flat_intervals.append((current_interval[0], current_interval[-1]))
            current_interval = []

# Handle the last interval if it was flat
if current_interval:
    flat_intervals.append((current_interval[0], current_interval[-1]))

# Print the flat intervals
print("Flat intervals:", flat_intervals)

# Plot the flat regions on the original plot
plt.plot(data['x'], data['y'], label='Original Data')
for interval in flat_intervals:
    plt.axvspan(interval[0], interval[1], color='red', alpha=0.3)
plt.title("Flat Regions Detected")
plt.xlabel("x")
plt.ylabel("y")
plt.legend()
plt.show()
