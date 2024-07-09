import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Generate a simple polynomial in the form of an "M"
x = np.linspace(0, 10, 100)
y = -((x - 2)**2 - 4) * ((x - 8)**2 - 4) / 16

# Create the DataFrame
df = pd.DataFrame({'time': x, 'Ptot': y})

# Example windows
windows = [(1, 3), (7, 9)]

# Number of desired intervals
target_intervals = 2

def count_intervals_above_line(df, height, windows):
    above_line = df['Ptot'] > height
    intervals = []
    in_interval = False
    start = None

    for i, row in df.iterrows():
        if above_line[i]:
            if not in_interval:
                in_interval = True
                start = row['time']
        else:
            if in_interval:
                end = row['time']
                in_interval = False
                intervals.append((start, end))
    
    if in_interval:
        end = df.iloc[-1]['time']
        intervals.append((start, end))

    valid_intervals = []
    for interval in intervals:
        for window in windows:
            if interval[0] <= window[1] and interval[1] >= window[0]:
                valid_intervals.append(interval)
                break

    return len(valid_intervals), valid_intervals

def find_minimum_height(df, target_intervals, windows):
    unique_heights = sorted(df['Ptot'].unique())
    low, high = 0, len(unique_heights) - 1
    result = None
    valid_intervals = []

    while low <= high:
        mid = (low + high) // 2
        height = unique_heights[mid]
        count, intervals = count_intervals_above_line(df, height, windows)

        if count == target_intervals:
            result = height
            valid_intervals = intervals
            high = mid - 1
        elif count < target_intervals:
            high = mid - 1
        else:
            low = mid + 1

    return result, valid_intervals

min_height, valid_intervals = find_minimum_height(df, target_intervals, windows)
print("Minimum height:", min_height)

# Plotting the graph
plt.figure(figsize=(10, 6))
plt.plot(df['time'], df['Ptot'], label='Ptot')
plt.axhline(y=min_height, color='r', linestyle='--', label=f'Minimum height = {min_height}')

for interval in valid_intervals:
    plt.axvspan(interval[0], interval[1], color='green', alpha=0.3)

for window in windows:
    plt.axvspan(window[0], window[1], color='blue', alpha=0.2)

plt.xlabel('Time')
plt.ylabel('Ptot')
plt.title('Ptot against Time with Valid Intervals and Minimum Height Line')
plt.legend()
plt.show()
