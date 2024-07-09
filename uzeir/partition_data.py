import numpy as np
import pandas as pd

# Example DataFrame
data = {
    'time': np.arange(0, 100, 0.1),
    'Ptot': np.sin(np.arange(0, 100, 0.1)) + 0.5 * np.random.randn(len(np.arange(0, 100, 0.1))) + 5
}
df = pd.DataFrame(data)

def count_intervals(df, height):
    above = df['Ptot'] > height
    intervals = (above != above.shift()).cumsum()
    num_intervals = len(intervals[above].unique())
    return num_intervals

def find_min_height_for_intervals(df, target_intervals, precision=1e-5):
    low, high = df['Ptot'].min(), df['Ptot'].max()
    
    while high - low > precision:
        mid = (low + high) / 2
        intervals = count_intervals(df, mid)
        
        if intervals >= target_intervals:
            low = mid
        else:
            high = mid
            
    return (low + high) / 2

# Set the target number of intervals
target_intervals = 5

# Find the minimum height
min_height = find_min_height_for_intervals(df, target_intervals)

print(f"The minimum height for {target_intervals} intervals is: {min_height:.5f}")
