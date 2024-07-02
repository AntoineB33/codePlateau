import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Conv1D, MaxPooling1D, Flatten, Dense, Dropout
from tensorflow.keras.utils import to_categorical
from tensorflow.keras.optimizers import Adam
import matplotlib.pyplot as plt

# Load your data
data = pd.read_csv('your_data.csv')

# Preprocessing
# Normalize the weight data (P1 and P2)
scaler = StandardScaler()
data[['P1', 'P2']] = scaler.fit_transform(data[['P1', 'P2']])

# Assuming you have a list of tuples with time duration and a boolean indicating if it's a peak
# Example: segments = [(100, False), (200, True), ...]
segments = [(200, False), (300, True)]  # Replace with your actual segment information

# Function to segment the data
def create_segments_and_labels(data, segments):
    X = []
    y = []
    start_index = 0
    for duration, is_peak in segments:
        end_index = start_index + duration
        if end_index > len(data):
            break
        segment_data = data.iloc[start_index:end_index]
        X.append(segment_data[['P1', 'P2']].values)
        y.append(1 if is_peak else 0)
        start_index = end_index
    return np.array(X), np.array(y)

X, y = create_segments_and_labels(data, segments)

# Ensure all segments have the same length
max_length = max([len(segment) for segment in X])
X_padded = np.zeros((len(X), max_length, 2))

for i, segment in enumerate(X):
    X_padded[i, :len(segment), :] = segment

# Convert class labels to one-hot encoding
y = to_categorical(y)

# Split data into training, validation, and test sets
X_train, X_temp, y_train, y_temp = train_test_split(X_padded, y, test_size=0.3, random_state=42)
X_val, X_test, y_val, y_test = train_test_split(X_temp, y_temp, test_size=0.5, random_state=42)

# Build the CNN model
model = Sequential()
model.add(Conv1D(filters=64, kernel_size=3, activation='relu', input_shape=(X_train.shape[1], X_train.shape[2])))
model.add(MaxPooling1D(pool_size=2))
model.add(Dropout(0.5))
model.add(Flatten())
model.add(Dense(100, activation='relu'))
model.add(Dense(y.shape[1], activation='softmax'))

# Compile the model
model.compile(optimizer=Adam(), loss='categorical_crossentropy', metrics=['accuracy'])

# Train the model
history = model.fit(X_train, y_train, epochs=30, batch_size=32, validation_data=(X_val, y_val))

# Evaluate the model
test_loss, test_acc = model.evaluate(X_test, y_test)
print(f'Test accuracy: {test_acc}')

# Save the model
model.save('segment_peak_classifier.h5')

# Plot training history
plt.plot(history.history['accuracy'], label='train accuracy')
plt.plot(history.history['val_accuracy'], label='val accuracy')
plt.xlabel('Epoch')
plt.ylabel('Accuracy')
plt.legend(loc='lower right')
plt.show()
