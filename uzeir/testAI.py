import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Conv1D, MaxPooling1D, Flatten, Dense, Dropout
from tensorflow.keras.utils import to_categorical
from tensorflow.keras.optimizers import Adam

# Load your data
# Assuming data is a pandas DataFrame with columns 'time', 'weight', and 'class'
data = pd.read_csv(r"C:\Users\abarb\Documents\travail\stage et4\travail\codePlateau\data\A envoyer_antoine(non corrompue)\A envoyer\Expériences plateaux\18_06_24_Benjamin_Roxane.csv")

# Normalize the sensor data
scaler = StandardScaler()
data[['P1', 'P2']] = scaler.fit_transform(data[['P1', 'P2']])

# Define a function to create segments and labels
def create_segments_and_labels(data, window_size=50, step=10):
    segments = []
    labels = []
    for start in range(0, len(data) - window_size, step):
        end = start + window_size
        segment = data[['P1', 'P2']].iloc[start:end].values
        segments.append(segment)
        # Replace this with your actual way of getting labels
        labels.append(0)  # Dummy label, replace with actual class label for the segment
    return np.array(segments), np.array(labels)

# Create segments and labels
window_size = 50
step = 10
X, y = create_segments_and_labels(data, window_size, step)

# Convert class labels to one-hot encoding
y = to_categorical(y)

# Split data into training, validation, and test sets
X_train, X_temp, y_train, y_temp = train_test_split(X, y, test_size=0.3, random_state=42)
X_val, X_test, y_val, y_test = train_test_split(X_temp, y_temp, test_size=0.5, random_state=42)

# Build the CNN model
model = Sequential()
model.add(Conv1D(filters=64, kernel_size=3, activation='relu', input_shape=(window_size, 2)))
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
model.save('eating_utensil_classifier.h5')

# Plot training history
import matplotlib.pyplot as plt

plt.plot(history.history['accuracy'], label='train accuracy')
plt.plot(history.history['val_accuracy'], label='val accuracy')
plt.xlabel('Epoch')
plt.ylabel('Accuracy')
plt.legend(loc='lower right')
plt.show()