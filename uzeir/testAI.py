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

# Preprocessing
# Normalize the weight data
scaler = StandardScaler()
data['weight'] = scaler.fit_transform(data['weight'].values.reshape(-1, 1))

# Assuming each segment is identified by a unique segment_id
segments = data['segment_id'].unique()
X = []
y = []

for segment_id in segments:
    segment_data = data[data['segment_id'] == segment_id]
    X.append(segment_data['weight'].values)
    y.append(segment_data['class'].iloc[0])

X = np.array(X)
y = np.array(y)

# Convert class labels to one-hot encoding
y = to_categorical(y)

# Split data into training, validation, and test sets
X_train, X_temp, y_train, y_temp = train_test_split(X, y, test_size=0.3, random_state=42)
X_val, X_test, y_val, y_test = train_test_split(X_temp, y_temp, test_size=0.5, random_state=42)

# Build the CNN model
model = Sequential()
model.add(Conv1D(filters=64, kernel_size=3, activation='relu', input_shape=(X_train.shape[1], 1)))
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
