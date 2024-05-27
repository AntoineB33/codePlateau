import pandas as pd
import numpy as np
import os
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.metrics import classification_report
import matplotlib.pyplot as plt
import glob

# Step 1: Load CSV files
def load_csv_files(directory):
    files = glob.glob(os.path.join(directory, "*0.csv"))
    data = []
    for file in files:
        df = pd.read_csv(file)
        df.columns = ["time", "Ptot"]  # Assigner les noms de colonnes
        data.append(df)
    return data

# Step 2: Data Preprocessing
def preprocess_data(data):
    preprocessed_data = []
    for df in data:
        df = df.dropna()  # Drop missing values
        df['time'] = pd.to_datetime(df['time'])
        df = df.sort_values('time')  # Ensure data is sorted by time
        df['time_diff'] = df['time'].diff().dt.total_seconds().fillna(0)
        df['Ptot_diff'] = df['Ptot'].diff().fillna(0)
        preprocessed_data.append(df)
    return preprocessed_data

# Step 3: Feature Engineering
def extract_features(data):
    feature_data = []
    labels = []
    for df in data:
        print(df.head())
        for i in range(1, len(df)):
            features = {
                'time_diff': df.iloc[i]['time_diff'],
                'Ptot_diff': df.iloc[i]['Ptot_diff'],
                'Ptot': df.iloc[i]['Ptot']
            }
            feature_data.append(features)
            labels.append(1 if df.iloc[i]['Ptot_diff'] < -1 else 0)  # Label as bite if Ptot drop is significant
    return pd.DataFrame(feature_data), np.array(labels)

# Step 4: Model Training
def train_model(features, labels):
    X_train, X_test, y_train, y_test = train_test_split(features, labels, test_size=0.2, random_state=42)
    model = RandomForestClassifier(n_estimators=100, random_state=42)
    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)
    print(classification_report(y_test, y_pred))
    return model

# Step 5: Bite Detection
def detect_bites(model, data):
    predictions = []
    for df in data:
        features = {
            'time_diff': df['time_diff'],
            'Ptot_diff': df['Ptot_diff'],
            'Ptot': df['Ptot']
        }
        features_df = pd.DataFrame(features)
        preds = model.predict(features_df)
        df['bite'] = preds
        predictions.append(df)
    return predictions

# Main function to run the pipeline
def main():
    directory = r'data\Donnees_brutes_csv'  # Replace with your directory path
    # directory = r'..\donneexslx\donneexslx'  # Replace with your directory path
    data = load_csv_files(directory)
    preprocessed_data = preprocess_data(data)
    features, labels = extract_features(preprocessed_data)
    model = train_model(features, labels)
    predictions = detect_bites(model, preprocessed_data)
    
    # Visualize results for the first file as an example
    plt.figure(figsize=(12, 6))
    plt.plot(predictions[0]['time'], predictions[0]['Ptot'], label='Ptot')
    plt.scatter(predictions[0][predictions[0]['bite'] == 1]['time'], predictions[0][predictions[0]['bite'] == 1]['Ptot'], color='red', label='Bites')
    plt.xlabel('Time')
    plt.ylabel('Ptot')
    plt.legend()
    plt.show()

if __name__ == "__main__":
    main()
