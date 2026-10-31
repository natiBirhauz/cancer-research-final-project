import os
import glob
import pandas as pd
import numpy as np
import re
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.metrics.pairwise import cosine_similarity
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import LSTM, Dense, Dropout
from tensorflow.keras.optimizers import Adam

# Suppress TensorFlow informational messages
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '2' 

def create_sequences(dataset, window_size):
    """
    Creates sequential windows for LSTM model prediction.
    Requires dataset length to be strictly greater than window_size.
    """
    X, y = [], []
    for i in range(len(dataset) - window_size):
        X.append(dataset[i:(i + window_size)])
        y.append(dataset[i + window_size])
    return np.array(X), np.array(y)

# ==========================================
# 1. Scan files and find maximum dimension
# ==========================================
folder_path = 'All_Tail_Proteins_Tables' # Target directory
window_size = 3

excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))
if not excel_files:
    raise FileNotFoundError(f"No excel files found in directory: {folder_path}")

print(f"Found {len(excel_files)} process files.")

pattern = re.compile(r'G\d+[+-]VALUE')

processed_files = []
max_dimensions = 0

print("Starting processing and per-process normalization...\n")

for file in excel_files:
    try:
        df = pd.read_excel(file)
    except Exception as e:
        print(f"[Error] Failed reading file {file}: {e}")
        continue
        
    df.columns = df.columns.astype(str).str.strip().str.upper()
    target_columns = [col for col in df.columns if pattern.search(col)]
    
    if not target_columns:
        continue
        
    # --- CRITICAL FIX 1: Handle NaNs in Raw Data ---
    # Fill any empty cells from Excel with 0.0 and enforce float32 type
    df[target_columns] = df[target_columns].fillna(0.0)
    data = df[target_columns].astype(np.float32).values
    
    # Extra safety: replace any residual infinite values
    data = np.nan_to_num(data, nan=0.0, posinf=0.0, neginf=0.0)
    
    scaler = StandardScaler()
    data_scaled = scaler.fit_transform(data)
    
    current_dim = data_scaled.shape[1]
    if current_dim > max_dimensions:
        max_dimensions = current_dim
        
    processed_files.append({
        'file_name': os.path.basename(file),
        'data': data_scaled,
        'original_dim': current_dim,
        'num_rows': data_scaled.shape[0]
    })

print(f"Maximum protein dimension found across all processes: {max_dimensions}")

# ==========================================
# 2. Zero-Padding and Sequence Generation
# ==========================================
all_X, all_y = [], []
files_skipped_due_to_length = 0

for item in processed_files:
    data_scaled = item['data']
    current_dim = item['original_dim']
    num_rows = item['num_rows']
    
    if num_rows <= window_size:
        files_skipped_due_to_length += 1
        continue

    pad_width = max_dimensions - current_dim
    
    if pad_width > 0:
        padded_data = np.pad(data_scaled, pad_width=((0, 0), (0, pad_width)), mode='constant', constant_values=0)
    else:
        padded_data = data_scaled
        
    X_seq, y_seq = create_sequences(padded_data, window_size)
    
    if len(X_seq) > 0:
        all_X.append(X_seq)
        all_y.append(y_seq)

if len(all_X) == 0:
    raise ValueError(f"No sequences were generated! All processed files had {window_size} or fewer rows.")

if files_skipped_due_to_length > 0:
    print(f"\n[Info] Skipped {files_skipped_due_to_length} files because they didn't have enough rows.")

X = np.vstack(all_X)
y = np.vstack(all_y)

print(f"\nFinal dataset shape - Input (X): {X.shape}, Target (y): {y.shape}")

# ==========================================
# 3. Train/Test Split
# ==========================================
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.3, shuffle=False)

# ==========================================
# 4. Neural Network (LSTM) Definition
# ==========================================
model = Sequential([
    LSTM(64, activation='relu', return_sequences=True, input_shape=(window_size, max_dimensions)),
    Dropout(0.2), 
    LSTM(32, activation='relu', return_sequences=False),
    Dense(max_dimensions) 
])

# --- CRITICAL FIX 2: Gradient Clipping ---
# clipvalue=1.0 prevents the gradients from exploding, which causes NaNs in LSTMs
optimizer = Adam(learning_rate=0.001, clipvalue=1.0)
model.compile(optimizer=optimizer, loss='mse')

print("\nStarting Network Training...")
# Note: if dataset is extremely small, validation_split might fail if batch is too small. 
# Lowered batch_size to 8 to handle small datasets better.
model.fit(X_train, y_train, epochs=30, batch_size=8, validation_split=0.1, verbose=1)

# ==========================================
# 5. Prediction and Cosine Similarity Evaluation
# ==========================================
print("\nEvaluating predictions on Test data...")
predictions = model.predict(X_test)

# Safety check for NaNs in predictions
if np.isnan(predictions).any():
    print("\n[Error] The model predicted NaN values. The data might still contain unscaled anomalies or the learning rate is too high.")
else:
    similarities = []
    for i in range(len(predictions)):
        pred_vector = predictions[i].reshape(1, -1)
        true_vector = y_test[i].reshape(1, -1)
        
        sim = cosine_similarity(pred_vector, true_vector)[0][0]
        similarities.append(sim)

    average_similarity = np.mean(similarities)
    print(f"\n======================================")
    print(f"Average Cosine Similarity on Test Set: {average_similarity:.4f}")
    print(f"======================================")