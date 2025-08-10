import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import re

df = pd.read_excel("C:/Users/ce/Desktop/python/PandasPython/Company_Dataset_Irregular.xlsx")

# -----------------------
# 2. Fix wrong characters in object (string) columns
# -----------------------
Exclude = "Email"
for col in df.select_dtypes(include="object").columns:
    if col != Exclude:
        df[col] = df[col].astype(str) \
            .str.replace("3", "e", regex=False) \
            .str.replace("#", "h", regex=False) \
            .str.replace("@", "a", regex=False)
df["Email"] = df["Email"].astype(str) \
    .str.replace("3", "e", regex=False) \
    .str.replace("#", "h", regex=False)

# -----------------------
# 3. Convert numeric columns safely and Calculating Median of Salary
# -----------------------
numeric_cols = ["Salary", "Age", "WorkHours", "PerformanceScore"]
if "Salary" in df.columns:
    df["Salary"] = pd.to_numeric(df["Salary"], errors="coerce")
    df.loc[df["Salary"] > 1e9, "Salary"] = np.nan 
    Mid_Sal = df["Salary"].median(skipna=True)    
    df["Salary"] = df["Salary"].fillna(Mid_Sal)
for col in numeric_cols:
    if col in df.columns and col != "Salary":
        df[col] = pd.to_numeric(df[col], errors="coerce")
        median_val = df[col].median(skipna=True)
        df[col] = df[col].fillna(median_val)

# -----------------------
# 4. Fill numeric NaN with column mean
# -----------------------
for col in numeric_cols:
    if col in df.columns:
        mean_val = df[col].mean(skipna=True)
        df[col] = df[col].fillna(mean_val)

# -----------------------
# 5. Filling text NaN with "Unknown"
# -----------------------
text_cols = df.select_dtypes(include="object").columns
for col in text_cols:
    df[col] = df[col].replace("nan", np.nan) 
    df[col] = df[col].fillna("Unknown")

# -----------------------
# 6. Replacing "Unknown" in JoinDate with random dates
# -----------------------
if "JoinDate" in df.columns:
    def random_date(start_date, end_date):
        delta = end_date - start_date
        random_days = random.randint(0, delta.days)
        return start_date + timedelta(days=random_days)

    start = datetime(2018, 1, 1)
    end = datetime(2023, 12, 31)

    df["JoinDate"] = df["JoinDate"].astype(str)
    df["JoinDate"] = df["JoinDate"].apply(
        lambda x: random_date(start, end) if "unknown" in x.lower() else pd.to_datetime(x, errors="coerce")
    )
    df["JoinDate"] = df["JoinDate"].fillna("Unknown")

# -----------------------
# 7. Cleaning Phone Numbers
# -----------------------
def clean_phone(phone):
    phone = str(phone).strip()
    
    if phone.lower() in ["unknown", "nan", ""]:
        return np.nan
    
    phone = re.sub(r"x\d+", "", phone, flags=re.IGNORECASE)
    phone = re.sub(r"[a-zA-Z]", "", phone)
    phone = re.sub(r"[^\d+]", "", phone)
    return phone
if "PhoneNumber" in df.columns:
    df["PhoneNumber"] = df["PhoneNumber"].apply(clean_phone)
    df["PhoneNumber"]=df["PhoneNumber"].fillna("Not Filled")
# -----------------------
# 8. cleaned dataset
# -----------------------
df.to_excel("Cleaned.xlsx", index=False)
print(df.isnull().sum())
print(df.dtypes)
print(df.duplicated().sum())
print("Dataset cleaned and saved successfully!")
