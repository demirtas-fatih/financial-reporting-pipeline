"""
generate_data.py
Generates realistic sample financial data for the reporting pipeline.
"""

import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

np.random.seed(42)
random.seed(42)

DEPARTMENTS = ["Engineering", "Marketing", "Sales", "Operations", "HR", "Finance", "IT"]
CATEGORIES  = ["Personnel", "Software", "Hardware", "Travel", "Consulting", "Office", "Training"]
MONTHS      = pd.date_range("2024-01-01", periods=12, freq="MS")

rows = []
for month in MONTHS:
    for dept in DEPARTMENTS:
        budget   = round(random.uniform(20_000, 120_000), 2)
        variance = random.uniform(-0.25, 0.30)
        actual   = round(budget * (1 + variance), 2)
        for _ in range(random.randint(4, 10)):
            rows.append({
                "date":       month + timedelta(days=random.randint(0, 27)),
                "department": dept,
                "category":   random.choice(CATEGORIES),
                "budget":     round(budget / random.randint(3, 8), 2),
                "actual":     round(actual  / random.randint(3, 8), 2),
                "approved":   random.choice([True, True, True, False]),
            })

df = pd.DataFrame(rows)
df.sort_values("date", inplace=True)
df.reset_index(drop=True, inplace=True)
df.to_csv("data/financial_data.csv", index=False)
print(f"✅ Generated {len(df)} rows → data/financial_data.csv")
