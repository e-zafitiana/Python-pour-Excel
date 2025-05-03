import pandas as pd

df = pd.DataFrame({
    "Employee": ["Alice", "Charlie"],
    "Age": [23, 37],
    "Salary": [35000, 70000]
})

df.to_excel("Week0.xlsx", index=False, engine="openpyxl")
df