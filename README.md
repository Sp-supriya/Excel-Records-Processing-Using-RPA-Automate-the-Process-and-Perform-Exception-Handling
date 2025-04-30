
### ✅ *Excel Records Processing Using RPA – Automate the Process and Perform Exception Handling*

---

## 🔧 Step-by-Step Implementation:

### **Step 1: Define the Objective**
Automate reading an Excel sheet, validate records, perform data processing, and log errors without interrupting the whole process (exception handling).

---

### **Step 2: Prepare Sample Excel File (`employees.xlsx`)**
Example columns:
| ID | Name    | Age | Department | Salary |
|----|---------|-----|------------|--------|
| 1  | Alice   | 28  | IT         | 60000  |
| 2  | Bob     | 22  | HR         | 48000  |
| 3  | Charlie |     | Finance    | 55000  |
| 4  | David   | 27  | IT         | -30000 |

> You can create this in Excel and save as `employees.xlsx`.

---

### **Step 3: Upload Excel in Google Colab**
```python
from google.colab import files
uploaded = files.upload()

import pandas as pd
import io

df = pd.read_excel(io.BytesIO(uploaded['employees.xlsx']))
print(df.head())
```

---

### **Step 4: Install Required Libraries**
```python
!pip install pandas openpyxl
```

---

### **Step 5: Process Records with Exception Handling**
```python
success_log = []
error_log = []

for index, row in df.iterrows():
    try:
        # Basic validation
        if pd.isnull(row['Name']) or pd.isnull(row['Age']) or row['Salary'] <= 0:
            raise ValueError(f"Invalid data in row {index + 2}")
        
        # Simulated processing
        print(f"Processing: {row['Name']} - {row['Department']}")

        success_log.append(row.to_dict())
        
    except Exception as e:
        print(f"Error in row {index + 2}: {e}")
        error_log.append({'Row': index + 2, 'Error': str(e)})
```

---

### **Step 6: Save Logs to Excel**
```python
success_df = pd.DataFrame(success_log)
error_df = pd.DataFrame(error_log)

success_df.to_excel("processed_success.xlsx", index=False)
error_df.to_excel("processed_errors.xlsx", index=False)

from google.colab import files
files.download("processed_success.xlsx")
files.download("processed_errors.xlsx")
```

---

## 📝 Summary:
- ✅ Upload Excel file
- ✅ Validate each row (check for missing or wrong values)
- ✅ Process valid rows
- ✅ Log errors without stopping execution
- ✅ Save results and download logs

---

