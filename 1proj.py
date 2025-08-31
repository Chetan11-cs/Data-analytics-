import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt

# 🔹 Updated Clean Function (merged old + new)
def get_Null_values(file_path, sheet_Name=0):
    try:
        # Step 1: Load data
        df = pd.read_excel(file_path, sheet_name=sheet_Name)

        # Step 2: Clean data
        df_clean = df.dropna(how='all')              # Remove empty rows
        df_clean = df_clean.dropna(axis=1, how='all')  # Remove empty cols
        df_clean.fillna(0, inplace=True)             # Replace remaining NaN with 0

        # Step 3: Track empty cell addresses (like old code)
        non_null_address = []
        for row_ind, row in df_clean.iterrows():
            for col_ind, values in enumerate(row):
                if pd.notna(values):   # A1, B1, C1...
                    col_letter = chr(65 + col_ind)
                    non_null_address.append(f"{col_letter}{row_ind+1}")

        # Step 4: Show preview in GUI
        preview_data = df_clean.to_string(index=False)
        text_box.delete("1.0", tk.END)
        text_box.insert(tk.END, preview_data)

        # Step 5: Show graph of first numeric column
        numeric_cols = df_clean.select_dtypes(include='number').columns
        if not numeric_cols.empty:
            col = numeric_cols[0]
            plt.figure(figsize=(10, 5))
            df_clean[col].head(10).plot(kind='bar')
            plt.title(f"Bar Chart of '{col}'")
            plt.xlabel("Index")
            plt.ylabel(col)
            plt.tight_layout()
            plt.show()
        else:
            messagebox.showinfo("No Numeric Data", "No numeric columns found to plot.")

        return non_null_address

    except Exception as e:
        messagebox.showerror("Error", str(e))
        return []


# 🔹 GUI Function
def upload_file():
    file_path = filedialog.askopenfilename(
        title="Select the Excel file",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    result = get_Null_values(file_path)
    print("Non-null cell addresses:", result)


# 🔹 Tkinter GUI Setup
root = tk.Tk()
root.title("Excel Data Viewer")
root.geometry("1000x600")

upload_btn = tk.Button(root, text="Upload File", command=upload_file)
upload_btn.pack(pady=10)

text_box = tk.Text(root, wrap=tk.NONE, font=("Courier", 10))
text_box.pack(expand=True, fill="both", padx=10, pady=10)

scroll_y = tk.Scrollbar(root, orient="vertical", command=text_box.yview)
scroll_y.pack(side="right", fill="y")
text_box.configure(yscrollcommand=scroll_y.set)

scroll_x = tk.Scrollbar(root, orient="horizontal", command=text_box.xview)
scroll_x.pack(side="bottom", fill="x")
text_box.configure(xscrollcommand=scroll_x.set)

root.mainloop()
