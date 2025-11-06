from pathlib import Path
import pandas as pd
import sys
import os

def batch(FILENAME):
    downloads_folder = str(Path.home() / "Downloads")
    output_dir = os.path.join(downloads_folder, "ExcelBatcher")
    os.makedirs(output_dir, exist_ok=True)

    xls = pd.ExcelFile(FILENAME)
    try:
        if len(xls.sheet_names) == 1:
            df = pd.read_excel(xls, xls.sheet_names[0])
            for i in range(0, df.shape[0], 100):
                df_chunk = df.iloc[i:i+100]
                output_path = os.path.join(
                    output_dir,
                    f"{os.path.splitext(os.path.basename(FILENAME))[0]}_{i//100 + 1}.xlsx")
                df_chunk.to_excel(output_path, index=False)
            return output_dir
        else:
            for sheet_name in xls.sheet_names:
                safe_sheet_name = "".join(c if c.isalnum() or c in (' ', '_') else '_' for c in sheet_name)
                df = pd.read_excel(xls, sheet_name)
                for i in range(0, df.shape[0], 100):
                    df_chunk = df.iloc[i:i+100]
                    output_path = os.path.join(
                        output_dir,
                        f"{os.path.splitext(os.path.basename(FILENAME))[0]}_{safe_sheet_name}_{i//100 + 1}.xlsx")
                    df_chunk.to_excel(output_path, index=False)
            return output_dir
    finally:
        xls.close()

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python batch.py <FILENAME>")
    else:
        FILENAME = sys.argv[1]
        batch(FILENAME)