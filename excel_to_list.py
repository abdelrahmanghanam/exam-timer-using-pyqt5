import pandas as pd

def excel_to_list(file_path, column_name="rules"):
    try:
        df = pd.read_excel(file_path)
        column_list = df[column_name].tolist()
        return column_list
    except Exception as e:
        print(f"Error occurred: {e}")
        return None
