import pandas as pd
import os

def read_excel_data(dl_path):
    df = pd.read_excel(dl_path)
    return df

    
def delete_file():
    script_dir = os.getcwd()
    filename = "Test_Data_Logging.xlsx"
    os.remove(os.path.normcase(os.path.join(script_dir, filename)))

    


# def write_excel(df, dl_path):
#     df.to_excel(dl_path, index=False)
    

