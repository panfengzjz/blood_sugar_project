import pandas as pd
import os

def print_excel(test_name, log_list, fileNumber=1): 
    if (fileNumber == 1):
        file_path = "./shulie.xlsx"
    elif (fileNumber == 2):
        file_path = "./bianliang.xlsx"
    elif (fileNumber == 3):
        file_path = "./patientday.xlsx"

    dic1 = {test_name: log_list}
    cur_df = pd.DataFrame(dic1)
    if os.path.exists(file_path):
        old_df = pd.read_excel(file_path)
        cur_df = pd.concat([old_df, cur_df], axis=1)
    cur_df.to_excel(file_path, index=False, header=False)