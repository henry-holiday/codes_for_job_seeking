import pandas as pd
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

# pd.set_option('display.max_rows', None)
# pd.set_option('display.max_columns', None)

# basic pathlib set up 
target_folder= r"C:\Users\G1659028\Downloads\Granger data cleanning practice" # chang it for input () ? 
Input_folder=Path(target_folder)
output_path = Input_folder / "output_folder"
output_path.mkdir(parents=True, exist_ok=True)


# Use glob to find Excel files list that is ready for parse
file_list = list(Input_folder.glob("*.xlsx")) 
# 先读取第一个文件以确定列数，避免未定义变量错误
try:
    first_file = file_list[0]
    # 临时读取以获取列数，不保留数据
    first_file_df = pd.read_excel(first_file, sheet_name='CPK', skiprows=4, header=0)
except Exception as e:
    print(f"读取第一个文件失败以获取列数: {e}")

    # cols to keep for granger parsing
cols_to_drop=[4, 5,8,9,14,15,16,18,19,20,21,22,23,24,25,26,59,60,61,62,63,64,65]
cols_to_keep = [i for i in range(len(first_file_df.columns)) if i not in cols_to_drop]


def Parse_granger_dimension_data(file_list,output_folder):

    Granger_data=[]
    for file in file_list:
        try:
            
            df=pd.read_excel(file,sheet_name='CPK',
                          skiprows=5,
                          usecols=cols_to_keep)
        except ValueError as e:
            print(f"Warning: Sheet 'CPK' not found in {file}. Skipping file. Error: {e}")
            continue


        df=df.iloc[3:].reset_index(drop=True)
        IQC_Inspection_lot = pd.read_excel(file, sheet_name='CPK').iloc[0,27]
        # get IQC Inspection lot
        # format it to /Y/m/d format
        IQC_Inspection_lot =IQC_Inspection_lot.strftime('%Y/%m/%d')
        df["IQC_Inspection_lot"] =IQC_Inspection_lot
        # move IQC_Inspection_lot to be the first column
        df = df[['IQC_Inspection_lot'] + [col for col in df.columns if col != 'IQC_Inspection_lot']]

    # append the df outside the for loop
        Granger_data.append(df)

    # end the for loop here
    Granger_measurement_df= pd.concat(Granger_data, ignore_index=True)


    columns_to_melt = list(range(1, 33))
    # Melt only those columns
    melted_df = Granger_measurement_df.melt(
    id_vars=[col for col in Granger_measurement_df.columns if col not in columns_to_melt],  # Keep other columns as ID
    value_vars=columns_to_melt,
    var_name='sample_number',
    value_name='value'
    )
    # make all the column name upper case 
    Granger_measurement_df.columns = Granger_measurement_df.columns.astype(str).str.upper()
    melted_df.columns = melted_df.columns.str.upper()
    df_specs= melted_df.groupby('FAI. NO.')[['USL','LSL']].first().T
    print(df_specs.columns.name)
    df_specs.columns.name=None
    display(df_specs)
    melted_df['row_idx'] = melted_df.groupby(['IQC_INSPECTION_LOT', 'FAI. NO.']).cumcount()
    FAI_list_wide_format_df = melted_df .pivot(index=['IQC_INSPECTION_LOT', 'row_idx'], columns='FAI. NO.', values='VALUE')
    FAI_list_wide_format_df.columns.name = None
    
    df_wide_format = pd.concat([df_specs, FAI_list_wide_format_df])
    display(df_wide_format)
    # extract only the first element from the index tuple and brocast it to a index columns
    df_wide_format=df_wide_format.reset_index()
    df_wide_format['index']=df_wide_format['index'].map(lambda x : x[0] if isinstance(x,tuple) else x)
    # rename the index title to IQC LOT
    df_wide_format.rename(columns ={'index': 'IQC_LOT'},inplace=True)
    df_specs.to_csv(Input_folder/output_path/"Granger_dimension_spec.csv",index=False)
    df_wide_format.to_excel(Input_folder/output_path/"Granger_wide_data.xlsx",engine='openpyxl',index=True)


    melted_df.to_excel(Input_folder/output_path/"Granger_stacked_data.xlsx",engine='openpyxl',index=False)
    Granger_measurement_df.to_excel(Input_folder/output_path/"Granger_measurement_data.xlsx",engine='openpyxl',index=False)

Parse_granger_dimension_data(file_list,output_path)
print("done")
