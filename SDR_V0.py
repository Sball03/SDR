import pandas as pd

user_input_file = input("Please enter a file name or path for data input: ") 
#user_input_file_sheet = input("Please enter a sheet name for data input: ")
user_output_file = input("Please enter a file name or path for data output: ")


df = pd.read_excel(user_input_file)

df.to_excel(user_output_file,sheet_name = 'Sheet1',index = False)

print(df)

