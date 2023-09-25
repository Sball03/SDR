import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

# ask for prin
ask_if_print = input("Do you want to print the results to terminal? (y/n)")

# File input/output
input_file = input("Enter the path to the input Excel file: ").strip()
output_file = input("Enter the path to the output Excel file: ").strip()

# Specific sheet input/output
sheet_name_input = input("Please enter a sheet name to read from: ").strip()
sheet_name_output = input("Please enter a sheet name to write to: ").strip()

# Get number of cycles
num_cycles = int(input("Please enter a number of ranges you want to enter (int): "))

df_main = pd.DataFrame()

# Load the existing Excel workbook
book = load_workbook(input_file)

# Where to write to
write_loc = input("Please enter a cell to write to (top-leftmost cell in form A:10): ").strip()
write_loc_list = write_loc.split(":")
write_col = column_index_from_string(write_loc_list[0])  # Convert column letter to integer
write_row = int(write_loc_list[1])

#storing number of columns to use for sample id
width_vals = []
for i in range(num_cycles):

    # Time row and col
    time_col = input("Please enter the time column to read from: ").strip()
    time_rows = input("Please enter the range of rows for the time column to read from (1:100): ").strip()
    rows_list = time_rows.split(":")

    # Values row and col
    value_col = input("Please enter a range of columns to read from (A:B): ").strip()


    # Read the time/id column
    df_time = pd.read_excel(
        input_file,
        sheet_name = sheet_name_input,
        header = None,
        skiprows = range(1, int(rows_list[0])),
        nrows = int(rows_list[1]) - int(rows_list[0]),
        usecols = time_col
    )

    # Read the values 
    df_value = pd.read_excel(
        input_file,
        sheet_name = sheet_name_input,
        header = None,
        skiprows = range(1, int(rows_list[0])),
        nrows = int(rows_list[1]) - int(rows_list[0]),
        usecols = value_col
    )

    # Combine them into one data frame
    df_combined = pd.concat([df_time, df_value], axis = 1)

    # Remove NaN values
    df_combined.dropna(axis = 1, inplace = True)
    
    #assigning the width values after the data frame is the right shape
    width_vals.append(df_combined.shape[1] - 1)

    print(width_vals[i])

    # Melt the data frame and add it to main
    df_melted = df_combined.melt(id_vars = [0], var_name = 'var', value_name = 'Value')
    del df_melted['var']
    df_main = pd.concat([df_main, df_melted])

# Add a Measurement ID column to the left
measurement_ids = [f"Measurement{i}" for i in range(len(df_main))]

# Add signal and sample IDs
signal_ids = []
sample_ids = []

for i in range(num_cycles):
    for j in range(len(df_main)):
        signal_ids.append(f"Signal{i + 1}")
    for l in range((width_vals[i])):
        for k in range(int(rows_list[1]) - 1):
            sample_ids.append(f"Sample{l + 1}")


# Adding Signal, Sample, and Measurement IDs
df_main.insert(0, "MeasurementID", measurement_ids)
df_main.insert(1, "SampleID", sample_ids)
df_main.insert(2, "SignalID", signal_ids)

# Create a Pandas Excel writer using openpyxl engine
with pd.ExcelWriter(output_file, engine = 'openpyxl') as writer:
    writer.book = book

    # Write the cleaned data to the output Excel sheet without overwriting
    df_main.to_excel(writer, sheet_name = sheet_name_output, index = False, header = False, startrow = write_row , startcol = write_col )

if ask_if_print == "y":
    print(df_main)
