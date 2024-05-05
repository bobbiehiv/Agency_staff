
import pandas as pd

def generate_file_path(base_directory, date_str):
    return f"{base_directory}\\{date_str}.xlsx"  # MSPNetworkActivityReport 

def compare_excel_sheets(file1, file2, sheet_name):
    data1 = pd.read_excel(file1, sheet_name=sheet_name, usecols="E,L", names=['Name', 'Date'])
    data2 = pd.read_excel(file2, sheet_name=sheet_name, usecols="E,L", names=['Name', 'Date'])
    merged_data = pd.merge(data1, data2, on='Name', how='outer', indicator=True, suffixes=('_old', '_new'))
    
    updates = merged_data[(merged_data['_merge'] == 'both') & (merged_data['Date_old'] != merged_data['Date_new'])]
    updates = updates[['Name', 'Date_new']]
    new_names = merged_data[merged_data['_merge'] == 'right_only'][['Name', 'Date_new']]
    removed_names = merged_data[merged_data['_merge'] == 'left_only'][['Name', 'Date_old']]

    return {'Updates': updates, 'Added': new_names, 'Removed': removed_names}

def compare_agency_lists(base_directory, date1, date2):
    file1 = generate_file_path(base_directory, date1)
    file2 = generate_file_path(base_directory, date2)
    results = {
        'On Assignment': compare_excel_sheets(file1, file2, 'On Assignment'),
        'Future Extension': compare_excel_sheets(file1, file2, 'Future Extension')
    }
    return results

def print_right_aligned(df):
    # Convert DataFrame to string
    df_string = df.to_string(index=False)
    # Determine console width
    width = 35  # You can adjust this width based on your console's width
    # Right align each line of the DataFrame string
    formatted_string = "\n".join("{:>{width}}".format(line, width=width) for line in df_string.split('\n'))
    print(formatted_string)


if __name__ == "__main__":
    base_directory = "C:\\Agency lists"
    print("Please enter the entire name of the first file:")
    date1 = input()
    print("Please enter the entire name of the second file:")
    date2 = input()

    results = compare_agency_lists(base_directory, date1, date2)



print('\n'"Updates in On Assignment:"'\n')

print_right_aligned(results['On Assignment']['Updates'])

print('\n'"Added in On assignment:"'\n')

print_right_aligned(results['On Assignment']['Added'])

print('\n'"Removed from On Assignment:"'\n')

print_right_aligned(results['On Assignment']['Removed'])

print('\n'"Updates in Future Extension:"'\n')

print_right_aligned(results['Future Extension']['Updates'])

print('\n'"Added in Future Extensions:"'\n')

print_right_aligned(results['Future Extension']['Added'])

print('\n'"Removed from Future Extensions:"'\n')

print_right_aligned(results['Future Extension']['Removed'])

print('\n',' Completed '.center(45,'*'),'\n')