''' pysimpleGUI for analyzing behavior data and creating all possible group assignments with cage mates being 
    assigned to identical groups and total group assignments have equal subjects within them. The optimal group 
    assignment out of all possible, is determined by finding mean values from hypothetical possible group 
    assignments for each columns data. The group assigment that results in the minimum value when the sum of 
    the mean's between groups is calculated, is assigned as the 'optimal' group. Each step of these 
    calculations create a new data frame that is saved to an Excel sheet along with the original data for 
    reference. The optimal group assigment is graphed for comparision.  
'''

# Import required libraries
import PySimpleGUI as sg
import pandas as pd
from itertools import combinations, permutations, product
import re
import matplotlib.pyplot as plt
import os
from pathlib import Path

# validate that the file paths are entered correctly
def is_valid_path(filepath):
    if filepath and Path(filepath).exists():
        return True
    sg.popup_error("A selected file path is incorrect or has been left empty.")
    return False

# window appears when the program successfully completes
def nom_window(input_filename):
    layout = [[sg.Text("\n"
    " All Systems Nominal.\n"
    f" {input_filename} \n"
    " has been modified\n"
    " to contain info & calcs for\n"
    " optimal group assignments."
    "\n"
    "")]]
    window = sg.Window((""), layout, modal=True)
    choice = None
    while True:
        event, values = window.read()
        if event == "Exit" or event == sg.WIN_CLOSED:
            break
    window.close()
    
# Define the location of the directory
def extract_values_from_excel(input_filename, output_folder):
    name = Path(output_folder)

    # Change the directory
    os.chdir(output_folder)
    print(output_folder)

    # creation of a maximum value for the progress bar function
    max = 7
    prog_bar_update_val = 0

    file_name = input_filename
    # Read the Excel file (assuming it's named 'data.xlsx')
    df = pd.read_excel(file_name)

    # Display the initial DataFrame
    #print("Initial DataFrame:")
    #print(df)

    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))

    ###################### All Possible Group Combinations ########################

    # Find unique cages
    unique_cages = df['cage'].unique()

    # Generate all possible group assignments for each unique cage
    groups = ['a', 'b']
    possible_assignments = list(product(groups, repeat=len(unique_cages)))

    # Create a list to store the resulting DataFrames
    combo_dfs = []

    # Counter for iteration number
    iteration_count = 1

    # Iterate over each possible group assignment
    for assignment in possible_assignments:
        # Check if the assignment has an equal number of 'a' and 'b'
        if assignment.count('a') == assignment.count('b'):
            # Create a new DataFrame to store the results
            result_df = df.copy()
            
            # Assign groups based on cage values
            group_assignments = dict(zip(unique_cages, assignment))
            result_df['group'] = result_df['cage'].map(group_assignments)
            result_df['combo'] = iteration_count
            iteration_count += 1
            
            # Add this DataFrame to the list
            combo_dfs.append(result_df)

    # Concatenate all resulting DataFrames
    all_combos_df = pd.concat(combo_dfs, ignore_index=True)

    #print(all_combos_df)

    #Save the updated DataFrame to a new sheet in the same Excel file
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        all_combos_df.to_excel(writer, sheet_name='AllCombos', index=False)


    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))

    ###################### All Combination Means ########################

    df = all_combos_df

    # Extract column names except for 'group', 'id' and 'combo'
    data_columns = [col for col in df.columns if col not in ['group', 'combo', 'id', 'cage']]

    # Create a new DataFrame to store mean values for each combination
    mean_values_df = pd.DataFrame(columns=['combo', 'group'] + data_columns)

    # Group by 'combo' and 'group' and calculate mean values for each 'ext_chunk' column
    for name, group in all_combos_df.groupby(['combo', 'group']):
        combo, group_name = name
        mean_values = group[data_columns].mean().to_frame().T
        mean_values['combo'] = combo
        mean_values['group'] = group_name

        # Concatenate 'id' values into a comma-separated string
        mean_values['id'] = ','.join(map(str, group['id'].tolist()))
        
        mean_values_df = pd.concat([mean_values_df, mean_values], ignore_index=True)

    # Reorder columns
    mean_values_df = mean_values_df[['combo', 'group', 'id'] + data_columns]


    # Display the resulting DataFrame
    #print(mean_values_df)

    # Save the updated DataFrame to a new sheet in the same Excel file
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        mean_values_df.to_excel(writer, sheet_name='AllMeans', index=False)

    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))

    ###################### Abs Diff from Calculated Mean Vals ######################

    # Read data from Excel file
    df = mean_values_df

    # Create a new dataframe for absolute differences
    abs_diff_df = pd.DataFrame()

    # Extract time point columns
    time_point_columns = [col for col in df.columns if col not in ['group', 'combo', 'id', 'cage']]

    # Iterate through time points
    for i in range(len(time_point_columns)):
        # Extract 'a' and 'b' values for the current time point
        a_values = df[df['group'] == 'a'][time_point_columns[i]].values
        b_values = df[df['group'] == 'b'][time_point_columns[i]].values

        # Extract time point number from column name
        time_point = int(time_point_columns[i][5:])

        # Calculate the absolute difference and create a new column in the new dataframe
        abs_diff_df[f'AbsDiff_{time_point}'] = abs(a_values - b_values)

    # Collect subjIDs and groups used for this time point
        subjIDs_a = df[df['group'] == 'a']['id'].values
        subjIDs_b = df[df['group'] == 'b']['id'].values

    # Create a list of subjIDs with corresponding groups
        subjIDsNgroups = [f"{group_a}: {subjIDs_a}, {group_b}: {subjID_b}" for group_a, subjIDs_a, group_b, subjID_b in zip(['a']*len(subjIDs_a), subjIDs_a, ['b']*len(subjIDs_b), subjIDs_b)]

    # Add the 'subjIDsNgroups' column to abs_diff_df
    abs_diff_df['Groups_subjID'] = subjIDsNgroups

    # Display the new dataframe
    #print("Absolute Mean Value Differences")
    #print(abs_diff_df)

    # Save the updated DataFrame to a new sheet in the same Excel file
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        abs_diff_df.to_excel(writer, sheet_name='MeanValsAbsDiff', index=False)

    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))

    ###################### Find Min Mean Value Grouping and list id's of this group ######################

    # Sum the values of each row
    row_sums = abs_diff_df.iloc[:, :-1].sum(axis=1)

    # Find the minimum sum value(s) and their index(es)
    min_sum = row_sums.min()
    min_indices = row_sums[row_sums == min_sum].index

    # Create a new DataFrame to hold the lowest sum values and their 'Groups_subjID' information
    min_sum_values_df = pd.DataFrame(columns=['MinSumValue', 'Groups_subjID'])

    # Populate the new DataFrame with the lowest sum values and their 'subjIDsNgroups' information
    for idx in min_indices:
        new_row = pd.DataFrame({
            'MinSumValue': [min_sum],
            'Groups_subjID': [abs_diff_df.loc[idx, 'Groups_subjID']]
            
        })
        min_sum_values_df = pd.concat([min_sum_values_df, new_row], ignore_index=True)

    # Save the updated DataFrame to a new sheet in the same Excel file
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        min_sum_values_df.to_excel(writer, sheet_name='MinMeansIDsGroups', index=False)

    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))

    ###################### Extract 'a', 'b' group assignments from list and create dictionary ######################
        
    # List of assigned groups
    string_list_element = [abs_diff_df.loc[idx, 'Groups_subjID']]
    # exp. string_list_element = ['a: 3,4,5,6,7,8,15,16, b: 1,2,9,10,11,12,13,14']
    print(string_list_element)

    # Join the list element into a single string
    string = ''.join(string_list_element)

    # Use regular expression to extract key-value pairs
    pairs = re.findall(r'(\w+):\s*([\d,]+)', string)

    # Initialize dictionary
    result_dict = {}

    # Iterate over pairs and populate dictionary
    for key, values_str in pairs:
        values = [int(value) for value in values_str.split(',') if value.strip()]
        result_dict[key] = values

    print(result_dict)

    prog_bar_update_val += 1
    # records progress by updating prog bar with each file compiled
    window["-Progress_BAR-"].update(max = max, current_count=int(prog_bar_update_val))

    ###################### Use groups assignment dictionary to create an new df with optimal groups ######################

    # OG DataFrame with 'id' column
    for_optimal_groups_df = pd.read_excel(file_name)

    # Dictionary of group assignments
    group_assignments = result_dict

    # Function to map group assignments
    def map_group(id):
        for group, ids in group_assignments.items():
            if id in ids:
                return group
        return None

    # Apply mapping function to 'id' column to create 'group' column
    for_optimal_groups_df['group'] = for_optimal_groups_df['id'].apply(map_group)

    # Display the new DataFrame
    print(for_optimal_groups_df)

    # Save the updated DataFrame to a new sheet in the same Excel file
    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        for_optimal_groups_df.to_excel(writer, sheet_name='OptimalAssignment', index=False)

    # last prog bar addition indicating the end of the program run
    window["-Progress_BAR-"].update(current_count=int(max))

    ###################### Use 'OptimalGroup' df to graph these groups ######################
        
    df = for_optimal_groups_df

    # Extract time point columns
    time_point_columns = [col for col in df.columns if col not in ['group', 'combo', 'id', 'cage']]

    # Calculate average values for 'a' and 'b' groups for each column data point
    a_avg = df[df['group'] == 'a'][time_point_columns].mean()
    b_avg = df[df['group'] == 'b'][time_point_columns].mean()

    # Plot the data
    plt.figure(figsize=(10, 6))

    # Plot 'a' average data
    plt.plot(time_point_columns, 
            a_avg, marker='o', label='a', color='blue')

    # Plot 'b' average data
    plt.plot(time_point_columns, 
            b_avg, marker='o', label='b', color='red')

    plt.title('Group Avg\na vs b')
    plt.xlabel('Min of Ext Learn Chunk')
    plt.ylabel('Pct FZn')
    plt.legend()
    # Rotate x-axis labels by 90 degrees
    plt.xticks(rotation=90,  fontsize=8)
    plt.show()

    # last prog bar addition indicating the end of the program run
    window["-Progress_BAR-"].update(current_count=int(prog_bar_update_val +1))

    # window telling the user the program functioned correctly
    nom_window(input_filename)   

# main GUI creation and GUI elements
sg.theme('DarkBlue9')

layout = [
    [sg.Text("Select the Excel file                    \n"
             "containing the behavioral data           \n" 
             "with cage mate pairs to be assigned groups."),
    sg.Input(key="-IN-"),
    sg.FileBrowse()],

    [sg.Text("Select a file to store the new Excel file.\n"
                "Data will be copied & transferred to this file. \n"),
    sg.Input(key="-OUT-"),
    sg.FolderBrowse()],

    [sg.Exit(), sg.Button("Press to assign subjects to least different groups"), 
    sg.Text("eBot's progress..."),
    sg.ProgressBar(20, orientation='horizontal', size=(15,10), 
                border_width=4, bar_color=("Blue", "Grey"),
                key="-Progress_BAR-")]
    
]

# create the window
window = sg.Window("Welcome to eBot's Least Diff Group Assignor!", layout)

# create an event loop
while True:
    event, values = window.read()
    # end program if user closes window
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    if event == "Press to assign subjects to least different groups":
        # check file selections are valid
        if (is_valid_path(values["-IN-"])) and (is_valid_path(values["-OUT-"])):

            extract_values_from_excel(
            input_filename  = values["-IN-"],
            output_folder = values["-OUT-"])   

window.close