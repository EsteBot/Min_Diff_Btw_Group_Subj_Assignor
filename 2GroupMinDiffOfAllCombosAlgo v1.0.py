import pandas as pd
from itertools import combinations, permutations, product
import re
import matplotlib.pyplot as plt


file_name = "F:\ExtChunkAllDats2.xlsx"
# Read the Excel file (assuming it's named 'data.xlsx')
df = pd.read_excel(file_name)

# Display the initial DataFrame
#print("Initial DataFrame:")
#print(df)


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

###################### Abs Diff from Calculated Mean Vals ######################

# Read data from Excel file
df = mean_values_df

# Create a new dataframe for absolute differences
abs_diff_df = pd.DataFrame()

# Extract time point columns
time_point_columns = [col for col in df.columns if col not in ['group', 'combo', 'id', 'cage']]

# Iterate through time points
for i in range(len(time_point_columns)):
    # Extract 'cn' and 'tx' values for the current time point
    a_values = df[df['group'] == 'a'][time_point_columns[i]].values
    b_values = df[df['group'] == 'b'][time_point_columns[i]].values

    # Extract time point number from column name
    time_point = int(time_point_columns[i][10:])

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

###################### Extract 'a', 'b' group assignments from list and create dictionary ######################
    
# List of assigned groups
string_list_element = [abs_diff_df.loc[idx, 'Groups_subjID']]
#string_list_element = ['a: 3,4,5,6,7,8,15,16, b: 1,2,9,10,11,12,13,14']
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