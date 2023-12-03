import pandas as pd
import re

# Read the Excel file
df = pd.read_excel('your_excel_file.xlsx')

# Function to convert 'Duration' column to seconds
def convert_to_seconds(duration):
    total_seconds = 0
    pattern = r'(\d+)\s*(sec|s|min|m|hour|h)'
    matches = re.findall(pattern, duration)
    
    for match in matches:
        value, unit = match
        if unit in ['sec', 's']:
            total_seconds += int(value)
        elif unit in ['min', 'm']:
            total_seconds += int(value) * 60
        elif unit in ['hour', 'h']:
            total_seconds += int(value) * 3600

    return total_seconds

# Apply the conversion function to the duration column
df['Duration_Seconds'] = df['Duration'].apply(convert_to_seconds)

# Fill empty 'Name' or 'Type' with 'Number'
df['Name'].fillna(df['Number'], inplace=True)
df['Type'].fillna(df['Number'], inplace=True)

# Grouping by 'Name' and 'Call Type' to get call type counts for each person
grouped = df.groupby(['Name', 'Call Type']).size().unstack(fill_value=0)

# Create a DataFrame for the call statistics
call_stats = pd.DataFrame({
    'Total_Calls': df['Name'].groupby(df['Name']).count(),
    'Total_Duration_Seconds': df['Duration_Seconds'].groupby(df['Name']).sum(),
    'Total_Duration_Minutes': df['Duration_Seconds'].groupby(df['Name']).sum() // 60,
    'Total_Duration_Hours': df['Duration_Seconds'].groupby(df['Name']).sum() // 3600
})

# Calculate counts for each call type (Outgoing, Incoming, Missed)
call_type_counts = df.groupby(['Name', 'Call Type']).size().unstack(fill_value=0)
call_stats = call_stats.merge(call_type_counts, on='Name', how='outer')

# Calculate percentages for each call type
call_stats['Outgoing_Percentage'] = (call_stats['Outgoing'] / call_stats['Total_Calls']) * 100
call_stats['Incoming_Percentage'] = (call_stats['Incoming'] / call_stats['Total_Calls']) * 100
call_stats['Missed_Percentage'] = (call_stats['Missed'] / call_stats['Total_Calls']) * 100

# Reset the index to bring 'Name' back as a column
call_stats = call_stats.reset_index()

# Rearrange columns in the desired order
columns_order = ['Name', 'Total_Calls', 'Total_Duration_Seconds', 'Total_Duration_Minutes', 'Total_Duration_Hours',
                 'Incoming', 'Outgoing', 'Missed', 'Incoming_Percentage', 'Outgoing_Percentage', 'Missed_Percentage']

call_stats = call_stats[columns_order]

# Sort the data by 'Total_Calls' in descending order
call_stats = call_stats.sort_values(by='Total_Duration_Seconds', ascending=False)

# Save the sorted merged table with counts and percentages into an Excel file
call_stats.to_excel('merged_data_sorted_by_calls.xlsx', index=False)

print("Merged data sorted by number of calls and saved to 'merged_data_sorted_by_calls.xlsx' file.")
