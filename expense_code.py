import pandas as pd

# Read the Excel file
df = pd.read_excel('expenses.xlsx')

# Convert the date column to a datetime data type
df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y')

# Define the range of months and years you're interested in
start_month = 1
start_year = 2023
end_month = 5
end_year = 2023

# Filter the expenses for the specified range of months and years
filtered_expenses = df[
    (df['Date'].dt.month >= start_month) & (df['Date'].dt.year >= start_year) &
    (df['Date'].dt.month <= end_month) & (df['Date'].dt.year <= end_year)
]

# Sum the expenses for the specified range of months
total_expenses = filtered_expenses['Expense'].sum()

# Find all sent e-transfers in the specified range of months
sent_transfers = filtered_expenses[filtered_expenses['Transaction Type'].str.startswith('SEND E-TFR')]

# Sum the sent e-transfers for the specified range of months
total_sent_transfers = sent_transfers['Expense'].sum()

# Create a DataFrame with the totals
totals_df = pd.DataFrame({
    'Transaction Type': ['Total expenses', 'Total sent e-transfers (without rent)'],
    'Expense': [total_expenses, total_sent_transfers-1000]
})

# Append the totals DataFrame to the original DataFrame
df = pd.concat([df, totals_df], ignore_index=True)

# Write the updated DataFrame back to the Excel file
df.to_excel('expenses.xlsx', index=False)

# Output the results with 2 decimal places
print(f"Total expenses from {start_month}/{start_year} to {end_month}/{end_year}: ${total_expenses:.2f}")
print(f"Total sent e-transfers (without rent) from {start_month}/{start_year} to {end_month}/{end_year}: ${total_sent_transfers-1000:.2f}")
