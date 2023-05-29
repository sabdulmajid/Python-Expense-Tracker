import pandas as pd

# Read the Excel file
df = pd.read_excel('expenses.xlsx')

# Convert the date column to a datetime data type
df['Date'] = pd.to_datetime(df['Date'], format='%m/%d/%Y')

# Define the month and year you're interested in
target_month = 1
target_year = 2023

# Filter the expenses for the specific month and year
filtered_expenses = df[(df['Date'].dt.month == target_month) & (df['Date'].dt.year == target_year)]

# Sum the expenses for the specific month
total_expenses = filtered_expenses['Expense'].sum()

# Find all sent e-transfers in the specific month
sent_transfers = filtered_expenses[filtered_expenses['Transaction Type'].str.startswith('SEND E-TFR')]

# Sum the sent e-transfers for the specific month
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
print(f"Total expenses for {target_month}/{target_year}: ${total_expenses:.2f}")
print(f"Total sent e-transfers (without rent) for {target_month}/{target_year}: ${total_sent_transfers-1000:.2f}")
