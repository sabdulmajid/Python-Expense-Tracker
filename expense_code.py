import pandas as pd

# Read the Excel file
df = pd.read_excel('expenses.xlsx')

# Convert the date column to a datetime data type
df['Date'] = pd.to_datetime(df['Date'], format='%Y-%m-%d')

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

# Exclude rent from the total expenses
total_expenses = filtered_expenses[filtered_expenses['Expense'] != 1000]['Expense'].sum()

# Find all sent e-transfers in the specified range of months, excluding rent
sent_transfers = filtered_expenses[
    (filtered_expenses['Transaction Type'].str.startswith('SEND E-TFR')) &
    (filtered_expenses['Expense'] != 1000)
]

# Sum the sent e-transfers for the specified range of months
total_sent_transfers = sent_transfers['Expense'].sum()

# Find all FreshCo and InstaCart entries and calculate the total as groceries expenses
groceries_expenses = filtered_expenses[
    filtered_expenses['Transaction Type'].str.contains('FRESHCO|INSTACART', case=False)
]['Expense'].sum()

# Find all outside fast food purchases and calculate the total
fast_food_expenses = filtered_expenses[
    filtered_expenses['Transaction Type'].str.contains('DD|TAHINI|DOMINO|KABOB SHACK|MATH COFFEE|TIM HORTONS|UBER\* EATS|MAC SUSHI|SHIN WA|HARVEY|SWEET DREAMS|GONG CHA', case=False)
]

fast_food_total = 0
for _, row in fast_food_expenses.iterrows():
    try:
        expense = row['Expense']
        if expense > 30:
            expense = expense / 2
        fast_food_total += expense
    except TypeError:
        pass

# Create a DataFrame with the totals
totals_df = pd.DataFrame({
    'Transaction Type': ['Total expenses', 'Total sent e-transfers (without rent)', 'Total FreshCo & InstaCart', 'Total Outside Fast Food'],
    'Expense': [total_expenses, total_sent_transfers, groceries_expenses, fast_food_total]
})

# Append the totals DataFrame to the original DataFrame
df = pd.concat([df, totals_df], ignore_index=True)

# Write the updated DataFrame back to the Excel file
df.to_excel('expenses.xlsx', index=False)

# Output the results with 2 decimal places
print(f"Total expenses from {start_month}/{start_year} to {end_month}/{end_year}: ${total_expenses:.2f}")
print(f"Total sent e-transfers (without rent) from {start_month}/{start_year} to {end_month}/{end_year}: ${total_sent_transfers:.2f}")
print(f"Total FreshCo & InstaCart (Ayman's Part) from {start_month}/{start_year} to {end_month}/{end_year}: ${groceries_expenses:.2f}")
print(f"Total Outside Fast Food from {start_month}/{start_year} to {end_month}/{end_year}: ${fast_food_total:.2f}")
