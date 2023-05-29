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

filtered_expenses = df[
    (df['Date'].dt.month >= start_month) & (df['Date'].dt.year >= start_year) &
    (df['Date'].dt.month <= end_month) & (df['Date'].dt.year <= end_year)
]

total_expenses = filtered_expenses[filtered_expenses['Expense'] != 1000]['Expense'].sum()

sent_transfers = filtered_expenses[
    (filtered_expenses['Transaction Type'].str.startswith('SEND E-TFR')) &
    (filtered_expenses['Expense'] != 1000)
]

total_sent_transfers = sent_transfers['Expense'].sum()

groceries_expenses = filtered_expenses[
    filtered_expenses['Transaction Type'].str.contains('FRESHCO|INSTACART', case=False)
]['Expense'].sum()

fast_food_expenses = filtered_expenses[
    filtered_expenses['Transaction Type'].str.contains('DD|TAHINI|DOMINO|KABOB SHACK|MATH COFFEE|TIM HORTONS|UBER\* EATS|MAC SUSHI|SHIN WA|HARVEY|SWEET DREAMS|GONG CHA', case=False)
]

fast_food_total = 0
for _, row in fast_food_expenses.iterrows():
    expense = row['Expense']
    if expense > 30:
        expense = expense / 2
    fast_food_total += expense

totals_df = pd.DataFrame({
    'Transaction Type': ['Total expenses', 'Total sent e-transfers (without rent)', 'Total FreshCo & InstaCart', 'Total Outside Fast Food'],
    'Expense': [total_expenses, total_sent_transfers, groceries_expenses, fast_food_total]
})

df = pd.concat([df, totals_df], ignore_index=True)

df.to_excel('expenses.xlsx', index=False)

print(f"Total expenses from {start_month}/{start_year} to {end_month}/{end_year}: ${total_expenses:.2f}")
print(f"Total sent e-transfers (without rent) from {start_month}/{start_year} to {end_month}/{end_year}: ${total_sent_transfers:.2f}")
print(f"Total FreshCo & InstaCart from {start_month}/{start_year} to {end_month}/{end_year}: ${groceries_expenses:.2f}")
print(f"Total Outside Fast Food from {start_month}/{start_year} to {end_month}/{end_year}: ${fast_food_total:.2f}")
