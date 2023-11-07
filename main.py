import xpress as xp
import pandas as pd

# Load data from Excel
excel_path = 'InvestmentData.xlsx'  # Replace with your actual Excel file path
df = pd.read_excel(excel_path)

# Define the model
model = xp.problem()

# Initialize the decision variables for each investment option and year
investment_vars = {(row['Option'], year): xp.var(name=f'invest_{row["Option"]}_{year}', lb=0)
                   for index, row in df.iterrows()
                   for year in range(1, 6) if year >= row['StartYear']}

# Add variables to the model
model.addVariable(*investment_vars.values())

# Constraints
# 1. Initial budget constraint
initial_budget_constraint = xp.Sum(investment_vars[option, 1] for option in df['Option'] if (option, 1) in investment_vars) <= 1000000
model.addConstraint(initial_budget_constraint)

# 2. Reinvestment constraints (available funds in a year must be >= investments made in that year)
for year in range(2, 6):
    reinvestment_constraint = xp.Sum(investment_vars[option, year] for option in df['Option'] if (option, year) in investment_vars) <= \
                              xp.Sum(investment_vars[option, year - duration] * row['ReturnRate'] for index, row in df.iterrows() for option, duration in [(row['Option'], row['Duration'])] if (option, year - duration) in investment_vars)
    model.addConstraint(reinvestment_constraint)

# Objective function: Maximize the final value of the portfolio
# This will include the initial investments that have matured and the final year investments that are withdrawn without reinvestment
final_returns = xp.Sum(investment_vars[option, 5] for option in df['Option'] if (option, 5) in investment_vars)
mature_investments = xp.Sum(investment_vars[option, 5 - duration] * row['ReturnRate'] for index, row in df.iterrows() for option, duration in [(row['Option'], row['Duration'])] if (option, 5 - duration) in investment_vars)
model.setObjective(final_returns + mature_investments, sense=xp.maximize)

# Solve the model
model.solve()

# Print the results
for var in investment_vars:
    print(f"{investment_vars[var].name}: {model.getSolution(investment_vars[var])}")
    
print(f"Objective function value: {model.getObjVal()}") 