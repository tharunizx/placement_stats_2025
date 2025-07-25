import pandas as pd
from tabulate import tabulate
import numpy as np


# Load Excel file
file_path = "placements.xlsx"
df = pd.read_excel(file_path)

# Clean column names but preserve '#' for serial number
df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")
df = df.rename(columns={'#': 'serial'})

# Clean the CTC column by removing non-numeric parts
df['ctc'] = df['ctc'].astype(str).str.replace('[^0-9.]', '', regex=True)
df['ctc'] = pd.to_numeric(df['ctc'], errors='coerce')

# Group by company and branch
company_branch_counts = df.groupby(['company', 'branch']).size().unstack(fill_value=0)

# Add Total per company
company_branch_counts['Total'] = company_branch_counts.sum(axis=1)

# Reorder columns: Total first
cols = ['Total'] + [col for col in company_branch_counts.columns if col != 'Total']
company_branch_counts = company_branch_counts[cols]

# Sort companies by total placed
company_branch_counts = company_branch_counts.sort_values(by='Total', ascending=False)

# Add final row with total placements from each branch (i.e., column-wise sum)
branch_totals = company_branch_counts.sum(numeric_only=True)
branch_totals.name = 'TOTAL'
company_branch_counts = pd.concat([company_branch_counts, pd.DataFrame([branch_totals])])

# Ensure CTC column is clean
df['ctc'] = df['ctc'].astype(str).str.replace('[^0-9.]', '', regex=True)
df['ctc'] = pd.to_numeric(df['ctc'], errors='coerce')

# # Extract students with multiple offers (i.e., same roll_no appears more than once)
# multi_offer_df = df[df.duplicated('roll_no', keep=False)]

# # For tabulate display
# students_with_multiple_offers = []

# for roll_no, group in multi_offer_df.groupby('roll_no'):
#     name = group.iloc[0]['student_name']
#     offers = [f"{row['company']} ({row['ctc']} LPA)" for _, row in group.iterrows()]
#     students_with_multiple_offers.append((name, roll_no, offers))

# # Convert for tabulate
# from tabulate import tabulate
# table_data = []
# for name, roll, offers in students_with_multiple_offers:
#     offer_str = " | ".join(offers)
#     table_data.append([name, roll, offer_str])

# # Display
# print("\n--- Students with Multiple Offers ---\n")
# print(tabulate(table_data, headers=["Name", "Roll No", "Offers"], tablefmt="grid"))


# # Drop duplicate student entries, keeping the one with the highest CTC
# # Assumes 'roll_no' or 'student_name' uniquely identifies a student

# if 'roll_no' in df.columns:
#     df = df.sort_values(by='ctc', ascending=False).drop_duplicates(subset='roll_no', keep='first')
# else:
#     df = df.sort_values(by='ctc', ascending=False).drop_duplicates(subset='student_name', keep='first')



# -------------------------------

# Print the full result
print("\n--- Company-wise Placement by Branch (with Branch Totals) ---")
print(company_branch_counts.to_string())

# General stats
total_students = len(df)
total_companies = df['company'].nunique()
average_ctc = round(df['ctc'].mean(), 2)
highest_ctc = df['ctc'].max()
lowest_ctc = df['ctc'].min()
median_ctc = df['ctc'].median()
std_ctc = round(df['ctc'].std(), 2)

# Company-wise stats
company_stats = df.groupby('company').agg(
    num_students=('roll_no', 'count'),
    avg_ctc=('ctc', 'mean')
).sort_values(by='num_students', ascending=False).head(5)


# CTC distribution
bins = [0, 5, 7, 10, np.inf]
labels = ['3–5', '5–7', '7–10', '10+']
df['ctc_range'] = pd.cut(df['ctc'], bins=bins, labels=labels, right=False)
ctc_distribution = df['ctc_range'].value_counts().sort_index()

# Top 5 highest paid students
top_students = df.sort_values(by='ctc', ascending=False).head(5)[['student_name', 'company', 'designation', 'ctc']]

print("\n--- Overall Gender Ratio ---")
gender_counts = df['gender'].value_counts()
gender_percent = df['gender'].value_counts(normalize=True) * 100
overall_gender_ratio = pd.DataFrame({
    'count': gender_counts,
    'percent': gender_percent.round(2)
})
print(overall_gender_ratio.to_string())

print("\n--- Company-wise Gender Ratio (with Percentage) ---")

# Group by company and gender, get counts
company_gender_ratio = df.groupby(['company', 'gender']).size().unstack(fill_value=0)

# Calculate total placed per company
company_gender_ratio['Total'] = company_gender_ratio.sum(axis=1)

# Calculate percentage columns
if 'Male' in company_gender_ratio.columns:
    company_gender_ratio['Male %'] = (company_gender_ratio['Male'] / company_gender_ratio['Total'] * 100).round(2)
else:
    company_gender_ratio['Male %'] = 0.0

if 'Female' in company_gender_ratio.columns:
    company_gender_ratio['Female %'] = (company_gender_ratio['Female'] / company_gender_ratio['Total'] * 100).round(2)
else:
    company_gender_ratio['Female %'] = 0.0

company_gender_ratio = company_gender_ratio.sort_values(by='Male %', ascending=False)


print(company_gender_ratio.to_string())

# Output
print("\n--- Placement Statistics ---")
print(f"Total Students Placed: {total_students}")
print(f"Total Companies Visited: {total_companies}")
print(f"Average CTC: ₹{average_ctc} LPA")
print(f"Highest CTC: ₹{highest_ctc} LPA")
print(f"Lowest CTC: ₹{lowest_ctc} LPA")
print(f"Median CTC: ₹{median_ctc} LPA")
print(f"Standard Deviation of CTC: ₹{std_ctc} LPA")

print("\n--- Top 5 Companies by Placement Count ---")
print(company_stats)

print("\n--- CTC Distribution ---")
print(ctc_distribution)

print("\n--- Top 5 Highest Paid Students ---")
print(top_students.to_string(index=False))
