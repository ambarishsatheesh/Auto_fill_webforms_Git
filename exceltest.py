# Code to test if reading from excel for mass data entry program works properly
import pandas as pd

# Open Excel and read cells
df = pd.read_excel('file_name.xlsx', sheet_name='sheet_name', usecols="A:G")

# Input column data into lists
email = df['Email Address'].tolist()
f_name = df['First Name'].tolist()
l_name = df['Last Name'].tolist()
c_name = df['Company Name'].tolist()
t_no = df['Telephone No'].tolist()
county = df['County'].tolist()
town = df['Town'].tolist()

# Loop through lists and print results
for i in range(371,378):
    print(email[i])
    print(f_name[i])
    print(l_name[i])
