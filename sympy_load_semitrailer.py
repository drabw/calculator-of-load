from sympy import symbols, Eq, solve
import pandas as pd
from openpyxl import Workbook
import pickle
import matplotlib.pyplot as plt
import os

# definition path for current catalog of script
catalog_path = os.path.dirname(os.path.abspath(__file__))

# definition path for data in same catalog
data_path = os.path.join(catalog_path, 'dictionary_only_semi.pkl')

# Load dictionary from file
with open(data_path, 'rb') as plik:
    dictionary_only_semi = pickle.load(plik)

# symbols for constant
Rd0, Rc0,  RdQ, RcQ, Q, Rc, Rd, Ql = \
    symbols('Rd0 Rc0 RdQ RcQ Q  Rc Rd Ql')

# def value of constant from dictionary
Qs = int(dictionary_only_semi["Qs"])
Xs = int(dictionary_only_semi["Xs"])
Qr = int(dictionary_only_semi["Qr"])
Xr = int(dictionary_only_semi["Xr"])
Qz = int(dictionary_only_semi["Qz"])
Xz = int(dictionary_only_semi["Xz"])
dmc = int(dictionary_only_semi["dmc"])
Rd_max = int(dictionary_only_semi["Rd_max"])
Rc_max = int(dictionary_only_semi["Rc_max"])
Xl_start, Xl_end, Xl_step = 100, 10500, 100  # Adjusted Xl parameters

data_xl = []
data_11500 = []
data_24000 = []

# def eq without load

eq1 = Eq(Qs * Xs + Qr * Xr + Qz * Xz - Rd0 * Xz, 0)
eq2 = Eq(Rc0 + Rd0 - Qs - Qr - Qz, 0)

# def eq with load
eq6 = Eq(RcQ - Ql + RdQ, 0)
eq11 = Eq(Rc0 + RcQ, Rc)
eq12 = Eq(Rd0 + RdQ, Rd)

# def condition
eq15 = Eq(Rc, Rc_max)
eq16 = Eq(Rd, Rd_max)

# Iterate through different Xl values for Rb = 11.500
for Xl_val in range(Xl_start, Xl_end + Xl_step, Xl_step):

    # condition no divide by zero
    div_0 = Xl_val - Xz
    # Add Xl to the list of equations
    if div_0 == 0:
        # set value for divide by zero
        eq5 = Eq(Ql, dmc)
    else:
        # set value for no divide by zero
        eq5 = Eq(Ql * Xl_val - RdQ * Xz, 0)

    # Combine all equations
    equations = [eq1, eq2, eq5, eq6, eq11, eq12, eq15]

    # Solve the system of equations
    solutions = solve(equations, (Rd0, Rc0, Rc, Rd, RcQ, RdQ, Ql, Q))

    # Round the solutions
    rounded_solutions = [(var, round(val, 2)) for var, val in solutions.items()]

    # loop by var
    for var, val in rounded_solutions:
        # if val == Q
        if var == Ql:
            # match new variable to val
            row_result = [val]
            # add Ql to list
            data_11500.append(row_result)
            # add position to list
            data_xl.append(Xl_val)

# Iterate through different Xl values for Rd = 24000
for Xl_val in range(Xl_start, Xl_end + Xl_step, Xl_step):

    # Add Xl to the list of equations
    eq5 = Eq(Ql * Xl_val - RdQ * Xz, 0)

    # Combine all equations
    equations = [eq1, eq2, eq5, eq6, eq11, eq12, eq16]

    # Solve the system of equations
    solutions = solve(equations, (Rd0, Rc0, Rc, Rd,
                                  RcQ, RdQ, Ql, Q))

    # Round the solutions
    rounded_solutions = [(var, round(val, 2)) for var, val in solutions.items()]

    # loop by var
    for var, val in rounded_solutions:
        # if val == Q
        if var == Ql:
            # match new variable to val
            row_result = [val]
            # add Ql to list
            data_24000.append(row_result)

# Create df from list
data_xl = pd.DataFrame(data_xl)
df_11500 = pd.DataFrame(data_11500)
df_24000 = pd.DataFrame(data_24000)

# save df in excel
wb = Workbook()
ws = wb.active
df = pd.concat([data_xl, df_11500, df_24000], axis=1)
df.columns = ['location from king pin [mm]',
              f'Ql max  {Rc_max} t for kingpin', f'Ql max  {Rd_max} t for axes']

# check the minimum Ql
columns_range = [f'Ql max  {Rc_max} t for kingpin',
                 f'Ql max  {Rd_max} t for axes']

# selected columns in df
selected_columns_df = df[columns_range]

# created empty list of min value
min_values_list = []

#  loop for row
for index, row in selected_columns_df.iterrows():
    # find min value
    min_value = row[row >= 0].min()

    # append to list min_value
    min_values_list.append(min_value)

# create df min
df_min = pd.DataFrame(min_values_list)

# update df (add df_min)
df = pd.concat([data_xl, df_11500, df_24000, df_min], axis=1)
df.columns = ['location from king pin [mm]',
              f'Ql max  {Rc_max} t for kingpin',
              f'Ql max  {Rd_max} t for axes', 'Qlmax']


# col for change value
target_column = 'Qlmax'

# calculate Qmax
Q_max = dmc - Qs - Qr - Qz

# function of lambda for replace max value
replace_func = lambda x: Q_max if x > Q_max else x

# use function for selected col
df[target_column] = df[target_column].apply(replace_func)

# option when user has not excel
try:
    import xlwings as xw

    # crate new one excel --> name book
    book = xw.Book()

    sheet1 = book.sheets[0]

    # put df to cell A6 without header and index
    sheet1["A1"].options(header=True, index=False).value = df

except ImportError:
    # none becouse I can use now xw in another part of code
    xw = None  # Definiowanie xw w bloku except
    # if user has not excel
    # save data in openpyxl
    df.to_excel('nazwa_pliku.xlsx', index=False)
    print("Uwaga: xlwings nie jest dostępne. Dane zostały zapisane do pliku "
          "Excela za pomocą openpyxl.")

# create plot
plt.plot(df['location from king pin [mm]'], df['Qlmax'])
plt.xlabel('X')
plt.ylabel('Y')
plt.title('Depend of max load of location')

# show plot
plt.show()
