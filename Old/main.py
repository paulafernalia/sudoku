import io_controller
import opt_model
import xlwings as xw
import os

# TODO: Handle infeasible solutions

# Basic parameters

# Select file and sheet
path = os.path.join(os.getcwd(), 'sudoku_app.xlsm')
wb = xw.Book(path) # for running from Python
# wb = xw.Book.caller() # for Excel macro

# Select always first tab
sht = wb.sheets[0]

# Create ranges of columns and rows
cols_in = list(map(chr, range(66,75)))
cols_out = list(map(chr, range(76,85)))
rows = list(range(3,12))
index_ = list(range(0,9))

# Read hard constraints from excel file
constr_matrix = io_controller.read_constraints(
    sht=sht,
    cols=cols_in,
    rows=rows,
    index_=index_)

# Solve sudoku using constraints from excel
sol_matrix = opt_model.solve_sudoku(constraint_matrix=constr_matrix)

# Write solution to excel
io_controller.write_solution(
    sht=sht,
    cols=cols_out,
    rows=rows,
    index_=index_,
    constr_matrix=constr_matrix,
    sol_matrix=sol_matrix)
