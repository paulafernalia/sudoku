# Sudoku Solver


Algorithm to solve a sodoku using linear programming.

The program reads the initial constraints from an excel file using the python library xlwings and it soves the sudoku with branch and bound using Pyomo and the open source solver glpk.

When more than one possible solution exists, it will return the first solution found. It includes handling of infeasible solutions.

The solution is written back to the excel file.
