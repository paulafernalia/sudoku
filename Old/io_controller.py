import pandas as pd
import numpy as np


def read_constraints(
    sht, 
    cols,
    rows,
    index_):

    # Read initial values from excel file and store in nummpy array
    constr_matrix = np.zeros([9,9], dtype=int)
    for c_i in index_:
        for r_i in index_:
            # Ignore cells with no values
            this_cell = sht.range(cols[c_i] + str(rows[r_i])).value
            if this_cell is not None:
                constr_matrix[r_i, c_i] = this_cell  

    return constr_matrix


def write_solution(
    sht,
    cols,
    rows,
    index_,
    constr_matrix,
    sol_matrix):

    for c_i in index_:
        for r_i in index_: 
            if constr_matrix[r_i, c_i] != 0:
                sht.range(cols[c_i] + str(rows[r_i])).value = constr_matrix[r_i, c_i]
                #sht.range(cols[c_i] + str(rows[r_i])).api.Font.Bold = True

    for c_i in index_:
        for r_i in index_:
            sht.range(cols[c_i] + str(rows[r_i])).value = sol_matrix[r_i, c_i]