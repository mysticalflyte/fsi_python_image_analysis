# Python 3.6.5 script for post-data processing
# Specifically eliminates common particles that may be stuck between different
# test runs and eliminating them from the second run's data.
# Written by Jenny Liu

# Imports:
import sys
import numpy as np
import math
from itertools import compress
import xlsxwriter as xls
import pandas as pd
from determineParticleSizes import write_xl_summaries
import pathlib


# Global constants - FEEL FREE TO MODIFY
ORIGINAL_XL_FILENAME = str(pathlib.Path("../Test Images/First Sample Images/img_results/results_test.xlsx"))
NEW_XL_FILENAME = str(pathlib.Path("../Test Images/First Sample Images/img_results/results_filtered.xlsx"))
POSITIONAL_REMOVAL_RANGE = 0
AREA_REMOVAL_RANGE = 0
MAX_PERCENT_REMOVED = 5


def main():
    xl_file = pd.ExcelFile(ORIGINAL_XL_FILENAME)
    df_data = pd.read_excel(xl_file, 'data', header = 0)
    # print(df_data)

    idx_to_delete = cmp_particles_in_imgs(df_data)
    idx_to_delete = set(idx_to_delete)
    # print(idx_to_delete)

    updated_df_data = df_data[~df_data.index.isin(idx_to_delete)]
    # print(updated_df_data)

    # Write out updated dataframe to a new excel file
    write_new_file(df_data, updated_df_data)

    inform_dirty_state(df_data, idx_to_delete)



# Creates and returns a list of indexes corresponding to rows in the excel sheet
# that should be eliminated based on criteria (determined by the global ranges)
def cmp_particles_in_imgs(df_data):
    idx_to_delete = []
    numRows, numCols = df_data.shape

    # iterate through each individual particle to compare against others
    for current_idx in range(numRows):
        current_img = df_data.iloc[current_idx, 0]
        # print("CURRENT IMAGE TO COMPARE AGAINST: " + current_img)

        for cmp_idx in range(current_idx + 1, numRows):
            # Skip particles associated with this same image
            cmp_img = df_data.iloc[cmp_idx, 0]
            # print("cmp_img: " + cmp_img)

            # If there are any repeats/similarities based on the global ranges,
            # add them to the list to be removed
            if (is_similar_particle(df_data, current_idx, cmp_idx)):
                idx_to_delete.append(cmp_idx)

    return idx_to_delete


# Determine whether the particle being compared against the current particle
# is the same (in terms of location or area)
def is_similar_particle(df_data, current_idx, cmp_idx):
    x_col_num = 12
    y_col_num = 13
    area_col_num = 3

    # current_row = df_data.iloc[[current_idx]]
    x_min = df_data.iloc[current_idx, x_col_num] - POSITIONAL_REMOVAL_RANGE

    x_max = x_min + 2 * POSITIONAL_REMOVAL_RANGE
    y_min = df_data.iloc[current_idx, y_col_num] - POSITIONAL_REMOVAL_RANGE
    y_max = y_min + 2 * POSITIONAL_REMOVAL_RANGE
    area_min = df_data.iloc[current_idx, area_col_num] - AREA_REMOVAL_RANGE
    area_max = area_min + 2 * AREA_REMOVAL_RANGE

    x_cmp = df_data.iloc[cmp_idx, x_col_num]
    y_cmp = df_data.iloc[cmp_idx, y_col_num]
    area_cmp = df_data.iloc[cmp_idx, area_col_num]

    x_similar = x_cmp >= x_min and x_cmp <= x_max
    y_similar = y_cmp >= y_min and y_cmp <= y_max
    area_similar = area_cmp >= area_min and area_cmp <= area_max

    return area_similar and (x_similar and y_similar)


# Write out updated filtered data to a new excel file, keeping the old data
# in a 3rd sheet
def write_new_file(df_data, updated_df_data):
    # Rewrite old data and put in new filtered data
    writer = pd.ExcelWriter(NEW_XL_FILENAME, engine = "xlsxwriter")
    updated_df_data.to_excel(writer, sheet_name = "filtered_data", index = False)
    df_data.to_excel(writer, sheet_name = "original_data", index = False)

    wb = writer.book

    # Set column widths for data sheets
    for worksheet in wb.worksheets():
        worksheet.set_column(0, 13, 15)

    # Make summary sheet, set width and write out summary excel functions
    summary_sheet = wb.add_worksheet("summary")
    summary_sheet.set_column(0, 1, 25)
    filtered_sheet = wb.get_worksheet_by_name("filtered_data")
    numRows, numCols = updated_df_data.shape
    write_xl_summaries(numRows, wb, filtered_sheet, summary_sheet)

    writer.save()


# If the percentage of removed particles reaches MAX_PERCENT_REMOVED, notify the
# user that they should clean the flowcell.
def inform_dirty_state(df_data, idx_to_delete):
    numRows, numCols = df_data.shape
    percent_removed = len(idx_to_delete) / numRows * 100

    if (percent_removed >= MAX_PERCENT_REMOVED):
        print("The flowcell is dirty and " + str(percent_removed) + "% of particles are repeats.")
        print("Please consider cleaning and starting anew.")


# Run the main program
if __name__ == "__main__":
    main()
