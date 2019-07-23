import os
from typing import List

import xlrd
import xlsxwriter


def accumulate_sheets(input_file: str, output_file: str):
    if not os.path.isfile(input_file):
        raise FileNotFoundError(input_file)
    book = xlrd.open_workbook(input_file, on_demand=True)
    data_matrix = []
    for sheet_name in book.sheet_names():
        sheet = book.sheet_by_name(sheet_name)
        for row in range(sheet.nrows):
            for column in range(sheet.ncols):
                if row == len(data_matrix):
                    data_matrix.append(list())

                cell_value = sheet.cell_value(row, column)
                if column == len(data_matrix[row]):
                    data_matrix[row].append(cell_value)
                else:
                    if (isinstance(data_matrix[row][column], float)
                            and isinstance(cell_value, float)):
                        data_matrix[row][column] += cell_value
                    elif cell_value:
                        data_matrix[row][column] = cell_value
        book.unload_sheet(sheet_name)
    book.release_resources()
    output_matrix(output_file, data_matrix)


def output_matrix(output_file: str, data_matrix: List[List]):
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet()

    for row, row_data in enumerate(data_matrix):
        for column, cell_value in enumerate(row_data):
            if cell_value is not None:
                worksheet.write(row, column, cell_value)
    workbook.close()


if __name__ == '__main__':
    import argparse
    import sys

    argument_parser = argparse.ArgumentParser()
    argument_parser.add_argument('input_file')
    argument_parser.add_argument('output_file')
    parsed = argument_parser.parse_args()

    if not os.path.isfile(parsed.input_file):
        print('Input file does not exist')
        sys.exit(1)
    accumulate_sheets(parsed.input_file, parsed.output_file)
