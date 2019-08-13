import os
from collections import OrderedDict
from typing import List, Generator, Any, Tuple

import xlrd
import xlsxwriter


def _iter_sheets(input_file: str, on_demand=True) -> Generator[
        xlrd.sheet.Sheet, None, None]:
    if not os.path.isfile(input_file):
        raise FileNotFoundError(input_file)
    book = xlrd.open_workbook(input_file, on_demand=on_demand)
    for sheet_name in book.sheet_names():
        sheet = book.sheet_by_name(sheet_name)
        yield sheet
        book.unload_sheet(sheet_name)
    book.release_resources()


def _get_cell_value(sheet: xlrd.sheet.Sheet, row: int, column: int) -> Any:
    cell_value = sheet.cell_value(row, column)
    if isinstance(cell_value, str) and not cell_value:
        cell_value = None
    return cell_value


def _iter_cell_values(sheet: xlrd.sheet.Sheet,
                      by_row: bool = True) -> Generator[Tuple[int, int, Any],
                                                        None, None]:
    for row in range(sheet.nrows if by_row else sheet.ncols):
        for column in range(sheet.ncols if by_row else sheet.nrows):
            yield row, column, _get_cell_value(sheet, row, column)


def _add_data_to_index(source: List, index: int, value: Any):
    if index == len(source):
        source.append(value)
    else:
        if source[index] is None:
            source[index] = value
        elif isinstance(source[index], float) and isinstance(value, float):
            source[index] += value
        elif value is not None:
            source[index] = value


def accumulate_sheets(input_file: str, output_file: str):
    data_matrix = []
    for sheet in _iter_sheets(input_file):
        for row, column, cell_value in _iter_cell_values(sheet):
            if row == len(data_matrix):
                data_matrix.append(list())

            cell_value = sheet.cell_value(row, column)
            _add_data_to_index(data_matrix[row], column, cell_value)
    output_matrix(output_file, data_matrix)


def accumulate_sheets_row_grouped(input_file: str, output_file: str,
                                  skip_rows=0):
    untouched_matrix = []
    grouped_rows = OrderedDict()
    empty_rows_after_key = OrderedDict()
    for sheet_index, sheet in enumerate(_iter_sheets(input_file)):
        for row in range(sheet.nrows):
            current_key = None
            left_intact = False
            for column in range(sheet.ncols):
                cell_value = _get_cell_value(sheet, row, column)

                if row < skip_rows:
                    # skipping rows, simply add it to the matrix overwriting
                    # previous values
                    if len(untouched_matrix) == row:
                        untouched_matrix.append(list())
                    _add_data_to_index(
                        untouched_matrix[row], column, cell_value)
                    left_intact = True
                    continue
                # First we need a key for this row
                if current_key is None and cell_value is not None:
                    current_key = cell_value
                    if current_key not in grouped_rows:
                        grouped_rows[current_key] = dict(
                            column=column, values=list())
                    continue
                elif current_key is not None:
                    offset_col = column - grouped_rows[current_key][
                        'column'] - 1
                    data = grouped_rows[current_key]['values']
                    _add_data_to_index(data, offset_col, cell_value)
            # If not left intact and no key available, the row was skipped
            # Only works for the first sheet
            if sheet_index == 0 and not left_intact and current_key is None:
                # Grab the last key
                keys = list(grouped_rows.keys())
                if not keys:
                    continue
                last_key = keys[-1]
                if last_key in empty_rows_after_key:
                    empty_rows_after_key[last_key] += 1
                else:
                    empty_rows_after_key[last_key] = 1

    # Append the untouched matrix with data from the grouped rows
    for key, data in grouped_rows.items():
        row_data = [None for _ in range(data['column'])]
        row_data.append(key)
        row_data.extend(data['values'])
        untouched_matrix.append(row_data)
        # Add the empty rows
        for _ in range(empty_rows_after_key.get(key, 0)):
            untouched_matrix.append([])
    output_matrix(output_file, untouched_matrix)


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
    argument_parser.add_argument('--group-by-row', action='store_true')
    argument_parser.add_argument(
        '--skip-initial-lines', type=int, default=0,
        help='Only available when grouping by rows')
    parsed = argument_parser.parse_args()

    if not os.path.isfile(parsed.input_file):
        print('Input file does not exist')
        sys.exit(1)
    file_arguments = (parsed.input_file, parsed.output_file)
    if parsed.group_by_row:
        accumulate_sheets_row_grouped(
            *file_arguments, skip_rows=parsed.skip_initial_lines)
    else:
        accumulate_sheets(*file_arguments)
