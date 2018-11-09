"""A program that analyzes and solves a power flow."""

import argparse
import numpy
import openpyxl

# Program constants.
DEFAULT_INPUT_WORKBOOK = 'data/Data.xlsx'
DEFAULT_LINE_DATA_WORKSHEET = 'Line data'


def parse_arguments():
    """Parses command line arguments.

    Returns:
        An object containing program arguments.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument('--input_workbook', default=DEFAULT_INPUT_WORKBOOK,
                        help='An Excel workbook containing input data.')
    parser.add_argument('--line_data_worksheet', default=DEFAULT_LINE_DATA_WORKSHEET,
                        help='The worksheet containing line data.')
    return parser.parse_args()


def count_rows(ws):
    """Counts the number of non-empty, contiguous rows in the worksheet.

    This method is necessary because openpyxl.worksheet.worksheet.max_row does not always correctly detect empty cells.

    Args:
        ws: The worksheet to read.

    Returns:
        The number of non-empty, contiguous rows.
    """
    n = 0
    for row in ws.iter_rows():
        if not row[0].value:
            break

        n += 1

    return n


def create_admittance_matrix(ws):
    """Creates an admittance matrix from an input worksheet.

    Each row is expected to contain the following data:
        - column 1: source bus
        - column 2: destination bus
        - column 3: R total (pu)
        - column 4: X total (pu)
        - column 5: B total (pu)
        - column 6: F max (MVA)

    Furthermore, it is expected that the destination bus will always be greater than the source bus. This constraint
    ensures that line data is not duplicated.

    Args:
        ws: The worksheet containing line data.

    Returns:
        A square matrix containing line-to-line admittances.
    """
    max_row = count_rows(ws) - 1  # skip the header row
    admittances = numpy.zeros((max_row, max_row)) * 1j  # multiply by 1j to make values complex

    # Compute matrix entries.
    for row in ws.iter_rows(row_offset=1, max_row=max_row):
        # Subtract 1 since arrays start at 0, but buses start at 1.
        src = row[0].value - 1
        dst = row[1].value - 1

        # Series and parallel impedances.
        rtotal = row[2].value
        xtotal = row[3].value
        btotal = row[4].value
        # fmax = row[5].value

        y_series = 1 / (rtotal + 1j * xtotal)
        y_parallel = 1j * btotal / 2

        if src != dst:
            # Subtract series admittances from off-diagonal entries.
            admittances[src][dst] -= y_series
            admittances[dst][src] -= y_series

        # Add series and parallel admittances to diagonal entries.
        admittances[src][src] += y_series + y_parallel
        admittances[dst][dst] += y_series + y_parallel

    return admittances


def main():
    """Reads an input file containing system data and initiates power flow computations."""
    args = parse_arguments()

    # Read input data.
    wb = openpyxl.load_workbook(args.input_workbook, read_only=True)
    line_data_ws = wb[args.line_data_worksheet]

    # Create admittance matrix.
    print(create_admittance_matrix(line_data_ws))

    # Close input data.
    wb.close()


if __name__ == '__main__':
    main()
