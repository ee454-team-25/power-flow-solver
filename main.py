"""A program that analyzes and solves a power flow."""

import argparse
import numpy
import openpyxl


def parse_arguments():
    """Parses command line arguments.

    Returns:
        An object containing program arguments.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument('--input_workbook', default='Data.xlsx', help='An Excel workbook containing input data.')
    parser.add_argument('--line_data_worksheet', default='Line data', help='The worksheet containing line data.')
    parser.add_argument('--num_buses', default=14, help='The number of buses in the system.')
    return parser.parse_args()


def create_admittance_matrix(ws, num_buses):
    """Creates an admittance matrix from an input worksheet.

    Each row is expected to contain the following data:
        - column 1: source/destination bus (in the format <source>-<destination>)
        - column 2: R total (pu)
        - column 3: X total (pu)
        - column 4: B total (pu)
        - column 5: F max (MVA)

    Args:
        ws: The worksheet containing line data.
        num_buses: The number of buses in the system.

    Returns:
        A square matrix containing line-to-line admittances.
    """
    admittances = numpy.zeros((num_buses, num_buses))
    for row in ws.iter_rows(row_offset=1):
        srcdst = row[0].value
        rtotal = row[1].value
        # xtotal = row[2].value
        # btotal = row[3].value
        # fmax = row[4].value

        # TODO(kjiwa): Compute the admittance correctly.
        src, dst = [int(i) for i in srcdst.split('-')]
        admittances[src - 1][dst - 1] = rtotal

    return admittances


def main():
    """Reads an input file containing system data and initiates power flow computations."""
    args = parse_arguments()

    # Read input data.
    wb = openpyxl.load_workbook(args.input_workbook, read_only=True)
    line_data_ws = wb[args.line_data_worksheet]

    # Create admittance matrix.
    print(create_admittance_matrix(line_data_ws, args.num_buses))

    # Close input data.
    wb.close()


if __name__ == '__main__':
    main()
