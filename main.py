"""A program that analyzes and solves a power flow."""

import numpy
import power_flow
import argparse
import openpyxl

# Input data constants.
DEFAULT_INPUT_WORKBOOK = 'data/Data.xlsx'
DEFAULT_BUS_DATA_WORKSHEET = 'Bus data'
DEFAULT_LINE_DATA_WORKSHEET = 'Line data'

# Power flow constants.
DEFAULT_SLACK_BUS = 1
DEFAULT_POWER_BASE_MVA = 100
DEFAULT_MAX_MISMATCH_MW = 0.1
DEFAULT_MAX_MISMATCH_MVAR = 0.1
FLAT_START_VOLTAGE_MAGNITUDE = 1.0
FLAT_START_VOLTAGE_ANGLE_DEG = 0.0


def parse_arguments():
    """Parses command line arguments.

    Returns:
        An object containing program arguments.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument('--input_workbook', default=DEFAULT_INPUT_WORKBOOK,
                        help='An Excel workbook containing input data.')
    parser.add_argument('--bus_data_worksheet', default=DEFAULT_BUS_DATA_WORKSHEET,
                        help='The worksheet containing bus data.')
    parser.add_argument('--line_data_worksheet', default=DEFAULT_LINE_DATA_WORKSHEET,
                        help='The worksheet containing line data.')
    parser.add_argument('--start_voltage_magnitude', type=int, default=FLAT_START_VOLTAGE_MAGNITUDE,
                        help='The initial voltage to use at each bus in V.')
    parser.add_argument('--start_voltage_angle_deg', type=int, default=FLAT_START_VOLTAGE_ANGLE_DEG,
                        help='The initial voltage angle to use at each bus in degrees.')
    parser.add_argument('--slack_bus', type=int, default=DEFAULT_SLACK_BUS, help='The system slack bus.')
    parser.add_argument('--power_base_mva', type=int, default=DEFAULT_POWER_BASE_MVA,
                        help='The base apparent power value for the system in MVA.')
    parser.add_argument('--max_mismatch_mw', type=int, default=DEFAULT_MAX_MISMATCH_MW,
                        help='The maximum real power mismatch in MW for convergence testing.')
    parser.add_argument('--max_mismatch_mvar', type=int, default=DEFAULT_MAX_MISMATCH_MW,
                        help='The maximum reactive power mismatch in Mvar for convergence testing.')
    return parser.parse_args()


def main():
    """Reads an input file containing system data and initiates power flow computations."""
    args = parse_arguments()

    # Open input file.
    wb = openpyxl.load_workbook(args.input_workbook, read_only=True)
    bus_data = wb[args.bus_data_worksheet]
    line_data = wb[args.line_data_worksheet]

    # Initialize the power flow.
    start_voltage = args.start_voltage_magnitude * numpy.exp(1j * numpy.deg2rad(args.start_voltage_angle_deg))
    pf = power_flow.PowerFlow(
        bus_data, line_data, args.slack_bus, start_voltage, args.power_base_mva, args.max_mismatch_mw,
        args.max_mismatch_mvar)

    # Execute the power flow.
    buses = pf.execute()

    # TODO(kjiwa): Report on any buses or lines exceeding their operating conditions.
    print(buses)

    # Close input file.
    wb.close()


if __name__ == '__main__':
    main()
