"""A program that analyzes and solves a power flow.

This program reads bus and line data from an Excel file and executes a power flow analysis to determine bus voltages.
A report about the system is then generated containing the following information:

    1. The voltage (magnitude and angle) at each bus in the system.
    2. The active and reactive power produced by each generator in the system.
    3. The active, reactive, and apparent power flowing through each line.
    4. Any violation of normal operating voltages and power limits.

The input file is expected to have two worksheets: one for bus data, and the other for line data. By default, the
worksheet names are expected to be 'Bus data' and 'Line data,' but these names may be overridden at the command line.

Bus data: The bus data worksheet should contain load and generation data for each bus in the system. Each bus is fully
specified by four variables: real power consumption, reactive power consumption, voltage, and phase angle. At least two
of these variables must be specified for each bus. The worksheet is expected to have the following structure:

    1. Bus number
    2. Real power consumed (MW)
    3. Reactive power consumed (Mvar)
    4. Real power delivered (MW)
    5. Voltage (pu)

Line data: The line data worksheet should contain parameters for each transmission line and transformer in the system.
In this analysis program, shunt conductances are assumed to be negligible and are ignored. The worksheet is expected to
have the following structure:

    1. Source bus
    2. Destination bus
    3. Total resistance (pu)
    4. Total reactance (pu)
    5. Total susceptance (pu)
    6. Maximum apparent power (MVA)

Program arguments: In addition to bus and line data, power flow analysis requires several parameters to be specified:

    1. Slack bus: A particular bus should be selected as the slack bus. The power consumption at this node is set to
       balance all power consumption and generation in the system. By default, the first bus is selected.

    2. Power base: The calculations executed during analysis assume that system data is represented in per-unit
       values. The power base value is used to perform this conversion. By default, the power base is set to 100 MVA.

    3. Maximum allowable power mismatch: The maximum allowable real and reactive power mismatches provide bounds that
       are used to determine if the power flow analysis has converged to a solution. By default, the maximum allowable
       mismatches are 0.1 MW and 0.1 Mvar.

    4. Start voltage: When the system executes, any unknown voltages are given a default value and then iteratively
       refined. This initial value is typically set to be 1.0 per-unit with 0 phase angle, also known as a "flat
       start."
"""

import argparse
import numpy
import openpyxl
import power_flow

# Input data constants.
DEFAULT_INPUT_WORKBOOK = 'data/Data.xlsx'
DEFAULT_BUS_DATA_WORKSHEET = 'Bus data'
DEFAULT_LINE_DATA_WORKSHEET = 'Line data'

# Power flow constants.
DEFAULT_SLACK_BUS_NUMBER = 1
DEFAULT_POWER_BASE_MVA = 100
DEFAULT_MAX_MISMATCH_MW = 0.1
DEFAULT_MAX_MISMATCH_MVAR = 0.1
START_VOLTAGE_PU = 1.0 + 0j


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
    parser.add_argument('--start_voltage_magnitude_pu', type=float, default=numpy.abs(START_VOLTAGE_PU),
                        help='The initial voltage to use at each bus in per-unit.')
    parser.add_argument('--start_voltage_angle_deg', type=float, default=numpy.rad2deg(numpy.angle(START_VOLTAGE_PU)),
                        help='The initial voltage angle to use at each bus in degrees.')
    parser.add_argument('--slack_bus_number', type=int, default=DEFAULT_SLACK_BUS_NUMBER,
                        help='The system slack bus number.')
    parser.add_argument('--power_base_mva', type=float, default=DEFAULT_POWER_BASE_MVA,
                        help='The base apparent power value for the system in MVA.')
    parser.add_argument('--max_mismatch_mw', type=float, default=DEFAULT_MAX_MISMATCH_MW,
                        help='The maximum real power mismatch in MW for convergence testing.')
    parser.add_argument('--max_mismatch_mvar', type=float, default=DEFAULT_MAX_MISMATCH_MW,
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
    start_voltage = args.start_voltage_magnitude_pu * numpy.exp(1j * numpy.deg2rad(args.start_voltage_angle_deg))
    pf = power_flow.PowerFlow(
        bus_data, line_data, args.slack_bus_number, start_voltage, args.power_base_mva, args.max_mismatch_mw,
        args.max_mismatch_mvar)

    # Execute the power flow.
    buses = pf.execute()

    # TODO(kjiwa): Report on any buses or lines exceeding their operating conditions.
    for bus in buses:
        print(bus)

    # Close input file.
    wb.close()


if __name__ == '__main__':
    main()
