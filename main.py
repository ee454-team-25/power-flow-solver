"""A program that analyzes and solves a power flow.

This program reads bus and line data from an Excel file and executes a power flow analysis to determine bus voltages. A
report about the system is then generated containing the following information:

    1. The voltage (magnitude and angle) at each bus in the system.
    2. The active and reactive power produced by each generator in the system.
    3. The active, reactive, and apparent power flowing through each line.
    4. Any violation of normal operating voltages and power limits.

The input file is expected to have two worksheets: one for bus data, and the other for line data. By default, the
worksheet names are expected to be 'Bus data' and 'Line data,' but these names may be overridden at the command line.

Bus data: The bus data worksheet should contain load and generation data for each bus in the system. Each bus is fully
specified by four variables: real power delivered, reactive power delivered, voltage, and phase angle. At least two of
these variables must be specified for each bus. The worksheet is expected to have the following structure:

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

    1. Swing bus: A particular bus should be selected as the swing bus. The power consumption at this node is set to
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
import power_flow_solver
import power_system_builder
import power_system_reporter

# Input data constants.
DEFAULT_INPUT_WORKBOOK = 'data/Data.xlsx'
DEFAULT_BUS_DATA_WORKSHEET_NAME = 'Bus data'
DEFAULT_LINE_DATA_WORKSHEET_NAME = 'Line data'

# Power flow constants.
DEFAULT_SWING_BUS_NUMBER = 1
DEFAULT_START_VOLTAGE = 1 + 0j
DEFAULT_POWER_BASE = 100
DEFAULT_MAX_ACTIVE_POWER_ERROR = 0.1
DEFAULT_MAX_REACTIVE_POWER_ERROR = 0.1
DEFAULT_MIN_OPERATING_VOLTAGE = 0.95
DEFAULT_MAX_OPERATING_VOLTAGE = 1.05


def parse_arguments():
    """Parses command line arguments.

    Returns:
        And object containing program arguments.
    """
    parser = argparse.ArgumentParser()
    parser.add_argument('--input_workbook', default=DEFAULT_INPUT_WORKBOOK,
                        help='An Excel workbook containing bus and line data.')
    parser.add_argument('--bus_data_worksheet', default=DEFAULT_BUS_DATA_WORKSHEET_NAME,
                        help='The name of the worksheet containing bus data.')
    parser.add_argument('--line_data_worksheet', default=DEFAULT_LINE_DATA_WORKSHEET_NAME,
                        help='The name of the worksheet containing line data.')
    parser.add_argument('--swing_bus_number', type=int, default=DEFAULT_SWING_BUS_NUMBER, help='The swing bus number.')
    parser.add_argument('--start_voltage_magnitude', type=float, default=numpy.abs(DEFAULT_START_VOLTAGE),
                        help='The initial voltage magnitude in volts to use when solving the power flow.')
    parser.add_argument('--start_voltage_angle', type=float, default=numpy.rad2deg(numpy.angle(DEFAULT_START_VOLTAGE)),
                        help='The initial voltage angle in degrees to use when solving the power flow.')
    parser.add_argument('--power_base', type=float, default=DEFAULT_POWER_BASE, help='The base power quantity in MVA.')
    parser.add_argument('--max_active_power_error', type=float, default=DEFAULT_MAX_ACTIVE_POWER_ERROR,
                        help='The maximum allowed mismatch between computed and actual megawatts at each bus.')
    parser.add_argument('--max_reactive_power_error', type=float, default=DEFAULT_MAX_REACTIVE_POWER_ERROR,
                        help='The maximum allowed mismatch between computed and actual megavars at each bus.')
    parser.add_argument('--min_operating_voltage', type=float, default=DEFAULT_MIN_OPERATING_VOLTAGE,
                        help='The minimum acceptable per-unit voltage magnitude at a bus.')
    parser.add_argument('--max_operating_voltage', type=float, default=DEFAULT_MAX_OPERATING_VOLTAGE,
                        help='The maximum acceptable per-unit voltage magnitude at a bus.')
    return parser.parse_args()


def main():
    """Reads an input file containing system data and initiates power flow computations."""
    args = parse_arguments()

    # Build the power system from an input file.
    start_voltage = args.start_voltage_magnitude * numpy.exp(1j * numpy.deg2rad(args.start_voltage_angle))
    builder = power_system_builder.ExcelPowerSystemBuilder(
        args.input_workbook, args.bus_data_worksheet, args.line_data_worksheet, start_voltage, args.power_base)
    system = builder.build_system()

    # Initialize the power flow.
    solver = power_flow_solver.PowerFlowSolver(system, args.swing_bus_number,
                                               args.max_active_power_error / args.power_base,
                                               args.max_reactive_power_error / args.power_base)

    # Iterate towards a solution.
    iteration = 1
    print(power_system_reporter.largest_power_mismatch_report(iteration, solver.estimates, args.power_base))
    while not solver.has_converged():
        solver.step()
        print(power_system_reporter.largest_power_mismatch_report(iteration, solver.estimates, args.power_base))
        iteration += 1

    # Produce system reports.
    print(power_system_reporter.bus_voltage_report(system, args.min_operating_voltage, args.max_operating_voltage))
    print(power_system_reporter.power_injection_report(
        solver.estimates, args.power_base, args.max_active_power_error, args.max_reactive_power_error))
    print(power_system_reporter.line_power_report(system, args.power_base))


if __name__ == '__main__':
    main()
