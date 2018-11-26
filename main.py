import argparse
import numpy
import power_flow_solver
import power_system_builder

DEFAULT_INPUT_WORKBOOK = 'data/Data.xlsx'
DEFAULT_BUS_DATA_WORKSHEET_NAME = 'Bus data'
DEFAULT_LINE_DATA_WORKSHEET_NAME = 'Line data'
DEFAULT_START_VOLTAGE = 1 + 0j
DEFAULT_POWER_BASE = 100
DEFAULT_MAX_ACTIVE_POWER_ERROR = 0.1
DEFAULT_MAX_REACTIVE_POWER_ERROR = 0.1


def parse_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input_workbook', default=DEFAULT_INPUT_WORKBOOK,
                        help='An Excel workbook containing bus and line data.')
    parser.add_argument('--bus_data_worksheet', default=DEFAULT_BUS_DATA_WORKSHEET_NAME,
                        help='The name of the worksheet containing bus data.')
    parser.add_argument('--line_data_worksheet', default=DEFAULT_LINE_DATA_WORKSHEET_NAME,
                        help='The name of the worksheet containing line data.')
    parser.add_argument('--start_voltage_magnitude', type=float, default=numpy.abs(DEFAULT_START_VOLTAGE),
                        help='The initial voltage magnitude in volts to use when solving the power flow.')
    parser.add_argument('--start_voltage_angle', type=float, default=numpy.rad2deg(numpy.angle(DEFAULT_START_VOLTAGE)),
                        help='The initial voltage angle in degrees to use when solving the power flow.')
    parser.add_argument('--power_base', type=float, default=DEFAULT_POWER_BASE, help='The base power quantity in MVA.')
    parser.add_argument('--max_active_power_error', type=float, default=DEFAULT_MAX_ACTIVE_POWER_ERROR,
                        help='The maximum allowed mismatch between computed and actual megawatts at each bus.')
    parser.add_argument('--max_reactive_power_error', type=float, default=DEFAULT_MAX_REACTIVE_POWER_ERROR,
                        help='The maximum allowed mismatch between computed and actual megavars at each bus.')
    return parser.parse_args()


def main():
    args = parse_arguments()

    start_voltage = args.start_voltage_magnitude * numpy.exp(1j * numpy.deg2rad(args.start_voltage_angle))
    builder = power_system_builder.ExcelPowerSystemBuilder(
        args.input_workbook, args.bus_data_worksheet, args.line_data_worksheet, start_voltage, args.power_base)
    system = builder.build_system()

    solver = power_flow_solver.PowerFlowSolver(system, args.max_active_power_error, args.max_reactive_power_error)
    while not solver.has_converged():
        solver.step()

    for bus in system.buses:
        print(u'Bus: {}, Voltage: {:.4f}\u2220{:.4f} deg'.format(bus.number, numpy.abs(bus.voltage),
                                                                 numpy.rad2deg(numpy.angle(bus.voltage))))


if __name__ == '__main__':
    main()
