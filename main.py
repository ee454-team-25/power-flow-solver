import argparse
import numpy
import power_system_builder

DEFAULT_INPUT_WORKBOOK = 'data/Data.xlsx'
DEFAULT_BUS_DATA_WORKSHEET_NAME = 'Bus data'
DEFAULT_LINE_DATA_WORKSHEET_NAME = 'Line data'
DEFAULT_START_VOLTAGE = 1 + 0j
DEFAULT_POWER_BASE = 100


def parse_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input_workbook', default=DEFAULT_INPUT_WORKBOOK,
                        help='An Excel workbook containing bus and line data.')
    parser.add_argument('--bus_data_worksheet', default=DEFAULT_BUS_DATA_WORKSHEET_NAME,
                        help='The name of the worksheet containing bus data.')
    parser.add_argument('--line_data_worksheet', default=DEFAULT_LINE_DATA_WORKSHEET_NAME,
                        help='The name of the worksheet containing line data.')
    parser.add_argument('--start_voltage_magnitude', default=numpy.abs(DEFAULT_START_VOLTAGE),
                        help='The initial voltage magnitude in volts to use when solving the power flow.')
    parser.add_argument('--start_voltage_angle', default=numpy.rad2deg(numpy.angle(DEFAULT_START_VOLTAGE)),
                        help='The initial voltage angle in degrees to use when solving the power flow.')
    parser.add_argument('--power_base', default=DEFAULT_POWER_BASE, help='The base power quantity in megawatts.')
    return parser.parse_args()


def main():
    args = parse_arguments()

    start_voltage = args.start_voltage_magnitude * numpy.exp(1j * numpy.deg2rad(args.start_voltage_angle))
    builder = power_system_builder.ExcelPowerSystemBuilder(
        args.input_workbook, args.bus_data_worksheet, args.line_data_worksheet, start_voltage, args.power_base)
    system = builder.build_system()
    print(system)


if __name__ == '__main__':
    main()
