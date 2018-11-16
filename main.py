import argparse
import power_system_builder

DEFAULT_INPUT_WORKBOOK = 'data/Data.xlsx'
DEFAULT_BUS_DATA_WORKSHEET_NAME = 'Bus data'
DEFAULT_LINE_DATA_WORKSHEET_NAME = 'Line data'


def parse_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input_workbook', default=DEFAULT_INPUT_WORKBOOK,
                        help='An Excel workbook containing bus and line data.')
    parser.add_argument('--bus_data_worksheet', default=DEFAULT_BUS_DATA_WORKSHEET_NAME,
                        help='The name of the worksheet containing bus data.')
    parser.add_argument('--line_data_worksheet', default=DEFAULT_LINE_DATA_WORKSHEET_NAME,
                        help='The name of the worksheet containing line data.')
    return parser.parse_args()


def main():
    args = parse_arguments()
    builder = power_system_builder.ExcelPowerSystemBuilder(args.input_workbook, args.bus_data_worksheet,
                                                           args.line_data_worksheet)
    system = builder.system()
    print(system)


if __name__ == '__main__':
    main()
