import openpyxl
import power_system

FLAT_START_VOLTAGE = 1 + 0j
DEFAULT_BUS_DATA_WORKSHEET_NAME = 'Bus data'
DEFAULT_LINE_DATA_WORKSHEET_NAME = 'Line data'
DEFAULT_POWER_BASE = 100


class ExcelPowerSystemBuilder:
    def __init__(self, filename, bus_data_worksheet_name=DEFAULT_BUS_DATA_WORKSHEET_NAME,
                 line_data_worksheet_name=DEFAULT_LINE_DATA_WORKSHEET_NAME, start_voltage=FLAT_START_VOLTAGE,
                 power_base=DEFAULT_POWER_BASE):
        self._workbook = openpyxl.load_workbook(filename, read_only=True)
        self._bus_data_worksheet = self._workbook[bus_data_worksheet_name]
        self._line_data_worksheet = self._workbook[line_data_worksheet_name]
        self._start_voltage = start_voltage
        self._power_base = power_base

    def build_buses(self):
        result = []
        for row in self._bus_data_worksheet.iter_rows(row_offset=1):
            bus_number = row[0].value
            if not bus_number:
                break

            p_load = (row[1].value or 0) / self._power_base
            q_load = (row[2].value or 0) / self._power_base
            p_generator = (row[3].value or 0) / self._power_base

            p_voltage = row[4].value or self._start_voltage
            result.append(power_system.Bus(bus_number, p_load, q_load, p_generator, p_voltage))

        return result

    def build_lines(self):
        result = []
        for row in self._line_data_worksheet.iter_rows(row_offset=1):
            source_bus_number = row[0].value
            destination_bus_number = row[1].value
            if not source_bus_number or not destination_bus_number:
                break

            r_distributed = row[2].value or 0
            x_distributed = row[3].value or 0
            z_distributed = r_distributed + 1j * x_distributed

            y_shunt = 1j * row[4].value or 0j
            max_power = row[5].value

            result.append(
                power_system.Line(source_bus_number, destination_bus_number, z_distributed, y_shunt, max_power))

        return result

    def build_system(self):
        return power_system.PowerSystem(self.build_buses(), self.build_lines())
