import openpyxl
import power_system


class ExcelPowerSystemBuilder:
    def __init__(self, filename, bus_data_worksheet_name, line_data_worksheet_name):
        self._workbook = openpyxl.load_workbook(filename, read_only=True)
        self._bus_data_worksheet = self._workbook[bus_data_worksheet_name]
        self._line_data_worksheet = self._workbook[line_data_worksheet_name]

    def buses(self):
        result = {}
        for row in self._bus_data_worksheet.iter_rows(row_offset=1):
            bus_number = row[0].value
            if not bus_number:
                break

            bus = power_system.Bus(bus_number, [], [])
            result[bus_number] = bus

            p_load = row[1].value or 0
            q_load = row[2].value or 0
            p_generator = row[3].value or 0
            p_voltage = row[4].value or 0

            if p_load or q_load:
                bus.loads.append(power_system.Load(p_load, q_load))

            if p_generator and p_voltage:
                bus.generators.append(power_system.Generator(p_generator, p_voltage))

        return result

    def lines(self):
        result = []
        for row in self._line_data_worksheet.iter_rows(row_offset=1):
            source_bus_number = row[0].value
            destination_bus_number = row[1].value

            r_distributed = row[2].value or 0
            x_distributed = row[3].value or 0
            z_distributed = r_distributed + 1j * x_distributed

            y_shunt = 1j * row[4].value or 0j
            max_power = row[5].value

            result.append(
                power_system.Line(source_bus_number, destination_bus_number, z_distributed, y_shunt, max_power))

        return result

    def system(self):
        return power_system.PowerSystem(self.buses(), self.lines())
