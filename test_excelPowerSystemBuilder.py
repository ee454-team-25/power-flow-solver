import power_system
import power_system_builder
import unittest


class TestExcelPowerSystemBuilder(unittest.TestCase):
    def test_buses(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)

        actual = builder.build_buses()
        expected = [
            power_system.Bus(1, 0, 0, 0, 1),
            power_system.Bus(2, 0.4, 0.2, 0, 1),
            power_system.Bus(3, 0.25, 0.15, 0, 1),
            power_system.Bus(4, 0.4, 0.2, 0, 1),
            power_system.Bus(5, 0.5, 0.2, 0, 1)
        ]

        self.assertListEqual(expected, actual)

    def test_lines(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)

        actual = builder.build_lines()
        expected = [
            power_system.Line(1, 2, 0.05 + 0.11j, 0.02j, None),
            power_system.Line(1, 3, 0.05 + 0.11j, 0.02j, None),
            power_system.Line(1, 5, 0.03 + 0.08j, 0.02j, None),
            power_system.Line(2, 3, 0.04 + 0.09j, 0.02j, None),
            power_system.Line(2, 5, 0.04 + 0.09j, 0.02j, None),
            power_system.Line(3, 4, 0.06 + 0.13j, 0.03j, None),
            power_system.Line(4, 5, 0.04 + 0.09j, 0.02j, None)
        ]

        self.assertListEqual(expected, actual)
