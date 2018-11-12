"""Unit tests for admittance matrix creation."""

import numpy.testing
import openpyxl
import power_flow
from unittest import TestCase


class TestRead_admittance_matrix(TestCase):
    @staticmethod
    def create_worksheet(data):
        """Creates an in-memory worksheet with test data."""
        ws = openpyxl.Workbook().create_sheet('Line data')
        ws.append(['Header'])
        for row in data:
            ws.append(row)

        return ws

    def test_three_bus_system(self):
        """Tests that the admittance matrix for a three bus system is created correctly."""
        num_buses = 3
        line_data = [[1, 2, 0, 0.1, 0],
                     [1, 3, 0, 0.2, 0],
                     [2, 3, 0, 0.25, 0]]

        expected = [[-15j, 10j, 5j],
                    [10j, -14j, 4j],
                    [5j, 4j, -9j]]

        # Create admittance matrix.
        ws = TestRead_admittance_matrix.create_worksheet(line_data)
        actual = power_flow.read_admittance_matrix(ws, num_buses)
        numpy.testing.assert_array_equal(expected, actual)

    def test_five_bus_system(self):
        """Tests that the admittance matrix for a five bus system is created correctly.

        This system is taken from Figure 3.1 of "Power System Load Flow Analysis," by Lynn Powell.
        """
        num_buses = 5
        line_data = [[1, 2, 0.05, 0.11, 0.02],
                     [1, 3, 0.05, 0.11, 0.02],
                     [1, 5, 0.03, 0.08, 0.02],
                     [2, 3, 0.04, 0.09, 0.02],
                     [2, 5, 0.04, 0.09, 0.02],
                     [3, 4, 0.06, 0.13, 0.03],
                     [4, 5, 0.04, 0.09, 0.02]]

        expected = numpy.array([
            [10.958904 - 25.997397j, -3.424658 + 7.534247j, -3.424658 + 7.534247j, 0j, -4.109589 + 10.958904j],
            [-3.424658 + 7.534247j, 11.672080 - 26.060948j, -4.123711 + 9.278351j, 0j, -4.123711 + 9.278351j],
            [-3.424658 + 7.534247j, -4.123711 + 9.278351j, 10.475198 - 23.119061j, -2.926829 + 6.341463j, 0j],
            [0j, 0j, -2.926829 + 6.341463j, 7.050541 - 15.594814j, -4.123711 + 9.278351j],
            [-4.109589 + 10.958904j, -4.123711 + 9.278351j, 0j, -4.123711 + 9.278351j, 12.357012 - 29.485605j]])

        ws = TestRead_admittance_matrix.create_worksheet(line_data)
        actual = power_flow.read_admittance_matrix(ws, num_buses)
        numpy.testing.assert_array_almost_equal(expected, actual)
