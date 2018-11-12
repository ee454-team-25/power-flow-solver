"""Unit tests for Jacobian matrix creation."""

import numpy.testing
import openpyxl
import power_flow
import power_flow_jacobian
from unittest import TestCase


class TestCreate_jacobian_matrix(TestCase):
    @staticmethod
    def load_state():
        """Loads sample data from Sample-Powell-3.1.xslx."""
        wb = openpyxl.load_workbook('data/Sample-Powell-3.1.xlsx')
        bus_data_ws = wb['Bus data']
        line_data_ws = wb['Line data']

        bus_states = power_flow.read_bus_states(bus_data_ws, 1, 1.0 + 0j, 100)
        admittances = power_flow.read_admittance_matrix(line_data_ws, len(bus_states))
        return bus_states, admittances

    def test_create_j11_matrix(self):
        """Tests that the J11 matrix is calculated correctly."""
        expected = [[26.09095, -9.278350, 0, -9.278350],
                    [-9.278350, 23.154060, -6.341464, 0],
                    [0, -6.341464, 15.619813, -9.278350],
                    [-9.278350, 0, -9.278350, 29.515605]]

        bus_states, admittances = TestCreate_jacobian_matrix.load_state()
        bus_estimates = power_flow.create_bus_estimates(bus_states, admittances)

        load_bus_estimates = [bus for bus in bus_estimates.values() if bus.type == power_flow.BusType.LOAD]
        gen_bus_estimates = [bus for bus in bus_estimates.values() if bus.type == power_flow.BusType.GENERATOR]
        actual = power_flow_jacobian.create_j11_matrix(load_bus_estimates, gen_bus_estimates, admittances)

        numpy.testing.assert_array_almost_equal(expected, actual, 5)

    def test_create_j12_matrix(self):
        """Tests that the J12 matrix is computed correctly."""
        expected = [[11.672080, -4.123711, 0, -4.123711],
                    [-4.123711, 10.475198, -2.926829, 0],
                    [0, -2.926829, 7.050541, -4.123711],
                    [-4.123711, 0, -4.123711, 12.357011]]

        bus_states, admittances = TestCreate_jacobian_matrix.load_state()
        bus_estimates = power_flow.create_bus_estimates(bus_states, admittances)

        load_bus_estimates = [bus for bus in bus_estimates.values() if bus.type == power_flow.BusType.LOAD]
        gen_bus_estimates = [bus for bus in bus_estimates.values() if bus.type == power_flow.BusType.GENERATOR]
        actual = power_flow_jacobian.create_j12_matrix(load_bus_estimates, gen_bus_estimates, admittances)

        numpy.testing.assert_array_almost_equal(expected, actual, 5)

    def test_create_j21_matrix(self):
        """Tests that the J21 matrix is computed correctly."""
        expected = [[-11.672080, 4.123711, 0, 4.123711],
                    [4.123711, -10.475198, 2.926829, 0],
                    [0, 2.926829, -7.050541, 4.123711, 0],
                    [4.123711, 0, 4.123711, -12.357011]]

        bus_states, admittances = TestCreate_jacobian_matrix.load_state()
        bus_estimates = power_flow.create_bus_estimates(bus_states, admittances)

        load_bus_estimates = [bus for bus in bus_estimates.values() if bus.type == power_flow.BusType.LOAD]
        gen_bus_estimates = [bus for bus in bus_estimates.values() if bus.type == power_flow.BusType.GENERATOR]
        actual = power_flow_jacobian.create_j21_matrix(load_bus_estimates, gen_bus_estimates, admittances)

        numpy.testing.assert_array_almost_equal(expected, actual, 5)

    def test_create_j22_matrix(self):
        """Tests that the J22 matrix is computed correctly."""
        expected = [[26.03095, -9.278350, 0, -9.278350],
                    [-9.278350, 23.08406, -6.341464, 0],
                    [0, -6.341464, 15.56981, -9.278350],
                    [-9.278350, 0, -9.278350, 29.45561]]

        bus_states, admittances = TestCreate_jacobian_matrix.load_state()
        bus_estimates = power_flow.create_bus_estimates(bus_states, admittances)

        load_bus_estimates = [bus for bus in bus_estimates.values() if bus.type == power_flow.BusType.LOAD]
        actual = power_flow_jacobian.create_j22_matrix(load_bus_estimates, admittances)

        numpy.testing.assert_array_almost_equal(expected, actual, 5)

    def test_create_jacobian_matrix(self):
        """Tests that the Jacobian matrix is computed correctly."""
        expected = [[26.09095, -9.278350, 0, -9.278350, 11.672080, -4.123711, 0, -4.123711],
                    [9.278350, 23.154060, -6.341464, 0, -4.123711, 10.475198, -2.926829, 0],
                    [0, -6.341464, 15.619813, -9.278350, 0, -2.926829, 7.050541, -4.123711],
                    [-9.278350, 0, -9.278350, 29.515605, -4.123711, 0, -4.123711, 12.357011],
                    [-11.672080, 4.123711, 0, 4.123711, 26.03095, -9.278350, 0, -9.278350],
                    [4.123711, -10.475198, 2.926829, 0, -9.278350, 23.108406, -6.341464, 0],
                    [0, 2.926829, -7.050541, 4.123711, 0, -6.341464, 15.56981, -9.278350],
                    [4.123711, 0, 4.123711, -12.357011, -9.278350, 0, -9.278350, 29.45561]]

        bus_states, admittances = TestCreate_jacobian_matrix.load_state()
        bus_estimates = power_flow.create_bus_estimates(bus_states, admittances)

        actual = power_flow_jacobian.create_jacobian_matrix(bus_estimates, admittances)
        numpy.testing.assert_array_almost_equal(expected, actual, 5)
