"""Unit tests for the power flow module."""

import numpy.testing
import openpyxl
import power_flow
from unittest import TestCase


class TestExecute_power_flow(TestCase):
    @staticmethod
    def load_state():
        """Loads sample data from Sample-Powell-3.1.xlsx."""
        wb = openpyxl.load_workbook('data/Sample-Powell-3.1.xlsx')
        bus_data_ws = wb['Bus data']
        line_data_ws = wb['Line data']

        bus_states = power_flow.read_bus_states(bus_data_ws, 1, 1.0 + 0j, 100)
        admittances = power_flow.read_admittance_matrix(line_data_ws, len(bus_states))
        return bus_states, admittances

    def test_create_bus_estimates(self):
        """Tests that bus estimates are calculated correctly."""
        bus_states, admittances = TestExecute_power_flow.load_state()
        bus_estimates = power_flow.create_bus_estimates(bus_states, admittances)

        numpy.testing.assert_almost_equal(bus_estimates[2].power, 0 - 0.03j)
        numpy.testing.assert_almost_equal(bus_estimates[3].power, 0 - 0.035j)
        numpy.testing.assert_almost_equal(bus_estimates[4].power, 0 - 0.025j)
        numpy.testing.assert_almost_equal(bus_estimates[5].power, 0 - 0.03j)

    def test_calculate_state_error(self):
        """Tests that bus errors are calculated correctly."""
        bus_states, admittances = TestExecute_power_flow.load_state()
        bus_estimates = power_flow.create_bus_estimates(bus_states, admittances)
        bus_errors = power_flow.calculate_state_error(bus_states, bus_estimates)

        numpy.testing.assert_almost_equal(bus_errors[2].power, -0.4 - 0.17j)
        numpy.testing.assert_almost_equal(bus_errors[3].power, -0.25 - 0.115j)
        numpy.testing.assert_almost_equal(bus_errors[4].power, -0.4 - 0.175j)
        numpy.testing.assert_almost_equal(bus_errors[5].power, -0.5 - 0.17j)
