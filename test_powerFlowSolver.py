import numpy
import power_flow_solver
import power_system_builder
import unittest


class TestPowerFlowSolver(unittest.TestCase):
    def test_errors(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()
        solver = power_flow_solver.PowerFlowSolver(system)

        p_expected = [-0.4, -0.25, -0.4, -0.5]
        q_expected = [-0.17, -0.115, -0.175, -0.17]

        actual = solver._bus_power_estimates()
        p_actual = [estimate.active_power_error for estimate in actual.values()
                    if estimate.type in (power_flow_solver._BusType.PQ, power_flow_solver._BusType.PV)]
        q_actual = [estimate.reactive_power_error for estimate in actual.values()
                    if estimate.type == power_flow_solver._BusType.PQ]

        numpy.testing.assert_array_almost_equal(p_expected, p_actual)
        numpy.testing.assert_array_almost_equal(q_expected, q_actual)

    def test_jacobian_11(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()
        solver = power_flow_solver.PowerFlowSolver(system)

        expected = [[26.030948, -9.278351, 0, -9.278350],
                    [-9.278351, 23.084061, -6.341463, 0],
                    [0, -6.341463, 15.569814, -9.278351],
                    [-9.278351, 0, -9.278351, 29.455605]]
        actual = solver._jacobian_11(solver._bus_power_estimates())
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian_12(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()
        solver = power_flow_solver.PowerFlowSolver(system)
        estimates = solver._bus_power_estimates()
        pq_estimates = {i.bus.number: i for i in estimates.values() if i.type == power_flow_solver._BusType.PQ}

        expected = [[11.672080, -4.123711, 0, -4.123711],
                    [-4.123711, 10.475198, -2.926829, 0],
                    [0, -2.926829, 7.050541, -4.123711],
                    [-4.123711, 0, -4.123711, 12.357011]]

        actual = solver._jacobian_12(estimates, pq_estimates)
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian_21(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()
        solver = power_flow_solver.PowerFlowSolver(system)
        estimates = solver._bus_power_estimates()
        pq_estimates = {i.bus.number: i for i in estimates.values() if i.type == power_flow_solver._BusType.PQ}

        expected = [[-11.672080, 4.123711, 0, 4.123711],
                    [4.123711, -10.475198, 2.926829, 0],
                    [0, 2.926829, -7.050541, 4.123711],
                    [4.123711, 0, 4.123711, -12.357011]]
        actual = solver._jacobian_21(estimates, pq_estimates)
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian_22(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()
        solver = power_flow_solver.PowerFlowSolver(system)
        estimates = solver._bus_power_estimates()
        pq_estimates = {i.bus.number: i for i in estimates.values() if i.type == power_flow_solver._BusType.PQ}

        expected = [[26.090948, -9.278351, 0, -9.278351],
                    [-9.278351, 23.154061, -6.341463, 0],
                    [0, -6.341463, 15.619814, -9.278351],
                    [-9.278351, 0, -9.278351, 29.515605]]
        actual = solver._jacobian_22(pq_estimates)
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_jacobian(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()
        solver = power_flow_solver.PowerFlowSolver(system)
        estimates = solver._bus_power_estimates()
        pq_estimates = {i.bus.number: i for i in estimates.values() if i.type == power_flow_solver._BusType.PQ}

        expected = [[26.030948, -9.278351, 0, -9.278350, 11.672080, -4.123711, 0, -4.123711],
                    [-9.278351, 23.084061, -6.341463, 0, -4.123711, 10.475198, -2.926829, 0],
                    [0, -6.341463, 15.569814, -9.278351, 0, -2.926829, 7.050541, -4.123711],
                    [-9.278351, 0, -9.278351, 29.455605, -4.123711, 0, -4.123711, 12.357011],
                    [-11.672080, 4.123711, 0, 4.123711, 26.090948, -9.278351, 0, -9.278351],
                    [4.123711, -10.475198, 2.926829, 0, -9.278351, 23.154061, -6.341463, 0],
                    [0, 2.926829, -7.050541, 4.123711, 0, -6.341463, 15.619814, -9.278351],
                    [4.123711, 0, 4.123711, -12.357011, -9.278351, 0, -9.278351, 29.515605]]
        actual = solver._jacobian(estimates, pq_estimates)
        numpy.testing.assert_array_almost_equal(expected, actual)

    def test_corrections(self):
        filename = 'data/Sample-Powell-3.1.xlsx'
        builder = power_system_builder.ExcelPowerSystemBuilder(filename)
        system = builder.build_system()
        solver = power_flow_solver.PowerFlowSolver(system)
        estimates = solver._bus_power_estimates()
        pq_estimates = {i.bus.number: i for i in estimates.values() if i.type == power_flow_solver._BusType.PQ}
        jacobian = solver._jacobian(estimates, pq_estimates)

        expected = numpy.transpose(
            [[-0.040599, -0.040003, -0.060281, -0.04515, -0.041184, -0.041492, -0.061248, -0.042768]])
        actual = solver._corrections(jacobian, estimates, pq_estimates)
        numpy.testing.assert_array_almost_equal(expected, actual)
